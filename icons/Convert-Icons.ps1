#Requires -Version 5.1
<#
.SYNOPSIS
  Convert product PNGs to multi-size alpha-preserving ICO files and verify transparency.

  Uses a hybrid ICO format for reliable Explorer / app embedding:
  - 16/32/48/64: classic 32bpp BMP (XOR+AND) with alpha channel
  - 256: PNG-compressed frame
#>
[CmdletBinding()]
param(
    [switch]$SkipPreview
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

$root = $PSScriptRoot
$pictures = Join-Path $root 'pictures'
$previewDir = Join-Path $root '_preview'

$jobs = @(
    @{ Name = 'Word';       Pattern = '*Word_0.png';       Out = 'Word.ico' }
    @{ Name = 'Excel';      Pattern = '*Excel*_0.png';     Out = 'Excel.ico' }
    @{ Name = 'PowerPoint'; Pattern = '*PowerPoint_0.png'; Out = 'PowerPoint.ico' }
)

# Common Windows icon sizes (including 64 for taskbar / shortcuts).
$sizes = @(16, 32, 48, 64, 256)

$code = @'
using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;

public static class MosIconBuilderV2
{
    public static Bitmap ResizeArgb(Bitmap source, int size)
    {
        var dest = new Bitmap(size, size, PixelFormat.Format32bppArgb);
        dest.SetResolution(source.HorizontalResolution, source.VerticalResolution);
        using (var g = Graphics.FromImage(dest))
        {
            g.Clear(Color.Transparent);
            g.CompositingMode = CompositingMode.SourceOver;
            g.CompositingQuality = CompositingQuality.HighQuality;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.SmoothingMode = SmoothingMode.HighQuality;
            g.PixelOffsetMode = PixelOffsetMode.HighQuality;
            g.DrawImage(source, new Rectangle(0, 0, size, size));
        }
        return dest;
    }

    public static byte[] BuildClassic32bppFrame(Bitmap bmp)
    {
        int w = bmp.Width;
        int h = bmp.Height;
        int xorStride = w * 4;
        int andStride = ((w + 31) / 32) * 4;
        var xor = new byte[xorStride * h];
        var and = new byte[andStride * h];

        for (int y = 0; y < h; y++)
        {
            int destY = h - 1 - y;
            for (int x = 0; x < w; x++)
            {
                Color c = bmp.GetPixel(x, y);
                int i = destY * xorStride + x * 4;
                xor[i] = c.B;
                xor[i + 1] = c.G;
                xor[i + 2] = c.R;
                xor[i + 3] = c.A;
            }
        }

        using (var ms = new MemoryStream())
        using (var bw = new BinaryWriter(ms))
        {
            bw.Write(40);                 // biSize
            bw.Write(w);                  // biWidth
            bw.Write(h * 2);              // biHeight (XOR + AND)
            bw.Write((ushort)1);          // biPlanes
            bw.Write((ushort)32);         // biBitCount
            bw.Write(0);                  // BI_RGB
            bw.Write(xor.Length + and.Length);
            bw.Write(0);
            bw.Write(0);
            bw.Write(0);
            bw.Write(0);
            bw.Write(xor);
            bw.Write(and);
            return ms.ToArray();
        }
    }

    public static void SaveHybridIcon(string path, Bitmap[] frames)
    {
        using (var fs = new FileStream(path, FileMode.Create, FileAccess.Write))
        using (var bw = new BinaryWriter(fs))
        {
            bw.Write((ushort)0);
            bw.Write((ushort)1);
            bw.Write((ushort)frames.Length);

            var blobs = new byte[frames.Length][];
            for (int i = 0; i < frames.Length; i++)
            {
                int dim = Math.Max(frames[i].Width, frames[i].Height);
                if (dim >= 256)
                {
                    using (var ms = new MemoryStream())
                    {
                        frames[i].Save(ms, ImageFormat.Png);
                        blobs[i] = ms.ToArray();
                    }
                }
                else
                {
                    blobs[i] = BuildClassic32bppFrame(frames[i]);
                }
            }

            int offset = 6 + (16 * frames.Length);
            for (int i = 0; i < frames.Length; i++)
            {
                int w = frames[i].Width;
                int h = frames[i].Height;
                bw.Write((byte)(w >= 256 ? 0 : w));
                bw.Write((byte)(h >= 256 ? 0 : h));
                bw.Write((byte)0);
                bw.Write((byte)0);
                bw.Write((ushort)1);
                bw.Write((ushort)32);
                bw.Write(blobs[i].Length);
                bw.Write(offset);
                offset += blobs[i].Length;
            }

            for (int i = 0; i < blobs.Length; i++)
            {
                bw.Write(blobs[i]);
            }
        }
    }

    public struct AlphaStats
    {
        public int Transparent;
        public int Partial;
        public int Opaque;
        public int CornerMinAlpha;
        public int Total;
    }

    public static AlphaStats Analyze(Bitmap bmp)
    {
        var s = new AlphaStats();
        s.Total = bmp.Width * bmp.Height;
        int c0 = bmp.GetPixel(0, 0).A;
        int c1 = bmp.GetPixel(bmp.Width - 1, 0).A;
        int c2 = bmp.GetPixel(0, bmp.Height - 1).A;
        int c3 = bmp.GetPixel(bmp.Width - 1, bmp.Height - 1).A;
        s.CornerMinAlpha = Math.Min(Math.Min(c0, c1), Math.Min(c2, c3));

        for (int y = 0; y < bmp.Height; y++)
        {
            for (int x = 0; x < bmp.Width; x++)
            {
                int a = bmp.GetPixel(x, y).A;
                if (a == 0) s.Transparent++;
                else if (a == 255) s.Opaque++;
                else s.Partial++;
            }
        }
        return s;
    }

    public static Bitmap LoadLargestFrameFromIco(string icoPath)
    {
        byte[] data = File.ReadAllBytes(icoPath);
        if (data.Length < 6) throw new InvalidDataException("ICO too small");
        ushort type = BitConverter.ToUInt16(data, 2);
        ushort count = BitConverter.ToUInt16(data, 4);
        if (type != 1 || count == 0) throw new InvalidDataException("Not an icon");

        int bestOffset = -1;
        int bestSize = -1;
        int bestDim = -1;
        for (int i = 0; i < count; i++)
        {
            int entry = 6 + i * 16;
            int w = data[entry] == 0 ? 256 : data[entry];
            int h = data[entry + 1] == 0 ? 256 : data[entry + 1];
            int bytes = BitConverter.ToInt32(data, entry + 8);
            int offset = BitConverter.ToInt32(data, entry + 12);
            int dim = Math.Max(w, h);
            if (dim > bestDim)
            {
                bestDim = dim;
                bestOffset = offset;
                bestSize = bytes;
            }
        }

        if (bestOffset < 0 || bestOffset + bestSize > data.Length)
            throw new InvalidDataException("Bad ICO frame");

        if (data[bestOffset] == 0x89 && data[bestOffset + 1] == 0x50)
        {
            using (var ms = new MemoryStream(data, bestOffset, bestSize))
            {
                return new Bitmap(ms);
            }
        }

        // Classic BMP frame: BITMAPINFOHEADER + XOR + AND
        int biWidth = BitConverter.ToInt32(data, bestOffset + 4);
        int biHeight = BitConverter.ToInt32(data, bestOffset + 8) / 2;
        int biBitCount = BitConverter.ToUInt16(data, bestOffset + 14);
        if (biBitCount != 32) throw new InvalidDataException("Expected 32bpp classic icon frame");

        var bmp = new Bitmap(biWidth, biHeight, PixelFormat.Format32bppArgb);
        int xorStride = biWidth * 4;
        int xorStart = bestOffset + 40;
        for (int y = 0; y < biHeight; y++)
        {
            int srcY = biHeight - 1 - y;
            for (int x = 0; x < biWidth; x++)
            {
                int i = xorStart + srcY * xorStride + x * 4;
                byte b = data[i];
                byte g = data[i + 1];
                byte r = data[i + 2];
                byte a = data[i + 3];
                bmp.SetPixel(x, y, Color.FromArgb(a, r, g, b));
            }
        }
        return bmp;
    }

    public static string DescribeIco(string icoPath)
    {
        byte[] data = File.ReadAllBytes(icoPath);
        ushort count = BitConverter.ToUInt16(data, 4);
        var parts = new System.Collections.Generic.List<string>();
        for (int i = 0; i < count; i++)
        {
            int entry = 6 + i * 16;
            int w = data[entry] == 0 ? 256 : data[entry];
            int h = data[entry + 1] == 0 ? 256 : data[entry + 1];
            int bytes = BitConverter.ToInt32(data, entry + 8);
            int offset = BitConverter.ToInt32(data, entry + 12);
            bool png = data[offset] == 0x89 && data[offset + 1] == 0x50;
            parts.Add(string.Format("{0}x{1}:{2}/{3}b", w, h, png ? "PNG" : "BMP", bytes));
        }
        return string.Join(", ", parts.ToArray());
    }

    public static void WriteCheckerPreview(Bitmap icon, string path)
    {
        int size = icon.Width;
        var preview = new Bitmap(size, size, PixelFormat.Format32bppArgb);
        using (var g = Graphics.FromImage(preview))
        {
            int cell = Math.Max(4, size / 16);
            for (int y = 0; y < size; y += cell)
            {
                for (int x = 0; x < size; x += cell)
                {
                    bool light = ((x / cell) + (y / cell)) % 2 == 0;
                    using (var b = new SolidBrush(light ? Color.FromArgb(220, 220, 220) : Color.FromArgb(160, 160, 160)))
                    {
                        g.FillRectangle(b, x, y, cell, cell);
                    }
                }
            }
            g.CompositingMode = CompositingMode.SourceOver;
            g.DrawImage(icon, 0, 0, size, size);
        }
        preview.Save(path, ImageFormat.Png);
        preview.Dispose();
    }
}
'@

try {
    Add-Type -TypeDefinition $code -ReferencedAssemblies System.Drawing -ErrorAction Stop
}
catch {
    if ($_.Exception.Message -notmatch 'already exists') { throw }
}

function Assert-AlphaOk {
    param(
        [string]$Label,
        [object]$Stats,
        [switch]$RequirePartial
    )
    if ($Stats.CornerMinAlpha -gt 8) {
        throw "$Label`: corners are not transparent (min A=$($Stats.CornerMinAlpha))"
    }
    if ($Stats.Transparent -lt ([math]::Max(50, [int]($Stats.Total * 0.2)))) {
        throw "$Label`: too few transparent pixels ($($Stats.Transparent))"
    }
    if ($RequirePartial -and $Stats.Partial -le 0) {
        throw "$Label`: expected semi-transparent edge pixels, found none"
    }
    $pctT = [math]::Round(100.0 * $Stats.Transparent / $Stats.Total, 1)
    $pctP = [math]::Round(100.0 * $Stats.Partial / $Stats.Total, 1)
    $pctO = [math]::Round(100.0 * $Stats.Opaque / $Stats.Total, 1)
    Write-Host ("  {0}: A0={1}% Partial={2}% Opaque={3}% cornerMinA={4}" -f `
        $Label, $pctT, $pctP, $pctO, $Stats.CornerMinAlpha)
}

if (-not $SkipPreview) {
    New-Item -ItemType Directory -Force -Path $previewDir | Out-Null
}

foreach ($job in $jobs) {
    $match = Get-ChildItem -LiteralPath $pictures -Filter $job.Pattern -File | Select-Object -First 1
    if (-not $match) {
        throw "Source PNG not found for pattern: $($job.Pattern) in $pictures"
    }
    $srcPath = $match.FullName
    $outPath = Join-Path $root $job.Out

    Write-Host "=== $($job.Name) ==="
    Write-Host "  Source $($match.Name)"
    $source = New-Object System.Drawing.Bitmap $srcPath
    try {
        if ($source.PixelFormat -ne [System.Drawing.Imaging.PixelFormat]::Format32bppArgb) {
            $converted = New-Object System.Drawing.Bitmap $source.Width, $source.Height, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
            $g = [System.Drawing.Graphics]::FromImage($converted)
            try {
                $g.DrawImage($source, 0, 0, $source.Width, $source.Height)
            }
            finally {
                $g.Dispose()
            }
            $source.Dispose()
            $source = $converted
        }

        $srcStats = [MosIconBuilderV2]::Analyze($source)
        Assert-AlphaOk -Label "$($job.Name) PNG" -Stats $srcStats -RequirePartial

        $frames = @()
        foreach ($size in $sizes) {
            $frames += [MosIconBuilderV2]::ResizeArgb($source, $size)
        }

        try {
            [MosIconBuilderV2]::SaveHybridIcon($outPath, $frames)
            Write-Host "  Wrote $outPath"
            Write-Host ("  Frames: {0}" -f [MosIconBuilderV2]::DescribeIco($outPath))

            $icoBmp = [MosIconBuilderV2]::LoadLargestFrameFromIco($outPath)
            try {
                $icoStats = [MosIconBuilderV2]::Analyze($icoBmp)
                Assert-AlphaOk -Label "$($job.Name) ICO/256" -Stats $icoStats -RequirePartial

                if (-not $SkipPreview) {
                    $previewPath = Join-Path $previewDir ("{0}_checker.png" -f $job.Name)
                    [MosIconBuilderV2]::WriteCheckerPreview($icoBmp, $previewPath)
                    Write-Host "  Preview $previewPath"
                }
            }
            finally {
                $icoBmp.Dispose()
            }
        }
        finally {
            foreach ($f in $frames) { $f.Dispose() }
        }
    }
    finally {
        $source.Dispose()
    }
}

Write-Host "=== All icons converted and alpha-verified ==="

# Keep project-local AppIcon.ico in sync (ApplicationIcon must not use ..\ paths;
# ClickOnce/SxS manifests reject parent-directory file names).
$projectIcons = @(
    @{ Src = 'Word.ico';       Dest = Join-Path $root '..\MOSapp\MOS Word app\AppIcon.ico' }
    @{ Src = 'Excel.ico';      Dest = Join-Path $root '..\MOSapp\mos_xaml_app\AppIcon.ico' }
    @{ Src = 'PowerPoint.ico'; Dest = Join-Path $root '..\MOSapp\Mos PowerPoint Mogi App\AppIcon.ico' }
    # Cover (MOSapp) uses Word rabbit as the shared brand icon for ClickOnce Start Menu.
    @{ Src = 'Word.ico';       Dest = Join-Path $root '..\MOSapp\MOSapp\AppIcon.ico' }
)
foreach ($pi in $projectIcons) {
    $src = Join-Path $root $pi.Src
    $dest = [IO.Path]::GetFullPath($pi.Dest)
    Copy-Item -LiteralPath $src -Destination $dest -Force
    Write-Host "  Synced $dest"
}
