using System;
using System.Runtime.InteropServices;
using System.Reflection;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Libraries;

namespace Libraries.Group1
{
    public class WordChecker1_5
    {
        public bool CheckTask_1_5_01() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_5_01(filePath); } catch { return false; } }
        public bool CheckTask_1_5_02() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_5_02(filePath); } catch { return false; } }
        public bool CheckTask_1_5_03() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_5_03(filePath); } catch { return false; } }
        public bool CheckTask_1_5_04() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_5_04(filePath); } catch { return false; } }
        public bool CheckTask_1_5_05() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_5_05(filePath); } catch { return false; } }
        public bool CheckTask_1_5_06() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_5_06(filePath); } catch { return false; } }
        public bool CheckTask_1_5_07() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_5_07(filePath); } catch { return false; } }
        public bool CheckTask_1_5_08() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_5_08(filePath); } catch { return false; } }

        private bool CheckTask_1_5_01(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { wordApp = new Application(); wordApp.Visible = true; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;
                Range searchRange = document.Content; Find find = searchRange.Find; find.ClearFormatting(); find.Text = "5月21日より5日間の"; find.Execute();
                if (!find.Found) { Marshal.ReleaseComObject(find); Marshal.ReleaseComObject(searchRange); return false; }
                Range paraRange = searchRange.Paragraphs[1].Range;
                int paraStart = paraRange.Start;
                int paraEnd = paraRange.End;

                // 「5月21日より5日間の...」段落の先頭付近に行内画像があるか
                bool result = false;

                InlineShapes ils = document.InlineShapes;
                for (int i = 1; i <= ils.Count && !result; i++)
                {
                    InlineShape il = null;
                    try
                    {
                        il = ils[i];
                        int s = il.Range.Start;
                        if (s >= paraStart && s <= Math.Min(paraStart + 5, paraEnd)) result = true;
                    }
                    catch { }
                    finally { if (il != null) Marshal.ReleaseComObject(il); }
                }
                Marshal.ReleaseComObject(ils);

                // 画像が Shape（フロート）として存在し、折り返しが「行内」相当になっているケースも許容
                if (!result)
                {
                    Microsoft.Office.Interop.Word.Shapes shapes = document.Shapes;
                    for (int i = 1; i <= shapes.Count && !result; i++)
                    {
                        Microsoft.Office.Interop.Word.Shape sh = null;
                        try
                        {
                            sh = shapes[i];
                            int anchor = sh.Anchor != null ? sh.Anchor.Start : -1;
                            if (anchor >= paraStart && anchor <= paraEnd)
                            {
                                // 行内に見える設定（厳密には InlineShapes が望ましいが、誤判定回避のため許容）
                                WrapFormat wf = sh.WrapFormat;
                                try { if (wf != null && wf.Type == WdWrapType.wdWrapInline) result = true; }
                                finally { if (wf != null) Marshal.ReleaseComObject(wf); }
                            }
                        }
                        catch { }
                        finally { if (sh != null) Marshal.ReleaseComObject(sh); }
                    }
                    Marshal.ReleaseComObject(shapes);
                }

                Marshal.ReleaseComObject(paraRange);
                Marshal.ReleaseComObject(find);
                Marshal.ReleaseComObject(searchRange);
                return result;
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_5_02(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { wordApp = new Application(); wordApp.Visible = true; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;
                // 「5月21日より5日間の」の先頭の画像の文字列の折り返しが「四角形」= wdWrapSquare
                Range searchRange = document.Content; Find find = searchRange.Find; find.ClearFormatting(); find.Text = "5月21日より5日間の"; find.Execute();
                if (!find.Found) { Marshal.ReleaseComObject(find); Marshal.ReleaseComObject(searchRange); return false; }
                bool result = false;
                foreach (Microsoft.Office.Interop.Word.Shape sh in document.Shapes)
                {
                    try
                    {
                        WrapFormat wf = sh.WrapFormat;
                        if (wf.Type == WdWrapType.wdWrapSquare) { result = true; Marshal.ReleaseComObject(wf); break; }
                        Marshal.ReleaseComObject(wf);
                    }
                    catch { }
                }
                Marshal.ReleaseComObject(find); Marshal.ReleaseComObject(searchRange);
                return result;
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_5_03(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { wordApp = new Application(); wordApp.Visible = true; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;
                // アート効果「水彩：スポンジ」: Word PIA から直接の効果名取得が困難なため、
                // 対象画像（「5月21日より5日間の...」段落付近）で Brightness/Contrast/ColorType がデフォルトから変化していることを「効果あり」として判定
                Range searchRange = document.Content; Find find = searchRange.Find; find.ClearFormatting(); find.Text = "5月21日より5日間の"; find.Execute();
                if (!find.Found) { Marshal.ReleaseComObject(find); Marshal.ReleaseComObject(searchRange); return false; }
                Range paraRange = searchRange.Paragraphs[1].Range;
                int paraStart = paraRange.Start;
                int paraEnd = paraRange.End;

                bool effected = false;

                InlineShapes ils = document.InlineShapes;
                for (int i = 1; i <= ils.Count && !effected; i++)
                {
                    InlineShape il = null;
                    try
                    {
                        il = ils[i];
                        int s = il.Range.Start;
                        if (s < paraStart || s > paraEnd) continue;
                        Microsoft.Office.Interop.Word.PictureFormat pf = il.PictureFormat;
                        try
                        {
                            effected =
                                Math.Abs(pf.Brightness - 0.5f) > 0.01f ||
                                Math.Abs(pf.Contrast - 0.5f) > 0.01f ||
                                pf.ColorType != MsoPictureColorType.msoPictureAutomatic;
                        }
                        finally { if (pf != null) Marshal.ReleaseComObject(pf); }
                    }
                    catch { }
                    finally { if (il != null) Marshal.ReleaseComObject(il); }
                }
                Marshal.ReleaseComObject(ils);

                if (!effected)
                {
                    Microsoft.Office.Interop.Word.Shapes shapes = document.Shapes;
                    for (int i = 1; i <= shapes.Count && !effected; i++)
                    {
                        Microsoft.Office.Interop.Word.Shape sh = null;
                        try
                        {
                            sh = shapes[i];
                            int anchor = sh.Anchor != null ? sh.Anchor.Start : -1;
                            if (anchor < paraStart || anchor > paraEnd) continue;
                            Microsoft.Office.Interop.Word.PictureFormat pf = sh.PictureFormat;
                            try
                            {
                                effected =
                                    Math.Abs(pf.Brightness - 0.5f) > 0.01f ||
                                    Math.Abs(pf.Contrast - 0.5f) > 0.01f ||
                                    pf.ColorType != MsoPictureColorType.msoPictureAutomatic;
                            }
                            finally { if (pf != null) Marshal.ReleaseComObject(pf); }
                        }
                        catch { }
                        finally { if (sh != null) Marshal.ReleaseComObject(sh); }
                    }
                    Marshal.ReleaseComObject(shapes);
                }

                Marshal.ReleaseComObject(paraRange);
                Marshal.ReleaseComObject(find);
                Marshal.ReleaseComObject(searchRange);
                return effected;
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_5_04(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { wordApp = new Application(); wordApp.Visible = true; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;
                // 図の効果「ぼかし25ポイント」= SoftEdge(Radius=25) として判定（Shape 側に反映されることを期待）
                Range searchRange = document.Content; Find find = searchRange.Find; find.ClearFormatting(); find.Text = "5月21日より5日間の"; find.Execute();
                if (!find.Found) { Marshal.ReleaseComObject(find); Marshal.ReleaseComObject(searchRange); return false; }
                Range paraRange = searchRange.Paragraphs[1].Range;
                int paraStart = paraRange.Start;
                int paraEnd = paraRange.End;

                bool ok = false;
                const float expected = 25f;
                const float tolerance = 1.5f;

                Microsoft.Office.Interop.Word.Shapes shapes = document.Shapes;
                for (int i = 1; i <= shapes.Count && !ok; i++)
                {
                    Microsoft.Office.Interop.Word.Shape sh = null;
                    try
                    {
                        sh = shapes[i];
                        int anchor = sh.Anchor != null ? sh.Anchor.Start : -1;
                        if (anchor < paraStart || anchor > paraEnd) continue;
                        Microsoft.Office.Interop.Word.SoftEdgeFormat se = sh.SoftEdge;
                        try { if (se != null && Math.Abs(se.Radius - expected) <= tolerance) ok = true; }
                        finally { if (se != null) Marshal.ReleaseComObject(se); }
                    }
                    catch { }
                    finally { if (sh != null) Marshal.ReleaseComObject(sh); }
                }
                Marshal.ReleaseComObject(shapes);

                Marshal.ReleaseComObject(paraRange);
                Marshal.ReleaseComObject(find);
                Marshal.ReleaseComObject(searchRange);
                return ok;
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_5_05(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { wordApp = new Application(); wordApp.Visible = true; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;
                // 図の効果「面取り ハードエッジ」= ThreeD.BevelTopType が HardEdge の場合に合格
                Range searchRange = document.Content; Find find = searchRange.Find; find.ClearFormatting(); find.Text = "TOEICテスト対策セミナー"; find.Execute();
                if (!find.Found) { Marshal.ReleaseComObject(find); Marshal.ReleaseComObject(searchRange); return false; }
                Range paraRange = searchRange.Paragraphs[1].Range;
                int paraStart = paraRange.Start;
                int paraEnd = paraRange.End;

                bool ok = false;
                Microsoft.Office.Interop.Word.Shapes shapes = document.Shapes;
                for (int i = 1; i <= shapes.Count && !ok; i++)
                {
                    Microsoft.Office.Interop.Word.Shape sh = null;
                    try
                    {
                        sh = shapes[i];
                        int anchor = sh.Anchor != null ? sh.Anchor.Start : -1;
                        if (anchor < paraStart || anchor > paraEnd) continue;
                        Microsoft.Office.Interop.Word.ThreeDFormat td = sh.ThreeD;
                        try { if (td != null && td.BevelTopType == MsoBevelType.msoBevelHardEdge) ok = true; }
                        finally { if (td != null) Marshal.ReleaseComObject(td); }
                    }
                    catch { }
                    finally { if (sh != null) Marshal.ReleaseComObject(sh); }
                }
                Marshal.ReleaseComObject(shapes);

                Marshal.ReleaseComObject(paraRange);
                Marshal.ReleaseComObject(find);
                Marshal.ReleaseComObject(searchRange);
                return ok;
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_5_06(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { wordApp = new Application(); wordApp.Visible = true; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;
                // 「TOEICテスト対策セミナー･･･」のタイトルの右側の画像の代替テキストが「セミナー案内」か
                bool result = false;
                foreach (Microsoft.Office.Interop.Word.Shape sh in document.Shapes)
                {
                    try
                    {
                        string alt = sh.AlternativeText ?? "";
                        if (alt.Trim().Contains("セミナー案内")) { result = true; break; }
                    }
                    catch { }
                }
                return result;
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_5_07(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { wordApp = new Application(); wordApp.Visible = true; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;
                // 「装飾用（Decorative）」が ON かを優先して判定（無ければ代替テキストが空であることを許容）
                Range searchRange = document.Content; Find find = searchRange.Find; find.ClearFormatting(); find.Text = "5月21日より5日間の"; find.Execute();
                if (!find.Found) { Marshal.ReleaseComObject(find); Marshal.ReleaseComObject(searchRange); return false; }
                Range paraRange = searchRange.Paragraphs[1].Range;
                int paraStart = paraRange.Start;
                int paraEnd = paraRange.End;

                bool ok = false;

                // まず Shape を確認
                Microsoft.Office.Interop.Word.Shapes shapes = document.Shapes;
                for (int i = 1; i <= shapes.Count && !ok; i++)
                {
                    Microsoft.Office.Interop.Word.Shape sh = null;
                    try
                    {
                        sh = shapes[i];
                        int anchor = sh.Anchor != null ? sh.Anchor.Start : -1;
                        if (anchor < paraStart || anchor > paraEnd) continue;

                        // Decorative プロパティは環境によっては Interop に露出しないため reflection で取得
                        object decorative = sh.GetType().InvokeMember("Decorative", BindingFlags.GetProperty, null, sh, null);
                        if (decorative is int di && di == -1) ok = true; // msoTrue = -1
                        else if (decorative is bool db && db) ok = true;
                        else if (string.IsNullOrWhiteSpace(sh.AlternativeText)) ok = true;
                    }
                    catch
                    {
                        try { if (sh != null && string.IsNullOrWhiteSpace(sh.AlternativeText)) ok = true; } catch { }
                    }
                    finally { if (sh != null) Marshal.ReleaseComObject(sh); }
                }
                Marshal.ReleaseComObject(shapes);

                // InlineShape も確認（decorative が取れない場合は AltText 空で許容）
                if (!ok)
                {
                    InlineShapes ils = document.InlineShapes;
                    for (int i = 1; i <= ils.Count && !ok; i++)
                    {
                        InlineShape il = null;
                        try
                        {
                            il = ils[i];
                            int s = il.Range.Start;
                            if (s < paraStart || s > paraEnd) continue;
                            object decorative = il.GetType().InvokeMember("Decorative", BindingFlags.GetProperty, null, il, null);
                            if (decorative is int di && di == -1) ok = true;
                            else if (decorative is bool db && db) ok = true;
                            else if (string.IsNullOrWhiteSpace(il.AlternativeText)) ok = true;
                        }
                        catch
                        {
                            try { if (il != null && string.IsNullOrWhiteSpace(il.AlternativeText)) ok = true; } catch { }
                        }
                        finally { if (il != null) Marshal.ReleaseComObject(il); }
                    }
                    Marshal.ReleaseComObject(ils);
                }

                Marshal.ReleaseComObject(paraRange);
                Marshal.ReleaseComObject(find);
                Marshal.ReleaseComObject(searchRange);
                return ok;
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_5_08(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { wordApp = new Application(); wordApp.Visible = true; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;
                // 「背景の削除」は Word PIA から直接取得が難しいため、透明背景が設定されているかで近似判定
                // 文末の画像を優先して確認（最後の InlineShape/Shape）
                bool transparentLike = false;
                InlineShape lastInline = null;
                try
                {
                    InlineShapes ils = document.InlineShapes;
                    if (ils.Count >= 1) lastInline = ils[ils.Count];
                    Marshal.ReleaseComObject(ils);
                }
                catch { }

                if (lastInline != null)
                {
                    try
                    {
                        Microsoft.Office.Interop.Word.PictureFormat pf = lastInline.PictureFormat;
                        try
                        {
                            transparentLike = pf.TransparentBackground == MsoTriState.msoTrue || pf.TransparencyColor != 0;
                        }
                        finally { if (pf != null) Marshal.ReleaseComObject(pf); }
                    }
                    catch { }
                    finally { Marshal.ReleaseComObject(lastInline); }
                }

                if (!transparentLike)
                {
                    try
                    {
                        Microsoft.Office.Interop.Word.Shapes shapes = document.Shapes;
                        if (shapes.Count >= 1)
                        {
                            Microsoft.Office.Interop.Word.Shape sh = shapes[shapes.Count];
                            try
                            {
                                Microsoft.Office.Interop.Word.PictureFormat pf = sh.PictureFormat;
                                try
                                {
                                    transparentLike = pf.TransparentBackground == MsoTriState.msoTrue || pf.TransparencyColor != 0;
                                }
                                finally { if (pf != null) Marshal.ReleaseComObject(pf); }
                            }
                            finally { Marshal.ReleaseComObject(sh); Marshal.ReleaseComObject(shapes); }
                        }
                        else
                        {
                            Marshal.ReleaseComObject(shapes);
                        }
                    }
                    catch { }
                }

                // 背景の削除コマンドが実行されたログがあり、かつ透明背景相当の状態になっている場合のみ正解とする
                bool logOk = LogReader.HasCommandExecuted("PictureBackgroundRemoval");
                return transparentLike && logOk;
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private string GetCurrentWordFilePath()
        {
            Application wordApp = null;
            try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); if (wordApp.ActiveDocument != null) return wordApp.ActiveDocument.FullName; return null; }
            catch (COMException) { return null; }
            finally { if (wordApp != null) Marshal.ReleaseComObject(wordApp); }
        }
    }
}

