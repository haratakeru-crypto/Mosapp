using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Libraries.Group1;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace Libraries
{
    public static class PPSnapshotChecker
    {
        private static readonly string SnapshotFilePath = Path.Combine(Path.GetTempPath(), "mos_ppt_snapshot.txt");

        public static List<string> CompareAndGetErrors(int projectId, int taskId, PPValidationExemptFlags exemptFlags)
        {
            List<string> errors = new List<string>();
            var swTotal = Stopwatch.StartNew();

            if (!File.Exists(SnapshotFilePath))
            {
                System.Diagnostics.Debug.WriteLine("[Validation] Snapshot file not found.");
                PPGradingPerf.Log("PPSnapshotChecker.CompareAndGetErrors", swTotal.ElapsedMilliseconds, "no snapshot file");
                return errors;
            }

            // スナップショットの読み込み
            var swLoad = Stopwatch.StartNew();
            var snapshot = LoadSnapshot();
            PPGradingPerf.Log("PPSnapshotChecker.LoadSnapshot", swLoad.ElapsedMilliseconds, $"P{projectId}-T{taskId}");
            if (snapshot == null)
            {
                PPGradingPerf.Log("PPSnapshotChecker.CompareAndGetErrors", swTotal.ElapsedMilliseconds, "load returned null");
                return errors;
            }

            // 重要：現在採点中のタスクと、スナップショットが取られたタスクが一致する場合のみチェックを行う。
            if (snapshot.ProjectId != projectId || snapshot.TaskId != taskId)
            {
                System.Diagnostics.Debug.WriteLine($"[Validation] ID Mismatch (Skipping): Snapshot={snapshot.ProjectId}-{snapshot.TaskId}, Grading={projectId}-{taskId}");
                PPGradingPerf.Log("PPSnapshotChecker.CompareAndGetErrors", swTotal.ElapsedMilliseconds, $"skipped id mismatch snap={snapshot.ProjectId}-{snapshot.TaskId}");
                return errors;
            }
            System.Diagnostics.Debug.WriteLine($"[Validation] Starting Check for Project{projectId} Task{taskId}");

            // 現在の状態の取得（CloseAllPowerPointPresentations / OpenProjectDocument と採点スレッドの競合を防ぐ）
            var swCom = Stopwatch.StartNew();
            int perfSlideCount = -1;
            PowerPoint.Application pptApp = null;
            lock (PowerPointCheckerCommon.PowerPointComInteropSync)
            {
                try
                {
                    pptApp = (PowerPoint.Application)Marshal.GetActiveObject("PowerPoint.Application");
                if (pptApp == null) return errors;

                PowerPoint.Presentation pres = null;
                try
                {
                    pres = pptApp.ActivePresentation;
                    if (pres == null) return errors;
                    try { perfSlideCount = pres.Slides.Count; } catch { }

                    // 1. スライド数の比較
                    if (!exemptFlags.HasFlag(PPValidationExemptFlags.SlidesCount))
                    {
                        if (pres.Slides.Count != snapshot.SlidesCount)
                        {
                            errors.Add($"SlidesCount changed: expected {snapshot.SlidesCount}, but is {pres.Slides.Count}");
                        }
                        else
                        {
                            // スライド名（順序・構成）の比較
                            for (int i = 1; i <= pres.Slides.Count; i++)
                            {
                                PowerPoint.Slide slide = null;
                                try
                                {
                                    slide = pres.Slides[i];
                                    if (i <= snapshot.SlideNames.Count && slide.Name != snapshot.SlideNames[i - 1])
                                    {
                                        errors.Add($"Slide name mismatch at position {i}: expected {snapshot.SlideNames[i - 1]}, but is {slide.Name}");
                                    }
                                }
                                catch { }
                                finally { if (slide != null) Marshal.ReleaseComObject(slide); }
                            }
                        }
                    }

                    // 2. 図形数・テキスト・アニメーションの比較（スライドごと）
                    long currentTotalTextLength = 0;
                    for (int i = 1; i <= pres.Slides.Count; i++)
                    {
                        PowerPoint.Slide slide = null;
                        PowerPoint.Shapes shapes = null;
                        PowerPoint.Shape shape = null;
                        try
                        {
                            slide = pres.Slides[i];
                            
                            // 図形数
                            int allowedDelta = PPTaskValidationConfig.GetAllowedShapesCountDelta(projectId, taskId, i);
                            bool hasShapesExemptFlag = exemptFlags.HasFlag(PPValidationExemptFlags.ShapesCount);

                            if (!hasShapesExemptFlag || allowedDelta != int.MaxValue)
                            {
                                if (snapshot.ShapesCounts.TryGetValue(i, out int expectedShapesCount))
                                {
                                    int actualCount = slide.Shapes.Count;
                                    int actualDelta = actualCount - expectedShapesCount;
                                    bool deltaOk = !hasShapesExemptFlag
                                        ? (actualDelta == 0)
                                        : PPTaskValidationConfig.IsAllowedShapesCountDelta(projectId, taskId, i, allowedDelta, actualDelta);

                                    if (!deltaOk)
                                    {
                                        if (hasShapesExemptFlag)
                                        {
                                            errors.Add(PPTaskValidationConfig.FormatDestructiveShapesCountMessage(i, projectId, taskId, allowedDelta, actualDelta));
                                        }
                                        else
                                        {
                                            errors.Add($"ShapesCount changed on slide {i}: expected {expectedShapesCount}, but is {actualCount}");
                                        }
                                    }
                                }
                            }

                            shapes = slide.Shapes;
                            int shapesCount = shapes.Count;
                            long slideTextLength = 0;

                            // 図形座標・サイズの比較設定（図形走査はテキスト算出と同一ループで実施）
                            bool exemptFullShapePosition = exemptFlags.HasFlag(PPValidationExemptFlags.ShapePosition);
                            bool onlyNewShapesExempt = PPTaskValidationConfig.IsShapePositionExemptForNewShapesOnly(projectId, taskId);
                            int allowedExistingChangesCount = PPTaskValidationConfig.GetAllowedExistingShapePositionChangeCount(projectId, taskId);
                            bool needsShapePositionCheck = !exemptFullShapePosition || onlyNewShapesExempt || allowedExistingChangesCount >= 0;
                            int changedExistingShapesCount = 0;

                            for (int j = 1; j <= shapesCount; j++)
                            {
                                try
                                {
                                    shape = shapes[j];
                                    // Text length (TextFrame2 を優先的に参照)
                                    try
                                    {
                                        dynamic tf2 = shape.TextFrame2;
                                        if (tf2 != null && (int)tf2.HasText == -1)
                                        {
                                            slideTextLength += tf2.TextRange.Length;
                                        }
                                        else if (shape.HasTextFrame == Office.MsoTriState.msoTrue && shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                                        {
                                            slideTextLength += shape.TextFrame.TextRange.Length;
                                        }
                                    }
                                    catch { }

                                    // 図形座標・サイズの比較
                                    if (needsShapePositionCheck)
                                    {
                                        string key = i + "_" + shape.Id;
                                        if (snapshot.ShapePositions.TryGetValue(key, out var old))
                                        {
                                            float left = (float)shape.Left;
                                            float top = (float)shape.Top;
                                            float w = (float)shape.Width;
                                            float h = (float)shape.Height;

                                            if (Math.Abs(old.Item1 - left) > 0.5f || Math.Abs(old.Item2 - top) > 0.5f ||
                                                Math.Abs(old.Item3 - w) > 0.5f || Math.Abs(old.Item4 - h) > 0.5f)
                                            {
                                                if (exemptFullShapePosition)
                                                {
                                                    if (onlyNewShapesExempt)
                                                    {
                                                        // 既存図形の位置変更が一切許されないタスクでの違反
                                                        errors.Add($"不正な図形変更: スライド {i} で指示外の既存図形(ID:{shape.Id})の位置・サイズが変更されています。");
                                                    }
                                                    else if (allowedExistingChangesCount >= 0)
                                                    {
                                                        changedExistingShapesCount++;
                                                        if (changedExistingShapesCount > allowedExistingChangesCount)
                                                        {
                                                            // 許可された個数以上の既存図形が変更された
                                                            errors.Add($"上限超過の図形変更: スライド {i} で許可された数以上の既存図形(ID:{shape.Id})が変更されています。");
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    errors.Add($"Shape position/size changed on slide {i} (ShapeId {shape.Id})");
                                                }
                                            }
                                        }
                                    }
                                }
                                catch { }
                                finally { if (shape != null) { Marshal.ReleaseComObject(shape); shape = null; } }
                            }

                            // 算出されたスライドのテキスト文字数を全体の合計に加算
                            currentTotalTextLength += slideTextLength;

                            // スライドごとの文字増減の厳格チェック
                            int allowedTextDelta = PPTaskValidationConfig.GetAllowedTextLengthDelta(projectId, taskId, i);
                            bool hasTextExemptFlag = exemptFlags.HasFlag(PPValidationExemptFlags.TextLength);

                            if (!hasTextExemptFlag || allowedTextDelta != int.MaxValue)
                            {
                                if (snapshot.SlideTextLengths.TryGetValue(i, out long expectedSlideTextLength))
                                {
                                    long actualTextDelta = slideTextLength - expectedSlideTextLength;
                                    bool textOk = !hasTextExemptFlag
                                        ? (actualTextDelta == 0)
                                        : PPTaskValidationConfig.IsAllowedTextLengthDelta(projectId, taskId, i, allowedTextDelta, actualTextDelta);

                                    if (!textOk)
                                    {
                                        if (hasTextExemptFlag)
                                        {
                                            errors.Add(PPTaskValidationConfig.FormatDestructiveTextLengthMessage(i, projectId, taskId, allowedTextDelta, actualTextDelta));
                                        }
                                        else
                                        {
                                            // 以前は全体だけだったが、スライド単位でも変化がないかチェック
                                            errors.Add($"TextLength changed on slide {i}: expected {expectedSlideTextLength}, but is {slideTextLength}");
                                        }
                                    }
                                }
                            }

                            // アニメーション数（減少のみ不合格）
                            if (!exemptFlags.HasFlag(PPValidationExemptFlags.AnimationRemoved))
                            {
                                if (snapshot.AnimationCounts.TryGetValue(i, out int expectedAnimCount))
                                {
                                    int currentAnimCount = slide.TimeLine.MainSequence.Count;
                                    if (currentAnimCount < expectedAnimCount)
                                    {
                                        errors.Add($"Animation removed on slide {i}: expected at least {expectedAnimCount}, but is {currentAnimCount}");
                                    }
                                }
                            }

                        }
                        catch { }
                        finally { if (slide != null) Marshal.ReleaseComObject(slide); }
                    }

                    // 合計文字数の比較
                    if (!exemptFlags.HasFlag(PPValidationExemptFlags.TextLength))
                    {
                        System.Diagnostics.Debug.WriteLine($"[Validation] TextLength: Current={currentTotalTextLength}, Snapshot={snapshot.TotalTextLength}");
                        if (currentTotalTextLength != snapshot.TotalTextLength)
                        {
                            errors.Add($"TotalTextLength changed: expected {snapshot.TotalTextLength}, but is {currentTotalTextLength}");
                        }
                    }
                }
                catch { }
                finally { if (pres != null) Marshal.ReleaseComObject(pres); }
                }
                catch { }
                finally
                {
                    if (pptApp != null) Marshal.ReleaseComObject(pptApp);
                    string slideInfo = perfSlideCount >= 0 ? $"slides={perfSlideCount}" : "slides=?";
                    PPGradingPerf.Log("PPSnapshotChecker.comActivePresCompare", swCom.ElapsedMilliseconds, $"P{projectId}-T{taskId} {slideInfo}");
                }
            }

            PPGradingPerf.Log("PPSnapshotChecker.CompareAndGetErrors.total", swTotal.ElapsedMilliseconds, $"P{projectId}-T{taskId} errors={errors.Count}");
            return errors;
        }

        private class SnapshotData
        {
            public int ProjectId;
            public int TaskId;
            public int SlidesCount;
            public List<string> SlideNames = new List<string>();
            public Dictionary<int, int> ShapesCounts = new Dictionary<int, int>();
            public long TotalTextLength;
            public Dictionary<int, long> SlideTextLengths = new Dictionary<int, long>();
            public Dictionary<int, int> AnimationCounts = new Dictionary<int, int>();
            public Dictionary<string, Tuple<float, float, float, float>> ShapePositions = new Dictionary<string, Tuple<float, float, float, float>>();
        }

        private static SnapshotData LoadSnapshot()
        {
            try
            {
                var data = new SnapshotData();
                var lines = File.ReadAllLines(SnapshotFilePath);
                foreach (var line in lines)
                {
                    if (string.IsNullOrEmpty(line)) continue;
                    int colonIndex = line.IndexOf(':');
                    if (colonIndex < 0) continue;

                    string key = line.Substring(0, colonIndex);
                    string value = line.Substring(colonIndex + 1);

                    switch (key)
                    {
                        case "TaskId":
                            var ids = value.Split(',');
                            if (ids.Length == 2)
                            {
                                int.TryParse(ids[0], out data.ProjectId);
                                int.TryParse(ids[1], out data.TaskId);
                            }
                            break;
                        case "SlidesCount":
                            int.TryParse(value, out data.SlidesCount);
                            break;
                        case "SlideNames":
                            data.SlideNames = value.Split(new[] { '|' }, StringSplitOptions.None).ToList();
                            break;
                        case "ShapesCounts":
                            ParsePairs(value, data.ShapesCounts);
                            break;
                        case "TotalTextLength":
                            long.TryParse(value, out data.TotalTextLength);
                            break;
                        case "SlideTextLengths":
                            ParseLongPairs(value, data.SlideTextLengths);
                            break;
                        case "AnimationCounts":
                            ParsePairs(value, data.AnimationCounts);
                            break;
                        case "ShapePositions":
                            ParsePositions(value, data.ShapePositions);
                            break;
                    }
                }
                return data;
            }
            catch { return null; }
        }

        private static void ParsePairs(string value, Dictionary<int, int> dict)
        {
            var pairs = value.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var pair in pairs)
            {
                var parts = pair.Split(':');
                if (parts.Length == 2 && int.TryParse(parts[0], out int k) && int.TryParse(parts[1], out int v))
                {
                    dict[k] = v;
                }
            }
        }

        private static void ParseLongPairs(string value, Dictionary<int, long> dict)
        {
            var pairs = value.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var pair in pairs)
            {
                var parts = pair.Split(':');
                if (parts.Length == 2 && int.TryParse(parts[0], out int k) && long.TryParse(parts[1], out long v))
                {
                    dict[k] = v;
                }
            }
        }

        private static void ParsePositions(string value, Dictionary<string, Tuple<float, float, float, float>> dict)
        {
            var items = value.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var item in items)
            {
                var parts = item.Split(':');
                if (parts.Length == 2)
                {
                    string id = parts[0];
                    var vals = parts[1].Split(',');
                    if (vals.Length == 4)
                    {
                        float.TryParse(vals[0], out float l);
                        float.TryParse(vals[1], out float t);
                        float.TryParse(vals[2], out float w);
                        float.TryParse(vals[3], out float h);
                        dict[id] = Tuple.Create(l, t, w, h);
                    }
                }
            }
        }
    }
}
