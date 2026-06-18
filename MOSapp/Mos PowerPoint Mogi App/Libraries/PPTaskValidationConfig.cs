using System;
using System.Collections.Generic;

namespace Libraries
{
    [Flags]
    public enum PPValidationExemptFlags
    {
        None = 0,
        ShapesCount = 1,       // 図形追加・削除・グループ化
        TextLength = 2,        // テキスト書き換え
        SlidesCount = 4,       // スライド追加・削除・非表示
        AnimationRemoved = 8,  // アニメーション削除
        ShapePosition = 16,    // 図形の移動・サイズ変更
        All = 31
    }

    public static class PPTaskValidationConfig
    {
        /// <summary>
        /// 指定されたタスクにおいて、チェックを「免除する（無視する）」フラグを返します。
        /// </summary>
                public static PPValidationExemptFlags GetExemptFlags(int projectId, int taskId)
        {
            PPValidationExemptFlags flags = PPValidationExemptFlags.None;

            // プロジェクトごとの破壊的操作免除設定
            if (projectId == 1)
            {
                switch (taskId)
                {
                    case 1: // 1-1 スライド追加（論理4枚目に挿入）
                        // SlidesCount は COM 採点で検証。挿入スライドのみ緩和、他スライドは GetAllowed*Delta で 0 固定。
                        flags |= PPValidationExemptFlags.SlidesCount | PPValidationExemptFlags.ShapesCount | PPValidationExemptFlags.ShapePosition | PPValidationExemptFlags.TextLength;
                        break;
                    case 2: // 1-2 スライド非表示
                        // 非表示設定のみのため免除不要
                        break;
                    case 3: // 1-3 スライド5のレイアウト変更＋テキスト入力
                        // 対象スライド（論理5）のみ緩和。他スライドは GetAllowed*Delta で 0 固定し厳格化する。
                        flags |= PPValidationExemptFlags.ShapesCount | PPValidationExemptFlags.ShapePosition | PPValidationExemptFlags.TextLength;
                        break;
                    case 4: // 1-4 スライド6の箇条書き2段組み
                        // 2段組みによりテキストフレーム等のサイズ・位置が変化するため、ShapePositionを免除
                        flags |= PPValidationExemptFlags.ShapePosition;
                        break;
                    case 5: // 1-5 文字間隔を広げる
                        // 書式変更のみでオブジェクト数は不変のため免除不要
                        break;
                    case 6: // 1-6 スライド8にセクション追加
                        // セクション追加のみのため免除不要
                        break;
                    case 7: // 1-7 スライド1のセクション名変更
                        // セクション名変更のみのため免除不要
                        break;
                    case 8: // 1-8 サマリーズーム挿入（スライド2に1枚追加）
                        // 挿入スライド（2）のみ緩和。他スライドはマッピング後デルタ0。枚数は+1のみ許可。
                        flags |= PPValidationExemptFlags.ShapesCount | PPValidationExemptFlags.TextLength | PPValidationExemptFlags.ShapePosition;
                        break;
                }
            }
            else if (projectId == 2)
            {
                // P2: taskId 1〜8 = P2-1〜P2-8。現行 CheckTask_1_2_01〜08（Legacy は旧2-x 番号のまま）
                // 画面切り替え・アニメーション中心のため、プロジェクト全体で免除フラグは未設定
            }
            else if (projectId == 3)
            {
                switch (taskId)
                {
                    case 1: // P3-1 SmartArt挿入（スライド7）
                    case 3: // P3-3 SmartArt変換（スライド6）
                        flags |= PPValidationExemptFlags.ShapesCount | PPValidationExemptFlags.TextLength | PPValidationExemptFlags.ShapePosition;
                        break;
                    case 4: // P3-4 3Dモデル挿入（スライド1）
                        flags |= PPValidationExemptFlags.ShapesCount | PPValidationExemptFlags.ShapePosition;
                        break;
                    case 5: // P3-5 3Dモデルサイズ・ビュー変更（スライド10）
                        flags |= PPValidationExemptFlags.ShapePosition;
                        break;
                    case 6: // P3-6 スライドズーム挿入（3件）
                        flags |= PPValidationExemptFlags.ShapesCount | PPValidationExemptFlags.ShapePosition;
                        break;
                    case 7: // P3-7 セクションズーム挿入（スライド2）— セクション操作の副作用で文字数・他スライド図形が変わる
                        flags |= PPValidationExemptFlags.ShapesCount | PPValidationExemptFlags.TextLength | PPValidationExemptFlags.ShapePosition;
                        break;
                }
            }
            else if (projectId == 4)
            {
                switch (taskId)
                {
                    case 1: // P4-1 テキスト入力（教育者必見）
                        flags |= PPValidationExemptFlags.TextLength;
                        break;
                    case 5: // P4-5 画像の配置
                    case 6: // P4-6 画像のトリミング
                    case 8: // P4-8 テキストボックス垂直中央配置
                        flags |= PPValidationExemptFlags.ShapePosition;
                        break;
                }
            }
            else if (projectId == 5)
            {
                switch (taskId)
                {
                    case 3: // 5-3 図形変更
                    case 4: // 5-4 図形のサイズ変更
                    case 5: // 5-5 図形のグループ化
                        // 図形の結合や変形、リサイズが発生するため、ShapesCount, ShapePosition免除が必要
                        flags |= PPValidationExemptFlags.ShapesCount | PPValidationExemptFlags.ShapePosition;
                        break;
                }
            }
            else if (projectId == 6)
            {
                switch (taskId)
                {
                    case 3: // 6-3 3Dモデル挿入
                        // 3Dモデルが追加および配置されるため、ShapesCount, ShapePosition免除が必要
                        flags |= PPValidationExemptFlags.ShapesCount | PPValidationExemptFlags.ShapePosition;
                        break;
                    case 4: // 6-4 3Dモデルのサイズ変更
                        // オブジェクトの寸法が変化するため、ShapePosition免除が必要
                        flags |= PPValidationExemptFlags.ShapePosition;
                        break;
                }
            }
            else if (projectId == 7)
            {
                switch (taskId)
                {
                    case 2: // 7-2 スライド再利用
                    case 3: // 7-3 アウトラインからスライド
                        // 大量のスライドやコンテンツが外部から流入するため、SlidesCount, ShapesCount, TextLength, ShapePosition免除が必要
                        flags |= PPValidationExemptFlags.SlidesCount | PPValidationExemptFlags.ShapesCount | PPValidationExemptFlags.TextLength | PPValidationExemptFlags.ShapePosition;
                        break;
                }
            }
            else if (projectId == 8)
            {
                switch (taskId)
                {
                    case 1: // 8-1 ビデオ挿入
                    case 2: // 8-2 ビデオ挿入
                        // ビデオオブジェクトが追加されるため、ShapesCount免除が必要
                        flags |= PPValidationExemptFlags.ShapesCount;
                        break;
                }
            }
            else if (projectId == 9)
            {
                switch (taskId)
                {
                    case 1: // 9-1 表を元にグラフ作成
                        // グラフオブジェクトが新規作成されるため、ShapesCount, ShapePosition免除が必要
                        flags |= PPValidationExemptFlags.ShapesCount | PPValidationExemptFlags.ShapePosition;
                        break;
                    case 4: // 9-4, 5 フッター
                    case 5:
                        // フッターやスライド番号を有効にすると、各スライドにプレースホルダー（図形）が実体化して追加されるため、
                        // ShapesCountおよびShapePositionの免除が必要。またテキスト入力のためTextLengthも免除。
                        flags |= PPValidationExemptFlags.ShapesCount | PPValidationExemptFlags.ShapePosition | PPValidationExemptFlags.TextLength;
                        break;
                    case 6: // 9-6 ハイパーリンク
                        // リンク設定によりテキスト内容と配置が変わるため、TextLength, ShapePosition免除が必要
                        flags |= PPValidationExemptFlags.TextLength | PPValidationExemptFlags.ShapePosition;
                        break;
                    case 7: // 9-7 サイズ変更
                        // レイアウト全体の再計算が発生するため、ShapePosition免除が必要
                        flags |= PPValidationExemptFlags.ShapePosition;
                        break;
                }
            }
            else if (projectId == 10)
            {
                switch (taskId)
                {
                    case 5: // 10-5 テーマ変更
                    case 7: // 10-7 プレースホルダー追加
                        // レイアウト変更やオブジェクト追加が発生するため、ShapesCount, ShapePosition免除が必要
                        flags |= PPValidationExemptFlags.ShapesCount | PPValidationExemptFlags.ShapePosition;
                        break;
                }
            }
            else if (projectId == 11)
            {
                switch (taskId)
                {
                    case 1: // 11-1 サイズ変更
                        // スライドのサイズ変更（サイズに合わせて調整）により、全体レイアウトの再計算やプレースホルダーの再構築が発生し、図形数が増減することがあるため ShapesCount と ShapePosition 免除が必要
                        flags |= PPValidationExemptFlags.ShapesCount | PPValidationExemptFlags.ShapePosition;
                        break;
                    case 6: // 11-6 整列（上下中央揃え）
                        // 図形の移動が発生するため、ShapePosition免除が必要
                        flags |= PPValidationExemptFlags.ShapePosition;
                        break;
                }
            }

            // 従属フラグの自動付与（スライド数や図形数が変化すると、付随するアニメーションも削除/ズレるため除外する）
            // 1-1 / 1-3 / 1-8 は対象スライド以外のアニメーション削除を検知したいため除外しない。
            if ((flags.HasFlag(PPValidationExemptFlags.SlidesCount) || flags.HasFlag(PPValidationExemptFlags.ShapesCount))
                && !(projectId == 1 && (taskId == 1 || taskId == 3 || taskId == 8)))
            {
                flags |= PPValidationExemptFlags.AnimationRemoved;
            }

            return flags;
        }

        /// <summary>プロジェクト1の論理スライド番号を、1-8 挿入後の物理番号に補正する。</summary>
        public static int GetProject1AdjustedSlideNumber(int logicalSlideNumber)
        {
            if (logicalSlideNumber > 1 && PPLogReader.HasTask1_8SummaryZoomExecutedGlobally())
                return logicalSlideNumber + 1;
            return logicalSlideNumber;
        }

        /// <summary>1-8 サマリーズーム挿入位置（物理スライド番号）。</summary>
        public const int Project1Task1_8InsertSlideIndex = 2;

        /// <summary>1-1 スライド挿入位置（論理スライド番号）。</summary>
        public const int Project1Task1_1InsertAtLogical = 4;

        /// <summary>1-1 挿入時の 1-8 オフセット（スライド2以降に +1）。</summary>
        public static int GetProject1Task1_1OffsetAfterSlide1()
        {
            return PPLogReader.HasTask1_8SummaryZoomExecutedGlobally() ? 1 : 0;
        }

        /// <summary>1-1 の操作対象スライド（新規挿入スライド、物理番号）。</summary>
        public static int GetProject1Task1_1InsertSlideIndex()
        {
            return Project1Task1_1InsertAtLogical + GetProject1Task1_1OffsetAfterSlide1();
        }

        /// <summary>1-1 の操作対象スライド（新規挿入スライド）か。</summary>
        public static bool IsProject1Task1_1TargetSlide(int currentSlideIndex)
        {
            return currentSlideIndex == GetProject1Task1_1InsertSlideIndex();
        }

        /// <summary>1-1: スナップショットが挿入前（枚数+1）か。</summary>
        public static bool IsProject1Task1_1InsertApplied(int snapshotSlidesCount, int currentSlidesCount)
        {
            return currentSlidesCount == snapshotSlidesCount + 1;
        }

        /// <summary>1-1: スナップショット取得時点ですでに挿入済み（一括採点など）。</summary>
        public static bool IsProject1Task1_1SnapshotPostInsert(int snapshotSlidesCount, int currentSlidesCount)
        {
            return currentSlidesCount == snapshotSlidesCount;
        }

        /// <summary>1-3 の操作対象スライド（論理5）か。</summary>
        public static bool IsProject1Task1_3TargetSlide(int slideIndex)
        {
            return slideIndex == GetProject1AdjustedSlideNumber(5);
        }

        /// <summary>1-8 の操作対象スライド（新規挿入スライド2）か。</summary>
        public static bool IsProject1Task1_8TargetSlide(int currentSlideIndex)
        {
            return currentSlideIndex == Project1Task1_8InsertSlideIndex;
        }

        /// <summary>スライド挿入によりスナップショット番号と現在番号の対応付けが必要か。</summary>
        public static bool UsesSlideIndexMapping(int projectId, int taskId)
        {
            return projectId == 1 && (taskId == 1 || taskId == 8);
        }

        /// <summary>1-8: スナップショットが挿入前（枚数+1）か挿入後（同数）か。</summary>
        public static bool IsProject1Task1_8InsertApplied(int snapshotSlidesCount, int currentSlidesCount)
        {
            return currentSlidesCount == snapshotSlidesCount + 1;
        }

        /// <summary>1-8: スナップショット取得時点ですでに挿入済み（一括採点など）。</summary>
        public static bool IsProject1Task1_8SnapshotPostInsert(int snapshotSlidesCount, int currentSlidesCount)
        {
            return currentSlidesCount == snapshotSlidesCount;
        }

        /// <summary>スライド枚数がタスク操作として許容されるか。</summary>
        public static bool IsSlidesCountValidForTask(int projectId, int taskId, int snapshotSlidesCount, int currentSlidesCount)
        {
            if (projectId == 1 && (taskId == 1 || taskId == 8))
            {
                if (taskId == 1)
                {
                    return IsProject1Task1_1InsertApplied(snapshotSlidesCount, currentSlidesCount)
                        || IsProject1Task1_1SnapshotPostInsert(snapshotSlidesCount, currentSlidesCount);
                }
                return IsProject1Task1_8InsertApplied(snapshotSlidesCount, currentSlidesCount)
                    || IsProject1Task1_8SnapshotPostInsert(snapshotSlidesCount, currentSlidesCount);
            }
            return currentSlidesCount == snapshotSlidesCount;
        }

        private static int MapProject1Task1_1SnapshotToCurrent(int snapshotSlideIndex)
        {
            int offset = GetProject1Task1_1OffsetAfterSlide1();
            if (snapshotSlideIndex < Project1Task1_1InsertAtLogical)
                return snapshotSlideIndex + (snapshotSlideIndex >= 2 ? offset : 0);
            return snapshotSlideIndex + 1 + offset;
        }

        private static bool TryMapProject1Task1_1CurrentToSnapshot(
            int currentSlideIndex,
            out int snapshotSlideIndex,
            out bool isInsertedSlide)
        {
            int offset = GetProject1Task1_1OffsetAfterSlide1();
            int insertedSlideNum = GetProject1Task1_1InsertSlideIndex();

            if (currentSlideIndex == insertedSlideNum)
            {
                snapshotSlideIndex = 0;
                isInsertedSlide = true;
                return true;
            }

            isInsertedSlide = false;
            if (currentSlideIndex < insertedSlideNum)
            {
                if (currentSlideIndex == 1)
                    snapshotSlideIndex = 1;
                else if (offset == 1)
                    snapshotSlideIndex = currentSlideIndex - 1;
                else
                    snapshotSlideIndex = currentSlideIndex;
                return true;
            }

            snapshotSlideIndex = currentSlideIndex - 1 - offset;
            return true;
        }

        /// <summary>現在スライド番号をスナップショット側番号に対応付ける（1-1: 論理4 / 1-8: スライド2に挿入）。</summary>
        public static bool TryMapCurrentSlideToSnapshot(
            int projectId,
            int taskId,
            int snapshotSlidesCount,
            int currentSlidesCount,
            int currentSlideIndex,
            out int snapshotSlideIndex,
            out bool isInsertedSlide)
        {
            snapshotSlideIndex = currentSlideIndex;
            isInsertedSlide = false;
            if (projectId == 1 && taskId == 1)
            {
                if (IsProject1Task1_1InsertApplied(snapshotSlidesCount, currentSlidesCount))
                    return TryMapProject1Task1_1CurrentToSnapshot(currentSlideIndex, out snapshotSlideIndex, out isInsertedSlide);

                if (IsProject1Task1_1SnapshotPostInsert(snapshotSlidesCount, currentSlidesCount))
                {
                    snapshotSlideIndex = currentSlideIndex;
                    isInsertedSlide = IsProject1Task1_1TargetSlide(currentSlideIndex);
                    return true;
                }

                return false;
            }
            if (projectId == 1 && taskId == 8)
            {
                if (IsProject1Task1_8InsertApplied(snapshotSlidesCount, currentSlidesCount))
                {
                    if (currentSlideIndex == Project1Task1_8InsertSlideIndex)
                    {
                        isInsertedSlide = true;
                        return true;
                    }
                    if (currentSlideIndex < Project1Task1_8InsertSlideIndex)
                    {
                        snapshotSlideIndex = currentSlideIndex;
                        return true;
                    }
                    snapshotSlideIndex = currentSlideIndex - 1;
                    return true;
                }

                if (IsProject1Task1_8SnapshotPostInsert(snapshotSlidesCount, currentSlidesCount))
                {
                    snapshotSlideIndex = currentSlideIndex;
                    isInsertedSlide = currentSlideIndex == Project1Task1_8InsertSlideIndex;
                    return true;
                }

                return false;
            }
            return true;
        }

        /// <summary>スナップショット側番号を現在スライド番号に対応付ける（ShapePositions 逆引き用）。</summary>
        public static int MapSnapshotSlideToCurrent(
            int projectId,
            int taskId,
            int snapshotSlideIndex,
            int snapshotSlidesCount,
            int currentSlidesCount)
        {
            if (projectId == 1 && taskId == 1)
            {
                if (IsProject1Task1_1InsertApplied(snapshotSlidesCount, currentSlidesCount))
                    return MapProject1Task1_1SnapshotToCurrent(snapshotSlideIndex);
                return snapshotSlideIndex;
            }
            if (projectId == 1 && taskId == 8)
            {
                if (IsProject1Task1_8InsertApplied(snapshotSlidesCount, currentSlidesCount))
                {
                    if (snapshotSlideIndex >= Project1Task1_8InsertSlideIndex)
                        return snapshotSlideIndex + 1;
                    return snapshotSlideIndex;
                }
                return snapshotSlideIndex;
            }
            return snapshotSlideIndex;
        }

        /// <summary>ShapePosition 免除をスライド単位で適用するタスクか。</summary>
        public static bool UsesPerSlideShapePositionExempt(int projectId, int taskId)
        {
            return projectId == 1 && (taskId == 1 || taskId == 3 || taskId == 8);
        }

        /// <summary>指定スライドで ShapePosition チェックを免除するか。</summary>
        public static bool IsShapePositionExemptForSlide(int projectId, int taskId, int slideIndex)
        {
            if (projectId == 1 && taskId == 1)
                return IsProject1Task1_1TargetSlide(slideIndex);
            if (projectId == 1 && taskId == 3)
                return IsProject1Task1_3TargetSlide(slideIndex);
            if (projectId == 1 && taskId == 8)
                return IsProject1Task1_8TargetSlide(slideIndex);
            return false;
        }

        /// <summary>対象スライドはレイアウト/挿入でアニメーションが変わるため AnimationRemoved を緩和する。</summary>
        public static bool IsAnimationRemovedCheckExemptForSlide(int projectId, int taskId, int slideIndex)
        {
            if (projectId == 1 && taskId == 1)
                return IsProject1Task1_1TargetSlide(slideIndex);
            if (projectId == 1 && taskId == 3)
                return IsProject1Task1_3TargetSlide(slideIndex);
            if (projectId == 1 && taskId == 8)
                return IsProject1Task1_8TargetSlide(slideIndex);
            return false;
        }

        /// <summary>
        /// ShapesCount免除フラグが有効な場合でも、特定のスライドにおいて許可される「図形数の増減（デルタ）」を返します。
        /// 1 なら +1個（挿入）、-2 なら -2個（グループ化など）を意味します。
        /// 厳格にチェックすべきでないタスクやスライドの場合は int.MaxValue を返すと無制限になります。
        /// </summary>
        public static int GetAllowedShapesCountDelta(int projectId, int taskId, int slideIndex)
        {
            // 厳格に図形数の増減を管理するタスク
            // 目的のスライド以外からの呼び出しに対しては「0（増減禁止）」を返すことで、他スライドへの変更をブロックします。
            if (projectId == 3 && taskId == 1) return slideIndex == 7 ? 0 : 0; // P3-1 SmartArt挿入
            if (projectId == 3 && taskId == 3) return slideIndex == 6 ? 0 : 0; // P3-3 SmartArt変換
            if (projectId == 3 && taskId == 4) return slideIndex == 1 ? 1 : 0; // P3-4 3Dモデル挿入
            if (projectId == 3 && taskId == 6) return 0;                       // P3-6 スライドズーム（タイトル指定スライド）
            if (projectId == 3 && taskId == 7)
            {
                if (slideIndex == 2) return 2; // セクションズーム2件
                if (slideIndex == 1) return 1; // セクション操作の副作用（図形+1）
                return 0;
            }
            if (projectId == 5 && taskId == 3) return 0;                       // 5-3 図形変更 (全スライド不変)
            if (projectId == 5 && taskId == 5) return slideIndex == 3 ? -2 : 0; // 5-5 グループ化
            if (projectId == 6 && taskId == 3) return slideIndex == 1 ? 1 : 0; // 6-3 3Dモデル挿入
            if (projectId == 9 && taskId == 1) return slideIndex == 2 ? 0 : 0; // 9-1 グラフ作成 (プレースホルダー内挿入のため不変)
            if (projectId == 1 && taskId == 1)
                return IsProject1Task1_1TargetSlide(slideIndex) ? int.MaxValue : 0;
            if (projectId == 1 && taskId == 3)
                return IsProject1Task1_3TargetSlide(slideIndex) ? int.MaxValue : 0;
            if (projectId == 1 && taskId == 8)
                return IsProject1Task1_8TargetSlide(slideIndex) ? int.MaxValue : 0;

            // フッター関連やレイアウト変更、スライド追加等のタスクは複雑に変動するため無制限
            return int.MaxValue;
        }

        /// <summary>
        /// ShapesCount 免除フラグが有効な場合の「許容デルタ」判定。
        /// 既定は allowedDelta と actualDelta の厳密一致。
        /// 例外的に、6-3（3Dモデル挿入）はスナップショット取得タイミング等で +1 が 0 として観測される場合があるため、
        /// スライド1に限り 0 または +1 を許容する（削除や大量追加は引き続き不許可）。
        /// 3-4（スライドズーム挿入）は結果表示時にスナップショットが挿入後状態で取られると actualDelta が 0 になるため、
        /// スライド1では 0 または +2 を許容する。
        /// 5-5（グループ化）は結果表示時にスナップショットがグループ化後状態で取られると actualDelta が 0 になるため、
        /// スライド3では 0 または -2 を許容する。
        /// </summary>
        public static bool IsAllowedShapesCountDelta(int projectId, int taskId, int slideIndex, int allowedDelta, int actualDelta)
        {
            if (allowedDelta == int.MaxValue) return true;

            // 6-3: Slide 1 only: allow 0 or +1. Disallow deletions (<0) and bulk additions (>1).
            if (projectId == 6 && taskId == 3 && slideIndex == 1)
            {
                return actualDelta == 0 || actualDelta == 1;
            }

            // P3-4: Slide 1 only: allow 0 or +1 (3Dモデル挿入)。
            if (projectId == 3 && taskId == 4 && slideIndex == 1)
            {
                return actualDelta == 0 || actualDelta == 1;
            }

            // P3-6: allow 0 or +3 on any slide（ズーム追加スライドはタイトル指定のため番号固定不可）。
            if (projectId == 3 && taskId == 6)
            {
                return actualDelta == 0 || actualDelta == 3;
            }

            // P3-7: Slide 2: allow 0 or +2。Slide 1: allow 0 or +1（セクション操作の副作用）。
            if (projectId == 3 && taskId == 7 && slideIndex == 2)
            {
                return actualDelta == 0 || actualDelta == 2;
            }
            if (projectId == 3 && taskId == 7 && slideIndex == 1)
            {
                return actualDelta == 0 || actualDelta == 1;
            }

            // 5-5: Slide 3 only: allow 0 or -2 (snapshot may be taken after group is already created during result grading).
            if (projectId == 5 && taskId == 5 && slideIndex == 3)
            {
                return actualDelta == 0 || actualDelta == -2;
            }

            return actualDelta == allowedDelta;
        }

        /// <summary>
        /// TextLength免除フラグが有効な場合でも、特定のスライドにおいて許可される「テキスト文字数の増減（デルタ）」を返します。
        /// 厳格にチェックすべきでないタスクやスライドの場合は int.MaxValue を返すと無制限になります。
        /// </summary>
        public static int GetAllowedTextLengthDelta(int projectId, int taskId, int slideIndex)
        {
            // 9-6: URLを「お問い合わせ」に変更 (スライド1の63文字のURLが6文字の「お問い合わせ」に置き換わるため -57文字)
            if (projectId == 9 && taskId == 6) return slideIndex == 1 ? -57 : 0;
            if (projectId == 1 && taskId == 1)
                return IsProject1Task1_1TargetSlide(slideIndex) ? int.MaxValue : 0;
            if (projectId == 1 && taskId == 3)
                return IsProject1Task1_3TargetSlide(slideIndex) ? int.MaxValue : 0;
            if (projectId == 1 && taskId == 8)
                return IsProject1Task1_8TargetSlide(slideIndex) ? int.MaxValue : 0;
            if (projectId == 4 && taskId == 1)
                return slideIndex == 1 ? int.MaxValue : 0; // P4-1 スライド1へのテキスト入力

            // 変換、削除、インポートなど文字数が可変なものはチェックを省略
            return int.MaxValue;
        }

        /// <summary>
        /// TextLength 免除時のスライド別デルタ判定。既定は allowedDelta と actualDelta の厳密一致。
        /// 9-6（ハイパーリンク）スライド1は、スナップショット時点で既に置換済みの場合 actualDelta が 0 となるため、
        /// 0 または -57（想定の URL→お問い合わせ）のみ許容する。
        /// </summary>
        public static bool IsAllowedTextLengthDelta(int projectId, int taskId, int slideIndex, int allowedDelta, long actualDelta)
        {
            if (allowedDelta == int.MaxValue) return true;

            if (projectId == 9 && taskId == 6 && slideIndex == 1)
            {
                return actualDelta == 0 || actualDelta == -57;
            }

            return actualDelta == allowedDelta;
        }

        /// <summary>破壊的操作ログ用。6-3は「0 または +1」、3-4 スライド1は「0 または +2」、5-5 スライド3は「0 または -2」、それ以外は従来の期待値表記。</summary>
        public static string FormatDestructiveShapesCountMessage(int slideIndex, int projectId, int taskId, int allowedDelta, int actualDelta)
        {
            if (projectId == 6 && taskId == 3 && slideIndex == 1)
            {
                return $"不正な図形操作: スライド {slideIndex} で指示外の図形の増減が検知されました（許容: 図形数の変化は 0 または +1、実際の変化: {actualDelta}）";
            }
            if (projectId == 3 && taskId == 4 && slideIndex == 1)
            {
                return $"不正な図形操作: スライド {slideIndex} で指示外の図形の増減が検知されました（許容: 図形数の変化は 0 または +1、実際の変化: {actualDelta}）";
            }
            if (projectId == 3 && taskId == 6)
            {
                return $"不正な図形操作: スライド {slideIndex} で指示外の図形の増減が検知されました（許容: 図形数の変化は 0 または +3、実際の変化: {actualDelta}）";
            }
            if (projectId == 3 && taskId == 7 && slideIndex == 2)
            {
                return $"不正な図形操作: スライド {slideIndex} で指示外の図形の増減が検知されました（許容: 図形数の変化は 0 または +2、実際の変化: {actualDelta}）";
            }
            if (projectId == 3 && taskId == 7 && slideIndex == 1)
            {
                return $"不正な図形操作: スライド {slideIndex} で指示外の図形の増減が検知されました（許容: 図形数の変化は 0 または +1、実際の変化: {actualDelta}）";
            }
            if (projectId == 5 && taskId == 5 && slideIndex == 3)
            {
                return $"不正な図形操作: スライド {slideIndex} で指示外の図形の増減が検知されました（許容: 図形数の変化は 0 または -2、実際の変化: {actualDelta}）";
            }
            return $"不正な図形操作: スライド {slideIndex} で指示外の図形の増減が検知されました（期待される変化数: {allowedDelta}、実際: {actualDelta}）";
        }

        /// <summary>破壊的操作ログ用。9-6 スライド1は「0 または -57」、それ以外は従来表記。</summary>
        public static string FormatDestructiveTextLengthMessage(int slideIndex, int projectId, int taskId, int allowedDelta, long actualDelta)
        {
            if (projectId == 9 && taskId == 6 && slideIndex == 1)
            {
                return $"不正なテキスト変更: スライド {slideIndex} で指示外のテキスト変更が検知されました（許容: 文字数の変化は 0 または -57、実際の変化: {actualDelta}）";
            }
            return $"不正なテキスト変更: スライド {slideIndex} で指示外のテキスト変更が検知されました（期待される文字数変化: {allowedDelta}、実際: {actualDelta}）";
        }

        /// <summary>
        /// ShapePosition免除フラグが有効な場合でも、既存図形の位置・サイズ変更を一切許可しない（新規追加図形の免除のみとする）タスクかどうかを返します。
        /// </summary>
        public static bool IsShapePositionExemptForNewShapesOnly(int projectId, int taskId)
        {
            if (projectId == 3 && (taskId == 1 || taskId == 3 || taskId == 4 || taskId == 6)) return true;
            if (projectId == 5 && (taskId == 3 || taskId == 5)) return true; // 5-3 図形変更, 5-5 グループ化
            if (projectId == 6 && taskId == 3) return true; // 6-3 3Dモデル挿入
            if (projectId == 9 && taskId == 1) return true; // 9-1 グラフ作成
            if (projectId == 10 && taskId == 7) return true; // 10-7 プレースホルダー追加
            return false;
        }

        /// <summary>
        /// ShapePosition免除フラグが有効な場合で、既存図形の変更を一部許可するタスクにおける「変更許可上限数」を返します。
        /// 制約を設けない場合は -1 を返します。
        /// </summary>
        public static int GetAllowedExistingShapePositionChangeCount(int projectId, int taskId)
        {
            if (projectId == 4 && taskId == 5) return 1; // P4-5 画像の配置
            if (projectId == 4 && taskId == 6) return 1; // P4-6 画像のトリミング
            if (projectId == 4 && taskId == 8) return 1; // P4-8 垂直中央配置
            if (projectId == 5 && taskId == 4) return 1; // 5-4 図形のサイズ変更
            if (projectId == 3 && taskId == 5) return 1; // P3-5 3Dモデルのサイズ変更
            // P3-7: セクションズームの副作用で複数スライドの既存図形がずれるため上限なし（-1）。図形数は GetAllowedShapesCountDelta で厳格化。
            if (projectId == 6 && taskId == 4) return 1; // 6-4 3Dモデルのサイズ変更
            if (projectId == 9 && taskId == 6) return 1; // 9-6 ハイパーリンク (書き換えによるサイズ変化を許容)
            if (projectId == 11 && taskId == 6) return 1; // 11-6 整列
            return -1;
        }
    }
}
