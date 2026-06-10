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
                    case 1: // 1-1 スライド追加（4枚目に挿入）
                        // SlidesCount は COM 採点（枚数+1・4枚目挿入）で検証するため免除。図形・位置・TextLength も挿入に伴い不一致になる。
                        flags |= PPValidationExemptFlags.SlidesCount | PPValidationExemptFlags.ShapesCount | PPValidationExemptFlags.ShapePosition | PPValidationExemptFlags.TextLength;
                        break;
                    case 2: // 1-2 スライド非表示
                        // 非表示設定のみのため免除不要
                        break;
                    case 3: // 1-3 スライド5のレイアウト変更＋テキスト入力
                        // プレースホルダーの再配置およびテキスト入力が発生するため、ShapesCount, ShapePosition, TextLength免除が必要
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
                    case 8: // 1-8 サマリーズーム挿入
                        // スライドの追加や、挿入による以降のスライド番号ズレ（すべての座標と文字数の不一致）を回避するため、すべて免除
                        flags |= PPValidationExemptFlags.SlidesCount | PPValidationExemptFlags.ShapesCount | PPValidationExemptFlags.TextLength | PPValidationExemptFlags.ShapePosition;
                        break;
                }
            }
            else if (projectId == 2)
            {
                // セクション操作や画面切り替えのため免除不要
            }
            else if (projectId == 3)
            {
                switch (taskId)
                {
                    case 1: // 3-1 SmartArt挿入
                    case 3: // 3-3 SmartArt変換
                        // オブジェクトの新規追加または置換が発生するため、ShapesCount, TextLength, ShapePosition免除が必要
                        flags |= PPValidationExemptFlags.ShapesCount | PPValidationExemptFlags.TextLength | PPValidationExemptFlags.ShapePosition;
                        break;
                    case 4: // 3-4 スライドズーム挿入
                        // スライドズームオブジェクトが追加されるため、ShapesCount, ShapePosition免除が必要
                        flags |= PPValidationExemptFlags.ShapesCount | PPValidationExemptFlags.ShapePosition;
                        break;
                }
            }
            else if (projectId == 4)
            {
                switch (taskId)
                {
                    case 4: // 4-4 画像のトリミング
                    case 5: // 4-5 画像の配置
                    case 6: // 4-6 順序入れ替え
                        // 画像のサイズや重なり順、座標が変化するため、ShapePosition免除が必要
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
            if (flags.HasFlag(PPValidationExemptFlags.SlidesCount) || flags.HasFlag(PPValidationExemptFlags.ShapesCount))
            {
                flags |= PPValidationExemptFlags.AnimationRemoved;
            }

            return flags;
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
            if (projectId == 3 && taskId == 1) return slideIndex == 5 ? 0 : 0; // 3-1 SmartArt挿入
            if (projectId == 3 && taskId == 3) return slideIndex == 6 ? 0 : 0; // 3-3 SmartArt変換
            if (projectId == 3 && taskId == 4) return slideIndex == 1 ? 2 : 0; // 3-4 スライドズーム
            if (projectId == 5 && taskId == 3) return 0;                       // 5-3 図形変更 (全スライド不変)
            if (projectId == 5 && taskId == 5) return slideIndex == 3 ? -2 : 0; // 5-5 グループ化
            if (projectId == 6 && taskId == 3) return slideIndex == 1 ? 1 : 0; // 6-3 3Dモデル挿入
            if (projectId == 9 && taskId == 1) return slideIndex == 2 ? 0 : 0; // 9-1 グラフ作成 (プレースホルダー内挿入のため不変)

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

            // 3-4: Slide 1 only: allow 0 or +2 (snapshot may be taken after zooms are already inserted during result grading).
            if (projectId == 3 && taskId == 4 && slideIndex == 1)
            {
                return actualDelta == 0 || actualDelta == 2;
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
                return $"不正な図形操作: スライド {slideIndex} で指示外の図形の増減が検知されました（許容: 図形数の変化は 0 または +2、実際の変化: {actualDelta}）";
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
            if (projectId == 3 && (taskId == 1 || taskId == 3 || taskId == 4)) return true; // 3-1, 3-3 SmartArt関連, 3-4 スライドズーム
            if (projectId == 4 && taskId == 6) return true; // 4-6 順序入れ替え (座標は不変)
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
            if (projectId == 4 && taskId == 4) return 1; // 4-4 画像のトリミング
            if (projectId == 4 && taskId == 5) return 1; // 4-5 画像の配置
            if (projectId == 5 && taskId == 4) return 1; // 5-4 図形のサイズ変更
            if (projectId == 6 && taskId == 4) return 1; // 6-4 3Dモデルのサイズ変更
            if (projectId == 9 && taskId == 6) return 1; // 9-6 ハイパーリンク (書き換えによるサイズ変化を許容)
            if (projectId == 11 && taskId == 6) return 1; // 11-6 整列
            return -1;
        }
    }
}
