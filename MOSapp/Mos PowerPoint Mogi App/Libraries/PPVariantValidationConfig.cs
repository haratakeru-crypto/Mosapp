using System;

namespace Libraries
{
    /// <summary>
    /// 類題セットのスナップショット免除。教材の <see cref="PPTaskValidationConfig"/> とは別体系。
    /// プロジェクト単位で拡張する（現状はセット1・Project1〜6）。
    /// </summary>
    public static class PPVariantValidationConfig
    {
        /// <summary>類題セットが有効ならその番号、否則 0。</summary>
        public static int ActiveVariantSetNo()
        {
            return PPLogReader.TryReadVariantSetNo(out int setNo) ? setNo : 0;
        }

        public static bool IsVariantActive()
        {
            return ActiveVariantSetNo() > 0;
        }

        /// <summary>
        /// 類題中は類題免除、教材中は教材免除。
        /// </summary>
        public static PPValidationExemptFlags ResolveExemptFlags(int projectId, int taskId)
        {
            int setNo = ActiveVariantSetNo();
            if (setNo > 0)
                return GetExemptFlags(setNo, projectId, taskId);
            return PPTaskValidationConfig.GetExemptFlags(projectId, taskId);
        }

        public static PPValidationExemptFlags ResolveExemptFlags(int variantSetNo, int projectId, int taskId)
        {
            if (variantSetNo > 0)
                return GetExemptFlags(variantSetNo, projectId, taskId);
            return PPTaskValidationConfig.GetExemptFlags(projectId, taskId);
        }

        public static PPValidationExemptFlags GetExemptFlags(int variantSetNo, int projectId, int taskId)
        {
            if (variantSetNo != 1)
                return PPValidationExemptFlags.None;

            PPValidationExemptFlags flags = PPValidationExemptFlags.None;

            if (projectId == 1)
            {
                switch (taskId)
                {
                    case 1: // 箇条書き2段組み — テキストフレームの位置・サイズが変わる
                        flags |= PPValidationExemptFlags.ShapePosition;
                        break;
                    case 2: // アニメーション効果オプションのみ
                        break;
                    case 3: // 吹き出しへテキスト入力
                        flags |= PPValidationExemptFlags.TextLength;
                        break;
                    case 4: // 稲妻→スマイル（図形数は不変）
                        flags |= PPValidationExemptFlags.ShapePosition;
                        break;
                    case 5: // コメント挿入
                        break;
                    case 6: // 背景色変更
                        break;
                }
            }
            else if (projectId == 2)
            {
                switch (taskId)
                {
                    case 1: // 表スライドを1枚挿入 — 枚数・図形・位置・文言がずれる
                        flags |= PPValidationExemptFlags.SlidesCount
                            | PPValidationExemptFlags.ShapesCount
                            | PPValidationExemptFlags.ShapePosition
                            | PPValidationExemptFlags.TextLength;
                        break;
                    case 2: // 画面切り替え（プッシュ）のみ
                        break;
                    case 3: // SmartArt 挿入＋テキスト入力＋不要図形削除
                        flags |= PPValidationExemptFlags.ShapesCount
                            | PPValidationExemptFlags.TextLength
                            | PPValidationExemptFlags.ShapePosition;
                        break;
                    case 4: // 3図形のグループ化 — 子図形の集計が変わる
                        flags |= PPValidationExemptFlags.ShapesCount
                            | PPValidationExemptFlags.TextLength
                            | PPValidationExemptFlags.ShapePosition;
                        break;
                    case 5: // 自動プレゼンテーション設定のみ
                        break;
                    case 6: // 表から集合横棒グラフ作成
                        flags |= PPValidationExemptFlags.ShapesCount
                            | PPValidationExemptFlags.ShapePosition;
                        break;
                }
            }
            else if (projectId == 3)
            {
                switch (taskId)
                {
                    case 1: // 既存スライドのレイアウトを「表スライド」に変更
                        flags |= PPValidationExemptFlags.ShapesCount
                            | PPValidationExemptFlags.ShapePosition
                            | PPValidationExemptFlags.TextLength;
                        break;
                    case 2: // SmartArt の色変更のみ
                        break;
                    case 3: // 右写真を左写真の上端に合わせる（水平は動かさない）
                        flags |= PPValidationExemptFlags.ShapePosition;
                        break;
                    case 4: // ドキュメント検査（コメント・個人情報削除）
                        break;
                    case 5: // URL を「申し込み」表示のハイパーリンクに置換
                        flags |= PPValidationExemptFlags.TextLength;
                        break;
                    case 6: // スライドマスターのテーマを「イオン」に変更
                        flags |= PPValidationExemptFlags.ShapesCount
                            | PPValidationExemptFlags.ShapePosition
                            | PPValidationExemptFlags.TextLength;
                        break;
                }
            }
            else if (projectId == 4)
            {
                switch (taskId)
                {
                    case 1: // 3Dモデルにターンテーブルアニメ — 追加のみのため免除なし
                        break;
                    case 2: // 図形の塗りつぶし・枠線 — 境界ボックスが微変動し得る
                        flags |= PPValidationExemptFlags.ShapePosition;
                        break;
                    case 3: // 目的別スライドショー作成のみ
                        break;
                    case 4: // 表スタイル変更のみ
                        break;
                    case 5: // 動画挿入（位置不問）
                        flags |= PPValidationExemptFlags.ShapesCount
                            | PPValidationExemptFlags.ShapePosition;
                        break;
                    case 6: // 「白紙」レイアウト複製＋プレースホルダー配置
                        flags |= PPValidationExemptFlags.ShapesCount
                            | PPValidationExemptFlags.ShapePosition;
                        break;
                }
            }
            else if (projectId == 5)
            {
                switch (taskId)
                {
                    case 1: // 表紙にスライドズーム2件を挿入して配置
                        flags |= PPValidationExemptFlags.ShapesCount
                            | PPValidationExemptFlags.ShapePosition;
                        break;
                    case 2: // 代替テキストを装飾用にマークのみ
                        break;
                    case 3: // 常に読み取り専用で開く
                        break;
                    case 4: // 表紙以外にフッター＋スライド番号
                        flags |= PPValidationExemptFlags.ShapesCount
                            | PPValidationExemptFlags.ShapePosition
                            | PPValidationExemptFlags.TextLength;
                        break;
                    case 5: // グレースケール表示＋写真の反転グレースケール
                        break;
                    case 6: // 音声のスライド間再生・フェードイン
                        break;
                }
            }
            else if (projectId == 6)
            {
                switch (taskId)
                {
                    case 1: // 特定2スライドのみフッター「手元に残す」
                        flags |= PPValidationExemptFlags.ShapesCount
                            | PPValidationExemptFlags.ShapePosition
                            | PPValidationExemptFlags.TextLength;
                        break;
                    case 2: // スライドサイズ変更（サイズに合わせて調整）
                        flags |= PPValidationExemptFlags.ShapesCount
                            | PPValidationExemptFlags.ShapePosition;
                        break;
                    case 3: // 動画のトリミングのみ
                        break;
                    case 4: // レイアウトの背景グラフィック非表示のみ
                        break;
                    case 5: // 配布資料マスターの日付非表示＋フッター
                        flags |= PPValidationExemptFlags.ShapesCount
                            | PPValidationExemptFlags.TextLength;
                        break;
                }
            }

            // スライド数・図形数が変わる操作ではアニメーションも付随変化し得る
            if (flags.HasFlag(PPValidationExemptFlags.SlidesCount)
                || flags.HasFlag(PPValidationExemptFlags.ShapesCount))
            {
                flags |= PPValidationExemptFlags.AnimationRemoved;
            }

            return flags;
        }

        /// <summary>
        /// TextLength 免除時のスライド別許容デルタ。未設定タスクは無制限（教材と同様の既定）。
        /// </summary>
        public static int GetAllowedTextLengthDelta(int variantSetNo, int projectId, int taskId, int slideIndex)
        {
            if (variantSetNo != 1)
                return int.MaxValue;

            if (projectId == 1 && taskId == 3)
                return slideIndex == 1 ? int.MaxValue : 0; // スライド1の吹き出しのみ

            // 3-5: 対象スライドのみ URL→「申し込み」で文字数が減る（他スライドは厳格）
            if (projectId == 3 && taskId == 5)
                return int.MaxValue;

            return int.MaxValue;
        }

        public static int GetAllowedShapesCountDelta(int variantSetNo, int projectId, int taskId, int slideIndex)
        {
            // 類題は ShapesCount 免除フラグ側で制御。厳密デルタは教材専用。
            if (variantSetNo < 1)
                return int.MaxValue;
            return int.MaxValue;
        }

        /// <summary>
        /// 類題は ShapePosition 免除時に既存図形の移動も許可する（教材の NewShapesOnly を流用しない）。
        /// </summary>
        public static bool IsShapePositionExemptForNewShapesOnly(int variantSetNo, int projectId, int taskId)
        {
            return false;
        }

        /// <summary>
        /// 類題は既存図形の変更上限を設けない（-1）。必要なタスクだけ将来ここで上書きする。
        /// </summary>
        public static int GetAllowedExistingShapePositionChangeCount(int variantSetNo, int projectId, int taskId)
        {
            return -1;
        }

        /// <summary>類題は教材 P1 のスライド単位 ShapePosition 免除を使わない。</summary>
        public static bool UsesPerSlideShapePositionExempt(int variantSetNo, int projectId, int taskId)
        {
            return false;
        }

        public static bool IsShapePositionExemptForSlide(int variantSetNo, int projectId, int taskId, int slideIndex)
        {
            return false;
        }

        public static bool IsAnimationRemovedCheckExemptForSlide(int variantSetNo, int projectId, int taskId, int slideIndex)
        {
            return false;
        }

        /// <summary>類題は教材 P1 の挿入スライド番号マッピングを使わない。</summary>
        public static bool UsesSlideIndexMapping(int variantSetNo, int projectId, int taskId)
        {
            return false;
        }

        /// <summary>類題の ShapesCount 厳密デルタ特例は無し（allowedDelta との一致／MaxValue のみ）。</summary>
        public static bool IsAllowedShapesCountDelta(
            int variantSetNo, int projectId, int taskId, int slideIndex, int allowedDelta, int actualDelta)
        {
            if (allowedDelta == int.MaxValue) return true;
            return actualDelta == allowedDelta;
        }
    }
}
