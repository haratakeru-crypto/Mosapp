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
                    case 1: // 1-1 スライド追加
                        // 新しいスライドが追加されるため、SlidesCount, ShapesCount, ShapePosition免除が必要
                        flags |= PPValidationExemptFlags.SlidesCount | PPValidationExemptFlags.ShapesCount | PPValidationExemptFlags.ShapePosition;
                        break;
                    case 2: // 1-2 スライド複製
                    case 4: // 1-4 スライド削除
                        // スライドの構成が大きく変わるため、SlidesCount, ShapesCount, TextLength, ShapePosition免除が必要
                        flags |= PPValidationExemptFlags.SlidesCount | PPValidationExemptFlags.ShapesCount | PPValidationExemptFlags.TextLength | PPValidationExemptFlags.ShapePosition;
                        break;
                    case 3: // 1-3 スライド非表示
                        // 非表示設定のみで物理的な変化はないため免除不要
                        break;
                    case 5: // 1-5 レイアウト変更
                        // プレースホルダーの再配置が発生するため、ShapesCount, ShapePosition免除が必要
                        flags |= PPValidationExemptFlags.ShapesCount | PPValidationExemptFlags.ShapePosition;
                        break;
                    case 6: // 1-6 箇条書きを2段組みに設定
                        // 書式変更のみでオブジェクト数は不変のため免除不要
                        break;
                    case 7: // 1-7 吹き出しへのテキスト入力
                        // 図形内に文字を書き込むため、TextLength免除が必要
                        flags |= PPValidationExemptFlags.TextLength;
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
    }
}
