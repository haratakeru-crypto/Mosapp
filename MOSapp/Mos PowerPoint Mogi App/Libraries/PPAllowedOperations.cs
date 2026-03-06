using System;
using System.Collections.Generic;

namespace Libraries
{
    /// <summary>
    /// タスクごとの「許可する操作タイプ」と「座標変化（画像・プレースホルダー移動）を許可するか」を定義する。
    /// 採点時に PPLogReader と組み合わせて、余計な操作・不正な座標変化を検出する。
    /// </summary>
    public static class PPAllowedOperations
    {
        /// <summary>ログに記録される操作タイプ: Ribbon コマンド実行。</summary>
        public const string OpTypeRibbonCommand = "RibbonCommand";

        /// <summary>ログに記録される操作タイプ: 図形・プレースホルダーの座標変化。</summary>
        public const string OpTypeShapePositionChange = "ShapePositionChange";

        private static readonly HashSet<string> AllowedRibbonOnly = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            OpTypeRibbonCommand
        };

        private static readonly HashSet<string> AllowedRibbonAndShapeMove = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            OpTypeRibbonCommand,
            OpTypeShapePositionChange
        };

        /// <summary>
        /// 指定タスクで許可する操作タイプの集合。呼び出し元で変更しないこと。
        /// </summary>
        public static HashSet<string> GetAllowedOperationTypes(int projectId, int taskId)
        {
            if (IsShapePositionChangeAllowed(projectId, taskId))
                return new HashSet<string>(AllowedRibbonAndShapeMove, StringComparer.OrdinalIgnoreCase);
            return new HashSet<string>(AllowedRibbonOnly, StringComparer.OrdinalIgnoreCase);
        }

        /// <summary>
        /// 指定タスクで画像・プレースホルダーの座標変化（移動）を許可するか。
        /// 問題文で「配置」「上揃え」「トリミング」等の指示があるタスク、またはスライド追加・レイアウト変更等で
        /// 正答操作により座標が変わるタスクは true。
        /// </summary>
        public static bool IsShapePositionChangeAllowed(int projectId, int taskId)
        {
            switch (projectId)
            {
                case 1:
                    return taskId == 1 || taskId == 2 || taskId == 4 || taskId == 5; // 1-1 スライド追加, 1-2 複製, 1-4 削除, 1-5 レイアウト変更
                case 3:
                    return taskId == 1 || taskId == 3 || taskId == 4; // 3-1 SmartArt挿入, 3-3 箇条書き→SmartArt, 3-4 スライドズーム配置
                case 4:
                    return taskId == 4 || taskId == 5 || taskId == 6; // 4-4 トリミング, 4-5 上揃え, 4-6 重ね順
                case 5:
                    return taskId == 3 || taskId == 4 || taskId == 5; // 5-3 図形変更, 5-4 幅合わせ, 5-5 グループ化
                case 6:
                    return taskId == 3; // 6-3: 3Dモデル幅・中央の枠に
                case 7:
                    return taskId == 2 || taskId == 3; // 7-2 スライド再利用, 7-3 アウトラインから挿入（スライド追加で番号ずれ）
                case 9:
                    return taskId == 1 || taskId == 7; // 9-1 表→グラフ, 9-7 スライドサイズ変更＋サイズに合わせて調整
                case 10:
                    return taskId == 5 || taskId == 7; // 10-5 マスターのテーマ変更, 10-7 レイアウト複製・プレースホルダー配置
                case 11:
                    return taskId == 1 || taskId == 6; // 11-1 スライドサイズ16:10＋調整, 11-6 上下中央揃え
                default:
                    return false;
            }
        }
    }
}
