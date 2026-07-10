using System.Collections.Generic;
using System.IO;

namespace MOSExcelMogiApp.Models
{
    /// <summary>
    /// 試験結果を保存するための静的ストレージクラス
    /// </summary>
    public static class ExamResultStorage
    {
        public delegate void ProjectResultsChangedHandler(int projectId);
        /// <summary>
        /// 採点結果が更新されたときに発火（復習採点などでResultWindowを即時更新する用途）
        /// </summary>
        public static event ProjectResultsChangedHandler ResultsChanged;

        // プロジェクトID -> タスクの採点結果（true=正解、false=不正解）
        private static Dictionary<int, List<bool>> _projectResults = new Dictionary<int, List<bool>>();
        // 「採点直後」のスナップショット（復習前の正答率計算用）
        private static Dictionary<int, List<bool>> _initialProjectResults = new Dictionary<int, List<bool>>();

        /// <summary>
        /// プロジェクトの採点結果を保存
        /// </summary>
        public static void SaveProjectResult(int projectId, List<bool> results)
        {
            if (_projectResults == null)
            {
                _projectResults = new Dictionary<int, List<bool>>();
            }
            
            _projectResults[projectId] = new List<bool>(results);
            try
            {
                ResultsChanged?.Invoke(projectId);
            }
            catch
            {
                // UI更新イベントなので例外は握りつぶす
            }
        }

        /// <summary>
        /// 採点直後の全プロジェクト結果をスナップショットとして保存する。
        /// 既にスナップショットが存在する場合は上書きしない（最初の採点結果を保持する）。
        /// </summary>
        public static void SaveInitialResultsIfEmpty(Dictionary<int, List<bool>> allResults)
        {
            if (allResults == null || allResults.Count == 0)
            {
                return;
            }

            if (_initialProjectResults != null && _initialProjectResults.Count > 0)
            {
                // すでに初回採点結果が保存されている場合は何もしない
                return;
            }

            _initialProjectResults = new Dictionary<int, List<bool>>();
            foreach (var kv in allResults)
            {
                _initialProjectResults[kv.Key] = new List<bool>(kv.Value);
            }
        }

        /// <summary>
        /// すべてのプロジェクトの採点結果を取得
        /// </summary>
        public static Dictionary<int, List<bool>> GetAllResults()
        {
            return _projectResults ?? new Dictionary<int, List<bool>>();
        }

        /// <summary>
        /// 採点直後のスナップショット結果を取得する。
        /// </summary>
        public static Dictionary<int, List<bool>> GetInitialResults()
        {
            return _initialProjectResults ?? new Dictionary<int, List<bool>>();
        }

        /// <summary>
        /// 採点結果をクリア
        /// </summary>
        public static void Clear()
        {
            _projectResults?.Clear();
            _initialProjectResults?.Clear();
        }
    }
}






