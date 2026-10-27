using System;
using System.Collections.Generic;
using System.Linq;

namespace Libraries
{
    /// <summary>
    /// 類題の総合点は参照DLLに委譲する。件数不一致や読み込み失敗では点数を作らない。
    /// </summary>
    public static class PracticeScoreClient
    {
        public const int RequiredResultCount = 35;

        public static bool TryCompute(IList<bool> results, out int total, out string error)
        {
            total = 0;
            error = null;

            if (results == null || results.Count != RequiredResultCount)
            {
                error = "採点結果の件数が一致しないため、総合点を表示できません。";
                return false;
            }

            try
            {
                int computed = MosExcel.PracticeScoring.PracticeTotal.Compute(results.ToArray());
                if (computed < 0 || computed > 1000)
                {
                    error = "総合点を表示できません。";
                    return false;
                }

                total = computed;
                return true;
            }
            catch (Exception)
            {
                error = "総合点を計算できませんでした。";
                return false;
            }
        }
    }
}
