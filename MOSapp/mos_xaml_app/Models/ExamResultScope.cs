namespace MOSExcelMogiApp.Models
{
    /// <summary>教材と類題セットで採点結果を分けるための試験コンテキスト。</summary>
    public sealed class ExamResultScope
    {
        private ExamResultScope(bool isVariant, int variantSetNo)
        {
            IsVariant = isVariant;
            VariantSetNo = isVariant ? variantSetNo : 0;
        }

        public bool IsVariant { get; }

        public int VariantSetNo { get; }

        public string Key => IsVariant ? "v" + VariantSetNo : "t";

        public static ExamResultScope Textbook { get; } = new ExamResultScope(false, 0);

        public static ExamResultScope Variant(int variantSetNo)
        {
            int setNo = variantSetNo < 1 ? 1 : variantSetNo;
            return new ExamResultScope(true, setNo);
        }
    }
}
