using NPOI.SS.UserModel;
namespace InsuranceCompareTool.Models {
    public class SheetTemplate
    {
        public ISheet Sheet { get; set; }
        private IRow mTitleRow;
        public IRow TitleRow
        {
            get { return mTitleRow?? (mTitleRow = Sheet.GetRow(Sheet.FirstRowNum)); }
        }
        private IRow mHeadRow;
        public IRow HeadRow
        {
            get { return mHeadRow ?? (mHeadRow = Sheet.GetRow(Sheet.FirstRowNum + 1)); }
        }
        private IRow mDataRow;
        public IRow DataRow
        {
            get
            {
                return mDataRow?? (mDataRow = Sheet.GetRow( Sheet.FirstRowNum + 2 )); 
            }
        }
        private IRow mSumRow;
        public IRow SumRow
        {
            get { return mSumRow??  (mSumRow = Sheet.GetRow(Sheet.FirstRowNum + 4 )); }
        }
    }

    public class ExCellStyle 
    {
        public ICellStyle CellStyle { get; set; } 
    }
}