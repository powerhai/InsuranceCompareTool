using System.Collections.Generic;
using InsuranceCompareTool.Models;
using NPOI.SS.UserModel;
namespace InsuranceCompareTool.Services {
    public class SheetReader
    {
        private readonly ISheet mSheet;

        public SheetReader(ISheet sheet)
        {
            mSheet = sheet;
        }

        public List<SheetColumn> GetColumns()
        {
            var columns = new List<SheetColumn>();
            var headerRow = mSheet.GetRow(mSheet.FirstRowNum);
            if(headerRow != null)
            {
                foreach(var cell in headerRow.Cells)
                {
                    cell.SetCellType( CellType.String );
                    var col = new SheetColumn();
                    col.Index = cell.ColumnIndex;
                    col.Title = cell.StringCellValue;
                    columns.Add(col);
                }                
            } 
            return columns;
        }
    }
}