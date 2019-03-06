using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.OpenXmlFormats.Dml;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.Streaming;
using NPOI.XSSF.UserModel;
namespace InsuranceCompareTool
{
    class Exporter
    {
        private const string ColNameA = "上期服务人员";
        private const string ColNameB = "客服专员工号";

        public  string   Convert(string sourceFile, string targetPath)
        {

            var filename = targetPath + "\\" + GetExportFileName() + ".xlsx";
            
            IWorkbook excel = new XSSFWorkbook(sourceFile);
            
            IWorkbook tExcel = new XSSFWorkbook();
            
            var tsheet = tExcel.CreateSheet("main");


            var sheet = excel.GetSheetAt(0);
              
            var cols = GetColumns(sheet);

            CopyRow(sheet.GetRow(sheet.FirstRowNum), tsheet, true );
            var indexA = cols.IndexOf(ColNameA);
            var indexB = cols.IndexOf(ColNameB);

            for (int i = sheet.FirstRowNum + 1 ; i <= sheet.LastRowNum; i++)
            {
                var row = sheet.GetRow(i);
                if(row.Cells.Count <= indexA || row.Cells.Count <= indexB)
                {
                    continue;
                }
                var aValue = row.Cells[indexA]?.StringCellValue;
                var bValue = row.Cells[indexB]?.StringCellValue;

                if(string.IsNullOrEmpty(aValue))
                { 
                    continue;
                }
                if(string.IsNullOrEmpty(bValue))
                    continue;
                if (aValue.Equals(bValue, StringComparison.CurrentCultureIgnoreCase))
                { 
                    continue;
                } 
                CopyRow(row, tsheet);
            }
            var file = new FileStream(filename, FileMode.CreateNew, FileAccess.Write);
            tExcel.Write(file);
            excel.Close();
            tExcel.Close();
            file.Close();
            return filename;
        }
        private void CopyRow(IRow row, ISheet sheet, bool isFirstRow = false)
        {
            var rowIndex = isFirstRow ? sheet.LastRowNum : sheet.LastRowNum + 1;
            var trow = sheet.CreateRow(rowIndex);  
            foreach(ICell cell in row.Cells)
            {
                var ncell =   trow.CreateCell(cell.ColumnIndex); 
                ncell.CellStyle = sheet.Workbook.CreateCellStyle();
                ncell.CellStyle.CloneStyleFrom(cell.CellStyle);
                ncell.SetCellType(cell.CellType);

                switch(cell.CellType)
                {
                    case CellType.String:
                    {
                        ncell.SetCellValue(cell.StringCellValue);
                        break;
                    }
                    case CellType.Boolean:
                    {
                        ncell.SetCellValue( cell.BooleanCellValue);
                        break;
                    }
                    case CellType.Error:
                    {
                        ncell.SetCellValue(cell.ErrorCellValue);
                        break;
                    }
                    case CellType.Numeric:
                    {
                        ncell.SetCellValue(cell.NumericCellValue);
                        break;
                    } 
                    default:
                    {
                        ncell.SetCellValue(cell.StringCellValue);
                        break;
                    } 
                }  
                if (isFirstRow)
                { 
                    sheet.SetColumnWidth(cell.ColumnIndex, row.Sheet.GetColumnWidth(cell.ColumnIndex));
                } 
            } 
            
        }
        public string GetExportFileName()
        {
            return DateTime.Now.ToString("yyyy-MMM-dd hh mm ss"); 
        }

        private List<String> GetColumns(ISheet sheet)
        {
            var items = new List<String>();
            var row = sheet.GetRow(sheet.FirstRowNum); 
            foreach(ICell cell in row.Cells)
            {
                items.Add(cell.StringCellValue);
            } 
            
            return items;

        }
    }
}
