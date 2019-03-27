using System.IO;
using System.Linq;
using InsuranceCompareTool.Models;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
namespace InsuranceCompareTool.Services {
    public class ServiceBillsStatisticsWriter
    {
        public readonly SheetColumn[] Columns =
        {
            new SheetColumn
            {
                Title = MemberSheetColumns.ID,
                Index = 0,
                Width = 900 * 4
            },
            new SheetColumn
            {
                Title = MemberSheetColumns.NAME,
                Index = 1,
                Width = 900 * 4
            },
            new SheetColumn
            {
                Title = StatisticsColumns.COUNT,
                Index = 2,
                Width = 900 * 6,
                Alignment = HorizontalAlignment.Right
            },
            new SheetColumn
            {
                Title = StatisticsColumns.SUM,
                Index = 3,
                Width = 900 * 8,
                Alignment = HorizontalAlignment.Right 
            }

        };
        private readonly XSSFWorkbook mWorkbook;
        private readonly ISheet mSheet;

        public ServiceBillsStatisticsWriter()
        {
            mWorkbook = new XSSFWorkbook();
            mSheet = mWorkbook.CreateSheet("统计");
            WriteHeaderRow(mSheet);
        }
        private void WriteHeaderRow(ISheet sheet )
        {
            var row = sheet.CreateRow(sheet.LastRowNum  ); 
            foreach (var col in Columns)
            {
                var cell = row.CreateCell(col.Index, CellType.String); 
                cell.SetCellValue(col.Title);
                sheet.SetColumnWidth(col.Index, col.Width); 
               
            }
        }
        public void WriteLine(string id, string name, string count,string sum)
        {
            var row = mSheet.CreateRow(mSheet.LastRowNum + 1);
            MakeCells(row);
            row.GetCell(Columns.First(a=>a.Title.Equals(MemberSheetColumns.ID)).Index).SetCellValue(id);
            row.GetCell(Columns.First(a=>a.Title.Equals(MemberSheetColumns.NAME)).Index).SetCellValue(name);
            row.GetCell(Columns.First(a=>a.Title.Equals(StatisticsColumns.COUNT)).Index).SetCellValue(count);
            row.GetCell(Columns.First(a=>a.Title.Equals(StatisticsColumns.SUM)).Index).SetCellValue(sum);
        }
        private void MakeCells(IRow row)
        {
            foreach(var col in Columns)
            {
                var cell = row.CreateCell(col.Index);
                var style = mWorkbook.CreateCellStyle();
                style.Alignment = col.Alignment;
                cell.CellStyle = style;
            }
        }
        public void Save(string filename)
        {
            if (File.Exists(filename))
            {
                File.Delete(filename);
            }
            var file = new FileStream(filename, FileMode.CreateNew, FileAccess.Write);
            mWorkbook.Write(file);
            file.Close();
            mWorkbook?.Close();
        }
    }
}