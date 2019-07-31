using System;
using System.Collections.Generic;
using System.IO;
using InsuranceCompareTool.Models;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
namespace InsuranceCompareTool.Core {
    /// <summary>
    /// 导出用户保单统计
    /// </summary>
    public class MemberExportHelper
    {
        private const string SHEET_NAME = "保单统计";
        private static readonly SheetColumn COL_ID = new SheetColumn(){ Title = "工号" , Index = 0, Alignment =  HorizontalAlignment.Left, CellType = CellType.String, Width = 2800};
        private static readonly SheetColumn COL_NAME = new SheetColumn()
        {
            Title = "姓名",
            Index = 0,
            Alignment = HorizontalAlignment.Left,
            CellType = CellType.String,
            Width = 2800
        };

        private static readonly SheetColumn COL_COUNT = new SheetColumn()
        {
            Title = "数量",
            Index = 0,
            Alignment = HorizontalAlignment.Right,
            CellType = CellType.String,
            Width = 2800
        };
        private static readonly SheetColumn COL_PRICE = new SheetColumn()
        {
            Title = "金额",
            Index = 0,
            Alignment = HorizontalAlignment.Right,
            CellType = CellType.String,
            Width = 5500
        };

        private static readonly SheetColumn COL_AREA = new SheetColumn()
        {
            Title = "地区",
            Index = 0,
            Alignment = HorizontalAlignment.Center,
            CellType = CellType.String,
            Width = 1800
        };


        private  List<SheetColumn> mColumns = new List<SheetColumn>()
        {
            COL_ID,
            COL_NAME,
            COL_COUNT,
            COL_PRICE,
            COL_AREA 
        };

        public void Export(string targetPath, List<MemberA> members)
        {
            try
            {
                if (!System.IO.Directory.Exists(targetPath))
                {
                    Directory.CreateDirectory(targetPath);
                   
                }
                else
                {
                    var files = Directory.GetFiles(targetPath);
                    foreach(var file in files)
                    {
                        File.Delete(file);
                    }
                }
 
                var fileName = $"{targetPath}\\保单统计.xlsx";

                WriteBook(fileName, members);
            }
            catch (Exception e)
            {
                throw new Exception($"写入表时产生错误" + e.Message);
            }


        }

        private void WriteBook(string filename, List<MemberA> members)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet(SHEET_NAME);
            WriteHeaderRow(sheet);
            WriteDataRows(sheet, members);

            var file = new FileStream(filename, FileMode.CreateNew, FileAccess.Write);
            workbook.Write(file);
            file.Close();
            workbook?.Close();
        }

        private void WriteHeaderRow(ISheet sheet )
        {
            var row = sheet.CreateRow(0);

            for (var i = 0; i < mColumns.Count; i++)
            {
                var col = mColumns[i];
                var cell = row.CreateCell(i);
                cell.SetCellValue(col.Title);
                sheet.SetColumnWidth(i, col.Width);
            } 
        }

        private void WriteDataRows(ISheet sheet, List<MemberA> members)
        {
            foreach (var mem  in members)
            {
                var row = sheet.CreateRow(sheet.LastRowNum + 1);
                row.CreateCell(0).SetCellValue(mem.ID);
                row.CreateCell(1).SetCellValue(mem.Name);
                row.CreateCell(2).SetCellValue(mem.Count);
                row.CreateCell(3).SetCellValue(mem.Price);
                row.CreateCell(4).SetCellValue(mem.Area); 
                 
            }
        }

    }
}