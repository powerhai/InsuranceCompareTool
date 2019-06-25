using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using InsuranceCompareTool.Models;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
namespace InsuranceCompareTool.Services
{
    public class BillLoadHelper
    {
        public DataTable Load(string file)
        {
            string tempFile = Path.GetTempFileName();
            File.Copy(file, tempFile, true);
            IWorkbook excel = new XSSFWorkbook(tempFile);

            try
            {
                if (excel.NumberOfSheets <= 0)
                {
                    throw new Exception("保单数据表文件缺少数据");
                }

                ISheet sheet = excel.GetSheetAt(0);

                SheetReader sheetReader = new SheetReader(sheet);
                var columns = sheetReader.GetColumns();
                if (columns.Count <= 0)
                {
                    throw new Exception("保单数据表文件缺少表头数据");
                }

                DataTable dt = new DataTable("DataTable");
                var headerRow = sheet.GetRow(sheet.FirstRowNum);
                var firstDataRow = sheet.GetRow(sheet.FirstRowNum + 1);

                for (int i =  headerRow.FirstCellNum ; i < headerRow.LastCellNum; i++)
                {
                    var hcell = headerRow.GetCell(i);
                    var dcell = firstDataRow.GetCell(i);
                    var cellType = GetType(dcell?.CellType);
                    var col = new DataColumn(hcell.StringCellValue, cellType );
                    dt.Columns.Add(col);
                }

                for(int i = sheet.FirstRowNum + 1; i < sheet.LastRowNum; i++)
                {
                    var row = sheet.GetRow(i);
                    DataRow dataRow = dt.NewRow();
                    for(var j = row.FirstCellNum; j < row.LastCellNum; j++)
                    {
                        var cell = row.GetCell(j);
                        if(cell != null)
                        {
                            switch(cell.CellType)
                            {
                                
                                case CellType.Boolean:
                                {
                                    dataRow[j] = cell.BooleanCellValue;
                                    break;
                                }

                                case CellType.Numeric:
                                {
                                    dataRow[j] = cell.NumericCellValue;
                                    break;
                                }
                                case CellType.Error:
                                {
                                    dataRow[j] = cell.ErrorCellValue;
                                    break; 
                                }
                                case CellType.String:
                                default:
                                {
                                    dataRow[j] = cell.StringCellValue;
                                    break;
                                }
                            }
                        } 
                    }

                    dt.Rows.Add(dataRow);
                }

                return dt;

            }
            catch (Exception ex)
            {
                throw;
            }
            finally
            {
                excel.Close();
            }

        }

        private Type GetType(CellType? cellType)
        {
            if(!cellType.HasValue)
            {
                return typeof(String);
            }
            switch(cellType.Value)
            {
                case CellType.Numeric:
                {
                    return typeof(String);
                }
                case CellType.Boolean:
                {
                    return typeof(Boolean);
                }
                case CellType.String:
                default:
                {
                    return typeof(String);
                }
            }
        }
    }
}
