using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using InsuranceCompareTool.Domain;
using InsuranceCompareTool.Models;
using log4net;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
namespace InsuranceCompareTool.Services
{
    public class BillLoadHelper
    {
        private readonly ILog mLogger;
        public BillLoadHelper(ILog logger)
        {
            mLogger = logger;
        }
        private readonly List<ColumnDefine> mColumns = new List<ColumnDefine>()
        {
            new ColumnDefine(){Name = BillSheetColumns.PAY_DATE, Type = typeof(DateTime)},
            new ColumnDefine(){Name = BillSheetColumns.PAY_DATE2, Type = typeof(DateTime)},
            new ColumnDefine(){Name = BillSheetColumns.PAY_NO, Type = typeof(int)},
            new ColumnDefine(){Name = BillSheetColumns.BILL_PRICE, Type = typeof(double)},
            new ColumnDefine(){Name = BillSheetColumns.PREVIOUS_PAY_DATE, Type = typeof(DateTime)},

        };

        public DataTable Load(string file)
        {
            string tempFile = Path.GetTempFileName();
            File.Copy(file, tempFile, true);
            IWorkbook excel = new XSSFWorkbook(tempFile);
            mLogger.Debug($"开始加载表格: {file}");
            try
            {
                if (excel.NumberOfSheets <= 0)
                {
                    mLogger.Error($"数据表缺少数据");
                    throw new Exception("保单数据表文件缺少数据");
                }

                ISheet sheet = excel.GetSheetAt(0);

                SheetReader sheetReader = new SheetReader(sheet);
                var columns = sheetReader.GetColumns();
                if (columns.Count <= 0)
                {
                    mLogger.Error($"缺少表头数据");
                    throw new Exception("保单数据表文件缺少表头数据");
                }

                DataTable dt = new DataTable("DataTable");
                var headerRow = sheet.GetRow(sheet.FirstRowNum);
                var firstDataRow = sheet.GetRow(sheet.FirstRowNum + 1);

                for (int i =  headerRow.FirstCellNum ; i < headerRow.LastCellNum; i++)
                {
                    var hcell = headerRow.GetCell(i);
                    var colData  = mColumns.FirstOrDefault(a => a.Name.Equals(hcell.StringCellValue)); 
                    var dcell = firstDataRow.GetCell(i);
                    var cellType = colData == null?  GetType(dcell?.CellType) : colData.Type ;
                    var col = new DataColumn(hcell.StringCellValue, cellType );
                    dt.Columns.Add(col);
                }
                AttachColumns(dt);
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
                                    var col = dt.Columns[j];
                                    if(col.DataType == typeof(DateTime))
                                    {
                                        dataRow[j] = cell.DateCellValue;
                                    }
                                    else 
                                    {
                                        dataRow[j] = cell.NumericCellValue;
                                    }
                                    
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
                mLogger.Debug($"表格加载完成");
                return dt;

            }
            catch (Exception ex)
            {
                mLogger.Error($"出现错误: {ex}");
                throw;
            }
            finally
            {
                excel.Close();
            }
            
        }

        private void AttachColumns(DataTable dt)
        {
            if(dt.Columns.Contains(BillSheetColumns.PREVIOUS_SERVICE))
            {
                if(!dt.Columns.Contains(BillSheetColumns.PREVIOUS_SERVICE_ID))
                {
                    dt.Columns.Add(new DataColumn(BillSheetColumns.PREVIOUS_SERVICE_ID, typeof(String)));
                }
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
