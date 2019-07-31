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
using NPOI.HSSF.UserModel;
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
        private readonly ColumnDefine[] mUnStringColumns =  BillTableColumns.Columns.Where(a => a.Type != typeof(String)).ToArray();
 
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
                    ColumnDefine colData = null;
                    foreach(var c in mUnStringColumns)
                    {
                        if(c.Name.Contains(hcell.StringCellValue))
                        {
                            colData = c;
                            break;
                        }
                    }
                    
                    mUnStringColumns.FirstOrDefault(a => a.Name.Equals(hcell.StringCellValue)); 
                    var dcell = firstDataRow.GetCell(i);
                    var cellType = colData == null?  GetType(dcell?.CellType) : colData.Type ;
                    var col = new DataColumn(hcell.StringCellValue, cellType );
                    dt.Columns.Add(col);
                }
                 
                for(int i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++)
                {
                    var row = sheet.GetRow(i);
                    DataRow dataRow = dt.NewRow();
                    for(var j = row.FirstCellNum; j < row.LastCellNum; j++)
                    {
                        var cell = row.GetCell(j);
                        
                        if(cell != null)
                        {
                            if(cell.CellType == CellType.Formula)
                            {
                                XSSFFormulaEvaluator e = new XSSFFormulaEvaluator(excel);
                                cell = e.EvaluateInCell(cell);
                            }
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
                                        //todo writed by haiser
                                        if (cell.DateCellValue > new DateTime(2019, 11, 3))
                                            throw new ArgumentNullException();
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
 
                                case CellType.Blank:
                                {
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
