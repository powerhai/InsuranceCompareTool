﻿using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using InsuranceCompareTool.Properties;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
namespace InsuranceCompareTool.Core
{
    /// <summary>
    /// 导出订单 - 用于上传系统，500条
    /// </summary>
    public class BillExportTypeD
    {
        private const string SHEET_NAME = "Bills";
        public void Export(string targetPath, DataTable bills, List<string> columns)
        { 
            try
            {
                if(!System.IO.Directory.Exists(targetPath))
                {
                    Directory.CreateDirectory(targetPath);
                }
                var fileNo = 1;
                for(var i = 0; i < bills.Rows.Count; i += Settings.Default.ExportCount)
                {
                    var rows = new List<DataRow>();
                    var fileName = $"{targetPath}\\上传报表 - {fileNo++}.xlsx";
                    for(var j = 0; j < Settings.Default.ExportCount; j++)
                    {
                        var index = i + j;
                        if(index >= bills.Rows.Count)
                        {
                            break;
                        }
                        
                        var row = bills.Rows[index];
                        rows.Add(row);
                    }
                    
                    WriteBook(fileName,columns, bills.Columns, rows);
                    
                }
                
            }
            catch (Exception e)
            {
                throw new Exception($"写入表时产生错误" + e.Message);
            }
        }

        private void WriteBook(string filename, List<string> columns, DataColumnCollection billsColumns, List<DataRow> rows)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet(SHEET_NAME);
            WriteHeaderRow(sheet, columns);
            WriteDataRows(sheet, columns, billsColumns, rows);
            var file = new FileStream(filename, FileMode.CreateNew, FileAccess.Write);
            workbook.Write(file);
            file.Close();
            workbook?.Close();
        }
        private void WriteHeaderRow(ISheet sheet, List<string> columns)
        {
            var row = sheet.CreateRow(0);
            for(var i = 0; i < columns.Count; i++)
            {
                var cell = row.CreateCell(i);
                cell.SetCellValue(columns[i]);
            }
            
        }
        private void WriteDataRows(ISheet sheet, List<string> columnNames, DataColumnCollection billsColumns, List<DataRow> rows)
        {
            foreach(var dr in rows)
            {
                var row = sheet.CreateRow(sheet.LastRowNum + 1);
                for(int i = 0; i < columnNames.Count; i++)
                {
                    var colName = columnNames[i];
                    var cellValue = dr[colName];
                    var cell = row.CreateCell(i);
                    var dataType = billsColumns[colName].DataType;
                    if(dataType == typeof(double))
                    {
                        cell.SetCellValue((double)cellValue);
                    }
                    else if(dataType == typeof(Boolean))
                    {
                        cell.SetCellValue((bool) cellValue);
                    }
                    else if(dataType == typeof(DBNull))
                    {

                    }
                    else 
                    {
                        if(cellValue.GetType() != typeof(System.DBNull))
                        { 
                            cell.SetCellValue((string)cellValue);
                        } 
                    }
                }
            }
        }
    }
}
