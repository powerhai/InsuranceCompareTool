using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using InsuranceCompareTool.Domain;
using InsuranceCompareTool.Properties;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
namespace InsuranceCompareTool.Core {
    /// <summary>
    /// 导出订单 - 用于上传系统，500条， 导出保单号、工号、应缴日期三列
    /// </summary>
    public class BillExportTypeK
    { 
        private const string SHEET_NAME = "Bills";
        public void Export(string targetPath, List<DataRow> dataRows, DataTable dt)
        {
            try
            {
                if (!System.IO.Directory.Exists(targetPath))
                {
                    Directory.CreateDirectory(targetPath);
                }
                var fileNo = 1;
                for (var i = 0; i < dataRows.Count; i += Settings.Default.ExportCount)
                {
                    var rows = new List<DataRow>();
                    var fileName = $"{targetPath}\\上传报表 - {fileNo++}.xls";
                    for (var j = 0; j < Settings.Default.ExportCount; j++)
                    {
                        var index = i + j;
                        if (index >= dataRows.Count)
                        {
                            break;
                        }

                        var row = dataRows[index];
                        rows.Add(row);
                    }

                    WriteBook(fileName, rows, dt);

                }

            }
            catch (Exception e)
            {
                throw new Exception($"写入表时产生错误" + e.Message);
            }
        }

        private void WriteBook(string filename, List<DataRow> rows, DataTable dt)
        {
            IWorkbook workbook = new HSSFWorkbook(); //new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet(SHEET_NAME);
            WriteHeaderRow(sheet, new List<string>() { BillTableColumns.COL_BILL_ID.ExportName, BillTableColumns.COL_CURRENT_SERVICE_ID.ExportName, BillTableColumns.COL_PAY_DATE.ExportName, });
            WriteDataRows(sheet, rows, dt);
            var file = new FileStream(filename, FileMode.CreateNew, FileAccess.Write);
            workbook.Write(file);
            file.Close();
            workbook?.Close();
        }
        private void WriteHeaderRow(ISheet sheet, List<string> columns)
        {
            var row = sheet.CreateRow(0);
            for (var i = 0; i < columns.Count; i++)
            {
                sheet.SetColumnWidth(i, 300* 18);
                var cell = row.CreateCell(i);
                cell.SetCellValue(columns[i]);
            }

        }
        private void WriteDataRows(ISheet sheet, List<DataRow> rows, DataTable dt)
        {
            var colBillId = "";
            var colServiceId = "";
            var colPayDate = "";
            foreach (var col in BillTableColumns.COL_BILL_ID.Name)
            {
                if (dt.Columns.Contains(col))
                    colBillId = col;
            }

            foreach (var col in BillTableColumns.COL_CURRENT_SERVICE_ID.Name)
            {
                if (dt.Columns.Contains(col))
                    colServiceId = col;
            }

            foreach (var col in BillTableColumns.COL_PAY_DATE.Name)
            {
                if (dt.Columns.Contains(col))
                    colPayDate = col;
            }


            var rowNo = 0;
            foreach (var dr in rows)
            {
                rowNo++;
                var row = sheet.CreateRow(sheet.LastRowNum + 1);
                var billId = "";
                var serviceId = "";
                DateTime date;

                if (!dr.IsNull(colBillId))
                {
                    billId = ((string)dr[colBillId]).Trim();
                }
                else
                {
                    throw new Exception($"第{rowNo}行缺少保单号数据");
                }

                if (!dr.IsNull(colServiceId))
                {
                    serviceId = ((string)dr[colServiceId]).Trim();
                }
                else
                {
                    throw new Exception($"第{rowNo}行缺少客服专员工号数据");
                }

                if (!dr.IsNull(colPayDate))
                {
                    date = ((DateTime)dr[colPayDate]);
                }
                else
                {
                    throw new Exception($"第{rowNo}行缺少应缴日期数据");
                }

                {
                    var cell = row.CreateCell(0);
                    cell.SetCellValue(billId);
                }
                {
                    var cell = row.CreateCell(1);
                    cell.SetCellValue(serviceId);
                }

                {
                    var cell = row.CreateCell(2);
                    cell.SetCellValue(serviceId);
                    var cellStyle = sheet.Workbook.CreateCellStyle();
                    var format = sheet.Workbook.CreateDataFormat();
                    cellStyle.DataFormat = format.GetFormat("YYYY-MM-DD");
                    cell.CellStyle = cellStyle;
                    cell.SetCellValue((DateTime)date);
                }



            }
        }
    }
}