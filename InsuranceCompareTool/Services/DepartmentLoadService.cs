using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using InsuranceCompareTool.Models;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
namespace InsuranceCompareTool.Services
{
    public class DepartmentLoadService
    {
        private static DepartmentLoadService Singleton = null;
        public static DepartmentLoadService CreateInstance()
        {
            if (Singleton == null)
            {
                Singleton = new DepartmentLoadService();
            }
            return Singleton;
        }
        private List<Department> mDepartments;
        public List<Department> GetDepartments()
        {
            return mDepartments;
        }
        public void Load(string file)
        {
            string tempFile = Path.GetTempFileName();
            File.Copy(file, tempFile, true);
            IWorkbook excel = new XSSFWorkbook(tempFile);
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
            InitColumnsData(columns);

            mDepartments = new List<Department>();

            for (int i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++)
            {
                var row = sheet.GetRow(i);
                var department = GetDepartment(row);
                if (department != null)
                {
                    mDepartments.Add(department);
                }
            }
            excel.Close();
        }

        private void InitColumnsData(List<SheetColumn> columns)
        {
            foreach (var col in columns)
            {
                switch (col.Title)
                {
                    case DepartmentSheetColumns.ID:
                    {
                        mColID.Index = col.Index;
                        break;
                    }
                    case DepartmentSheetColumns.NAME:
                    {
                        mColName.Index = col.Index;
                        break;
                    }
                    case DepartmentSheetColumns.AREA:
                    {
                        mColArea.Index = col.Index;
                        break;
                    } 
                }
            }
        }

        private Department GetDepartment(IRow row)
        {
            var bill = new Department();
            foreach (var cell in row.Cells)
            {
                cell.SetCellType(CellType.String);
            }

            try
            {
                bill.ID = row.GetCell(mColID.Index)?.StringCellValue.Trim();
                bill.Area = row.GetCell(mColArea.Index)?.StringCellValue.Trim();
                bill.Name = row.GetCell(mColName.Index)?.StringCellValue.Trim();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                throw;
            } 
            return bill;
        } 
        private readonly SheetColumn mColID = new SheetColumn() { Title = DepartmentSheetColumns.ID };
        private readonly SheetColumn mColName = new SheetColumn() { Title = DepartmentSheetColumns.NAME };
        private readonly SheetColumn mColArea = new SheetColumn() { Title = DepartmentSheetColumns.AREA }; 
    }
}
