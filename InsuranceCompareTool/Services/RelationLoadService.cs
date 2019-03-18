using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using InsuranceCompareTool.Models;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
namespace InsuranceCompareTool.Services {
    public class RelationLoadService
    {
        private static RelationLoadService Singleton = null;
        public static RelationLoadService CreateInstance()
        {
            if(Singleton == null)
            {
                Singleton = new RelationLoadService();
            }

            return Singleton;
        }
        private List<Relation> mRelations;
        public List<Relation> GetRelations()
        {
            return mRelations;
        }
        public void Load(string file)
        { 
            IWorkbook excel = new XSSFWorkbook(file);
            if (excel.NumberOfSheets <= 0)
            {
                throw new Exception("绑定表文件缺少数据");
            }
            ISheet sheet = excel.GetSheetAt(0);
            SheetReader sheetReader = new SheetReader(sheet);
            var columns = sheetReader.GetColumns();
            if (columns.Count <= 0)
            {
                throw new Exception("绑定表文件缺少表头数据");
            }
            InitColumnsData(columns);

            mRelations = new List<Relation>();

            for (int i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++)
            {
                var row = sheet.GetRow(i);
                var department = GetRelation(row);
                if (department != null)
                {
                    mRelations.Add(department);
                }
            }
            excel.Close();
        }

        private Relation GetRelation(IRow row)
        {
            var bill = new Relation();
            foreach (var cell in row.Cells)
            {
                cell.SetCellType(CellType.String);
            }

            try
            { 
                bill.Area = row.GetCell(mColArea.Index)?.StringCellValue.Trim();
                bill.ServiceID = row.GetCell(mColServiceID.Index)?.StringCellValue.Trim();
                bill.ServiceName = row.GetCell(mColServiceName.Index)?.StringCellValue.Trim();
                bill.BindType = row.GetCell(mColBindType.Index)?.StringCellValue.Trim();
                bill.BindString = row.GetCell(mColBindString.Index)?.StringCellValue.Trim();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                throw;
            }
            return bill;
        }
        private void InitColumnsData(List<SheetColumn> columns)
        {
            foreach (var col in columns)
            {
                switch (col.Title)
                {
                    case RelationSheetColumns.SERVICE_NAME:
                    {
                        mColServiceName.Index = col.Index;
                        break;
                    }
                    case RelationSheetColumns.SERVICE_ID:
                    {
                        mColServiceID.Index = col.Index;
                        break;
                    }
                    case RelationSheetColumns.BIND_TYPE:
                    {
                        mColBindType.Index = col.Index;
                        break;
                    }
                    case RelationSheetColumns.BIND_STRING:
                    {
                        mColBindString.Index = col.Index;
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
        private readonly SheetColumn mColServiceName = new SheetColumn(){ Title = RelationSheetColumns.SERVICE_NAME };
        private readonly SheetColumn mColServiceID = new SheetColumn() { Title = RelationSheetColumns.SERVICE_ID };
        private readonly SheetColumn mColBindType = new SheetColumn() { Title = RelationSheetColumns.BIND_TYPE };
        private readonly SheetColumn mColBindString = new SheetColumn() { Title = RelationSheetColumns.BIND_STRING }; 
        private readonly SheetColumn mColArea = new SheetColumn() { Title = DepartmentSheetColumns.AREA };
    }
}