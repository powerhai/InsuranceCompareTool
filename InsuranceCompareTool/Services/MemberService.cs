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
    public class MemberService
    {
        private static MemberService Singleton = null;
        private readonly SheetColumn mColID = new SheetColumn() { Title = MemberSheetColumns.ID };
        private readonly SheetColumn mColName = new SheetColumn() { Title = MemberSheetColumns.NAME };
        private readonly SheetColumn mColPosition = new SheetColumn() { Title = MemberSheetColumns.POSITION };
        private readonly SheetColumn mColStatus = new SheetColumn() { Title = MemberSheetColumns.STATUS };
        private readonly SheetColumn mColArea = new SheetColumn() { Title = MemberSheetColumns.AREA }; 
        private readonly SheetColumn mColVirtual = new SheetColumn() { Title = MemberSheetColumns.VIRTUAL };
        private readonly SheetColumn mColReportable = new SheetColumn() { Title = MemberSheetColumns.REPORTABLE };

        public static MemberService CreateInstance()
        {
            if (Singleton == null)
            { 
                Singleton = new MemberService();
            }
            return Singleton;
        }
        private void InitColumnsData(List<SheetColumn> columns)
        {
            foreach (var col in columns)
            {
                switch (col.Title)
                {
                    case MemberSheetColumns.ID:
                    {
                        mColID.Index = col.Index;
                        break;
                    }
                    case MemberSheetColumns.NAME:
                    {
                        mColName.Index = col.Index;
                        break;
                    }
                    case MemberSheetColumns.POSITION:
                    {
                        mColPosition.Index = col.Index;
                        break;
                    }
                    case MemberSheetColumns.STATUS:
                    {
                        mColStatus.Index = col.Index;
                        break;
                    }
                    case MemberSheetColumns.AREA:
                    {
                        mColArea.Index = col.Index;
                        break;
                    }
                    case MemberSheetColumns.VIRTUAL:
                    {
                        mColVirtual.Index = col.Index;
                        break;
                    }
                    case MemberSheetColumns.REPORTABLE:
                    {
                        mColReportable.Index = col.Index;
                        break;
                    }
                }
            }
        }
        public void Load(string file)
        {
            string tempFile = Path.GetTempFileName();
            File.Copy(file, tempFile, true);
            IWorkbook excel = new XSSFWorkbook(tempFile);
            try {
                if (excel.NumberOfSheets <= 0)
                {
                    throw new Exception("职员数据表文件缺少数据");
                }

                mMembers = new List<Member>();

                for (int c = 0; c < excel.NumberOfSheets; c++)
                {
                    ISheet sheet = excel.GetSheetAt(c);
                    SheetReader sheetReader = new SheetReader(sheet);
                    var columns = sheetReader.GetColumns();
                    if (columns.Count <= 0)
                    {
                        throw new Exception("职员数据表文件缺少表头数据");
                    }
                    InitColumnsData(columns);
                    ValidateColumns();
                    
                    for (int i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++)
                    {
                        var row = sheet.GetRow(i);
                        var bill = GetMember(row);
                        if (bill != null)
                        {
                            mMembers.Add(bill);
                        }
                    } 
                }
            }
            catch(Exception ex)
            {
                throw;
            }
            finally
            {
                excel.Close();
            } 
        }

        private void ValidateColumns()
        {
            List<string> mis = new List<string>();
            if (mColID.Index < 0)
                mis.Add(MemberSheetColumns.ID); 
            if (mColName.Index < 0)
                mis.Add(MemberSheetColumns.NAME );
            if (mColPosition.Index < 0)
                mis.Add(MemberSheetColumns.POSITION);
            if (mColStatus.Index < 0)
                mis.Add(MemberSheetColumns.STATUS);
            if (mColArea.Index < 0)
                mis.Add(MemberSheetColumns.AREA);
            
            if (mis.Count > 0)
            {
                string v = "人员表缺少以下列： " + string.Join(", ", mis);
                throw new Exception(v);
            }
        }
        private Member GetMember(IRow row)
        {
            var bill = new Member();
            foreach (var cell in row.Cells)
            {
                cell.SetCellType(CellType.String);
            }

            try
            {
                bill.ID = row.GetCell(mColID.Index)?.StringCellValue.Trim(); 
                bill.Name = row.GetCell(mColName.Index)?.StringCellValue.Trim();
                bill.Position = row.GetCell(mColPosition.Index)?.StringCellValue.Trim();
                bill.Status = row.GetCell(mColStatus.Index)?.StringCellValue.Trim();
                bill.Area = row.GetCell(mColArea.Index)?.StringCellValue.Trim();
                if(mColReportable.Index > 0 && row.GetCell(mColReportable.Index) != null)
                {
                    bill.Reportable = row.GetCell(mColReportable.Index).StringCellValue.Trim().Equals("Yes", StringComparison.CurrentCultureIgnoreCase);
                }
                if (mColVirtual.Index > 0 && row.GetCell(mColVirtual.Index) != null)
                {
                    bill.VirtualMember = row.GetCell(mColVirtual.Index).StringCellValue.Trim().Equals("Yes", StringComparison.CurrentCultureIgnoreCase);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                throw;
            }

            if(string.IsNullOrEmpty(bill.ID))
            {
                return null;
            }

            return bill;
        }

        

        private List<Member> mMembers;
        public List<Member> GetMembers()
        { 
            return mMembers;
        }

        public Member GetMusterOfArea(string area)
        {
            return mMembers.FirstOrDefault(a => a.Area.Equals(area) && a.Position == PositionNames.MUSTER && a.Status == StatusNames.ZAI_ZI);
        }
        public List<Member> GetZaiziMembers()
        {
            return mMembers.Where(a => a.Status == StatusNames.ZAI_ZI).ToList();
        }
        public Member GetMemberBy(string id)
        {
            return mMembers.FirstOrDefault(a => a.ID == id);
        }
        public Member GetMemberByName(string name)
        {
            return mMembers.FirstOrDefault(a => a.Name == name);
        }
    }
}
