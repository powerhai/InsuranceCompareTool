using System.Collections.Generic;
using System.Data;
using System.IO;
using InsuranceCompareTool.Models;
namespace InsuranceCompareTool.Core {
    /// <summary>
    /// 导出订单 - 
    /// </summary>
    public class BillExportTypeC
    {
        public void Export(string targetPath, DataTable bills, List<Member> members)
        {
            InitDirAndClearFiles(targetPath);
            var areas = new List<string>();
            foreach(DataRow dr in bills.Rows)
            {
                var area = dr["area"] as string;
                if(!areas.Contains((string)dr["asdfs"]))
                {

                }
            }
        }

        private void InitDirAndClearFiles(string targetPath)
        {
            var dir = new DirectoryInfo(targetPath);
            if (!dir.Exists)
            {
                dir.Create();
            }
            else
            {
 
            }
        }

    }
}