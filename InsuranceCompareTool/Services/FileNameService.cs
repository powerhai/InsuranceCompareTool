using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DateTime = System.DateTime;
namespace InsuranceCompareTool.Services
{
    public class FileNameService
    {
        private readonly ConfigService mConfigService;
        public FileNameService(ConfigService configService)
        {
            mConfigService = configService;
        }

        public string MainFile
        {
            get { return $"{mConfigService.TargetFile}\\{DateTime.Now.AddMonths(1):yyyy-MM}.xlsx"; }
        }
        public string MonthPath
        {
            get { return $"{mConfigService.TargetFile}\\{DateTime.Now.AddMonths(1):yyyy-MM}"; }
        }
        public string ServicePath
        {
            get { return $"{MonthPath}\\地区收费清单"; }
        }
        public string TemplateFile
        {
            get { return "template.xlsx"; }
        }
    }
}
