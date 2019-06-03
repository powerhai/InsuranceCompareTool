using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InsuranceCompareTool.Services
{
    public class ConfigService
    {
        public string MembersFile { get; set; }
        public string RelationFile { get; set; }
        public string TemplateFile { get; set; }
        public string DepartmentsFile { get; set; }

        public ConfigService()
        {
            Load();
        }
        private void Load()
        {
            var cfa = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            SourceFile = cfa.AppSettings.Settings[nameof(SourceFile)].Value;
            TargetFile = cfa.AppSettings.Settings[nameof(TargetFile)].Value;
            MembersFile = cfa.AppSettings.Settings[nameof(MembersFile)].Value;
            DepartmentsFile = cfa.AppSettings.Settings[nameof(DepartmentsFile)].Value;
            RelationFile = cfa.AppSettings.Settings[nameof(RelationFile)].Value;
            TemplateFile = cfa.AppSettings.Settings[nameof(TemplateFile)].Value;
        }
        public string TargetFile { get; set; }
        public string SourceFile { get; set; }
        public void Save()
        {
            var cfa = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            cfa.AppSettings.Settings[nameof(SourceFile)].Value = SourceFile;
            cfa.AppSettings.Settings[nameof(TargetFile)].Value = TargetFile;
            cfa.AppSettings.Settings[nameof(MembersFile)].Value = MembersFile;
            cfa.AppSettings.Settings[nameof(DepartmentsFile)].Value = DepartmentsFile;
            cfa.AppSettings.Settings[nameof(RelationFile)].Value = RelationFile;
            cfa.AppSettings.Settings[nameof(TemplateFile)].Value = TemplateFile;
            cfa.Save();
        }
    }
}
