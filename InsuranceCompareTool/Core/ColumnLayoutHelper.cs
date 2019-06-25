using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Xml.Serialization;
using InsuranceCompareTool.Models;
using InsuranceCompareTool.Models.Dispatch;
using InsuranceCompareTool.Properties;
using InsuranceCompareTool.ViewModels;
namespace InsuranceCompareTool.Core
{
    public class ProjectCacheHelper
    {
        public void SaveProjects(List<Project> projects)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                XmlSerializer xml = new XmlSerializer(typeof(List<Project>));
                xml.Serialize(ms,projects);
                ms.Seek(0, SeekOrigin.Begin);
                var bytes = new byte[ms.Length];
                ms.Read(bytes, 0, bytes.Length);
                var str = System.Text.Encoding.UTF8.GetString(bytes);
                Settings.Default.Projects = str;
                Settings.Default.Save();
            }
        }
        public List<Project> GetProjects()
        {
            var list = new List<Project>();
            using (MemoryStream ms = new MemoryStream())
            {
                var bytes = System.Text.Encoding.UTF8.GetBytes(Settings.Default.Projects);
                ms.Write(bytes, 0, bytes.Length);
                ms.Seek(0, SeekOrigin.Begin);
                XmlSerializer xml = new XmlSerializer(typeof(List<Project>));
                list  = xml.Deserialize(ms) as List<Project>; 
            }

            return list;
        }
        
    }
    public class ColumnLayoutHelper
    {
        public List<ColumnVisible> GetColumnLayouts()
        {
            var list = new List<ColumnVisible>();
            if (string.IsNullOrEmpty(Settings.Default.ColumnLayouts))
                return list;
            try
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    var bytes = System.Text.Encoding.UTF8.GetBytes(Settings.Default.ColumnLayouts);
                    ms.Write(bytes, 0, bytes.Length);
                    ms.Seek(0, SeekOrigin.Begin);
                    XmlSerializer xml = new XmlSerializer(typeof(List<ColumnVisible>));
                    list   = xml.Deserialize(ms) as List<ColumnVisible>; 
                }
            }
            catch (Exception ex)
            {
                //Console.WriteLine(ex); 
            }

            return list;
        }
        public void SaveColumnLayout(List<ColumnVisible> columns)
        {
            try
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    XmlSerializer xml = new XmlSerializer(typeof(List<ColumnVisible>));
                    xml.Serialize(ms, columns);
                    ms.Seek(0, SeekOrigin.Begin);
                    var bytes = new byte[ms.Length];
                    ms.Read(bytes, 0, bytes.Length);
                    var str = System.Text.Encoding.UTF8.GetString(bytes);
                    Settings.Default.ColumnLayouts = str;
                    Settings.Default.Save();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "保存失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
