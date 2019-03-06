using System;
using System.Configuration;
using System.Diagnostics;
using System.Windows.Forms;
using System.Windows.Input;
using Prism.Commands;
using Prism.Mvvm;

namespace InsuranceCompareTool.ViewModels
{
    public class MainWindowViewModel : BindableBase
    { 
        private string mTitle = "Document Comparing Tool";
        private string mSourceFile;
        private string mTargetFile;
        public string SourceFile
        {
            get => mSourceFile;
            set
            {
                mSourceFile = value;
                RaisePropertyChanged();
            }
        }
        public string TargetFile
        {
            get => mTargetFile;
            set
            {
                mTargetFile = value;
                RaisePropertyChanged();
            }
        }
        public string Title
        {
            get { return mTitle; }
            set { SetProperty(ref mTitle, value); }
        }

        public MainWindowViewModel()
        {
            LoadData();
        }

        private void LoadData()
        {
            var cfa = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            this.SourceFile = cfa.AppSettings.Settings[nameof(SourceFile)].Value;
            this.TargetFile = cfa.AppSettings.Settings[nameof(TargetFile)].Value;
        }


        public ICommand SelectSourceFileCommand
        {
            get { return new DelegateCommand(() =>
            {
                var openFileDialog = new Microsoft.Win32.OpenFileDialog()
                {
                    Filter = "Excel Files (*.xlsx)|*.xlsx;*.xls"
                };
                var result = openFileDialog.ShowDialog();
                if (result == true)
                {
                    SourceFile = openFileDialog.FileName;
                }
            });}
        }

        public ICommand SelectTargetFileCommand
        {
            get
            {
                return new DelegateCommand(() =>
                {
                    var openFileDialog = new FolderBrowserDialog();

                    var result = openFileDialog.ShowDialog();
                    if (result == System.Windows.Forms.DialogResult.OK)
                    {
                        TargetFile  = openFileDialog.SelectedPath;
                    }
                });
            }
        }

        public ICommand ExportCommand
        {
            get
            {
                return new DelegateCommand(() =>
                {
                    try
                    {
                        Exporter exporter = new Exporter();

                        string file = exporter.Convert(SourceFile, TargetFile);
                        Process.Start(file);
                    }
                    catch(System.InvalidOperationException e1)
                    {
                        MessageBox.Show("操作不能继续：文件正被使用!", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    catch(Exception ex)
                    { 
                        MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error); 
                    }

                });
            }
        }
        public ICommand SaveCommand
        {
            get
            {
                return new DelegateCommand(() =>
                {
                    var cfa = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                    cfa.AppSettings.Settings[nameof(SourceFile)].Value = SourceFile;
                    cfa.AppSettings.Settings[nameof(TargetFile)].Value = TargetFile;
                    cfa.Save();
                });
            }
        }
    }
}
