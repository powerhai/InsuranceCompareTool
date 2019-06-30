using Prism.Commands;
using Prism.Mvvm;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Input;
using InsuranceCompareTool.Models;
using InsuranceCompareTool.Properties;
using InsuranceCompareTool.Services;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;
namespace InsuranceCompareTool.ViewModels
{
	public class SettingsViewViewModel : TabViewModelBase
    {
        private readonly ConfigService mConfigService;
        private string mMembersFile;
        private string mSourceFile;
        private string mTargetFile; 
        private string mDepartmentsFile;
        private string mRelationFile; 
        private string mTemplateFile;
        public override string Title { get; set; } = "Settings";
        

        public string SourceFile
        {
            get => mSourceFile;
            set
            {
                mSourceFile = value;
                RaisePropertyChanged(); 
            }
        }
        public string MembersFile
        {
            get => mMembersFile;
            set
            {
                mMembersFile = value;
                RaisePropertyChanged(); 
            }
        }
        public string RelationFile
        {
            get => mRelationFile;
            set
            {
                SetProperty(ref mRelationFile, value); 
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


        public string TemplateFile
        {
            get => mTemplateFile;
            set
            {
                SetProperty(ref mTemplateFile, value); 
            }
        }
        public string DepartmentsFile
        {
            get => mDepartmentsFile;
            set
            {
                SetProperty(ref mDepartmentsFile, value); 
            }
        }

        public SettingsViewViewModel(ConfigService configService)
        {
            mConfigService = configService;
            LoadData();
        }

        public ICommand SelectMembersFileCommand
        {
            get
            {
                return new DelegateCommand(() =>
                {
                    var openFileDialog = new OpenFileDialog
                    {
                        Filter = "Excel Files (*.xlsx)|*.xlsx",

                    };
                    if(!string.IsNullOrEmpty(MembersFile))
                    {
                        var file = new FileInfo(MembersFile);
                        if(file?.Directory?.FullName != null)
                        {
                            openFileDialog.InitialDirectory = file.Directory.FullName;
                        } 
                    }

                    var result = openFileDialog.ShowDialog();
                    if (result == true)
                        MembersFile = openFileDialog.FileName;
                });
            }
        }
        public ICommand SelectDepartmentsFileCommand
        {
            get
            {
                return new DelegateCommand(() =>
                {
                    var openFileDialog = new OpenFileDialog
                    {
                        Filter = "Excel Files (*.xlsx)|*.xlsx"
                    };
                    if (!string.IsNullOrEmpty(DepartmentsFile))
                    {
                        var file = new FileInfo(DepartmentsFile);
                        if (file?.Directory?.FullName != null)
                        {
                            openFileDialog.InitialDirectory = file.Directory.FullName;
                        }
                    }

                    var result = openFileDialog.ShowDialog();
                    if (result == true)
                        DepartmentsFile = openFileDialog.FileName;
                });
            }
        }
        public ICommand SelectTemplateFileCommand
        {
            get
            {
                return new DelegateCommand(() =>
                {
                    var openFileDialog = new OpenFileDialog
                    {
                        Filter = "Excel Files (*.xlsx)|*.xlsx"
                    };

                    if (!string.IsNullOrEmpty(TemplateFile))
                    {
                        var file = new FileInfo(TemplateFile);
                        if (file?.Directory?.FullName != null)
                        {
                            openFileDialog.InitialDirectory = file.Directory.FullName;
                        }
                    }

                    var result = openFileDialog.ShowDialog();
                    if (result == true)
                        TemplateFile = openFileDialog.FileName;

                });
            }
        }


        public ICommand SelectTargetFileCommand
        {
            get
            {
                return new DelegateCommand(() =>
                {
                    var openFileDialog = new FolderBrowserDialog(); 
                    var result = openFileDialog.ShowDialog();
                    if (result == DialogResult.OK)
                        TargetFile = openFileDialog.SelectedPath;
                });
            }
        }


        public ICommand OpenWorkPositionCommand
        {
            get
            {
                return new DelegateCommand(() =>
                {
                    try
                    {
                        if (!string.IsNullOrEmpty(TargetFile))
                            Process.Start(TargetFile);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                });
            }
        }

        public ICommand OpenTemplateCommand
        {
            get
            {
                return new DelegateCommand(() =>
                {
                    try
                    {
                        if (!string.IsNullOrEmpty(TemplateFile))
                            Process.Start(TemplateFile);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                });
            }
        }
        public ICommand OpenMemberFileCommand
        {
            get
            {
                return new DelegateCommand(() =>
                {
                    try
                    {
                        if (!string.IsNullOrEmpty(MembersFile))
                            Process.Start(MembersFile);
                    }
                    catch (InvalidOperationException e1)
                    {
                        MessageBox.Show("操作不能继续：文件正被使用!", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                });
            }
        }
        public ICommand OpenDepartmentsFileCommand
        {
            get
            {
                return new DelegateCommand(() =>
                {
                    try
                    {
                        if (!string.IsNullOrEmpty(DepartmentsFile))
                            Process.Start(DepartmentsFile);
                    }
                    catch (InvalidOperationException e1)
                    {
                        MessageBox.Show("操作不能继续：文件正被使用!", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                });
            }
        }

        public ICommand OpenRelationFileFileCommand
        {
            get
            {
                return new DelegateCommand(() =>
                {
                    try
                    {
                        if (!string.IsNullOrEmpty(RelationFile))
                            Process.Start(RelationFile);
                    }
                    catch (InvalidOperationException e1)
                    {
                        MessageBox.Show("操作不能继续：文件正被使用!", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                });
            }
        }

        public ICommand SelectRelationFileCommand
        {
            get
            {
                return new DelegateCommand(() =>
                {
                    var openFileDialog = new OpenFileDialog
                    {
                        Filter = "Excel Files (*.xlsx)|*.xlsx"
                    };
                    var result = openFileDialog.ShowDialog();
                    if (result == true)
                        RelationFile = openFileDialog.FileName;
                });
            }
        }

        private bool mIsLoaded = false;
        private void LoadData()
        {
            if(mIsLoaded == true)
                return;
            this.DepartmentsFile = Settings.Default.DepartmentsFile;
            this.MembersFile = Settings.Default.MembersFile;
            this.RelationFile = Settings.Default.RelationFile;
            this.TemplateFile = Settings.Default.TemplateFile;
            this.TargetFile = Settings.Default.WorkPath; 
            mIsLoaded = true;
        }
 
 
        public override void Leave()
        {
            if (mIsLoaded == false)
                return;
            Settings.Default.WorkPath = TargetFile;
            Settings.Default.MembersFile = MembersFile;
            Settings.Default.TemplateFile = TemplateFile;
            Settings.Default.RelationFile = RelationFile;
            Settings.Default.DepartmentsFile = DepartmentsFile;
         
        }
        public override void Close()
        { 
            Settings.Default.Save();
        }
        public override void Enter()
        {
            LoadData();
        }
    } 
}
