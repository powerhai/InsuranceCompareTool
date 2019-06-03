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
using InsuranceCompareTool.Services;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;
namespace InsuranceCompareTool.ViewModels
{
	public class SettingsViewViewModel : ViewModelBase
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
                UpdateData();
            }
        }
        public string MembersFile
        {
            get => mMembersFile;
            set
            {
                mMembersFile = value;
                RaisePropertyChanged();
                UpdateData();
            }
        }
        public string RelationFile
        {
            get => mRelationFile;
            set
            {
                SetProperty(ref mRelationFile, value);
                UpdateData();
            }
        }
        public string TargetFile
        {
            get => mTargetFile;
            set
            {
                mTargetFile = value;
                RaisePropertyChanged();
                UpdateData();
            }
        }


        public string TemplateFile
        {
            get => mTemplateFile;
            set
            {
                SetProperty(ref mTemplateFile, value);
                UpdateData();
            }
        }
        public string DepartmentsFile
        {
            get => mDepartmentsFile;
            set
            {
                SetProperty(ref mDepartmentsFile, value);
                UpdateData();
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
            this.DepartmentsFile = mConfigService.DepartmentsFile;
            this.MembersFile = mConfigService.MembersFile;
            this.RelationFile = mConfigService.RelationFile;
            this.TemplateFile = mConfigService.TemplateFile;
            this.TargetFile = mConfigService.TargetFile;
            mIsLoaded = true;
        }
        private void UpdateData()
        {
            if(mIsLoaded)
            {
                mConfigService.DepartmentsFile = DepartmentsFile;
                mConfigService.MembersFile = MembersFile;
                mConfigService.RelationFile = RelationFile;
                mConfigService.TemplateFile = TemplateFile;
                mConfigService.TargetFile = TargetFile;
                SaveData();
            } 
        } 
        private void SaveData()
        {
            mConfigService.Save(); 
        } 
    } 
}
