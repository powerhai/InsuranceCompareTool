using System;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Input;
using InsuranceCompareTool.Services;
using Prism.Commands;
using Prism.Mvvm;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;
namespace InsuranceCompareTool.ViewModels
{
    public class MainWindowViewModel : BindableBase
    {
        #region Fields
        private BillLoadService mBillLoadService;
        private bool mIsEnabled = true;
        private MemberService mMemberService;
        private BillStatusService mBillStatusService;
        private BillFormatService mBillFormatService;
        private BillAreaService mBillAreaService;
        private BillMemberService mBillMemberService;
        private RelationLoadService mRelationLoadService;
        private DepartmentLoadService mDepartmentLoadService;
        private FileNameService mFileNameService;
        private BillExportTypeCService mBillExportTypeCService;
        private string mMembersFile;
        private string mSourceFile;
        private string mTargetFile;
        private string mTitle = "Document Comparing Tool";
        private string mDepartmentsFile;
        private string mRelationFile;
        private bool mIsMainTableEnabled;
        #endregion
        #region Properties
        public bool IsEnabled
        {
            get => mIsEnabled;
            set
            {
                mIsEnabled = value;
                RaisePropertyChanged();
            }
        }
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
            set =>  SetProperty(ref mRelationFile, value);
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
        public string DepartmentsFile
        {
            get => mDepartmentsFile;
            set { SetProperty(ref mDepartmentsFile, value); }
        }
        public string Title
        {
            get => mTitle;
            set => SetProperty(ref mTitle, value);
        }
        public bool IsMainTableEnabled
        {
            get => mIsMainTableEnabled;
            set => SetProperty(ref mIsMainTableEnabled, value);
        }
        public ICommand SelectMembersFileCommand
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
                    if(result == true)
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
                    var result = openFileDialog.ShowDialog();
                    if (result == true)
                        DepartmentsFile = openFileDialog.FileName;
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
                    if(result == DialogResult.OK)
                        TargetFile = openFileDialog.SelectedPath;
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
                        var exporter = new Exporter();
                        var file = exporter.Convert(SourceFile, TargetFile);
                        Process.Start(file);
                    }
                    catch(InvalidOperationException e1)
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
                    cfa.AppSettings.Settings[nameof(MembersFile)].Value = MembersFile;
                    cfa.AppSettings.Settings[nameof(DepartmentsFile)].Value = DepartmentsFile;
                    cfa.AppSettings.Settings[nameof(RelationFile)].Value = RelationFile;
                    cfa.Save();
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
                        if(!string.IsNullOrEmpty(TargetFile))
                            Process.Start(TargetFile);
                    }
                    catch(Exception ex)
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
                        if(!string.IsNullOrEmpty(MembersFile))
                            Process.Start(MembersFile);
                    }
                    catch(InvalidOperationException e1)
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

        public ICommand CreateDataCommand => new DelegateCommand(() => { CreateData(); });
        public ICommand ImportDataCommand => new DelegateCommand(ImportData);
        public ICommand BuildReportCommand => new DelegateCommand(BuildReportData);
        public ICommand ExportTypeCBillCommand => new DelegateCommand(ExportTypeCBills);
        #endregion
        #region Constructors
        public MainWindowViewModel(DepartmentLoadService departmentLoadService)
        {
            mDepartmentLoadService = DepartmentLoadService.CreateInstance();
            mMemberService = MemberService.CreateInstance();
            mBillLoadService = BillLoadService.CreateInstance();
            mBillStatusService = BillStatusService.CreateInstance();
            mBillFormatService = BillFormatService.CreateInstance();
            mBillAreaService = BillAreaService.CreateInstance();
            mBillMemberService = BillMemberService.CreateInstance();
            mRelationLoadService = RelationLoadService.CreateInstance();
            mBillExportTypeCService = BillExportTypeCService.CreateInstance();

            LoadData();
        }
        #endregion
        #region Private or Protect Methods
        private void BuildReportData() { }
        private void CreateData()
        {
            try
            {
                

                if (string.IsNullOrEmpty(MembersFile))
                    throw new Exception("亲，请先填写职员数据表");
                if(!File.Exists(MembersFile))
                    throw new Exception("亲，职员数据文件不存在");
                var openFileDialog = new OpenFileDialog
                {
                    Filter = "Excel Files (*.xlsx)|*.xlsx"
                };
                var result = openFileDialog.ShowDialog();
                if(result == true)
                    SourceFile = openFileDialog.FileName;
                else
                    return;

                if (File.Exists(mFileNameService.MainFile))
                {
                   var rv =  MessageBox.Show( $"继续操作将覆盖原有数据，是否继续？","文件已存在", MessageBoxButtons.YesNo);
                   if(rv != DialogResult.Yes)
                       return;
                }
                IsEnabled = false;
                var t = new Task(() =>
                {
                    try
                    {
                        mBillLoadService.Load(this.SourceFile);

                        mMemberService.Load( this.MembersFile);
                       
                        mRelationLoadService.Load(this.RelationFile);
                       // mDepartmentLoadService.Load(this.DepartmentsFile);

                        var bills = mBillLoadService.GetBills();
                        var members = mMemberService.GetMembers();
                        var relations = mRelationLoadService.GetRelations();
                        //var departments = mDepartmentLoadService.GetDepartments();

                        DataValidateService.ValidateDate(relations, bills,  members);
                        //mBillFormatService.Format(bills,members);
                        mBillAreaService.CalculateArea(bills);
                        mBillMemberService.CalculateMembers(bills,members);
                        mBillStatusService.CalculateStatus(bills, members);
                        BillDispatchService.CreateInstance().DispatchBills(bills,members);
                        
                        BillExportTypeAService.CreateInstance().Export(this.SourceFile, mFileNameService.MainFile , bills);
                        Process.Start(mFileNameService.MainFile);
                    }
                    catch(Exception ex)
                    {
                        MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                });
                t.Start();
                t.Wait();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                IsEnabled = true;
            } 
        }
        private void CheckFiles()
        {
            if (string.IsNullOrEmpty(MembersFile))
                throw new Exception("亲，请先填写职员数据表");
            if (!File.Exists(MembersFile))
                throw new Exception("亲，职员数据文件不存在");
        }
        private void ExportTypeCBills()
        {
            try
            {
                CheckFiles();
                var openFileDialog = new OpenFileDialog
                {
                    Filter = "Excel Files (*.xlsx)|*.xlsx"
                };
                var result = openFileDialog.ShowDialog();
                if (result == true)
                    SourceFile = openFileDialog.FileName;
                else
                    return;

                
                IsEnabled = false;
                var t = new Task(() =>
                {
                    try
                    {
                        mBillLoadService.Load(this.SourceFile);
                        mMemberService.Load(this.MembersFile); 
                        var bills = mBillLoadService.GetBills();
                        var members = mMemberService.GetMembers(); 
                        mBillMemberService.CalculateMembers(bills, members);  
                        mBillAreaService.CalculateArea(bills);
                        mBillExportTypeCService.Export(mFileNameService.ServicePath, bills,members);
                        Process.Start(mFileNameService.ServicePath);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                });
                t.Start();
                t.Wait();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                IsEnabled = true;
            }
        }
        private void ImportData() { }
        private void LoadData()
        {
            var cfa = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            SourceFile = cfa.AppSettings.Settings[nameof(SourceFile)].Value;
            TargetFile = cfa.AppSettings.Settings[nameof(TargetFile)].Value;
            MembersFile = cfa.AppSettings.Settings[nameof(MembersFile)].Value;
            DepartmentsFile = cfa.AppSettings.Settings[nameof(DepartmentsFile)].Value;
            RelationFile = cfa.AppSettings.Settings[nameof(RelationFile)].Value;
            mFileNameService = new FileNameService(TargetFile);
            IsMainTableEnabled = File.Exists(mFileNameService.MainFile);
        }
        #endregion
    }
}