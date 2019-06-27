using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Forms;
using System.Windows.Input;
using InsuranceCompareTool.Models;
using InsuranceCompareTool.Properties;
using InsuranceCompareTool.Services;
using Prism.Commands;
using Prism.Mvvm;
using Unity;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;
namespace InsuranceCompareTool.ViewModels
{
    public class MainWindowViewModel : ViewModelBase
    {
        private readonly IUnityContainer mContainer;
        private readonly ConfigService mConfigService;
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
        private ExportTemplateService mExportTemplateService;

        private const string mTitle = "Document Comparing Tool";

        private bool mIsMainTableEnabled;
        private string mTemplateFile;
        private CollectionViewSource mScreenView;
        private ICommand mCloseTabCommand;
        #endregion
        #region Properties
        public ObservableCollection<TabViewModelBase> Views { get; } = new ObservableCollection<TabViewModelBase>();
        public CollectionViewSource ScreenView
        {
            get => mScreenView ?? (mScreenView = new CollectionViewSource() { Source = Views });
        }
        public bool IsEnabled
        {
            get => mIsEnabled;
            set
            {
                mIsEnabled = value;
                RaisePropertyChanged();
            }
        }

        public override string Title
        {
            get => mTitle;
            set { }
        }
        public bool IsMainTableEnabled
        {
            get => mIsMainTableEnabled;
            set => SetProperty(ref mIsMainTableEnabled, value);
        }

        public ICommand GotoSettingsView
        {
            get
            {
                return new DelegateCommand(() =>
                {
                    var view = Views.FirstOrDefault(a => a is SettingsViewViewModel);
                    if (view == null)
                    {
                        view = mContainer.Resolve<SettingsViewViewModel>();
                        this.Views.Add(view);
                    }

                    this.ScreenView.View.MoveCurrentTo(view);
                });
            }
        }
        public ICommand GotoPolicyView
        {
            get
            {
                return new DelegateCommand(() =>
                {
                    var view = Views.FirstOrDefault(a => a is PolicyViewViewModel);
                    if (view == null)
                    {
                        view = mContainer.Resolve<PolicyViewViewModel>();
                        this.Views.Add(view);
                    }

                    this.ScreenView.View.MoveCurrentTo(view);
                });
            }
        }

        public ICommand AddTabCommand => new DelegateCommand(() =>
        {
            this.Views.Add(mContainer.Resolve<AssignViewViewModel>());
        });

        public ICommand CloseTabCommand
        {
            get
            {
                if (mCloseTabCommand == null)
                {
                    mCloseTabCommand = new DelegateCommand<TabViewModelBase>((p) =>
                    {
                        if (p is AssignViewViewModel)
                        {
                            this.Views.Remove(p);
                        }
                    });
                }
                return mCloseTabCommand;
            }
        }
        public ICommand SelectTabCommand
        {
            get
            {
                return new DelegateCommand<SelectionChangedEventArgs>((o) =>
               {
                   if (o.RemovedItems != null && o.RemovedItems.Count > 0)
                   {
                       var old = o.RemovedItems[0] as TabViewModelBase;
                       old?.Leave();
                   }

                   if (o.AddedItems != null && o.AddedItems.Count > 0)
                   {
                       var newd = o.AddedItems[0] as TabViewModelBase;
                       newd?.Enter();
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
                   foreach (var view in Views)
                   {
                       view.Leave();
                   }
               });
            }
        }
    
        public ICommand ImportDataCommand => new DelegateCommand(ImportData);
        public ICommand BuildReportCommand => new DelegateCommand(BuildReportData);
        public ICommand ExportTypeCBillCommand => new DelegateCommand(ExportTypeCBills);
        #endregion
        #region Constructors
        public MainWindowViewModel(IUnityContainer container, ConfigService configService, FileNameService fileNameService)
        {
            mContainer = container;
            mConfigService = configService;
            mFileNameService = fileNameService;
            mDepartmentLoadService = DepartmentLoadService.CreateInstance();
            mMemberService = MemberService.CreateInstance();
            mBillLoadService = BillLoadService.CreateInstance();
            mBillStatusService = BillStatusService.CreateInstance();
            mBillFormatService = BillFormatService.CreateInstance();
            mBillAreaService = BillAreaService.CreateInstance();
            mBillMemberService = BillMemberService.CreateInstance();
            mRelationLoadService = RelationLoadService.CreateInstance();
            mBillExportTypeCService = BillExportTypeCService.CreateInstance();
            this.Title = InsuranceCompareTool.Properties.Settings.Default.Projects;
            LoadData();
            LoadScreen();
        }
        private void LoadScreen()
        {
            this.Views.Add(mContainer.Resolve<AssignViewViewModel>()); 
            this.Views.Add(mContainer.Resolve<PolicyViewViewModel>());
            this.Views.Add(mContainer.Resolve<SettingsViewViewModel>());
        }
        #endregion
        #region Private or Protect Methods
        private void BuildReportData() { }

        private void CheckFiles()
        {
            if (string.IsNullOrEmpty(Settings.Default.MembersFile))
                throw new Exception("亲，请先填写职员数据表");
            if (!File.Exists(Settings.Default.MembersFile))
                throw new Exception("亲，职员数据文件不存在");
            if (string.IsNullOrEmpty(Settings.Default.TemplateFile))
                throw new Exception("宝贝！请先填写导出模板数据表");
            if (!File.Exists(Settings.Default.TemplateFile))
                throw new Exception("宝贝！导出模板数据表文件不存在");
        }
        private void ExportTypeCBills()
        {
            try
            {
                String sourceFile = "";
                CheckFiles();
                var openFileDialog = new OpenFileDialog
                {
                    Filter = "Excel Files (*.xlsx)|*.xlsx"
                };
                var result = openFileDialog.ShowDialog();
                if (result == true)
                    sourceFile = openFileDialog.FileName;
                else
                    return;


                IsEnabled = false;
                var t = new Task(() =>
                {
                    try
                    {
                        mExportTemplateService.Load(Settings.Default.TemplateFile);
                        mBillLoadService.Load(sourceFile);
                        mMemberService.Load(Settings.Default.MembersFile);
                        var bills = mBillLoadService.GetBills();
                        var members = mMemberService.GetMembers();
                        mBillMemberService.CalculateMembers(bills, members);
                        mBillAreaService.CalculateArea(bills);
                        mBillExportTypeCService.Export(mFileNameService.ServicePath, bills, members);
                        Process.Start(mFileNameService.ServicePath);
                    }
                    //catch (InvalidOperationException e1)
                    //{
                    //    MessageBox.Show("操作不能继续：文件正被使用!", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //}
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

            mExportTemplateService = ExportTemplateService.CreateInstance();
            IsMainTableEnabled = File.Exists(mFileNameService.MainFile);
        }
        #endregion
    }
}