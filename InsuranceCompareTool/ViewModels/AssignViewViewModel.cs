using Prism.Commands;
using Prism.Mvvm;
using System;
using System.CodeDom;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Xml.Serialization;
using InsuranceCompareTool.Core;
using InsuranceCompareTool.Domain;
using InsuranceCompareTool.Models;
using InsuranceCompareTool.Models.Dispatch;
using InsuranceCompareTool.Properties;
using InsuranceCompareTool.Services;
using InsuranceCompareTool.ShareCommon;
using log4net;
using Prism.Interactivity.InteractionRequest;
using DataGrid = System.Windows.Controls.DataGrid;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;
namespace InsuranceCompareTool.ViewModels
{


    public class AssignViewViewModel : TabViewModelBase
    {
        #region Fields
        private readonly ILog mLogger;
        private const string SHOW_ALL = "显示全部";
        private ColumnLayoutHelper mColumnLayoutHelper = new ColumnLayoutHelper();
        private ProjectCacheHelper mProjectCacheHelper = new ProjectCacheHelper();
        private ObservableCollection<ColumnVisible> mColumnLayouts = new ObservableCollection<ColumnVisible>();
        private readonly ObservableCollection<Project> mProjects = new ObservableCollection<Project>();
        private List<DataRow> mSelectedRows = new List<DataRow>();
        private List<DataRow> mSelectedCachedRows = new List<DataRow>();
        private List<DataRow> mCurrentSelectedRows = new List<DataRow>();
        private DataTable mDataTable;
        private List<UIColumnData> mColumns;
        private CollectionViewSource mLayoutView;
        private ColumnVisible mLastLayout;
        private string mSourceFile;
        private bool mIsBusy;
        private ObservableCollection<FilterData> mFilters = new ObservableCollection<FilterData>();
        private CollectionViewSource mFilterView;

        private CollectionViewSource mMemberView;
        private DateTime mLastMemberFileWriteTime = new DateTime(1900);
        private bool mIsLoaded = false;
        private ICommand mAssignToSomebodyCommand;
        private List<string> mAreas = new List<string>() { AreaNames.JH, AreaNames.YK, AreaNames.YW, AreaNames.WY, AreaNames.PJ, AreaNames.LX, AreaNames.QT, AreaNames.DY };
        private CollectionViewSource mAreaView;
        private ObservableCollection<MemberA> mMembers = new ObservableCollection<MemberA>();
        private CollectionViewSource mProjectView;
        private ICollectionView mDataTableView;
        private List<Member> mAllMembers;
        private ICommand mSelectBillCommand;

        private ObservableCollection<MemberA> mSelectedMembers = new ObservableCollection<MemberA>();
        private ICommand mAssignToServiceCommand;
        private CollectionViewSource mSelectMemberView;
        private CollectionViewSource mReportMemberView;
        private string mCurrentArea;
        private DataTable mCacheDataTable;
        private ICommand mSelectCachedBillsCommand;
        private ICommand mCopyBillContentCommand;
        private List<string> mCopyableColumns;
        private ICommand mAssignToServiceBCommand;
        private ICommand mAssignToSomebodyByMenuCommand;
        private ICommand mSelectTabItemCommand;
        #endregion
        #region Properties
        public ObservableCollection<MemberA> SelectedMembers => mSelectedMembers;
        public InteractionRequest<ShowTabItemNotification> SelectItemNotification { get; } = new InteractionRequest<ShowTabItemNotification>();
        public List<string> Areas
        {
            get => mAreas;
        }

        public List<DataRow> SelectedCachedRows
        {
            get => mSelectedCachedRows;
            set => SetProperty(ref mSelectedCachedRows, value);
        }
        public List<DataRow> SelectedRows
        {
            get => mSelectedRows;
            set => SetProperty(ref mSelectedRows, value);

        }
        public CollectionViewSource AreaView
        {
            get
            {
                if (mAreaView == null)
                {
                    mAreaView = new CollectionViewSource()
                    {
                        Source = Areas
                    };
                    mAreaView.View.CurrentChanged += OnCurrentAreaChanged;
                }

                return mAreaView;
            }
        }
        public CollectionViewSource SelectMemberView
        {
            get
            {
                if (mSelectMemberView == null)
                {
                    mSelectMemberView = new CollectionViewSource()
                    {
                        Source = mSelectedMembers
                    }; 
                } 
                return mSelectMemberView;
            }
        }
 
        public CollectionViewSource ReportMembersView
        {
            get
            {
                if(mReportMemberView == null)
                {
                    mReportMemberView = new CollectionViewSource(){Source = Members };
                    mReportMemberView.SortDescriptions.Add(new SortDescription("Count", ListSortDirection.Descending));
                    mReportMemberView.View.Filter += FilterReportMember;
                } 
                return mReportMemberView;
            }
        }
        public ObservableCollection<MemberA> Members
        {
            get => mMembers;
        }
        public CollectionViewSource MemberView
        {
            get
            {
                if (mMemberView == null)
                {
                    mMemberView = new CollectionViewSource() { Source = Members };
                    mMemberView.SortDescriptions.Add(new SortDescription("Count", ListSortDirection.Descending));
                    mMemberView.View.Filter += FilterMember;
                }
                return mMemberView;
            }
        }
        public ObservableCollection<FilterData> Filters
        {
            get => mFilters;
        }
        public CollectionViewSource FilterView
        {
            get
            {
                if (mFilterView == null)
                {
                    mFilterView = new CollectionViewSource()
                    {
                        Source = mFilters
                    };
                    mFilterView.View.CurrentChanged += ViewCurrentChanged;
                }
                return mFilterView;
            }
        }
        public List<string> CopyableColumns
        {
            get => mCopyableColumns;
            set => SetProperty(ref mCopyableColumns, value);
        }
        public bool IsBusy
        {
            get => mIsBusy;
            set => SetProperty(ref mIsBusy, value);
        }
        public override string Title { get; set; } = "分单 ";
        public bool IsExportEnable
        {
            get { return DataTable != null; }
        }
        public string SourceFile
        {
            get => mSourceFile;
            set => SetProperty(ref mSourceFile, value);
        }
        public ObservableCollection<Project> Projects
        {
            get => mProjects;
        }
        public string TotalMoney
        {
            get
            {
                if (DataTable == null)
                    return "0";
                double val = 0;

                foreach (DataRow row in DataTable.Rows)
                {

                    if (row.Table.Columns.Contains(BillSheetColumns.BILL_PRICE))
                    {
                        double bf = 0;
                        double.TryParse(row[BillSheetColumns.BILL_PRICE].ToString(), out bf);
                        val += bf;
                    }
                }
                return val.ToString("C");
            }
        }
        public ICollectionView DataTableView
        {
            get
            {
                if (mDataTableView == null)
                {
                    mDataTableView = CollectionViewSource.GetDefaultView(DataTable);
                }

                return mDataTableView;
            }
        }
        public DataTable CacheDataTable
        {
            get => mCacheDataTable;
            set
            {
                SetProperty(ref mCacheDataTable, value);
                OnPropertyChanged();
            }
        }
        public DataTable DataTable
        {
            get => mDataTable;
            set
            {
                SetProperty(ref mDataTable, value);
                OnPropertyChanged(nameof(IsExportEnable));
                OnPropertyChanged(nameof(TotalMoney));

            }
        }
        public List<UIColumnData> Columns
        {
            get => mColumns;
            set => SetProperty(ref mColumns, value);
        }
        public CollectionViewSource LayoutView
        {
            get
            {
                if (mLayoutView == null)
                {
                    mLayoutView = new CollectionViewSource()
                    {
                        Source = mColumnLayouts
                    };

                    mLayoutView.View.CurrentChanged += LayoutChanged;
                }
                return mLayoutView;
            }
        }

        public CollectionViewSource ProjectView
        {
            get
            {
                if (mProjectView == null)
                {
                    mProjectView = new CollectionViewSource() { Source = mProjects };
                }

                return mProjectView;
            }
        }
        #endregion
        #region Methods
        private void LogAssgign(DataRow bill, MemberA mem)
        {
            var log = Convert.ToString(bill[BillSheetColumns.SYS_HISTORY]);
            var service = "";
            if (!bill.IsNull(BillSheetColumns.SYS_SERVICE))
            {
                service = Convert.ToString(bill[BillSheetColumns.SYS_SERVICE]);
            }
            var newlog = $"手工派单 -> {mem.Name} ({mem.ID} - {mem.Area})";
            var history = "";
            if (string.IsNullOrEmpty(log))
            {
                if (string.IsNullOrEmpty(service))
                {
                    history = newlog;
                }
                else
                {
                    history = $"原始客服专员： {service}\r\n{newlog}";
                }

            }
            else
            {
                history = $"{log}\r\n{newlog}";
            }

            bill[BillSheetColumns.SYS_HISTORY] = history;
        }
 
        private bool FilterMember(object obj)
        {
            var area = AreaView.View.CurrentItem as string;
            if (string.IsNullOrEmpty(area))
            {
                return true;
            }

            var ob = obj as MemberA;
            return ob.Area == area;
        }

        private bool FilterReportMember(object obj)
        {
             
            if (string.IsNullOrEmpty(mCurrentArea))
            {
                return true;
            }

            var ob = obj as MemberA;
            return ob.Area == mCurrentArea || ob.Count > 0;
        }
        private void ViewCurrentChanged(object sender, EventArgs e)
        {
            var filter = mFilterView.View.CurrentItem as FilterData;
            if (filter == null)
                return;
            if (filter.ShowType == FilterType.ShowAll)
                DataTable.DefaultView.RowFilter = "";
            else if (filter.ShowType == FilterType.ShowFilter)
            {
                DataTable.DefaultView.RowFilter = $"{BillSheetColumns.SYS_FILTER} = '{filter.Name}'";
            }
            else
            {
                DataTable.DefaultView.RowFilter = $"{BillSheetColumns.SYS_FILTER} = '' ";
            }

        }
        private void OnCurrentAreaChanged(object sender, EventArgs e)
        {
            mCurrentArea = AreaView.View.CurrentItem as string;
            ReportMembersView.View.Refresh();
            MemberView.View.Refresh();
            MemberView.View.MoveCurrentToFirst();
        }
        public AssignViewViewModel(ILog logger)
        {
            mLogger = logger;
        }
        private void CacheLayout()
        {
            if (mLastLayout == null || Columns == null)
                return;
            mLastLayout.Columns = this.Columns.Where(a => a.IsVisible == true).Select(a => a.Name).ToArray();
        }
        private void GenerateFilters(Project project)
        {
            mFilters.Clear();

            if (DataTable == null)
                return;
            mFilters.Add(new FilterData()
            {
                Count = DataTable.Rows.Count,
                Name = SHOW_ALL,
                ShowType = FilterType.ShowAll
            });
            var noHandledCount = 0;
            foreach (DataRow dr in DataTable.Rows)
            {
                if ((string)dr[BillSheetColumns.SYS_FILTER] == "")
                {
                    noHandledCount += 1;
                }
            }
            mFilters.Add(new FilterData()
            {
                Count = noHandledCount,
                Name = "无处理",
                ShowType = FilterType.ShowNoFilter
            });

            foreach (var step in project.Steps)
            {
                var filter = new FilterData() { Name = step.Title, ShowType = FilterType.ShowFilter };
                var count = 0;
                foreach (DataRow row in DataTable.Rows)
                {
                    if (row[BillSheetColumns.SYS_FILTER] != DBNull.Value && (string)row[BillSheetColumns.SYS_FILTER] == step.Title)
                    {
                        count += 1;
                    }
                }

                filter.Count = count;
                if (filter.Count > 0)
                    mFilters.Add(filter);
            }

        }

        private async void ExportDBills(List<DataRow> rows)
        {
            var exporter = new BillExportTypeK();
            var targetPath = $"{Settings.Default.WorkPath}\\上传保单\\{DateTime.Now:yyyy-MM-dd hh-mm-ss}";
            IsBusy = true;
            await Task.Run(() =>
            {
                try
                {

                    exporter.Export(targetPath, rows, DataTable);

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }

            });
            IsBusy = false;
            try
            {
                if (!string.IsNullOrEmpty(targetPath))
                    Process.Start(targetPath);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message, "错误", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }




        private async void ExportDBillsAll()
        { 
            var rows = new List<DataRow>();
            foreach (DataRow row in mCurrentDataTable.Rows)
            {
                rows.Add(row);
            }
            ExportDBills(rows);
        }
        private void SaveLayout()
        {

            var cur = mLayoutView.View.CurrentItem as ColumnVisible;
            if (cur != null && Columns != null)
            {
                cur.Columns = this.Columns.Where(a => a.IsVisible == true).Select(a => a.Name).ToArray();
            }

            mColumnLayoutHelper.SaveColumnLayout(mColumnLayouts.ToList());
        }
        public override void Leave()
        {
            this.SaveLayout();
        }
        public override void Close()
        {
            this.SaveLayout();
        }
        public override async void Enter()
        {
            IsBusy = true;
            await LoadMembers();
            this.LoadLayout();
            this.LoadProjects();
            IsBusy = false;
        }
        private void LoadLayout()
        {
            var layouts = mColumnLayoutHelper.GetColumnLayouts();

            foreach (var layout in layouts)
            {
                if (!mColumnLayouts.Select(a => a.Guid).Contains(layout.Guid))
                {
                    this.mColumnLayouts.Add(layout);
                }
                else
                {
                    var lay = mColumnLayouts.FirstOrDefault(a => a.Guid.Equals(layout.Guid));
                    lay.Columns = layout.Columns;

                }

                var cur = LayoutView.View.CurrentItem as ColumnVisible;
                if (cur == null)
                {
                    continue;
                }

                if (Columns == null)
                    continue;

                foreach (var col in Columns)
                {
                    col.IsVisible = cur.Columns.Contains(col.Name);
                }
            }

            foreach (var col in mColumnLayouts.ToArray())
            {
                if (!layouts.Select(a => a.Guid).Contains(col.Guid))
                {
                    mColumnLayouts.Remove(col);
                }
            }

            if (mColumnLayouts.Count <= 0)
            {
                var lay = new ColumnVisible()
                {
                    Title = "可视布局 1",
                    Guid = Guid.NewGuid()
                };
                mColumnLayouts.Add(lay);
            }

            this.LayoutView.View.MoveCurrentToFirst();
        }
        private void LoadProjects()
        {
            var curProject = this.ProjectView.View.CurrentItem as Project;

            this.Projects.Clear();
            var list = mProjectCacheHelper.GetProjects();

            foreach (var project in list)
            {
                this.Projects.Add(project);
            }

            this.ProjectView.View.MoveCurrentToFirst();
            if (curProject != null)
            {
                var pro = this.Projects.FirstOrDefault(a => a.Guid.Equals(curProject.Guid));
                if (pro != null)
                {
                    this.ProjectView.View.MoveCurrentTo(pro);
                }
            }
        }
        private async Task LoadMembers()
        {
            if (string.IsNullOrEmpty(Settings.Default.MembersFile))
                return;
            FileInfo fi = new FileInfo(Settings.Default.MembersFile);
            if (!fi.Exists)
                return;

            if (mIsLoaded == true)
            {
                if (fi.LastWriteTime == mLastMemberFileWriteTime)
                    return;
            }

            mIsLoaded = true;

            mLastMemberFileWriteTime = fi.LastWriteTime;

            var memService = new MemberService();
            await Task.Run(() => { memService.Load(Settings.Default.MembersFile); });
            Members.Clear();
            var mems = memService.GetZaiziMembers().Where(a => a.Position == "客服专员");
            mAllMembers = memService.GetMembers().ToList();
            foreach (var mem in mems)
            {
                Members.Add(new MemberA() { ID = mem.ID, Name = mem.Name, Count = 0, Area = mem.Area });
            }

            MemberView.View.MoveCurrentToFirst();
        }
        private void MakeColumns(DataTable dt)
        {
            var list = new List<UIColumnData>();
            var curLayout = LayoutView.View.CurrentItem as ColumnVisible;
            List<string> sysCols = new List<string>();
            List<string> copyableCols = new List<string>();

            foreach (ColumnDefine col in BillTableColumns.Columns.Where(a => a.IsSystemColumn))
            {
                sysCols.AddRange(col.Name);
            }

            var index = 0;
            foreach (DataColumn dataColumn in dt.Columns)
            {
                if (sysCols.Contains(dataColumn.ColumnName))
                    continue;
                ColumnDefine defineCol = null;
                foreach(var colData in BillTableColumns.Columns)
                {
                    if(colData.Name.Contains(dataColumn.ColumnName))
                    {
                        defineCol = colData;
                        if(colData.IsCopyAble)
                        {
                            copyableCols.Add(dataColumn.ColumnName);
                        }
                        break;
                    }
                }
                var col = new UIColumnData()
                {
                    Name = dataColumn.ColumnName,
                    IsVisible = true,
                    Sort =  defineCol?.Sort ?? 0,
                    Index = index ++
                };

                if (curLayout != null)
                {
                    col.IsVisible = curLayout.Columns.Contains(col.Name);
                }
                list.Add(col);
            }
            list.Sort((a, b) => { return b.Sort - a.Sort;});
            Columns = list;
            CopyableColumns = copyableCols;
        }
        private void LayoutChanged(object sender, EventArgs e)
        {
            if (Columns == null)
                return;

            CacheLayout();
            var cur = LayoutView.View.CurrentItem as ColumnVisible;
            if (cur == null)
            {
                return;
            }
            foreach (var col in Columns)
            {
                col.IsVisible = cur.Columns.Contains(col.Name);
            }

            mLastLayout = cur;
        }
        private async void ExportTypeCBills()
        {
            IsBusy = true;
            var exception = await Task.Run(() =>
          {
              try
              {
                  var templateService = new ExportTemplateService();
                  var billLoader = new BillLoadService();
                  var memberService = new MemberService();
                  var billMemberService = new BillMemberService();
                  var billAreaService = new BillAreaService();
                  var billExportTypeCService = new BillExportTypeCService();

                  templateService.Load(Settings.Default.TemplateFile);
                  billLoader.Load(DataTable);
                  memberService.Load(Settings.Default.MembersFile);
                  var bills = billLoader.GetBills();
                  var members = memberService.GetMembers();


                  billMemberService.CalculateMembers(bills, members);
                  billAreaService.CalculateArea(bills);
                  var path = Settings.Default.WorkPath + $"{DateTime.Now.AddMonths(1):yyyy-MM}\\地区收费清单";

                  billExportTypeCService.Export(path, bills, members);
                  Process.Start(path);
              }
              catch (InvalidOperationException)
              {
                  return new Exception("操作不能继续：文件正被使用!");
              }
              catch (Exception ex)
              {
                  return ex;
              }

              return null;
          });
            if (exception != null)
            {
                Xceed.Wpf.Toolkit.MessageBox.Show(exception.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            IsBusy = false;
        }
        private void CheckDatatable(DataTable dt)
        {
            foreach (DataRow dr in dt.Rows)
            {
                if (dt.Columns.Contains(BillSheetColumns.SELLER_STATE))
                {
                    if (dr.IsNull(BillSheetColumns.SELLER_STATE))
                    {

                        mLogger.Error($"{dr[BillSheetColumns.SELLER_NAME]} 的在职状态未知");
                        break;
                    }
                }


                if (dr.IsNull(BillSheetColumns.SYS_SERVICE_STATUS))
                {
                    mLogger.Error($"{dr[BillSheetColumns.CURRENT_SERVICE_NAME]} 的在职状态未知");
                    break;
                }
            }
        }
        private async void ExportTypeEBills()
        {
            var exporter = new BillExportTypeE();
            var targetPath = $"{Settings.Default.WorkPath}\\导出保单";
            IsBusy = true;
            await Task.Run(() =>
            {
                try
                {
                    exporter.Export(targetPath, DataTable, Columns.Select(a => a.Name).ToList());
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }

            });
            IsBusy = false;
            try
            {
                if (!string.IsNullOrEmpty(targetPath))
                    Process.Start(targetPath);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message, "错误", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        private DataRow GetRow(DataTable dt, DataRow dr)
        {
            var guid = (Guid)dr[BillSheetColumns.SYS_GUID];
            DataRow anotherRow = null;
            foreach (DataRow row in dt.Rows)
            {
                if (((Guid)row[BillSheetColumns.SYS_GUID]).Equals(guid))
                {
                    anotherRow = row;
                    break;
                }
            }

            return anotherRow;
        }

        private async void AssignBills()
        {
            mLogger.Info("开始分单");
            var cur = ProjectView.View.CurrentItem as Project;
            if (cur == null)
                return;

            var assignHelper = new BillAssignHelper(mLogger, mAllMembers);
            IsBusy = true;
            var ex = await Task.Run<Exception>(() =>
            {
                try
                {
                    assignHelper.AssignBills(DataTable, cur);
                    var rows = GetChangedRows();
                    foreach(var row in rows)
                    {
                        var dr = GetRow(CacheDataTable, row); 
                        if(dr != null)
                        {
                            CopyAssign(row,dr);
                        }
                    }
                }
                catch (Exception e)
                {
                    return e;
                }

                return null;
            });
            IsBusy = false;
            mLogger.Info("分单结束");
            if (ex != null)
            {
                Xceed.Wpf.Toolkit.MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            GenerateFilters(cur);

            Xceed.Wpf.Toolkit.MessageBox.Show("分单完成", "Finished", MessageBoxButton.OK, MessageBoxImage.Information);
            SelectItemNotification?.Raise(new ShowTabItemNotification() { Index = 2 });

        }
        private void CopyAssign(DataRow from, DataRow to)
        {
            to[BillSheetColumns.CURRENT_SERVICE_NAME] = from[BillSheetColumns.CURRENT_SERVICE_NAME];
            to[BillSheetColumns.CURRENT_SERVICE_ID] = from[BillSheetColumns.CURRENT_SERVICE_ID];
            to[BillSheetColumns.SYS_SERVICE_AREA] = from[BillSheetColumns.SYS_SERVICE_AREA];
            to[BillSheetColumns.SYS_HISTORY] = from[BillSheetColumns.SYS_HISTORY];
        }
        private void AssignTo(DataRow bill, MemberA member)
        {

            if(!bill.IsNull(BillSheetColumns.CURRENT_SERVICE_ID))
            {
                var serID = (string)bill[BillSheetColumns.CURRENT_SERVICE_ID];
                if(serID.Equals(member.ID, StringComparison.CurrentCultureIgnoreCase))
                {
                    return;
                }
            }

            bill[BillSheetColumns.CURRENT_SERVICE_NAME] = member.Name;
            bill[BillSheetColumns.CURRENT_SERVICE_ID] = member.ID;
            bill[BillSheetColumns.SYS_SERVICE_AREA] = member.Area;
            
            DataTable dt = null;
            if(bill.Table == DataTable)
            {
                dt = CacheDataTable;
            }
            else
            {
                dt = DataTable;
            }

            var guid = (Guid)bill[BillSheetColumns.SYS_GUID];
            DataRow anotherRow = null;
            foreach(DataRow row in dt.Rows)
            { 
                if(((Guid)row[BillSheetColumns.SYS_GUID]).Equals(guid))
                {
                    anotherRow = row;
                    break;
                }
            }

            if(anotherRow != null)
            {
                anotherRow[BillSheetColumns.CURRENT_SERVICE_NAME] = member.Name;
                anotherRow[BillSheetColumns.CURRENT_SERVICE_ID] = member.ID;
                anotherRow[BillSheetColumns.SYS_SERVICE_AREA] = member.Area;
                LogAssgign(anotherRow, member);
            }
            LogAssgign(bill, member);
        }

        private void CalculateCountForMembers(DataTable dt)
        {
            if(dt == null)
                return;

            foreach (MemberA member in Members)
            {
                member.Count = 0;
                member.Price = 0;
            }

            bool hasPriceCol = dt.Columns.Contains(BillSheetColumns.BILL_PRICE);

            if (dt.Columns.Contains(BillSheetColumns.CURRENT_SERVICE_ID))
            {
                foreach (DataRow row in dt.Rows)
                {
                    if (row.IsNull(BillSheetColumns.CURRENT_SERVICE_ID))
                        continue;
                    var serID = (string)row[BillSheetColumns.CURRENT_SERVICE_ID];
                    var mem = Members.FirstOrDefault(a =>
                        a.ID.Equals(serID, StringComparison.CurrentCultureIgnoreCase));
                    if (mem == null)
                        continue;
                    if (hasPriceCol)
                    {
                        if (!row.IsNull(BillSheetColumns.BILL_PRICE))
                        {
                            var price = Convert.ToDouble(row[BillSheetColumns.BILL_PRICE]);
                            mem.Price += price;
                        }
                    }
                    mem.Count++;
                }
            }

            var firMem = Members.FirstOrDefault(a => a.Count > 0);
            if (firMem != null)
            {
                mCurrentArea = this.Areas.FirstOrDefault(a => a == firMem.Area);
                this.AreaView.View.MoveCurrentTo(mCurrentArea); 
            }
            
            MemberView.View.Refresh();
            ReportMembersView.View.Refresh();
        }
        #endregion
        #region Commands
        public ICommand AssignToServiceBCommand
        {
            get
            {
                if(mAssignToServiceBCommand == null)
                {
                    mAssignToServiceBCommand = new DelegateCommand<MemberA>((m) =>
                    { 
                        if (m == null)
                            return;

                        foreach (var row in mCurrentSelectedRows)
                        {
                            AssignTo(row,m); 
                        }

                        this.CalculateCountForMembers(DataTable);
                    });
                }
                return mAssignToServiceBCommand;
            }
        }
        public ICommand AssignToServiceCommand
        {
            get
            {
                if (mAssignToServiceCommand == null)
                {
                    mAssignToServiceCommand = new DelegateCommand(() =>
                    {

                        var mem = SelectMemberView.View.CurrentItem as MemberA;
                        if (mem == null)
                            return;

                        foreach (var row in SelectedRows)
                        {  
                            AssignTo(row, mem); 
                        }
                        this.CalculateCountForMembers(DataTable);

                    });
                }

                return mAssignToServiceCommand;
            }
        }
        public ICommand AssignToSomebodyCommand
        {
            get
            {
                return mAssignToSomebodyCommand ?? (mAssignToSomebodyCommand = new DelegateCommand(() =>
                {
                    var mem = MemberView.View.CurrentItem as MemberA;
                    if (mem == null)
                        return;

                    foreach (var row in mCurrentSelectedRows)
                    {
                        AssignTo(row, mem);
                    }
                    this.CalculateCountForMembers(DataTable);
                }));
            }
        }
        public ICommand AssignToSomebodyByMenuCommand
        {
            get
            {
                return mAssignToSomebodyByMenuCommand ?? (mAssignToSomebodyByMenuCommand = new DelegateCommand <MemberA>(( mem) =>
                { 
                    if (mem == null)
                        return; 
                    foreach (var row in mCurrentSelectedRows)
                    {
                        AssignTo(row, mem);
                    }
                    this.CalculateCountForMembers(DataTable);
                }));
            } 
        }

        public ICommand SearchCommand
        {
            get
            {
                return new DelegateCommand<string>((s) =>
                {
                    this.FilterView.View.MoveCurrentTo(null);
                    List<string> unstrCols = new List<string>();
                    foreach (var col in BillTableColumns.Columns.Where(a => a.Type != typeof(String)))
                    {
                        unstrCols.AddRange(col.Name);
                    }
                    var str = "";
                    var cols = Columns.Where(a => a.IsVisible == true);
                    foreach (var col in cols)
                    {
                        if (unstrCols.Contains(col.Name))
                        {
                            continue;
                        }

                        if (!string.IsNullOrEmpty(str))
                        {
                            str += " or ";
                        }

                        str += $"{col.Name} like '%{s}%'";
                    }

                    if (string.IsNullOrEmpty(s))
                    {
                        CurrentDataTable.DefaultView.RowFilter = "";
                    }
                    else
                    {
                        CurrentDataTable.DefaultView.RowFilter = str;
                    }

                });
            }
        }
        public ICommand LoadBillCommand
        {
            get
            {
                return new DelegateCommand(async () =>
                {


                    var openFileDialog = new OpenFileDialog
                    {
                        Filter = "Excel Files (*.xlsx)|*.xlsx",

                    };

                    var result = openFileDialog.ShowDialog();
                    if (result == true)
                    {
                        var loader = new BillLoadHelper(mLogger);
                        SourceFile = openFileDialog.FileName;
                        IsBusy = true;
                        try
                        {

                            DataTable dt = await Task.Run(() =>
                            {
                                var dct = loader.Load(openFileDialog.FileName);
                                var memhelp = new MemberService();
                                memhelp.Load(Settings.Default.MembersFile);
                                var mems = memhelp.GetMembers();

                                var formater = new BillFormatHelper(mLogger, mems);
                                formater.Format(dct);

                                CheckDatatable(dct);
                                return dct;

                            });

                            CalculateCountForMembers(dt);
                            MakeColumns(dt);
                            DataTable = dt;
                             
                            this.CacheDataTable = dt.Clone();
                        }
                        catch (Exception ex)
                        {
                            Xceed.Wpf.Toolkit.MessageBox.Show(ex.Message, "error", MessageBoxButton.OK,
                                MessageBoxImage.Error);
                        }
                        finally
                        {
                            mFilters.Clear();
                            IsBusy = false;
                        }
                        ReportMembersView.View.Refresh();
                        SelectItemNotification?.Raise(new ShowTabItemNotification(){ Index = 2 });
                    }
                });
            }
        }
        public ICommand SelectCachedBillsCommand
        {
            get
            {
                if(mSelectCachedBillsCommand == null)
                {
                    mSelectCachedBillsCommand = new DelegateCommand<ObservableCollection<object>>((b) =>
                    {
                        var list = new List<DataRow>();
                        foreach(DataRowView row in b)
                        {
                            list.Add(row.Row);
                        } 
                        SelectedCachedRows = list;
                    }); 
                }

                return mSelectCachedBillsCommand;
            }

        }
        public ICommand SelectBillCommand
        {
            get
            {
                if (mSelectBillCommand == null)
                {
                    mSelectBillCommand = new DelegateCommand<ObservableCollection<object>>((b) =>
                    {
                        this.SelectedMembers.Clear();
                        var list = new System.Collections.Generic.Dictionary<string, MemberA>();
                        var selectedRows = new List<DataRow>();

                        foreach (DataRowView row in b)
                        {
                            var bill = row.Row;
                            selectedRows.Add(bill);
                            if (!bill.IsNull(BillSheetColumns.CURRENT_SERVICE_ID))
                            {
                                var sid = Convert.ToString(bill[BillSheetColumns.CURRENT_SERVICE_ID]);
                                if (list.ContainsKey(sid))
                                    continue;
                                var mem = mMembers.FirstOrDefault(a => a.ID.Equals(sid));
                                if (mem == null)
                                    continue;
                                //if(mem.Status != StatusNames.ZAI_ZI)
                                //    continue;
                                //list[sid]  = new MemberA(){ ID = sid, Area = mem.Area, Name = mem.Name };
                                list[sid] = mem;
                              
                            }
                        }
                        this.SelectedRows = selectedRows;
                        if (list.Count <= 1)
                            return;
                        foreach (KeyValuePair<string, MemberA> mem in list)
                        {
                            this.SelectedMembers.Add(mem.Value);
                        }

                        this.SelectMemberView.View.MoveCurrentToFirst();
                        
                    });
                }

                return mSelectBillCommand;
            }
        }

        public ICommand SaveMemberCommand
        {
            get
            {
                return new DelegateCommand(() =>
                {
                    var targetPath = $"{Settings.Default.WorkPath}\\调单保单统计";

                    var list = new List<MemberA>();
                    this.ReportMembersView.View.MoveCurrentToFirst();
                    do
                    {
                        list.Add(this.ReportMembersView.View.CurrentItem as MemberA);
                    } while(this.ReportMembersView.View.MoveCurrentToNext());
                    

                    try
                    {
                        var helper = new MemberExportHelper(); 
                        helper.Export(targetPath, list);
                        if (!string.IsNullOrEmpty(targetPath))
                            Process.Start(targetPath);
                    }
                    catch (Exception ex)
                    {
                        Xceed.Wpf.Toolkit.MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    }


                });
            }
        }
        public ICommand AddLayoutCommand
        {
            get
            {
                return new DelegateCommand(() =>
               {
                   var layout = new ColumnVisible()
                   {
                       Title = $"可视布局 {mColumnLayouts.Count + 1}",
                       Guid = Guid.NewGuid()
                   };
                   layout.Columns = this.Columns.Where(a => a.IsVisible == true).Select(a => a.Name).ToArray();
                   this.mColumnLayouts.Add(layout);
                   this.mLayoutView.View.MoveCurrentTo(layout);
               });
            }
        }
        public ICommand DelLayoutCommand
        {
            get
            {
                return new DelegateCommand<ColumnVisible>((p) =>
                {
                    if (this.mColumnLayouts.Count <= 1)
                    {
                        return;
                    }
                    this.mColumnLayouts.Remove(p);
                    this.LayoutView.View.MoveCurrentToFirst();
                });
            }
        }

        public ICommand ExportCCommand
        {
            get
            {
                return new DelegateCommand(ExportTypeCBills);
            }
        }
        public ICommand ExportDCommand
        {
            get
            {
                return new DelegateCommand(ExportDBillsAll);
            }
        }
        public ICommand ExportFCommand
        {
            get
            {
                return new DelegateCommand(() =>
                {  
                    ExportDBills(mCurrentSelectedRows);
                });
            }
        }
        public ICommand CopyMemberCommand
        {
            get
            {
                return new DelegateCommand(() =>
                {
                    StringBuilder sb = new StringBuilder();
                    ReportMembersView.View.MoveCurrentToFirst(); 
                     
                    do
                    {
                        var mem = ReportMembersView.View.CurrentItem as MemberA;
                        var c = ' ';
                        sb.AppendLine($"{mem.ID.PadRight(10,c)}\t{mem.Name.ToString().PadRight(5, c)}\t{mem.Count.ToString().PadLeft(5,c)}\t\t{mem.Price.ToString("N").PadLeft(13,c)}\t\t{mem.Area}");
                    } while(ReportMembersView.View.MoveCurrentToNext());
                    Clipboard.SetText(sb.ToString());
                    Xceed.Wpf.Toolkit.MessageBox.Show("营销员数据已拷贝到剪贴板,","成功", MessageBoxButton.OK, MessageBoxImage.Information);
                });
            }
        }
        public ICommand SaveToCacheBoxCommand
        {
            get
            {
                return  new DelegateCommand (() =>
                { 
                    foreach (var row in mSelectedRows)
                    { 
                        var guid = (Guid)row[BillSheetColumns.SYS_GUID];
                        var ex = false;
                        foreach(DataRow r in CacheDataTable.Rows)
                        {
                            if(((Guid)r[BillSheetColumns.SYS_GUID]).Equals(guid))
                            {
                                ex = true;
                                break;
                            }
                        }
                        if(ex == false)
                            CacheDataTable.ImportRow(row);
                    }
                });
            }
        }

        public ICommand DelCachedBillsCommand
        {
            get
            {
                return new DelegateCommand(() =>
                {
                    var list = mSelectedCachedRows.ToArray();
                    foreach(var dr in list)
                    {
                        CacheDataTable.Rows.Remove(dr);
                    }
                     
                });
            }
        }
        public List<DataRow> GetChangedRows()
        {
            var rows = new List<DataRow>();
            foreach (DataRow row in DataTable.Rows)
            {
                var oldServiceId = "";
                var newServiceId = "";
                if (!row.IsNull(BillSheetColumns.CURRENT_SERVICE_ID))
                    newServiceId = (string)row[BillSheetColumns.CURRENT_SERVICE_ID];
                if (!row.IsNull(BillSheetColumns.SYS_SERVICE_ID))
                    oldServiceId = (string)row[BillSheetColumns.SYS_SERVICE_ID];
                if (oldServiceId != newServiceId)
                {
                    rows.Add(row);
                }
            }

            return rows;
        }
        public ICommand ExportGCommand
        {
            get
            {
                return new DelegateCommand(() =>
                {
                    var rows = GetChangedRows(); 
                    ExportDBills(rows);
                });
            }
        }

        public ICommand AddSelectAllLayoutCommand
        {
            get
            {
                return new DelegateCommand(() =>
                { 
                    foreach(var col  in Columns)
                    {
                        col.IsVisible = true;
                    }
                });
            }
        }

        public ICommand AddUnSelectAllLayoutCommand
        {
            get
            {
                return new DelegateCommand(() =>
                {
                    foreach (var col in Columns)
                    {
                        col.IsVisible = false;
                    }
                });
            }
        }


        public ICommand ExportECommand
        {
            get
            {
                return new DelegateCommand(ExportTypeEBills);
            }
        }
        public ICommand DoAssignCommand
        {
            get
            {
                return new DelegateCommand(AssignBills);
            }
        }
        public ICommand CopyBillContentCommand
        {
            get
            {
                if(mCopyBillContentCommand == null)
                {
                    mCopyBillContentCommand = new DelegateCommand<string>((c) =>
                    {
                        var row = this.mCurrentSelectedRows.FirstOrDefault();
                        if(row == null)
                            return;
                        if(!row.IsNull(c))
                        {
                            var value = Convert.ToString(row[c]);
                            Clipboard.SetText(value);
                        }
                        else
                        {
                            Clipboard.SetText("空");
                        }

                    });
                }

                return mCopyBillContentCommand;
            }
        }
        private DataTable mCurrentDataTable;
        private DataTable CurrentDataTable
        {
            get
            {
                if(mCurrentDataTable == null)
                {
                    mCurrentDataTable = DataTable;
                }

                return mCurrentDataTable;
            }
        }
        public ICommand SelectTabItemCommand
        {
            get
            {
                if(mSelectTabItemCommand == null)
                {
                    mSelectTabItemCommand = new DelegateCommand<object>((i) =>
                    {
                        var inti = Convert.ToInt32(i);
                        if(inti == 0)
                        {
                            mCurrentDataTable = DataTable;
                            mCurrentSelectedRows = SelectedRows;
                        }

                        if(inti == 1)
                        {
                            mCurrentDataTable = CacheDataTable;
                            mCurrentSelectedRows = mSelectedCachedRows;
                        }
                    });
                }

                return mSelectTabItemCommand;
            }
        }
        #endregion
    }



    public class UIColumnData : BindableBase
    {
        private bool mIsVisible;
        public string Name { get; set; }
        public bool IsVisible
        {
            get => mIsVisible;
            set => SetProperty(ref mIsVisible, value);
        }
        public int Sort { get; set; }
        public int Index { get; set; }
    }
    public class FilterData
    {
        public string Name { get; set; }
        public int Count { get; set; }
        public FilterType ShowType { get; set; }
    }
    public enum FilterType
    {
        ShowAll,
        ShowFilter,
        ShowNoFilter

    }
}
