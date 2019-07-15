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
using log4net;
using DataGrid = System.Windows.Controls.DataGrid;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;
namespace InsuranceCompareTool.ViewModels
{
    public class MemberA : BindableBase
    {
        public string Name
        {
            get => mName;
            set => SetProperty(ref mName, value);
        }
        private int mCount;
        private string mName;
        private string mArea;
        private double mPrice;
        public int Count
        {
            get => mCount;
            set {
                SetProperty(ref mCount, value);
                RaisePropertyChanged(nameof(Title));
                RaisePropertyChanged(nameof(Title2));
            }
        }
        public double Price
        {
            get => mPrice;
            set
            {
                SetProperty(ref mPrice, value);
                RaisePropertyChanged(nameof(Title));
                RaisePropertyChanged(nameof(Title2));
            }
        }
        public string Area
        {
            get => mArea;
            set => SetProperty(ref mArea, value);
        }
        public string Title
        {
            get
            {
                var name = Name.PadRight(6, ' ');
                var count = Count.ToString().PadLeft(5, ' ');
                var price = Price.ToString("C").PadLeft(12, ' ');
                return $"{name}\t{count}\t{price}\t{ID}";
            }
        }

        public string Title2
        {
            get
            {
                var name = Name.PadRight(6, ' ');
                var count = Count.ToString().PadLeft(5, ' ');
                var price = Price.ToString("C").PadLeft(12, ' ');
                return $"{name}\t{count}\t{price}\t{ID}\t{Area}";
            }
        }

        public string ID { get; set; }
    }
    public class AssignViewViewModel : TabViewModelBase
    {
        private readonly ILog mLogger;
        #region Fields
        private const string SHOW_ALL = "显示全部";
        private ColumnLayoutHelper mColumnLayoutHelper = new ColumnLayoutHelper();
        private ProjectCacheHelper mProjectCacheHelper = new ProjectCacheHelper();
        private ObservableCollection<ColumnVisible> mColumnLayouts = new ObservableCollection<ColumnVisible>();
        private readonly ObservableCollection<Project> mProjects = new ObservableCollection<Project>();

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
        private List<string> mAreas = new List<string>(){AreaNames.JH, AreaNames.YK, AreaNames.YW, AreaNames.WY, AreaNames.PJ, AreaNames.LX, AreaNames.QT, AreaNames.DY};
        private CollectionViewSource mAreaView;
        private ObservableCollection<MemberA> mMembers = new ObservableCollection<MemberA>();
        private CollectionViewSource mProjectView;
        private ICollectionView mDataTableView;
        private List<Member> mAllMembers;
        private ICommand mSelectBillCommand;

        private ObservableCollection<MemberA> mSelectedMembers = new ObservableCollection<MemberA>();
        private ICommand mAssignToServiceCommand;
        private CollectionViewSource mSelectMemberView;
        public ObservableCollection<MemberA> SelectedMembers => mSelectedMembers;
        #endregion
        #region Properties
        public List<string> Areas
        {
            get => mAreas; 
        }
        public CollectionViewSource AreaView
        {
            get
            {
                if(mAreaView == null)
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
                if(mSelectMemberView == null)
                {
                    mSelectMemberView = new CollectionViewSource()
                    {
                        Source =  mSelectedMembers
                    };
                    
                }

                return mSelectMemberView;
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
                    mMemberView.SortDescriptions.Add(new SortDescription("Price", ListSortDirection.Ascending));
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
                    mFilterView.View.CurrentChanged += View_CurrentChanged;
                }
                return mFilterView;
            }
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
                        double.TryParse( row[BillSheetColumns.BILL_PRICE].ToString(), out bf);
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
                if(mDataTableView == null)
                {
                    mDataTableView = CollectionViewSource.GetDefaultView(DataTable);
                }

                return mDataTableView;
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
                if(mProjectView == null)
                {
                    mProjectView = new CollectionViewSource() { Source = mProjects };
                }

                return mProjectView;
            }
        }
        #endregion
        #region Methods
        private bool FilterMember(object obj)
        {
            var area = AreaView.View.CurrentItem as string;
            if(string.IsNullOrEmpty(area))
            {
                return true;
            }

            var ob = obj as MemberA;
            return ob.Area == area;
        }
        private void View_CurrentChanged(object sender, EventArgs e)
        {
            var filter = mFilterView.View.CurrentItem as FilterData;
            if (filter == null)
                return;
            if (filter.ShowType == FilterType.ShowAll)
                DataTable.DefaultView.RowFilter = "";
            else if (filter.ShowType == FilterType.ShowFilter)
            {
                DataTable.DefaultView.RowFilter = $"{BillSheetColumns.SYS_FILTER} = '{filter.Name}'";
            }else
            {
                DataTable.DefaultView.RowFilter = $"{BillSheetColumns.SYS_FILTER} = '' "; 
            }
             
        }
        private void OnCurrentAreaChanged(object sender, EventArgs e)
        {  
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
            foreach(DataRow dr in DataTable.Rows)
            {
                if((string)dr[BillSheetColumns.SYS_FILTER] == "")
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
                var filter = new FilterData() { Name = step.Title , ShowType = FilterType.ShowFilter};
                var count = 0;
                foreach (DataRow row in DataTable.Rows)
                {
                    if (row[BillSheetColumns.SYS_FILTER] != DBNull.Value && (string)row[BillSheetColumns.SYS_FILTER] == step.Title)
                    {
                        count += 1;
                    }
                }

                filter.Count = count;
                mFilters.Add(filter);
            }

        }
        private async void ExportTypeDBills()
        {
            var exporter = new BillExportTypeD();
            var targetPath = $"{Settings.Default.WorkPath}\\上传保单\\{DateTime.Now:yyyy-MM-dd--ss-mm}";
            IsBusy = true;
            await Task.Run(() =>
            {
                try
                {
                    exporter.Export(targetPath, DataTable, Columns.Where(a => a.IsVisible).Select(a => a.Name).ToList());

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
            if(curProject != null)
            {
                var pro = this.Projects.FirstOrDefault(a => a.Guid.Equals(curProject.Guid));
                if(pro != null)
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
            if(!fi.Exists)
                return;

            if(mIsLoaded == true)
            {
                if(   fi.LastWriteTime == mLastMemberFileWriteTime)
                    return;
            }

            mIsLoaded = true;

            mLastMemberFileWriteTime = fi.LastWriteTime;

            var memService = new MemberService();
            await Task.Run(() => {  memService.Load(Settings.Default.MembersFile);});
            Members.Clear();
            var mems = memService.GetZaiziMembers().Where(a=>a.Position == "客服专员");
            mAllMembers = memService.GetMembers().ToList();
            foreach (var  mem in mems)
            {
                Members.Add(new MemberA(){ ID = mem.ID, Name = mem.Name, Count =  0 , Area = mem.Area });
            }

            MemberView.View.MoveCurrentToFirst();
        }
        private void MakeColumns(DataTable dt)
        {
            var list = new List<UIColumnData>();
            var curLayout = LayoutView.View.CurrentItem as ColumnVisible;
            var sysCols = BillTableColumns.Columns.Where(a => a.IsSystemColumn).Select(a=>a.Name).ToArray();

            foreach (DataColumn dataColumn in dt.Columns)
            {
                if(sysCols.Contains(dataColumn.ColumnName))
                    continue;
                var col = new UIColumnData()
                {
                    Name = dataColumn.ColumnName,
                    IsVisible = true
                };
                if (curLayout != null)
                {
                    col.IsVisible = curLayout.Columns.Contains(col.Name);
                }
                list.Add(col);
            }

            Columns = list;
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
            await  Task.Run(() =>
            {
                try
                {
                    var templateService = new ExportTemplateService();
                    var billLoader = new BillLoadService();
                    var memberService =new MemberService();
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
                catch (InvalidOperationException e1)
                {
                    MessageBox.Show("操作不能继续：文件正被使用!", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            });
             
            IsBusy = false;
        }
        private void CheckDatatable(DataTable dt)
        {
            string message = "表格中存在在职状态未知的营销员或服务专员";
            foreach(DataRow dr in dt.Rows)
            {
                if(dr.IsNull(BillSheetColumns.SELLER_STATE))
                {
                    Xceed.Wpf.Toolkit.MessageBox.Show(message, "错误", MessageBoxButton.OK, MessageBoxImage.Warning);
                    break;
                }

                if(dr.IsNull(BillSheetColumns.SYS_SERVICE_STATUS))
                {
                    Xceed.Wpf.Toolkit.MessageBox.Show(message, "错误", MessageBoxButton.OK, MessageBoxImage.Warning);
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

        private async void AssignBills()
        {
            mLogger.Info("开始分单");
            var cur = ProjectView.View.CurrentItem as Project;
            if (cur == null)
                return;
            
            var assignHelper = new BillAssignHelper(mLogger,mAllMembers);
            IsBusy = true;
            var ex = await Task.Run<Exception>(() =>
            {
                try
                {
                    assignHelper.AssignBills(DataTable, cur);
                }
                catch(Exception e)
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

            
        }
        private void CalculateCountForMembers(DataTable dt)
        {
            foreach(MemberA member in Members)
            {
                member.Count = 0;
                member.Price = 0;
            }

            bool hasPriceCol = dt.Columns.Contains(BillSheetColumns.BILL_PRICE);

            if(dt.Columns.Contains(BillSheetColumns.CURRENT_SERVICE_ID))
            {
                foreach(DataRow row in dt.Rows)
                {
                    if(row.IsNull(BillSheetColumns.CURRENT_SERVICE_ID)) 
                        continue;
                    var serID = (string) row[BillSheetColumns.CURRENT_SERVICE_ID];
                    var mem = Members.FirstOrDefault(a =>
                        a.ID.Equals(serID, StringComparison.CurrentCultureIgnoreCase));
                    if(mem == null)
                        continue;
                    if(hasPriceCol)
                    {
                        if(!row.IsNull(BillSheetColumns.BILL_PRICE))
                        {
                            var price = (double) row[BillSheetColumns.BILL_PRICE];
                            mem.Price += price; 
                        }
                    }
                    mem.Count++;
                }
            }

            var firMem = Members.FirstOrDefault(a => a.Count > 0);
            if(firMem != null)
            {
                var area = this.Areas.FirstOrDefault(a => a == firMem.Area);
                this.AreaView.View.MoveCurrentTo(area);

            }
            MemberView.View.Refresh();
        }
        #endregion
        #region Commands

        public ICommand SearchCommand
        {
            get
            {
                return new DelegateCommand<string>((s) =>
                {
                    this.FilterView.View.MoveCurrentTo(null);
                    var unstrCols = BillTableColumns.Columns.Where(a => a.Type != typeof(String)).Select(a=>a.Name).ToArray();
                    var str = "";
                    var cols = Columns.Where(a => a.IsVisible == true);
                    foreach(var col in cols)
                    {
                        if(unstrCols.Contains(col.Name))
                        {
                            continue;
                        }

                        if(!string.IsNullOrEmpty(str))
                        {
                            str += " or ";
                        }

                        str += $"{col.Name} like '%{s}%'";
                    }

                    if(string.IsNullOrEmpty(s))
                    {
                        DataTable.DefaultView.RowFilter = "";
                    }
                    else
                    {
                        DataTable.DefaultView.RowFilter = str;
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

                            DataTable dt = await Task<DataTable>.Run(() =>
                            {
                                return loader.Load(openFileDialog.FileName);
                            });
                            
                            var memhelp = new MemberService();
                            memhelp.Load(Settings.Default.MembersFile);
                            var mems = memhelp.GetMembers();
                            var formater = new BillFormatHelper(mLogger, mems);
                            formater.Format(dt);
                            CalculateCountForMembers(dt);
                            CheckDatatable(dt);
                            MakeColumns(dt);
                            DataTable = dt;

                        }
                        catch (Exception ex) { }
                        finally
                        {
                            mFilters.Clear();
                            IsBusy = false;
                        }

                    }
                });
            }
        }
        public ICommand SelectBillCommand
        {
            get
            {
                if(mSelectBillCommand == null)
                {
                    mSelectBillCommand = new DelegateCommand<ObservableCollection<object>>((b) =>
                    {
                        this.SelectedMembers.Clear();
                        var list  = new System.Collections.Generic.Dictionary<string,MemberA>();
                        foreach(DataRowView row in b)
                        {
                            var bill = row.Row;
                            if(!bill.IsNull(BillSheetColumns.CURRENT_SERVICE_ID))
                            {
                                var sid = Convert.ToString(bill[BillSheetColumns.CURRENT_SERVICE_ID]);
                                if(list.ContainsKey(sid))
                                    continue;
                                var mem = mMembers.FirstOrDefault(a => a.ID.Equals(sid));
                                if(mem == null)
                                    continue;
                                //if(mem.Status != StatusNames.ZAI_ZI)
                                //    continue;
                                //list[sid]  = new MemberA(){ ID = sid, Area = mem.Area, Name = mem.Name };
                                list[sid] = mem;

                            }
                        }
                        if(list.Count <=1)
                            return;
                        foreach(KeyValuePair<string, MemberA> mem in list)
                        {
                            this.SelectedMembers.Add(mem.Value);
                        }

                        this.SelectMemberView.View.MoveCurrentToFirst();
                    });
                }

                return mSelectBillCommand;
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
                return new DelegateCommand(ExportTypeDBills);
            }
        }
        public ICommand ExportECommand
        {
            get
            {
                return new DelegateCommand(ExportTypeEBills);
            }
        }
        public ICommand AssignToServiceCommand
        {
            get
            {
                if(mAssignToServiceCommand == null)
                {
                    mAssignToServiceCommand = new DelegateCommand<IEnumerable<object>>((m) =>
                    { 

                        var mem = SelectMemberView.View.CurrentItem as MemberA;
                        if(mem == null)
                            return;

                        foreach(DataRowView dr in m)
                        { 
                            var row = dr.Row;
                            row[BillSheetColumns.CURRENT_SERVICE_NAME] = mem.Name;
                            row[BillSheetColumns.CURRENT_SERVICE_ID] = mem.ID;
                            LogAssgign(row,mem);
                        }
                        this.CalculateCountForMembers(DataTable);

                    });
                }

                return mAssignToServiceCommand;
            } 
        }
        private void AssignTo(MemberA mem)
        {
            
        }
        public ICommand AssignToSomebodyCommand
        {
            get
            {
                return mAssignToSomebodyCommand?? (mAssignToSomebodyCommand = new DelegateCommand<IEnumerable<object>>((mmm) =>
                {
                    var mem = MemberView.View.CurrentItem as MemberA;
                    if(mem == null)
                        return;
                    var list = mmm.Cast<DataRowView>();
                    foreach(var row in list)
                    {  

                        row.Row[BillSheetColumns.CURRENT_SERVICE_NAME] = mem.Name;
                        row.Row[BillSheetColumns.CURRENT_SERVICE_ID] = mem.ID;
                        LogAssgign(row.Row, mem);
                    }
                    this.CalculateCountForMembers(DataTable);
                }));
            }
            
        }
        public ICommand DoAssignCommand
        {
            get
            {
                return new DelegateCommand(AssignBills);
            }
        }
        #endregion

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
    }
    public class FilterData
    {
        public string Name { get; set; }
        public int Count { get; set; }
        public FilterType ShowType { get; set; }
    }
    public enum  FilterType
    {
        ShowAll,
        ShowFilter,
        ShowNoFilter

    }
}
