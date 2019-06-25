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
using InsuranceCompareTool.Models;
using InsuranceCompareTool.Models.Dispatch;
using InsuranceCompareTool.Properties;
using InsuranceCompareTool.Services;
using DataGrid = System.Windows.Controls.DataGrid;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;
namespace InsuranceCompareTool.ViewModels
{
    public class AssignViewViewModel : TabViewModelBase
    {
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
        private ObservableCollection<Member> mMembers = new ObservableCollection<Member>();
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
        public ObservableCollection<Member> Members
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
                    if (!row.IsNull("本期应缴保费"))
                    {
                        double bf = 0;
                        double.TryParse((string)row["本期应缴保费"], out bf);
                        val += bf;
                    }
                }
                return val.ToString("C");
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

        public CollectionViewSource ProjectView => new CollectionViewSource() { Source = mProjects };
        #endregion
        #region Methods
        private bool FilterMember(object obj)
        {
            var area = AreaView.View.CurrentItem as string;
            if(string.IsNullOrEmpty(area))
            {
                return true;
            }

            var ob = obj as Member;
            return ob.Area == area;
        }
        private void View_CurrentChanged(object sender, EventArgs e)
        {
            var filter = mFilterView.View.CurrentItem as FilterData;
            if (filter == null)
                return;
            if (filter.ShowAll == true)
                DataTable.DefaultView.RowFilter = "";
            else
                DataTable.DefaultView.RowFilter = $"{BillSheetColumns.SYS_FILTER} = '{filter.Name}'";
        }
        private void OnCurrentAreaChanged(object sender, EventArgs e)
        {  
            MemberView.View.Refresh();
            MemberView.View.MoveCurrentToFirst();
        }
        public AssignViewViewModel()
        {
        }
        private void CacheLayout()
        {
            if (mLastLayout == null || Columns == null)
                return;
            mLastLayout.Columns = this.Columns.Where(a => a.IsVisible == true).Select(a => a.Name).ToArray();
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
                ShowAll = true

            });
            foreach (var step in project.Steps)
            {
                var filter = new FilterData() { Name = step.Title };
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
        private void LoadProjects()
        {
            this.Projects.Clear();
            var list = mProjectCacheHelper.GetProjects();
            foreach (var project in list)
            {
                this.Projects.Add(project);
            }

            this.ProjectView.View.MoveCurrentToFirst();
        }
        private async Task LoadMembers()
        { 
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
            foreach(var  mem in mems)
            {
                Members.Add(mem);
            }

            MemberView.View.MoveCurrentToFirst();
        }
        private void MakeColumns(DataTable dt)
        {
            var list = new List<UIColumnData>();
            var curLayout = LayoutView.View.CurrentItem as ColumnVisible;

            foreach (DataColumn dataColumn in dt.Columns)
            {
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
        private async void AssignBills()
        {
            var cur = ProjectView.View.CurrentItem as Project;
            if (cur == null)
                return;
            var assignHelper = new BillAssignHelper();
            IsBusy = true;
            await Task.Run(() =>
            {
                assignHelper.AssignBills(DataTable, cur);

            });
            GenerateFilters(cur);
            IsBusy = false;
            Xceed.Wpf.Toolkit.MessageBox.Show("分单完成", "Finished", MessageBoxButton.OK, MessageBoxImage.Information);


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
                    var str = "";
                    var cols = Columns.Where(a => a.IsVisible == true);
                    foreach(var col in cols)
                    {
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
                        var loader = new BillLoadHelper();
                        SourceFile = openFileDialog.FileName;
                        IsBusy = true;
                        try
                        {

                            DataTable dt = await Task<DataTable>.Run(() =>
                            {
                                return loader.Load(openFileDialog.FileName);
                            });
                            MakeColumns(dt);
                            DataTable = dt;

                        }
                        catch (Exception ex) { }
                        finally
                        {
                            IsBusy = false;
                        }

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


        public ICommand ExportDCommand
        {
            get
            {
                return new DelegateCommand(ExportTypeDBills);
            }
        }
        public ICommand AssignToSomebodyCommand
        {
            get
            {
                return mAssignToSomebodyCommand?? (mAssignToSomebodyCommand = new DelegateCommand<IEnumerable<object>>((mmm) =>
                {
                    var mem = MemberView.View.CurrentItem as Member;
                    if(mem == null)
                        return;
                    var list = mmm.Cast<DataRowView>();
                    foreach(var row in list)
                    {  
                        row.Row[BillSheetColumns.CURRENT_SERVICE_NAME] = mem.Name;
                        row.Row[BillSheetColumns.CURRENT_SERVICE_ID] = mem.ID;
                    }
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
        public bool ShowAll { get; set; }
    }
}
