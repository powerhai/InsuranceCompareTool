using Prism.Commands;
using Prism.Mvvm;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;
using System.Xml.Serialization;
using InsuranceCompareTool.Core;
using InsuranceCompareTool.Models;
using InsuranceCompareTool.Models.Dispatch;
using InsuranceCompareTool.Properties;
using InsuranceCompareTool.Views;
using Microsoft.Win32;
using Prism.Interactivity.InteractionRequest;
using Unity;
namespace InsuranceCompareTool.ViewModels
{
    public class PolicyViewViewModel : TabViewModelBase
    {
        private readonly IUnityContainer mContainer;
        private ProjectCacheHelper mProjectCacheHelper = new ProjectCacheHelper();
        public override string Title { get; set; } = "分单方案";

        public InteractionRequest<EditProjectNotification> EditProjectNotification { get; } = new InteractionRequest<EditProjectNotification>();

        public PolicyViewViewModel(IUnityContainer container)
        {
            mContainer = container; 
        }
        private bool mIsLoaded = false;
        private void LoadData()
        {
            if(mIsLoaded == true)
                return;
            mIsLoaded = true;
            if(string.IsNullOrEmpty( Settings.Default.Projects) )
                return;
            try
            {
                var list = mProjectCacheHelper.GetProjects();
                foreach(var v in list)
                {
                    mProjects.Add(v);
                }
            }
            catch(Exception ex)
            {
                //Console.WriteLine(ex); 
            }

           
        }

        #region Commands
        public ICommand AddCommand
        {
            get
            {
                return new DelegateCommand(() =>
                { 
                    EditProjectNotification.Raise(new EditProjectNotification()
                    {
                        Title = "Edit project",
                        Project = new Project() { Title = $"新的分单方案 {DateTime.Now.ToString("g")}" }
                    }, r =>
                    {
                        if (r.Confirmed == true)
                        {
                            this.mProjects.Add(r.Project);
                            this.ProjectView.MoveCurrentToLast();
                             
                        }
                    }); 
                });
            }
        }
        public ICommand UpCommand
        {
            get
            {
                return new DelegateCommand<Project>((p) =>
                {
                    var index = mProjects.IndexOf(p);
                    if(index > 0)
                    {
                        mProjects.Move(index, index - 1);
                         
                    }

                    ProjectView.MoveCurrentTo(p);
                });
            }
        }
        public ICommand DownCommand
        {
            get
            {
                return new DelegateCommand<Project>((p) =>
                {
                    var index = mProjects.IndexOf(p);
                    if(index < mProjects.Count - 1)
                    {
                        mProjects.Move(index, index + 1);
                        
                    }

                    ProjectView.MoveCurrentTo(p);

                });
            }
        }
        public ICommand EditCommand
        {
            get
            {
                return new DelegateCommand<Project>((p) =>
                {
                    p.BeginEdit(); 
                    EditProjectNotification.Raise(new EditProjectNotification()
                    {
                        Title = "Edit project",
                        Project = p
                    }, r =>
                    {
                        if (r.Confirmed == true)
                        {
                            p.EndEdit();
                             
                        }
                        else
                        { 
                            p.CancelEdit();
                        }
                    });
                });
            }
        }
        public ICommand CopyCommand
        {
            get
            {
                return new DelegateCommand<Project>((p) =>
                    {
                        using (MemoryStream ms = new MemoryStream())
                        {
                            XmlSerializer xml = new XmlSerializer(typeof(Project));
                            xml.Serialize(ms,p);
                            ms.Seek(0, SeekOrigin.Begin);
                            var obj =xml.Deserialize(ms) as Project;
                            obj.Title += " - 2";
                            obj.Guid = Guid.NewGuid();
                            this.mProjects.Add(obj); 
                        }
                         
                    });
            }
        }

        public ICommand DelCommand
        {
            get
            {
                return new DelegateCommand<Project>((p) =>
                {
                    mProjects.Remove(p);
                    SaveCommand.Execute(null);
                });
            }
        }
        public ICommand SaveCommand
        {
            get
            {
                return new DelegateCommand(() =>
                {
                    if (mIsLoaded == false)
                        return;
                    try
                    {

                        mProjectCacheHelper.SaveProjects(mProjects.ToList()); 
                    }
                    catch(Exception ex)
                    {
                        MessageBox.Show(ex.ToString(), "保存失败", MessageBoxButton.OK, MessageBoxImage.Error);
                    }

                });
            }
        }
        public ICommand ImportCommand => new DelegateCommand(Import);
        public ICommand ExportCommand => new DelegateCommand(Export);
        #endregion

        private void Import()
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "Haiser Files (*.hai)|*.hai", 
            };

            var result = openFileDialog.ShowDialog();
            if(result == true)
            {
                var json = "";
                using(var c = openFileDialog.OpenFile())
                {
                    var bytes = new byte[c.Length];
                    c.Read(bytes, 0, bytes.Length);
                    json = System.Text.Encoding.UTF8.GetString(bytes);
                }
                 
                var rv = mProjectCacheHelper.SaveProjects(json);
                if(rv)
                {
                    Xceed.Wpf.Toolkit.MessageBox.Show("导入成功", "成功", MessageBoxButton.OK, MessageBoxImage.Information);
                    mIsLoaded = false;
                    mProjects.Clear();
                    LoadData();
                }
                else
                {
                    Xceed.Wpf.Toolkit.MessageBox.Show("导入失败", "失败", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
        private void Export()
        {
            var saveFile = new SaveFileDialog()
            {
                FileName = $"分单策略 - {DateTime.Now.ToString("MM-dd-hh-ss")}.hai"
            }; 
            var result = saveFile.ShowDialog();
            if(result == true)
            {
                try
                {
                     using (var c = saveFile.OpenFile())
                    {
                        var bytes = System.Text.Encoding.UTF8.GetBytes(Settings.Default.Projects);
                        c.Write(bytes,0, bytes.Length); 
                    }                    
                }
                catch(Exception ex)
                {
                    Xceed.Wpf.Toolkit.MessageBox.Show(ex.Message, "导出失败", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                } 
            }
        }
        private ListCollectionView mProjectView;
        private ObservableCollection<Project> mProjects = new ObservableCollection<Project>();
        public ListCollectionView ProjectView
        {
            get
            {
                if(mProjectView == null)
                {
                    mProjectView = new ListCollectionView(mProjects); 
                } 
                return mProjectView;
            }
        }
        public override void Leave()
        {
            SaveCommand.Execute(null);
        }
        public override void Close()
        {
            SaveCommand.Execute(null);
        }
        public override void Enter()
        {
            LoadData();
        }
    }

    public class EditProjectNotification : IConfirmation
    {
        public string Title { get; set; }
        public object Content { get; set; }
        public bool Confirmed { get; set; }
        public Project Project { get; set; }
    }
}
