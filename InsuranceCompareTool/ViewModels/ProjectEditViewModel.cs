using Prism.Commands;
using Prism.Mvvm;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Data;
using System.Windows.Input;
using InsuranceCompareTool.Domain;
using InsuranceCompareTool.Models;
using InsuranceCompareTool.Models.Dispatch;
using InsuranceCompareTool.ShareCommon;
using Prism.Interactivity.InteractionRequest;
namespace InsuranceCompareTool.ViewModels
{
    public class ProjectEditViewModel : ViewModelBase, IInteractionRequestAware
    {
        private EditProjectNotification mNotification;
        public override string Title { get; set; }
        public INotification Notification
        {
            get => mNotification;
            set
            {
                SetProperty(ref mNotification, (EditProjectNotification)value);
                Project = mNotification.Project;
            }
        } 
        public Action FinishInteraction { get; set; }
        public List<EnumPackager> BoolTypeList =>
        new List<EnumPackager>()
        {
            new EnumPackager() { Enum = BoolType.And },
            new EnumPackager() { Enum = BoolType.Or }
        };
        public List<EnumPackager> ColumnTypes => new List<EnumPackager>()
        {
            new EnumPackager(){ Enum = LogicColumnName.ServiceID },
            new EnumPackager(){ Enum = LogicColumnName.ServiceStatus },
            new EnumPackager(){ Enum = LogicColumnName.SellerStatus }, 
            new EnumPackager(){ Enum = LogicColumnName.CustomerPassportId }, 
            new EnumPackager(){ Enum = LogicColumnName.CurrentPrice }, 
            new EnumPackager(){ Enum = LogicColumnName.PAY_ADDRESS },
            new EnumPackager(){ Enum = LogicColumnName.PAY_ADDRESS2 },
            new EnumPackager(){ Enum = LogicColumnName.BillID },
        };
        
        public List<EnumPackager> OperatorTypes
        {
            get
            {
                if(mOperatorTypes == null)
                {
                    mOperatorTypes = new List<EnumPackager>()
                    {
                        new EnumPackager(){ Enum =  OperatorType.Equal},
                        new EnumPackager(){ Enum = OperatorType.Greater },
                        new EnumPackager(){ Enum = OperatorType.Less } ,
                        new EnumPackager(){ Enum = OperatorType.Contain } ,
                    };
                } 
                return mOperatorTypes;
            }
        }
        public List<EnumPackager> OperatorTypesForNumber {
            get
            {
                if(mOperatorTypesForNumber == null)
                {
                    mOperatorTypesForNumber = new List<EnumPackager>()
                    {
                        new EnumPackager(){ Enum =  OperatorType.Equal},
                        new EnumPackager(){ Enum = OperatorType.Greater },
                        new EnumPackager(){ Enum = OperatorType.Less } ,
                    };
                } 
                return mOperatorTypesForNumber;
            }
        }
        public List<EnumPackager> DispatchTypes => new List<EnumPackager>()
        {
            new EnumPackager(){Enum = DispatchType.DispatchToAreaManager},
            new EnumPackager(){Enum = DispatchType.DispatchToDesignated},
            new EnumPackager(){Enum = DispatchType.DispatchToLastService},
            new EnumPackager(){Enum = DispatchType.DispatchToManagerOfCustomer},
            new EnumPackager(){Enum = DispatchType.DispatchToManagerOfSeller}, 
        };
        public ProjectEditViewModel()
        { 
        } 
        private Project mProject;
        private List<EnumPackager> mOperatorTypesForNumber;
        private List<EnumPackager> mOperatorTypes;
        public Project Project
        {
            get => mProject;
            private set => SetProperty(ref mProject, value);
        }

        public ICommand SaveCommand
        {
            get
            {
                return new DelegateCommand(() =>
                {
                    mNotification.Confirmed = true;
                    FinishInteraction.Invoke();
                });
            }
        }
        public ICommand CancelCommand
        {
            get
            {
                return new DelegateCommand(() =>
               {
                   mNotification.Confirmed = false;
                   FinishInteraction.Invoke();
               });
            }
        }

        public ICommand AddStepCommand
        {
            get
            {
                return new DelegateCommand(() =>
                {
                    var step = new Step() { Title = $"步骤 {Project.Steps.Count + 1}" };
                    Project.Steps.Add(step);
                    Project.StepView.View.MoveCurrentToLast();
                    ResetStepsIndex();
                });
            }
        }
        public ICommand UpStepCommand
        {
            get
            {
                return new DelegateCommand<Step>((s) =>
                {
                    var index = Project.Steps.IndexOf(s);
                    if (index > 0)
                    {
                        Project.Steps.Move(index, index - 1);
                        Project.StepView.View.MoveCurrentTo(s);
                        ResetStepsIndex();
                    }

                });
            }
        }
        public ICommand DownStepCommand
        {
            get
            {
                return new DelegateCommand<Step>((s) =>
                {
                    var index = Project.Steps.IndexOf(s);
                    if (index < Project.Steps.Count - 1)
                    {
                        Project.Steps.Move(index, index + 1);
                        Project.StepView.View.MoveCurrentTo(s);
                        ResetStepsIndex();
                    }

                });
            }
        }
        public ICommand DelStepCommand
        {
            get
            {
                return new DelegateCommand<Step>((s) =>
                {
                    //Project.Steps.Remove(s);
                    Project.Steps.Remove(s);
                    ResetStepsIndex();
                });
            }
        }
        public ICommand AddFilterCommand
        {
            get
            {
                return new DelegateCommand(() =>
               {
                   var step = Project.StepView.View.CurrentItem as Step;
                   if (step == null)
                       return;
                   step.Filters.Add(new Filter()
                   {

                   });
               });
            }
        }

        public ICommand DelFilterCommand
        {
            get
            {
                return new DelegateCommand<Filter>((f) =>
               {
                   var step = Project.StepView.View.CurrentItem as Step;
                   if (step == null)
                       return;
                   step.Filters.Remove(f);
               });
            }
        }
        private void ResetStepsIndex()
        {
            for (int i = 0; i < Project.Steps.Count; i++)
            {
                var step = Project.Steps[i];
                step.Index = i + 1;
            }
        }
    }


}
