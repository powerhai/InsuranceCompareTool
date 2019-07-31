using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;
using InsuranceCompareTool.ShareCommon;
using Prism.Mvvm;
using System.Runtime.Serialization;
using System.Xml.Serialization;
using InsuranceCompareTool.Domain;
namespace InsuranceCompareTool.Models.Dispatch
{

    [Serializable]
    public class Project : EditableObject
    { 
        public Guid Guid { get; set; } = Guid.NewGuid();

        private string mTitle;
        [NonSerialized()]
        private CollectionViewSource mStepView;

        private   ObservableCollection<Step> mSteps = new ObservableCollection<Step>();

        public string Title
        {
            get => mTitle;
            set => SetProperty(ref mTitle, value);
        }

        public ObservableCollection<Step> Steps => mSteps;

        public CollectionViewSource StepView =>  (mStepView??(mStepView = new CollectionViewSource( ){Source = Steps}));
        public override void CancelEdit()
        { 
            Title = mBackup.Title;
            mSteps = mBackup.Steps; 
            OnPropertyChanged(nameof(Steps));
            StepView.Source = Steps; 
        }
        private Project mBackup;
        public override void BeginEdit()
        {
            mBackup = DeepCopyByXml(this); 
        }
        public override void EndEdit()
        {
            mBackup = null;
        }


        private Project DeepCopyByXml(Project obj)
        {
            object retval;
            using (MemoryStream ms = new MemoryStream())
            {
                XmlSerializer xml = new XmlSerializer(typeof(Project));
                xml.Serialize(ms, obj);
                ms.Seek(0, SeekOrigin.Begin);
                retval = xml.Deserialize(ms);
                ms.Close();
            }
            return (Project)retval;
        }
    }

    [Serializable]
    public class Step : EditableObject
    {
        public int Index
        {
            get => mIndex;
            set => SetProperty(ref mIndex,value);
        }
        private string mTitle; 
        private int mIndex;
        private DispatchType? mDispatchType;
        private string mDispatchDesignated;
        public string Title
        {
            get => mTitle;
            set => SetProperty(ref mTitle, value);
        }

        public ObservableCollection<Filter> Filters { get; } = new ObservableCollection<Filter>();
        public DispatchType? DispatchType
        {
            get => mDispatchType;
            set
            {
                if (mDispatchType == Domain.DispatchType.DispatchToDesignated && value != Domain.DispatchType.DispatchToDesignated)
                {
                    mBackDispatchDesignated = DispatchDesignated;
                    DispatchDesignated = "";
                }
                if (value == Domain.DispatchType.DispatchToDesignated && mDispatchType != Domain.DispatchType.DispatchToDesignated)
                {
                    DispatchDesignated = mBackDispatchDesignated;
                }
                SetProperty(ref mDispatchType, value);

            }
        }
        [NonSerialized]
        private string mBackDispatchDesignated;
        private bool mStopNextStep;
        public string DispatchDesignated
        {
            get => mDispatchDesignated;
            set => SetProperty(ref mDispatchDesignated, value);
        }
        public bool StopNextStep
        {
            get => mStopNextStep;
            set => SetProperty(ref mStopNextStep, value);
        }
    }

    [Serializable]
    public class Filter : EditableObject
    {
        private BoolType mBoolType;
        private LogicColumnName mColumn;
        public BoolType BoolType
        {
            get => mBoolType;
            set => SetProperty(ref mBoolType,value);
        }
        public LogicColumnName Column
        {
            get => mColumn;
            set
            {
                SetProperty(ref mColumn, value); 
                OnPropertyChanged(nameof(DataType));
            }

        }
        public OperatorType OperatorType { get; set; }
        public string Value { get; set; }
        public ColumnDataType DataType
        {
            get
            {
                switch(Column)
                { 
                    case LogicColumnName.PayNo:
                    case LogicColumnName.CurrentPrice:
                    {
                        return ColumnDataType.Number;
                    }
                    case LogicColumnName.PayAddress:
                    case LogicColumnName.ProductName:
                    {
                        return ColumnDataType.Address;
                    }
                    case LogicColumnName.DifferentArea:
                    case LogicColumnName.DifferentService:
                    case LogicColumnName.SellerIsService:
                    {
                        return ColumnDataType.Other;
                    }
                    case LogicColumnName.ServiceStatus:
                    case LogicColumnName.ServiceID:
                    case LogicColumnName.PreviousServiceID:
                    case LogicColumnName.ProductID:
                    default:
                    {
                        return ColumnDataType.String;
                    }
                }
            }
        }
    }
    [XmlInclude(typeof(Column_ServiceID))]
    [XmlInclude(typeof(Column_Current_Price))]
    [XmlInclude(typeof(Column_Service_Status))]
    public abstract class ColumnData
    {
        public abstract ColumnDataType ColumnDataType { get;  } 
        public abstract ColumnType ColumnType { get; }
        public abstract string Title { get; }
    }

    [Serializable] 
    public class Column_ServiceID : ColumnData
    {
        private ColumnDataType mColumnDataType = ColumnDataType.String;
        private string mTitle = "客服专员工号";
        private ColumnType mColumnType = ColumnType.RealColumn;
        public override ColumnDataType ColumnDataType => mColumnDataType;
        public override ColumnType ColumnType => mColumnType;
        public override string Title => mTitle;
    }
    [Serializable]
   
    public class Column_Current_Price : ColumnData
    {
        public override ColumnDataType ColumnDataType => ColumnDataType.Number;
        public override ColumnType ColumnType => ColumnType.RealColumn;
        public override string Title => "本期应缴保费";
    }
    [Serializable] 
    public class Column_Service_Status : ColumnData
    {
        public override ColumnDataType ColumnDataType => ColumnDataType.String;
        public override ColumnType ColumnType => ColumnType.LogicColumn;
        public override string Title => "客服专员在职状态";
    }


    [Serializable]
    public class Dispatcher : EditableObject
    {

    }
}
