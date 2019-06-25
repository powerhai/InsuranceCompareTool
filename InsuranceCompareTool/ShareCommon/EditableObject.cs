using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.Serialization;
using Prism.Mvvm;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.Serialization.Formatters.Binary;
using System.Xml.Serialization;
namespace InsuranceCompareTool.ShareCommon
{ 

    [Serializable]
    public abstract class EditableObject : INotifyPropertyChanged, IEditableObject
    {
        [NonSerialized]
        private object[] mCopy;
        [NonSerialized]
        private bool mIsChanged;
        [NonSerialized]
        private object mBackup;
         

        public virtual void BeginEdit()
        {
            //var members = FormatterServices.GetSerializableMembers(GetType());
            //mCopy = FormatterServices.GetObjectData(this, members);
            mBackup = Copy(this);

        }
        public virtual void EndEdit()
        {
            mCopy = null;
        }
        public virtual void CancelEdit()
        {

            var members = FormatterServices.GetSerializableMembers(GetType());
            var objs = FormatterServices.GetObjectData(mBackup, members);
            var obj = FormatterServices.PopulateObjectMembers(this, members, objs);
            
            mCopy = null; 
        }
        
        public bool IsChanged
        {
            get => mIsChanged;
            set => mIsChanged = value;
        }
        private void NoticeChanged()
        {
            IsChanged = true;
        }



        public object Copy(object obj)
        {

            Object targetDeepCopyObj;
            try
            {



                Type targetType = obj.GetType();
                //值类型  
                if (targetType.IsValueType == true)
                {
                    targetDeepCopyObj = obj;
                }
                //引用类型   
                else
                {
                    targetDeepCopyObj = System.Activator.CreateInstance(targetType);   //创建引用对象   
                    System.Reflection.MemberInfo[] memberCollection = obj.GetType().GetMembers();

                    foreach (System.Reflection.MemberInfo member in memberCollection)
                    {
                        if (member.MemberType == System.Reflection.MemberTypes.Field)
                        {
                            System.Reflection.FieldInfo field = (System.Reflection.FieldInfo)member;
                            Object fieldValue = field.GetValue(obj);
                            if (fieldValue is ICloneable)
                            {
                                field.SetValue(targetDeepCopyObj, (fieldValue as ICloneable).Clone());
                            }
                            else
                            {
                                field.SetValue(targetDeepCopyObj, Copy(fieldValue));
                            }

                        }
                        else if (member.MemberType == System.Reflection.MemberTypes.Property)
                        {
                            System.Reflection.PropertyInfo myProperty = (System.Reflection.PropertyInfo)member;
                            MethodInfo info = myProperty.GetSetMethod(false);
                            if (info != null)
                            {
                                object propertyValue = myProperty.GetValue(obj, null);
                                if (propertyValue is ICloneable)
                                {
                                    myProperty.SetValue(targetDeepCopyObj, (propertyValue as ICloneable).Clone(), null);
                                }
                                else
                                {
                                    myProperty.SetValue(targetDeepCopyObj, Copy(propertyValue), null);
                                }
                            }

                        }
                    }
                }
                return targetDeepCopyObj;
            }
            catch (Exception e)
            {

            }

            return null;
        }


        //public EditableObject DeepCopyByXml(EditableObject obj)
        //{
        //    object retval;
        //    using (MemoryStream ms = new MemoryStream())
        //    {
        //        XmlSerializer xml = new XmlSerializer(typeof(EditableObject));
        //        xml.Serialize(ms, obj);
        //        ms.Seek(0, SeekOrigin.Begin);
        //        retval = xml.Deserialize(ms);
        //        ms.Close();
        //    }
        //    return (EditableObject)retval;
        //}

        public EditableObject DeepCopyByBin(EditableObject obj)
        {
            object retval;
            using (MemoryStream ms = new MemoryStream())
            {
                BinaryFormatter bf = new BinaryFormatter();
                //序列化成流
                bf.Serialize(ms, obj);
                ms.Seek(0, SeekOrigin.Begin);
                //反序列化成对象
                retval = bf.Deserialize(ms);
                ms.Close();
            }
            return (EditableObject)retval;
        }

        [field: NonSerialized] 
        /// <summary>
        /// Occurs when a property value changes.
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged;
  
        /// <summary>
        /// Checks if a property already matches a desired value. Sets the property and
        /// notifies listeners only when necessary.
        /// </summary>
        /// <typeparam name="T">Type of the property.</typeparam>
        /// <param name="storage">Reference to a property with both getter and setter.</param>
        /// <param name="value">Desired value for the property.</param>
        /// <param name="propertyName">Name of the property used to notify listeners. This
        /// value is optional and can be provided automatically when invoked from compilers that
        /// support CallerMemberName.</param>
        /// <returns>True if the value was changed, false if the existing value matched the
        /// desired value.</returns>
        protected virtual bool SetProperty<T>(ref T storage, T value, [CallerMemberName] string propertyName = null)
        {
            if (EqualityComparer<T>.Default.Equals(storage, value)) return false;

            storage = value;
            RaisePropertyChanged(propertyName);

            return true;
        }
 
        /// <summary>
        /// Checks if a property already matches a desired value. Sets the property and
        /// notifies listeners only when necessary.
        /// </summary>
        /// <typeparam name="T">Type of the property.</typeparam>
        /// <param name="storage">Reference to a property with both getter and setter.</param>
        /// <param name="value">Desired value for the property.</param>
        /// <param name="propertyName">Name of the property used to notify listeners. This
        /// value is optional and can be provided automatically when invoked from compilers that
        /// support CallerMemberName.</param>
        /// <param name="onChanged">Action that is called after the property value has been changed.</param>
        /// <returns>True if the value was changed, false if the existing value matched the
        /// desired value.</returns>
        protected virtual bool SetProperty<T>(ref T storage, T value, Action onChanged, [CallerMemberName] string propertyName = null)
        {
            if (EqualityComparer<T>.Default.Equals(storage, value)) return false;

            storage = value;
            onChanged?.Invoke();
            RaisePropertyChanged(propertyName);

            return true;
        }
 
        /// <summary>
        /// Raises this object's PropertyChanged event.
        /// </summary>
        /// <param name="propertyName">Name of the property used to notify listeners. This
        /// value is optional and can be provided automatically when invoked from compilers
        /// that support <see cref="CallerMemberNameAttribute"/>.</param>
        protected void RaisePropertyChanged([CallerMemberName]string propertyName = null)
        {
            //TODO: when we remove the old OnPropertyChanged method we need to uncomment the below line
            //OnPropertyChanged(new PropertyChangedEventArgs(propertyName));
#pragma warning disable CS0618 // Type or member is obsolete
            OnPropertyChanged(propertyName);
#pragma warning restore CS0618 // Type or member is obsolete
        }
 
        /// <summary>
        /// Notifies listeners that a property value has changed.
        /// </summary>
        /// <param name="propertyName">Name of the property used to notify listeners. This
        /// value is optional and can be provided automatically when invoked from compilers
        /// that support <see cref="CallerMemberNameAttribute"/>.</param>
        [Obsolete("Please use the new RaisePropertyChanged method. This method will be removed to comply wth .NET coding standards. If you are overriding this method, you should overide the OnPropertyChanged(PropertyChangedEventArgs args) signature instead.")]
        [System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Never)]
        protected virtual void OnPropertyChanged([CallerMemberName]string propertyName = null)
        {
            OnPropertyChanged(new PropertyChangedEventArgs(propertyName));
        }
 
        /// <summary>
        /// Raises this object's PropertyChanged event.
        /// </summary>
        /// <param name="args">The PropertyChangedEventArgs</param>
        protected virtual void OnPropertyChanged(PropertyChangedEventArgs args)
        {
            PropertyChanged?.Invoke(this, args);
        }
 
        /// <summary>
        /// Raises this object's PropertyChanged event.
        /// </summary>
        /// <typeparam name="T">The type of the property that has a new value</typeparam>
        /// <param name="propertyExpression">A Lambda expression representing the property that has a new value.</param>
        [Obsolete("Please use RaisePropertyChanged(nameof(PropertyName)) instead. Expressions are slower, and the new nameof feature eliminates the magic strings.")]
        [System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Never)]
        protected virtual void OnPropertyChanged<T>(Expression<Func<T>> propertyExpression)
        {
            var propertyName = PropertySupport.ExtractPropertyName(propertyExpression);
            OnPropertyChanged(propertyName);
        }



    }
}
