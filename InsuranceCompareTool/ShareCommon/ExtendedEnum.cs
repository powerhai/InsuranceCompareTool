using System;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
namespace InsuranceCompareTool.ShareCommon
{
    public static class ExtendedEnum
    { 
        public static string GetDescription(this Enum enumItem)
        {
            if(enumItem == null)
            {
                return "";
            }
            Type enumType = enumItem.GetType();
            string sName = Enum.GetName(enumType, enumItem);
            if (sName == null)
            {
                return null;
            }
            FieldInfo fieldinfo = enumType.GetField(sName);
            Object[] attrs = fieldinfo.GetCustomAttributes(typeof(DescriptionAttribute), false);
            if (attrs == null || attrs.Length == 0)
            {
                return sName;
            }
            else
            {
                DescriptionAttribute descAttr = (DescriptionAttribute)attrs[0];
                return descAttr.Description;
            }
 
        }
      
    }
}
