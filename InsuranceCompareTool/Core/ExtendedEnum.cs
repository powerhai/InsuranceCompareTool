using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace InsuranceCompareTool.core
{
    public static class ExtendedEnum
    {
  
        public static string GetDescription(this Enum enumObj)
        {
            if (enumObj == null)
                return null;
            string rv = "";
             
            if (string.IsNullOrEmpty(rv))
            {
                FieldInfo fieldInfo = enumObj.GetType().GetField(enumObj.ToString());
                var attribArray = Attribute.GetCustomAttributes(fieldInfo, typeof(DescriptionAttribute), false);
                rv = enumObj.ToString();
                if (attribArray.Any())
                {
                    var att = attribArray[0] as DescriptionAttribute;
                    rv = att != null ? att.Description : enumObj.ToString();

                }
            }
            return rv;
        }
      
    }
}
