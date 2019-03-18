using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using InsuranceCompareTool.Models;
namespace InsuranceCompareTool.Services {
    public class BillFormatService
    {
        private static BillFormatService Singleton = null;
        public static BillFormatService CreateInstance()
        {
            if (Singleton == null)
            {
                Singleton = new BillFormatService();
            }
            return Singleton;
        }

        private Regex mRegex  = new Regex("[a-zA-Z0-9]+");

        public void Format(List<Bill> bills, List<Member> members)
        {
            
        }
    }
}