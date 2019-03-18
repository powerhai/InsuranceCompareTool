using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using InsuranceCompareTool.Models;
namespace InsuranceCompareTool.Services
{
    public class BillMemberService
    {
        private static BillMemberService Singleton = null;
        public static BillMemberService CreateInstance()
        {
            if (Singleton == null)
            {
                Singleton = new BillMemberService();
            }
            return Singleton;
        }

        public void CalculateMembers(List<Bill> bills, List<Member> members)
        {
            foreach(var bill in bills)
            {
                if(!string.IsNullOrEmpty(bill.SellerID))
                {
                    bill.SellerObj = members.FirstOrDefault(a => a.ID.Equals(bill.SellerID));
                }

                 
                if(!string.IsNullOrEmpty(bill.LastServiceID))
                {
                    bill.LastServiceObj = members.FirstOrDefault(a => a.ID.Equals(bill.LastServiceID));
                }
            }
        }
    }
}
