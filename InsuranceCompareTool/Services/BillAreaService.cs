using System.Collections.Generic;
using System.Linq;
using InsuranceCompareTool.Models;
namespace InsuranceCompareTool.Services {
    public class BillAreaService
    {
        private static BillAreaService Singleton = null;
        public static BillAreaService CreateInstance()
        {
            if (Singleton == null)
            {
                Singleton = new BillAreaService();
            }
            return Singleton;
        }
        


        private   string[] mJhAddress = new string[]{"永康街","义乌街","浦江街","兰溪街","东阳街","武义街","磐安街","婺城区", "金东区"};
         
        public void CalculateArea(List<Bill> bills  )
        {
            foreach (var bill in bills)
            {
                bill.CustomerArea = GetAddressArea(bill.CustomerAddress);
                bill.SellArea = GetAreaFormOrganization4(bill.Organization4);
            }
        }

        private string GetAreaFormOrganization4(string address)
        {
            if(string.IsNullOrEmpty(address))
                return "";
            if (address.Contains(AreaNames.YK))
            {
                return AreaNames.YK;
            }

            if (address.Contains(AreaNames.WY))
            {
                return AreaNames.WY;
            }
            if (address.Contains(AreaNames.YW))
            {
                return AreaNames.YW;
            }
            if (address.Contains(AreaNames.LX))
            {
                return AreaNames.LX;
            }
            if (address.Contains(AreaNames.PJ))
            {
                return AreaNames.PJ;
            }
            if (address.Contains(AreaNames.PA))
            {
                return AreaNames.PA;
            }
            if (address.Contains(AreaNames.DY))
            {
                return AreaNames.DY;
            }
            return AreaNames.JH;
        }
        private string GetAddressArea(string address)
        {
            if(string.IsNullOrEmpty(address))
            {
                return "";
            } 

            foreach(var jh in mJhAddress)
            {
                if(address.Contains(jh))
                {
                    return AreaNames.JH;
                }
            }
            if(address.Contains(AreaNames.YK))
            {
                return AreaNames.YK;
            }

            if(address.Contains(AreaNames.WY))
            {
                return AreaNames.WY;
            }
            if (address.Contains(AreaNames.YW))
            {
                return AreaNames.YW;
            }
            if (address.Contains(AreaNames.LX))
            {
                return AreaNames.LX;
            }
            if (address.Contains(AreaNames.PJ))
            {
                return AreaNames.PJ;
            }
            if (address.Contains(AreaNames.PA))
            {
                return AreaNames.PA;
            }
            if (address.Contains(AreaNames.DY))
            {
                return AreaNames.DY;
            } 
            return AreaNames.JH; 
        }
    }
}