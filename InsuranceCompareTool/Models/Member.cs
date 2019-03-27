using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;
namespace InsuranceCompareTool.Models
{

    public class Member
    {
        public string ID { get; set; }
        public string Name { get; set; } 
        public string Position { get; set; }
        public string Status { get; set; }
        public string Area { get; set; }
        public string Company { get; set; }
        public string CompanyID { get; set; }
        public string HostID { get; set; }
        
        public Member Slave { get; }
        public int BillCount { get; set; } = 0;
        public bool VirtualMember { get; set; } = false;
        public bool Reportable { get; set; } = false;
    }

    public class Bill
    {
        public string ID { get; set; }
        public string CustomerPassportID { get; set; }
        public string CustomerName { get; set; }
        public string CustomerAddress { get; set; }
        public string CustomerArea { get; set; } = "";

        public string SellArea { get; set; } = "";

        public string PayAddress { get; set; } 
        public string LastServiceID { get; set; }
        public string LastServiceName { get; set; }
        public string CurrentServiceID { get; set; }
        public string CurrentServiceName { get; set; }
        public string SellerID { get; set; }
        public string SellerName { get; set; }
        public int PayNo { get; set; }
        public DateTime PayDate { get; set; }
        public string CompanyID { get; set; }  

        public string SrcCurrentServiceName { get; set; }
        public Member CurrentServiceObj { get; set; }
        public Member LastServiceObj { get; set; }
        public Member SellerObj { get; set; }
        public List<BillStatus> Statuses { get;   } = new List<BillStatus>();

    
        public int RowNum { get; set; }

        public string PayModel { get; set; } = "";
        public string SellerState { get; set; }
        public string Organization4 { get; set; }
        public double Price { get; set; }
        public string ProductName { get; set; }
        public string Creditor { get; set; }
        public string PerPay { get; set; }
        
        public string CustomerAccount { get; set; } 
        public string MobilePhone { get; set; }
    }
}
