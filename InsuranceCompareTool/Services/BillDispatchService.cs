using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using InsuranceCompareTool.Models;
namespace InsuranceCompareTool.Services
{
    public class BillDispatchService
    {

        private static BillDispatchService Singleton = null;
        public static BillDispatchService CreateInstance()
        {
            if (Singleton == null)
            {
                Singleton = new BillDispatchService();
            }
            return Singleton;
        }

        private static Random mRandom = new Random();

        public void DispatchBills(List<Bill> bills, List<Member> members)
        {
            Member wifuMuster =
                members.FirstOrDefault(a => a.Area.Equals(AreaNames.YW) && a.Position.Equals(PositionNames.MUSTER));
            Member qtMuster = members.FirstOrDefault(a=>a.Area.Equals(AreaNames.QT) && a.Position.Equals(PositionNames.MUSTER));

            var masters = members.Where(a => a.Position.Equals(PositionNames.MUSTER)).ToArray();

            var sfzs = bills.Select(a => a.CustomerPassportID).Distinct().ToArray();

            int no = 1;
            foreach (var sfz in sfzs)
            {
                var sbills = bills.Where(a => a.CustomerPassportID.Equals(sfz)).ToArray();
                if (sbills.Length > 1)
                {
                    var services = sbills.Select(a => a.CurrentServiceID).Distinct().ToArray();
                    if (services.Length > 1)
                    {
                        var first = sbills.First();
                        var serviceID = first.CurrentServiceID;
                        var serviceName = first.CurrentServiceName;
                        foreach (var bill in sbills)
                        {
                            bill.SrcServiceName = bill.CurrentServiceName;
                            bill.CurrentServiceID = serviceID;
                            bill.CurrentServiceName = serviceName;
                        }
                    }
                }
            }
            /*
           foreach(var bill in bills)
           {
               bill.SrcCurrentServiceName = bill.CurrentServiceName;
               if(bill.Statuses.Contains(BillStatus.AreaSell)) //区拓
               {
                   if(qtMuster != null)
                   {
                       bill.CurrentServiceID = qtMuster?.ID;
                       bill.CurrentServiceName = qtMuster?.Name;
                       bill.CurrentServiceObj = qtMuster;
                   }
                   else
                   {
                       bill.Statuses.Add(BillStatus.Error);
                   }
                   continue;
               }
               if(bill.Statuses.Contains(BillStatus.Yiwu)) //义乌
               {
                   if(wifuMuster != null)
                   {
                       bill.CurrentServiceID = wifuMuster?.ID;
                       bill.CurrentServiceName = wifuMuster?.Name;
                       bill.CurrentServiceObj = wifuMuster;
                   }
                   else
                   {
                       bill.Statuses.Add(BillStatus.Error);
                   }

                   continue;
               }

               if(bill.Statuses.Contains(BillStatus.AcrossArea234)) //跨区234
               {
                   var mem = masters.FirstOrDefault(a => a.Area.Equals(bill.SellArea));
                   if(mem != null)
                   {
                       bill.CurrentServiceID = mem.ID;
                       bill.CurrentServiceName = mem.Name;
                       bill.CurrentServiceObj = mem;
                   }
                   else
                   {
                       bill.Statuses.Add(BillStatus.Error);
                   }
                   continue; 
               }

               if (bill.Statuses.Contains(BillStatus.AcrossArea5)) //跨区5
               {
                   var mem = masters.FirstOrDefault(a => a.Area.Equals(bill.CustomerArea));
                   if (mem != null)
                   {
                       bill.CurrentServiceID = mem.ID;
                       bill.CurrentServiceName = mem.Name;
                       bill.CurrentServiceObj = mem;
                   }
                   else
                   {
                       bill.Statuses.Add(BillStatus.Error);
                   }
                   continue;
               }

               if(bill.Statuses.Contains(BillStatus.BindService)) //绑定客服
               {
                   var mem = members.FirstOrDefault(a => a.ID.Equals(bill.SellerObj.Slave.ID));
                   if (mem != null)
                   {
                       bill.CurrentServiceID = mem.ID;
                       bill.CurrentServiceName = mem.Name;
                       bill.CurrentServiceObj = mem;
                   }
                   else
                   {
                       bill.Statuses.Add(BillStatus.Error);
                   }
               }

               if(bill.Statuses.Contains(BillStatus.SellerNotAvailable)) //营销员非在职
               {
                   var mt = masters.FirstOrDefault(a => a.Area.Equals(bill.SellArea));
                   var mems = masters.Where(a => a.Area.Equals(bill.SellArea) && a.Status.Equals(StatusNames.ZAI_ZI)).ToArray();
                   Member mem = null;
                   if(mt != null && mt.BillCount < 30)
                   {
                       mem = mt;
                   }
                   else
                   { 
                       if( mems.Length > 0)
                       {
                           mem = mems[mRandom.Next(0, mems.Length)]; 
                       }
                   }
                   if (mem != null)
                   {
                       bill.CurrentServiceID = mem.ID;
                       bill.CurrentServiceName = mem.Name;
                       bill.CurrentServiceObj = mem;
                   }
                   else
                   {
                       bill.Statuses.Add(BillStatus.Error);
                   }
                   continue;
               }

           }*/
        }
    }
}
