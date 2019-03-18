using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using InsuranceCompareTool.Models;
namespace InsuranceCompareTool.Services
{
    public class BillStatusService
    {
        private static BillStatusService Singleton = null;
        public static BillStatusService CreateInstance()
        {
            if (Singleton == null)
            {
                Singleton = new BillStatusService();
            }
            return Singleton;
        }

        public void CalculateStatus( List<Bill> bills, List<Member> members)
        {
            foreach(var bill in bills)
            {
                //区拓
                if(bill.PayModel.Equals(PayModelNames.AREA_SELL))
                {
                    bill.Statuses.Add(BillStatus.AreaSell);
                    continue;
                }

                //义乌
                if( bill.SellArea.Equals(AreaNames.YW))
                {
                    bill.Statuses.Add( BillStatus.Yiwu);    
                    continue;
                } 

                //跨区
                if( ! bill.CustomerArea.Equals(bill.SellArea))
                {
                    if(bill.PayNo >= 2 && bill.PayNo <= 4)
                    {
                        bill.Statuses.Add(BillStatus.AcrossArea234);
                    }

                    if(bill.PayNo >= 5)
                    {
                        bill.Statuses.Add(BillStatus.AcrossArea5);
                    }
                } 

                //客服不在职
                if(bill.CurrentServiceObj != null && !bill.CurrentServiceObj.Status.Equals(StatusNames.ZAI_ZI))
                {
                    bill.Statuses.Add( BillStatus.ServiceNotAvailable);
                }

                //客服空缺
                if(bill.LastServiceObj == null || bill.CurrentServiceObj == null)
                {
                    bill.Statuses.Add(BillStatus.NoService);
                }

                //与上期客服不一致
                if(bill.LastServiceObj != null && bill.CurrentServiceObj != null &&
                   bill.LastServiceObj != bill.CurrentServiceObj)
                {
                    bill.Statuses.Add(BillStatus.NoLastService);
                }

                //营销员是个服务员
                if(bill.SellerObj != null && bill.SellerObj.Position.Equals(PositionNames.SERVICE))
                {
                    bill.Statuses.Add( BillStatus.ServiceSameToSeller);
                }

                //绑定营销员与客服
                if(bill.SellerObj != null && bill.SellerObj.Slave != null)
                {
                    bill.Statuses.Add(BillStatus.BindService);
                }

                //营销员离职
                if(!bill.SellerState.Equals(StatusNames.ZAI_ZI))
                {
                    bill.Statuses.Add(BillStatus.SellerNotAvailable);
                }
 
            }
        }
    }
}
