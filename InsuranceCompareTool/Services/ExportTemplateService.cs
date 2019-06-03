using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using InsuranceCompareTool.Models;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
namespace InsuranceCompareTool.Services
{
    public class ExportTemplateService
    {
        private static ExportTemplateService Singleton = null;
        public static ExportTemplateService CreateInstance( )
        {
            if (Singleton == null)
            {
                Singleton = new ExportTemplateService( );
            }
            return Singleton;
        } 
        private   SheetTemplate mServiceSheetTemplate;
        private object mLockObject = new object();
        private SheetTemplate mSellerSheetTemplate;
        private SheetTemplate mAreaSumSheetTemplate;
        private const string AREA_SUM_TEMP_SHEET_NAME = "地区统计";
        private const string SERVICE_TEMP_SHEET_NAME = "客服清单";
        private const string SELLER_TEMP_SHEET_NAME = "业务员清单";
        private const string ALL_TEMP_SHEET_NAME = "总清单";
        private bool mIsLoaded = false;
        private SheetTemplate mAllSheetTemplate;
 
        
        public void Load(string templateFile)
        {
            lock(mLockObject)
            {
                IWorkbook excel = new XSSFWorkbook(templateFile);
                try
                {
                    if(excel.GetSheet(AREA_SUM_TEMP_SHEET_NAME) == null)
                    {
                        throw new Exception($"未找到模板表： {AREA_SUM_TEMP_SHEET_NAME}");
                    }

                    if(excel.GetSheet(SERVICE_TEMP_SHEET_NAME) == null)
                    {
                        throw new Exception($"未找到模板表： {SERVICE_TEMP_SHEET_NAME}");
                    }

                    if(excel.GetSheet(SELLER_TEMP_SHEET_NAME) == null)
                    {
                        throw new Exception($"未找到模板表： {SELLER_TEMP_SHEET_NAME}");
                    }

                    if(excel.GetSheet(ALL_TEMP_SHEET_NAME) == null)
                    {
                        throw new Exception($"未找到模板表： {ALL_TEMP_SHEET_NAME}");
                    }

                    mServiceSheetTemplate = new SheetTemplate()
                    {
                        Sheet = excel.GetSheet(SERVICE_TEMP_SHEET_NAME)
                    };
                    mSellerSheetTemplate = new SheetTemplate()
                    {
                        Sheet = excel.GetSheet(SELLER_TEMP_SHEET_NAME)
                    };
                    mAreaSumSheetTemplate = new SheetTemplate()
                    {
                        Sheet = excel.GetSheet(AREA_SUM_TEMP_SHEET_NAME)
                    };
                    mAllSheetTemplate = new SheetTemplate()
                    {
                        Sheet = excel.GetSheet(ALL_TEMP_SHEET_NAME)
                    };
                    mIsLoaded = true;
                }
                catch(Exception ex)
                {
                    throw;
                }
                finally
                {
                    excel.Close();
                }
            }
           
        }
        public SheetTemplate ServiceSheetTemplate
        {
            get
            {
                 
                lock(mLockObject)
                {
                    return mServiceSheetTemplate;
                }
            }
        }
        public SheetTemplate SellerSheetTemplate
        {
            get
            {
                
                lock (mLockObject)
                {
                    return  mSellerSheetTemplate;
                } 
            } 
        }
        public SheetTemplate AreaSumSheetTemplate
        {
            get
            {
                
                lock (mLockObject)
                {
                    return mAreaSumSheetTemplate;
                }
            }
        }
        public SheetTemplate AllSheetTemplate
        {
            get
            {
                
                lock(mLockObject)
                {
                    return mAllSheetTemplate;
                }
            }
        }
    }
}
