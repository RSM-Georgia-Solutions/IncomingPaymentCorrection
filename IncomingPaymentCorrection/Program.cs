using System;
using System.Collections.Generic;
using SAPbobsCOM;
using SAPbouiCOM.Framework;

namespace Invoice_Income_Correction
{
    class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        /// 
        /// 
        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                Application oApp = null;
                oApp = args.Length < 1 ? new Application() : new Application(args[0]);
                DiCompany = (Company)Application.SBO_Application.Company.GetDICompany();
                Menu MyMenu = new Menu();
                MyMenu.AddMenuItems();
                oApp.RegisterMenuEventHandler(MyMenu.SBO_Application_MenuEvent);
                Application.SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
                Recordset recSet = (Recordset)DiCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

                string query = "SELECT LinkAct_25, LinkAct_21 FROM OACP where PeriodCat ='" +
                               DateTime.Now.Year + "'";
                recSet.DoQuery(query);
                ExchangeGain = recSet.Fields.Item("LinkAct_25").Value.ToString();
                ExchangeLoss = recSet.Fields.Item("LinkAct_21").Value.ToString();
                oApp.Run();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        public static Company DiCompany;
        public static string ExchangeGain { get; set; }
        public static string ExchangeLoss { get; set; }
        static void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    //Exit Add-On
                    System.Windows.Forms.Application.Exit();
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_FontChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    break;
                default:
                    break;
            }
        }
    }
}
