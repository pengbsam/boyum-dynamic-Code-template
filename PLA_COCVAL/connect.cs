using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Configuration;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using SAPbobsCOM;

namespace PLA_COCVAL
{
    class connect
    {
        public void SetApplication()
        { 
            string connectionstring = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056";
            SAPbouiCOM.SboGuiApi oGuiApi = new SAPbouiCOM.SboGuiApi();

            oGuiApi.Connect(connectionstring);

            Global.oapplication = oGuiApi.GetApplication(1);
        }

        public connect()
        {
            try
            { 
                SetApplication();
           
                Global.ocompany = new SAPbobsCOM.Company();
                Global.ocompany = (SAPbobsCOM.Company)Global.oapplication.Company.GetDICompany();
                Global.oapplication.SetStatusBarMessage("TEST Program connected:" + Global.ocompany.CompanyName.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, false);


                Global.oapplication.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
            }
            catch(Exception ex)
            {
                Global.oapplication.SetStatusBarMessage("TEST Program Connection failed:" + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
           

         }

        public void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            
            switch FormUID:
                case 
        }

    }
}
