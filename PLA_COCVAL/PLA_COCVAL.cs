using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PLA_COCVAL
{
    static class PLA_COCVAL
    {
        static public void PLA_COCVAL_DO(ref SAPbouiCOM.Application application, ref SAPbobsCOM.Company company)
        {

            //SAPbouiCOM.Application application = new SAPbouiCOM.Application();
            //application = Global.oapplication;

            //SAPbobsCOM.Company company = new SAPbobsCOM.Company();
            //company = Global.ocompany;

            SAPbouiCOM.Form form = application.Forms.ActiveForm;

            

            SAPbouiCOM.Form gridForm = null;
            SAPbouiCOM.Item oItem = null;
            SAPbouiCOM.Grid oGrid = null;
            SAPbouiCOM.Button oBtn = null;
            string COC1;
            string COC2 = "BOS_VW_COCP2_T1";
            string COC3 = "BOS_VW_COCP3";
            string COCP = "BOS_VW_COC_PIVOT";
            int errorSection = 0;
            string errorMsg = "";

            try
            {
                SAPbobsCOM.Recordset coc1Record = (SAPbobsCOM.Recordset)(company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset));
                SAPbobsCOM.Recordset oRecord2 = (SAPbobsCOM.Recordset)(company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset));

                form.DataSources.DataTables.Add()
                string docEntry = form.DataSources.DBDataSources.Item("ODLN").GetValue("DocEntry", 0);
                if (docEntry == null)
                { 
                    errorMsg += "No Document Number found" + "\r\n";
                    throw new Exception("COCERROR");
                }
                string draftkey = form.DataSources.DBDataSources.Item("ODLN").GetValue("draftKey", 0);
                bool draftDoc;
                if (draftkey == "-1")
                {
                    draftDoc = true;
                    COC1 = "BOS_VW_COCP1_DFT";
                }
                else
                {
                    draftDoc = false;
                    COC1 = "BOS_VW_COCP1";
                }

                coc1Record.DoQuery("select a.*, b.*, c.U_VAC_SPEC_VIS, U_VAC_SPEC1, U_VAC_SPEC2, U_VAC_SPEC3,U_VAC_SPEC4,U_VAC_SPEC5, U_VAC_SPEC6 " +
                    "from " + COC1 + " a left outer join OBTN b on a.\"BlockItem\" = b.\"ItemCode\" and a.\"BatchNum\" = b.\"DistNumber\" "+
                    " left outer join RDR1 C on a.\"RdrEntry\" = c.\"DocEntry\" and a.\"RdrLine\" = c.\"LineNum\" " +
                    " where a.\"DLNEntry\" =" + docEntry.ToString());
                if (coc1Record.RecordCount == 0)
                {
                    errorMsg += "No COC Header Document found for :" + docEntry.ToString() + "\r\n";
                    throw new Exception("COCERROR");
                }
                for (int i = 0; i < coc1Record.RecordCount; i ++)
                {
                    if (coc1Record.Fields.Item("RESULT").ToString() == "")
                    {
                        errorSection++;
                        errorMsg += "No COC Header Document found for :" + docEntry.ToString() + "\r\n";
                    }
                }
                int x = 0;
                x.ToString()
            }
            catch (Exception e)
            {
                errorSection++;
                errorMsg += "Exception caught: " + e.Message.ToString()+ "\r\n";
            }
            finally
            {
                if (errorMsg != "")
                {
                    application.MessageBox(errorMsg, 1, "OK", "", "");
                    return;
                }
            }
           

            //application.MessageBox("Found DocEntry : " + docEntry.ToString() + " and this is the draft key " + draftkey.ToString(), 1, "ok", "", "");
            //try
            //{
            //    gridForm = application.Forms.GetForm("PopForm", 0);
            //    gridForm.Close();
            //}
            //catch { }

            //gridForm = application.Forms.Add("PopForm", SAPbouiCOM.BoFormTypes.ft_Fixed);
            //gridForm.Title = "Found error";
            //gridForm.Left = 100;
            //gridForm.Width = 450;
            //gridForm.Top = 100;
            //gridForm.Height = 280;
            //gridForm.Visible = true;

            //gridForm.DataSources.DataTables.Add("myDataTable");
            //gridForm.DataSources.DataTables.Item(0).ExecuteQuery("select " + docEntry + " as DOCENTRY, " + draftkey + " as DRAFTKEY from Dummy;");

            //oItem = gridForm.Items.Add("1", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            //oItem.Left = 15;
            //oItem.Top = 220;
            //oItem.Width = 60;
            //oBtn = (SAPbouiCOM.Button)oItem.Specific;
            //oBtn.Caption = "OK";

            //oItem = gridForm.Items.Add("2", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            //oItem.Left = 100;
            //oItem.Top = 220;
            //oItem.Width = 60;
            //oBtn = (SAPbouiCOM.Button)oItem.Specific;
            //oBtn.Caption = "Cancel";

            //oItem = gridForm.Items.Add("myGrid", SAPbouiCOM.BoFormItemTypes.it_GRID);
            //oItem.Left = 15;
            //oItem.Top = 15;
            //oItem.Width = 420;
            //oItem.Height = 200;
            //oItem.Enabled = false;
            //oGrid = (SAPbouiCOM.Grid)oItem.Specific;
            //oGrid.DataTable = gridForm.DataSources.DataTables.Item("myDataTable");
            //oGrid.AutoResizeColumns();
        }
    }
}
