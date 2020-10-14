using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PLA_COCVAL
{
    class BankfileArchive
    {
        static public void BankfileArchive_DO(ref SAPbouiCOM.Application application, ref SAPbobsCOM.Company company)
        {
            SAPbouiCOM.Form form = application.Forms.ActiveForm;

            string vcardFolder = @"\\core.com\core\Shared\Accounting\Accounts Payable\Virtual Card\TEST\HUNT-VIRTUAL_VCARD\HUNT_Vcard";
            string achfolder = @"\\core.com\core\Shared\Accounting\Accounts Payable\Virtual Card\TEST\HUNT-VIRTUAL_VCARD\HUNT_ACH";
            string toFolder = @"\\core.com\core\Shared\Accounting\Accounts Payable\Virtual Card\TEST\HUNT-VIRTUAL_VCARD\Archive";
            string currentDTM = DateTime.Now.ToString("yyyyMMdd_HHmmss");

            string achfromfile = "";
            string vcardfromfile = "";
            string vcarddestFile = "";
            string achdestFile = "";

            int vcardcount = 0;
            int achcount = 0;
            int choice = 0;

            string showmsg = "";


            //Get varcfile

            if (System.IO.Directory.Exists(vcardFolder))
            {
                string[] files = System.IO.Directory.GetFiles(vcardFolder);

                foreach (string s in files)
                {
                    if (System.IO.Path.GetExtension(s).ToUpper() == ".CSV")
                    {
                        showmsg = showmsg + "Vcard file : " + System.IO.Path.GetFileNameWithoutExtension(s) + " Last Modified " + System.IO.File.GetLastWriteTime(s).ToString() + "\r\n";
                        vcardcount++;
                    }
                }
            }

            //get achfile

            if (System.IO.Directory.Exists(achfolder))
            {
                string[] files = System.IO.Directory.GetFiles(achfolder);

                foreach (string s in files)
                {
                    if (System.IO.Path.GetExtension(s).ToUpper() == ".CSV")
                    {
                        showmsg = showmsg + "ACH file : " + System.IO.Path.GetFileNameWithoutExtension(s) + " Last Modified " + System.IO.File.GetLastWriteTime(s).ToString() + "\r\n";
                        achcount++;
                    }
                }
            }
            //show user and choose
            if (showmsg != "")
            {
                showmsg = showmsg + " Please choose which file you want to archive";
                choice = application.MessageBox(showmsg, 3, "vcard", "ACH", "Both");
            }



            try
            {
                if (choice == 1 || choice == 3)
                {
                    application.MessageBox("run vcard", 1, "OK", "", "");
                    //vcard file conversion.
                    if (System.IO.Directory.Exists(vcardFolder))
                    {
                        string[] files = System.IO.Directory.GetFiles(vcardFolder);

                        foreach (string s in files)
                        {
                            if (System.IO.Path.GetExtension(s).ToUpper() == ".CSV")
                            {
                                vcardfromfile = "VCARD_" + System.IO.Path.GetFileNameWithoutExtension(s) + "_" + currentDTM + System.IO.Path.GetExtension(s);
                                vcarddestFile = System.IO.Path.Combine(toFolder, vcardfromfile);
                                System.IO.File.Copy(s, vcarddestFile, true);
                            }
                        }
                    }
                }
                if (choice == 2 || choice == 3)
                {
                    application.MessageBox("run ACH", 1, "OK", "", "");
                    //ach file conversion.
                    if (System.IO.Directory.Exists(achfolder))
                    {
                        string[] files = System.IO.Directory.GetFiles(achfolder);

                        foreach (string s in files)
                        {
                            if (System.IO.Path.GetExtension(s).ToUpper() == ".CSV")
                            {
                                achfromfile = "ACH_" + System.IO.Path.GetFileNameWithoutExtension(s) + "_" + currentDTM + System.IO.Path.GetExtension(s);
                                achdestFile = System.IO.Path.Combine(toFolder, achfromfile);
                                System.IO.File.Copy(s, achdestFile, true);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                application.SetStatusBarMessage("File Archived Failed" + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }
    }
}
