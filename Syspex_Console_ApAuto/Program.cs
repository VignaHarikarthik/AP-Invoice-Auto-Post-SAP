using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Syspex_Console_ApAuto
{
    class Program
    {
        #region ***** SAP UserName and Password*****
        public static string strSAPServer = "SYSPEXSAP04";
        public static string strSAPSQLUsername = "sa";
        public static string strSAPSQLPwd = "Password1111";
        public static string strSAPDBName;
        public static string strCompanyConn;
        public static string strSAPUsername = "PRO-061";
        public static string strSAPPwd = "96447338";
        #endregion

        #region ***** SAP Company Connection*****
        readonly private static string SGCompany = "SYSPEX_LIVE";
        static SqlConnection SGConnection = new SqlConnection("Server=192.168.1.21;Database=SYSPEX_LIVE;Uid=Sa;Pwd=Password1111;");
        static SqlConnection LocalConnection = new SqlConnection("Server=192.168.1.21;Database=AndriodAppDB;Uid=Sa;Pwd=Password1111;");

        #endregion

        #region ***** Error handling variables *****
        public static string sErrMsg;
        public static int lErrCode;
        #endregion

        static void Main(string[] args)
        {
            // PO in header 
            //1. One PO with one grn 
            //2. One PO with multiple grn total the sum, if got match post that grn for  apinovice 
            //3. One PO with multiple grn check each amount, if any thing got match post that grn for apinvoice 

            // PO in detail 
            //1. Mulitple PO with muliple grn combine together and check the line Total amount + tax  and tally with the extracted amount and post ap invoice 
            //2. Muitple PO with mulitple grn combine together and check the total line before discount + tax and tally the extracted amount and post ap invoice 

            // Uplaod the PDF to SAP
            // upload_pdf_sap();



            // Extracted Data Post to SAP
            post_apinvoices();

            //System.Threading.Thread.Sleep(1000);

            // Rename the file name after posted
            rename_file_name();



        } 

        public static void post_apinvoices()
        {
            DataTable table = GetExtractedData();
            if (table.Rows.Count > 0)
            {
                SAPbobsCOM.Company oCompany = CompanyConnection("1");
                if (oCompany != null)
                {
                    // loop the extracted data 
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        string docnum = post_to_sap(table.Rows[i]["po_no"].ToString(), table.Rows[i]["invoice_number"].ToString(), Convert.ToDouble(table.Rows[i]["amount"].ToString()), table.Rows[i]["line_amount"].ToString(), table.Rows[i]["po_number_detail"].ToString(), oCompany);
                        if (!string.IsNullOrEmpty(docnum))
                        {
                            // get the sellername
                            if (docnum.All(char.IsDigit) == true)
                            {
                                string strSql;
                                var recordset = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                strSql = "select CardName from OPCH where DocNum= '" + docnum + "'";
                                recordset.DoQuery(strSql);
                                string seller_name = Convert.ToString(recordset.Fields.Item("CardName").Value);
                                send_email("analisa@syspex.com,vigna@syspex.com,jiachi.eng@syspex.com,jessly.chai@syspex.com", table.Rows[i]["invoice_number"].ToString(), TruncateLongString(seller_name, 50), docnum, table.Rows[i]["amount"].ToString());
                                //1 for sucess
                                UpdateSAPSTATUS(table.Rows[i]["pdf_file_name"].ToString(), "1", docnum);
                            }
                            if (docnum.Contains("10001467 - There is already a record with duplicated customer/vendor reference number."))
                            {
                                string strSql;
                                var recordset = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                strSql = "select CardName, docnum from OPCH where NumAtCard like '%" + table.Rows[i]["invoice_number"].ToString() + "%'";
                                recordset.DoQuery(strSql);
                                string seller_name = Convert.ToString(recordset.Fields.Item("CardName").Value);
                                send_email("analisa@syspex.com,vigna@syspex.com,jiachi.eng@syspex.com", table.Rows[i]["invoice_number"].ToString(), TruncateLongString(seller_name, 50), docnum, table.Rows[i]["amount"].ToString());
                                //1 for sucess
                                UpdateSAPSTATUS(table.Rows[i]["pdf_file_name"].ToString(), "1", Convert.ToString(recordset.Fields.Item("docnum").Value));
                            }

                        }
                        else
                        {
                            UpdateSAPSTATUS(table.Rows[i]["pdf_file_name"].ToString(), "2", "");
                        }
                    }
                }
            }
        }
        static void rename_file_name()
        {
            Regex rx = new Regex("^[0-9]{14}|[0-9]{10}$|[EB]{2}", RegexOptions.Compiled | RegexOptions.IgnoreCase);
            StringBuilder sb_filename = new StringBuilder();
            foreach (var file_path in
                 Directory
     .EnumerateFiles(@"F:\apinvoice", "*.pdf")
   .Where(x => rx.IsMatch(Path.GetFileNameWithoutExtension(x))))
            {
                DataTable dt = Getfilename(Path.GetFileName(file_path).ToString());
                if (dt.Rows.Count > 0)
                {
                    if (dt.Rows[0]["sap_docnum"].ToString().Trim() != "")
                    {
                        //Once Pushed To SAP Then Rename the File and Path
                        UpdateSAPSTATUS(dt.Rows[0]["sap_docnum"].ToString());
                        System.IO.File.Move(file_path, @"F:\apinvoice\" + dt.Rows[0]["sap_docnum"].ToString().Replace(" ", "") + ".pdf");
                    }
                    else
                    {
                        sb_filename.Append(Path.GetFileName(file_path).ToString() + ",");

                    }

                }
            }

        }


        public static string upload_pdf_sap()
        {
            int errCode = 0;
            try
            {
                Regex rx = new Regex("^[0-9]{5}$", RegexOptions.Compiled | RegexOptions.IgnoreCase);

                foreach (var file_path in
                                Directory
                    .EnumerateFiles(@"F:\apinvoice", "*.pdf")
                  .Where(x => rx.IsMatch(Path.GetFileNameWithoutExtension(x))))
                {
                    SAPbobsCOM.Company oCompany = CompanyConnection("1");
                    SAPbobsCOM.Attachments2 oAtt = (SAPbobsCOM.Attachments2)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oAttachments2);
                    oAtt.Lines.SourcePath = @"F:\apinvoice";
                    oAtt.Lines.FileName = Path.GetFileNameWithoutExtension(file_path);
                    oAtt.Lines.FileExtension = "pdf";
                    int iErr = oAtt.Add();
                    int AttEntry = 0;
                    if (iErr == 0)
                    {
                        AttEntry = int.Parse(oCompany.GetNewObjectKey());
                        var apinvoice = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices);
                        if (apinvoice.GetByKey(Convert.ToInt32(Path.GetFileNameWithoutExtension(file_path))) == true)
                        {
                            apinvoice.AttachmentEntry = AttEntry;
                            errCode = apinvoice.Update();
                        }
                    }
                    if (iErr != 0)
                    {
                        (oCompany).GetLastError(out lErrCode, out sErrMsg);
                    }
                }
            }

            catch (Exception Ex)
            {
                sErrMsg = Ex.ToString();
            }

            return errCode.ToString();




        }
        public static string post_to_sap(string po_number, string invoice_number, double extracted_amount, string line_amount, string po_number_detail, SAPbobsCOM.Company ocompany)
        {
            DataTable dt = new DataTable();
            if (!string.IsNullOrEmpty(po_number))
                dt = GetDataSap(po_number);
            else if (!string.IsNullOrEmpty(po_number_detail))
                dt = GetLineDataSap(po_number_detail, line_amount);

            string post_sucess = string.Empty;
            double actual_amount = 0.00;
            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    // add the tax if the po in the detail 
                    if (!string.IsNullOrEmpty(po_number_detail))

                        actual_amount += Convert.ToDouble(dt.Rows[i]["grn total"].ToString()) + Convert.ToDouble(dt.Rows[i]["tax"].ToString());
                    else
                        actual_amount += Convert.ToDouble(dt.Rows[i]["grn total"].ToString());
                }
                //without rounding 
                actual_amount = Convert.ToDouble(string.Format("{0:0.00}", actual_amount));

                if (actual_amount == extracted_amount)
                {
                    // if the po got multiple grn match add the total and if total got match
                    post_sucess = create_apinvoice_with_multiple_grnentry(dt, invoice_number, ocompany);

                }
                else
                {
                    // if the po got multiple grn match any one of grn match then post invoice
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        if (extracted_amount == Convert.ToDouble(dt.Rows[i]["grn total"].ToString()))
                            post_sucess = create_apinvoice_with_single_grnentry(Convert.ToInt32(dt.Rows[i]["grn docentry"].ToString()), invoice_number, ocompany);
                    

                    }
                }

            }


            return post_sucess;
        }
        public static string create_apinvoice_with_multiple_grnentry(DataTable dt, string invoice_number, SAPbobsCOM.Company oCompany)
        {
            var grpo = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes);
            var apinvoice = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices);
            var recordset = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            // will get muliple same grn entry so making distinct
            DataView view = new DataView(dt);
            DataTable distinctValues = view.ToTable(true, "grn docentry");

            try
            {
                int iTotalPO_Line;
                for (int i = 0; i < distinctValues.Rows.Count; i++)
                {
                    if (grpo.GetByKey(Convert.ToInt32(distinctValues.Rows[i]["grn docentry"])) == true)
                    {
                        apinvoice.CardCode = grpo.CardCode;
                        apinvoice.JournalMemo = TruncateLongString(invoice_number + '#' + grpo.CardName, 50);
                        apinvoice.DocDate = DateTime.Now;
                       // apinvoice.DocDueDate = DateTime.Now;
                        apinvoice.DocRate = grpo.DocRate;
                        apinvoice.NumAtCard = TruncateLongString(invoice_number + '#' + grpo.CardName, 100);
                        apinvoice.Comments = "Created by Ap Invoice Automation " + DateTime.Now.ToShortDateString() + "";
                        iTotalPO_Line = grpo.Lines.Count;

                        //Update GRPO Document
                        grpo.JournalMemo = TruncateLongString(invoice_number + '#' + grpo.CardName, 50);
                        grpo.NumAtCard = TruncateLongString(invoice_number + '#' + grpo.CardName, 100);
                        grpo.Update();


                        int x;
                        for (x = 0; x <= iTotalPO_Line - 1; x++)
                        {
                            grpo.Lines.SetCurrentLine(x);

                            if (grpo.Lines.LineStatus == SAPbobsCOM.BoStatus.bost_Close)
                            {
                            }
                            else
                            {
                                apinvoice.Lines.ItemCode = grpo.Lines.ItemCode;
                                apinvoice.Lines.WarehouseCode = grpo.Lines.WarehouseCode;
                                apinvoice.Lines.Quantity = grpo.Lines.Quantity;
                                apinvoice.Lines.BaseType = 20;
                                apinvoice.Lines.BaseEntry = grpo.DocEntry;
                                apinvoice.Lines.BaseLine = grpo.Lines.LineNum;
                                apinvoice.Lines.Add();
                                
                            }
                        }

                    }

                }
                lErrCode = apinvoice.Add();
                if (lErrCode != 0)
                {
                    (oCompany).GetLastError(out lErrCode, out sErrMsg);
                }
            }

            catch (Exception Ex)
            {
                sErrMsg = Ex.ToString();
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(grpo);
                SAPbobsCOM.Documents documents = grpo = null;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(apinvoice);
                SAPbobsCOM.Documents documents1 = apinvoice = null;
                GC.Collect();

            }
            if (lErrCode == 0)
            {
                string strSql;
                strSql = "select DocNum from OPCH where DocNum= '" + Convert.ToString(((SAPbobsCOM.Company)oCompany).GetNewObjectKey()) + "'";
                recordset.DoQuery(strSql);
                return Convert.ToString(recordset.Fields.Item("DocNum").Value);
            }
            else
            {
                return sErrMsg;
            }

        }
        public static string create_apinvoice_with_single_grnentry(int docentry, string invoice_number, SAPbobsCOM.Company oCompany)
        {

            var grpo = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes);
            var apinvoice = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices);
            var recordset = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                int iTotalPO_Line;
                if (grpo.GetByKey(docentry) == true)
                {
                    apinvoice.JournalMemo = TruncateLongString(invoice_number + '#' + apinvoice.CardName, 50);
                    apinvoice.DocDate = DateTime.Now;
                   // apinvoice.DocDueDate = DateTime.Now;
                    apinvoice.NumAtCard = TruncateLongString(invoice_number + '#' + apinvoice.CardName, 100);
                    apinvoice.Comments = "Created by Ap Invoice Automation on " + DateTime.Now.ToShortDateString() + "";
                    iTotalPO_Line = grpo.Lines.Count;

                    //Update GRPO Document
                    grpo.JournalMemo = TruncateLongString(invoice_number + '#' + grpo.CardName, 50);
                    grpo.NumAtCard = TruncateLongString(invoice_number + '#' + grpo.CardName, 100);
                    grpo.Update();


                    int x;
                    for (x = 0; x <= iTotalPO_Line - 1; x++)
                    {
                        grpo.Lines.SetCurrentLine(x);

                        if (grpo.Lines.LineStatus == SAPbobsCOM.BoStatus.bost_Close)
                        {
                        }
                        else
                        {
                            apinvoice.Lines.ItemCode = grpo.Lines.ItemCode;
                            apinvoice.Lines.WarehouseCode = grpo.Lines.WarehouseCode;
                            apinvoice.Lines.Quantity = grpo.Lines.Quantity;
                            apinvoice.Lines.BaseType = 20;
                            apinvoice.Lines.BaseEntry = grpo.DocEntry;
                            apinvoice.Lines.BaseLine = grpo.Lines.LineNum;
                            apinvoice.Lines.Add();
                        }
                    }

                }
                lErrCode = apinvoice.Add();
                if (lErrCode != 0)
                {
                    (oCompany).GetLastError(out lErrCode, out sErrMsg);
                }
            }

            catch (Exception Ex)
            {
                sErrMsg = Ex.ToString();
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(grpo);
                SAPbobsCOM.Documents documents = grpo = null;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(apinvoice);
                SAPbobsCOM.Documents documents1 = apinvoice = null;
                GC.Collect();

            }
            if (lErrCode == 0)
            {
                string strSql;
                strSql = "select DocNum from OPCH where DocNum= '" + Convert.ToString(((SAPbobsCOM.Company)oCompany).GetNewObjectKey()) + "'";
                recordset.DoQuery(strSql);
                return Convert.ToString(recordset.Fields.Item("DocNum").Value);

            }
            else
            {
                return sErrMsg;
            }

        }
        static SAPbobsCOM.Company CompanyConnection(string RegionID)
        {
            int lErrCode = 0;
            int lRetCode;
            string sErrMsg = "";

            if (RegionID == "1")
                strSAPDBName = SGCompany;


            SAPbobsCOM.Company oCompany = new SAPbobsCOM.Company();
            string strLicenseName = strSAPServer + ":30000";
            oCompany.Server = strSAPServer;
            oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2016;
            oCompany.CompanyDB = strSAPDBName;
            oCompany.UserName = strSAPUsername;
            oCompany.Password = strSAPPwd;
            oCompany.DbUserName = strSAPSQLUsername;
            oCompany.DbPassword = strSAPSQLPwd;
            oCompany.UseTrusted = false;
            oCompany.LicenseServer = strLicenseName;
            //Try to connect
            lRetCode = oCompany.Connect();
            if (lRetCode != 0) // if the connection failed
            {
                int temp_int = lErrCode;
                string temp_string = sErrMsg;
                oCompany.GetLastError(out temp_int, out temp_string);
            }
            return oCompany;
        }
        private static string TruncateLongString(string str, int maxLength)
        {
            if (string.IsNullOrEmpty(str))
                return str;
            return str.Substring(0, Math.Min(str.Length, maxLength));
        }
        public static DataTable GetDataSap(string po_number)
        {

            StringBuilder sb = new StringBuilder();
            sb.AppendLine("select X.[ grn number],X.[grn docentry], CASE WHEN MAX(X.DocCur) != 'SGD' THEN FORMAT(sum(X.[grn total]),'#0.00') else FORMAT(MAX(X.[grn total]),'#0.00') END [grn total], X.[po number]");
            sb.AppendLine("from (");
            sb.AppendLine("SELECT DISTINCT T0.[DocNum] [po number], T3.DocCur, T3.[DocNum] [ grn number], T3.DocEntry [grn docentry], T4.lineNum, ");
            sb.AppendLine("CASE WHEN T3.DocCur != 'SGD' then (T4.PriceBefDi * T4.Quantity) ELSE T3.DocTotal END [grn total]  ");
            sb.AppendLine("FROM OPOR T0 INNER JOIN POR1 T1 ON T0.DocEntry = T1.DocEntry ");
            sb.AppendLine("INNER JOIN PDN1 T2 ON T1.DocEntry= T2.BaseEntry AND T1.LineNum= T2.BaseLine INNER JOIN OPDN T3 ON T2.DocEntry = T3.DocEntry inner join ");
            sb.AppendLine("PDN1 T4 on T4.DocEntry = T3.DocEntry  ");
            sb.AppendLine("where T4.TargetType ='-1')X ");
            sb.AppendLine("where X.[po number] = " + po_number + "  group by X.[ grn number],X.[grn docentry],X.[po number]");


            DataTable dsetItem = new DataTable();
            SqlCommand CmdItem = new SqlCommand(sb.ToString(), SGConnection)
            {
                CommandType = CommandType.Text
            };
            SqlDataAdapter AdptItm = new SqlDataAdapter(CmdItem);
            AdptItm.Fill(dsetItem);
            CmdItem.Dispose();
            AdptItm.Dispose();
            SGConnection.Close();
            return dsetItem;
        }
        public static DataTable GetLineDataSap(string po_number_detail, string line_amount)
        {

            StringBuilder sb = new StringBuilder();


            //sb.AppendLine("select * from (");
            //sb.AppendLine("SELECT  Distinct T0.[DocNum] [po number],T3.[DocNum] [ grn number], T3.DocEntry [grn docentry], T1.ItemCOde,T3.DocTotal,");
            //sb.AppendLine("CASE WHEN T3.DocCur != 'SGD' then (T4.PriceBefDi * T4.Quantity) ELSE T3.DocTotal + T3.DiscSum - T3.VatSum END [grn total] ,CASE WHEN T3.DocCur != 'SGD'  THEN T4.VatSumFrgn ELSE T3.VatSum END as [tax]");
            //sb.AppendLine(" FROM OPOR T0 INNER JOIN POR1 T1 ON T0.DocEntry = T1.DocEntry");
            //sb.AppendLine(" INNER JOIN  PDN1 T4 on  T1.DocEntry = T4.BaseEntry and T1.LineNum = T4.BaseLine ");
            //sb.AppendLine("		  INNER JOIN  OPDN T3 on T3.DocEntry = T4.DocEntry");
            //sb.AppendLine("where T4.TargetType ='-1' )X   where X.[po number] in (" + po_number_detail + ") and X.[grn total] in (" + line_amount + ")");


            sb.AppendLine("select * from (");
            sb.AppendLine("SELECT  Distinct T0.[DocNum] [po number],T3.[DocNum] [ grn number], T3.DocEntry [grn docentry], T1.ItemCOde,T3.DocTotal,");
            sb.AppendLine("CASE WHEN T3.DocCur != 'SGD' then (T4.PriceBefDi * T4.Quantity) ELSE T4.LineTotal END [grn total] ,CASE WHEN T3.DocCur != 'SGD'  THEN T4.VatSumFrgn ELSE T4.VatSum END as [tax]");
            sb.AppendLine(" FROM OPOR T0 INNER JOIN POR1 T1 ON T0.DocEntry = T1.DocEntry");
            sb.AppendLine(" INNER JOIN  PDN1 T4 on  T1.DocEntry = T4.BaseEntry and T1.LineNum = T4.BaseLine ");
            sb.AppendLine("		  INNER JOIN  OPDN T3 on T3.DocEntry = T4.DocEntry");
            //changed GRNTOTAL TO DOCTOTAL sometimes they split the lines 

            sb.AppendLine("where T4.TargetType ='-1' )X   where X.[po number] in (" + po_number_detail + ") and X.[grn total] in (" + line_amount + ")");

            DataTable dsetItem = new DataTable();
            SqlCommand CmdItem = new SqlCommand(sb.ToString(), SGConnection)
            { 
                CommandType = CommandType.Text 
            };
            SqlDataAdapter AdptItm = new SqlDataAdapter(CmdItem);
            AdptItm.Fill(dsetItem);
            CmdItem.Dispose();
            AdptItm.Dispose();
            SGConnection.Close();
            return dsetItem;
        }
        public static DataTable extract_only_po_amount_without_line_amount(string po_number_detail)
        {

            StringBuilder sb = new StringBuilder();

            // get the line amount 
            sb.AppendLine("select * from (");
            sb.AppendLine("SELECT  Distinct T0.[DocNum] [po number],T3.[DocNum] [ grn number], T3.DocEntry [grn docentry], T1.ItemCOde,");
            sb.AppendLine("CASE WHEN T3.DocCur != 'SGD' then (T4.PriceBefDi * T4.Quantity) ELSE T4.LineTotal END [grn total] ,CASE WHEN T3.DocCur != 'SGD'  THEN T4.VatSumFrgn ELSE T4.VatSum END as [tax]");
            sb.AppendLine(" FROM OPOR T0 INNER JOIN POR1 T1 ON T0.DocEntry = T1.DocEntry");
            sb.AppendLine(" INNER JOIN  PDN1 T4 on  T1.DocEntry = T4.BaseEntry and T1.LineNum = T4.BaseLine ");
            sb.AppendLine("		  INNER JOIN  OPDN T3 on T3.DocEntry = T4.DocEntry");
            sb.AppendLine("where T4.TargetType ='-1' )X   where X.[po number] in (" + po_number_detail + ")");

            DataTable dsetItem = new DataTable();
            SqlCommand CmdItem = new SqlCommand(sb.ToString(), SGConnection)
            {
                CommandType = CommandType.Text
            };
            SqlDataAdapter AdptItm = new SqlDataAdapter(CmdItem);
            AdptItm.Fill(dsetItem);
            CmdItem.Dispose();
            AdptItm.Dispose();
            SGConnection.Close();
            return dsetItem;
        }
        private static void send_email(string To, string invoice_number, string seller_name, string docnum, string amount)
        {
            //// Email Part 
            MailMessage mm = new MailMessage
            {
                From = new MailAddress("noreply@syspex.com")
            };
            foreach (var address in To.Split(new[] { ";" }, StringSplitOptions.RemoveEmptyEntries))
            {
                mm.To.Add(address);

            }

            mm.IsBodyHtml = true;
            mm.Subject = "AP Invoice DocNo #" + docnum + "  (" + amount + ") for " + invoice_number + " #" + seller_name;
            mm.Body = "<p>Dear Acccounts Payabale Team,</p> AP Invoice have been auto-created </p>" +
     "<p> Regards,</p>" +
    "<p> DS Team</p> ";

            SmtpClient smtp = new SmtpClient
            {
                Host = "smtp.gmail.com",
                EnableSsl = true
            };
            System.Net.NetworkCredential NetworkCred = new System.Net.NetworkCredential
            {
                UserName = "noreply@syspex.com",
                Password = "design360"
            };
            smtp.UseDefaultCredentials = true;
            smtp.Credentials = NetworkCred;
            smtp.Port = 587;
            smtp.Send(mm);

        }
        public static void SendAutomatedEmail(string htmlString, string recipient)

        {
            try
            {
                //// Email Part 
                MailMessage mm = new MailMessage
                {
                    From = new MailAddress("noreply@syspex.com")
                };
                foreach (var address in recipient.Split(new[] { ";" }, StringSplitOptions.RemoveEmptyEntries))
                {
                    mm.To.Add(address);

                }

                mm.IsBodyHtml = true;
                mm.Subject = "Failed Invoices Not Posted as of  " + DateTime.Now.ToString("dd/MM/yyyy");
                mm.Body = htmlString;
                SmtpClient smtp = new SmtpClient
                {
                    Host = "smtp.gmail.com",
                    EnableSsl = true
                };
                System.Net.NetworkCredential NetworkCred = new System.Net.NetworkCredential
                {
                    UserName = "noreply@syspex.com",
                    Password = "design360"
                };
                smtp.UseDefaultCredentials = true;
                smtp.Credentials = NetworkCred;
                smtp.Port = 587;
                smtp.Send(mm);

            }
            catch (Exception e)
            {

            }

        }
        private static DataTable GetExtractedData()
        {
            string query = @"select Top 10 * from
            [ap_invoice_ocr_extract] where created_date>= day(getdate()) and sap_status = '0' and amount != '' and (sap_docnum= isnull(sap_docnum,'') or sap_docnum is null) and (po_no != '' or po_number_detail !='') and company ='65ST' ";
            //string query = @"select * from [ap_invoice_ocr_extract]   where pdf_file_name = '16012023102133.pdf'";



            DataTable dsetItem = new DataTable();
            SqlCommand CmdItem = new SqlCommand(query, LocalConnection)
            {
                CommandType = CommandType.Text
            };
            SqlDataAdapter AdptItm = new SqlDataAdapter(CmdItem);
            AdptItm.Fill(dsetItem);
            CmdItem.Dispose();
            AdptItm.Dispose();
            LocalConnection.Close();
            return dsetItem;
        }
        public static DataTable Getfilename(string pdf_file_name)
        {
            string query = @"
            select * from  dbo.ap_invoice_ocr_extract where pdf_file_name = '" + pdf_file_name + "'";
            DataTable dsetItem = new DataTable();
            SqlCommand CmdItem = new SqlCommand(query, LocalConnection)
            {
                CommandType = CommandType.Text
            };
            SqlDataAdapter AdptItm = new SqlDataAdapter(CmdItem);
            AdptItm.Fill(dsetItem);
            CmdItem.Dispose();
            AdptItm.Dispose();
            LocalConnection.Close();
            return dsetItem;
        }

        private static void UpdateSAPSTATUS(string pdf_file_name, string sucess, string docnum)
        {
            if (LocalConnection.State == ConnectionState.Closed) { LocalConnection.Open(); }
            SqlCommand CmdOrdStatus = new SqlCommand();
            CmdOrdStatus = new SqlCommand("UPDATE [ap_invoice_ocr_extract] SET sap_status = " + sucess + ", sap_docnum = '" + docnum + "' where  pdf_file_name = '" + pdf_file_name + "'", LocalConnection);
            CmdOrdStatus.CommandType = CommandType.Text;
            CmdOrdStatus.ExecuteNonQuery();
            CmdOrdStatus.Dispose();
            LocalConnection.Close();
        }
        private static void UpdateSAPSTATUS(string docnum)
        {
            docnum = docnum.Trim();
            if (LocalConnection.State == ConnectionState.Closed) { LocalConnection.Open(); }
            SqlCommand CmdOrdStatus = new SqlCommand();
            CmdOrdStatus = new SqlCommand("UPDATE [ap_invoice_ocr_extract] SET file_path = 'F:\\apinvoice\\" + docnum.Replace(" ", "") + ".pdf', pdf_file_name = '" + docnum + ".pdf' where  sap_docnum = '" + docnum + "'", LocalConnection);
            CmdOrdStatus.CommandType = CommandType.Text;
            CmdOrdStatus.ExecuteNonQuery();
            CmdOrdStatus.Dispose();
            LocalConnection.Close();
        }


    }




}

