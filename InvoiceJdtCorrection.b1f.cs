using System;
using System.Collections.Generic;
using System.Reflection;
using System.Xml;
using SAPApi;
using SAPbobsCOM;
using SAPbouiCOM;
using SAPbouiCOM.Framework;
using Application = SAPbouiCOM.Framework.Application;

namespace Invoice_Income_Correction
{
    [FormAttribute("Invoice_Income_Correction.InvoiceJdtCorrection", "InvoiceJdtCorrection.b1f")]
    class InvoiceJdtCorrection : UserFormBase
    {
        public InvoiceJdtCorrection()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_0").Specific));
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("Item_1").Specific));
            this.ComboBox0.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.ComboBox0_ComboSelectAfter);
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_2").Specific));
            this.ComboBox1 = ((SAPbouiCOM.ComboBox)(this.GetItem("Item_3").Specific));
            this.ComboBox1.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.ComboBox1_ComboSelectAfter);
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_4").Specific));
            this.ComboBox2 = ((SAPbouiCOM.ComboBox)(this.GetItem("Item_5").Specific));
            this.ComboBox2.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.ComboBox2_ComboSelectAfter);
            this.Grid0 = ((SAPbouiCOM.Grid)(this.GetItem("Item_6").Specific));
            this.Grid0.DoubleClickAfter += new SAPbouiCOM._IGridEvents_DoubleClickAfterEventHandler(this.Grid0_DoubleClickAfter);
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_7").Specific));
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
            this.PictureBox1 = ((SAPbouiCOM.PictureBox)(this.GetItem("Item_11").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
        }

        private SAPbouiCOM.StaticText StaticText0;

        DIManager _diManager = new DIManager(Program.DiCompany);
        UiManager _uiManager = new UiManager();


        private void OnCustomInitialize()
        {
            _uiManager.GetLastYears(2018, ComboBox0, DateTime.Now.Year + 2);
            _uiManager.SetMonthNames(ComboBox1);
            ComboBox2.ValidValues.Add("1", "Incoming Payments");
            ComboBox2.ValidValues.Add("2", "Outgoing Payments");
            ComboBox0.ExpandType = BoExpandType.et_ValueOnly;
            ComboBox0.Select(DateTime.Now.Year.ToString());
            ComboBox1.Select(DateTime.Now.Month.ToString());
            ComboBox2.Select("1");
            RefreshForm();
            GetItem("Item_7").FontSize = 10;
            string path = System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
            PictureBox1.Picture = "" + path + "\\RSMxxxx.jpg";
     //       PictureBox1.Picture = @"D:\Users\nkurdadze\Desktop\RSMxxxx.jpg";
        }

        private void RefreshForm()
        {
            if (ComboBox2.Selected == null || ComboBox0.Selected == null || ComboBox1.Selected == null)
            {
                return;
            }

            string invType = ComboBox2.Selected.Value == "1" ? "13" : "18";

            if (invType == "13")
            {


                DateTime dt = new DateTime(Convert.ToInt32(ComboBox0.Value), Convert.ToInt32(ComboBox1.Value), 1);

                string databaseQuery = $@"SELECT ROW_NUMBER() OVER (ORDER BY U_CardCode) as [#], U_CardCode as [CardCode] , U_CardName as [CardName],    U_PaymentDocEntry as [PaymentDocEntry], U_PaymentDocNum as [Payment Number], U_Select as [Select],  U_PaymentTransId as [Payment JournalEntry Number], U_CorrectionTransId as [JournalEntry Number]  FROM [@RSM_IPHY] WHERE U_DocDate = '{dt:s}' and U_InvType = {invType}";

                string generatedQuery = $@" SELECT  ROW_NUMBER() OVER (ORDER BY a.CardCode) as [#] ,  * FROM (SELECT distinct  ORCT.CardCode, ORCT.CardName,     ORCT.DocEntry as [PaymentDocEntry], ORCT.DocNum as [Payment Number], 'N' as [Select],  JDT1.TransId as [Payment JournalEntry Number], '           ' as [JournalEntry Number]            
               FROM ORCT 
		       JOIN RCT2 on RCT2.DocNum = ORCT.DocEntry  
			   JOIN OCRD ON OCRD.CardCode = ORCT.CardCode      
               JOIN JDT1 on ORCT.TransID = JDT1.TransId   

                WHERE YEAR(ORCT.DocDate) = {ComboBox0.Selected.Value} AND MONTH(ORCT.DocDate) = {ComboBox1.Selected.Value}   AND  U_FixedRatePayer = N'კი'  AND ORCT.Canceled = 'N'  AND ((Account = '{Program.ExchangeGain}' AND JDT1.Credit != 0) OR (Account = '{Program.ExchangeLoss}' AND JDT1.Debit != 0 and RCT2.InvType = '13' ))) a";

                string MergedQuerys =
                    "select boo.#, boo.CardCode, boo.CardName,  boo.[Payment JournalEntry Number], boo.[Payment Number], boo.PaymentDocEntry, boo.[Select], foo.[JournalEntry Number] from (";
                MergedQuerys += generatedQuery + ") as boo left join (";

                MergedQuerys += databaseQuery +
                          ") as foo on foo.[Payment JournalEntry Number] = boo.[Payment JournalEntry Number]";
                Grid0.DataTable.ExecuteQuery(DIManager.QueryHanaTransalte(MergedQuerys));
                SAPbouiCOM.EditTextColumn incomingPayment = (EditTextColumn)Grid0.Columns.Item("PaymentDocEntry");
                incomingPayment.LinkedObjectType =   "24" ;
            }

         
           
            else
            {
                DateTime dt = new DateTime(Convert.ToInt32(ComboBox0.Value), Convert.ToInt32(ComboBox1.Value), 1);

                string databaseQuery = $@"SELECT ROW_NUMBER() OVER (ORDER BY U_CardCode) as [#], U_CardCode as [CardCode] , U_CardName as [CardName],    U_PaymentDocEntry as [PaymentDocEntry], U_PaymentDocNum as [Payment Number], U_Select as [Select],  U_PaymentTransId as [Payment JournalEntry Number], U_CorrectionTransId as [JournalEntry Number]  FROM [@RSM_IPHY] WHERE U_DocDate = '{dt:s}' and U_InvType = {invType}";

                string generatedQuery = $@" SELECT  ROW_NUMBER() OVER (ORDER BY a.CardCode) as [#] ,  * FROM (SELECT distinct  OVPM.CardCode, OVPM.CardName,     OVPM.DocEntry as [PaymentDocEntry], OVPM.DocNum as [Payment Number], 'N' as [Select],  JDT1.TransId as [Payment JournalEntry Number], '           ' as [JournalEntry Number]            
               FROM OVPM 
		       JOIN VPM2 on VPM2.DocNum = OVPM.DocEntry  
			   JOIN OCRD ON OCRD.CardCode = OVPM.CardCode      
               JOIN JDT1 on OVPM.TransID = JDT1.TransId   

                WHERE YEAR(OVPM.DocDate) = {ComboBox0.Selected.Value} AND MONTH(OVPM.DocDate) = {ComboBox1.Selected.Value}   AND  U_FixedRatePayer = N'კი'  AND OVPM.Canceled = 'N'  AND ((Account = '{Program.ExchangeGain}' AND JDT1.Credit != 0) OR (Account = '{Program.ExchangeLoss}' AND JDT1.Debit != 0 and VPM2.InvType = '18' ))) a";

                string MergedQuerys =
                    "select boo.#, boo.CardCode, boo.CardName,  boo.[Payment JournalEntry Number], boo.[Payment Number], boo.PaymentDocEntry, boo.[Select], foo.[JournalEntry Number] from (";
                MergedQuerys += generatedQuery + ") as boo left join (";

                MergedQuerys += databaseQuery +
                                ") as foo on foo.[Payment JournalEntry Number] = boo.[Payment JournalEntry Number]";
                Grid0.DataTable.ExecuteQuery(DIManager.QueryHanaTransalte(MergedQuerys));

                SAPbouiCOM.EditTextColumn outgoingPayment = (EditTextColumn)Grid0.Columns.Item("PaymentDocEntry");
                outgoingPayment.LinkedObjectType = "46";
            }


            Grid0.Columns.Item("Select").Type = BoGridColumnType.gct_CheckBox;
            Grid0.Columns.Item("CardCode").Editable = false;
            Grid0.Columns.Item("CardName").Editable = false;
            Grid0.Columns.Item("PaymentDocEntry").Editable = false;
            Grid0.Columns.Item("JournalEntry Number").Editable = false;
            Grid0.Columns.Item("Payment Number").Editable = false;
            Grid0.Columns.Item("Payment JournalEntry Number").Editable = false;
            Grid0.Columns.Item("#").Editable = false;
            Grid0.Columns.Item("Select").Editable = true;
            SAPbouiCOM.EditTextColumn journalEntry1 = (EditTextColumn)Grid0.Columns.Item("JournalEntry Number");
            journalEntry1.LinkedObjectType = "30";
            SAPbouiCOM.EditTextColumn bpColumn = (EditTextColumn)Grid0.Columns.Item("CardCode");
            bpColumn.LinkedObjectType = "2";

            SAPbouiCOM.EditTextColumn journalEntry = (EditTextColumn)Grid0.Columns.Item("Payment JournalEntry Number");
            journalEntry.LinkedObjectType = "30";


        }

        private SAPbouiCOM.ComboBox ComboBox0;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.ComboBox ComboBox1;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.ComboBox ComboBox2;
        private SAPbouiCOM.Grid Grid0;
        private Button Button0;

        private void Button0_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {
            int clicked = SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("ნამდვილად გსურთ მაკორეკტირებელი გატარებების გაკეთება?", 1, "დიახ", "არა");
            if (clicked == 2)
            {
                return;
            }

            SAPbobsCOM.JournalEntries journalEntryPayment =
                (SAPbobsCOM.JournalEntries)Program.DiCompany.GetBusinessObject(
                    SAPbobsCOM.BoObjectTypes.oJournalEntries);
            bool isIncoming = ComboBox2.Value == "1";
            for (int i = 0; i < Grid0.Rows.Count; i++)
            {
                bool checkboxChecked = Grid0.DataTable.Columns.Item("Select").Cells.Item(i).Value.ToString() == "Y";
                string cardCode = Grid0.DataTable.Columns.Item("CardCode").Cells.Item(i).Value.ToString();
                string CardName = Grid0.DataTable.Columns.Item("CardName").Cells.Item(i).Value.ToString();
                string PaymentDocEntry = Grid0.DataTable.Columns.Item("PaymentDocEntry").Cells.Item(i).Value.ToString();
                string PaymentNumber = Grid0.DataTable.Columns.Item("Payment Number").Cells.Item(i).Value.ToString();
                string Select = Grid0.DataTable.Columns.Item("Select").Cells.Item(i).Value.ToString();
                string paymentTransId = Grid0.DataTable.Columns.Item("Payment JournalEntry Number").Cells.Item(i).Value.ToString();
                string CorrectionTransId = string.IsNullOrWhiteSpace(Grid0.DataTable.Columns.Item("JournalEntry Number").Cells.Item(i).Value.ToString()) ? "0" : Grid0.DataTable.Columns.Item("JournalEntry Number").Cells.Item(i).Value.ToString();


                Recordset recSet = (Recordset)Program.DiCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                recSet.DoQuery("select DebPayAcct from OCRD where CardCode = '" + cardCode + "'");
                string BpControlAcc = recSet.Fields.Item("DebPayAcct").Value.ToString();
                string result = "0";
                Recordset recordset =
                    (Recordset)Program.DiCompany.GetBusinessObject(BoObjectTypes
                        .BoRecordset);

                recordset.DoQuery(DIManager.QueryHanaTransalte($@"SELECT * FROM [@RSM_IPHY] WHERE U_PaymentTransId = {paymentTransId}"));

                string corTr = recordset.Fields.Item("U_CorrectionTransId").Value.ToString();

                if (!recordset.EoF && corTr != "0")
                {
                    continue;
                }

                if (checkboxChecked)
                {
                    int transId = Convert.ToInt32(Grid0.DataTable.Columns.Item("Payment JournalEntry Number").Cells
                        .Item(i).Value);
                    journalEntryPayment.GetByKey(transId);
                    double exchangeRateGainCredit = 0;
                    double exchangeRateLossDebit = 0;

                    for (int j = 0; j < journalEntryPayment.Lines.Count; j++)
                    {
                        journalEntryPayment.Lines.SetCurrentLine(j);
                        if (journalEntryPayment.Lines.AccountCode == Program.ExchangeGain)
                        {
                            exchangeRateGainCredit += journalEntryPayment.Lines.Credit;
                        }
                        if (journalEntryPayment.Lines.AccountCode == Program.ExchangeLoss)
                        {
                            exchangeRateLossDebit += journalEntryPayment.Lines.Debit;
                        }
                    }


                    if (exchangeRateLossDebit == 0)
                    {
                        result = addJournalEntry(journalEntryPayment.ReferenceDate, journalEntryPayment.DueDate,
                            journalEntryPayment.TaxDate
                            , journalEntryPayment.Reference, journalEntryPayment.Lines.BPLID, Program.ExchangeGain,
                            exchangeRateGainCredit, BpControlAcc, cardCode, isIncoming);
                        Grid0.DataTable.SetValue("JournalEntry Number", i, result);
                        SAPbouiCOM.EditTextColumn journalEntry = (EditTextColumn)Grid0.Columns.Item("JournalEntry Number");
                        journalEntry.LinkedObjectType = "30";
                    }
                    if (exchangeRateGainCredit == 0)
                    {
                        result = addJournalEntry(journalEntryPayment.ReferenceDate, journalEntryPayment.DueDate,
                            journalEntryPayment.TaxDate
                            , journalEntryPayment.Reference, journalEntryPayment.Lines.BPLID, Program.ExchangeLoss,
                            exchangeRateLossDebit, BpControlAcc, cardCode, isIncoming);
                        Grid0.DataTable.SetValue("JournalEntry Number", i, result);
                        SAPbouiCOM.EditTextColumn journalEntry = (EditTextColumn)Grid0.Columns.Item("JournalEntry Number");
                        journalEntry.LinkedObjectType = "30";
                    }

                }



                if (!recordset.EoF && corTr == "0")
                {
                    Dictionary<string, dynamic> iphyInster = new Dictionary<string, dynamic>();
                    iphyInster.Add("CardCode", cardCode);
                    iphyInster.Add("CardName", CardName);
                    iphyInster.Add("PaymentDocEntry", PaymentDocEntry);
                    iphyInster.Add("PaymentDocNum", PaymentNumber);
                    iphyInster.Add("Select", Select);
                    iphyInster.Add("PaymentTransId", paymentTransId);
                    iphyInster.Add("CorrectionTransId", result);
                    iphyInster.Add("DocDate", new DateTime(Convert.ToInt32(ComboBox0.Value), Convert.ToInt32(ComboBox1.Value), 1));
                    iphyInster.Add("InvType", isIncoming ? "13" : "18");

                    DIManager.DbUpdate("RSM_IPHY", iphyInster, Program.DiCompany, recordset.Fields.Item("Code").Value.ToString());
                }

                if (recordset.EoF)
                {

                    Dictionary<string, dynamic> iphyInster = new Dictionary<string, dynamic>();
                    iphyInster.Add("CardCode", cardCode);
                    iphyInster.Add("CardName", CardName);
                    iphyInster.Add("PaymentDocEntry", PaymentDocEntry);
                    iphyInster.Add("PaymentDocNum", PaymentNumber);
                    iphyInster.Add("Select", Select);
                    iphyInster.Add("PaymentTransId", paymentTransId);
                    iphyInster.Add("CorrectionTransId", result);
                    iphyInster.Add("DocDate", new DateTime(Convert.ToInt32(ComboBox0.Value), Convert.ToInt32(ComboBox1.Value), 1));
                    iphyInster.Add("InvType", isIncoming ? "13" : "18");

                    DIManager.DbInsert("RSM_IPHY", iphyInster, Program.DiCompany);
                }


            }
        }

        private bool _checkeD = false;

        private void Grid0_DoubleClickAfter(object sboObject, SBOItemEventArg pVal)
        {
            if (pVal.Row == -1 && pVal.ColUID == "Select")
            {
                if (_checkeD)
                {
                    for (int i = 0; i < Grid0.Rows.Count; i++)
                    {
                        SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.Freeze(true);
                        Grid0.DataTable.SetValue(pVal.ColUID, i, "N");
                        SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.Freeze(false);
                    }
                    _checkeD = false;
                }
                else if (!_checkeD)
                {
                    for (int i = 0; i < Grid0.Rows.Count; i++)
                    {
                        SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.Freeze(true);
                        Grid0.DataTable.SetValue(pVal.ColUID, i, "Y");
                        SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.Freeze(false);
                    }
                    _checkeD = true;
                }
            }
        }

        private void ComboBox0_ComboSelectAfter(object sboObject, SBOItemEventArg pVal)
        {
            RefreshForm();
        }

        private void ComboBox1_ComboSelectAfter(object sboObject, SBOItemEventArg pVal)
        {
            RefreshForm();
        }

        private void ComboBox2_ComboSelectAfter(object sboObject, SBOItemEventArg pVal)
        {
            RefreshForm();
        }




        public string addJournalEntry(DateTime ReferenceDate, DateTime DueDate, DateTime TaxDate, string Reference, int BPLID, string AccountCode, double Amount, string ControlAccount, string ControlAccName, bool isIncoming)
        {

            SAPbobsCOM.JournalEntries journalEntries = (SAPbobsCOM.JournalEntries)Program.DiCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
            journalEntries.ReferenceDate = ReferenceDate;
            journalEntries.DueDate = DueDate;
            journalEntries.TaxDate = TaxDate;
            journalEntries.Memo = "JDT CORRECTION";
            journalEntries.Reference = Reference;
            journalEntries.TransactionCode = "3";

            journalEntries.Lines.BPLID = BPLID;
            journalEntries.Lines.AccountCode = AccountCode;
            journalEntries.Lines.Debit = AccountCode == Program.ExchangeGain ? 0 : -Amount;
            journalEntries.Lines.Credit = AccountCode == Program.ExchangeGain ? -Amount : 0;
            journalEntries.Lines.FCCredit = 0;
            journalEntries.Lines.FCDebit = 0;
            journalEntries.Lines.Add();

            journalEntries.Lines.BPLID = BPLID;
            journalEntries.Lines.ControlAccount = ControlAccount;
            journalEntries.Lines.ShortName = ControlAccName;
            journalEntries.Lines.Debit = isIncoming ? 0 : AccountCode == Program.ExchangeGain ? -Amount : Amount;
            journalEntries.Lines.Credit = isIncoming ? AccountCode == Program.ExchangeLoss ? -Amount : Amount : 0;
            journalEntries.Lines.FCCredit = 0;
            journalEntries.Lines.FCDebit = 0;

            journalEntries.Lines.Add();
            int iq = journalEntries.Add();
            if (iq == 0)
            {
                return Program.DiCompany.GetNewObjectKey();
            }
            else
            {
                Application.SBO_Application.SetStatusBarMessage(Program.DiCompany.GetLastErrorDescription(),
                    BoMessageTime.bmt_Short, true);
                return "0";
            }
        }

        private PictureBox PictureBox1;

      
    }
}