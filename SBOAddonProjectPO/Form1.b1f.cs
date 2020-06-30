using System;
using System.Collections.Generic;
using System.Xml;
using SAPbouiCOM.Framework;

namespace SBOAddonProjectPO
{
    [FormAttribute("SBOAddonProjectPO.Form1", "Form1.b1f")]
    class Form1 : UserFormBase
    {
        
        public Form1()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Declare(); 
            this.OnCustomInitialize(); 
        }

        public void Declare()
        {
            oApplication = (SAPbouiCOM.Application)Application.SBO_Application;
            oForm = (SAPbouiCOM.Form)oApplication.Forms.ActiveForm;
            oCompany = (SAPbobsCOM.Company)oApplication.Company.GetDICompany();

            this.lblPONo = ((SAPbouiCOM.StaticText)(this.GetItem("lblPONo").Specific));
            this.txtPONo = ((SAPbouiCOM.EditText)(this.GetItem("txtPONo").Specific));
            this.lbtnPO = ((SAPbouiCOM.LinkedButton)(this.GetItem("lbtnPO").Specific));
            this.matrixItem = ((SAPbouiCOM.Matrix)(this.GetItem("matrixItem").Specific));
            this.btnFind = ((SAPbouiCOM.Button)(this.GetItem("btnFind").Specific));
            this.btnFind.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.btnFind_ClickBefore);
            this.btnAdd = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.btnAdd.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.btnAdd_ClickBefore);
            this.btnCancel = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));

            ButtonsDisable();
        }

        public void ButtonsEnable()
        {
            btnAdd.Item.Enabled = true;
            btnCancel.Item.Enabled = true;
        }
        public void ButtonsDisable()
        {
            btnAdd.Item.Enabled = false;
            btnCancel.Item.Enabled = false;
        }
        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
        }

        private void OnCustomInitialize()
        {
        }

        private void btnFind_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            oRecordset = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            if (string.IsNullOrEmpty(txtPONo.Value.ToString()))
            {
                //BubbleEvent = false;
                Application.SBO_Application.SetStatusBarMessage("Please select a Purchase Order Number", 
                                                                    SAPbouiCOM.BoMessageTime.bmt_Medium, true);
            }
            else
            {
                ButtonsEnable();
                docEntry = txtPONo.Value.ToString();
                query = "select t0.ItemCode, t0.Dscription, t0.Quantity from POR1 t0 left join OPOR t1 on t1.DocEntry = t0.DocEntry where t1.DocEntry=" + docEntry;
                oRecordset.DoQuery(query);

                if (oRecordset.RecordCount > 0)
                {
                    for (int i = 0; i < oRecordset.RecordCount; i++)
                    {
                        matrixItem.AddRow();

                        ((SAPbouiCOM.EditText)matrixItem.Columns.Item("colCode").Cells.Item(i + 1).Specific).Value = 
                            oRecordset.Fields.Item("ItemCode").Value.ToString();
                        ((SAPbouiCOM.EditText)matrixItem.Columns.Item("colName").Cells.Item(i + 1).Specific).Value = 
                            oRecordset.Fields.Item("Dscription").Value.ToString();
                        ((SAPbouiCOM.EditText)matrixItem.Columns.Item("colQty").Cells.Item(i + 1).Specific).Value = 
                            oRecordset.Fields.Item("Quantity").Value.ToString();

                        oRecordset.MoveNext();
                    }
                }
            }
        }
        private void btnAdd_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
            {
                if (string.IsNullOrEmpty(txtPONo.Value.ToString()))
                {
                    Application.SBO_Application.SetStatusBarMessage("Please select a Purchase Order Number", 
                                                                        SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                }
                else
                {
                    ButtonsEnable();

                    for (int i = 0; i < oRecordset.RecordCount; i++)
                    {
                        oUserTable = oCompany.UserTables.Item("GRQC");

                        oUserTable.Code = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                        oUserTable.Name = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                        
                        oUserTable.UserFields.Fields.Item("U_PONo").Value = txtPONo.Value.ToString();

                        oUserTable.UserFields.Fields.Item("U_ItemCode").Value = 
                            ((SAPbouiCOM.EditText)matrixItem.Columns.Item("colCode").Cells.Item(i + 1).Specific).Value.ToString();

                        oUserTable.UserFields.Fields.Item("U_ItemName").Value =
                           ((SAPbouiCOM.EditText)matrixItem.Columns.Item("colName").Cells.Item(i + 1).Specific).Value.ToString();

                        oUserTable.UserFields.Fields.Item("U_Quantity").Value =
                           ((SAPbouiCOM.EditText)matrixItem.Columns.Item("colQty").Cells.Item(i + 1).Specific).Value.ToString();

                        oUserTable.UserFields.Fields.Item("U_QtyRejected").Value = 
                            ((SAPbouiCOM.EditText)matrixItem.Columns.Item("colQtyR").Cells.Item(i + 1).Specific).Value.ToString();

                        int r = oUserTable.Add();

                        if (r != 0)
                        {
                            oApplication.SetStatusBarMessage("Error" + oCompany.GetLastErrorDescription(), 
                                                                SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                        }
                        else
                        {
                            oApplication.SetStatusBarMessage("Record inserted successfully", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                        }
                    }                    
                }
            }
            BubbleEvent = false;
            Clear();
        }

        public void Clear()
        {
            txtPONo.Value = string.Empty;
            matrixItem.Clear();
        }

        public SAPbobsCOM.Recordset oRecordset;
        public SAPbobsCOM.Company oCompany;
        public SAPbouiCOM.Form oForm;
        public SAPbouiCOM.Application oApplication;
        public SAPbobsCOM.UserTable oUserTable;

        private string docEntry;
        private string query;

        private SAPbouiCOM.StaticText lblPONo;
        private SAPbouiCOM.EditText txtPONo;
        private SAPbouiCOM.LinkedButton lbtnPO;
        private SAPbouiCOM.Button btnFind, btnAdd, btnCancel;
        private SAPbouiCOM.Matrix matrixItem;
    }
}