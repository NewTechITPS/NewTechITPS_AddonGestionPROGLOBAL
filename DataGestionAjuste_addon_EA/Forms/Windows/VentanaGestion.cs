using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using PROGLOBAL_DataGestionAjuste_addon_EA.Services;
using PROGLOBAL_DataGestionAjuste_addon_EA.Common;
using SAPbobsCOM;
using System.Linq.Expressions;
using REDFARM.Addons.Tools;
using PROGLOBAL_ReservationInvoiceCloser.Services;
using System.Globalization;
using PROGLOBAL_DataGestionAjuste_addon_EA.Models;
using System.Windows.Forms;
using System.Collections;
using System.Data;
using DocumentFormat.OpenXml.Office.SpreadSheetML.Y2023.MsForms;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Vml;

namespace PROGLOBAL_DataGestionAjuste_addon_EA.Forms.WINDOW
{

    public class VentanaGestion 
    {

        #region Atributos

        private static SAPbouiCOM.Form? _oForm;

        private static ReportExcelFormat? _reportExcelFormat;

        public const string frmUID = "60004"; 
        public const string menuUID = "VentanaGestion";

        private string _itemDateFrom = "Item_0";
        private string _itemDateTo = "Item_1";
        private string _itemBtnFilter = "Item_4";
        private string _itemBtnExport = "Item_5";
        private string _itemBtnApplyCommision = "Item_18";
        private string _itemGridGastos = "Item_8";
        private string _itemGridTotales = "Item_10";
        private string _itemGridVentas = "Item_12";
        private string _itemBtnApplyAjuste = "Item_13";
        private string _itemBtnSave = "Item_14";
        private string _itemGridSavedAjustes = "Item_15";
        private string _itemLoading = "Item_19";
        private string _itemSolapaVentas = "Item_11";
        private string _itemSolapaGastos = "Item_7";
        private string _itemSolapaTotalVentas = "Item_9";
        private string _itemSolapaTotalGastos = "Item_17";


        private string _colAcumulado = "Acumulado (1)";
        private string _colPorcAcum = "% s. ventas (5)";
        private string _colMensual = "Mensual";
        private string _colComision = "Comisiones";
        private string _colVentas = "Ventas";
        private string _colDetalle = "Detalle";
        #endregion

        public void OSAPB1appl_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if(pVal.MenuUID == menuUID && !pVal.BeforeAction)
            {
                VentanaGestionService.CreateWindow();
            }
            
        }

        public void OSAPB1appl_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (pVal.EventType == BoEventTypes.et_FORM_LOAD && pVal.ActionSuccess)
            {
                try
                {

                    _reportExcelFormat = new ReportExcelFormat();
                    _reportExcelFormat.FileName = $"{DateTime.Now.ToString("FFFFFFF")}_informeGestion_{DateTime.Now.ToString("yyyy-MM-dd")}_form{pVal.FormTypeCount}";

                    VentanaGestionService.FormUID = FormUID;
                    _oForm = ConnectionSDK.UIAPI!.Forms.Item(FormUID);

                } catch(Exception ex)
                {
                    ConnectionSDK.UIAPI!.MessageBox(ex.Message);
                }
            }

            if (pVal.EventType == BoEventTypes.et_FORM_RESIZE && pVal.ActionSuccess)
            {
                _oForm = ConnectionSDK.UIAPI!.Forms.Item(FormUID);
                Grid GAjusteSave = _oForm.Items.Item(_itemGridSavedAjustes).Specific;
                GAjusteSave.Item.Width = 180;
                GAjusteSave.Item.Height = 153;
            }

            // VALIDACION CAMPOS FECHAS
            if (pVal.EventType == BoEventTypes.et_VALIDATE && (pVal.ItemUID == _itemDateFrom || pVal.ItemUID == _itemDateTo) && pVal.ActionSuccess)
            {
                _oForm = ConnectionSDK.UIAPI!.Forms.Item(FormUID);
                SAPbouiCOM.Item oItem = _oForm!.Items.Item(_itemDateFrom);
                SAPbouiCOM.EditText ETDateFrom = oItem.Specific;

                oItem = _oForm.Items.Item(_itemDateTo);
                SAPbouiCOM.EditText ETDateTo = oItem.Specific;

                string valueDateFrom = ETDateFrom.Value;
                string valueDateTo = ETDateTo.Value;

                oItem = _oForm.Items.Item(_itemBtnFilter);
                oItem.Enabled = !string.IsNullOrEmpty(valueDateFrom) && !string.IsNullOrEmpty(valueDateTo);
                
            }


            if (pVal.EventType == BoEventTypes.et_FORM_CLOSE && pVal.ActionSuccess)
            {
                VentanaGestionService.TruncateUDOGestionAjuste();
            }

            // PRESIONAR BOTON "FILTRAR"
            if (pVal.EventType == BoEventTypes.et_ITEM_PRESSED && pVal.ItemUID == _itemBtnFilter && pVal.ActionSuccess) 
            {
                _oForm = ConnectionSDK.UIAPI!.Forms.Item(FormUID);
                SAPbouiCOM.Item oItemLoading = _oForm!.Items.Item(_itemLoading);
                oItemLoading.Visible = true;
                _oForm.Freeze(true);

                try
                {
                    var sheet = VentanaGestionService.CreateSheet();

                    
                    // GASTOS
                    VentanaGestionService.RefreshDataGastosGrid();
                    VentanaGestionService.CreateColumnsInDataTableExpenses(sheet.DataTableExpenses);
                    VentanaGestionService.LoadDataInDataTableExpenses(sheet);

                    // VENTAS
                    VentanaGestionService.RefreshDataVentasGrid();
                    VentanaGestionService.CreateColumnsInDataTableSales(sheet.DataTableSales);  
                    VentanaGestionService.LoadDataInDataTableSales(sheet);

                    // TOTALES
                    var totals = VentanaGestionService.RefreshDataTotalesVentasGrid(sheet);

                    VentanaGestionService.RefreshDataTotalesGastosGrid(totals);


                    //VentanaGestionService.CreateColumnsInDataTableTotales(sheet.DataTableTotals); 
                    //VentanaGestionService.LoadDataInDataTableTotals(sheet);


                    SAPbouiCOM.Item oItemBtnAjuste = _oForm.Items.Item(_itemBtnApplyAjuste);
                    SAPbouiCOM.Item oItemBtnSave = _oForm.Items.Item(_itemBtnSave);
                    oItemBtnAjuste.Enabled = true;
                    oItemBtnSave.Enabled = true;

                } catch (Exception ex)
                {
                    NotificationService.Error("Error al aplicar el filtro de fecha; Mensaje ->" + ex.Message);
                } finally
                {
                    oItemLoading.Visible = false;
                    _oForm.Freeze(false);
                }
            }

            // PRESIONAR BOTON "APLICAR AJUSTE"
            if(pVal.EventType == BoEventTypes.et_ITEM_PRESSED && pVal.ItemUID == _itemBtnApplyAjuste && pVal.ActionSuccess) 
            {
                _oForm = ConnectionSDK.UIAPI!.Forms.Item(FormUID);
                SAPbouiCOM.Item oItemLoading = _oForm!.Items.Item(_itemLoading);
                oItemLoading.Visible = true;
                _oForm.Freeze(true);
                try
                {
                    var sheet = VentanaGestionService.CreateSheet();

                    sheet.DataTableExpenses = new System.Data.DataTable();
                    sheet.DataTableSales = new System.Data.DataTable();
                    sheet.DataTableTotalsSales = new System.Data.DataTable();
                    sheet.DataTableTotalsExpenses = new System.Data.DataTable();

                    VentanaGestionService.InsertRecordsUDOGestionAjuste();

                    // GASTOS
                    VentanaGestionService.RefreshDataGastosGrid();
                    VentanaGestionService.CreateColumnsInDataTableExpenses(sheet.DataTableExpenses);
                    VentanaGestionService.LoadDataInDataTableExpenses(sheet);

                    // VENTAS
                    VentanaGestionService.RefreshDataVentasGrid();
                    VentanaGestionService.CreateColumnsInDataTableSales(sheet.DataTableSales);
                    VentanaGestionService.LoadDataInDataTableSales(sheet);

                    // TOTALES
                    var totals = VentanaGestionService.RefreshDataTotalesVentasGrid(sheet);

                    VentanaGestionService.RefreshDataTotalesGastosGrid(totals);

                    ConnectionSDK.UIAPI!.MessageBox("Ajuste aplicado con éxito");

                } catch (Exception ex)
                {
                    NotificationService.Error("Error al aplicar ajuste; Mensaje ->" + ex.Message);
                } finally
                {
                    oItemLoading.Visible = false;
                    _oForm.Freeze(false);
                }
                
            }

            // PRESIONAR BOTON "GUARDAR"
            if(pVal.EventType == BoEventTypes.et_ITEM_PRESSED && pVal.ItemUID == _itemBtnSave && pVal.ActionSuccess)  
            {
                _oForm = ConnectionSDK.UIAPI!.Forms.Item(FormUID);
                SAPbouiCOM.Item oItemLoading = _oForm.Items.Item(_itemLoading);
                oItemLoading.Visible = true;

                _oForm!.Freeze(true);
                try
                {
                    var sheet = VentanaGestionService.CreateSheet();
                    
                    VentanaGestionService.CreateColumnsInDataTableExpenses(sheet.DataTableExpenses);
                    VentanaGestionService.LoadDataInDataTableExpenses(sheet);
                    
                    VentanaGestionService.CreateColumnsInDataTableSales(sheet.DataTableSales);
                    VentanaGestionService.LoadDataInDataTableSales(sheet);

                    VentanaGestionService.CreateColumnsInDataTableTotales(sheet.DataTableTotalsSales);
                    VentanaGestionService.LoadDataInDataTableTotals(sheet);

                    VentanaGestionService.ResetGestionAjuste();
                    VentanaGestionService.LoadSheetNameInGrid(sheet);

                    _reportExcelFormat!.Sheets.Add(sheet);

                    ConnectionSDK.UIAPI!.MessageBox("Ajuste guardado con éxito -> Ajuste: " + sheet.SheetName);

                    SAPbouiCOM.Item oItemBtnExport = _oForm!.Items.Item(_itemBtnExport);
                    oItemBtnExport.Enabled = true;

                    SAPbouiCOM.Item oItemBtnAjuste = _oForm.Items.Item(_itemBtnApplyAjuste);
                    SAPbouiCOM.Item oItemBtnSave = _oForm.Items.Item(_itemBtnSave);
                    oItemBtnAjuste.Enabled = false;
                    oItemBtnSave.Enabled = false;

                } catch(Exception ex)
                {
                    NotificationService.Error("Error al guardar; Mensaje ->" + ex.Message);
                }
                finally
                {
                    _oForm!.Freeze(false);
                    oItemLoading.Visible = false;
                }

            }

            // PRESIONAR BOTON "EXPORTAR EXCEL"
            if (pVal.EventType == BoEventTypes.et_COMBO_SELECT && pVal.ItemUID == _itemBtnExport && pVal.ActionSuccess)  
            {
                _oForm = ConnectionSDK.UIAPI!.Forms.Item(FormUID);
                ButtonCombo oBCExport = (ButtonCombo)_oForm!.Items.Item(_itemBtnExport).Specific;

                if (oBCExport.Selected.Value == "Exportar Excel")
                {
                    string pathFile = VentanaGestionService.GetPathToSaveFile(_reportExcelFormat!.FileName);

                    if (pathFile != null)
                    {
                        VentanaGestionService.ExportExcel(_reportExcelFormat, pathFile);

                        NotificationService.Success("Documento creado con exito");
                    }
                }

            }

            // CALCULAR ACUMULADO Y SU PORCENTAJE
            if (pVal.EventType == BoEventTypes.et_VALIDATE && pVal.ColUID == _colComision && pVal.ItemUID == _itemGridTotales && pVal.ActionSuccess)
            {
                _oForm = ConnectionSDK.UIAPI!.Forms.Item(pVal.FormUID);

                _oForm.Freeze(true);
                Grid GTotales = _oForm.Items.Item(_itemGridTotales).Specific;

                double mensual = GTotales.DataTable.GetValue(_colMensual, pVal.Row);
                double comision = GTotales.DataTable.GetValue(_colComision, pVal.Row);
                double ventas = GTotales.DataTable.GetValue(_colVentas, pVal.Row);

                double acumulado = mensual + comision;
                GTotales.DataTable.Columns.Item(_colAcumulado).Cells.Item(pVal.Row).Value = acumulado;
                GTotales.DataTable.Columns.Item(_colPorcAcum).Cells.Item(pVal.Row).Value = ventas != 0 ? acumulado / ventas * 100 : 0;
                _oForm.Freeze(false);
            }


            if (pVal.EventType == BoEventTypes.et_ITEM_PRESSED && pVal.ItemUID == _itemBtnApplyCommision && pVal.ActionSuccess)
            {
                VentanaGestionService.CalculateTotals_Comisiones_Acumulado_PorcAcumulado();
            }

            if (pVal.EventType == BoEventTypes.et_ITEM_PRESSED && pVal.ActionSuccess)
            {
                _oForm = ConnectionSDK.UIAPI!.Forms.Item(pVal.FormUID);
                SAPbouiCOM.Item itemApplyAjuste = _oForm.Items.Item(_itemBtnApplyAjuste);
                SAPbouiCOM.Item itemApplyCommision = _oForm.Items.Item(_itemBtnApplyCommision);

                Folder solapaVentas = _oForm.Items.Item(_itemSolapaVentas).Specific;
                Folder solapaGastos = _oForm.Items.Item(_itemSolapaGastos).Specific;
                Folder solapaTotalVentas = _oForm.Items.Item(_itemSolapaTotalVentas).Specific;

                Grid GVentas = _oForm.Items.Item(_itemGridVentas).Specific;
                Grid GGastos = _oForm.Items.Item(_itemGridGastos).Specific;
                Grid GTotales = _oForm.Items.Item(_itemGridTotales).Specific;

                itemApplyCommision.Enabled = solapaTotalVentas.Selected && GTotales.Rows.Count > 0;
                itemApplyAjuste.Enabled = (solapaVentas.Selected && GVentas.Rows.Count > 0) || (solapaGastos.Selected && GGastos.Rows.Count > 0);
            }
        }

        public void OSAPB1appl_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;

        }
       
    }
}
