﻿using SAPbouiCOM;
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
using Newtonsoft.Json;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Drawing.Charts;
using static System.Runtime.InteropServices.JavaScript.JSType;
using DocumentFormat.OpenXml.Office2016.Excel;

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
                //VentanaGestionService.TruncateUDOGestionAjuste();
            }

            // PRESIONAR BOTON "FILTRAR"
            if (pVal.EventType == BoEventTypes.et_ITEM_PRESSED && pVal.ItemUID == _itemBtnFilter && pVal.ActionSuccess) 
            {
                SAPbouiCOM.ProgressBar progressBar = ConnectionSDK.UIAPI!.StatusBar.CreateProgressBar("Filtrando datos", 100, true);

                _oForm = ConnectionSDK.UIAPI!.Forms.Item(FormUID);
                SAPbouiCOM.StaticText oInfoProgress = _oForm.Items.Item(VentanaGestionService._itemInfoProgress).Specific;

                try
                {
                    var sheet = VentanaGestionService.CreateSheet();

                    // GASTOS
                    oInfoProgress.Caption = "Obteniendo los datos de Gastos..";
                    VentanaGestionService.RefreshDataGastosGrid();
                    progressBar.Value = 10;
                    VentanaGestionService.CreateColumnsInDataTableSystem(NameDataTables.tablaGastos.ToString(), sheet.DataTableExpenses);
                    progressBar.Value = 20;
                    VentanaGestionService.LoadDataInDataTableSystem(NameDataTables.tablaGastos.ToString(), sheet.DataTableExpenses);
                    progressBar.Value = 40;

                    // VENTAS
                    oInfoProgress.Caption = "Obteniendo los datos de Ventas..";
                    VentanaGestionService.RefreshDataVentasGrid();
                    progressBar.Value = 50;
                    VentanaGestionService.CreateColumnsInDataTableSystem(NameDataTables.tablaVentas.ToString(), sheet.DataTableSales);
                    progressBar.Value = 60;
                    VentanaGestionService.LoadDataInDataTableSystem(NameDataTables.tablaVentas.ToString(), sheet.DataTableSales);
                    progressBar.Value = 80;


                    // TOTALES
                    oInfoProgress.Caption = "Obteniendo los Totales de Ventas..";
                    var totals = VentanaGestionService.RefreshDataTotalesVentasGrid(sheet);
                    progressBar.Value = 90;
                    oInfoProgress.Caption = "Obteniendo los Totales de Gastos..";
                    VentanaGestionService.RefreshDataTotalesGastosGrid(totals!);
                    progressBar.Value = 100;

                    SAPbouiCOM.Item oItemBtnAjuste = _oForm.Items.Item(_itemBtnApplyAjuste);
                    SAPbouiCOM.Item oItemBtnSave = _oForm.Items.Item(_itemBtnSave);
                    oItemBtnAjuste.Enabled = true;
                    oItemBtnSave.Enabled = true;

                } catch (Exception ex)
                {
                    NotificationService.Error("Error al aplicar el filtro de fecha; Mensaje ->" + ex.Message);
                } finally
                {
                    oInfoProgress.Caption = "";
                    progressBar.Stop();
                    MarshalGC.ReleaseComObject(progressBar);
                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                }
            }

            // PRESIONAR BOTON "APLICAR AJUSTE"
            if(pVal.EventType == BoEventTypes.et_ITEM_PRESSED && pVal.ItemUID == _itemBtnApplyAjuste && pVal.ActionSuccess) 
            {
                SAPbouiCOM.ProgressBar progressBar = ConnectionSDK.UIAPI!.StatusBar.CreateProgressBar("Aplicando Ajustes", 100, false);
                _oForm = ConnectionSDK.UIAPI!.Forms.Item(FormUID);
                _oForm = ConnectionSDK.UIAPI!.Forms.Item(FormUID);
                SAPbouiCOM.StaticText oInfoProgress = _oForm.Items.Item(VentanaGestionService._itemInfoProgress).Specific;

                try
                {
                    var sheet = VentanaGestionService.CreateSheet();

                    sheet.DataTableExpenses = new System.Data.DataTable();
                    sheet.DataTableSales = new System.Data.DataTable();
                    sheet.DataTableTotalsSales = new System.Data.DataTable();
                    sheet.DataTableTotalsExpenses = new System.Data.DataTable();

                    oInfoProgress.Caption = "Guardando los Ajustes colocados..";
                    VentanaGestionService.InsertRecordsUDOGestionAjuste();
                    progressBar.Value += 10;
                    // GASTOS
                    oInfoProgress.Caption = "Actualizando información de los Gastos..";
                    VentanaGestionService.RefreshDataGastosGrid();
                    progressBar.Value += 10;
                    VentanaGestionService.CreateColumnsInDataTableSystem(NameDataTables.tablaGastos.ToString(), sheet.DataTableExpenses);
                    progressBar.Value += 10;
                    VentanaGestionService.LoadDataInDataTableSystem(NameDataTables.tablaGastos.ToString(), sheet.DataTableExpenses);
                    progressBar.Value += 10;

                    // VENTAS
                    oInfoProgress.Caption = "Actualizando información de las Ventas..";
                    VentanaGestionService.RefreshDataVentasGrid();
                    progressBar.Value += 10;
                    VentanaGestionService.CreateColumnsInDataTableSystem(NameDataTables.tablaVentas.ToString(), sheet.DataTableSales);
                    progressBar.Value += 10;
                    VentanaGestionService.LoadDataInDataTableSystem(NameDataTables.tablaVentas.ToString(), sheet.DataTableSales);
                    progressBar.Value += 10;

                    // TOTALES
                    oInfoProgress.Caption = "Actualizando información de los totales de Ventas..";
                    var totals = VentanaGestionService.RefreshDataTotalesVentasGrid(sheet);
                    progressBar.Value += 10;

                    oInfoProgress.Caption = "Actualizando información de los totales de Gastos..";
                    VentanaGestionService.RefreshDataTotalesGastosGrid(totals!);
                    progressBar.Value += 20;
                    ConnectionSDK.UIAPI!.MessageBox("Ajuste aplicado con éxito");
                    
                } catch (Exception ex)
                {

                    NotificationService.Error("Error al aplicar ajuste; Mensaje ->" + ex.Message);

                } finally
                {
                    oInfoProgress.Caption = "";
                    MarshalGC.ReleaseComObject(progressBar);
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
                
            }

            // PRESIONAR BOTON "GUARDAR"
            if(pVal.EventType == BoEventTypes.et_ITEM_PRESSED && pVal.ItemUID == _itemBtnSave && pVal.ActionSuccess)  
            {
                SAPbouiCOM.ProgressBar progressBar = ConnectionSDK.UIAPI!.StatusBar.CreateProgressBar("Guardando ajustes", 100, true);
                _oForm = ConnectionSDK.UIAPI!.Forms.Item(FormUID);
                SAPbouiCOM.StaticText oInfoProgress = _oForm.Items.Item(VentanaGestionService._itemInfoProgress).Specific;
                
                try
                {
                    SAPbouiCOM.EditText ETDateFrom = _oForm!.Items.Item(_itemDateFrom).Specific;
                    _reportExcelFormat!.FirstDate = ETDateFrom.Value;

                    var sheet = VentanaGestionService.CreateSheet();

                    oInfoProgress.Caption = "Guardando los datos de Gastos..";

                    VentanaGestionService.CreateColumnsInDataTableSystem(NameDataTables.tablaGastos.ToString(), sheet.DataTableExpenses);
                    VentanaGestionService.LoadDataInDataTableSystem(NameDataTables.tablaGastos.ToString(), sheet.DataTableExpenses);
                    progressBar.Value = 20;


                    oInfoProgress.Caption = "Guardando los datos de Ventas..";
                    VentanaGestionService.CreateColumnsInDataTableSystem(NameDataTables.tablaVentas.ToString(), sheet.DataTableSales);
                    VentanaGestionService.LoadDataInDataTableSystem(NameDataTables.tablaVentas.ToString(), sheet.DataTableSales);
                    progressBar.Value = 40;


                    oInfoProgress.Caption = "Guardando los Totales de Ventas..";
                    VentanaGestionService.CreateColumnsInDataTableSystem(NameDataTables.tablaTotalVentas.ToString(), sheet.DataTableTotalsSales);
                    VentanaGestionService.LoadDataInDataTableSystem(NameDataTables.tablaTotalVentas.ToString(), sheet.DataTableTotalsSales);
                    progressBar.Value = 60;


                    oInfoProgress.Caption = "Guardando los Totales de Gastos..";
                    VentanaGestionService.CreateColumnsInDataTableSystem(NameDataTables.tablaTotalGastos.ToString(), sheet.DataTableTotalsExpenses);
                    VentanaGestionService.LoadDataInDataTableSystem(NameDataTables.tablaTotalGastos.ToString(), sheet.DataTableTotalsExpenses);
                    progressBar.Value = 80;


                    VentanaGestionService.ResetGestionAjuste();
                    VentanaGestionService.LoadSheetNameInGrid(sheet);
                    progressBar.Value = 100;

                    _reportExcelFormat!.Sheets.Add(sheet);

                    ConnectionSDK.UIAPI!.MessageBox("Ajuste guardado con éxito -> Ajuste: " + sheet.SheetName);

                    SAPbouiCOM.Item oItemBtnAjuste = _oForm.Items.Item(_itemBtnApplyAjuste);
                    SAPbouiCOM.Item oItemBtnSave = _oForm.Items.Item(_itemBtnSave);
                    oItemBtnAjuste.Enabled = false;
                    oItemBtnSave.Enabled = false;

                    ButtonCombo oBCExport = (ButtonCombo)_oForm!.Items.Item(_itemBtnExport).Specific;
                    Grid GSavedAjuste = _oForm!.Items.Item(_itemGridSavedAjustes).Specific;

                    if (GSavedAjuste.Rows.Count == 3)
                    {
                        oBCExport.Item.Enabled = true;
                    }


                } catch(Exception ex)
                {
                    NotificationService.Error("Error al guardar; Mensaje ->" + ex.Message);
                }
                finally
                {
                    oInfoProgress.Caption = "";
                    progressBar.Stop();
                    MarshalGC.ReleaseComObject(progressBar);
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }

            }

            // PRESIONAR BOTON "EXPORTAR EXCEL"
            if (pVal.EventType == BoEventTypes.et_COMBO_SELECT && pVal.ItemUID == _itemBtnExport && pVal.ActionSuccess)
            {
                SAPbouiCOM.ProgressBar progressBar = ConnectionSDK.UIAPI!.StatusBar.CreateProgressBar("Comenzando proceso de exportación", 100, false);

                _oForm = ConnectionSDK.UIAPI!.Forms.Item(FormUID);
                SAPbouiCOM.StaticText oInfoProgress = _oForm.Items.Item(VentanaGestionService._itemInfoProgress).Specific;
                ButtonCombo oBCExport = (ButtonCombo)_oForm!.Items.Item(_itemBtnExport).Specific;

                try {                

                    progressBar.Value = 10;
                    oInfoProgress.Caption = "Calculando información del trimestre actual..";

                    var sCurrTrimestral = VentanaGestionService.CalcCurrentTrimestral(_reportExcelFormat!);
                    progressBar.Value = 50;

                    _reportExcelFormat!.Sheets.Add(sCurrTrimestral!);

                    var monthCalcAnnual = VentanaGestionService.GetMonthCalcAnnual(_reportExcelFormat);

                    if(monthCalcAnnual != null)
                    {
                        oInfoProgress.Caption = "Calculando información anual..";

                        var sheetsAnnual = VentanaGestionService.GetSheetsAnnual(monthCalcAnnual, _oForm);

                        ReportExcelFormatSheet sAnnual = new();
                        VentanaGestionService.CloneDataTableSheetAnnualFrom(sAnnual, sCurrTrimestral!);
                        progressBar.Value = 60;

                        // GASTOS
                        oInfoProgress.Caption = "Calculando información anual.. Obteniendo informacion de Gastos.";
                        VentanaGestionService.LoadDataExpensesInSheetAnnual(sCurrTrimestral!, sAnnual, sheetsAnnual!);
                        progressBar.Value = 70;

                        // VENTAS
                        oInfoProgress.Caption = "Calculando información anual.. Obteniendo informacion de Ventas.";
                        VentanaGestionService.LoadDataSalesInSheetAnnual(sCurrTrimestral!, sAnnual, sheetsAnnual!);
                        progressBar.Value = 80;

                        // TOTALES DE VENTAS
                        oInfoProgress.Caption = "Calculando información anual.. Obteniendo informacion de los Totales de Ventas.";
                        VentanaGestionService.LoadDataTotalsSalesInSheetAnnual(sCurrTrimestral!, sAnnual, sheetsAnnual!);
                        progressBar.Value = 90;

                        // TOTALES DE GASTOS
                        oInfoProgress.Caption = "Calculando información anual.. Obteniendo informacion de los Totales de Gastos.";
                        VentanaGestionService.LoadDataTotalsExpensesInSheetAnnual(sCurrTrimestral!, sAnnual, sheetsAnnual!);
                        progressBar.Value = 100;

                        // --------------- // 
                        sAnnual.SheetName = "ANUAL";
                        sAnnual.TitleSales = "VENTAS";
                        sAnnual.TitleExpenses = "GASTOS";
                        sAnnual.TitleTotalsExpenses = "TOTALES GASTOS";
                        sAnnual.TitleTotalsSales = "TOTALES VENTAS";

                        _reportExcelFormat!.Sheets.Add(sAnnual!);
                    }

                    if (oBCExport.Selected.Value == "Exportar Excel")
                    {
                        string pathFile = VentanaGestionService.GetPathToSaveFile(_reportExcelFormat!.FileName);

                        if (pathFile != null)
                        {
                            VentanaGestionService.ExportExcel(_reportExcelFormat, pathFile);

                            NotificationService.Success("Documento creado con exito");
                        }

                        ConnectionSDK.UIAPI.MessageBox("Informe finalizado");

                        try
                        {
                            _oForm.Close();
                        }catch {}

                    }

                }
                catch (Exception ex)
                {
                    NotificationService.Error("Error al guardar; Mensaje ->" + ex.Message);
                }
                finally
                {
                    progressBar.Stop();
                    MarshalGC.ReleaseComObject(progressBar);
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }

            }

            // CALCULAR ACUMULADO Y SU PORCENTAJE
            if (pVal.EventType == BoEventTypes.et_VALIDATE && pVal.ColUID == _colComision && pVal.ItemUID == _itemGridTotales && pVal.ActionSuccess)
            {
                _oForm = ConnectionSDK.UIAPI!.Forms.Item(pVal.FormUID);
                _oForm.Freeze(true);

                try
                {
                    Grid GTotales = _oForm.Items.Item(_itemGridTotales).Specific;

                    double mensual = GTotales.DataTable.GetValue(_colMensual, pVal.Row);
                    double comision = GTotales.DataTable.GetValue(_colComision, pVal.Row);
                    double ventas = GTotales.DataTable.GetValue(_colVentas, pVal.Row);

                    double acumulado = mensual + comision;
                    GTotales.DataTable.Columns.Item(_colAcumulado).Cells.Item(pVal.Row).Value = acumulado;
                    GTotales.DataTable.Columns.Item(_colPorcAcum).Cells.Item(pVal.Row).Value = ventas != 0 ? acumulado / ventas * 100 : 0;
                } catch(Exception ex)
                {
                    NotificationService.Error(ex.Message);
                } finally
                {
                    _oForm.Freeze(false);
                }
            }


            // APLICAR COMISIONES
            if (pVal.EventType == BoEventTypes.et_ITEM_PRESSED && pVal.ItemUID == _itemBtnApplyCommision && pVal.ActionSuccess)
            {
                SAPbouiCOM.ProgressBar progressBar = ConnectionSDK.UIAPI!.StatusBar.CreateProgressBar("Aplicando comisiones", 100, false);
                progressBar.Value += 50;
                VentanaGestionService.CalculateTotals_Comisiones_Acumulado_PorcAcumulado();
                progressBar.Value += 50;
                
            }

            // BLOQUEAR Y DESBLOQUEAR BOTONES "APLICAR COMISIONES" Y "APLICAR AJUSTE"
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

            // CALCULAR TRIMESTRE AUTOMATICAMENTE
            if (pVal.EventType == BoEventTypes.et_VALIDATE && pVal.ItemUID == _itemDateFrom && pVal.ActionSuccess)
            {
                _oForm = ConnectionSDK.UIAPI!.Forms.Item(pVal.FormUID);
                SAPbouiCOM.StaticText lblNumTrimestral = _oForm.Items.Item(VentanaGestionService._itemLblNumTrimestral).Specific;
                SAPbouiCOM.EditText dateFrom = _oForm.Items.Item(_itemDateFrom).Specific;

                if(!string.IsNullOrEmpty(dateFrom.Value))
                {
                    int year = Convert.ToInt32(dateFrom.Value.Substring(0, 4)); // 20250101  -> 2025
                    int month = Convert.ToInt32(dateFrom.Value.Substring(4, 2)); // 20250101  -> 01 -> 1

                    int[] trimestral1 = [10, 11, 12];
                    int[] trimestral2 = [1, 2, 3];
                    int[] trimestral3 = [4, 5, 6];
                    int[] trimestral4 = [7, 8, 9];

                    bool isTrimestral1 = trimestral1.Any(m => m == month);
                    bool isTrimestral2 = trimestral2.Any(m => m == month);
                    bool isTrimestral3 = trimestral3.Any(m => m == month);
                    bool isTrimestral4 = trimestral4.Any(m => m == month);

                    int? currentSearchTrimestral =
                        isTrimestral1 ? 1 :
                        isTrimestral2 ? 2 :
                        isTrimestral3 ? 3 :
                        isTrimestral4 ? 4 : null;

                    lblNumTrimestral.Caption = Convert.ToString(currentSearchTrimestral);
                }
            }
        }

        public void OSAPB1appl_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;

        }
       
    }
}
