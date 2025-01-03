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

namespace PROGLOBAL_DataGestionAjuste_addon_EA.Forms.WINDOW
{

    public class VentanaGestion 
    {

        #region Atributos

        private static SAPbouiCOM.Form? _oForm;

        private static ReportExcelFormat? _reportExcelFormat;

        public const string frmUID = "60006"; 
        public const string menuUID = "VentanaGestion";

        private string _itemDateFrom = "Item_0";
        private string _itemDateTo = "Item_1";
        private string _itemBtnFilter = "Item_4";
        private string _itemBtnExport = "Item_5";
        private string _itemSolapaGastos = "Item_7";
        private string _itemGridGastos = "Item_8";
        private string _itemGridTotales = "Item_10";
        private string _itemSolapaVentas = "Item_11";
        private string _itemGridVentas = "Item_12";
        private string _itemBtnApplyAjuste = "Item_13";
        private string _itemBtnSave = "Item_14";
        private string _itemGridSavedAjustes = "Item_15";
        private string _itemLoading = "Item_19";
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
                if (!string.IsNullOrEmpty(valueDateFrom) && !string.IsNullOrEmpty(valueDateTo))
                {
                    oItem.Enabled = true;
                }
                else
                {
                    oItem.Enabled = false;
                }
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

                    VentanaGestionService.TruncateUDOGestionAjuste();

                    VentanaGestionService.RefreshDataGastosGrid();
                    VentanaGestionService.CreateColumnsInDataTableExpenses(sheet.DataTableExpenses);
                    VentanaGestionService.LoadDataInDataTableExpenses(sheet);

                    VentanaGestionService.RefreshDataVentasGrid();
                    VentanaGestionService.CreateColumnsInDataTableSales(sheet.DataTableSales);  
                    VentanaGestionService.LoadDataInDataTableSales(sheet);

                    VentanaGestionService.RefreshDataTotalesGrid(sheet);
                    //VentanaGestionService.CreateColumnsInDataTableTotales(sheet.DataTableTotals); 
                    //VentanaGestionService.LoadDataInDataTableTotals(sheet);

                    // TODO GASTOS

                    //DataRow rowSys = sheet.DataTableTotals!.NewRow();
                    //sheet.DataTableTotals.Rows.Add(rowSys);

                    //// INDUSTRIA
                    //rowSys["PROGLOBAL"] = "DIVISION INDUSTRIA";
                    //var recordsIND = totals.Where(data => Regex.IsMatch(data.CC, "^IND"));
                    //rowSys["Ventas"] = recordsIND.Sum(i => i.Ventas);
                    //rowSys["Costos"] = recordsIND.Sum(i => i.Costo);


                    //foreach (var cc in recordsIND)
                    //{
                    //    rowSys = sheet.DataTableTotals!.NewRow();
                    //    rowSys["PROGLOBAL"] = cc.CC;
                    //    rowSys["Ventas"] = cc.Ventas;
                    //    rowSys["Costos"] = cc.Costo;
                    //    rowSys["Margen"] = cc.Ventas - cc.Costo;
                    //    rowSys["% s. ventas (1)"] = 0;

                    //    rowSys["Directos"] = cc.Directo;
                    //    rowSys["% s. ventas (2)"] = 0;
                    //    rowSys["Indirectos"] = cc.Indirecto;
                    //    rowSys["% s. ventas (3)"] = 0;

                    //    sheet.DataTableTotals.Rows.Add(rowSys);
                    //}


                    //// AGRO
                    //rowSys = sheet.DataTableTotals!.NewRow();
                    //sheet.DataTableTotals.Rows.Add(rowSys);
                    //rowSys["PROGLOBAL"] = "DIVISION AGRO";

                    //var recordsAGRO = totals.Where(data => Regex.IsMatch(data.CC, "^AGRO"));
                    //foreach (var cc in recordsAGRO)
                    //{
                    //    rowSys = sheet.DataTableTotals!.NewRow();
                    //    rowSys["PROGLOBAL"] = cc.CC;
                    //    rowSys["Ventas"] = cc.Ventas;
                    //    rowSys["Costos"] = cc.Costo;
                    //    rowSys["Margen"] = 0;
                    //    rowSys["% s. ventas (1)"] = 0;

                    //    rowSys["Directos"] = cc.Directo;
                    //    rowSys["% s. ventas (2)"] = 0;
                    //    rowSys["Indirectos"] = cc.Indirecto;
                    //    rowSys["% s. ventas (3)"] = 0;
                    //    sheet.DataTableTotals.Rows.Add(rowSys);
                    //}


                    //// EQUIPO TECNICO
                    //rowSys = sheet.DataTableTotals!.NewRow();
                    //sheet.DataTableTotals.Rows.Add(rowSys);
                    //rowSys["PROGLOBAL"] = "DIVISION EQUIPO TECNICO";

                    //var recordsET = totals.Where(data => Regex.IsMatch(data.CC, "^ET"));
                    //foreach (var cc in recordsET)
                    //{
                    //    rowSys = sheet.DataTableTotals!.NewRow();
                    //    rowSys["PROGLOBAL"] = cc.CC;
                    //    rowSys["Ventas"] = cc.Ventas;
                    //    rowSys["Costos"] = cc.Costo;
                    //    rowSys["Margen"] = 0;
                    //    rowSys["% s. ventas (1)"] = 0;

                    //    rowSys["Directos"] = cc.Directo;
                    //    rowSys["% s. ventas (2)"] = 0;
                    //    rowSys["Indirectos"] = cc.Indirecto;
                    //    rowSys["% s. ventas (3)"] = 0;
                    //    sheet.DataTableTotals.Rows.Add(rowSys);
                    //}


                    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////
                    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////
                    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////

                    //VentanaGestionService.LoadDataInDataTableTotals(sheet);

                    //VentanaGestionService.ResetGestionAjuste();
                    //VentanaGestionService.LoadSheetNameInGrid(sheet);

                    //
                    //////////////////

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
                    VentanaGestionService.TruncateUDOGestionAjuste();

                    VentanaGestionService.InsertRecordsUDOGestionAjuste();

                    VentanaGestionService.RefreshDataGastosGrid();
                    VentanaGestionService.RefreshDataVentasGrid();
                    //VentanaGestionService.RefreshDataTotalesGrid();

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
                    VentanaGestionService.CreateColumnsInDataTableSales(sheet.DataTableSales);  // TO DO
                    //VentanaGestionService.CreateColumnsInDataTableTotales(sheet.DataTableTotals); // TO DO


                    VentanaGestionService.LoadDataInDataTableExpenses(sheet);
                    VentanaGestionService.LoadDataInDataTableSales(sheet);


                    //////////////////////////////////////////////////////
                    // TODO GASTOS
                    var expensesData = sheet.DataTableExpenses!.AsEnumerable();
                    var colsGastos = sheet.DataTableExpenses!.Columns.Cast<System.Data.DataColumn>().Where(col => Regex.IsMatch(col.ColumnName, @"Di$|In$")).ToList();

                    var dataGastos = colsGastos
                        .Select(col =>
                        {
                            string uniqueCC = Regex.Replace(col.ColumnName, @"Di$|In$", "").Trim();

                            double totalDirecto = expensesData
                                                    .Where(data => !data.IsNull(uniqueCC + " Di"))
                                                    .Select(data => data.Field<double>(uniqueCC + " Di")).Sum();

                            double totalIndirecto = expensesData
                                                    .Where(data => !data.IsNull(uniqueCC + " In"))
                                                    .Select(data => data.Field<double>(uniqueCC + " In")).Sum();

                            return new
                            {
                                Code = uniqueCC,
                                TotalDirecto = totalDirecto,
                                TotalIndirecto = totalIndirecto,
                                TotalCC = totalDirecto + totalIndirecto
                            };
                        }).Distinct();


                    //////////////////////////////////////////////////////
                    // TODO VENTAS
                    var SalesData = sheet.DataTableSales!.AsEnumerable();
                    var colsVentas = sheet.DataTableSales!.Columns.Cast<System.Data.DataColumn>().Where(col => Regex.IsMatch(col.ColumnName, @"^Costo\s+(IND|AGRO|ET)|^Venta\s+(IND|AGRO|ET)")).ToList();
                    var dataVentasSinAgrupar = colsVentas
                        .Select(col =>
                        {
                            string uniqueCC = Regex.Replace(col.ColumnName, @"^Costo|^Venta", "").Trim();

                            double totalVentas = SalesData
                                                    .Where(data => !data.IsNull(col.ColumnName))
                                                    .Select(data => Regex.IsMatch(col.ColumnName, @"^Venta\s+(IND|AGRO|ET)") ? data.Field<double>(col.ColumnName) : 0).Sum();
                            double totalCosto = SalesData
                                                    .Where(data => !data.IsNull(col.ColumnName))
                                                    .Select(data => Regex.IsMatch(col.ColumnName, @"^Costo\s+(IND|AGRO|ET)") ? data.Field<double>(col.ColumnName) : 0).Sum();

                            return new
                            {
                                Code = uniqueCC,
                                Ventas = totalVentas,
                                Costo = totalCosto,
                            };

                        }).Distinct();

                    var dataVentasAgrupado = dataVentasSinAgrupar
                        .GroupBy(d => d.Code)
                        .Select(d =>
                        {
                            return new
                            {
                                Code = d.Key,
                                Ventas = d.Sum(data => data.Ventas),
                                Costos = d.Sum(data => data.Costo)
                            };
                        }).ToList();

                    // TOTALES
                    var totals = from ventas in dataVentasAgrupado
                                 join gastos in dataGastos
                                 on ventas.Code equals gastos.Code
                                 select new
                                 {
                                     CC = ventas.Code,
                                     Ventas = ventas.Ventas,
                                     Costo = ventas.Costos,
                                     Directo = gastos.TotalDirecto,
                                     Indirecto = gastos.TotalIndirecto,
                                     TotalGasto = gastos.TotalCC
                                 };

                    // TOTALES
                    sheet.DataTableTotals!.Columns.Add("PROGLOBAL");

                    // VENTAS
                    sheet.DataTableTotals!.Columns.Add("Ventas");
                    sheet.DataTableTotals!.Columns.Add("Costos");
                    sheet.DataTableTotals!.Columns.Add("Margen");
                    sheet.DataTableTotals!.Columns.Add("% s. ventas (1)");

                    // GASTOS
                    sheet.DataTableTotals!.Columns.Add("Directos"); 
                    sheet.DataTableTotals!.Columns.Add("% s. ventas (2)");
                    sheet.DataTableTotals!.Columns.Add("Indirectos");
                    sheet.DataTableTotals!.Columns.Add("% s. ventas (3)");

                    //// RESULTADOS
                    //sheet.DataTableTotals!.Columns.Add("Mensual");
                    //sheet.DataTableTotals!.Columns.Add("% s. ventas (4)");
                    //sheet.DataTableTotals!.Columns.Add("Comisiones");
                    //sheet.DataTableTotals!.Columns.Add("Acumulados");
                    //sheet.DataTableTotals!.Columns.Add("% s. ventas (5)");

                    //// INTERESES
                    //sheet.DataTableTotals!.Columns.Add("Intereses");
                    //sheet.DataTableTotals!.Columns.Add("Acumulados_");
                    //sheet.DataTableTotals!.Columns.Add("% s. ventas (6)");



                    DataRow rowSys = sheet.DataTableTotals!.NewRow();
                    sheet.DataTableTotals.Rows.Add(rowSys);

                    // INDUSTRIA
                    rowSys["PROGLOBAL"] = "DIVISION INDUSTRIA";
                    var recordsIND = totals.Where(data => Regex.IsMatch(data.CC,"^IND"));
                    rowSys["Ventas"] = recordsIND.Sum(i => i.Ventas);
                    rowSys["Costos"] = recordsIND.Sum(i => i.Costo);


                    foreach (var cc in recordsIND)
                    {
                        rowSys = sheet.DataTableTotals!.NewRow();
                        rowSys["PROGLOBAL"] = cc.CC;
                        rowSys["Ventas"] = cc.Ventas;
                        rowSys["Costos"] = cc.Costo;
                        rowSys["Margen"] = cc.Ventas - cc.Costo;
                        rowSys["% s. ventas (1)"] = 0;

                        rowSys["Directos"] = cc.Directo;
                        rowSys["% s. ventas (2)"] = 0;
                        rowSys["Indirectos"] = cc.Indirecto;
                        rowSys["% s. ventas (3)"] = 0;

                        sheet.DataTableTotals.Rows.Add(rowSys);
                    }


                    // AGRO
                    rowSys = sheet.DataTableTotals!.NewRow();
                    sheet.DataTableTotals.Rows.Add(rowSys);
                    rowSys["PROGLOBAL"] = "DIVISION AGRO";

                    var recordsAGRO = totals.Where(data => Regex.IsMatch(data.CC, "^AGRO"));
                    foreach (var cc in recordsAGRO)
                    {
                        rowSys = sheet.DataTableTotals!.NewRow();
                        rowSys["PROGLOBAL"] = cc.CC;
                        rowSys["Ventas"] = cc.Ventas;
                        rowSys["Costos"] = cc.Costo;
                        rowSys["Margen"] = 0;
                        rowSys["% s. ventas (1)"] = 0;

                        rowSys["Directos"] = cc.Directo;
                        rowSys["% s. ventas (2)"] = 0;
                        rowSys["Indirectos"] = cc.Indirecto;
                        rowSys["% s. ventas (3)"] = 0;
                        sheet.DataTableTotals.Rows.Add(rowSys);
                    }


                    // EQUIPO TECNICO
                    rowSys = sheet.DataTableTotals!.NewRow();
                    sheet.DataTableTotals.Rows.Add(rowSys);
                    rowSys["PROGLOBAL"] = "DIVISION EQUIPO TECNICO";

                    var recordsET = totals.Where(data => Regex.IsMatch(data.CC, "^ET"));
                    foreach (var cc in recordsET)
                    {
                        rowSys = sheet.DataTableTotals!.NewRow();
                        rowSys["PROGLOBAL"] = cc.CC;
                        rowSys["Ventas"] = cc.Ventas;
                        rowSys["Costos"] = cc.Costo;
                        rowSys["Margen"] = 0;
                        rowSys["% s. ventas (1)"] = 0;

                        rowSys["Directos"] = cc.Directo;
                        rowSys["% s. ventas (2)"] = 0;
                        rowSys["Indirectos"] = cc.Indirecto;
                        rowSys["% s. ventas (3)"] = 0;
                        sheet.DataTableTotals.Rows.Add(rowSys);
                    }


                    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////
                    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////
                    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////

                    //VentanaGestionService.LoadDataInDataTableTotals(sheet);

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

           
        }

        public void OSAPB1appl_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;

        }
       
    }
}
