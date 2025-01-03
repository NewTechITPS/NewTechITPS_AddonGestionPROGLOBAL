using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office.SpreadSheetML.Y2023.MsForms;
using DocumentFormat.OpenXml.Office2016.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using PROGLOBAL_DataGestionAjuste_addon_EA.Common;
using PROGLOBAL_DataGestionAjuste_addon_EA.Models;
using PROGLOBAL_ReservationInvoiceCloser.Services;
using REDFARM.Addons.Tools;
using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PROGLOBAL_DataGestionAjuste_addon_EA.Services
{

    public class VentanaGestionService
    {
        public static string? FormUID { get; set; }

        private static string _FormUID => FormUID ?? throw new NotImplementedException("Colocar el FORM");

        private static SAPbouiCOM.Form? _oForm;
        private static Recordset? _oRecordset;
        private static int _countForm = 0;

        private static string _itemDateFrom = "Item_0";
        private static string _itemDateTo = "Item_1";
        private static string _itemBtnFilter = "Item_4";
        private static string _itemBtnExport = "Item_5";
        private static string _itemSolapaGastos = "Item_7";
        private static string _itemGridGastos = "Item_8";
        private static string _itemGridTotales = "Item_10";
        private static string _itemSolapaVentas = "Item_11";
        private static string _itemGridVentas = "Item_12";
        private static string _itemBtnSave = "Item_14";
        private static string _itemGridSavedAjustes = "Item_15";
        private static string _itemLoading = "Item_19";

        public static void CreateMenu()
        {
            try
            {
                Menus Menus = ConnectionSDK.UIAPI!.Menus;
                Menus = Menus.Item("1536").SubMenus; // 43550  - 43546

                if (Menus.Exists("VentanaGestion"))
                {
                    Menus.Remove(Menus.Item("VentanaGestion"));
                }

                MenuItem oMenu = Menus.Add("VentanaGestion", "Informe de Gestion", BoMenuType.mt_STRING, 0);
                oMenu.Enabled = true;

            } catch (Exception ex)
            {
                NotificationService.Error($"CreateMenu Error -> {ex.Message}");
            }

        }


        public static void CreateWindow()
        {
            try
            {


                FormCreationParams p = ConnectionSDK.UIAPI!.CreateObject(BoCreatableObjectType.cot_FormCreationParams);
                p.UniqueID = "VentanaGestion_" + _countForm.ToString();

                string pathFileXml = Environment.CurrentDirectory + "\\xmlFormInformeGestion.xml";
                p.XmlData = System.IO.File.ReadAllText(pathFileXml);

                _oForm = ConnectionSDK.UIAPI!.Forms.AddEx(p);

                Folder oFolderGastos = _oForm.Items.Item(_itemSolapaGastos).Specific;
                oFolderGastos.Select();

                SAPbouiCOM.Item oItemExport = _oForm.Items.Item(_itemBtnExport);
                ButtonCombo oBtnComboExport = oItemExport.Specific;
                oBtnComboExport.ValidValues.Add("Exportar Excel", "Exportar Excel");

                _oForm.Visible = true;

                _countForm++;

            }
            finally
            {
                MarshalGC.ReleaseComObject(_oForm);
                
            }
        }

        public static void RefreshDataGastosGrid()
        {
            try { 
                _oForm = ConnectionSDK.UIAPI!.Forms.Item(_FormUID);

                string formatDate = "yyyyMMdd";
                string formatDateSP = "yyyy-MM-dd";

                SAPbouiCOM.Item item = _oForm!.Items.Item(_itemDateFrom);
                EditText ETDateFrom = item.Specific;
                string sDateFrom = ETDateFrom.Value;
                string DateFrom = DateTime.ParseExact(sDateFrom, formatDate, CultureInfo.InvariantCulture).ToString(formatDateSP);

                item = _oForm.Items.Item(_itemDateTo);
                EditText ETDateTo = item.Specific;
                string sDateTo = ETDateTo.Value;
                string DateTo = DateTime.ParseExact(sDateTo, formatDate, CultureInfo.InvariantCulture).ToString(formatDateSP);

                item = _oForm.Items.Item(_itemGridGastos);
                Grid GRIDGastos = item.Specific;

                SAPbouiCOM.DataTable oDataTableGastos;


                try
                {
                    oDataTableGastos = _oForm.DataSources.DataTables.Item("tableGastos");
                }
                catch
                {
                    oDataTableGastos = _oForm.DataSources.DataTables.Add("tableGastos");
                }

                oDataTableGastos.ExecuteQuery($"CALL INFORME_EA_GESTION('{DateFrom}','{DateTo}', 0)");

                GRIDGastos.DataTable = oDataTableGastos;

                for (int i = 0; i < GRIDGastos.Columns.Count; i++)
                {

                    GridColumn col = GRIDGastos.Columns.Item(i);

                    col.Editable = col.UniqueID switch
                    {
                        "U_Ajuste" => true,
                        _ => false,
                    };
                }

            }
            finally
            {
                MarshalGC.ReleaseComObject(_oForm);
        
            }

        }
        
        
        public static void RefreshDataVentasGrid()
        {
            try { 
                _oForm = ConnectionSDK.UIAPI!.Forms.Item(_FormUID);

                string formatDate = "yyyyMMdd";
                string formatDateSP = "yyyy-MM-dd";

                SAPbouiCOM.Item item = _oForm!.Items.Item(_itemDateFrom);
                EditText ETDateFrom = item.Specific;
                string sDateFrom = ETDateFrom.Value;
                string DateFrom = DateTime.ParseExact(sDateFrom, formatDate, CultureInfo.InvariantCulture).ToString(formatDateSP);

                item = _oForm.Items.Item(_itemDateTo);
                EditText ETDateTo = item.Specific;
                string sDateTo = ETDateTo.Value;
                string DateTo = DateTime.ParseExact(sDateTo, formatDate, CultureInfo.InvariantCulture).ToString(formatDateSP);

                item = _oForm.Items.Item(_itemGridVentas);
                Grid GRIDVentas = item.Specific;

                SAPbouiCOM.DataTable oDataTableVentas;

                try
                {
                    oDataTableVentas = _oForm.DataSources.DataTables.Item("tableVentas");
                }
                catch
                {
                    oDataTableVentas = _oForm.DataSources.DataTables.Add("tableVentas");
                }

                oDataTableVentas.ExecuteQuery($"CALL INFORME_EA_GESTION('{DateFrom}','{DateTo}', 1)");


                GRIDVentas.DataTable = oDataTableVentas;

                for (int i = 0; i < GRIDVentas.Columns.Count; i++)
                {

                    GridColumn col = GRIDVentas.Columns.Item(i);

                    col.Editable = col.UniqueID switch
                    {
                        "U_Ajuste" => true,
                        _ => false,
                    };
                }
            }
            finally
            {
                MarshalGC.ReleaseComObject(_oForm);
   
            }
}

        public static void RefreshDataTotalesGrid(ReportExcelFormatSheet sheet)
        {
            try {
                _oForm = ConnectionSDK.UIAPI!.Forms.Item(_FormUID);

                string formatDate = "yyyyMMdd";
                string formatDateSP = "yyyy-MM-dd";

                SAPbouiCOM.Item item = _oForm!.Items.Item(_itemDateFrom);
                EditText ETDateFrom = item.Specific;
                string sDateFrom = ETDateFrom.Value;
                string DateFrom = DateTime.ParseExact(sDateFrom, formatDate, CultureInfo.InvariantCulture).ToString(formatDateSP);

                item = _oForm.Items.Item(_itemDateTo);
                EditText ETDateTo = item.Specific;
                string sDateTo = ETDateTo.Value;
                string DateTo = DateTime.ParseExact(sDateTo, formatDate, CultureInfo.InvariantCulture).ToString(formatDateSP);

                item = _oForm.Items.Item(_itemGridTotales);
                Grid GRIDTotales = item.Specific;

                SAPbouiCOM.DataTable oDataTableTotales;

                try
                {
                    oDataTableTotales = _oForm.DataSources.DataTables.Item("tableTotales");
                }
                catch
                {
                    oDataTableTotales = _oForm.DataSources.DataTables.Add("tableTotales");
                }

                var dataTotals = GetDataTotals(sheet);

                List<string> columns = ["Ventas", "Costos", "Margen", "% s. ventas (1)", "Directos", "% s. ventas (2)", "Indirectos", "% s. ventas (3)"];

                oDataTableTotales.Columns.Add("Detalle", BoFieldsType.ft_AlphaNumeric);
                columns.ForEach(colName =>
                {
                    oDataTableTotales.Columns.Add(colName, BoFieldsType.ft_Float);
                });


                // INDUSTRIA
                var recordsIND = dataTotals.Where(data => Regex.IsMatch(data.CC!, "^IND"));

                var totalVentaIndustria = recordsIND.Sum(r => r.Ventas);
                var totalCostoIndustria = recordsIND.Sum(r => r.Costo);
                var totalMargenIndustria = totalVentaIndustria - totalCostoIndustria;
                var totalPorcVentaIndustria = totalVentaIndustria != 0 ? (totalMargenIndustria / totalVentaIndustria * 100) : 0;

                var totalDirectoIndustria = recordsIND.Sum(r => r.Directo);
                var totalPorcVtaDirectIndustria = totalVentaIndustria != 0 ? (totalDirectoIndustria / totalVentaIndustria * 100) : 0;
                var totalIndirectoIndustria = recordsIND.Sum(r => r.Indirecto);
                var totalPorcVtaIndirectIndustria = totalVentaIndustria != 0 ? (totalIndirectoIndustria / totalVentaIndustria * 100) : 0;

                oDataTableTotales.Rows.Add(1);
                oDataTableTotales.SetValue("Detalle", 0, "DIVISION INDUSTRIA");
                oDataTableTotales.SetValue("Ventas", 0, totalVentaIndustria);
                oDataTableTotales.SetValue("Costos", 0, totalCostoIndustria);
                oDataTableTotales.SetValue("Margen", 0, totalMargenIndustria);
                oDataTableTotales.SetValue("% s. ventas (1)", 0, totalPorcVentaIndustria);

                oDataTableTotales.SetValue("Directos", 0, totalDirectoIndustria);
                oDataTableTotales.SetValue("% s. ventas (2)", 0, totalPorcVtaDirectIndustria);
                oDataTableTotales.SetValue("Indirectos", 0, totalIndirectoIndustria);
                oDataTableTotales.SetValue("% s. ventas (3)", 0, totalPorcVtaIndirectIndustria);

                int countIndex = oDataTableTotales.Rows.Count;

                oDataTableTotales.Rows.Add(recordsIND.Count());

                
                for (int idata = 0; idata < recordsIND.Count(); idata++)
                {
                    var data = recordsIND.ToList();
                    var detalle = data[idata].CC;
                    var ventas = data[idata].Ventas;
                    var costos = data[idata].Costo;
                    var margen = ventas - costos;
                    var porcVenta = ventas != 0 ? (margen / ventas * 100) : 0;

                    oDataTableTotales.SetValue("Detalle", countIndex, detalle);
                    oDataTableTotales.SetValue("Ventas", countIndex, ventas);
                    oDataTableTotales.SetValue("Costos", countIndex, costos);
                    oDataTableTotales.SetValue("Margen", countIndex, margen);
                    oDataTableTotales.SetValue("% s. ventas (1)", countIndex, porcVenta);

                    var directo = data[idata].Directo;
                    var porcVtaDirect = ventas != 0 ? (directo / ventas * 100) : 0;
                    var indirecto = data[idata].Indirecto;
                    var porcVtaIndirect = ventas != 0 ? (indirecto / ventas * 100) : 0;

                    oDataTableTotales.SetValue("Directos", countIndex, directo);
                    oDataTableTotales.SetValue("% s. ventas (2)", countIndex, porcVtaDirect);
                    oDataTableTotales.SetValue("Indirectos", countIndex, indirecto);
                    oDataTableTotales.SetValue("% s. ventas (3)", countIndex, porcVtaIndirect);

                    countIndex++;
                }

                //// AGRO
                var recordsAGRO = dataTotals.Where(data => Regex.IsMatch(data.CC!, "^AGRO"));

                var totalVentaAGRO = recordsAGRO.Sum(r => r.Ventas);
                var totalCostoAGRO = recordsAGRO.Sum(r => r.Costo);
                var totalMargenAGRO = totalVentaAGRO - totalCostoAGRO;
                var totalPorcVentaAGRO = totalVentaAGRO != 0 ? (totalMargenAGRO / totalVentaAGRO * 100) : 0;

                var totalDirectoAGRO = recordsAGRO.Sum(r => r.Directo);
                var totalPorcVtaDirectAGRO = totalVentaAGRO != 0 ? (totalDirectoAGRO / totalVentaAGRO * 100) : 0;
                var totalIndirectoAGRO = recordsAGRO.Sum(r => r.Indirecto);
                var totalPorcVtaIndirectAGRO = totalVentaAGRO != 0 ? (totalIndirectoAGRO / totalVentaAGRO * 100) : 0;

                oDataTableTotales.Rows.Add(1);
                oDataTableTotales.SetValue("Detalle", countIndex, "DIVISION AGRO");
                oDataTableTotales.SetValue("Ventas", countIndex, totalVentaAGRO);
                oDataTableTotales.SetValue("Costos", countIndex, totalCostoAGRO);
                oDataTableTotales.SetValue("Margen", countIndex, totalMargenAGRO);
                oDataTableTotales.SetValue("% s. ventas (1)", countIndex, totalPorcVentaAGRO);

                oDataTableTotales.SetValue("Directos", countIndex, totalDirectoAGRO);
                oDataTableTotales.SetValue("% s. ventas (2)", countIndex, totalPorcVtaDirectAGRO);
                oDataTableTotales.SetValue("Indirectos", countIndex, totalIndirectoAGRO);
                oDataTableTotales.SetValue("% s. ventas (3)", countIndex, totalPorcVtaIndirectAGRO);


                countIndex = oDataTableTotales.Rows.Count;

                oDataTableTotales.Rows.Add(recordsAGRO.Count());
                for (int idata = 0; idata < recordsAGRO.Count(); idata++)
                {
                    var data = recordsAGRO.ToList();
                    var detalle = data[idata].CC;
                    var ventas = data[idata].Ventas;
                    var costos = data[idata].Costo;
                    var margen = ventas - costos;
                    var porcVenta = ventas != 0 ? (margen / ventas * 100) : 0;

                    oDataTableTotales.SetValue("Detalle", countIndex, detalle);
                    oDataTableTotales.SetValue("Ventas", countIndex, ventas);
                    oDataTableTotales.SetValue("Costos", countIndex, costos);
                    oDataTableTotales.SetValue("Margen", countIndex, margen);
                    oDataTableTotales.SetValue("% s. ventas (1)", countIndex, porcVenta);

                    var directo = data[idata].Directo;
                    var porcVtaDirect = ventas != 0 ? (directo / ventas * 100) : 0;
                    var indirecto = data[idata].Indirecto;
                    var porcVtaIndirect = ventas != 0 ? (indirecto / ventas * 100) : 0;

                    oDataTableTotales.SetValue("Directos", countIndex, directo);
                    oDataTableTotales.SetValue("% s. ventas (2)", countIndex, porcVtaDirect);
                    oDataTableTotales.SetValue("Indirectos", countIndex, indirecto);
                    oDataTableTotales.SetValue("% s. ventas (3)", countIndex, porcVtaIndirect);

                    countIndex++;
                }


                //// EQUIPO TECNICO
                var recordsET = dataTotals.Where(data => Regex.IsMatch(data.CC!, "^ET"));

                var totalVentaET = recordsET.Sum(r => r.Ventas);
                var totalCostoET = recordsET.Sum(r => r.Costo);
                var totalMargenET = totalVentaET - totalCostoET;
                var totalPorcVentaET = totalVentaET != 0 ? (totalMargenET / totalVentaET * 100) : 0;

                var totalDirectoET = recordsET.Sum(r => r.Directo);
                var totalPorcVtaDirectET = totalVentaET != 0 ? (totalDirectoET / totalVentaET * 100) : 0;
                var totalIndirectoET = recordsET.Sum(r => r.Indirecto);
                var totalPorcVtaIndirectET = totalVentaET != 0 ? (totalIndirectoET / totalVentaET * 100) : 0;

                oDataTableTotales.Rows.Add(1);
                oDataTableTotales.SetValue("Detalle", countIndex, "DIVISION EQUIPO TECNICO");
                oDataTableTotales.SetValue("Ventas", countIndex, totalVentaET);
                oDataTableTotales.SetValue("Costos", countIndex, totalCostoET);
                oDataTableTotales.SetValue("Margen", countIndex, totalMargenET);
                oDataTableTotales.SetValue("% s. ventas (1)", countIndex, totalPorcVentaET);

                oDataTableTotales.SetValue("Directos", countIndex, totalDirectoET);
                oDataTableTotales.SetValue("% s. ventas (2)", countIndex, totalPorcVtaDirectET);
                oDataTableTotales.SetValue("Indirectos", countIndex, totalIndirectoET);
                oDataTableTotales.SetValue("% s. ventas (3)", countIndex, totalPorcVtaIndirectET);


                countIndex = oDataTableTotales.Rows.Count;

                oDataTableTotales.Rows.Add(recordsET.Count());
                for (int idata = 0; idata < recordsET.Count(); idata++)
                {
                    var data = recordsET.ToList();
                    var detalle = data[idata].CC;
                    var ventas = data[idata].Ventas;
                    var costos = data[idata].Costo;
                    var margen = ventas - costos;
                    var porcVenta = ventas != 0 ? (margen / ventas * 100) : 0;

                    oDataTableTotales.SetValue("Detalle", countIndex, detalle);
                    oDataTableTotales.SetValue("Ventas", countIndex, ventas);
                    oDataTableTotales.SetValue("Costos", countIndex, costos);
                    oDataTableTotales.SetValue("Margen", countIndex, margen);
                    oDataTableTotales.SetValue("% s. ventas (1)", countIndex, porcVenta);

                    var directo = data[idata].Directo;
                    var porcVtaDirect = ventas != 0 ? (directo / ventas * 100) : 0;
                    var indirecto = data[idata].Indirecto;
                    var porcVtaIndirect = ventas != 0 ? (indirecto / ventas * 100) : 0;

                    oDataTableTotales.SetValue("Directos", countIndex, directo);
                    oDataTableTotales.SetValue("% s. ventas (2)", countIndex, porcVtaDirect);
                    oDataTableTotales.SetValue("Indirectos", countIndex, indirecto);
                    oDataTableTotales.SetValue("% s. ventas (3)", countIndex, porcVtaIndirect);

                    countIndex++;
                }


                // TOTAL COMPONENTES
                var totalComponentes = new
                {
                    Detalle = "TOTAL COMPONENTES",
                    Ventas = totalVentaIndustria + totalVentaAGRO,
                    Costos = totalCostoIndustria + totalCostoAGRO,
                    Margen = totalMargenIndustria + totalMargenAGRO,
                    PorcVenta_1 = totalPorcVentaIndustria + totalPorcVentaAGRO,
                    Directo = totalDirectoIndustria + totalDirectoAGRO,
                    PorcVenta_2 = totalPorcVtaDirectIndustria + totalPorcVtaDirectAGRO,
                    Indirecto = totalIndirectoIndustria + totalIndirectoAGRO,
                    PorcVenta_3 = totalPorcVtaIndirectIndustria + totalPorcVtaIndirectAGRO
                };

                oDataTableTotales.Rows.Add(1);
                oDataTableTotales.SetValue("Detalle", countIndex, totalComponentes.Detalle);
                oDataTableTotales.SetValue("Ventas", countIndex, totalComponentes.Ventas);
                oDataTableTotales.SetValue("Costos", countIndex, totalComponentes.Costos);
                oDataTableTotales.SetValue("Margen", countIndex, totalComponentes.Margen);
                oDataTableTotales.SetValue("% s. ventas (1)", countIndex, totalComponentes.PorcVenta_1);

                oDataTableTotales.SetValue("Directos", countIndex, totalComponentes.Directo);
                oDataTableTotales.SetValue("% s. ventas (2)", countIndex, totalComponentes.PorcVenta_2);
                oDataTableTotales.SetValue("Indirectos", countIndex, totalComponentes.Indirecto);
                oDataTableTotales.SetValue("% s. ventas (3)", countIndex, totalComponentes.PorcVenta_3);

                countIndex++;

                // TOTAL PROGLOBAL
                var totalProglobal = new
                {
                    Detalle = "TOTAL PROGLOBAL",
                    Ventas = totalComponentes.Ventas + totalVentaET,
                    Costos = totalComponentes.Costos + totalCostoET,
                    Margen = totalComponentes.Margen + totalMargenET,
                    PorcVenta_1 = totalComponentes.PorcVenta_1 + totalPorcVentaET,
                    Directo = totalComponentes.Directo + totalDirectoET,
                    PorcVenta_2 = totalComponentes.PorcVenta_2 + totalPorcVtaDirectET,
                    Indirecto = totalComponentes.Indirecto + totalIndirectoET,
                    PorcVenta_3 = totalComponentes.PorcVenta_3 + totalPorcVtaIndirectET
                };

                oDataTableTotales.Rows.Add(1);
                oDataTableTotales.SetValue("Detalle", countIndex, totalProglobal.Detalle);
                oDataTableTotales.SetValue("Ventas", countIndex, totalProglobal.Ventas);
                oDataTableTotales.SetValue("Costos", countIndex, totalProglobal.Costos);
                oDataTableTotales.SetValue("Margen", countIndex, totalProglobal.Margen);
                oDataTableTotales.SetValue("% s. ventas (1)", countIndex, totalProglobal.PorcVenta_1);

                oDataTableTotales.SetValue("Directos", countIndex, totalProglobal.Directo);
                oDataTableTotales.SetValue("% s. ventas (2)", countIndex, totalProglobal.PorcVenta_2);
                oDataTableTotales.SetValue("Indirectos", countIndex, totalProglobal.Indirecto);
                oDataTableTotales.SetValue("% s. ventas (3)", countIndex, totalProglobal.PorcVenta_3);


                GRIDTotales.DataTable = oDataTableTotales;

                for (int i = 0; i < GRIDTotales.DataTable.Rows.Count; i++)
                {
                    string detalle = GRIDTotales.DataTable.GetValue(0, i);
                    if (Regex.IsMatch(detalle, @"^TOTAL COMPONENTES"))
                    {
                        GRIDTotales.CommonSetting.SetRowBackColor(i + 1, 16776960);   
                    }
                    else if (Regex.IsMatch(detalle, @"^TOTAL PROGLOBAL"))
                    {
                        GRIDTotales.CommonSetting.SetRowBackColor(i + 1, 13808780);
                    }
                    else if (Regex.IsMatch(detalle, @"^DIVISION"))
                    {
                        GRIDTotales.CommonSetting.SetRowBackColor(i + 1, 16777138);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                MarshalGC.ReleaseComObject(_oForm);
            }
}

        public static void TruncateUDOGestionAjuste()
        {
            try { 
                _oRecordset = ConnectionSDK.DIAPI!.GetBusinessObject(BoObjectTypes.BoRecordset);
                _oRecordset!.DoQuery(@"SELECT ""Code"" FROM ""@GESTIONAJUSTE""");

                if (_oRecordset.RecordCount > 0)
                {
                    SAPbobsCOM.CompanyService? companyService = ConnectionSDK.DIAPI!.GetCompanyService();
                    SAPbobsCOM.GeneralService? generalService = companyService.GetGeneralService("GESTIONAJUSTE");

                    while (!_oRecordset.EoF)
                    {
                        string Code = _oRecordset.Fields.Item(0).Value;

                        GeneralDataParams generalDataParams = (GeneralDataParams)generalService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralDataParams);

                        generalDataParams.SetProperty("Code", Code);
                        generalService.Delete(generalDataParams);

                        _oRecordset.MoveNext();
                    }
                }
            }
            finally
            {
                MarshalGC.ReleaseComObject(_oRecordset);
      
            }

}

        public static void InsertRecordsUDOGestionAjuste()
        {
            try { 
                _oForm = ConnectionSDK.UIAPI!.Forms.Item(_FormUID);

                SAPbouiCOM.Item item = _oForm!.Items.Item(_itemGridGastos);
                DateTime dateFrom = _oForm.Items.Item(_itemDateFrom).Specific.Value;
                DateTime dateTo = _oForm.Items.Item(_itemDateTo).Specific.Value;
                Grid GRIDGastos = item.Specific;

                for (int i = 0; i < GRIDGastos.Rows.Count; i++)
                {

                    string account = GRIDGastos.DataTable.Columns.Item(5).Cells.Item(i).Value;
                    double ajuste = GRIDGastos.DataTable.Columns.Item(8).Cells.Item(i).Value;

                    if (ajuste > 0)
                    {
                        SAPbobsCOM.CompanyService? companyService = ConnectionSDK.DIAPI!.GetCompanyService();
                        SAPbobsCOM.GeneralService? generalService = companyService.GetGeneralService("GESTIONAJUSTE");
                        SAPbobsCOM.GeneralData generalData = (SAPbobsCOM.GeneralData)generalService!.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);

                        generalData.SetProperty("Code", account);
                        generalData.SetProperty("U_Ajuste", ajuste);
                        generalData.SetProperty("U_DateFrom", dateFrom);
                        generalData.SetProperty("U_DateTo", dateTo);
                        generalData.SetProperty("U_Entity", "0");

                        generalService.Add(generalData);
                    }
                }
            }
            finally
            {
                MarshalGC.ReleaseComObject(_oForm);
   
            }
        }

        public static void ResetGestionAjuste()
        {
            try { 
                _oForm = ConnectionSDK.UIAPI!.Forms.Item(_FormUID);

                EditText ETDateFrom = _oForm!.Items.Item(_itemDateFrom).Specific;
                EditText ETDateTo = _oForm!.Items.Item(_itemDateTo).Specific;

                ETDateFrom.Value = null;
                ETDateTo.Value = null;
                _oForm.DataSources.DataTables.Item("tableGastos").Clear();
                _oForm.DataSources.DataTables.Item("tableVentas").Clear();
                _oForm.DataSources.DataTables.Item("tableTotales").Clear();

            }
            finally
            {
                MarshalGC.ReleaseComObject(_oForm);
            }
        }


        public static void CreateColumnsInDataTableExpenses(System.Data.DataTable? table)
        {

            _oForm = ConnectionSDK.UIAPI!.Forms.Item(_FormUID);
            SAPbouiCOM.DataTable dtGastos = _oForm!.DataSources.DataTables.Item("tableGastos");


            DataColumnCollection cols = table!.Columns;
            for (int col = 0; col < dtGastos.Columns.Count; col++)
            {
                string colName = dtGastos.Columns.Item(col).Name;
                var colType = (dtGastos.Columns.Item(col).Cells.Item(1).Value)!.GetType();

                cols.Add(colName, colType);
            }
        }

        public static void CreateColumnsInDataTableSales(System.Data.DataTable? table)
        {
            _oForm = ConnectionSDK.UIAPI!.Forms.Item(_FormUID);
            SAPbouiCOM.DataTable dtVentas = _oForm!.DataSources.DataTables.Item("tableVentas");

            DataColumnCollection cols = table!.Columns;
            for (int col = 0; col < dtVentas.Columns.Count; col++)
            {
                string colName = dtVentas.Columns.Item(col).Name;
                var colType = (dtVentas.Columns.Item(col).Cells.Item(1).Value)!.GetType();

                cols.Add(colName, colType);
            }
        }

        public static void CreateColumnsInDataTableTotales(System.Data.DataTable? table)
        {
            _oForm = ConnectionSDK.UIAPI!.Forms.Item(_FormUID);
            SAPbouiCOM.DataTable dtTotales = _oForm!.DataSources.DataTables.Item("tableTotales");

            DataColumnCollection cols = table!.Columns;
            for (int col = 0; col < dtTotales.Columns.Count; col++)
            {
                string colName = dtTotales.Columns.Item(col).Name;
                var colType = (dtTotales.Columns.Item(col).Cells.Item(1).Value)!.GetType();

                cols.Add(colName, colType);
            }
        }


        public static void LoadDataInDataTableExpenses(ReportExcelFormatSheet sheet)
        {
            try { 
                _oForm = ConnectionSDK.UIAPI!.Forms.Item(_FormUID);
                SAPbouiCOM.DataTable dtGastos = _oForm!.DataSources.DataTables.Item("tableGastos");
                for (int r = 0; r < dtGastos.Rows.Count; r++)
                {
                    DataRow rowSys = sheet.DataTableExpenses!.NewRow();

                    for (int c = 0; c < dtGastos.Columns.Count; c++)
                    {
                        rowSys[c] = dtGastos.GetValue(c, r)?.ToString();
                    }
                    sheet.DataTableExpenses.Rows.Add(rowSys);
                }
            }
            finally
            {
                MarshalGC.ReleaseComObject(_oForm);
            }
        } 
        
        public static void LoadDataInDataTableSales(ReportExcelFormatSheet sheet)
        {
            try { 
                _oForm = ConnectionSDK.UIAPI!.Forms.Item(_FormUID);
                SAPbouiCOM.DataTable dtVentas = _oForm!.DataSources.DataTables.Item("tableVentas");
                for (int r = 0; r < dtVentas.Rows.Count; r++)
                {
                    DataRow rowSys = sheet.DataTableSales!.NewRow();

                    for (int c = 0; c < dtVentas.Columns.Count; c++)
                    {
                        rowSys[c] = dtVentas.GetValue(c, r)?.ToString();
                    }
                    sheet.DataTableSales.Rows.Add(rowSys);
                }
            }
            finally
            {
                MarshalGC.ReleaseComObject(_oForm);
            }
        }
        
        public static void LoadDataInDataTableTotals(ReportExcelFormatSheet sheet)
        {
            try {
                _oForm = ConnectionSDK.UIAPI!.Forms.Item(_FormUID);
                SAPbouiCOM.DataTable dtTotales = _oForm!.DataSources.DataTables.Item("tableTotales");
                for (int r = 0; r < dtTotales.Rows.Count; r++)
                {
                    DataRow rowSys = sheet.DataTableTotals!.NewRow();

                    for (int c = 0; c < dtTotales.Columns.Count; c++)
                    {
                        rowSys[c] = dtTotales.GetValue(c, r)?.ToString();
                    }
                    sheet.DataTableTotals.Rows.Add(rowSys);
                }
            }
            finally
            {
                MarshalGC.ReleaseComObject(_oForm);
            }
        }

     
        
        public static string GetSheetName()
        {
            try { 
                _oForm = ConnectionSDK.UIAPI!.Forms.Item(_FormUID);
                EditText ETDateFrom = _oForm!.Items.Item(_itemDateFrom).Specific;
                EditText ETDateTo = _oForm!.Items.Item(_itemDateTo).Specific;

                DateTime dDateFrom = DateTime.ParseExact(ETDateFrom.Value, "yyyyMMdd", CultureInfo.InvariantCulture);

                string[] Months = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
                string month = Months[dDateFrom.Month - 1];
                int year = dDateFrom.Year;
                return $"{month} {year}";
            }
            finally
            {
                MarshalGC.ReleaseComObject(_oForm);
            }
        }

        public static ReportExcelFormatSheet CreateSheet()
        {
            var sheet = new ReportExcelFormatSheet();
            sheet.SheetName = GetSheetName();

            sheet.DataTableExpenses = new System.Data.DataTable("SysTableGastos");
            sheet.DataTableSales = new System.Data.DataTable("SysTableVentas");
            sheet.DataTableTotals = new System.Data.DataTable("SysTableTotales");

            sheet.TitleExpenses = "GASTOS";
            sheet.TitleSales = "VENTAS";
            sheet.TitleTotals = "TOTALES";

            return sheet;
        }

        public static void LoadSheetNameInGrid(ReportExcelFormatSheet sheet)
        {
            try { 
                _oForm = ConnectionSDK.UIAPI!.Forms.Item(_FormUID);
                Grid GSavedAjuste = _oForm!.Items.Item(_itemGridSavedAjustes).Specific;
                GSavedAjuste.DataTable.Rows.Add();
                GSavedAjuste.DataTable.SetValue(0, GSavedAjuste.DataTable.Rows.Count - 1, sheet.SheetName);
            }
            finally
            {
                MarshalGC.ReleaseComObject(_oForm);
            } 
        }

        public static string GetPathToSaveFile(string? filename)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            saveFileDialog.Filter = "Archivos Excel (*.xlsx)|*.xlsx|Todos los archivos (*.*)|*.*";
            saveFileDialog.Title = "Guardar archivo como";
            saveFileDialog.FileName = filename;

            string pathFilename = string.Empty;

            var thread = new Thread(() =>
            {
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    pathFilename = saveFileDialog.FileName;
                }
            });

            thread.ApartmentState = ApartmentState.STA;
            thread.Start();
            thread.Join();

            return pathFilename;
        }

        public static void ExportExcel(ReportExcelFormat reportExcelFormat, string pathFile)
        {
            var workbook = new XLWorkbook();

            foreach (var sheet in reportExcelFormat!.Sheets)
            {
                var sheetExcel = workbook.Worksheets.Add(sheet.SheetName!);

                sheetExcel.Cell("A1").SetValue(sheet.TitleExpenses);
                sheetExcel.Cell("A1").Style.Font.Bold = true;
                sheetExcel.Cell("A3").InsertTable(sheet.DataTableExpenses);

                IXLCell LastCell1 = sheetExcel.Column(1).LastCellUsed();
                int rowLastDataExpenses = LastCell1.Address.RowNumber;

                sheetExcel.Cell(rowLastDataExpenses + 2, 1).SetValue(sheet.TitleSales);
                sheetExcel.Cell(rowLastDataExpenses + 2, 1).Style.Font.Bold = true;
                sheetExcel.Cell(rowLastDataExpenses + 4, 1).InsertTable(sheet.DataTableSales);

                IXLCell LastCell2 = sheetExcel.Column(1).LastCellUsed();
                int rowLastDataSales = LastCell2.Address.RowNumber;

                sheetExcel.Cell(rowLastDataSales + 2, 1).SetValue(sheet.TitleTotals);
                sheetExcel.Cell(rowLastDataSales + 2, 1).Style.Font.Bold = true;
                sheetExcel.Cell(rowLastDataSales + 4, 1).InsertTable(sheet.DataTableTotals);

                sheetExcel.Columns().AdjustToContents();
            }

            workbook.SaveAs(pathFile);
        }

        public static List<VentasGastoTotales> GetDataTotals(ReportExcelFormatSheet sheet)
        {
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
                         select new VentasGastoTotales()
                         {
                             CC = ventas.Code,
                             Ventas = ventas.Ventas,
                             Costo = ventas.Costos,
                             Directo = gastos.TotalDirecto,
                             Indirecto = gastos.TotalIndirecto,
                             TotalGasto = gastos.TotalCC
                         };


            return totals.ToList();
        }
    }

    public class VentasGastoTotales
    {
        public string? CC { get; set; }
        public double? Ventas { get; set; }
        public double? Costo { get; set; }
        public double? Directo { get; set; }
        public double? Indirecto { get; set; }
        public double? TotalGasto { get; set; }
    }
}
