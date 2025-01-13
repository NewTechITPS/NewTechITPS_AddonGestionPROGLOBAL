using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office.SpreadSheetML.Y2023.MsForms;
using DocumentFormat.OpenXml.Office2016.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Vml;
using PROGLOBAL_DataGestionAjuste_addon_EA.Common;
using PROGLOBAL_DataGestionAjuste_addon_EA.Models;
using PROGLOBAL_ReservationInvoiceCloser.Services;
using REDFARM.Addons.Tools;
using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;

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
                oDataTableGastos.Clear();
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
                oDataTableVentas.Clear();
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
                string dateFrom = _oForm.Items.Item(_itemDateFrom).Specific.Value;
                string dateTo = _oForm.Items.Item(_itemDateTo).Specific.Value;

                DateTime dateFromParser = DateTime.ParseExact(dateFrom, "yyyyMMdd", CultureInfo.InvariantCulture);
                DateTime dateToParser = DateTime.ParseExact(dateTo, "yyyyMMdd", CultureInfo.InvariantCulture);

                Grid GRIDGastos = item.Specific;

                for (int i = 0; i < GRIDGastos.Rows.Count; i++)
                {

                    string account = GRIDGastos.DataTable.Columns.Item(5).Cells.Item(i).Value;
                    double ajuste = GRIDGastos.DataTable.Columns.Item(8).Cells.Item(i).Value;

                    if (ajuste > 0)
                    {

                        SAPbobsCOM.CompanyService? companyService = ConnectionSDK.DIAPI!.GetCompanyService();
                        SAPbobsCOM.GeneralService? generalService = companyService.GetGeneralService("GESTIONAJUSTE");
                        GeneralData generalData;
                        _oRecordset = ConnectionSDK.DIAPI.GetBusinessObject(BoObjectTypes.BoRecordset);

                        _oRecordset.DoQuery(@$"SELECT TOP 1 ""Code"" FROM ""@GESTIONAJUSTE"" WHERE ""U_Detail"" = '{account}' AND ""U_DateFrom"" = '{dateFromParser.ToString("yyyy-MM-dd")}' AND ""U_DateTo"" = '{dateToParser.ToString("yyyy-MM-dd")}' AND ""U_Entity"" = '0'");
                        string? code = _oRecordset.Fields.Item(0).Value;
                        bool existAjuste = !string.IsNullOrEmpty(code);
                        if (existAjuste)
                        {
                            GeneralDataParams generalDataParams = (GeneralDataParams)generalService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralDataParams);
                            generalDataParams.SetProperty("Code", code);

                            generalData = generalService.GetByParams(generalDataParams);

                            generalData.SetProperty("U_Ajuste", ajuste);
                            generalService.Update(generalData);
                        } else
                        {

                            generalData = (SAPbobsCOM.GeneralData)generalService!.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);
                        
                            _oRecordset.DoQuery(@"SELECT MAX(""Code"") FROM ""@GESTIONAJUSTE""");
                            code = string.IsNullOrEmpty(_oRecordset.Fields.Item(0).Value) ? "1" : ((int)int.Parse(_oRecordset.Fields.Item(0).Value) + 1).ToString();

                            generalData.SetProperty("Code", code);
                            generalData.SetProperty("U_Detail", account);
                            generalData.SetProperty("U_Ajuste", ajuste);
                            generalData.SetProperty("U_DateFrom", dateFromParser);
                            generalData.SetProperty("U_DateTo", dateToParser);
                            generalData.SetProperty("U_Entity", "0");

                            generalService.Add(generalData);

                        }


                    }
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
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
                        TotalIndirecto = totalIndirecto
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
                        Costos = totalCosto,
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
                        Costos = d.Sum(data => data.Costos)
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
                             Costos = ventas.Costos,
                             Directo = gastos.TotalDirecto,
                             Indirecto = gastos.TotalIndirecto
                         };


            return totals.ToList();
        }

        public static void RefreshDataTotalesGrid(ReportExcelFormatSheet sheet)
        {
            try
            {
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
                oDataTableTotales.Clear();

                List<string> columns = ["Ventas", "Costos", "Margen", "% s. ventas (1)", "Directos", "% s. ventas (2)", "Indirectos",
                    "% s. ventas (3)", "T. Gastos", "Mensual", "% s. ventas (4)" , "Comisiones", "Acumulado (1)", "% s. ventas (5)", "Intereses", "Acumulado (2)", "% s. ventas (6)"];

                oDataTableTotales.Columns.Add("Detalle", BoFieldsType.ft_AlphaNumeric);
                columns.ForEach(colName => oDataTableTotales.Columns.Add(colName, BoFieldsType.ft_Float));

                var dataTotals = GetDataTotals(sheet);
                int countIndex = 0;

                // INDUSTRIA
                var totals_IND = LoadDataInDataTable(dataTotals, "^IND", oDataTableTotales, countIndex);
                countIndex = oDataTableTotales.Rows.Count;

                // AGRO
                var totals_AGRO = LoadDataInDataTable(dataTotals, "^AGRO", oDataTableTotales, countIndex);
                countIndex = oDataTableTotales.Rows.Count;

                // TOTAL COMPONENTES
                LoadInTable_TOTALCOMPONENTES(totals_IND, totals_AGRO, oDataTableTotales, countIndex);
                countIndex++;

                //// EQUIPO TECNICO
                var totals_ET = LoadDataInDataTable(dataTotals, "^ET", oDataTableTotales, countIndex);
                countIndex = oDataTableTotales.Rows.Count;

                // TOTAL PROGLOBAL
                LoadInTable_TOTALPROGLOBAL(totals_IND, totals_AGRO, totals_ET, oDataTableTotales, countIndex);


                GRIDTotales.DataTable = oDataTableTotales;

                for (int i = 0; i < GRIDTotales.DataTable.Rows.Count; i++)
                {
                    string detalle = GRIDTotales.DataTable.GetValue(0, i);
                    switch (detalle)
                    {
                        case "DIVISION INDUSTRIA":
                        case "DIVISION AGRO":
                            GRIDTotales.CommonSetting.SetRowBackColor(i + 1, 16777138);
                            break;
                        case "TOTAL COMPONENTES":
                        case "DIVISION EQUIPO TECNICO":
                            GRIDTotales.CommonSetting.SetRowBackColor(i + 1, 16776960);
                            break;
                        case "TOTAL PROGLOBAL":
                            GRIDTotales.CommonSetting.SetRowBackColor(i + 1, 14483330);
                            break;
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


        public static TotalsEntity LoadDataInDataTable(List<VentasGastoTotales> dataTotals, string regexEntity, SAPbouiCOM.DataTable oDataTableTotales, int countIndex)
        {
            string Detalle = regexEntity switch
            {
                "^IND" => "DIVISION INDUSTRIA",
                "^AGRO" => "DIVISION AGRO",
                "^ET" => "DIVISION EQUIPO TECNICO",
                _ => "No definido"
            };

            var records = dataTotals.Where(data => Regex.IsMatch(data.CC!, $"^{regexEntity}"));

            var totalVenta = records.Sum(r => r.Ventas);
            var totalCosto = records.Sum(r => r.Costos);
            var totalMargen = totalVenta - totalCosto; // Ventas - Costos
            var totalPorcVenta = totalVenta != 0 ? totalMargen / totalVenta * 100 : 0; // Margen / Ventas * 100

            var totalDirecto = records.Sum(r => r.Directo);
            var totalPorcVtaDirect = totalVenta != 0 ? totalDirecto / totalVenta * 100 : 0; // Directo / Ventas * 100
            var totalIndirecto = records.Sum(r => r.Indirecto);
            var totalPorcVtaIndirect = totalVenta != 0 ? totalIndirecto / totalVenta * 100 : 0; // Indirecto / Ventas
            var totalTotalGastos = totalDirecto + totalIndirecto;   // Directo + Indirecto 

            var totalMensual = totalMargen - totalDirecto - totalIndirecto; // MargenVentas - Directo - Indirecto
            var totalPorcVtaMensual = totalVenta != 0 ? totalMensual / totalVenta * 100 : 0; // Mensual / Ventas * 100
            var totalComisiones = records.Sum(r => r.Comision); // 0 (manual)
            var totalAcumuladoMensualComision = totalMensual + totalComisiones;  // Mensual + Comision
            var totalPorcVtaAcumuladoMensualComision = totalVenta != 0 ? totalAcumuladoMensualComision / totalVenta * 100 : 0; // Acumulado (1) / Ventas * 100

            var totalIntereses = records.Sum(r => r.Intereses);
            var totalAcumuladoIntereses = totalAcumuladoMensualComision + totalIntereses;
            var totalPorcAcumuladoIntereses = totalVenta != 0 ? totalAcumuladoIntereses / totalVenta * 100 : 0; // Acumulado (2) / Ventas * 100

            oDataTableTotales.Rows.Add();
            oDataTableTotales.SetValue("Detalle", countIndex, Detalle);
            oDataTableTotales.SetValue("Ventas", countIndex, totalVenta);
            oDataTableTotales.SetValue("Costos", countIndex, totalCosto);
            oDataTableTotales.SetValue("Margen", countIndex, totalMargen);
            oDataTableTotales.SetValue("% s. ventas (1)", countIndex, totalPorcVenta);

            oDataTableTotales.SetValue("Directos", countIndex, totalDirecto);
            oDataTableTotales.SetValue("% s. ventas (2)", countIndex, totalPorcVtaDirect);
            oDataTableTotales.SetValue("Indirectos", countIndex, totalIndirecto);
            oDataTableTotales.SetValue("% s. ventas (3)", countIndex, totalPorcVtaIndirect);
            oDataTableTotales.SetValue("T. Gastos", countIndex, totalTotalGastos);

            oDataTableTotales.SetValue("Mensual", countIndex, totalMensual);
            oDataTableTotales.SetValue("% s. ventas (4)", countIndex, totalPorcVtaMensual);
            oDataTableTotales.SetValue("Comisiones", countIndex, totalComisiones);
            oDataTableTotales.SetValue("Acumulado (1)", countIndex, totalAcumuladoMensualComision);
            oDataTableTotales.SetValue("% s. ventas (5)", countIndex, totalPorcVtaAcumuladoMensualComision);

            oDataTableTotales.SetValue("Intereses", countIndex, totalIntereses);
            oDataTableTotales.SetValue("Acumulado (2)", countIndex, totalAcumuladoIntereses);
            oDataTableTotales.SetValue("% s. ventas (6)", countIndex, totalPorcAcumuladoIntereses);

            countIndex = oDataTableTotales.Rows.Count;
            oDataTableTotales.Rows.Add(records.Count());

            for (int idata = 0; idata < records.Count(); idata++)
            {
                var data = records.ToList()[idata];

                oDataTableTotales.SetValue("Detalle", countIndex, data.CC);
                oDataTableTotales.SetValue("Ventas", countIndex, data.Ventas);
                oDataTableTotales.SetValue("Costos", countIndex, data.Costos);
                oDataTableTotales.SetValue("Margen", countIndex, data.MargenVentas);
                oDataTableTotales.SetValue("% s. ventas (1)", countIndex, data.PorcentajeVentas);

                oDataTableTotales.SetValue("Directos", countIndex, data.Directo);
                oDataTableTotales.SetValue("% s. ventas (2)", countIndex, data.PorcentajeVentaDirecto);
                oDataTableTotales.SetValue("Indirectos", countIndex, data.Indirecto);
                oDataTableTotales.SetValue("% s. ventas (3)", countIndex, data.PorcentajeVentaIndirecto);
                oDataTableTotales.SetValue("T. Gastos", countIndex, data.TotalGastos);

                oDataTableTotales.SetValue("Mensual", countIndex, data.Mensual);
                oDataTableTotales.SetValue("% s. ventas (4)", countIndex, data.PorcentajeVentaMensual);
                oDataTableTotales.SetValue("Comisiones", countIndex, data.Comision);
                oDataTableTotales.SetValue("Acumulado (1)", countIndex, data.AcumuladoMensualComision);
                oDataTableTotales.SetValue("% s. ventas (5)", countIndex, data.PorcentajeVentaAcumuladoMensualComision);

                oDataTableTotales.SetValue("Intereses", countIndex, data.Intereses);
                oDataTableTotales.SetValue("Acumulado (2)", countIndex, data.AcumuladoIntereses);
                oDataTableTotales.SetValue("% s. ventas (6)", countIndex, data.PorcentajeAcumuladoIntereses);

                countIndex++;
            }

            return new TotalsEntity()
            {
                totalVenta = totalVenta,
                totalCosto = totalCosto,
                totalMargen = totalMargen,
                totalPorcVenta = totalPorcVenta,
                totalDirecto = totalDirecto,
                totalPorcVtaDirect = totalPorcVtaDirect,
                totalIndirecto = totalIndirecto,
                totalPorcVtaIndirect = totalPorcVtaIndirect,
                totalTotalGastos = totalTotalGastos,
                totalMensual = totalMensual,
                totalPorcVtaMensual = totalPorcVtaMensual,
                totalComisiones = totalComisiones,
                totalAcumuladoMensualComision = totalAcumuladoMensualComision,
                totalPorcVtaAcumuladoMensualComision = totalPorcVtaAcumuladoMensualComision,
                totalIntereses = totalIntereses,
                totalAcumuladoIntereses = totalAcumuladoIntereses,
                totalPorcAcumuladoIntereses = totalPorcAcumuladoIntereses,         
            };
        }

        public static void LoadInTable_TOTALCOMPONENTES(TotalsEntity totals_IND, TotalsEntity totals_AGRO, SAPbouiCOM.DataTable oDataTableTotales, int countIndex)
        {
            // TOTAL COMPONENTES
            var totalComponentes = new
            {
                Detalle = "TOTAL COMPONENTES",
                Ventas = totals_IND.totalVenta + totals_AGRO.totalVenta,
                Costos = totals_IND.totalCosto + totals_AGRO.totalCosto,
                Margen = totals_IND.totalMargen + totals_AGRO.totalMargen,
                PorcVenta_1 = totals_IND.totalPorcVenta + totals_AGRO.totalPorcVenta,
                Directo = totals_IND.totalDirecto + totals_AGRO.totalDirecto,
                PorcVenta_2 = totals_IND.totalPorcVtaDirect + totals_AGRO.totalPorcVtaDirect,
                Indirecto = totals_IND.totalIndirecto + totals_AGRO.totalIndirecto,
                PorcVenta_3 = totals_IND.totalPorcVtaIndirect + totals_AGRO.totalPorcVtaIndirect,
                TotalGastos = totals_IND.totalTotalGastos + totals_AGRO.totalTotalGastos,
                Mensual = totals_IND.totalMensual + totals_AGRO.totalMensual,
                PorcentajeVentaMensual = totals_IND.totalPorcVtaMensual + totals_AGRO.totalPorcVtaMensual,
                Comisiones = totals_IND.totalComisiones + totals_AGRO.totalComisiones,
                AcumuladoMensualComisiones = totals_IND.totalAcumuladoMensualComision + totals_AGRO.totalAcumuladoMensualComision,
                PorcentajeVentaAcumuladoMensualComisiones = totals_IND.totalPorcVtaAcumuladoMensualComision + totals_AGRO.totalPorcVtaAcumuladoMensualComision,
                Intereses = totals_IND.totalIntereses + totals_AGRO.totalIntereses,
                AcumuladoIntereses = totals_IND.totalAcumuladoIntereses + totals_AGRO.totalAcumuladoIntereses,
                PorcentageAcumuladoIntereses = totals_IND.totalPorcAcumuladoIntereses + totals_AGRO.totalPorcAcumuladoIntereses
            };

                oDataTableTotales.Rows.Add();
                oDataTableTotales.SetValue("Detalle", countIndex, totalComponentes.Detalle);
                oDataTableTotales.SetValue("Ventas", countIndex, totalComponentes.Ventas);
                oDataTableTotales.SetValue("Costos", countIndex, totalComponentes.Costos);
                oDataTableTotales.SetValue("Margen", countIndex, totalComponentes.Margen);
                oDataTableTotales.SetValue("% s. ventas (1)", countIndex, totalComponentes.PorcVenta_1);

                oDataTableTotales.SetValue("Directos", countIndex, totalComponentes.Directo);
                oDataTableTotales.SetValue("% s. ventas (2)", countIndex, totalComponentes.PorcVenta_2);
                oDataTableTotales.SetValue("Indirectos", countIndex, totalComponentes.Indirecto);
                oDataTableTotales.SetValue("% s. ventas (3)", countIndex, totalComponentes.PorcVenta_3);
                oDataTableTotales.SetValue("T. Gastos", countIndex, totalComponentes.TotalGastos);

                oDataTableTotales.SetValue("Mensual", countIndex, totalComponentes.Mensual);
                oDataTableTotales.SetValue("% s. ventas (4)", countIndex, totalComponentes.PorcentajeVentaMensual);
                oDataTableTotales.SetValue("Comisiones", countIndex, totalComponentes.Comisiones);
                oDataTableTotales.SetValue("Acumulado (1)", countIndex, totalComponentes.AcumuladoMensualComisiones);
                oDataTableTotales.SetValue("% s. ventas (5)", countIndex, totalComponentes.PorcentajeVentaAcumuladoMensualComisiones);

                oDataTableTotales.SetValue("Intereses", countIndex, totalComponentes.Intereses);
                oDataTableTotales.SetValue("Acumulado (2)", countIndex, totalComponentes.AcumuladoIntereses);
                oDataTableTotales.SetValue("% s. ventas (6)", countIndex, totalComponentes.PorcentageAcumuladoIntereses);
                
            }


        public static void LoadInTable_TOTALPROGLOBAL(TotalsEntity totals_IND, TotalsEntity totals_AGRO, TotalsEntity totals_ET, SAPbouiCOM.DataTable oDataTableTotales, int countIndex)
        {
            // TOTAL COMPONENTES
            var totalComponentes = new
            {
                Detalle = "TOTAL PROGLOBAL",
                Ventas = totals_IND.totalVenta + totals_AGRO.totalVenta + totals_ET.totalVenta,
                Costos = totals_IND.totalCosto + totals_AGRO.totalCosto + totals_ET.totalCosto,
                Margen = totals_IND.totalMargen + totals_AGRO.totalMargen + totals_ET.totalMargen,
                PorcVenta_1 = totals_IND.totalPorcVenta + totals_AGRO.totalPorcVenta + totals_ET.totalPorcVenta,
                Directo = totals_IND.totalDirecto + totals_AGRO.totalDirecto + totals_ET.totalDirecto,
                PorcVenta_2 = totals_IND.totalPorcVtaDirect + totals_AGRO.totalPorcVtaDirect + totals_ET.totalPorcVtaDirect,
                Indirecto = totals_IND.totalIndirecto + totals_AGRO.totalIndirecto + totals_ET.totalIndirecto,
                PorcVenta_3 = totals_IND.totalPorcVtaIndirect + totals_AGRO.totalPorcVtaIndirect + totals_ET.totalPorcVtaIndirect,
                TotalGastos = totals_IND.totalTotalGastos + totals_AGRO.totalTotalGastos + totals_ET.totalTotalGastos,
                Mensual = totals_IND.totalMensual + totals_AGRO.totalMensual + totals_ET.totalMensual,
                PorcentajeVentaMensual = totals_IND.totalPorcVtaMensual + totals_AGRO.totalPorcVtaMensual + totals_ET.totalPorcVtaMensual,
                Comisiones = totals_IND.totalComisiones + totals_AGRO.totalComisiones + totals_ET.totalComisiones,
                AcumuladoMensualComisiones = totals_IND.totalAcumuladoMensualComision + totals_AGRO.totalAcumuladoMensualComision + totals_ET.totalAcumuladoMensualComision,
                PorcentajeVentaAcumuladoMensualComisiones = totals_IND.totalPorcVtaAcumuladoMensualComision + totals_AGRO.totalPorcVtaAcumuladoMensualComision + totals_ET.totalPorcVtaAcumuladoMensualComision,
                Intereses = totals_IND.totalIntereses + totals_AGRO.totalIntereses + totals_ET.totalIntereses,
                AcumuladoIntereses = totals_IND.totalAcumuladoIntereses + totals_AGRO.totalAcumuladoIntereses + totals_ET.totalAcumuladoIntereses,
                PorcentageAcumuladoIntereses = totals_IND.totalPorcAcumuladoIntereses + totals_AGRO.totalPorcAcumuladoIntereses + totals_ET.totalPorcAcumuladoIntereses
            };

            oDataTableTotales.Rows.Add();
            oDataTableTotales.SetValue("Detalle", countIndex, totalComponentes.Detalle);
            oDataTableTotales.SetValue("Ventas", countIndex, totalComponentes.Ventas);
            oDataTableTotales.SetValue("Costos", countIndex, totalComponentes.Costos);
            oDataTableTotales.SetValue("Margen", countIndex, totalComponentes.Margen);
            oDataTableTotales.SetValue("% s. ventas (1)", countIndex, totalComponentes.PorcVenta_1);

            oDataTableTotales.SetValue("Directos", countIndex, totalComponentes.Directo);
            oDataTableTotales.SetValue("% s. ventas (2)", countIndex, totalComponentes.PorcVenta_2);
            oDataTableTotales.SetValue("Indirectos", countIndex, totalComponentes.Indirecto);
            oDataTableTotales.SetValue("% s. ventas (3)", countIndex, totalComponentes.PorcVenta_3);
            oDataTableTotales.SetValue("T. Gastos", countIndex, totalComponentes.TotalGastos);

            oDataTableTotales.SetValue("Mensual", countIndex, totalComponentes.Mensual);
            oDataTableTotales.SetValue("% s. ventas (4)", countIndex, totalComponentes.PorcentajeVentaMensual);
            oDataTableTotales.SetValue("Comisiones", countIndex, totalComponentes.Comisiones);
            oDataTableTotales.SetValue("Acumulado (1)", countIndex, totalComponentes.AcumuladoMensualComisiones);
            oDataTableTotales.SetValue("% s. ventas (5)", countIndex, totalComponentes.PorcentajeVentaAcumuladoMensualComisiones);

            oDataTableTotales.SetValue("Intereses", countIndex, totalComponentes.Intereses);
            oDataTableTotales.SetValue("Acumulado (2)", countIndex, totalComponentes.AcumuladoIntereses);
            oDataTableTotales.SetValue("% s. ventas (6)", countIndex, totalComponentes.PorcentageAcumuladoIntereses);

        }
    }

    public class VentasGastoTotales
    {
        public string? CC { get; set; }
        public double? Ventas { get; set; }
        public double? Costos { get; set; }
        public double? MargenVentas { get => Ventas - Costos; }
        public double? PorcentajeVentas { get => Ventas != 0 ? (Ventas - Costos) / Ventas * 100 : 0; }
        public double? Directo { get; set; }
        public double? PorcentajeVentaDirecto { get => Ventas != 0 ? Directo / Ventas * 100 : 0; }
        public double? Indirecto { get; set; }
        public double? PorcentajeVentaIndirecto { get => Ventas != 0 ? Indirecto / Ventas * 100 : 0; }
        public double? TotalGastos { get => Directo + Indirecto; }
        public double? Mensual { get => MargenVentas - Directo - Indirecto; }
        public double? PorcentajeVentaMensual { get => Ventas != 0 ? Mensual / Ventas * 100 : 0; }
        public double? Comision { get => 0; }
        public double? AcumuladoMensualComision { get => Mensual + Comision; }
        public double? PorcentajeVentaAcumuladoMensualComision { get => Ventas != 0 ? AcumuladoMensualComision / Ventas * 100 : 0; }
        public double? Intereses { get => 0; }
        public double? AcumuladoIntereses { get => 0; }
        public double? PorcentajeAcumuladoIntereses { get => 0; }

    }


   public class TotalsEntity
    {
        public double? totalVenta { get; set; }
        public double? totalCosto{ get; set; }
        public double? totalMargen { get; set; }
        public double? totalPorcVenta { get; set; }
        public double? totalDirecto { get; set; }
        public double? totalPorcVtaDirect { get; set; }
        public double? totalIndirecto { get; set; }
        public double? totalPorcVtaIndirect { get; set; }
        public double? totalTotalGastos { get; set; }
        public double? totalMensual { get; set; }
        public double? totalPorcVtaMensual { get; set; }
        public double? totalComisiones { get; set; }
        public double? totalAcumuladoMensualComision { get; set; }
        public double? totalPorcVtaAcumuladoMensualComision { get; set; }
        public double? totalIntereses { get; set; }
        public double? totalAcumuladoIntereses { get; set; }
        public double? totalPorcAcumuladoIntereses { get; set; }
}
}
