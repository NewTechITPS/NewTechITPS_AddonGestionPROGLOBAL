using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PROGLOBAL_DataGestionAjuste_addon_EA.Models
{
    public class ReportExcelFormat
    {
        public string? FileName { get; set; }
        public List<ReportExcelFormatSheet> Sheets { get; set; }

        public ReportExcelFormat()
        {
            Sheets = new List<ReportExcelFormatSheet>();
        }
        

        
        private static ReportExcelFormat? _instance;

        public static ReportExcelFormat Instance
        {
            get
            {
                
                if (_instance == null)
                {
                    _instance = new ReportExcelFormat();
                }
                return _instance;
                
            }
        }
    }
    public class ReportExcelFormatSheet
    {
        public string? SheetName { get; set; }

        public string? TitleExpenses { get; set; }
        public DataTable? DataTableExpenses { get; set; }

        public string? TitleSales { get; set; }
        public DataTable? DataTableSales { get; set; }

        public string? TitleTotals { get; set; }
        public DataTable? DataTableTotals { get; set; }

        public ReportExcelFormatSheet()
        {
            DataTableExpenses = new DataTable();
            DataTableSales = new DataTable();
            DataTableTotals = new DataTable();
        }

        //private static ReportExcelFormatSheet? _instance;
        //public static ReportExcelFormatSheet Instance
        //{
        //    get
        //    {

        //        if (_instance == null)
        //        {
        //            _instance = new ReportExcelFormatSheet();
        //        }
        //        return _instance;

        //    }
        //}
    }



}
