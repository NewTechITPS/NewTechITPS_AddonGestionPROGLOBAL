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

        private string _firstDate;
        public string FirstDate { 
            
            get => _firstDate; 
            
            set {
                if (_firstDate == null)
                {
                    _firstDate = value;
                }
            } }


        public List<DateTime>? InformeStartDates { get; set; }
        public List<DateTime>? InformeEndDates { get; set; }
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
        public string? DateFrom { get; set; }
        public string? DateTo { get; set; }

        public string? SheetName { get; set; }

        public string? TitleExpenses { get; set; }
        public DataTable? DataTableExpenses { get; set; }

        public string? TitleSales { get; set; }
        public DataTable? DataTableSales { get; set; }

        public string? TitleTotalsSales { get; set; }
        public DataTable? DataTableTotalsSales { get; set; }

        public string? TitleTotalsExpenses { get; set; }
        public DataTable? DataTableTotalsExpenses { get; set; }

        public ReportExcelFormatSheet()
        {
            DataTableExpenses = new();
            DataTableSales = new();
            DataTableTotalsSales = new();
            DataTableTotalsExpenses = new();
        }
    }

}
