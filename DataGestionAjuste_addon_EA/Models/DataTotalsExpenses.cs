using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PROGLOBAL_DataGestionAjuste_addon_EA.Models
{
    public class DataTotalsExpenses
    {
        public string? CodExternal { get; set; } 
        public double INDDirect { get; set; }  
        public double AGRODirect { get; set; } 
        public double ETDirect { get; set; }   
        public double INDIndirect { get; set; } 
        public double AGROIndirect { get; set; } 
        public double ETIndirect { get; set; }   
        
    }
}
