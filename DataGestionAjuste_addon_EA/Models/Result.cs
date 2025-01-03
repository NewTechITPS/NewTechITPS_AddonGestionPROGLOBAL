using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PROGLOBAL_DataGestionAjuste_addon_EA.Models
{
    public class Result<T>
    {
        public string? Error{ get; set; }
        public string? Message { get; set; }
        public T? Data { get; set; }
    }
}
