using Newtonsoft.Json;
using PROGLOBAL_DataGestionAjuste_addon_EA.Services;
using REDFARM.Addons;
using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Drawing.Text;
using System.Linq;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using System.Xml;

namespace PROGLOBAL_DataGestionAjuste_addon_EA.Common
{
    public interface IExecutionsApp
    {

    }

    public class ExecutionsApp : IExecutionsApp
    {

        public ExecutionsApp() 
        {
            VentanaGestionService.CreateMenu();
        }



    }
}
