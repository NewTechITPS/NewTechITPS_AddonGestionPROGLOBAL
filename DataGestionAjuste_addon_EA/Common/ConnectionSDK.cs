using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PROGLOBAL_DataGestionAjuste_addon_EA.Common
{
    public interface IConnectionSDK
    {
        static SAPbouiCOM.Application? UIAPI { get; }
        static SAPbobsCOM.Company? DIAPI { get; }
        static bool Connected { get; }
    }

    public class ConnectionSDK : IConnectionSDK
    {
        protected static SAPbouiCOM.Application? _UIAPI;
        protected static SAPbobsCOM.Company? _DIAPI;

        public static SAPbouiCOM.Application? UIAPI => _UIAPI ?? throw new Exception("UIAPI no definido");
        public static SAPbobsCOM.Company? DIAPI => _DIAPI ?? throw new Exception("DIAPI no definido");

        public static bool Connected => _UIAPI != null && _DIAPI!.Connected;


        public static void Singlenton()
        {
            _UIAPI = GetApplication();

            if (_UIAPI != null)
            {
                _DIAPI = _UIAPI.Company.GetDICompany();
            }
        }



        public static SAPbouiCOM.Application? GetApplication()
        {
            SboGuiApi api = new()
            {
                AddonIdentifier = "DataGestionAjuste"
            };

            string[] commands = Environment.GetCommandLineArgs();
            string strConnection;

            if (commands.Length == 1) strConnection = commands[0];
            else
                if (commands[0].LastIndexOf("\\") > 0)
            {
                strConnection = commands[1];
            }
            else
            {
                strConnection = commands[0];
            }

            try
            {
                api.Connect(strConnection);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            return api.GetApplication();
        }
    }
}
