using PROGLOBAL_DataGestionAjuste_addon_EA.Common;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PROGLOBAL_ReservationInvoiceCloser.Services
{
    public interface INotificationService
    {
        void Error(string msg);
        void Warn(string msg);
        void Success(string msg);
    }

    public class NotificationService
    {

        public static void Error(string msg) => ConnectionSDK.UIAPI!.StatusBar.SetText(msg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);

        public static void Warn(string msg) => ConnectionSDK.UIAPI!.StatusBar.SetText(msg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);

        public static void Success(string msg) => ConnectionSDK.UIAPI!.StatusBar.SetText(msg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);


    }
}
