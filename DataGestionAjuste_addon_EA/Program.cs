using PROGLOBAL_DataGestionAjuste_addon_EA.Common;
using PROGLOBAL_DataGestionAjuste_addon_EA.Services;
using SAPbouiCOM;


try
{
    ConnectionSDK.Singlenton();
    //QueryServices.Singlenton();

    if (ConnectionSDK.Connected)
    {
        ConnectionSDK.UIAPI?.StatusBar.SetText("Addon DataGestionAjuste Connected", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Success);
        var oEvents = new Events();
        var oExecutions = new ExecutionsApp();
      
        GC.KeepAlive(oEvents);
        GC.KeepAlive(oExecutions);

        System.Windows.Forms.Application.Run();
    }
}
catch (Exception ex)
{
    ConnectionSDK.UIAPI?.MessageBox($"Fatal error in addon DataGestionAjuste: {ex.Message}");
}

        


