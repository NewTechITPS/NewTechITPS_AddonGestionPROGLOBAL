using PROGLOBAL_DataGestionAjuste_addon_EA.Forms.WINDOW;
using PROGLOBAL_DataGestionAjuste_addon_EA.Services;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Drawing.Text;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PROGLOBAL_DataGestionAjuste_addon_EA.Common
{
    public interface IEvents
    {

    }

    public class Events : IEvents
    {
        public Events()
        {
            
            try
            {
                ConnectionSDK.UIAPI!.AppEvent += OSAPB1appl_AppEvent;
                ConnectionSDK.UIAPI.MenuEvent += OSAPB1appl_MenuEvent;
                ConnectionSDK.UIAPI.ItemEvent += OSAPB1appl_ItemEvent;
                ConnectionSDK.UIAPI.FormDataEvent += OSAPB1appl_FormDataEvent;

                FilterEvents();
            }
            catch (Exception ex)
            {
                ConnectionSDK.UIAPI?.StatusBar.SetText("CONSTRUCTOR Eventos()" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        #region Eventos
        private void OSAPB1appl_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                switch (pVal.FormTypeEx)
                {
                    case VentanaGestion.frmUID:
                        VentanaGestion oVentanaGestion = new VentanaGestion();
                        oVentanaGestion.OSAPB1appl_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                    break;

                }
            }
            catch (Exception ex)
            {
                ConnectionSDK.UIAPI?.StatusBar.SetText("ITEM Eventos()" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        private void OSAPB1appl_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                switch (BusinessObjectInfo.FormTypeEx)
                {

                    case VentanaGestion.frmUID:
                        VentanaGestion oVentanaGestion = new VentanaGestion();
                        oVentanaGestion.OSAPB1appl_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                    break;

                }
            }
            catch (Exception ex)
            {
                ConnectionSDK.UIAPI?.StatusBar.SetText("FORM DATA Eventos()" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        private void OSAPB1appl_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                VentanaGestion oVentanaGestion = new VentanaGestion();
                oVentanaGestion.OSAPB1appl_MenuEvent(ref pVal, out BubbleEvent);
               
            } 
            catch (Exception ex)
            {
                ConnectionSDK.UIAPI?.StatusBar.SetText("MENU Eventos()" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        private void OSAPB1appl_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            try
            {
                //Switch para los diferentes tipos de eventos.
                //De esta forma terminamos la aplicación en cualquiera de todos estos eventos.
                switch (EventType)
                {
                    //Cuando se cierra la aplicación
                    case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                        ConnectionSDK.UIAPI?.StatusBar.SetText("Finalizando Addon DataGestionAjuste...");
                        System.Windows.Forms.Application.Exit();
                        break;
                    //Cuando se cambia la fuente o el lenguaje, nuestro formulario debe cargar de nuevo                    
                    case SAPbouiCOM.BoAppEventTypes.aet_FontChanged:
                    case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                        ConnectionSDK.UIAPI?.StatusBar.SetText("Reiniciando Addon DataGestionAjuste...");
                        System.Windows.Forms.Application.Restart();
                        break;
                }
            }
            catch (Exception ex)
            {
                ConnectionSDK.UIAPI?.StatusBar.SetText("APP Eventos()" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        #endregion

        #region Metodos

        private void FilterEvents()
        {
            //Revisar documentacion SDK
            SAPbouiCOM.EventFilters? oFiltros;
            SAPbouiCOM.EventFilter? oFiltro;

            try
            {
                //De esta forma filtro el evento que quiero escuchar en mi aplicación para que no este
                //escuchando todos los itemsevents.
                oFiltros = new SAPbouiCOM.EventFilters();
                oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD);
                oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD);
                oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE);
                oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE);
                oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST);
                oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD);
                oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE);
                oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD);
                oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN);
                oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_FORM_LOAD);
                oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_FORM_CLOSE);
                oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_COMBO_SELECT);
                oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_VALIDATE);
                oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_CLICK);
                oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE);
                oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE);
                oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_LOST_FOCUS);
                oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_KEY_DOWN);
                oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED);
                oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK);

                //oFiltro.AddEx(frmEntrega.TipoFormulariostr);
                //oFiltro.AddEx(frmAbono.strAbono);
                //oFiltro.AddEx(frmComisiones.frmComisionesfrm);
                //oFiltro.AddEx(frmPromociones.frmPromocionesstr);
                //oFiltro.AddEx(Formularios.FormMonitor.strFormulario);
                //oFiltro.AddEx(Formularios.FormMonitor.strFormulario);

                //Esto funciona en cascada, se define el filtro y los formularios que estan debajo, es para donde aplican esos filtros
                //De esa forma podemos filtrar en que formulario en especifico queremos que se use dicho filtro
                //De esta forma seteamos los filtros
                ConnectionSDK.UIAPI?.SetFilter(oFiltros);
            }
            catch (Exception ex)
            {
                ConnectionSDK.UIAPI?.StatusBar.SetText("CONFIG Eventos()" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        #endregion
    }
}
