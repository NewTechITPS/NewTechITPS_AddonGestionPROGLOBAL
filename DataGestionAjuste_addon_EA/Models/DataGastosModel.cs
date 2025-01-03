using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PROGLOBAL_DataGestionAjuste_addon_EA.Models
{
    public class DataGastosModel
    {


        [JsonProperty("LineID")]
        public int LineID { get; set; }

        [JsonProperty("Grupo Codigo")]
        public string? GrupoCodigo { get; set; }

        [JsonProperty("Grupo Nombre")]
        public string? GrupoNombre { get; set; }

        [JsonProperty("SubGrupo Codigo")]
        public string? SubGrupoCodigo { get; set; }

        [JsonProperty("SubGrupo Nombre")]
        public string? SubGrupoNombre { get; set; }

        [JsonProperty("Cuenta")]
        public string? Cuenta { get; set; }

        [JsonProperty("Nombre")]
        public string? Nombre { get; set; }

        [JsonProperty("Total de gastos")]
        public decimal TotalDeGastos { get; set; }

        [JsonProperty("U_Ajuste")]
        public decimal UAjuste { get; set; }

        [JsonProperty("Industria Directo")]
        public decimal IndustriaDirecto { get; set; }

        [JsonProperty("Industria Indirecto")]
        public decimal IndustriaIndirecto { get; set; }

        [JsonProperty("AGRO Directo")]
        public decimal AgroDirecto { get; set; }

        [JsonProperty("AGRO Indirecto")]
        public decimal AgroIndirecto { get; set; }

        [JsonProperty("ET (Equipo tecnico) Directo")]
        public decimal ETDirecto { get; set; }

        [JsonProperty("ET (Equipo tecnico) Indirecto")]
        public decimal ETIndirecto { get; set; }

        [JsonProperty("IND Puertos y Aceiteras Directo")]
        public decimal IndPuertosYAceiterasDirecto { get; set; }

        [JsonProperty("IND Puertos y Aceiteras Indirecto")]
        public decimal IndPuertosYAceiterasIndirecto { get; set; }

        [JsonProperty("IND  Medianos Directo")]
        public decimal IndMedianosDirecto { get; set; }

        [JsonProperty("IND  Medianos Indirecto")]
        public decimal IndMedianosIndirecto { get; set; }

        [JsonProperty("IND Pesados Directo")]
        public decimal IndPesadosDirecto { get; set; }

        [JsonProperty("IND Pesados Indirecto")]
        public decimal IndPesadosIndirecto { get; set; }

        [JsonProperty("Repuestos Directo")]
        public decimal RepuestosDirecto { get; set; }

        [JsonProperty("Repuestos Indirecto")]
        public decimal RepuestosIndirecto { get; set; }

        [JsonProperty("IND Industria Directo Directo")]
        public decimal IndIndustriaDirectoDirecto { get; set; }

        [JsonProperty("IND Industria Directo Indirecto")]
        public decimal IndIndustriaDirectoIndirecto { get; set; }

        [JsonProperty("IND Agro O.E.M Directo")]
        public decimal IndAgroOEMDirecto { get; set; }

        [JsonProperty("IND Agro O.E.M Indirecto")]
        public decimal IndAgroOEMIndirecto { get; set; }

        [JsonProperty("IND Agro Directo Directo")]
        public decimal IndAgroDirectoDirecto { get; set; }

        [JsonProperty("IND Agro Directo Indirecto")]
        public decimal IndAgroDirectoIndirecto { get; set; }

        [JsonProperty("IND Agro Reventa Directo")]
        public decimal IndAgroReventaDirecto { get; set; }

        [JsonProperty("IND Agro Reventa Indirecto")]
        public decimal IndAgroReventaIndirecto { get; set; }

        [JsonProperty("AGRO Comercio Norte Directo")]
        public decimal AgroComercioNorteDirecto { get; set; }

        [JsonProperty("AGRO Comercio Norte Indirecto")]
        public decimal AgroComercioNorteIndirecto { get; set; }

        [JsonProperty("AGRO Comercio Sur Directo")]
        public decimal AgroComercioSurDirecto { get; set; }

        [JsonProperty("AGRO Comercio Sur Indirecto")]
        public decimal AgroComercioSurIndirecto { get; set; }

        [JsonProperty("AGRO Acopios Norte Directo")]
        public decimal AgroAcopiosNorteDirecto { get; set; }

        [JsonProperty("AGRO Acopios Norte Indirecto")]
        public decimal AgroAcopiosNorteIndirecto { get; set; }

        [JsonProperty("AGRO Acopios Sur Directo")]
        public decimal AgroAcopiosSurDirecto { get; set; }

        [JsonProperty("AGRO Acopios Sur Indirecto")]
        public decimal AgroAcopiosSurIndirecto { get; set; }

        [JsonProperty("AGRO Acopios Litoral Directo")]
        public decimal AgroAcopiosLitoralDirecto { get; set; }

        [JsonProperty("AGRO Acopios Litoral Indirecto")]
        public decimal AgroAcopiosLitoralIndirecto { get; set; }

        [JsonProperty("AGRO Acopios Backoffice Directo")]
        public decimal AgroAcopiosBackofficeDirecto { get; set; }

        [JsonProperty("AGRO Acopios Backoffice Indirecto")]
        public decimal AgroAcopiosBackofficeIndirecto { get; set; }

        [JsonProperty("AGRO Acopios Directo Directo")]
        public decimal AgroAcopiosDirectoDirecto { get; set; }

        [JsonProperty("AGRO Acopios Directo Indirecto")]
        public decimal AgroAcopiosDirectoIndirecto { get; set; }

        [JsonProperty("ET Lights Directo")]
        public decimal ETLightsDirecto { get; set; }

        [JsonProperty("ET Lights Indirecto")]
        public decimal ETLightsIndirecto { get; set; }

        [JsonProperty("ET Oils and Ports Directo")]
        public decimal ETOilsAndPortsDirecto { get; set; }

        [JsonProperty("ET Oils and Ports Indirecto")]
        public decimal ETOilsAndPortsIndirecto { get; set; }

        [JsonProperty("ET Industries Directo")]
        public decimal ETIndustriesDirecto { get; set; }

        [JsonProperty("ET Industries Indirecto")]
        public decimal ETIndustriesIndirecto { get; set; }

        [JsonProperty("ET Spare Parts Directo")]
        public decimal ETSparePartsDirecto { get; set; }

        [JsonProperty("ET Spare Parts Indirecto")]
        public decimal ETSparePartsIndirecto { get; set; }
    }

}
