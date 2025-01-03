using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using PROGLOBAL_DataGestionAjuste_addon_EA.Models;
using System.Windows.Forms;
using Sap.Data.Hana;

namespace PROGLOBAL_DataGestionAjuste_addon_EA.Services
{
    public class QueryServices 
    {
        protected static IDbConnection? _dbConnection;

        public static void Singlenton()
        {
            if(_dbConnection == null)
            {
                _dbConnection = new HanaConnection("SERVERNODE={saphaproglobal:30015};DSN=HANA;UID=B1ADMIN;PWD=@lbZ0i!034wa5Y*;CS=PROCESS_SRL_ITPS;databaseName=SK1");
            }
        }

        public static Result<string> CreateTemporalTable_Gastos()
        {
            Result<string> result = new();
            if (_dbConnection!.State != ConnectionState.Open) _dbConnection.Open();

            IDbCommand command = _dbConnection.CreateCommand();
            command.CommandText = @$"CALL INFORME_EA_GESTION_ACTIONS('CREATE_TABLE', 'GASTOS')";

            try
            {
                IDataReader r = command.ExecuteReader();
                r.Read();
                result.Message = "Tabla temporal creada con éxito!";
            }
            catch (Exception ex)
            {
                result.Message = "Tabla temporal no creada!";
                result.Error = ex.Message;
            }
            finally
            {
                if (_dbConnection.State == ConnectionState.Open) _dbConnection.Close();
            }

            return result;
        }

        public static Result<string> DropTemporalTable_Gastos()
        {
            Result<string> result = new();
            if (_dbConnection!.State != ConnectionState.Open) _dbConnection.Open();

            IDbCommand command = _dbConnection.CreateCommand();
            command.CommandText = @$"CALL INFORME_EA_GESTION_ACTIONS('DROP_TABLE', 'GASTOS')";

            try
            {
                IDataReader r = command.ExecuteReader();
                r.Read();
                result.Message = "Table temporal eliminada con éxito!";
            }
            catch (Exception ex)
            {
                result.Message = "Table temporal no eliminada";
                result.Error = ex.Message;
            }
            finally
            {
                if (_dbConnection.State == ConnectionState.Open) _dbConnection.Close();
            }

            return result;
        }

        public static Result<List<DataGastosModel>> SelectTemporalTable_Gastos()
        {
            Result<List<DataGastosModel>> result = new();

            if (_dbConnection!.State != ConnectionState.Open) _dbConnection.Open();

            IDbCommand command = _dbConnection.CreateCommand();
            command.CommandText = @$"CALL INFORME_EA_GESTION_ACTIONS('SELECT', 'GASTOS')";
            IDataReader r = command.ExecuteReader();
            
            try
            {
                while (r.Read())
                {
                    DataGastosModel data = new();
                    data.LineID = (int)r.GetValue(0);
                    data.GrupoCodigo = (string)r.GetValue(1);
                    data.GrupoNombre = (string)r.GetValue(2);
                    data.SubGrupoCodigo = (string)r.GetValue(3);
                    data.SubGrupoNombre = (string)r.GetValue(4);
                    data.Cuenta = (string)r.GetValue(5);
                    data.Nombre = (string)r.GetValue(6);
                    data.TotalDeGastos = (decimal)r.GetValue(7);
                    data.UAjuste = (decimal)r.GetValue(8);
                    data.IndustriaDirecto = (decimal)r.GetValue(9);
                    data.IndustriaIndirecto = (decimal)r.GetValue(10);
                    data.AgroDirecto =  (decimal)r.GetValue(11);
                    data.AgroIndirecto =  (decimal)r.GetValue(12);
                    data.ETDirecto =  (decimal)r.GetValue(13);
                    data.ETIndirecto =  (decimal)r.GetValue(14);
                    data.IndPuertosYAceiterasDirecto =  (decimal)r.GetValue(15);
                    data.IndPuertosYAceiterasIndirecto =  (decimal)r.GetValue(16);
                    data.IndMedianosDirecto =  (decimal)r.GetValue(17);
                    data.IndMedianosIndirecto =  (decimal)r.GetValue(18);
                    data.IndPesadosDirecto =  (decimal)r.GetValue(19);
                    data.IndPesadosIndirecto =  (decimal)r.GetValue(20);
                    data.RepuestosDirecto =  (decimal)r.GetValue(21);
                    data.RepuestosIndirecto =  (decimal)r.GetValue(22);
                    data.IndIndustriaDirectoDirecto =  (decimal)r.GetValue(23);
                    data.IndIndustriaDirectoIndirecto =  (decimal)r.GetValue(24);
                    data.IndAgroOEMDirecto =  (decimal)r.GetValue(25);
                    data.IndAgroOEMIndirecto =  (decimal)r.GetValue(26);
                    data.IndAgroDirectoDirecto =  (decimal)r.GetValue(27);
                    data.IndAgroDirectoIndirecto =  (decimal)r.GetValue(28);
                    data.IndAgroReventaDirecto =  (decimal)r.GetValue(29);
                    data.IndAgroReventaIndirecto =  (decimal)r.GetValue(30);
                    data.AgroComercioNorteDirecto =  (decimal)r.GetValue(31);
                    data.AgroComercioNorteIndirecto =  (decimal)r.GetValue(32);
                    data.AgroComercioSurDirecto =  (decimal)r.GetValue(33);
                    data.AgroComercioSurIndirecto =  (decimal)r.GetValue(34);
                    data.AgroAcopiosNorteDirecto =  (decimal)r.GetValue(35);
                    data.AgroAcopiosNorteIndirecto =  (decimal)r.GetValue(36);
                    data.AgroAcopiosSurDirecto =  (decimal)r.GetValue(37);
                    data.AgroAcopiosSurIndirecto =  (decimal)r.GetValue(38);
                    data.AgroAcopiosLitoralDirecto =  (decimal)r.GetValue(39);
                    data.AgroAcopiosLitoralIndirecto =  (decimal)r.GetValue(40);
                    data.AgroAcopiosBackofficeDirecto =  (decimal)r.GetValue(41);
                    data.AgroAcopiosBackofficeIndirecto =  (decimal)r.GetValue(42);
                    data.AgroAcopiosDirectoDirecto =  (decimal)r.GetValue(43);
                    data.AgroAcopiosDirectoIndirecto =  (decimal)r.GetValue(44);
                    data.ETLightsDirecto =  (decimal)r.GetValue(45);
                    data.ETLightsIndirecto =  (decimal)r.GetValue(46);
                    data.ETOilsAndPortsDirecto =  (decimal)r.GetValue(47);
                    data.ETOilsAndPortsIndirecto =  (decimal)r.GetValue(48);
                    data.ETIndustriesDirecto =  (decimal)r.GetValue(49);
                    data.ETIndustriesIndirecto =  (decimal)r.GetValue(50);
                    data.ETSparePartsDirecto =  (decimal)r.GetValue(51);
                    data.ETSparePartsIndirecto = (decimal)r.GetValue(52);

                    result.Data!.Add(data);

                }
                    result.Message = "Operación completada";
            }
            catch(Exception ex)
            {
                result.Message = ex.Message;
            }
            finally
            {
                if (_dbConnection.State == ConnectionState.Open) _dbConnection.Close();
            }

            return result;
        }

        public static Result<string> InsertDataDefaultTemporalTable_Gastos(string dateFrom, string dateTo)
        {
            Result<string> result = new();

            if (_dbConnection!.State != ConnectionState.Open) _dbConnection.Open();

            IDbCommand command = _dbConnection.CreateCommand();
            command.CommandText = @$"CALL INFORME_EA_GESTION_ACTIONS('INSERT', 'GASTOS', '{dateFrom}', '{dateTo}')";
            IDataReader r = command.ExecuteReader();

            try
            {
                r.Read();
                
                result.Message = "Operación completada";
            }
            catch (Exception ex)
            {
                result.Message = ex.Message;
            }
            finally
            {
                if (_dbConnection.State == ConnectionState.Open) _dbConnection.Close();
            }

            return result;
        }


        //public IList<ListaPrecio> GetListPrice()
        //{
        //    try
        //    {
        //        if (dbConnection.State != ConnectionState.Open)
        //            dbConnection.Open();

        //        using IDbCommand command = dbConnection.CreateCommand();
        //        command.CommandText = $@"SELECT T0.""ItemCode"",T0.""Price"",T0.""Currency"" 
        //                                    FROM ITM1 T0
        //                                    INNER JOIN OITM T1 ON T0.""ItemCode"" = T1.""ItemCode"" 
        //                                    WHERE T0.""PriceList"" = 16 AND T0.""Price"" > 0 AND T1.""QryGroup1"" = 'Y' ";

        //        using IDataReader reader = command.ExecuteReader();

        //        var listSN = new List<ListaPrecio>();

        //        while (reader.Read())
        //        {
        //            listSN.Add(new ListaPrecio
        //            {
        //                ItemCode = reader.GetString(reader.GetOrdinal("ItemCode")),
        //                Price = reader.GetDouble(reader.GetOrdinal("Price")),
        //                Currency = reader.GetString(reader.GetOrdinal("Currency"))
        //            });
        //        }

        //        return listSN;
        //    }
        //    finally
        //    {
        //        if (dbConnection.State == ConnectionState.Open)
        //            dbConnection.Close();
        //    }
        //}

        //public IList<FormaEnvio> GetFormaEnvio()
        //{
        //    try
        //    {
        //        if (dbConnection.State != ConnectionState.Open)
        //            dbConnection.Open();

        //        using IDbCommand command = dbConnection.CreateCommand();
        //        command.CommandText = $@"SELECT ""TrnspCode"",""TrnspName"" FROM OSHP";

        //        using IDataReader reader = command.ExecuteReader();

        //        var listSN = new List<FormaEnvio>();

        //        while (reader.Read())
        //        {
        //            listSN.Add(new FormaEnvio
        //            {
        //                id = reader.GetInt32(reader.GetOrdinal("TrnspCode")),
        //                name = reader.GetString(reader.GetOrdinal("TrnspName"))
        //            });
        //        }

        //        return listSN;
        //    }
        //    finally
        //    {
        //        if (dbConnection.State == ConnectionState.Open)
        //            dbConnection.Close();
        //    }
        //}

        //public int GetIdVtex(string code)
        //{
        //    try
        //    {
        //        if (dbConnection.State != ConnectionState.Open)
        //            dbConnection.Open();

        //        using IDbCommand command = dbConnection.CreateCommand();
        //        command.CommandText = @"SELECT ""DocEntry"" FROM ORDR WHERE ""U_Interfaz"" = 'VTEX' AND ""U_Code"" = ? AND CANCELED = 'N'";

        //        IDbDataParameter docEntryParameter = command.CreateParameter();
        //        docEntryParameter.ParameterName = "@Code";
        //        docEntryParameter.Value = code;
        //        command.Parameters.Add(docEntryParameter);

        //        using IDataReader reader = command.ExecuteReader();

        //        if (reader.Read())
        //        {
        //            return reader.GetInt32(0);
        //        }
        //    }
        //    finally
        //    {
        //        if (dbConnection.State == ConnectionState.Open)
        //            dbConnection.Close();
        //    }

        //    return 0;
        //}

        //public int GetUnDefault(string itemCode)
        //{
        //    try
        //    {
        //        if (dbConnection.State != ConnectionState.Open)
        //            dbConnection.Open();

        //        using IDbCommand command = dbConnection.CreateCommand();
        //        command.CommandText = @"SELECT ""U_DefaultUNVtex"" FROM OITM WHERE IFNULL(""U_DefaultUNVtex"",0) <> 0 AND ""QryGroup1"" = 'Y' AND ""ItemCode"" = ?";

        //        IDbDataParameter docEntryParameter = command.CreateParameter();
        //        docEntryParameter.ParameterName = "@ItemCode";
        //        docEntryParameter.Value = itemCode;
        //        command.Parameters.Add(docEntryParameter);

        //        using IDataReader reader = command.ExecuteReader();

        //        if (reader.Read())
        //        {
        //            return reader.GetInt32(0);
        //        }
        //    }
        //    finally
        //    {
        //        if (dbConnection.State == ConnectionState.Open)
        //            dbConnection.Close();
        //    }

        //    return 0;
        //}

        //public List<ResultGetOrdersVtexByStatus> GetOrdersVtexByStatus(string statusvtex = "PP")
        //{
        //    try
        //    {
        //        if (dbConnection.State != ConnectionState.Open)
        //            dbConnection.Open();

        //        using IDbCommand command = dbConnection.CreateCommand();
        //        command.CommandText = $@"SELECT 
        //                                T0.""DocEntry"", 
        //                                T0.""U_Code"",
        //                                COUNT(T1.""LineNum"") AS ""QuantityLines""
        //                                FROM ORDR T0 
        //                                INNER JOIN RDR1 T1 
        //                                ON T1.""DocEntry"" = T0.""DocEntry"" 
        //                                WHERE T0.CANCELED = 'N'
        //                                AND T0.""U_Interfaz"" = 'VTEX' 
        //                                AND IFNULL(T0.""U_Code"", '') <> '' 
        //                                AND T0.""U_Vtex_Status"" = ?
        //                                GROUP BY T0.""DocEntry"", T0.""U_Code""";

        //        IDbDataParameter param = command.CreateParameter();
        //        param.ParameterName = "@U_Vtex_Status";
        //        param.Value = statusvtex;
        //        command.Parameters.Add(param);

        //        using IDataReader reader = command.ExecuteReader();

        //        var orders = new List<ResultGetOrdersVtexByStatus>();

        //        while (reader.Read())
        //        {
        //            orders.Add(new ResultGetOrdersVtexByStatus
        //            {
        //                DocEntry = reader.GetInt32(reader.GetOrdinal("DocEntry")),
        //                U_Code = reader.GetString(reader.GetOrdinal("U_Code")),
        //                QuantityLines = reader.GetInt32(reader.GetOrdinal("QuantityLines"))
        //            });
        //        }

        //        return orders;
        //    }
        //    finally
        //    {
        //        if (dbConnection.State == ConnectionState.Open)
        //            dbConnection.Close();
        //    }

        //}

        //public ResultGetOrderVtexByDocEntry GetOrderVtexByDocEntry(int? DocEntry)
        //{
        //    try
        //    {
        //        if (dbConnection.State != ConnectionState.Open)
        //            dbConnection.Open();

        //        using IDbCommand command = dbConnection.CreateCommand();
        //        command.CommandText = $@"SELECT 
        //                                T0.""DocEntry"", 
        //                                T0.""U_Code"",
        //                                COUNT(T1.""LineNum"") AS ""QuantityLines""
        //                                FROM ORDR T0 
        //                                INNER JOIN RDR1 T1 
        //                                ON T1.""DocEntry"" = T0.""DocEntry"" 
        //                                WHERE T0.CANCELED = 'N'
        //                                AND T0.""U_Interfaz"" = 'VTEX' 
        //                                AND IFNULL(T0.""U_Code"", '') <> '' 
        //                                AND T0.""DocEntry"" = ?
        //                                GROUP BY T0.""DocEntry"", T0.""U_Code""";

        //        IDbDataParameter param = command.CreateParameter();
        //        param.ParameterName = "@DocEntry";
        //        param.Value = DocEntry;
        //        command.Parameters.Add(param);

        //        using IDataReader reader = command.ExecuteReader();

        //        reader.Read();

        //        var order = new ResultGetOrderVtexByDocEntry
        //        {
        //            DocEntry = reader.GetInt32(reader.GetOrdinal("DocEntry")),
        //            U_Code = reader.GetString(reader.GetOrdinal("U_Code")),
        //            QuantityLines = reader.GetInt32(reader.GetOrdinal("QuantityLines"))
        //        };

        //        return order;
        //    }
        //    finally
        //    {
        //        if (dbConnection.State == ConnectionState.Open)
        //            dbConnection.Close();
        //    }

        //}
    }
}
