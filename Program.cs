using Newtonsoft.Json;
using System;
using System.Configuration;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace ActualizarBaseZKAccess
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            // Ruta a la base de datos
            string rutaMDB = ConfigurationManager
                                .ConnectionStrings["ZKAccess"]
                                .ConnectionString;

            //string rutaMDB = "C:/ZKTeco/ZKAccess3.5/Access.mdb";

            // Cargamos una sola vez la lista de usuarios
            var usuarios = ObtenerUsuarios(rutaMDB);

            while (true)
            {
                Console.Clear();
                Console.WriteLine("===== SISTEMA DE ACTUALIZACIÓN DE ACCESOS =====");
                Console.WriteLine("1. Extender acceso a socios habilitados");
                Console.WriteLine("2. Gestionar acceso de socios habilitados y no habilitados");
                Console.WriteLine("0. Salir");
                Console.Write("Seleccione una opción válida: ");

                string opcion = Console.ReadLine();
                Console.WriteLine();

                switch (opcion)
                {
                    case "1":
                        await RecorrerUsuarios(rutaMDB, usuarios);
                        break;

                    case "2":
                        await RecorrerUsuariosConNoHabilitados(rutaMDB, usuarios);
                        break;

                    case "0":
                        Console.WriteLine("Saliendo...");
                        return;

                    default:
                        Console.WriteLine("Opción no válida. Presione una tecla para reintentar.");
                        Console.ReadKey();
                        break;
                }
            }
        }

        class Usuario
        {
            public int USERID { get; set; }
            public string CardNo { get; set; }

            public string Name { get; set; }
        }

        static List<Usuario> ObtenerUsuarios(string rutaMDB)
        {
            var lista = new List<Usuario>();
            string cadenaConexion = $@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={rutaMDB};";

            using (OleDbConnection conexion = new OleDbConnection(cadenaConexion))
            {
                conexion.Open();
                string consulta = "SELECT USERID, Name, CardNo  FROM USERINFO WHERE CardNo IS NOT NULL AND CardNo <> ''";

                using (OleDbCommand comando = new OleDbCommand(consulta, conexion))
                using (OleDbDataReader lector = comando.ExecuteReader())
                {
                    while (lector.Read())
                    {
                        lista.Add(new Usuario
                        {
                            USERID = Convert.ToInt32(lector["USERID"]),
                            Name = lector["Name"].ToString(),
                            CardNo = lector["CardNo"].ToString()

                        });
                    }
                }
            }

            return lista;
        }

        static async Task RecorrerUsuarios(string rutaMDB, List<Usuario> usuarios)
        {
            // 1) Calcular la próxima fecha 10
            DateTime hoy = DateTime.Today;
            int mesObjetivo = (hoy.Day <= 10 ? hoy.Month : hoy.Month + 1);
            int añoObjetivo = hoy.Year;
            if (mesObjetivo == 13) { mesObjetivo = 1; añoObjetivo++; }
            DateTime fechaFin = new DateTime(añoObjetivo, mesObjetivo, 10);

            // La fecha de inicio se puede quedar como el Today si lo deseas:
            DateTime fechaInicio = hoy;

            Console.WriteLine($" FechaFin objetivo: {fechaFin:yyyy-MM-dd}");

            foreach (var u in usuarios)
            {
                Console.WriteLine($"Procesando {u.Name} (ID {u.USERID}), Tarjeta N°{u.CardNo}...");
                bool permitido = await ConsultarClubOnlineAsync(u.CardNo);

                if (!permitido)
                {
                    Console.WriteLine("   Estado: cuota impaga. No se modifica fecha de acceso.");
                    Console.WriteLine();
                    Console.WriteLine();
                    continue;
                }

                // 2) Leer la fecha fin actual desde la BD
                DateTime? fechaFinActual = ObtenerFechaFinUsuario(rutaMDB, u.USERID);

                // 3) Comparar fecha completa (día+mes+año)
                if (fechaFinActual.HasValue && fechaFinActual.Value.Date == fechaFin.Date)
                {
                    Console.WriteLine($"   Ya tiene acceso vigente hasta {fechaFinActual:yyyy-MM-dd}. No se realiza cambio.");
                }
                else
                {
                    ModificarValidezUsuario(rutaMDB, u.USERID, fechaInicio, fechaFin, true);
                    Console.WriteLine($"   Acceso extendido hasta {fechaFin:yyyy-MM-dd}.");
                }
                Console.WriteLine();
                Console.WriteLine();

            }

            Console.WriteLine("\n Proceso finalizado. Pulsa una tecla para salir.");
            Console.ReadKey();
        }

        static async Task RecorrerUsuariosConNoHabilitados(string rutaMDB, List<Usuario> usuarios)
        {
            DateTime hoy = DateTime.Today;
            

            // Próximo día 10
            int mesProx = (hoy.Day <= 10 ? hoy.Month : hoy.Month + 1);
            int añoProx = hoy.Year;
            if (mesProx == 13) { mesProx = 1; añoProx++; }
            DateTime fechaProx10 = new DateTime(añoProx, mesProx, 10);

            // Día 10 del mes anterior
            int mesAnt = hoy.Month - 1;
            int añoAnt = hoy.Year;
            if (mesAnt == 0) { mesAnt = 12; añoAnt--; }
            DateTime fechaAnt10 = new DateTime(añoAnt, mesAnt, 10);
            DateTime diaAntNo = new DateTime(añoAnt, mesAnt, 9); // Día siguiente al 10 del mes anterior

            Console.WriteLine($" Objetivos: FechaFinAccesoExtendido = {fechaProx10:yyyy-MM-dd}, FechaFinAccesoCancelado = {fechaAnt10:yyyy-MM-dd}");

            foreach (var u in usuarios)
            {
                Console.WriteLine($"Procesando {u.Name} (ID {u.USERID}), Tarjeta N°{u.CardNo}...");
                string status = await ConsultarClubOnlineAsyncConNoHabilitados(u.CardNo);

                switch (status)
                {
                    case "INEXISTENTE":
                        Console.WriteLine("   Usuario inexistente en ClubOnline. Sin cambios.");
                        Console.WriteLine();
                        Console.WriteLine();
                        continue;

                    case "OK":
                        // Leer fecha actual
                        DateTime? finOK = ObtenerFechaFinUsuario(rutaMDB, u.USERID);
                        if (finOK.HasValue && finOK.Value.Date == fechaProx10.Date)
                        {
                            Console.WriteLine($"   Ya tiene acceso vigente hasta {finOK:yyyy-MM-dd}. No se realiza cambio.");
                        }
                        else
                        {
                            ModificarValidezUsuario(rutaMDB, u.USERID, hoy, fechaProx10, true);
                            Console.WriteLine($"   Acceso extendido hasta {fechaProx10:yyyy-MM-dd}.");
                        }
                        break;

                    case "NO":
                        DateTime? finNO = ObtenerFechaFinUsuario(rutaMDB, u.USERID);
                        if (finNO.HasValue && finNO.Value.Date == fechaAnt10.Date)
                        {
                            Console.WriteLine($"   Acceso cancelado. Fecha de corte: {fechaAnt10:yyyy-MM-dd}. Ya estaba aplicado");
                        }
                        else
                        {
                            ModificarValidezUsuario(rutaMDB, u.USERID, diaAntNo, fechaAnt10, true);
                            Console.WriteLine($"   Acceso cancelado. Fecha de corte: {fechaAnt10:yyyy-MM-dd}.");
                        }
                        break;

                    default:
                        Console.WriteLine("    Respuesta desconocida. Verificar.");
                        break;
                }
                Console.WriteLine();
                Console.WriteLine();
            }
            Console.WriteLine("\nOperación completada.");
            Console.Write("Presione cualquier tecla para regresar al menú...");
            Console.ReadKey();
        }

        static DateTime? ObtenerFechaFinUsuario(string rutaMDB, int userId)
        {
            string connStr = $@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={rutaMDB};";
            using (var conexion = new OleDbConnection(connStr))
            {
                conexion.Open();
                string sql = "SELECT acc_enddate FROM USERINFO WHERE USERID = ?";
                using (var cmd = new OleDbCommand(sql, conexion))
                {
                    cmd.Parameters.AddWithValue("?", userId);
                    var result = cmd.ExecuteScalar();
                    if (result != DBNull.Value && result != null)
                        return Convert.ToDateTime(result);
                }
            }
            return null;
        }

        static void ModificarValidezUsuario(string rutaMDB, int userId, DateTime fechaInicio, DateTime fechaFin, bool establecerValidez)
        {
            string cadenaConexion = $@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={rutaMDB};";
            using (OleDbConnection conexion = new OleDbConnection(cadenaConexion))
            {
                conexion.Open();
                string consulta = "UPDATE USERINFO SET set_valid_time = ?, acc_startdate = ?, acc_enddate = ? WHERE USERID = ?";
                using (OleDbCommand comando = new OleDbCommand(consulta, conexion))
                {
                    comando.Parameters.AddWithValue("?", establecerValidez);
                    comando.Parameters.AddWithValue("?", fechaInicio);
                    comando.Parameters.AddWithValue("?", fechaFin);
                    comando.Parameters.AddWithValue("?", userId);
                    int filasAfectadas = comando.ExecuteNonQuery();
                    Console.WriteLine($"✅ Usuario {userId} actualizado ({filasAfectadas} fila/s).");
                }
            }
        }

        static async Task<bool> ConsultarClubOnlineAsync(string tarjeta)
        {
            using (var client = new HttpClient())
            {
                var payload = new
                {
                    idCompany = 63,
                    idAccessControlDevice = 1,
                    accessControlId = tarjeta
                };
                var json = JsonConvert.SerializeObject(payload);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                try
                {
                    var resp = await client.PostAsync("http://s008.myclubonline.com.ar:8080/Sasha/Door/API/RFID", content);
                    var txt = await resp.Content.ReadAsStringAsync();
                    Console.WriteLine($"   API respondió: {txt}");
                    return txt.Contains("\"opc\":\"OK\"");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("   ❌ Error en API: " + ex.Message);
                    return false;
                }
            }
        }

        static async Task<string> ConsultarClubOnlineAsyncConNoHabilitados(string tarjeta)
        {
            using (var client = new HttpClient())
            {
                var payload = new
                {
                    idCompany = 63,
                    idAccessControlDevice = 1,
                    accessControlId = tarjeta
                };
                var json = JsonConvert.SerializeObject(payload);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                try
                {
                    var resp = await client.PostAsync(
                        "http://s008.myclubonline.com.ar:8080/Sasha/Door/API/RFID",
                        content
                    );
                    var txt = await resp.Content.ReadAsStringAsync();
                    Console.WriteLine($"   API respondió: {txt}");

                    if (txt.Contains("\"opc\":\"OK\"")) return "OK";
                    else if (txt.Contains("\"renglon2\":\"NO HABILITADO\"")) return "NO";
                    else return "INEXISTENTE";
                }
                catch
                {
                    return "INEXISTENTE";
                }
            }
        }

    }
}
