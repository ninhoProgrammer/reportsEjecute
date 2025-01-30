using System;
using System.CodeDom;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Reporting.WinForms;


namespace repotsEjecute
{
    public class CreateReportsToPDF
    {
        ReportViewer reportViewer1;
        private System.Data.SqlClient.SqlConnection conexion;
        private DataSet ds;

        static void Main(string[] args)
        {
            CreateReportsToPDF createReportsToPDF = new CreateReportsToPDF();
            // Validar parámetros
            if (args.Length < 8)
            {
                Console.WriteLine("Error: Se necesitan 8 parámetros.");
                string[] parameters =
                {
                    "Z:\\RDLReportes\\RC_-_Balanza_catorcena.rdl",
                    "1",
                    "reporte1",
                    "2",
                    "2024",
                    "1",
                    "0",
                    "true"
                };
                args = parameters;
            }


            string path = args[0];
            string ReCodigo = args[1];
            string ReNombre = args[2];
            string peTipo = args[3];
            string peAnio = args[4];
            string peNumero = args[5];
            string otroParametro = args[6];
            bool Activo = true;


            // Llamar a la función existente
            string result = createReportsToPDF.creationsReports(path, ReCodigo, ReNombre, peTipo, peAnio, peNumero, otroParametro, Activo, null);

            // Guardar el resultado en un archivo .txt
            //createReportsToPDF.returnMessage(result);
            //Console.WriteLine(outputPath + path + ReCodigo+ ReNombre+ peTipo+ peAnio+ peNumero+ otroParametro+ Activo+ null);
            //File.WriteAllText(outputPath, result);

            // Imprimir el resultado en la salida estándar
            Console.WriteLine(result);
        }

        public CreateReportsToPDF()
        { 
        }

        public string creationsReports(string path, string ReCodigo, string ReNombre, string peTipo, string peAnio, string peNumero, string colabora, bool Activo, DataTable dataTable)
        {
            try
            {
                reportViewer1 = new ReportViewer();
                string activo = "Si";
                reportViewer1.LocalReport.ReportPath = path;

                if (File.Exists(path))
                {
                    // El reporte está listo para abrirse
                    if (Activo)
                    {
                        activo = "S";
                    }
                    else
                    {
                        activo = "N";
                    }
                    //vDatosPeriodo:
                    //[0]Año, [1]Tipo(Numero), [2]Numero, [3]Empleado, [4]Tipo(Texto), [5]Fecha Inicio, [6]Fecha Fin, [7]Nombre Empleado

                    List<ReportParameter> parametros = new List<ReportParameter>
                    {
                        new ReportParameter("EMPRESA", "BARRON"),
                        new ReportParameter("REPORTE", ReNombre),
                        new ReportParameter("ACTIVO", activo),
                        new ReportParameter("AÑO", peAnio),
                        new ReportParameter("PERIODO", peNumero),
                        new ReportParameter("TIPO", peTipo),
                        new ReportParameter("EMPLEADO", colabora)
                    };

                    var validParameters = reportViewer1.LocalReport.GetParameters();

                    // Filtrar y establecer solo los parámetros que están definidos en el informe
                    var parametersToSet = parametros
                        .Where(p => validParameters.Any(vp => vp.Name.Equals(p.Name, StringComparison.OrdinalIgnoreCase)))
                        .ToList();

                    if (parametersToSet.Any())
                    {
                        reportViewer1.LocalReport.SetParameters(parametersToSet);
                    }


                    //formReports.reportViewer1.Clear();
                    var parameters = reportViewer1.LocalReport.GetParameters();
                    var dataSourceNames = reportViewer1.LocalReport.GetDataSourceNames();

                    ReportDataSource reportDataSource = new ReportDataSource(dataSourceNames[0].ToString(), DataReporte(GetSqlQueryFromRdl(path), reportViewer1));
                    //ReportDataSource reportDataSource2 = new ReportDataSource("DataSet1", repRLDC.Tables["KARDEX"]);
                    if (reportDataSource.Value != null)
                    {
                        reportViewer1.LocalReport.DataSources.Clear();
                        reportViewer1.LocalReport.DataSources.Add(reportDataSource);

                        string outputFilePath = "C:\\Users\\mario\\Documents\\GitHub\\Reporteadores\\wwwroot\\Temp\\ScreenshotsReporte.pdf";

                        // Configura los parámetros para exportar
                        Warning[] warnings;
                        string[] streamIds;
                        string mimeType = string.Empty;
                        string encoding = string.Empty;
                        string extension = string.Empty;

                        // Renderiza el reporte en formato PDF
                        byte[] bytes = reportViewer1.LocalReport.Render(
                            "PDF", // Formato de exportación
                            null, // DeviceInfo (puede ser null para usar configuración por defecto)
                            out mimeType,
                            out encoding,
                            out extension,
                            out streamIds,
                            out warnings
                        );

                        // Guarda el PDF en el sistema de archivos
                        using (FileStream fs = new FileStream(outputFilePath, FileMode.Create))
                        {
                            fs.Write(bytes, 0, bytes.Length);
                        }

                        // Mensaje de confirmación
                        return outputFilePath;
                    }
                }
                else
                {
                    returnMessage("Message: sdas");
                }
            }
            catch (Exception ex)
            {
                returnMessage(ex.Message);
                return " ";
            }
            return " ";
        }

        public void returnMessage(string message)
        {
            string outputPath = "C:\\wwwroot\\Temp\\resultado.txt";
            Console.WriteLine(outputPath);
            File.WriteAllText(outputPath, message);
        }

        public string GetSqlQueryFromRdl(string rdlPath)
        {
            // Leer todo el contenido del archivo
            string content = null;
            string sqlQuery = null;
            try
            {
                content = File.ReadAllText(rdlPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error leyendo el archivo: {ex.Message}");
                return null;
            }

            // Usar una expresión regular para encontrar el texto entre las etiquetas <CommandText> y </CommandText>
            string pattern = @"<CommandText>(.*?)<\/CommandText>";
            Match match = Regex.Match(content, pattern, RegexOptions.Singleline);

            if (match.Success)
            {
                sqlQuery = match.Groups[1].Value.Trim();
                //Console.WriteLine("Consulta SQL extraída:");
                //Console.WriteLine(sqlQuery);
            }
            else
            {
                Console.WriteLine("No se encontró ninguna consulta SQL entre las etiquetas <CommandText>.");
            }
            return sqlQuery;
        }

        public DataTable DataReporte(string QUERY, ReportViewer parametros)
        {
            try
            {
                conexion = ObtenerConexion();
                if (conexion.State != System.Data.ConnectionState.Open)
                {
                    conexion.Open();
                }
                QUERY = QUERY.Replace("&gt;", ">").Replace("&lt;", "<");
                SqlCommand cmd = new SqlCommand(QUERY, conexion);
                var parameters = parametros.LocalReport.GetParameters();
                foreach (var param in parameters)
                {
                    cmd.Parameters.AddWithValue("@" + param.Prompt, param.Values[0].ToString());
                }
                SqlDataAdapter ad = new SqlDataAdapter(cmd);
                ds = new DataSet();
                ad.Fill(ds, "tabla");
                conexion.Close();
                //return ds;
                return ds.Tables["tabla"];
            }
            catch (Exception es)
            {
                conexion.Close();
                MessageBox.Show("Elementos no validos: " + es.Message);

                return null;
            }

        }

        public static System.Data.SqlClient.SqlConnection ObtenerConexion()
        {
            System.Data.SqlClient.SqlConnection Conexion = new System.Data.SqlClient.SqlConnection(@"Data Source=192.168.101.100; Initial Catalog= BARRON; MultipleActiveResultSets = true;User ID=sa; Password=Admin_sqlABG");
            return Conexion;
        }
    }
}
