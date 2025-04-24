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


namespace ReportsEjecute
{
    public class CreateReportsToPDF
    {
        private ReportViewer ReportViewerCreations { get; set; }
        private System.Data.SqlClient.SqlConnection _Conexion { get; set; }
        private DataSet _DatasetCreations { get; set; }
        private static string _Company { get; set; }

        static void Main(string[] args)
        {
            CreateReportsToPDF createReportsToPDF = new CreateReportsToPDF();
            // Validar parámetros
            if (args.Length < 11)
            {
                Console.WriteLine("Entraste a modo compilado");
                string[] parameters =
                {
                    "Z:\\RDLReportes\\RC_-_Balanza_catorcena.rdl",
                    "1",
                    "reporte1",
                    "2",
                    "2024",
                    "1",
                    "0",
                    "true",
                    Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Temp", "ScreenshotsReporte.pdf"),
                    "pdf",
                    "BARRON"
                };
                args = parameters;
            }
            
            string path = args[0];
            string ReCodigo = args[1];
            string ReNombre = args[2];
            string peTipo = args[3];
            string peAnio = args[4];
            string peNumero = args[5];
            string cbCodigo = args[6];
            bool Activo = bool.Parse(args[7]);
            string outputFilePath = args[8];
            string format = args[9];
            _Company = args[10];

            // Llamar a la función existente
            string result = createReportsToPDF.CreationsReports(path, ReCodigo, ReNombre, peTipo, peAnio, peNumero, cbCodigo, Activo, null, outputFilePath, format);

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

        public string CreationsReports(string path, string ReCodigo, string ReNombre, string peTipo, string peAnio, string peNumero, string colabora, bool Activo, DataTable dataTable, string outputFilePath, string format)
        {
            try
            {
                ReportViewerCreations = new ReportViewer();
                string activo = "Si";
                ReportViewerCreations.LocalReport.ReportPath = path;
                if (File.Exists(path))
                {
                    // El reporte está listo para abrirse
                    activo = Activo ? "S" : "N";

                    List<ReportParameter> parametros = new List<ReportParameter>
                    {
                        new ReportParameter("EMPRESA", _Company),
                        new ReportParameter("REPORTE", ReNombre),
                        new ReportParameter("ACTIVO", activo),
                        new ReportParameter("AÑO", peAnio),
                        new ReportParameter("PERIODO", peNumero),
                        new ReportParameter("TIPO", peTipo),
                        new ReportParameter("EMPLEADO", colabora)
                    };

                    var validParameters = ReportViewerCreations.LocalReport.GetParameters();

                    // Filtrar y establecer solo los parámetros que están definidos en el informe
                    var parametersToSet = parametros
                        .Where(p => validParameters.Any(vp => vp.Name.Equals(p.Name, StringComparison.OrdinalIgnoreCase)))
                        .ToList();

                    if (parametersToSet.Any())
                    {
                        ReportViewerCreations.LocalReport.SetParameters(parametersToSet);
                    }

                    var parameters = ReportViewerCreations.LocalReport.GetParameters();
                    var dataSourceNames = ReportViewerCreations.LocalReport.GetDataSourceNames();

                    ReportDataSource reportDataSource = new ReportDataSource(dataSourceNames[0].ToString(), DataReporte(GetSqlQueryFromRdl(path), ReportViewerCreations));
                    if (reportDataSource.Value != null)
                    {
                        ReportViewerCreations.LocalReport.DataSources.Clear();
                        ReportViewerCreations.LocalReport.DataSources.Add(reportDataSource);

                        Warning[] warnings;
                        string[] streamIds;
                        string mimeType = string.Empty;
                        string encoding = string.Empty;
                        string extension = string.Empty;
                        byte[] bytes;
                        // Renderiza el reporte en el formato especificado
                        if (format == "pdf")
                        {
                            bytes = ReportViewerCreations.LocalReport.Render(
                               "PDF", // Formato de exportación
                               null, // DeviceInfo (puede ser null para usar configuración por defecto)
                               out mimeType,
                               out encoding,
                               out extension,
                               out streamIds,
                               out warnings
                            );
                        }
                        else
                        {
                            bytes = ReportViewerCreations.LocalReport.Render(
                               "Excel", // Formato de exportación
                               null, // DeviceInfo (puede ser null para usar configuración por defecto)
                               out mimeType,
                               out encoding,
                               out extension,
                               out streamIds,
                               out warnings
                            );
                        }

                        // Guarda el archivo en el sistema de archivos
                        using (FileStream fs = new FileStream(outputFilePath, FileMode.Create))
                        {
                            fs.Write(bytes, 0, bytes.Length);
                        }

                        return outputFilePath;
                    }
                }
                else
                {
                    returnMessage("El archivo de reporte no existe.", outputFilePath);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, outputFilePath);
                returnMessage(ex.Message, outputFilePath);
                return " ";
            }
            return " ";
        }

        public void returnMessage(string message, string outputFilePath)
        {
            string outputDirectory = Path.GetDirectoryName(outputFilePath);
            string outputPath = Path.Combine(outputDirectory, "resultado.txt");
            string directoryPath = Path.GetDirectoryName(outputPath);

            if (!Directory.Exists(directoryPath))
            {
                Directory.CreateDirectory(directoryPath);
            }

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
                _Conexion = ObtenerConexion();
                if (_Conexion.State != System.Data.ConnectionState.Open)
                {
                    _Conexion.Open();
                }
                QUERY = QUERY.Replace("&gt;", ">").Replace("&lt;", "<");
                SqlCommand cmd = new SqlCommand(QUERY, _Conexion);
                var parameters = parametros.LocalReport.GetParameters();
                foreach (var param in parameters)
                {
                    cmd.Parameters.AddWithValue("@" + param.Prompt, param.Values[0].ToString());
                }
                SqlDataAdapter ad = new SqlDataAdapter(cmd);
                _DatasetCreations = new DataSet();
                ad.Fill(_DatasetCreations, "tabla");
                _Conexion.Close();
                //return ds;
                return _DatasetCreations.Tables["tabla"];
            }
            catch (Exception es)
            {
                _Conexion.Close();
                MessageBox.Show("Elementos no validos: " + es.Message + QUERY);

                return null;
            }

        }

        public static System.Data.SqlClient.SqlConnection ObtenerConexion()
        {
            System.Data.SqlClient.SqlConnection Conexion = new System.Data.SqlClient.SqlConnection(@"Data Source=192.168.101.100; Initial Catalog=" + _Company + " ; MultipleActiveResultSets = true;User ID=sa; Password=Admin_sqlABG");
            return Conexion;
        }
    }
}
