using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Activities;
using System.ComponentModel;

using System.Data;
using System.IO;

using RestSharp;

using System.Data.SqlClient;
using Newtonsoft.Json.Linq;
using System.Configuration;
using System.Xml;
using System.Xml.Linq;

// Librería de WinSCP
// Más información: https://winscp.net/eng/docs/library
// Tipo de licencia MPL 2.0
// Usada para LoadToSFTP
using WinSCP;

// Librería de Renci.SshNet
// Más información: https://github.com/sshnet/SSH.NET/tree/master
// Tipo de licencia MIT License
// Usada para LoadToSFTP
using Renci.SshNet;

namespace LleidaEmail
{

    [Description("Crea el encabezado requerido para cargar a Lleida mediante SFTP")]
    public class BuildHeaders : CodeActivity
    {
        public BuildHeaders()
        {
            CERTIFY = true;
            ATTACHID = 0;
        }
        [Category("Input")]
        [DisplayName("Scheduled")]
        [Description("Especificar en formato yyyymmddHHmmss. Dejar vacío para no programar envío.")]
        public InArgument<String> SCHEDULED { get; set; }

        [Category("Input")]
        [DisplayName("Template")]
        [Description("Nombre de la plantilla creada en Mailer.")]
        public InArgument<String> TEMPLATE { get; set; }

        [Category("Input")]
        [DisplayName("ID")]
        public InArgument<String> ID { get; set; }

        [Category("Input")]
        [DisplayName("Mail ID")]
        public InArgument<String> MAILID { get; set; }

        [Category("Input")]
        [DisplayName("Certify")]
        public Boolean CERTIFY { get; set; }

        [Category("Input")]
        [DisplayName("Attach")]
        [Description("Especificar el nombre de los archivos a adjuntar. Deben ir separados por el delimitador |.")]
        public InArgument<String> ATTACH { get; set; }

        [Category("Input")]
        [DisplayName("Attach ID")]
        [Description("Especificar el número de columna que contiene el nombre de archivo adjunto. Dejar cero si no se especifica valor.")]
        public InArgument<int> ATTACHID { get; set; }

        [Category("Output")]
        [DisplayName("Data Table")]
        [Description("Respuesta en formato DataTable.")]
        public OutArgument<DataTable> DataTable { get; set; }

        [Category("zException Handling")]
        [RequiredArgument]
        [DisplayName("Exception Handling")]
        [Description("True si desea manejar las excepciones manualmente.")]
        public Boolean ExceptionHandling { get; set; }

        [Category("zException Handling")]
        [DisplayName("Resultado")]
        [Description("Resultado de la ejecución.")]
        public OutArgument<bool> Result { get; set; }

        [Category("zException Handling")]
        [DisplayName("What")]
        [Description("Mensaje de Error en caso de tenerlo.")]
        public OutArgument<string> What { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            try
            {
                //Capturar o crear parametros
                System.Data.DataTable vDtbEncabezados = new DataTable();
                vDtbEncabezados.Columns.Add("Encabezados");

                // Si se especifica envío programado
                if (!String.IsNullOrWhiteSpace(SCHEDULED.Get(context)))
                {
                    vDtbEncabezados.Rows.Add("#SCHEDULED:" + SCHEDULED.Get(context));
                }
                // Escribir el nombre de la plantilla
                vDtbEncabezados.Rows.Add("#TEMPLATE:" + TEMPLATE.Get(context));
                // Escribir el ID de envío
                vDtbEncabezados.Rows.Add("#ID:" + ID.Get(context));
                // Escribir el MAILID
                vDtbEncabezados.Rows.Add("#MAILID:" + MAILID.Get(context));
                // Validar si es Email Certificado
                if (CERTIFY)
                {
                    vDtbEncabezados.Rows.Add("#CERTIFY");
                }
                // Validar si se especifica el nombre de archivo adjunto
                if (String.IsNullOrWhiteSpace(ATTACH.Get(context)))
                {
                    vDtbEncabezados.Rows.Add("#ATTACH:" + ATTACH.Get(context));
                }
                // Validar si se especifica ID de documento adjunto
                if (ATTACHID.Get(context) > 0)
                {
                    vDtbEncabezados.Rows.Add("#ATTACHID:" + ATTACHID.Get(context));
                }
                // Asignar variables de salida
                DataTable.Set(context, vDtbEncabezados);
                Result.Set(context, true);
            }
            catch (Exception e)
            {
                string error = "Mensaje Error: " + e.Message + Environment.NewLine + "InnerException: " + e.InnerException;
                What.Set(context, error);
                Result.Set(context, false);

                Boolean vExceptionHandling = ExceptionHandling;
                if (vExceptionHandling == false)
                {
                    throw e;
                }

            }

        }
    }

    [Description("Cargar archivo a Lleida mediante SFTP")]
    public class LoadToSFTP : CodeActivity
    {
        [Category("HostOptions")]
        [DisplayName("Host Name")]
        public InArgument<String> HostName { get; set; }

        [Category("HostOptions")]
        [DisplayName("Port Number")]
        public InArgument<int> PortNumber { get; set; }

        [Category("Authentication")]
        [DisplayName("User Name")]
        public InArgument<String> UserName { get; set; }

        [Category("Authentication")]
        [DisplayName("Password")]
        public InArgument<String> Password { get; set; }

        [Category("Input")]
        [DisplayName("File Path")]
        public InArgument<String> FilePath { get; set; }

        [Category("Input")]
        [DisplayName("Target Directory")]
        [Description("Especificar en formato /rootDirectory/Directory/...")]
        public InArgument<String> TargetDirectory { get; set; }

        [Category("zException Handling")]
        [RequiredArgument]
        [DisplayName("Exception Handling")]
        [Description("True si desea manejar las excepciones manualmente.")]
        public Boolean ExceptionHandling { get; set; }

        [Category("zException Handling")]
        [DisplayName("Resultado")]
        [Description("Resultado de la ejecución.")]
        public OutArgument<bool> Result { get; set; }

        [Category("zException Handling")]
        [DisplayName("What")]
        [Description("Mensaje de Error en caso de tenerlo.")]
        public OutArgument<string> What { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            try
            {
                // Obtener argumentos
                String vStrHostName = HostName.Get(context);
                int vIntPortNumber = PortNumber.Get(context);
                String vStrUserName = UserName.Get(context);
                String vStrPassword = Password.Get(context);
                String vStrFilePath = FilePath.Get(context);
                String vStrTargetDirectory = TargetDirectory.Get(context);

            // Ejecutar tarea
                RenciSSHNet(vStrHostName, vIntPortNumber, vStrUserName, vStrPassword, vStrFilePath, vStrTargetDirectory);

                Result.Set(context, true);
            }
            catch (Exception e)
            {
                string error = "Mensaje Error: " + e.Message + Environment.NewLine + "InnerException: " + e.InnerException;
                What.Set(context, error);
                Result.Set(context, false);

                Boolean vExceptionHandling = ExceptionHandling;
                if (vExceptionHandling == false)
                {
                    throw e;
                }

            }

        }

        public void WinSCP(String Host, int Port, String User, String Pass, String FilePath, String TargetDirectory)
        {
            // Setup session options
            SessionOptions sessionOptions = new SessionOptions
            {
                Protocol = Protocol.Sftp,
                HostName = Host,
                UserName = User,
                Password = Pass,
                SshHostKeyFingerprint = "ssh-rsa 2048 xxxxxxxxxxx...="
            };

            using (WinSCP.Session session = new WinSCP.Session())
            {
                // Connect
                session.Open(sessionOptions);

                // Upload files
                session.PutFiles(FilePath, TargetDirectory).Check();
            }
        }

        public static void RenciSSHNet(String Host, int Port, String User, String Pass, String FilePath, String TargetDirectory)
        {
            Console.WriteLine("Creating client and connecting");
            using (var client = new SftpClient(Host, Port, User, Pass))
            {
                client.Connect();
                Console.WriteLine("Connected to {0}", Host);

                client.ChangeDirectory(TargetDirectory);
                Console.WriteLine("Changed directory to {0}", TargetDirectory);

                var listDirectory = client.ListDirectory(TargetDirectory);
                Console.WriteLine("Listing directory:");
                foreach (var fi in listDirectory)
                {
                    Console.WriteLine(" - " + fi.Name);
                }

                using (var fileStream = new FileStream(FilePath, FileMode.Open))
                {
                    Console.WriteLine("Uploading {0} ({1:N0} bytes)", FilePath, fileStream.Length);
                    client.BufferSize = 4 * 1024; // bypass Payload error large files
                    client.UploadFile(fileStream, Path.GetFileName(FilePath));
                }
            }
        }
    }

    [Description("Consumo API mailcertapi.cgi para capturar resultado de Mails enviados por Lleida")]
    public class ReadLastListPDF : CodeActivity
    {
        public ReadLastListPDF()
        {

        }

        [Category("Input")]
        [RequiredArgument]
        [DisplayName("User")]
        [Description("Usuario para consumir el servicio")]
        public InArgument<string> user { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [DisplayName("Password")]
        [Description("Contraseña para consumir el servicio")]
        public InArgument<string> password { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [DisplayName("Fecha Minima")]
        [Description("Fecha desde donde se van a traer registros: AAAAMMDDHHMMSS")]
        public InArgument<string> fechaMin { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [DisplayName("Fecha Maxima")]
        [Description("Fecha hasta donde se van a traer registros: AAAAMMDDHHMMSS")]
        public InArgument<string> fechaMax { get; set; }

        [Category("Output")]
        [DisplayName("Respuesta API")]
        [Description("Respuesta API")]
        public OutArgument<DataTable> Respuesta { get; set; }

        [Category("zException Handling")]
        [RequiredArgument]
        [DisplayName("Exception Handling")]
        [Description("True si desea manejar las excepciones manualmente.")]
        public Boolean ExceptionHandling { get; set; }

        [Category("zException Handling")]
        [DisplayName("Resultado")]
        [Description("Resultado de la ejecución.")]
        public OutArgument<bool> Result { get; set; }

        [Category("zException Handling")]
        [DisplayName("What")]
        [Description("Mensaje de Error en caso de tenerlo.")]
        public OutArgument<string> What { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            try
            {
                // Obtener argumentos
                String vUser = user.Get(context);
                String vPass = password.Get(context);
                String vMailDateMin = fechaMin.Get(context);
                String vMailDateMax = fechaMax.Get(context);

                //Crear Datatable
                DataTable dtMails = new DataTable();
                dtMails.Columns.Add("mail_id");
                dtMails.Columns.Add("gstatus");
                dtMails.Columns.Add("mail_date");
                dtMails.Columns.Add("mail_to");
                dtMails.Columns.Add("mail_from");
                dtMails.Columns.Add("file_id");
                dtMails.Columns.Add("file_status");
                dtMails.Columns.Add("file_name");
                dtMails.Columns.Add("mail_subj");
                dtMails.Columns.Add("LlaveRegistro");

                var client = new RestClient("https://tsa.lleida.net/cgi-bin/mailcertapi.cgi");
                client.Timeout = -1;
                var request = new RestRequest(Method.POST);
                request.AddHeader("Content-Type", "application/x-www-form-urlencoded");
                request.AddHeader("Accept", "application/xml");
                request.AddParameter("action", "list_pdf");
                request.AddParameter("user", vUser);
                request.AddParameter("password", vPass);
                request.AddParameter("mail_date_min", vMailDateMin);
                request.AddParameter("mail_date_max", vMailDateMax);//20200317231805
                IRestResponse response;
                response = client.Execute(request);

                var xDoc = XDocument.Parse(response.Content);

                if (response.StatusCode == 0)
                {
                   throw new System.ArgumentException("Web Service   Down", "Reponse Web Service ");
                }
                else
                {

                    // Status
                    var Status = xDoc.Descendants("status").Single();
                    //Console.WriteLine("status Consumo");
                    //Console.WriteLine(Status.Value);
                    //Msg
                    var Msg = xDoc.Descendants("msg").Single();
                    //Console.WriteLine("msg");
                    //Console.WriteLine(Msg.Value);

                    //rows_found
                    var response2 = xDoc.Descendants("pdf_list").Single();
                    var attr = response2.Attribute("rows_found");
                    //Console.WriteLine("rows_found");
                    //Console.WriteLine(attr.Value);
                    string rowsFound = attr.Value;

                    //rows_ok
                    attr = response2.Attribute("rows_ok");
                    //Console.WriteLine("rows_ok");
                    //Console.WriteLine(attr.Value);
                    string rowsOK = attr.Value;

                    //rows_ok
                    attr = response2.Attribute("rows_ko");
                    //Console.WriteLine("rows_KO");
                    //Console.WriteLine(attr.Value);
                    string rowsKO = attr.Value;

                
                    //Obtener Pdf_List
                    var pdf_list = xDoc.Root.Element("pdf_list");



                    //Recorrer cada Pdf_row
                    foreach (var pdfList in pdf_list.Descendants("pdf_row"))
                    {
                        DataRow newRow = dtMails.NewRow(); 
                        newRow[0] = pdfList.Element("mail_id").Value;
                        newRow[1] = pdfList.Element("gstatus").Value;
                        newRow[2] = pdfList.Element("mail_date").Value;
                        newRow[3] = pdfList.Element("mail_to").Value;
                        newRow[4] = pdfList.Element("mail_from").Value;
                        try
                        {
                            newRow[5] = pdfList.Element("file_id").Value;
                        }
                        catch { }

                        try
                        {
                            newRow[6] = pdfList.Element("file_status").Value;
                        }
                        catch { }

                        try
                        {
                            newRow[7] = pdfList.Element("file_name").Value;
                        }
                        catch { }
                    
                        newRow[8] = pdfList.Element("mail_subj").Value;
                        newRow[9] = "Llave" + pdfList.Element("mail_subj").Value;

                    
                        dtMails.Rows.Add(newRow);
                    }

                }

                // Establecer argumentos de salida
                Respuesta.Set(context, dtMails);
            }
            catch (Exception e)
            {
                string error = "Mensaje Error: " + e.Message + Environment.NewLine + "InnerException: " + e.InnerException;
                What.Set(context, error);
                Result.Set(context, false);

                Boolean vExceptionHandling = ExceptionHandling;
                if (vExceptionHandling == false)
                {
                    throw e;
                }

            }          

        }


    }

    // Tipos de variables creadas
    public class EmailCertificadoReport
    {
        public bool success { get; set; }
        public int httpResponseCode { get; set; }
        public String error { get; set; }
        public string message { get; set; }
        public IList<string> data { get; set; }
        public string fechaActual { get; set; }
    }

    class LoadSFTP : CodeActivity
    {

        protected override void Execute(CodeActivityContext context)
        {
            throw new NotImplementedException();
        }
    }

}