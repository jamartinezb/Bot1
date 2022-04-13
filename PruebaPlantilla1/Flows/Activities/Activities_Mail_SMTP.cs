using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Activities;
using System.ComponentModel;

using System.Net.Mail;
using System.Net;
using System.Text.RegularExpressions;

namespace Mail_SMTP
{

    /// <summary>
    /// Descripción: Enviar Mail
    /// Autor: Carlos Cardona
    /// Fecha Ultima Versión: 2020-06-16
    /// </summary>
    [Description("Enviar mails a multiples correos al tiempo")]
    public class Mail_SMTP : CodeActivity
    {
        public Mail_SMTP()
        {
            EmailHTMLbody = false;
            EnableSSL = true;
        }

        Regex RegexEmail = new Regex(@"^([\w\.\-]+)@([\w\-]+)((\.(\w){2,})+)$");
        Regex RegexPath = new Regex(@"^(?:[a-zA-Z]\:|\\\\[\w\.]+\\[\w.$]+)\\(?:[\w]+\\)*\w([\w.])+$");

        [Category("Sender details")]
        [Description("email desde donde se enviará el correo.")]
        [RequiredArgument]
        [DisplayName("Sender email")]
        public InArgument<String> SenderEmail { get; set; }

        [Category("Sender details")]
        [Description("contraseña email de envio.")]
        [RequiredArgument]
        [DisplayName("Sender password")]
        public InArgument<String> SenderPassword { get; set; }

        [Category("Sender details")]
        [Description("Server SMTP. Depende del proveedor Gmail, hotmail, etc.")]
        [RequiredArgument]
        [DisplayName("SMTP server")]
        public InArgument<String> SMTPServer { get; set; }

        [Category("Sender details")]
        [Description("Número de puerto. Por defecto 587.")]
        [RequiredArgument]
        [DisplayName("Port number")]
        public InArgument<int> SMTPPortNumber { get; set; }

        [Category("Receiver details")]
        [Description("emails de destino. Separados por punto y coma (;)")]
        [RequiredArgument]
        [DisplayName("To")]
        public InArgument<String> ReceiverTo { get; set; }

        [Category("Receiver details")]
        [Description("e-mails de destino en copia. Separados por punto y coma (;)")]
        [DisplayName("CC")]
        public InArgument<String> ReceiverCC { get; set; }

        [Category("Receiver details")]
        [Description("e-mail de destino ocultos. Separados por punto y coma (;)")]
        [DisplayName("BCC")]
        public InArgument<String> ReceiverBCC { get; set; }

        [Category("Email")]
        [Description("Asunto del correo")]
        [RequiredArgument]
        [DisplayName("Subject")]
        public InArgument<String> EmailSubject { get; set; }

        [Category("Email")]
        [Description("Cuerpo del correo")]
        [RequiredArgument]
        [DisplayName("Body")]
        public InArgument<String> EmailBody { get; set; }

        [Category("Email")]
        [DisplayName("Attachments")]
        [Description("Especificar las rutas de adjuntos separadas por el caracter punto y coma")]
        public InArgument<String> EmailAttachements { get; set; }

        [Category("Options")]
        [Description("Indica si el cuerpo del correo esta en formato HTML")]
        [RequiredArgument]
        [DisplayName("HTML body")]
        public Boolean EmailHTMLbody { get; set; }

        [Category("Options")]
        [Description("habilitar SSL")]
        [RequiredArgument]
        [DisplayName("Enable SSL")]
        public Boolean EnableSSL { get; set; }

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
                // Obtener Variables
                String To = ReceiverTo.Get(context);
                String Cc = ReceiverCC.Get(context);
                String Bcc = ReceiverBCC.Get(context);


                //// Determine when the activity has been configured in an invalid way.  
                //if (!RegexEmail.IsMatch(this.SenderEmail.Get(context)))
                //{
                //    // Add a validation error with a custom message.
                //    throw new System.Exception("Write a valid sender email");
                //}
                //if (RegexEmail.Match(this.ReceiverTo.Get(context)).Captures.Count == 0)
                //{
                //    // Add a validation error with a custom message.  
                //    throw new System.Exception("Write at least one valid To email");
                //}
                //if (RegexEmail.Match(this.ReceiverCC.Get(context)).Groups[0].Captures.Count == 0 && !String.IsNullOrWhiteSpace(this.ReceiverCC.Get(context)))
                //{
                //    // Add a validation error with a custom message.  
                //    throw new System.Exception("Write at least one valid CC email");
                //}
                //if (RegexEmail.Match(this.ReceiverBCC.Get(context)).Groups[0].Captures.Count == 0 && !String.IsNullOrWhiteSpace(this.ReceiverBCC.Get(context)))
                //{
                //    // Add a validation error with a custom message.  
                //    throw new System.Exception("Write at least one valid BCC email");
                //}

                //Crear Correo
                MailMessage msg = new MailMessage();
                msg.From = new MailAddress(SenderEmail.Get(context));
                // Agregar destinatarios
                foreach (var Email in To.Split(';'))
                {
                    if (Email != "")
                    {
                        msg.To.Add(new MailAddress(Email));
                    }
                }
                // Agregar destinatarios con copia
                foreach (var Email in Cc.Split(';'))
                {
                    if (Email != "")
                    {
                        msg.CC.Add(new MailAddress(Email));
                    }
                }
                // Agregar destinatarios con copia oculta
                foreach (var Email in Bcc.Split(';'))
                {
                    if (Email != "")
                    {
                        msg.Bcc.Add(new MailAddress(Email));
                    }
                }
                // Especificar si el cuerpo del correo es HTML
                msg.IsBodyHtml = EmailHTMLbody;
                // Especificar el asunto del correo
                msg.Subject = EmailSubject.Get(context);
                // Especificar el cuerpo del correo
                msg.Body = EmailBody.Get(context);
                // Agregar adjuntos
                foreach (Capture Att in RegexPath.Match(this.EmailAttachements.Get(context)).Groups[0].Captures)
                {
                    msg.Attachments.Add(new Attachment(Att.Value));
                }

                SmtpClient client = new SmtpClient();
                client.UseDefaultCredentials = true;
                client.Host = this.SMTPServer.Get(context);
                client.Port = this.SMTPPortNumber.Get(context);
                client.EnableSsl = this.EnableSSL;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.Credentials = new NetworkCredential(this.SenderEmail.Get(context), this.SenderPassword.Get(context));
                client.Timeout = 20000;

                //Enviar mensaje
                client.Send(msg);
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

    [Description("Enviar mails con verificación de correo electronico a 1 solo correo a la vez")]
    public class Mail_SMTP_Verificacion1correo : CodeActivity
    {
        public Mail_SMTP_Verificacion1correo()
        {
            EmailHTMLbody = false;
            EnableSSL = true;
        }

        Regex RegexEmail = new Regex(@"^([\w\.\-]+)@([\w\-]+)((\.(\w){2,})+)$");
        Regex RegexPath = new Regex(@"^(?:[a-zA-Z]\:|\\\\[\w\.]+\\[\w.$]+)\\(?:[\w]+\\)*\w([\w.])+$");

        [Category("Sender details")]
        [RequiredArgument]
        [DisplayName("Sender email")]
        public InArgument<String> SenderEmail { get; set; }

        [Category("Sender details")]
        [RequiredArgument]
        [DisplayName("Sender password")]
        public InArgument<String> SenderPassword { get; set; }

        [Category("Sender details")]
        [RequiredArgument]
        [DisplayName("SMTP server")]
        public InArgument<String> SMTPServer { get; set; }

        [Category("Sender details")]
        [RequiredArgument]
        [DisplayName("Port number")]
        public InArgument<int> SMTPPortNumber { get; set; }

        [Category("Receiver details")]
        [RequiredArgument]
        [DisplayName("To")]
        public InArgument<String> ReceiverTo { get; set; }

        [Category("Receiver details")]
        [DisplayName("CC")]
        public InArgument<String> ReceiverCC { get; set; }

        [Category("Receiver details")]
        [DisplayName("BCC")]
        public InArgument<String> ReceiverBCC { get; set; }

        [Category("Email")]
        [RequiredArgument]
        [DisplayName("Subject")]
        public InArgument<String> EmailSubject { get; set; }

        [Category("Email")]
        [RequiredArgument]
        [DisplayName("Body")]
        public InArgument<String> EmailBody { get; set; }

        [Category("Email")]
        [DisplayName("Attachments")]
        [Description("Especificar las rutas de adjuntos separadas por el caracter punto y coma")]
        public InArgument<String> EmailAttachements { get; set; }

        [Category("Options")]
        [RequiredArgument]
        [DisplayName("HTML body")]
        public Boolean EmailHTMLbody { get; set; }

        [Category("Options")]
        [RequiredArgument]
        [DisplayName("Enable SSL")]
        public Boolean EnableSSL { get; set; }

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
                // Agregar validaciones

                // Determine when the activity has been configured in an invalid way.  
                if (!RegexEmail.IsMatch(this.SenderEmail.Get(context)))
                {
                    // Add a validation error with a custom message.
                    throw new System.Exception("Write a valid sender email");
                }
                if (RegexEmail.Match(this.ReceiverTo.Get(context)).Captures.Count == 0)
                {
                    // Add a validation error with a custom message.  
                    throw new System.Exception("Write at least one valid To email");
                }
                if (RegexEmail.Match(this.ReceiverCC.Get(context)).Groups[0].Captures.Count == 0 && !String.IsNullOrWhiteSpace(this.ReceiverCC.Get(context)))
                {
                    // Add a validation error with a custom message.  
                    throw new System.Exception("Write at least one valid CC email");
                }
                if (RegexEmail.Match(this.ReceiverBCC.Get(context)).Groups[0].Captures.Count == 0 && !String.IsNullOrWhiteSpace(this.ReceiverBCC.Get(context)))
                {
                    // Add a validation error with a custom message.  
                    throw new System.Exception("Write at least one valid BCC email");
                }

                MailMessage msg = new MailMessage();
                msg.From = new MailAddress(SenderEmail.Get(context));
                // Agregar destinatarios
                foreach (Capture Email in RegexEmail.Match(ReceiverTo.Get(context)).Groups[0].Captures)
                {
                    msg.To.Add(new MailAddress(Email.Value));
                }
                // Agregar destinatarios con copia
                foreach (Capture Email in RegexEmail.Match(ReceiverCC.Get(context)).Groups[0].Captures)
                {
                    msg.CC.Add(new MailAddress(Email.Value));
                }
                // Agregar destinatarios con copia oculta
                foreach (Capture Email in RegexEmail.Match(ReceiverBCC.Get(context)).Groups[0].Captures)
                {
                    msg.Bcc.Add(new MailAddress(Email.Value));
                }
                // Especificar si el cuerpo del correo es HTML
                msg.IsBodyHtml = EmailHTMLbody;
                // Especificar el asunto del correo
                msg.Subject = EmailSubject.Get(context);
                // Especificar el cuerpo del correo
                msg.Body = EmailBody.Get(context);
                // Agregar adjuntos
                // MR = ajustes porque no funciona con multiple archivos (FOR que esta 76 linaes mas abajo) 
                // foreach (Capture Att in RegexPath.Match(this.EmailAttachements.Get(context)).Groups[0].Captures)
                //
                string vCadena = EmailAttachements.Get(context);
                string[] partsCadenaString = vCadena.Split(new string[] { ";" }, StringSplitOptions.None);
                //
                foreach (var fileAtt in partsCadenaString)
                {
                    msg.Attachments.Add(new Attachment(fileAtt));
                }

                SmtpClient client = new SmtpClient();
                client.UseDefaultCredentials = true;
                client.Host = this.SMTPServer.Get(context);
                client.Port = this.SMTPPortNumber.Get(context);
                client.EnableSsl = this.EnableSSL;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.Credentials = new NetworkCredential(this.SenderEmail.Get(context), this.SenderPassword.Get(context));
                client.Timeout = 20000;

                //enviar mensaje
                client.Send(msg);

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


}
