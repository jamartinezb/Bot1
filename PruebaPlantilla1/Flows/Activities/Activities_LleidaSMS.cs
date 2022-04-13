using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Activities;
using System.ComponentModel;

using System.Net;
//using System.Net.Http.Headers;

using System.IO;

using RestSharp;
using Newtonsoft;

namespace LleidaSMS
{
    [Description("Enviar Mensaje de Texto mediante Lleida")]
    public class SendSMS : CodeActivity
    {
        [Category("Optional Parameters")]
        [DisplayName("Message Sender or DS")]
        [Description("Optional. Message sender or DS.")]
        public InArgument<String> src { get; set; }

        [Category("Optional Parameters")]
        [DisplayName("Data Coding")]
        [Description("Optional. Specifies data coding.")]
        public InArgument<String> data_coding { get; set; }

        [Category("Optional Parameters")]
        [DisplayName("Delivery receipt")]
        [Description("Optional. Activates delivery receipt or DR.")]
        public InArgument<String> delivery_receipt { get; set; }

        [Category("Optional Parameters")]
        [DisplayName("Allow Answer")]
        [Description("Optional. Activates the sending with long numerical sender.")]
        public InArgument<String> allow_answer { get; set; }

        [Category("Optional Parameters")]
        [DisplayName("User Id")]
        [Description("Optional. Specifies a unique ID for the message.")]
        public InArgument<String> user_id { get; set; }

        [Category("Optional Parameters")]
        [DisplayName("Schedule")]
        [Description("Optional. It specifies the exact time the message has to be delivered.")]
        public InArgument<String> schedule { get; set; }

        [Category("Parameters")]
        [DisplayName("User Name")]
        [RequiredArgument]
        [Description("Lleida.net user account name.")]
        public InArgument<String> user { get; set; }

        [Category("Parameters")]
        [DisplayName("Password")]
        [RequiredArgument]
        [Description("User password.")]
        public InArgument<String> password { get; set; }

        [Category("Parameters")]
        [DisplayName("Texto del mensaje")]
        [RequiredArgument]
        [Description("Text message.")]
        public InArgument<String> txt { get; set; }

        [Category("Parameters")]
        [DisplayName("Números destino")]
        [RequiredArgument]
        [Description("Contains one or more num elements, with the SMS recipient numbers.")]
        public InArgument<List<String>> dst { get; set; }

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
                String vStrsrc = src.Get(context);
                String vStrdata_coding = data_coding.Get(context);
                String vStrdelivery_receipt = delivery_receipt.Get(context);
                String vStrallow_answer = allow_answer.Get(context);
                String vStruser_id = user_id.Get(context);
                String vStrschedule = schedule.Get(context);
                String vStruser = user.Get(context);
                String vStrpassword = password.Get(context);
                String vStrtxt = txt.Get(context);
                List<String> vLstStrdst = dst.Get(context);

                // Crear objeto con cuerpo de mensaje
                requestSMS json = new requestSMS();
                // Escribir usuario
                json.sms.user = vStruser;
                // Escribir contraseña
                json.sms.password = vStrpassword;
                // Escribir src
                json.sms.src = "";// vStrsrc;
                // Escribir destinatarios
                json.sms.dst.num = vLstStrdst;
                // Escribir texto
                json.sms.txt = vStrtxt;
                // Escribir user_id
                json.sms.user_id = vStruser_id;
                // Escribir allow_answer
                json.sms.allow_answer = "+573186198601";

                String vStrjson = Newtonsoft.Json.JsonConvert.SerializeObject(json, Newtonsoft.Json.Formatting.Indented);
                vStrjson = vStrjson.Replace("allow_answer", "allow_answer/");
                Console.WriteLine(vStrjson);
                //var httpWebRequest = (HttpWebRequest)WebRequest.Create("http://api.lleida.net/sms/v2/");

                var client = new RestClient("http://api.lleida.net/sms/v2/");
                var request = new RestRequest(Method.POST);
                request.AddHeader("cache-control", "no-cache");
                request.AddHeader("Content-Type", "application/json");
                request.AddHeader("Accept", "application/json");
                request.AddParameter("undefined", vStrjson, ParameterType.RequestBody);
                IRestResponse response = client.Execute(request);

                responseSMS vObjResponse = Newtonsoft.Json.JsonConvert.DeserializeObject<responseSMS>(response.Content);

                //Console.WriteLine("request: " + vObjResponse.request);
                //Console.WriteLine("status: " + vObjResponse.status);
                //Console.WriteLine("code: " + vObjResponse.code.ToString());
                //Console.WriteLine("newcredit: " + vObjResponse.newcredit.ToString());

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

    // Objeto json de petición
    public class obj1
    {
        public List<string> num { get; set; }
    }

    public class obj2
    {
        public string user { get; set; }
        public string password{ get; set; }
        public string src { get; set; }
        public obj1 dst{ get; set; }
        public string txt{ get; set; }
        public string user_id { get; set; }
        public string allow_answer { get; set; }

        public obj2()
        {
            dst = new obj1();
        }
    }

    public class requestSMS
    {
        public obj2 sms{ get; set; }
        public requestSMS()
        {
            sms = new obj2();
        }
    }

    // Objeto json de respuesta
    public class responseSMS
    {
        public string request { get; set; }
        public string status { get; set; }
        public Int32 code { get; set; }
        public double newcredit { get; set; }
    }
}

