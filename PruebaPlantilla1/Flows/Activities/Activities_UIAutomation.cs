using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Activities;
using System.ComponentModel;

using System.Runtime.InteropServices;
using System.Diagnostics;

using SelectorManager;
using System.Threading; // Utilizado para agregar delays

namespace UIAutomation
{
    [Description("Dar click a un elemento mediante libreria UIAutomation")]
    public class Click : CodeActivity
    {
        public Click()
        {
            DelayAfter = 300;
            DelayBefore = 1000;
        }

        [Category("Common")]
        [DisplayName("Tiempo de espera al finalizar")]
        [Description("Cantidad de tiempo (en milisegundos) después de ejecutar la actividad.")]
        [RequiredArgument]
        public InArgument<Int32> DelayAfter { get; set; }


        [Category("Common")]
        [DisplayName("Tiempo de espera al Iniciar")]
        [Description("Cantidad de tiempo (en milisegundos) antes de ejecutar la actividad.")]
        [RequiredArgument]
        public InArgument<int> DelayBefore { get; set; }


        [Category("Input")]
        [RequiredArgument]
        public InArgument<String> Selector { get; set; }

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
                String vStrSelector = Selector.Get(context);
                int vDelayAfter = DelayAfter.Get(context);
                int vDelayBefore = DelayBefore.Get(context);

                // Tiempo de espera antes de iniciar la actividad
                Thread.Sleep(vDelayBefore);

                //capturar selector
                Dictionary<string, UiSelector> vDicSelector = SelectorManager.Funciones.TransformIntoUiSelector(Selector.Get(context));
                //this is a constant indicating the window that we want to send a text message
                const int BM_CLICK = 0x00F5;

                //getting application window
                IntPtr vWndAplicacion = LibWrap.FindWindowA(vDicSelector["0"].Attribute.className, vDicSelector["0"].Attribute.title);

                Console.WriteLine("Llaves en el diccionario " + vDicSelector.Count.ToString());
                //getting application's textbox handle from the main window's handle

                IntPtr Button = new IntPtr();

                for (int i = 1; i < vDicSelector.Count; i++)
                {
                    Console.WriteLine("Iteration number " + i.ToString());
                    //the field is called 'Button'
                    Button = LibWrap.FindWindowEx(vWndAplicacion, IntPtr.Zero, vDicSelector[i.ToString()].Attribute.className, vDicSelector[i.ToString()].Attribute.title);

                    if (Button != null)
                    {
                        Console.WriteLine("No es nulo " + Button.ToString());

                    }
                }

                Console.WriteLine("Click button");

                LibWrap.SendMessage(Button, BM_CLICK, "", "");

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

    [Description("Escribir texto dentro de un elemento mediante libreria UIAutomation")]
    public class TypeInto : CodeActivity
    {
        public TypeInto(){
            DelayAfter = 300;
            DelayBefore = 1000;
            TimeoutMS = 30000;
        }

        [Category("Common")]
        [DisplayName("Tiempo de espera al finalizar")]
        [Description("Cantidad de tiempo (en milisegundos) después de ejecutar la actividad.")]
        [RequiredArgument]
        public InArgument<Int32> DelayAfter { get; set; }


        [Category("Common")]
        [DisplayName("Tiempo de espera al Iniciar")]
        [Description("Cantidad de tiempo (en milisegundos) antes de ejecutar la actividad.")]
        [RequiredArgument]
        public InArgument<int> DelayBefore { get; set; }

        [Category("Input")]
        [DisplayName("Selector")]
        [Description("Selector del elemento al cual se le aplicará la acción.")]
        [RequiredArgument]
        public InArgument<String> Selector { get; set; }

        [Category("Input")]
        [DisplayName("Texto")]
        [Description("Texto que se va a escribir en el elemento.")]
        public InArgument<String> Text { get; set; }

        [Category("Input")]
        [DefaultValue(true)]
        [DisplayName("Selector")]
        [Description("Cantidad de tiempo (en milisegundos) antes de ejecutar la actividad.")]
        public InArgument<Int32> TimeoutMS { get; set; }

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
                String vStrSelector = Selector.Get(context);
                String vStrText = Text.Get(context);
                int vDelayAfter = DelayAfter.Get(context);
                int vDelayBefore = DelayBefore.Get(context);

                // Tiempo de espera antes de iniciar la actividad
                Thread.Sleep(vDelayBefore);


                Dictionary<string, UiSelector> vDicSelector = SelectorManager.Funciones.TransformIntoUiSelector(Selector.Get(context));
                //this is a constant indicating the window that we want to send a text message
                const int WM_SETTEXT = 0X000C;
                const int WM_GETTEXT = 0x000D;

                //getting application window
                IntPtr vWndAplicacion = LibWrap.FindWindowA(vDicSelector["0"].Attribute.className, vDicSelector["0"].Attribute.title);


                //getting application's textbox handle from the main window's handle
                //the field is called 'Edit'
                //IntPtr Textbox = LibWrap.FindWindowEx(vWndAplicacion, IntPtr.Zero, vDicSelector["1"].Attribute.className, vDicSelector["1"].Attribute.title);

                IntPtr Textbox = new IntPtr();

                for (int i = 1; i < vDicSelector.Count; i++)
                {
                    Console.WriteLine("Iteration number " + i.ToString());
                    //the field is called 'Edit'
                    Textbox = LibWrap.FindWindowEx(vWndAplicacion, IntPtr.Zero, vDicSelector[i.ToString()].Attribute.className, vDicSelector[i.ToString()].Attribute.title);
                    if (Textbox != null)
                    {
                        Console.WriteLine("No es nulo " + Textbox.ToString());

                    }
                }

                //sending the message to the textbox
                StringBuilder title = new StringBuilder(100);
                LibWrap.SendMessage(Textbox, WM_GETTEXT, title.Capacity, title);
                Console.WriteLine("title: " + title);
                LibWrap.SendMessage(Textbox, WM_SETTEXT, "0", vStrText);

                Result.Set(context, true);

                // Tiempo de espera después de finalizada la actividad
                Thread.Sleep(vDelayAfter);
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

    public class LibWrap
    {
        //include FindWindowA
        [DllImport("user32.dll")]
        public static extern IntPtr FindWindowA(string lpszClass, string lpszWindowName);

        //include FindWindowEx
        [DllImport("user32.dll")]
        public static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        //include SendMessage
        [DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, uint uMsg, string wParam, string lParam);

        //[DllImport("user32.dll", SetLastError = true)]
        //public static extern IntPtr SendMessage(int hWnd, int Msg, int wparam, int lparam);

        [DllImport("user32.dll", EntryPoint = "SendMessage", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        public static extern bool SendMessage(IntPtr hWnd, uint Msg, int wParam, StringBuilder lParam);
    }
}