using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Activities;
using System.ComponentModel;

using SelectorManager;

using System.Threading; // Utilizado para agregar delays
using AutoIt;
using System.Runtime.InteropServices;
using System.Configuration;
using System.Diagnostics;
//using AutoItX3Lib;

//using Tesseract;

namespace AutoIt
{
    [Description("Click en unas coordenadas especificas mediante AutoIT.")]
    public class ClickByCoordinates : CodeActivity
    {
        public ClickByCoordinates()
        {
            DelayAfter = 300;
            DelayBefore = 200;
            OffsetX = 0;
            OffsetY = 0;
        }

        public enum Button
        {
            left,
            right,
            middle,
            main,
            menu,
            primary,
            secondary
        }

        [Category("Common")]
        [DisplayName("Delay After")]
        [Description("Cantidad de tiempo (en milisegundos) después de ejecutar la actividad.")]
        [RequiredArgument]
        public InArgument<Int32> DelayAfter { get; set; }

        [Category("Common")]
        [DisplayName("Delay Before")]
        [Description("Cantidad de tiempo (en milisegundos) antes de ejecutar la actividad.")]
        [RequiredArgument]
        public InArgument<int> DelayBefore { get; set; }

        [Category("Input")]
        [DefaultValue(true)]
        [RequiredArgument]
        [DisplayName("Tipo De Click")]
        [Description("The button to click, \"left\", \"right\", \"middle\", \"main\", \"menu\", \"primary\", \"secondary\". Default is the left button, in that case, left this field empty.")]
        public Button MouseButton { get; set; }

        [Category("ControlClick Coords")]
        [DefaultValue(true)]
        [RequiredArgument]
        [DisplayName("Posición X")]
        [Description("Asignar lo establecido en ControlClick Coords para X de la herramienta Au3Info.exe")] // Actualmente se envía cero como valor, no se ha logrado hacer click en elemento relativo
        public InArgument<Int32> OffsetX { get; set; }

        [Category("ControlClick Coords")]
        [DefaultValue(true)]
        [RequiredArgument]
        [DisplayName("Posición Y")]
        [Description("Asignar lo establecido en ControlClick Coords para Y de la herramienta Au3Info.exe")] // Actualmente se envía cero como valor, no se ha logrado hacer click en elemento relativo
        public InArgument<Int32> OffsetY { get; set; }

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
               // String vStrSelector = Selector.Get(context);
               // Dictionary<string, UiSelector> vDicSelector = SelectorManager.Funciones.TransformIntoUiSelector(Selector.Get(context));
                String vStrMouseButton = MouseButton.ToString();
                Int32 vIntOffsetX = OffsetX.Get(context);
                Int32 vIntOffsetY = OffsetY.Get(context);
                Int32 vIntDelayAfter = DelayAfter.Get(context);
                Int32 vIntDelayBefore = DelayBefore.Get(context);

                // Tiempo de espera antes de iniciar la actividad
                Thread.Sleep(vIntDelayBefore);

                AutoItX.MouseClick(vStrMouseButton, vIntOffsetX, vIntOffsetY, 1);

                // Tiempo de espera después de finalizada la actividad
                Thread.Sleep(vIntDelayAfter);
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

    [Description("-En Construcción-")]
	public class ClickText : CodeActivity
	{

		protected override void Execute(CodeActivityContext context)
		{
			throw new NotImplementedException();
		}
	}

    [Description("-En Construcción-")]
    public class ClickElement : CodeActivity
    {

        protected override void Execute(CodeActivityContext context)
        {
            throw new NotImplementedException();
        }
    }

    [Description("Capturar texto de un elemento mediante AutoIT.")]
    public class GetText : CodeActivity
    {
        public GetText()
        {
            TimeoutMS = 30000;
            DelayAfter = 300;
            DelayBefore = 200;
        }

        [Category("Common")]
        [DisplayName("Delay After")]
        [Description("Cantidad de tiempo (en milisegundos) después de ejecutar la actividad.")]
        [RequiredArgument]
        public InArgument<Int32> DelayAfter { get; set; }


        [Category("Common")]
        [DisplayName("Delay Before")]
        [Description("Cantidad de tiempo (en milisegundos) antes de ejecutar la actividad.")]
        [RequiredArgument]
        public InArgument<int> DelayBefore { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [DisplayName("Selector")]
        [Description("Selector Tipo Uipath")]
        public InArgument<String> Selector { get; set; }

        [Category("Input")]
        [DefaultValue(true)]
        [DisplayName("Time Out")]
        [Description("Tiempo de busqueda maxima del elemento.")]
        public InArgument<Int32> TimeoutMS { get; set; }

        [Category("Output")]
        [DisplayName("Value")]
        [Description("Texto capturado del elemento")]
        public OutArgument<String> Value { get; set; }

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
                Dictionary<string, UiSelector> vDicSelector = SelectorManager.Funciones.TransformIntoUiSelector(Selector.Get(context));
                Int32 vIntTimeoutMS = TimeoutMS.Get(context);
                Int32 vIntDelayAfter = DelayAfter.Get(context);
                Int32 vIntDelayBefore = DelayBefore.Get(context);

                // Tiempo de espera antes de iniciar la actividad
                Thread.Sleep(vIntDelayBefore);

                // Crear los CssSelector
                String vStrAutoItSelector = Misc.CreateAutoItSelector(vDicSelector);
                Console.WriteLine("Se creó el vStrAutoItSelector: " + vStrAutoItSelector);

                // Activar la ventana antes de realizar la acción
                AutoItX.WinActivate(vDicSelector["0"].Attribute.title);

                // Ejecutar la acción
                String vStrReturValue = AutoItX.ControlGetText(vDicSelector["0"].Attribute.title, "", vStrAutoItSelector);

                // Establecer argumento de salida
                Value.Set(context, vStrReturValue);

                // Tiempo de espera después de finalizada la actividad
                Thread.Sleep(vIntDelayAfter);
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

    [Description("Presionar una key (boton especial) en el equipo. EJ: Enter.")]
    public class PressKeyC : CodeActivity
    {
        public PressKeyC()
        {
            DelayAfter = 300;
            DelayBefore = 200;
      
        }

        [Category("Common")]
        [DisplayName("Delay After")]
        [Description("Cantidad de tiempo (en milisegundos) después de ejecutar la actividad.")]
        [RequiredArgument]
        public InArgument<Int32> DelayAfter { get; set; }


        [Category("Common")]
        [DisplayName("Delay Before")]
        [Description("Cantidad de tiempo (en milisegundos) antes de ejecutar la actividad.")]
        [RequiredArgument]
        public InArgument<int> DelayBefore { get; set; }

        [Category("Input")]
        [DisplayName("Key to press")]
        [Description("Key to Press: {ENTER},ETC")]
        [RequiredArgument]
        public InArgument<String>  Key { get; set; }

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


        [DllImport("USER32.DLL", CharSet = CharSet.Unicode)]
        public static extern IntPtr FindWindow(string lpClassName,
            string lpWindowName);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr FindWindowEx(IntPtr parentHandle, IntPtr childAfter, string className, string windowTitle);


        // Activate an application window.
        [DllImport("USER32.DLL")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern int SendMessage(int hWnd, int msg, int wParam, IntPtr lParam);

        private const int BN_CLICKED = 245;

        protected override void Execute(CodeActivityContext context)
        {
            try
            {

                // Obtener argumentos
                Int32 vIntDelayAfter = DelayAfter.Get(context);
                Int32 vIntDelayBefore = DelayBefore.Get(context);
                Thread.Sleep(vIntDelayBefore);

                IntPtr hwndChild = IntPtr.Zero;
                IntPtr calculatorHandle = FindWindow("Button", "Abrir");
                hwndChild = FindWindowEx((IntPtr)calculatorHandle, IntPtr.Zero, "Button", "1");

                //send BN_CLICKED message
                SendMessage((int)hwndChild, BN_CLICKED, 0, IntPtr.Zero);

                //  SetForegroundWindow(calculatorHandle);

                //System.Windows.Forms.SendKeys.SendWait("{ENTER}");
                // Tiempo de espera después de finalizada la actividad
                Thread.Sleep(vIntDelayAfter);
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

    [Description("-En Construcción-")]
    public class SelectWinMenuItem : CodeActivity
    {

        protected override void Execute(CodeActivityContext context)
        {

            Console.WriteLine("Actividad sin configurar");
            // Obtener argumentos
            //AutoItX3 au3 = new AutoItX3();
            //au3.WinMenuSelectItem("[CLASS:Notepad]", "", "&Archivo", "Guardar como...");
            //AutoItX.WinActivate("[CLASS:Notepad]");
            //var resp = Misc.AU3_WinMenuSelectItem("Sin título: Bloc de notas", "", "&Archivo", "", "", "", "", "", "","");
            //Console.WriteLine("End click: " + resp.ToString());

            //Int32 ReturnValue = AutoItX.ControlClick("Sin título: Bloc de notas", "", "[NAME:Aplicación]", button: "", numClicks: 1, x: 10, y: 15);

            //Console.WriteLine("Return value " + ReturnValue.ToString()); // 1 for success 0 for error

            //AutoIt.SelectWinMenuItem();
            //AutoItX3 au3 = new AutoItX3();

        }
    }

    [Description("Enviar HotKeys (atajos con botones especiales) a un elemento mediante AutoIT. EJ: Alt + F4")]
    public class SendHotKeys : CodeActivity
    {
        public SendHotKeys()
        {
            DelayAfter = 300;
            DelayBefore = 1000;
            TimeoutMS = 30000;
            Activate = true;
        }
        public enum KeysStr
        {
            ALT_F4,

        }

        [Category("Common")]
        [DisplayName("Delay After")]
        [Description("Cantidad de tiempo (en milisegundos) después de ejecutar la actividad.")]
        [RequiredArgument]
        public InArgument<Int32> DelayAfter { get; set; }

        [Category("Common")]
        [DisplayName("Delay Before")]
        [Description("Cantidad de tiempo (en milisegundos) antes de ejecutar la actividad.")]
        [RequiredArgument]
        public InArgument<int> DelayBefore { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [DisplayName("Hot Key")]
        [Description("Hot Key to Press, EJ: ALT_F4")]
        public KeysStr Key { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [DisplayName("Selector")]
        [Description("Selector Tipo Uipath")]
        public InArgument<String> Selector { get; set; }

        [Category("Input")]
        [DisplayName("Time Out")]
        [Description("Tiempo de busqueda maxima del elemento.")]
        [DefaultValue(true)]
        public InArgument<Int32> TimeoutMS { get; set; }

        [Category("Options")]
        [DefaultValue(true)]
        [DisplayName("Activate")]
        [Description("Se usa el ControlID del selector, en caso de no especificarse no activará el elemento antes de ejecutar la actividad.")]
        public Boolean Activate { get; set; }

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
                Int32 vIntDelayAfter = DelayAfter.Get(context);
                Int32 vIntDelayBefore = DelayBefore.Get(context);
                Dictionary<string, UiSelector> vDicSelector;

                Dictionary<String, String> vDicKey = Misc.KeyCombinedDic();
                String vStrText = vDicKey[Key.ToString()].ToString();
                Int32 vIntTimeoutMS = TimeoutMS.Get(context);

                // Tiempo de espera antes de iniciar la actividad
                Thread.Sleep(vIntDelayBefore);

                // Crear los CssSelector
                if (Activate)
                {
                    vDicSelector = SelectorManager.Funciones.TransformIntoUiSelector(Selector.Get(context));
                    String vStrAutoItSelector = Misc.CreateAutoItSelector(vDicSelector);

                    // Activar la ventana antes de realizar la acción
                    AutoItX.WinActivate(vDicSelector["0"].Attribute.title);
                    if (!String.IsNullOrWhiteSpace(vDicSelector["1"].Attribute.id))
                    {
                        AutoItX.ControlFocus(vDicSelector["0"].Attribute.title, "", vDicSelector["1"].Attribute.id);
                    }
                }

                // Ejecutar la acción
                AutoItX.Send(vStrText);

                // Tiempo de espera después de finalizada la actividad
                Thread.Sleep(vIntDelayAfter);
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

    [Description("Enviar Keys (botones especiales) a un elemento mediante AutoIT.")]
    public class SendKeys : CodeActivity
    {
        public SendKeys()
        {
            DelayAfter = 300;
            DelayBefore = 1000;
            TimeoutMS = 30000;
            Activate = true;
        }
        
        public enum KeysStr
        {
            SPACE,
            ENTER,
            ALT,
            BACKSPACE,
            DELETE,
            Up,
            Down,
            Left,
            Right,
            HOME,
            END,
            ESCAPE,
            INSERT,
            PageUp,
            PageDown,
            Function,
            F1,
            F2,
            F3,
            F4,
            F5,
            F6,
            F7,
            F8,
            F9,
            F10,
            F11,
            F12,
            TAB,
            PrintScreen,
            LeftWindows,
            RightWindows,
            BreakProcessing,
            PAUSE,
            Numpad0,
            Numpad1,
            Numpad2,
            Numpad3,
            Numpad4,
            Numpad5,
            Numpad6,
            Numpad7,
            Numpad8,
            Numpad9,
            NumpadMultiply,
            NumpadAdd,
            NumpadSubtract,
            NumpadDivide,
            NumpadPeriod,
            NumpadEnter,
            WindowsAppkey,
            LeftALT,
            RightALT,
            LeftCTRL,
            RightCTRL,
            LeftShift,
            RightShift,
            SLEEP,
            BrowserBack,
            BrowserForward,
            BrowserRefresh,
            BrowserStop,
            BrowserSearch,
            BrowserFavorites,
            BrowserHome,
            MuteVolume,
            ReduceVolume,
            IncreaseVolume,
            NextTrackMediaPlayer,
            PreviousTrackMediaPlayer,
            StopMediaPlayer,
            PlayPauseMediaPlayer,
            LaunchEmailApplication,
            LaunchMediaPlayer,
            ControlC,
            ControlV,
        }

        [Category("Common")]
        [DisplayName("Delay After")]
        [Description("Cantidad de tiempo (en milisegundos) después de ejecutar la actividad.")]
        [RequiredArgument]
        public InArgument<Int32> DelayAfter { get; set; }

        [Category("Common")]
        [DisplayName("Delay Before")]
        [Description("Cantidad de tiempo (en milisegundos) antes de ejecutar la actividad.")]
        [RequiredArgument]
        public InArgument<int> DelayBefore { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [DisplayName("Key to send")]
        [Description("Key to send: SPACE, ENTER, ALT, ETC")]
        public KeysStr Key { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [DisplayName("Selector")]
        [Description("Selector Tipo Uipath")]
        public InArgument<String> Selector { get; set; }

        [Category("Input")]
        [DefaultValue(true)]
        [DisplayName("Time Out")]
        [Description("Tiempo de busqueda maxima del elemento.")]
        public InArgument<Int32> TimeoutMS { get; set; }

        [Category("Options")]
        [DefaultValue(true)]
        [DisplayName("Activate")]
        [Description("Se usa el ControlID del selector, en caso de no especificarse no activará el elemento antes de ejecutar la actividad.")]
        public Boolean Activate { get; set; }

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
                Int32 vIntDelayAfter = DelayAfter.Get(context);
                Int32 vIntDelayBefore = DelayBefore.Get(context);
                Dictionary<string, UiSelector> vDicSelector = SelectorManager.Funciones.TransformIntoUiSelector(Selector.Get(context));
                Dictionary<String, String> vDicKey = Misc.KeyDic();
                String vStrText = vDicKey[Key.ToString()].ToString();
                Int32 vIntTimeoutMS = TimeoutMS.Get(context);

                // Tiempo de espera antes de iniciar la actividad
                Thread.Sleep(vIntDelayBefore);

                // Crear los CssSelector
                String vStrAutoItSelector = Misc.CreateAutoItSelector(vDicSelector);
                //Console.WriteLine("Se creó el vStrAutoItSelector: " + vStrAutoItSelector);
                if (Activate)
                {
                    // Activar la ventana antes de realizar la acción
                    AutoItX.WinActivate(vDicSelector["0"].Attribute.title);
                    if (!String.IsNullOrWhiteSpace(vDicSelector["1"].Attribute.id))
                    {
                        AutoItX.ControlFocus(vDicSelector["0"].Attribute.title, "", vDicSelector["1"].Attribute.id);
                    }

                }

                // Ejecutar la acción
                AutoItX.Send(vStrText);

                // Tiempo de espera después de finalizada la actividad
                Thread.Sleep(vIntDelayAfter);
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

    [Description("Escribir texto de un elemento mediante AutoIT.")]
    public class SetText : CodeActivity
    {
        public SetText()
        {
            DelayAfter = 300;
            DelayBefore = 1000;
            TimeoutMS = 30000;
        }

        [Category("Common")]
        [DisplayName("Delay After")]
        [Description("Cantidad de tiempo (en milisegundos) después de ejecutar la actividad.")]
        [RequiredArgument]
        public InArgument<Int32> DelayAfter { get; set; }

        [Category("Common")]
        [DisplayName("Delay Before")]
        [Description("Cantidad de tiempo (en milisegundos) antes de ejecutar la actividad.")]
        [RequiredArgument]
        public InArgument<int> DelayBefore { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [DisplayName("Texto")]
        [Description("Texto que se quiere ingresar en el elemento.")]
        public InArgument<String> Text { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [DisplayName("Selector")]
        [Description("Selector Tipo Uipath")]
        public InArgument<String> Selector { get; set; }

        [Category("Input")]
        [DefaultValue(true)]
        [DisplayName("Time Out")]
        [Description("Tiempo de busqueda maxima del elemento.")]
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
                Int32 vIntDelayAfter = DelayAfter.Get(context);
                Int32 vIntDelayBefore = DelayBefore.Get(context);
                Dictionary<string, UiSelector> vDicSelector = SelectorManager.Funciones.TransformIntoUiSelector(Selector.Get(context));
                String vStrText = Text.Get(context);
                Int32 vIntTimeoutMS = TimeoutMS.Get(context);

                // Tiempo de espera antes de iniciar la actividad
                Thread.Sleep(vIntDelayBefore);

                // Crear los CssSelector
                String vStrAutoItSelector = Misc.CreateAutoItSelector(vDicSelector);
                Console.WriteLine("Se creó el vStrAutoItSelector: " + vStrAutoItSelector);

                // Activar la ventana antes de realizar la acción
                AutoItX.WinActivate(vDicSelector["0"].Attribute.title);

                // Ejecutar la acción
                Int32 ReturnValue = AutoItX.ControlSetText(vDicSelector["0"].Attribute.title, "", vStrAutoItSelector, vStrText);

                // Imprimir el resultado de la acción
                Console.WriteLine("Return value " + ReturnValue.ToString()); // 1 for success 0 for error

                // Tiempo de espera después de finalizada la actividad
                Thread.Sleep(vIntDelayAfter);
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

    [Description("Escribir en un elemento mediante AutoIT.")]
    public class TypeInto : CodeActivity
    {
        public TypeInto()
        {
            DelayAfter = 300;
            DelayBefore = 1000;
            TimeoutMS = 30000;
            EmptyField = false;
        }

        [Category("Common")]
        [DisplayName("Delay After")]
        [Description("Cantidad de tiempo (en milisegundos) después de ejecutar la actividad.")]
        [RequiredArgument]
        public InArgument<Int32> DelayAfter { get; set; }


        [Category("Common")]
        [DisplayName("Delay Before")]
        [Description("Cantidad de tiempo (en milisegundos) antes de ejecutar la actividad.")]
        [RequiredArgument]
        public InArgument<int> DelayBefore { get; set; }


        [Category("Input")]
        [RequiredArgument]
        [DisplayName("Selector")]
        [Description("Selector Tipo Uipath")]
        public InArgument<String> Selector { get; set; }

        [Category("Input")]
        [DisplayName("Texto")]
        [Description("Texto que se quiere ingresar en el elemento.")]
        public InArgument<String> Text { get; set; }

        [Category("Input")]
        [DefaultValue(true)]
        [DisplayName("Time Out")]
        [Description("Tiempo de busqueda maxima del elemento.")]
        public InArgument<Int32> TimeoutMS { get; set; }

        [Category("Options")]
        [DefaultValue(true)]
        [DisplayName("Limpiar Campo - En construcción")]
        [Description("Seleccionar True para limpiar el campo antes de escribir.")]
        public InArgument<Boolean> EmptyField { get; set; }

        protected override void Execute(CodeActivityContext context)
        {

            // Obtener argumentos
            String vStrSelector = Selector.Get(context);
            String vStrText = Text.Get(context);
            Dictionary<string, UiSelector> vDicSelector = SelectorManager.Funciones.TransformIntoUiSelector(Selector.Get(context));
            Int32 vIntDelayAfter = DelayAfter.Get(context);
            Int32 vIntDelayBefore = DelayBefore.Get(context);

            // Tiempo de espera antes de iniciar la actividad
            Thread.Sleep(vIntDelayBefore);

            AutoItX.WinWait(vDicSelector["0"].Attribute.title, "", TimeoutMS.Get(context) / 1000);
            AutoItX.WinActivate(vDicSelector["0"].Attribute.title);

            // Crear los CssSelector
            String vStrAutoItSelector = Misc.CreateAutoItSelector(vDicSelector);
            Console.WriteLine("Se creó el vStrAutoItSelector: " + vStrAutoItSelector);
            Console.WriteLine("title: " + vDicSelector["0"].Attribute.title);
            //AutoIt.AutoItX.WinWaitActive(vDicSelector["0"].Attribute.title);

            DateTime vDatInicio = DateTime.Now;
            for (int i = 0; i < TimeoutMS.Get(context); i = (int)(DateTime.Now - vDatInicio).TotalMilliseconds)
            {
                try
                {
                    int vIntResultado;
                    // Enviar Ctrl+A al selector para seleccionar todo el texto dentro de este
                    vIntResultado = AutoItX.ControlSend(vDicSelector["0"].Attribute.title, "", vStrAutoItSelector, "^a", 0);
                    if (vIntResultado == 0)
                    {
                        throw new System.Exception("Window/control is not found");
                    }
                    // Enviar el texto a selector especificado
                    vIntResultado = AutoItX.ControlSend(vDicSelector["0"].Attribute.title, "", vStrAutoItSelector, vStrText, 0);
                    if (vIntResultado == 0)
                    {
                        throw new System.Exception("Window/control is not found");
                    }
                    break;
                }
                catch (Exception)
                {
                    Thread.Sleep(3000);
                }
            }

            // Tiempo de espera después de finalizada la actividad
            Thread.Sleep(vIntDelayAfter);
        }
    }

//Funciones
    public class Misc
    {

        //AU3_API long WINAPI AU3_WinMenuSelectItem(const char *szTitle, /*[in,defaultvalue("")]*/const char *szText
        //, const char *szItem1, const char *szItem2, const char *szItem3, const char *szItem4, const char *szItem5
        //, const char *szItem6, const char *szItem7, const char *szItem8);
        [DllImport("AutoItX3.dll", SetLastError = true, CharSet = CharSet.Auto)]
        static public extern int AU3_WinMenuSelectItem([MarshalAs(UnmanagedType.LPStr)]string Title
            , [MarshalAs(UnmanagedType.LPStr)] string Text, [MarshalAs(UnmanagedType.LPStr)] string Item1
            , [MarshalAs(UnmanagedType.LPStr)] string Item2, [MarshalAs(UnmanagedType.LPStr)] string Item3
            , [MarshalAs(UnmanagedType.LPStr)] string Item4, [MarshalAs(UnmanagedType.LPStr)] string Item5
            , [MarshalAs(UnmanagedType.LPStr)] string Item6, [MarshalAs(UnmanagedType.LPStr)] string Item7
            , [MarshalAs(UnmanagedType.LPStr)] string Item8);

        public static string CreateAutoItSelector(Dictionary<string, UiSelector> vDicSelector)
        {
            String vStrAutoItSelector = "";
            UiSelector vObjChildSelector = vDicSelector["1"];

            // Si existe id
            // ID - The internal control ID. The Control ID is the internal numeric identifier that windows gives to each control.
            // It is generally the best method of identifying controls. In addition to the AutoIt Window Info Tool, 
            // other applications such as screen readers for the blind and Microsoft tools/APIs may allow you
            // to get this Control ID
            if (!String.IsNullOrWhiteSpace(vObjChildSelector.Attribute.id)) { vStrAutoItSelector = vStrAutoItSelector + " ID:" + vObjChildSelector.Attribute.id + ";"; }

            // Si existe text
            // TEXT - The text on a control, for example "&Next" on a button
            if (!String.IsNullOrWhiteSpace(vObjChildSelector.Attribute.text)) { vStrAutoItSelector = vStrAutoItSelector + " TEXT:" + vObjChildSelector.Attribute.text + ";"; }

            // Si existe Class
            // CLASS - The internal control classname such as "Edit" or "Button"
            if (!String.IsNullOrWhiteSpace(vObjChildSelector.Attribute.className)) { vStrAutoItSelector = vStrAutoItSelector + " CLASS:" + vObjChildSelector.Attribute.className + ";"; }

            // Si existe name
            // NAME - The internal .NET Framework WinForms name (if available)
            if (!String.IsNullOrWhiteSpace(vObjChildSelector.Attribute.name)) { vStrAutoItSelector = vStrAutoItSelector + " NAME:" + vObjChildSelector.Attribute.name + ";"; }

            // Si existe instance
            // INSTANCE - The 1-based instance when all given properties match.
            if (!String.IsNullOrWhiteSpace(vObjChildSelector.Attribute.instance)) { vStrAutoItSelector = vStrAutoItSelector + " INSTANCE:" + vObjChildSelector.Attribute.instance + ";"; }

            // Quitar el último ; y el espacio al inicio de la cadena de texto
            if (!String.IsNullOrWhiteSpace(vStrAutoItSelector))
            {
                vStrAutoItSelector = vStrAutoItSelector.Substring(1, vStrAutoItSelector.Length - 2);
            }

            // Agregar llaves
            vStrAutoItSelector = "[" + vStrAutoItSelector + "]";

            return vStrAutoItSelector;
        }

        public static Dictionary<String, String> KeyDic()
        {
            Dictionary<String, String> vDicKey = new Dictionary<String, String>
            {
                {"SPACE","{SPACE}"},
                {"ENTER","{ENTER}"},
                {"ALT","{ALT}"},
                {"BACKSPACE","{BACKSPACE}"},
                {"DELETE","{DEL}"},
                {"Up","{UP}"},
                {"Down","{DOWN}"},
                {"Left","{LEFT}"},
                {"Right","{RIGHT}"},
                {"HOME","{HOME}"},
                {"END","{END}"},
                {"ESCAPE","{ESC}"},
                {"INSERT","{INS}"},
                {"PageUp","{PGUP}"},
                {"PageDown","{PGDN}"},
                {"Function","{F1} - {F12}"},
                {"F1","{F1}"},
                {"F2","{F2}"},
                {"F3","{F3}"},
                {"F4","{F4}"},
                {"F5","{F5}"},
                {"F6","{F6}"},
                {"F7","{F7}"},
                {"F8","{F8}"},
                {"F9","{F9}"},
                {"F10","{F10}"},
                {"F11","{F11}"},
                {"F12","{F12}"},
                {"TAB","{TAB}"},
                {"PrintScreen","{PRINTSCREEN}"},
                {"LeftWindows","{LWIN}"},
                {"RightWindows","{RWIN}"},
                {"BreakProcessing","{BREAK}"},
                {"PAUSE","{PAUSE}"},
                {"Numpad0","{NUMPAD0}"},
                {"Numpad1","{NUMPAD1}"},
                {"Numpad2","{NUMPAD2}"},
                {"Numpad3","{NUMPAD3}"},
                {"Numpad4","{NUMPAD4}"},
                {"Numpad5","{NUMPAD5}"},
                {"Numpad6","{NUMPAD6}"},
                {"Numpad7","{NUMPAD7}"},
                {"Numpad8","{NUMPAD8}"},
                {"Numpad9","{NUMPAD9}"},
                {"NumpadMultiply","{NUMPADMULT}"},
                {"NumpadAdd","{NUMPADADD}"},
                {"NumpadSubtract","{NUMPADSUB}"},
                {"NumpadDivide","{NUMPADDIV}"},
                {"NumpadPeriod","{NUMPADDOT}"},
                {"NumpadEnter","{NUMPADENTER}"},
                {"WindowsAppkey","{APPSKEY}"},
                {"LeftALT","{LALT}"},
                {"RightALT","{RALT}"},
                {"LeftCTRL","{LCTRL}"},
                {"RightCTRL","{RCTRL}"},
                {"LeftShift","{LSHIFT}"},
                {"RightShift","{RSHIFT}"},
                {"SLEEP","{SLEEP}"},
                {"BrowserBack","{BROWSER_BACK}"},
                {"BrowserForward","{BROWSER_FORWARD}"},
                {"BrowserRefresh","{BROWSER_REFRESH}"},
                {"BrowserStop","{BROWSER_STOP}"},
                {"BrowserSearch","{BROWSER_SEARCH}"},
                {"BrowserFavorites","{BROWSER_FAVORITES}"},
                {"BrowserHome","{BROWSER_HOME}"},
                {"MuteVolume","{VOLUME_MUTE}"},
                {"ReduceVolume","{VOLUME_DOWN}"},
                {"IncreaseVolume","{VOLUME_UP}"},
                {"NextTrackMediaPlayer","{MEDIA_NEXT}"},
                {"PreviousTrackMediaPlayer","{MEDIA_PREV}"},
                {"StopMediaPlayer","{MEDIA_STOP}"},
                {"PlayPauseMediaPlayer","{MEDIA_PLAY_PAUSE}"},
                {"LaunchEmailApplication","{LAUNCH_MAIL}"},
                {"LaunchMediaPlayer","{LAUNCH_MEDIA}"},
                {"ControlC","^{C}"},
                {"ControlV","^{V}"},
            };

            return vDicKey;
        }




        public static Dictionary<String, String> KeyCombinedDic()
        {
            Dictionary<String, String> vDicKey = new Dictionary<String, String>
            {
                {"ALT_F4","!{F4}"},
                {"Not implemented","{ENTER}"},
            
            };

            return vDicKey;
        }


    }
}
