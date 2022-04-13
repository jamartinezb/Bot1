using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Activities;
using System.ComponentModel;
using System.Windows.Forms;
using System.Threading;
using System.Windows.Automation;
using System.Runtime.InteropServices;



namespace CapturarElementos
{
    [Description("Capturar detalles de elementos dentro de aplicaciones")]
    public class GetElements : CodeActivity
    {
        [DllImport("user32.dll")]
        static extern void mouse_event(int dwFlags, int dx, int dy, int dwData, int dwExtraInfo);

        private const int MOUSEEVENTF_LEFTDOWN = 0x02;
        private const int MOUSEEVENTF_LEFTUP = 0x04;
        private const int MOUSEEVENTF_RIGHTDOWN = 0x08;
        private const int MOUSEEVENTF_RIGHTUP = 0x10;

        public void ClickMouseLeftButton(System.Drawing.Point globalLocation)
        {
            System.Drawing.Point currLocation = Cursor.Position;

            Cursor.Position = globalLocation;

            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP,
                globalLocation.X, globalLocation.Y, 0, 0);

            //Cursor.Position = currLocation;
        }

        protected override void Execute(CodeActivityContext context)
        {

            while (true)
            {
                System.Windows.Point Posicion;
                System.Drawing.Point pt = Cursor.Position;
                AutomationElement el = AutomationElement.FromPoint(new System.Windows.Point(pt.X, pt.Y));
                //AutomationElement ol = AutomationElement.
                Console.WriteLine("--------------------------------------------------------------");
                Console.WriteLine("Name: " + el.Current.Name);
                Console.WriteLine("ClassName: " + el.Current.ClassName);
                //Console.WriteLine("ControlType" + el.Current.ControlType);
                Console.WriteLine("ProcessId: " + el.Current.ProcessId);
                Console.WriteLine("IsControlElement: " + el.Current.ProcessId);
                Console.WriteLine("AutomationID: " + el.Current.AutomationId);
                Console.WriteLine("FrameWorkID: " + el.Current.FrameworkId);
                Console.WriteLine("NativeWindowHandle: " + el.Current.NativeWindowHandle);
                Console.WriteLine("AcceleratorKey: " + el.Current.AcceleratorKey);
                Console.WriteLine("AccesKey: " + el.Current.AccessKey);
                Console.WriteLine("BoundingRectangle: " + el.Current.BoundingRectangle);
                //Console.WriteLine("ControlType: " + el.Current.ControlType);
                Console.WriteLine("HelpText: " + el.Current.HelpText);
                Console.WriteLine("IsControlElement: " + el.Current.IsControlElement);
                Console.WriteLine("Enabled: " + el.Current.IsEnabled);
                Console.WriteLine("KeyboardFocusable: " + el.Current.IsKeyboardFocusable);
                Console.WriteLine("Visible: " + el.Current.IsOffscreen);
                Console.WriteLine("ItemType: " + el.Current.ItemType);
                Console.WriteLine("LabeledBy: " + el.Current.LabeledBy);
                Console.WriteLine("LocalizedControlType: " + el.Current.LocalizedControlType);
                //Console.WriteLine("SupportedPatterns: " + el.GetSupportedPatterns());
                //Console.WriteLine("ParentName: " + el.GetCurrentPattern());

                Posicion = el.GetClickablePoint();
                Console.WriteLine("Posicion: " + Posicion.ToString());
                Console.WriteLine("RunTimeId:" + el.GetRuntimeId());

                //Posicion para el boton inicial
                System.Drawing.Point PosicionEnviar = new System.Drawing.Point(Convert.ToInt32(Posicion.X), Convert.ToInt32(Posicion.Y));

                //Posicion capturando el boton
                //IntPtr VentanaHandle = new IntPtr(el.Current.NativeWindowHandle);
                //AutomationElement el2 = AutomationElement.FromHandle(VentanaHandle);

                //Evento para saber cuando cambia el foco
                //                AutomationElement.AutomationFocusChangedEvent




                //Thread.Sleep(2000);
                //ClickMouseLeftButton(PosicionEnviar);
                //Thread.Sleep(3000);
            }
        }
    }
}
