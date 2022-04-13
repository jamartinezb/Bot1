
// *
// @author Carlos Cardona - Andres Tarazona Mora - Comdata
// * @version 1.00, 20/05/2019
// *

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.IE;

using System.Activities;
using System.ComponentModel;

using SelectorManager;

using System.Threading; // Utilizado para agregar delays

using System.Runtime.InteropServices;  // Utilizado para agregar parámetros opcionales en las funciones


namespace Selenium
{
    [Description("Validar si existe una alerta en un navegador mediante Selenium")]
    public class AlertExist : CodeActivity
    {
        public AlertExist()
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
        [DisplayName("Browser")]
        [Description("Objeto del navegador donde se realizará la actividad")]
        public InArgument<IWebDriver> SeleniumBrowser { get; set; }

        [Category("Input")]
        [DefaultValue(true)]
        [DisplayName("Time Out")]
        [Description("Tiempo de busqueda maxima del elemento.")]
        public InArgument<Int32> TimeoutMS { get; set; }

        [Category("Output")]
        [RequiredArgument]
        [DisplayName("Existe Alerta")]
        [Description("Resultado si existe o no.")]
        public OutArgument<Boolean> Exist { get; set; }

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
                // Tiempo de espera antes de iniciar la actividad
                Thread.Sleep(DelayBefore.Get(context));

                DateTime vDatInicio = DateTime.Now;
                for (int i = 0; i < TimeoutMS.Get(context); i = (int)(DateTime.Now - vDatInicio).TotalMilliseconds)
                {
                    try
                    {
                        IAlert al = SeleniumBrowser.Get(context).SwitchTo().Alert();
                        Exist.Set(context, true);
                        break;
                    }
                    catch (Exception)
                    {
                        Exist.Set(context, false);
                        Thread.Sleep(3000);
                    }
                }
                // Tiempo de espera despues de iniciar la actividad
                Thread.Sleep(DelayAfter.Get(context));

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

    [Description("Capturar texto sobre una alerta en un navegador mediante Selenium")]
    public class AlertGetText : CodeActivity
    {
        public AlertGetText()
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
        [DisplayName("Browser")]
        [Description("Objeto del navegador donde se realizará la actividad")]
        public InArgument<IWebDriver> SeleniumBrowser { get; set; }

        [Category("Input")]
        [DefaultValue(true)]
        [DisplayName("Time Out")]
        [Description("Tiempo de busqueda maxima del elemento.")]
        public InArgument<Int32> TimeoutMS { get; set; }

        [Category("Output")]
        [RequiredArgument]
        [DisplayName("Value")]
        [Description("Texto capturado del elemento")]
        public OutArgument<String> Text { get; set; }

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

                // Tiempo de espera antes de iniciar la actividad
                Thread.Sleep(DelayBefore.Get(context));

                DateTime vDatInicio = DateTime.Now;
                for (int i = 0; i < TimeoutMS.Get(context); i = (int)(DateTime.Now - vDatInicio).TotalMilliseconds)
                {
                    try
                    {
                        IAlert al = SeleniumBrowser.Get(context).SwitchTo().Alert();
                        Text.Set(context, al.Text);
                        break;
                    }
                    catch (Exception)
                    {
                        // Si no existe un continue on error y se excedió el tiempo de espera
                        if ((int)(DateTime.Now - vDatInicio).TotalMilliseconds > TimeoutMS.Get(context))
                        {
                            throw;
                        }
                        else
                        {
                            Text.Set(context, "");
                        }
                        Thread.Sleep(3000);
                    }
                }

                // Tiempo de espera despues de iniciar la actividad
                Thread.Sleep(DelayAfter.Get(context));
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

    [Description("Ejecutar actión sobre una alerta en un navegador mediante Selenium")]
    public class AlertSendAction : CodeActivity
    {
        public AlertSendAction()
        {
            DelayAfter = 300;
            DelayBefore = 1000;
            TimeoutMS = 30000;
        }

        public enum Actions
        {
            Accept,
            Dismiss,
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
        [DisplayName("Browser")]
        [Description("Objeto del navegador donde se realizará la actividad")]
        public InArgument<IWebDriver> SeleniumBrowser { get; set; }

        [Category("Input")]
        [DefaultValue(true)]
        [DisplayName("Time Out")]
        [Description("Tiempo de busqueda maxima del elemento.")]
        public InArgument<Int32> TimeoutMS { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [DisplayName("Select action")]
        [Description("Acción a realizar.")]
        public Actions SelectAction { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            // Obtener argumentos

            // Tiempo de espera antes de iniciar la actividad
            Thread.Sleep(DelayBefore.Get(context));

            DateTime vDatInicio = DateTime.Now;
            for (int i = 0; i < TimeoutMS.Get(context); i = (int)(DateTime.Now - vDatInicio).TotalMilliseconds)
            {
                try
                {
                    IAlert al = SeleniumBrowser.Get(context).SwitchTo().Alert();
                    switch (SelectAction)
                    {
                        case Actions.Accept:
                            al.Accept();
                            break;
                        case Actions.Dismiss:
                            al.Dismiss();
                            break;
                        default:
                            break;
                    }
                    break;
                }
                catch (Exception)
                {
                    // Si no existe un continue on error y se excedió el tiempo de espera
                    if ((int)(DateTime.Now - vDatInicio).TotalMilliseconds > TimeoutMS.Get(context))
                    {
                        throw;
                    }
                    Thread.Sleep(3000);
                }
            }
            // Tiempo de espera despues de iniciar la actividad
            Thread.Sleep(DelayAfter.Get(context));
        }
    }

    [Description("Enviar texto sobre una alerta en un navegador mediante Selenium")]
    public class AlertSendText : CodeActivity
    {
        public AlertSendText()
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
        [DisplayName("Browser")]
        [Description("Objeto del navegador donde se realizará la actividad")]
        public InArgument<IWebDriver> SeleniumBrowser { get; set; }

        [Category("Input")]
        [DefaultValue(true)]
        [DisplayName("Time Out")]
        [Description("Tiempo de busqueda maxima del elemento.")]
        public InArgument<Int32> TimeoutMS { get; set; }

        [Category("Output")]
        [RequiredArgument]
        [DisplayName("Texto")]
        [Description("Texto que se quiere ingresar en el elemento.")]
        public InArgument<String> Text { get; set; }

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

                // Tiempo de espera antes de iniciar la actividad
                Thread.Sleep(DelayBefore.Get(context));

                DateTime vDatInicio = DateTime.Now;
                for (int i = 0; i < TimeoutMS.Get(context); i = (int)(DateTime.Now - vDatInicio).TotalMilliseconds)
                {
                    try
                    {
                        IAlert al = SeleniumBrowser.Get(context).SwitchTo().Alert();
                        al.SendKeys(Text.Get(context));
                        break;
                    }
                    catch (Exception)
                    {
                        // Si no existe un continue on error y se excedió el tiempo de espera
                        if ((int)(DateTime.Now - vDatInicio).TotalMilliseconds > TimeoutMS.Get(context))
                        {
                            throw;
                        }
                        Thread.Sleep(3000);
                    }
                }

                // Tiempo de espera despues de iniciar la actividad
                Thread.Sleep(DelayAfter.Get(context));
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

    [Description("Click by CssSelector en un elemento web mediante Selenium")]
    public class Click : CodeActivity
    {

        public Click()
        {
            DelayAfter = 300;
            DelayBefore = 1000;
            TimeoutMS = 30000;
            //InjectJS = false;
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
        [DisplayName("Browser")]
        [Description("Objeto del navegador donde se realizará la actividad")]
        public InArgument<IWebDriver> SeleniumBrowser { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [DisplayName("Selector")]
        [Description("Selector tipo Css")]
        public InArgument<String> Selector { get; set; }

        [Category("Input")]
        [DefaultValue(true)]
        [DisplayName("Time Out")]
        [Description("Tiempo de busqueda maxima del elemento.")]
        public InArgument<Int32> TimeoutMS { get; set; }

        [Category("Options")]
        [DisplayName("Inyectar Javascript?")]
        [Description("True para Realizar la actividad mediante inyección de Javascript")]
        public Boolean InjectJS { get; set; }

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
                IWebDriver vObjBrowser = (IWebDriver)SeleniumBrowser.Get(context);
                Int32 vIntDelayAfter = DelayAfter.Get(context);
                Int32 vIntDelayBefore = DelayBefore.Get(context);
                Dictionary<string, UiSelector> vDicSelector = SelectorManager.Funciones.TransformIntoUiSelector(Selector.Get(context));
                Int32 vIntTimeoutMS = TimeoutMS.Get(context);

                // Tiempo de espera antes de iniciar la actividad
                Thread.Sleep(vIntDelayBefore);

                // Encontrar tab de Chrome por título

                foreach (var Handle in vObjBrowser.WindowHandles)
                {
                    // Esperar carga de página web
                    Misc.WaitForLoad(vObjBrowser.SwitchTo().Window(Handle));
                    if (vObjBrowser.SwitchTo().Window(Handle).Title.ToLower().Equals(vDicSelector["0"].Attribute.title.ToLower()))
                    {
                        // Coinciden los títulos
                        break;
                    }
                }

                // Crear los CssSelector
                String vStrSeleniumCssSelector = Misc.CreateSeleniumCssSelector(vDicSelector);

                String vStrJSCssSelector = Misc.CreateJSCssSelector(vDicSelector);

                // Validar si se construyeron selectores
                if (String.IsNullOrWhiteSpace(vStrSeleniumCssSelector) || String.IsNullOrWhiteSpace(vStrJSCssSelector))
                {
                    Console.WriteLine("No se construyeron selectores");
                    return;
                }

                // Enviar el click
                if (InjectJS)
                {
                    // Crear el objeto donde se ejecutará el script
                    IJavaScriptExecutor js = (IJavaScriptExecutor)vObjBrowser;

                    // Cambiar el tamaño del elemento usando JavaScript
                    js.ExecuteScript("document.querySelector(" + "\"" + vStrJSCssSelector + "\"" + ").style.height = '100%'");
                    js.ExecuteScript("document.querySelector(" + "\"" + vStrJSCssSelector + "\"" + ").style.width = '100%'");

                    if (String.IsNullOrWhiteSpace(vDicSelector["1"].Attribute.instance))
                    {
                        js.ExecuteScript("document.querySelector(" + "\"" + vStrJSCssSelector + "\"" + ").focus();");
                        js.ExecuteScript("document.querySelector(" + "\"" + vStrJSCssSelector + "\"" + ").click();");
                    }
                    else
                    {
                        js.ExecuteScript("document.querySelectorAll(" + "\"" + vStrJSCssSelector + "\"" + ")[" + (Int32.Parse(vDicSelector["1"].Attribute.instance) - 1).ToString() + "].click();");
                    }
                }
                else
                {

                    //try{
                    // Crear wait 
                    WebDriverWait wait = new WebDriverWait(vObjBrowser, TimeSpan.FromMilliseconds(vIntTimeoutMS));

                    // Esperar a que exista el elemento
                    Thread.Sleep(5000);

                    // Crear el elemento
                    IWebElement vEleElemento = vObjBrowser.FindElement(By.CssSelector(vStrSeleniumCssSelector));
                    vEleElemento.Click();
                    //}                catch (Exception)                {                    throw;                }

                }

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

    [Description("Click by item id en un elemento web mediante Selenium")]
    public class ClickById : CodeActivity
    {

        public ClickById()
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
        [DisplayName("Browser")]
        [Description("Objeto del navegador donde se realizará la actividad")]
        [RequiredArgument]
        public InArgument<Object> SeleniumBrowser { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [DisplayName("Item Id")]
        [Description("Valor del parametro Id del elemento en la página.")]
        public InArgument<String> ItemId { get; set; }

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
                IWebDriver vObjBrowser = (IWebDriver)SeleniumBrowser.Get(context);
                Int32 vIntDelayAfter = DelayAfter.Get(context);
                Int32 vIntDelayBefore = DelayBefore.Get(context);
                string ItemIdIn = ItemId.Get(context);

                //Dictionary<string, UiSelector> vDicSelector = SelectorManager.Funciones.TransformIntoUiSelector(Selector.Get(context));
                Int32 vIntTimeoutMS = TimeoutMS.Get(context);

                // Tiempo de espera antes de iniciar la actividad
                Thread.Sleep(vIntDelayBefore);

                // Crear wait 
                WebDriverWait wait = new WebDriverWait(vObjBrowser, TimeSpan.FromMilliseconds(vIntTimeoutMS));

                // Crear el objeto donde se ejecutará el script
                IJavaScriptExecutor js = (IJavaScriptExecutor)vObjBrowser;

                // Esperar a que exista el elemento
                //DONETSeleniumExtras  Nuget para quitar lo de obsoleto

                wait.Until(ExpectedConditions.ElementExists(By.Id(ItemIdIn)));

                // Crear el elemento
                IWebElement vEleElemento = vObjBrowser.FindElement(By.Id(ItemIdIn));

                // Cambiar el tamaño del elemento usando JavaScript
                //js.ExecuteScript("document.getElementById(" + "\"" + ItemIdIn + "\"" + ").style.height = '100%'");
                //js.ExecuteScript("document.getElementById(" + "\"" + ItemIdIn + "\"" + ").style.width = '100%'");

                // Enviar el click
                vEleElemento.Click();

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

   // [Description("Click by CssSelector en un elemento web mediante Selenium")]
    /*public class DoubleClick : CodeActivity
    {

        public DoubleClick()
        {
            DelayAfter = 300;
            DelayBefore = 1000;
            TimeoutMS = 30000;
            //InjectJS = false;
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
        [DisplayName("Browser")]
        [Description("Objeto del navegador donde se realizará la actividad")]
        public InArgument<IWebDriver> SeleniumBrowser { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [DisplayName("Selector")]
        [Description("Selector tipo Css")]
        public InArgument<String> Selector { get; set; }

        [Category("Input")]
        [DefaultValue(true)]
        [DisplayName("Time Out")]
        [Description("Tiempo de busqueda maxima del elemento.")]
        public InArgument<Int32> TimeoutMS { get; set; }

        [Category("Options")]
        [DisplayName("Inyectar Javascript?")]
        [Description("True para Realizar la actividad mediante inyección de Javascript")]
        public Boolean InjectJS { get; set; }

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
                IWebDriver vObjBrowser = (IWebDriver)SeleniumBrowser.Get(context);
                Int32 vIntDelayAfter = DelayAfter.Get(context);
                Int32 vIntDelayBefore = DelayBefore.Get(context);
                Dictionary<string, UiSelector> vDicSelector = SelectorManager.Funciones.TransformIntoUiSelector(Selector.Get(context));
                Int32 vIntTimeoutMS = TimeoutMS.Get(context);

                // Tiempo de espera antes de iniciar la actividad
                Thread.Sleep(vIntDelayBefore);

                // Encontrar tab de Chrome por título

                foreach (var Handle in vObjBrowser.WindowHandles)
                {
                    // Esperar carga de página web
                    Misc.WaitForLoad(vObjBrowser.SwitchTo().Window(Handle));
                    if (vObjBrowser.SwitchTo().Window(Handle).Title.ToLower().Equals(vDicSelector["0"].Attribute.title.ToLower()))
                    {
                        // Coinciden los títulos
                        break;
                    }
                }

                // Crear los CssSelector
                String vStrSeleniumCssSelector = Misc.CreateSeleniumCssSelector(vDicSelector);

                String vStrJSCssSelector = Misc.CreateJSCssSelector(vDicSelector);

                // Validar si se construyeron selectores
                if (String.IsNullOrWhiteSpace(vStrSeleniumCssSelector) || String.IsNullOrWhiteSpace(vStrJSCssSelector))
                {
                    Console.WriteLine("No se construyeron selectores");
                    return;
                }

                // Enviar el click
                if (InjectJS)
                {
                    // Crear el objeto donde se ejecutará el script
                    IJavaScriptExecutor js = (IJavaScriptExecutor)vObjBrowser;

                    // Cambiar el tamaño del elemento usando JavaScript
                    js.ExecuteScript("document.querySelector(" + "\"" + vStrJSCssSelector + "\"" + ").style.height = '100%'");
                    js.ExecuteScript("document.querySelector(" + "\"" + vStrJSCssSelector + "\"" + ").style.width = '100%'");

                    if (String.IsNullOrWhiteSpace(vDicSelector["1"].Attribute.instance))
                    {
                        js.ExecuteScript("document.querySelector(" + "\"" + vStrJSCssSelector + "\"" + ").focus();");
                        js.ExecuteScript("document.querySelector(" + "\"" + vStrJSCssSelector + "\"" + ").click();");
                    }
                    else
                    {
                        js.ExecuteScript("document.querySelectorAll(" + "\"" + vStrJSCssSelector + "\"" + ")[" + (Int32.Parse(vDicSelector["1"].Attribute.instance) - 1).ToString() + "].click();");
                    }
                }
                else
                {

                    //try{
                    // Crear wait 
                    WebDriverWait wait = new WebDriverWait(vObjBrowser, TimeSpan.FromMilliseconds(vIntTimeoutMS));

                    // Esperar a que exista el elemento
                    Thread.Sleep(5000);

                    // Crear el elemento
                    IWebElement vEleElemento = vObjBrowser.FindElement(By.CssSelector(vStrSeleniumCssSelector));
                    //vEleElemento
                    actions = 
                    new Actions(vObjBrowser).DoubleClick = new  Action(vObjBrowser);


                    //}                catch (Exception)                {                    throw;                }

                }

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
    }*/

    /// <summary>
    /// Descripción: Cierra un navegador de Chrome Previamente creado
    /// Autor: N/A
    /// Fecha Ultima Versión: 2020-06-23
    /// </summary>
    [Description("Cierra un navegador Previamente creado")]
    public class CloseBrowsser : CodeActivity
    {
        [Category("Input")]
        [DisplayName("Browser")]
        [Description("Objeto del navegador donde se realizará la actividad")]
        [RequiredArgument]
        public InArgument<IWebDriver> SeleniumBrowser { get; set; }

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
                IWebDriver vObjBrowser = (IWebDriver)SeleniumBrowser.Get(context);
                if (vObjBrowser != null)
                {
                    // Enviar orden de cierre Google Chrome
                    vObjBrowser.Quit();
                    Result.Set(context, true);
                }
                else
                {
                    Result.Set(context, false);
                }
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

    /// <summary>
    /// Descripción: Crea un nuevo navegador de Chrome o internet Explorer
    /// Autor: N/A
    /// Fecha Ultima Versión: 2020-06-23
    /// </summary>
    [Description("Crea un nuevo navegador de Chrome o internet Explorer")]
    public class CreateBrowser : CodeActivity
    {

        //Definir navegadores
        public enum Navegadores
        {
            Google_Chrome,
            Internet_Explorer,
        }

        [Category("In")]
        [Description("Navegador a utilizar")]
        [RequiredArgument]
        public Navegadores Navegador { get; set; }

        [Category("Output")]
        [DisplayName("Browser")]
        [Description("Driver el cual será solicitado para todas las actividades Selenium")]
        [RequiredArgument]
        public OutArgument<IWebDriver> SeleniumBrowser { get; set; }

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
                //Falta el try - catch
                string vNavegador = Navegador.ToString();

                //Definir variable
                IWebDriver vObjBrowser;

                //Crear WebDriver
                if (vNavegador == "Google_Chrome")
                {
                    //IWebDriver vObjCHDriver = (IWebDriver)new ChromeDriver();
                    //IWebDriver vObjCHDriver = (IWebDriver)new ChromeDriver(new ChromeOptions() { BrowserVersion= "79.0.3945.117" });
                    vObjBrowser = (IWebDriver)new ChromeDriver();

                    // Maximizar la ventana
                    vObjBrowser.Manage().Window.Maximize();

                    //Set Salidas:
                    Result.Set(context, true);
                    SeleniumBrowser.Set(context, vObjBrowser);
                }
                else if (vNavegador == "Internet_Explorer")
                {
                    vObjBrowser = new InternetExplorerDriver();

                    // Maximizar la ventana
                    vObjBrowser.Manage().Window.Maximize();

                    //Set Salidas:
                    Result.Set(context, true);
                    SeleniumBrowser.Set(context, vObjBrowser);
                }
                else
                {
                    Console.WriteLine("Error, por favor elegir entre Google Chrome o Internet Explorer");
                    //Set Salidas:
                    Result.Set(context, false);
                }
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

    [Description("Validar si existe un elemento by CssSelector de un elemento web mediante Selenium")]
    public class ElementExist : CodeActivity
    {
        public ElementExist()
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
        [DisplayName("Browser")]
        [Description("Objeto del navegador donde se realizará la actividad")]
        public InArgument<IWebDriver> SeleniumBrowser { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [DisplayName("Selector")]
        [Description("Selector tipo Css")]
        public InArgument<String> Selector { get; set; }

        [Category("Input")]
        [DefaultValue(true)]
        [DisplayName("Time Out")]
        [Description("Tiempo de busqueda maxima del elemento.")]
        public InArgument<Int32> TimeoutMS { get; set; }

        [Category("Options")]
        [DisplayName("Inyectar Javascript?")]
        [Description("True para Realizar la actividad mediante inyección de Javascript")]
        public Boolean InjectJS { get; set; }

        [Category("Output")]
        [RequiredArgument]
        [DisplayName("Existe Alerta")]
        [Description("Resultado si existe o no.")]
        public OutArgument<Boolean> Exist { get; set; }

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
                IWebDriver vObjBrowser = (IWebDriver)SeleniumBrowser.Get(context);
                Int32 vIntDelayAfter = DelayAfter.Get(context);
                Int32 vIntDelayBefore = DelayBefore.Get(context);
                Dictionary<string, UiSelector> vDicSelector = SelectorManager.Funciones.TransformIntoUiSelector(Selector.Get(context));
                Int32 vIntTimeoutMS = TimeoutMS.Get(context);
                DateTime vDatInicio = System.DateTime.Now;

                // Tiempo de espera antes de iniciar la actividad
                Thread.Sleep(vIntDelayBefore);

                // Crear los CssSelector
                String vStrSeleniumCssSelector = Misc.CreateSeleniumCssSelector(vDicSelector);

                String vStrJSCssSelector = Misc.CreateJSCssSelector(vDicSelector);

                // Validar si se construyeron selectores
                if (String.IsNullOrWhiteSpace(vStrSeleniumCssSelector) || String.IsNullOrWhiteSpace(vStrJSCssSelector))
                {
                    Console.WriteLine("No se construyeron selectores");
                    return;
                }

                if (InjectJS)
                {
                    // Crear el objeto donde se ejecutará el script
                    IJavaScriptExecutor js = (IJavaScriptExecutor)vObjBrowser;
                    int a = 0;
                    do
                    {
                        a = Convert.ToInt32(js.ExecuteScript("var ElementsFound = document.querySelectorAll(" + "\"" + vStrJSCssSelector + "\"" + ");" +
                            "if (ElementsFound == null){" +
                            "     return 0;" +
                            "} else {" +
                            "     return ElementsFound.length;" +
                            "}"));
                    } while (a == 0 && (int)(System.DateTime.Now - vDatInicio).TotalMilliseconds < vIntTimeoutMS);

                    if (String.IsNullOrWhiteSpace(vDicSelector["1"].Attribute.instance))
                    {
                        Exist.Set(context, a > 0 ? true : false);
                    }
                    else
                    {
                        Exist.Set(context, a >= Int32.Parse(vDicSelector["1"].Attribute.instance) ? true : false);
                    }
                }
                else
                {
                    // Crear wait 
                    WebDriverWait wait = new WebDriverWait(vObjBrowser, TimeSpan.FromMilliseconds(vIntTimeoutMS));

                    // Esperar a que exista el elemento
                    wait.Until(ExpectedConditions.ElementExists(By.CssSelector(vStrSeleniumCssSelector)));

                    // Crear el elemento
                    try
                    {
                        IWebElement vEleElemento = vObjBrowser.FindElement(By.CssSelector(vStrSeleniumCssSelector));
                        Exist.Set(context, true);
                    }
                    catch (Exception)
                    {
                        Exist.Set(context, false);
                    }
                }
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

    /// <summary>
    /// Descripción: Capturar string basado en query ejecutado en la página actual
    /// Autor: N/A
    /// Fecha Ultima Versión: 2020-06-23
    /// </summary>
    [Description("Capturar texto de un elemento basado en un query ejecutado en la página actual")]
    public class GetResultQuery : CodeActivity
    {
        public GetResultQuery()
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
        [DisplayName("Browser")]
        [Description("Objeto del navegador donde se realizará la actividad")]
        public InArgument<IWebDriver> SeleniumBrowser { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [DisplayName("Query")]
        [Description("Query que se va a inyectar tipo Javascript")]
        public InArgument<String> query { get; set; }

        [Category("Input")]
        [DefaultValue(true)]
        [DisplayName("Time Out")]
        [Description("Tiempo de busqueda maxima del elemento.")]
        public InArgument<Int32> TimeoutMS { get; set; }

        [Category("Output")]
        [RequiredArgument]
        [DisplayName("Value")]
        [Description("Texto capturado del elemento")]
        public OutArgument<String> Text { get; set; }

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
                IWebDriver vObjBrowser = (IWebDriver)SeleniumBrowser.Get(context);
                Int32 vIntDelayAfter = DelayAfter.Get(context);
                Int32 vIntDelayBefore = DelayBefore.Get(context);
                String vQuery = query.Get(context);
                Int32 vIntTimeoutMS = TimeoutMS.Get(context);

                // Tiempo de espera antes de iniciar la actividad
                Thread.Sleep(vIntDelayBefore);

                // Crear el objeto donde se ejecutará el script
                IJavaScriptExecutor js = (IJavaScriptExecutor)vObjBrowser;
                // Cambiar el tamaño del elemento usando JavaScript
                //js.ExecuteScript("document.querySelector(" + "\"" + vStrJSCssSelector + "\"" + ").value = \"" + vStrText + "\"");

                js.ExecuteScript(vQuery);

                String vStrInnerText = (String)js.ExecuteScript(
                           "var ElementFound = " + vQuery + ";" +
                           "if (ElementFound == null){" +
                           "     return \"\";" +
                           "} else {" +
                           "     return ElementFound.innerText;" +
                           "}");
                Text.Set(context, vStrInnerText);


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

    [Description("Capturar texto by CssSelector de un elemento web mediante Selenium")]
    public class GetText : CodeActivity
    {
        public GetText()
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
        [DisplayName("Browser")]
        [Description("Objeto del navegador donde se realizará la actividad")]
        public InArgument<IWebDriver> SeleniumBrowser { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [DisplayName("Selector")]
        [Description("Selector tipo Css")]
        public InArgument<String> Selector { get; set; }

        [Category("Input")]
        [DefaultValue(true)]
        [DisplayName("Time Out")]
        [Description("Tiempo de busqueda maxima del elemento.")]
        public InArgument<Int32> TimeoutMS { get; set; }

        [Category("Options")]
        [DisplayName("Inyectar Javascript?")]
        [Description("True para Realizar la actividad mediante inyección de Javascript")]
        public Boolean InjectJS { get; set; }

        [Category("Output")]
        [RequiredArgument]
        [DisplayName("Value")]
        [Description("Texto capturado del elemento")]
        public OutArgument<String> Text { get; set; }

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
                IWebDriver vObjBrowser = (IWebDriver)SeleniumBrowser.Get(context);
                Int32 vIntDelayAfter = DelayAfter.Get(context);
                Int32 vIntDelayBefore = DelayBefore.Get(context);
                Dictionary<string, UiSelector> vDicSelector = SelectorManager.Funciones.TransformIntoUiSelector(Selector.Get(context));
                String vStrText = Text.Get(context);
                Int32 vIntTimeoutMS = TimeoutMS.Get(context);
                DateTime vDatInicio = System.DateTime.Now;

                // Tiempo de espera antes de iniciar la actividad
                Thread.Sleep(vIntDelayBefore);

                // Crear los CssSelector
                String vStrSeleniumCssSelector = Misc.CreateSeleniumCssSelector(vDicSelector);

                String vStrJSCssSelector = Misc.CreateJSCssSelector(vDicSelector);

                // Validar si se construyeron selectores
                if (String.IsNullOrWhiteSpace(vStrSeleniumCssSelector) || String.IsNullOrWhiteSpace(vStrJSCssSelector))
                {
                    Console.WriteLine("No se construyeron selectores");
                    return;
                }

                if (InjectJS)
                {
                    // Crear el objeto donde se ejecutará el script
                    IJavaScriptExecutor js = (IJavaScriptExecutor)vObjBrowser;

                    if (String.IsNullOrWhiteSpace(vDicSelector["1"].Attribute.instance))
                    {
                        String vStrInnerText = (String)js.ExecuteScript(
                            "var ElementFound = document.querySelector(" + "\"" + vStrJSCssSelector + "\"" + ");" +
                            "if (ElementFound == null){" +
                            "     return \"\";" +
                            "} else {" +
                            "     return ElementFound.innerText;" +
                            "}");
                        Text.Set(context, vStrInnerText);
                    }
                    else
                    {
                        String vStrInnerText = (String)js.ExecuteScript(
                            "var ElementsFound = document.querySelectorAll(" + "\"" + vStrJSCssSelector + "\"" + ");" +
                            "if (ElementsFound == null || ElementsFound.length < parseInt(arguments[0])+1){" +
                            "     return \"\";" +
                            "} else {" +
                            "     return ElementsFound[parseInt(arguments[0])].innerText;" +
                            "}", (Int32.Parse(vDicSelector["1"].Attribute.instance) - 1).ToString());
                        Text.Set(context, vStrInnerText);
                    }
                }
                else
                {
                    // Crear wait 
                    WebDriverWait wait = new WebDriverWait(vObjBrowser, TimeSpan.FromMilliseconds(vIntTimeoutMS));

                    // Esperar a que exista el elemento
                    wait.Until(ExpectedConditions.ElementExists(By.CssSelector(vStrSeleniumCssSelector)));

                    // Crear el elemento
                    try
                    {
                        IWebElement vEleElemento = vObjBrowser.FindElement(By.CssSelector(vStrSeleniumCssSelector));
                        Text.Set(context, vEleElemento.Text);
                    }
                    catch (Exception)
                    {
                        Text.Set(context, "");
                    }
                }
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

    [Description("Capturar texto by item id de un elemento web mediante Selenium")]
    public class GetAttributeById : CodeActivity
    {
        public GetAttributeById()
        {
            DelayAfter = 300;
            DelayBefore = 1000;
            TimeoutMS = 30000;
            Attribute = "value";
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
        [DisplayName("Browser")]
        [Description("Objeto del navegador donde se realizará la actividad")]
        public InArgument<Object> SeleniumBrowser { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [DisplayName("Attributo a capturar")]
        [Description("Attributo a capturar ejemplo: value, class, etc.")]
        public InArgument<String> Attribute { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [DisplayName("Item Id")]
        [Description("Valor del parametro Id del elemento en la página.")]
        public InArgument<String> ItemId { get; set; }

        [Category("Input")]
        [DefaultValue(true)]
        [DisplayName("Time Out")]
        [Description("Tiempo de busqueda maxima del elemento.")]
        public InArgument<Int32> TimeoutMS { get; set; }

        [Category("Output")]
        [DisplayName("Value")]
        [Description("Texto capturado del atributo del elemento")]
        public OutArgument<string> Valor { get; set; }

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
                IWebDriver vObjBrowser = (IWebDriver)SeleniumBrowser.Get(context);
                Int32 vIntDelayAfter = DelayAfter.Get(context);
                Int32 vIntDelayBefore = DelayBefore.Get(context);
                String vItemIdIn = ItemId.Get(context);
                String vAttribute = Attribute.Get(context);
                Int32 vIntTimeoutMS = TimeoutMS.Get(context);

                // Tiempo de espera antes de iniciar la actividad
                Thread.Sleep(vIntDelayBefore);

                // Crear wait 
                WebDriverWait wait = new WebDriverWait(vObjBrowser, TimeSpan.FromMilliseconds(vIntTimeoutMS));

                // Crear el objeto donde se ejecutará el script
                IJavaScriptExecutor js = (IJavaScriptExecutor)vObjBrowser;

                // Esperar a que exista el elemento
                wait.Until(ExpectedConditions.ElementExists(By.Id(vItemIdIn)));

                // Crear el elemento
                IWebElement vEleElemento = vObjBrowser.FindElement(By.Id(vItemIdIn));

                // Leer con InnerHTML
                string Lectura = vEleElemento.GetAttribute(vAttribute);
                //Console.WriteLine(Lectura);

                //Send Output
                Valor.Set(context, Lectura);

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

    [Description("Capturar URL de la pestaña activa en el navegador seleccionado")]
    public class GetURL : CodeActivity
    {

        [Category("Input")]
        [RequiredArgument]
        [DisplayName("Browser")]
        [Description("Objeto del navegador donde se realizará la actividad")]
        public InArgument<Object> SeleniumBrowser { get; set; }

        [Category("Output")]
        [RequiredArgument]
        [DisplayName("URL")]
        [Description("Dirección URL capturada por la actividad.")]
        public OutArgument<String> URL { get; set; }

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
                IWebDriver vObjBrowser = (IWebDriver)SeleniumBrowser.Get(context);

                // Crear el objeto donde se ejecutará el script
                IJavaScriptExecutor js = (IJavaScriptExecutor)vObjBrowser;

                string Resultado = (string)js.ExecuteScript("return location.href");

                //Enviar Resultado URL
                URL.Set(context, Resultado);
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

    /// <summary>
    /// Descripción: Navegar a una página especifica
    /// Autor: N/A
    /// Fecha Ultima Versión: 2020-06-23
    /// </summary>
    [Description("Navegar a una página especifica")]
    public class NavigateTo : CodeActivity
    {
        public NavigateTo()
        {
            OpenNewTab = false;
        }

        [Category("Input")]
        [DisplayName("Browser")]
        [Description("Objeto del navegador donde se realizará la actividad")]
        [RequiredArgument]
        public InArgument<IWebDriver> SeleniumBrowser { get; set; }

        [Category("Input")]
        [DisplayName("URL")]
        [Description("URL a la que se quiere navegar. Debe contener todal a dirección, desde https://")]
        [RequiredArgument]
        public InArgument<String> URL { get; set; }

        [Category("Options")]
        [DisplayName("Open new tab?")]
        [Description("Desea abrir la URL en una ventana nueva?")]
        [DefaultValue(true)]
        public Boolean OpenNewTab { get; set; }

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
                IWebDriver vObjBrowser = (IWebDriver)SeleniumBrowser.Get(context);
                String vStrURL = URL.Get(context);
                IJavaScriptExecutor js = (IJavaScriptExecutor)vObjBrowser;

                // Navegar a URL
                if (OpenNewTab)
                {
                    js.ExecuteScript(String.Format("window.open('{0}', '_blank');", vStrURL));
                }
                else
                {
                    js.ExecuteScript(String.Format("window.self.location.assign('{0}');", vStrURL));
                }

                // Maximizar la ventana
                vObjBrowser.Manage().Window.Maximize();
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

    /// <summary>
    /// Descripción: Ejecutar Query Javascript en la página actual
    /// Autor: N/A
    /// Fecha Ultima Versión: 2020-06-23
    /// </summary>
    [Description("Ejecutar Query Javascript en la página actual")]
    public class SendQuery : CodeActivity
    {
        public SendQuery()
        {
            DelayAfter = 300;
            DelayBefore = 1000;
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
        [DisplayName("Browser")]
        [Description("Objeto del navegador donde se realizará la actividad")]
        public InArgument<IWebDriver> SeleniumBrowser { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [DisplayName("Query")]
        [Description("Query que se va a inyectar tipo Javascript")]
        public InArgument<String> query { get; set; }

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
                IWebDriver vObjBrowser = (IWebDriver)SeleniumBrowser.Get(context);
                Int32 vIntDelayAfter = DelayAfter.Get(context);
                Int32 vIntDelayBefore = DelayBefore.Get(context);
                String vQuery = query.Get(context);

                // Tiempo de espera antes de iniciar la actividad
                Thread.Sleep(vIntDelayBefore);

                // Crear el objeto donde se ejecutará el script
                IJavaScriptExecutor js = (IJavaScriptExecutor)vObjBrowser;
                // Cambiar el tamaño del elemento usando JavaScript
                //js.ExecuteScript("document.querySelector(" + "\"" + vStrJSCssSelector + "\"" + ").value = \"" + vStrText + "\"");

                js.ExecuteScript(vQuery);

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

    [Description("Escribir texto by CssSelector en un elemento web mediante Selenium")]
    public class TypeInto : CodeActivity
    {
        public TypeInto()
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
        [DisplayName("Browser")]
        [Description("Objeto del navegador donde se realizará la actividad")]
        public InArgument<IWebDriver> SeleniumBrowser { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [DisplayName("Texto")]
        [Description("Texto que se quiere ingresar en el elemento.")]
        public InArgument<String> Text { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [DisplayName("Selector")]
        [Description("Selector tipo Css")]
        public InArgument<String> Selector { get; set; }

        [Category("Input")]
        [DefaultValue(true)]
        [DisplayName("Time Out")]
        [Description("Tiempo de busqueda maxima del elemento.")]
        public InArgument<Int32> TimeoutMS { get; set; }

        [Category("Options")]
        [DisplayName("Inyectar Javascript?")]
        [Description("True para Realizar la actividad mediante inyección de Javascript")]
        public Boolean InjectJS { get; set; }

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
                IWebDriver vObjBrowser = (IWebDriver)SeleniumBrowser.Get(context);
                Int32 vIntDelayAfter = DelayAfter.Get(context);
                Int32 vIntDelayBefore = DelayBefore.Get(context);
                Dictionary<string, UiSelector> vDicSelector = SelectorManager.Funciones.TransformIntoUiSelector(Selector.Get(context));
                String vStrText = Text.Get(context);
                Int32 vIntTimeoutMS = TimeoutMS.Get(context);

                // Tiempo de espera antes de iniciar la actividad
                Thread.Sleep(vIntDelayBefore);

                // Encontrar tab de Chrome por título

                foreach (var Handle in vObjBrowser.WindowHandles)
                {
                    // Esperar carga de página web
                    Misc.WaitForLoad(vObjBrowser.SwitchTo().Window(Handle));
                    if (vObjBrowser.SwitchTo().Window(Handle).Title.ToLower().Equals(vDicSelector["0"].Attribute.title.ToLower()))
                    {
                        // Coinciden los títulos
                        break;
                    }
                }

                // Crear los CssSelector
                String vStrSeleniumCssSelector = Misc.CreateSeleniumCssSelector(vDicSelector);

                String vStrJSCssSelector = Misc.CreateJSCssSelector(vDicSelector);

                // Validar si se construyeron selectores
                if (String.IsNullOrWhiteSpace(vStrSeleniumCssSelector) || String.IsNullOrWhiteSpace(vStrJSCssSelector))
                {
                    Console.WriteLine("No se construyeron selectores");
                    return;
                }

                if (InjectJS)
                {
                    // Crear el objeto donde se ejecutará el script
                    IJavaScriptExecutor js = (IJavaScriptExecutor)vObjBrowser;
                    // Cambiar el tamaño del elemento usando JavaScript
                    //js.ExecuteScript("document.querySelector(" + "\"" + vStrJSCssSelector + "\"" + ").value = \"" + vStrText + "\"");

                    if (String.IsNullOrWhiteSpace(vDicSelector["1"].Attribute.instance))
                    {
                        js.ExecuteScript(
                            "var ElementFound = document.querySelector(" + "\"" + vStrJSCssSelector + "\"" + ");" +
                            "if (ElementFound != null){" +
                            "     ElementFound.value = \"" + vStrText + "\";" +
                            "}");
                    }
                    else
                    {
                        js.ExecuteScript(
                            "var ElementsFound = document.querySelectorAll(" + "\"" + vStrJSCssSelector + "\"" + ");" +
                            "if (ElementsFound != null && ElementsFound.length >= parseInt(arguments[0])+1){" +
                            "     ElementsFound[parseInt(arguments[0])].value = \"" + vStrText + "\";" +
                            "}", (Int32.Parse(vDicSelector["1"].Attribute.instance) - 1).ToString());
                    }
                }
                else
                {
                    // Crear wait 
                    WebDriverWait wait = new WebDriverWait(vObjBrowser, TimeSpan.FromMilliseconds(vIntTimeoutMS));

                    // Esperar a que exista el elemento
                    wait.Until(ExpectedConditions.ElementExists(By.CssSelector(vStrSeleniumCssSelector)));

                    // Crear el elemento
                    try
                    {
                        IWebElement vEleElemento = vObjBrowser.FindElement(By.CssSelector(vStrSeleniumCssSelector));

                        vEleElemento.Clear();
                        vEleElemento.SendKeys(vStrText);
                    }
                    catch (Exception)
                    {

                    }

                }

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

    [Description("Escribir texto by item id en un elemento web mediante Selenium")]
    public class TypeIntoById : CodeActivity
    {
        public TypeIntoById()
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
        [DisplayName("Browser")]
        [Description("Objeto del navegador donde se realizará la actividad")]
        public InArgument<Object> SeleniumBrowser { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [DisplayName("Texto")]
        [Description("Texto que se quiere ingresar en el elemento.")]
        public InArgument<String> Text { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [DisplayName("Item Id")]
        [Description("Valor del parametro Id del elemento en la página.")]
        public InArgument<String> ItemId { get; set; }

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
                IWebDriver vObjBrowser = (IWebDriver)SeleniumBrowser.Get(context);
                Int32 vIntDelayAfter = DelayAfter.Get(context);
                Int32 vIntDelayBefore = DelayBefore.Get(context);
                String ItemIdIn = ItemId.Get(context);
                String vStrText = Text.Get(context);
                Int32 vIntTimeoutMS = TimeoutMS.Get(context);

                // Tiempo de espera antes de iniciar la actividad
                Thread.Sleep(vIntDelayBefore);

                // Crear wait 
                WebDriverWait wait = new WebDriverWait(vObjBrowser, TimeSpan.FromMilliseconds(vIntTimeoutMS));

                // Crear el objeto donde se ejecutará el script
                IJavaScriptExecutor js = (IJavaScriptExecutor)vObjBrowser;

                // Esperar a que exista el elemento
                wait.Until(ExpectedConditions.ElementExists(By.Id(ItemIdIn)));

                // Crear el elemento
                IWebElement vEleElemento = vObjBrowser.FindElement(By.Id(ItemIdIn));

                // Cambiar el tamaño del elemento usando JavaScript
                js.ExecuteScript("document.getElementById(" + "\"" + ItemIdIn + "\"" + ").value = \"" + vStrText + "\"");

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


    //Funciones
    public class Misc
    {
        public static void WaitForLoad(IWebDriver driver, Int32 timeoutSec = 60)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            WebDriverWait wait = new WebDriverWait(driver, new TimeSpan(0, 0, timeoutSec));
            wait.Until(d => js.ExecuteScript("return document.readyState").Equals("complete"));
            //new WebDriverWait(driver, MyDefaultTimeout).Until(d => ((IJavaScriptExecutor)d).ExecuteScript("return document.readyState").Equals("complete"));
        }

        public static string CreateJSCssSelector(Dictionary<string, UiSelector> vDicSelector)
        {
            String vStrJSCssSelector = "";
            UiSelector vObjChildSelector = vDicSelector["1"];

            if (!String.IsNullOrWhiteSpace(vObjChildSelector.Attribute.className) || !String.IsNullOrWhiteSpace(vObjChildSelector.Attribute.id) || !String.IsNullOrWhiteSpace(vObjChildSelector.Attribute.name) || !String.IsNullOrWhiteSpace(vObjChildSelector.Attribute.href))
            {
                // Si existe Class
                if (!String.IsNullOrWhiteSpace(vObjChildSelector.Attribute.className)) { vStrJSCssSelector = vStrJSCssSelector + " class='" + vObjChildSelector.Attribute.className + "'"; }
                // Si existe id
                if (!String.IsNullOrWhiteSpace(vObjChildSelector.Attribute.id)) { vStrJSCssSelector = vStrJSCssSelector + " id='" + vObjChildSelector.Attribute.id + "'"; }
                // Si existe name
                if (!String.IsNullOrWhiteSpace(vObjChildSelector.Attribute.name)) { vStrJSCssSelector = vStrJSCssSelector + " name='" + vObjChildSelector.Attribute.name + "'"; }
                // Si existe href
                if (!String.IsNullOrWhiteSpace(vObjChildSelector.Attribute.href)) { vStrJSCssSelector = vStrJSCssSelector + " href='" + vObjChildSelector.Attribute.href + "'"; }
                // Si existe role
                if (!String.IsNullOrWhiteSpace(vObjChildSelector.Attribute.role)) { vStrJSCssSelector = vStrJSCssSelector + " role='" + vObjChildSelector.Attribute.role + "'"; }

                // Agregar llaver
                vStrJSCssSelector = "[" + vStrJSCssSelector + "]";

                // Agregar tag name
                vStrJSCssSelector = vObjChildSelector.Attribute.tag + vStrJSCssSelector;

                // Agregar parenId
                if (!String.IsNullOrWhiteSpace(vObjChildSelector.Attribute.parentid)) { vStrJSCssSelector = "#" + vObjChildSelector.Attribute.parentid + " " + vStrJSCssSelector; }

                // Agregar parentClass solo en caso de que el parentId no exista
                if (String.IsNullOrWhiteSpace(vObjChildSelector.Attribute.parentid) && !String.IsNullOrWhiteSpace(vObjChildSelector.Attribute.parentclassName)) { vStrJSCssSelector = "." + vObjChildSelector.Attribute.parentclassName + " " + vStrJSCssSelector; }
            }

            return vStrJSCssSelector;
        }

        // .FindElement or .FindElements by .CssSelector must be builded by following structure:
        // tagName[attributename=attributeValue]
        // Example 1: input[id=email]
        // Example 2: input[name=email][type=text]
        // For more information go to : https://www.seleniumeasy.com/selenium-tutorials/css-selectors-tutorial-for-selenium-with-examples
        
        public static string CreateSeleniumCssSelector(Dictionary<string, UiSelector> vDicSelector)
        {
            String vStrSeleniumCssSelector ="";
            UiSelector vObjChildSelector = vDicSelector["1"];

            // Si existe Class
            if (!String.IsNullOrWhiteSpace(vObjChildSelector.Attribute.className)) { vStrSeleniumCssSelector = vStrSeleniumCssSelector + "[class=" + vObjChildSelector.Attribute.className + "]"; }
            // Si existe id
            if (!String.IsNullOrWhiteSpace(vObjChildSelector.Attribute.id)) { vStrSeleniumCssSelector = vStrSeleniumCssSelector + "[id=" + vObjChildSelector.Attribute.id + "]"; }
            // Si existe name
            if (!String.IsNullOrWhiteSpace(vObjChildSelector.Attribute.name)) { vStrSeleniumCssSelector = vStrSeleniumCssSelector + "[name=" + vObjChildSelector.Attribute.name + "]"; }
            // Si existe href
            if (!String.IsNullOrWhiteSpace(vObjChildSelector.Attribute.href)) { vStrSeleniumCssSelector = vStrSeleniumCssSelector + "[href=" + vObjChildSelector.Attribute.href + "]"; }

            // Agregar nombre de tag
            vStrSeleniumCssSelector = vObjChildSelector.Attribute.tag + vStrSeleniumCssSelector;

            return vStrSeleniumCssSelector;
        }

    }

}