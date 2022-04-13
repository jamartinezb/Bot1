using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using System;
using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Mouse
{
    [Description("Permite administrar las funciones de los eventos del mouse.")]
    public class Activites_Mouse : CodeActivity

    {
        protected override void Execute(CodeActivityContext context)
        {
            throw new NotImplementedException();
        }
    }

    public class Activities_MouseOver : CodeActivity
    {
        public Activities_MouseOver()
        {
            EjecutarClic = false;
        }

        #region Propiedades
        [Category("Input")]
        [RequiredArgument]
        [Description("Driver del Navegador a Utilizar.")]
        public InArgument<IWebDriver> Navegador { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [Description("Elemento Xpath que se quiere interactuar con los eventos del mouse.")]
        public InArgument<string> ElementoABuscar { get; set; }

        [Category("Input")]
        [Description("Permite ejecutar un Clic en el elemento buscado.")]
        public InArgument<bool> EjecutarClic { get; set; }

        #endregion

        protected override void Execute(CodeActivityContext context)
        {
            IWebDriver vNavegador = (IWebDriver)Navegador.Get(context);
            string vElemento = ElementoABuscar.Get(context); 

            //Instantiate Action Class        
            Actions actions = new Actions(vNavegador);
            //Retrieve WebElement 'Music' to perform mouse hover 
            IWebElement oElemento = vNavegador.FindElement(By.XPath(vElemento));
            //Mouse hover menuOption 'Music'
            actions.MoveToElement(oElemento).Perform();

            if (EjecutarClic.Get(context))
            {
                oElemento.Click();
            }
        }
    }
}
