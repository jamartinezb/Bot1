using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Activities;
using System.ComponentModel;
using SelectorManager;
using System.Threading; // Utilizado para agregar delays
using System.Runtime.InteropServices;  // Utilizado para agregar parámetros opcionales en las funciones
using System.Diagnostics;
using System.Configuration;



namespace EnvironmentDevelop
{
    /// <summary>
    /// Descripción: Cerrar un proceso en el equipo.
    /// Autor: Carlos Cardona
    /// Fecha Ultima Versión: 2020-06-16
    /// </summary>
    [Description("Finalizar un proceso del equipo")]
    public class CloseProcess : CodeActivity
    {
        
        [Category("In")]
        [Description("Nombre del programa que se quiere cerrar.")]
        [RequiredArgument]
        public InArgument<string> Program { get; set; }

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
                String Proceso = Program.Get(context);
                Process[] chromeInstances = Process.GetProcessesByName(Proceso);

                foreach (Process p in chromeInstances)
                    p.Kill();

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

    [Description("Finalizar el flujo del robot.")]
    public class ExitProgram : CodeActivity
    {

        protected override void Execute(CodeActivityContext context)
        {

            Environment.Exit(1);
        }



    }

    /// <summary>
    /// Descripción: Eliminar Archivo
    /// Autor: Carlos Cardona
    /// Fecha Ultima Versión: 2020-06-16
    /// </summary>
    [Description("Eliminar un archivo")]
    public class FileDelete : CodeActivity
    {

        [Category("In")]
        [Description("Ruta completa del archivo a eliminar.")]
        [DisplayName("Ruta del archivo")]
        [RequiredArgument]
        public InArgument<string> PathFile { get; set; }

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
                String vPathFile = PathFile.Get(context);

                if (File.Exists(vPathFile))
                {
                    File.Delete(vPathFile);
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
    /// Descripción: Validar existencia de un archivo
    /// Autor: Carlos Cardona
    /// Fecha Ultima Versión: 2020-06-16
    /// </summary>
    [Description("Validar existencia de un archivo en una carpeta especifica")]
    public class FileExist : CodeActivity
    {

        [Category("In")]
        [Description("Ruta completa del archivo para validar.")]
        [DisplayName("Ruta del archivo")]
        [RequiredArgument]
        public InArgument<string> PathFile { get; set; }

        [Category("Output")]
        [DisplayName("Respuesta")]
        [Description("Respuesta de la actividad.")]
        public OutArgument<bool> Respuesta { get; set; }

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
                String vPathFile = PathFile.Get(context);

                if (File.Exists(vPathFile))
                {
                    Respuesta.Set(context, true);
                }
                else
                {
                    Respuesta.Set(context, false);
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
    /// Descripción: Mover archivo
    /// Autor: Carlos Cardona
    /// Fecha Ultima Versión: 2020-06-16
    /// </summary>
    [Description("Mover de ubicación un archivo")]
    public class FileMove : CodeActivity
    {

        [Category("In")]
        [Description("Ruta completa del archivo para validar.")]
        [DisplayName("Ruta ubicación inicial archivo")]
        [RequiredArgument]
        public InArgument<string> PathFileEntrada { get; set; }

        [Category("In")]
        [Description("Ruta completa del archivo para validar.")]
        [DisplayName("Ruta ubicación final archivo")]
        [RequiredArgument]
        public InArgument<string> PathFileSalida { get; set; }

        [Category("In")]
        [Description("Desea Eliminar el archivo de la ruta inicial?")]
        [DisplayName("Eliminar Archivo Ruta inicial")]
        [RequiredArgument]
        public InArgument<bool> Eliminar { get; set; }

        [Category("Output")]
        [DisplayName("Respuesta")]
        [Description("Respuesta de la actividad.")]
        public OutArgument<bool> Respuesta { get; set; }

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
                String vPathFileEntrada = PathFileEntrada.Get(context);
                String vPathFileSalida = PathFileSalida.Get(context);
                bool vEliminar = Eliminar.Get(context);

                if (!File.Exists(vPathFileSalida))
                {
                    File.Copy(vPathFileEntrada, vPathFileSalida, true);
                    if (vEliminar == true)
                    {
                        File.Delete(vPathFileEntrada);
                    }
                    Respuesta.Set(context, true);
                }
                else
                {
                    Respuesta.Set(context, false);
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
    /// Descripción: Capturar configuración de App.Config. Util para IdRobot
    /// Autor: Carlos Cardona
    /// Fecha Ultima Versión: 2020-06-16
    /// </summary>
    /*[Description("Capturar Configuración de App.Config .")]
    public class GetConfigurationSettingString : CodeActivity
    {
        public GetConfigurationSettingString()
        {

        }

        [Category("Input")]
        [RequiredArgument]
        [Description("Variable to get from config file.")]
        public InArgument<string> ConfigVariable { get; set; }

        [Category("Output")]
        [DisplayName("Valor obtenido.")]
        [Description("Valor de la variable obtenido del APP.Config .")]
        public OutArgument<String> ValorResult { get; set; }

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
                Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                string configVariable = config.AppSettings.Settings[ConfigVariable.Get(context)].Value.ToString();
                ValorResult.Set(context, configVariable);
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

    [Description("Asignar texto al portapapeles")]
    public class SetTextToClipboard : CodeActivity
    {

        [Category("Input")]
        [RequiredArgument]
        public InArgument<String> TextToCopy { get; set; }

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
                String vStrText = TextToCopy.Get(context);

                // Ejecutar la acción
                System.Windows.Forms.Clipboard.SetText(vStrText);

                // Imprimir el contenido del clipboard:
                //Console.WriteLine("Copy Text to Clipboard" + vStrText);
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
