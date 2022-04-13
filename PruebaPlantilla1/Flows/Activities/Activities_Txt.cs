using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Activities;
using System.ComponentModel;
using System.IO;

namespace Txt
{
    /// <summary>
    /// Descripción: Crear Archivo Txt
    /// Autor: Carlos Cardona
    /// Fecha Ultima Versión: 2020-06-19
    /// </summary>
    [Description("Crear un archivo tipo Txt")]
    public class CreateTxt : CodeActivity
    {

        [Category("In")]
        [Description("Ruta completa del archivo que se creará")]
        [DisplayName("Ruta del archivo")]
        [RequiredArgument]
        public InArgument<string> PathFile { get; set; }

        [Category("In")]
        [Description("Texto a escribir")]
        [RequiredArgument]
        public InArgument<string> Texto { get; set; }

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
                String vTexto = Texto.Get(context);

                if (!File.Exists(vPathFile))
                {
                    // Create a file to write to.
                    using (StreamWriter sw = File.CreateText(vPathFile))
                    {
                        sw.WriteLine(vTexto);
                    }
                }
                else
                {
                    Console.WriteLine("El archivo ya existe");
                    Result.Set(context, false);
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
    /// Descripción: Insertar texto en un Archivo Txt
    /// Autor: Carlos Cardona
    /// Fecha Ultima Versión: 2020-06-19
    /// </summary>
    [Description("Adicionar texto nuevo en un archivo tipo Txt")]
    public class InsertTextTxt : CodeActivity
    {

        [Category("In")]
        [Description("Ruta completa del archivo que se creará")]
        [DisplayName("Ruta del archivo")]
        [RequiredArgument]
        public InArgument<string> PathFile { get; set; }

        [Category("In")]
        [Description("Texto a escribir")]
        [RequiredArgument]
        public InArgument<string> Texto { get; set; }

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
                String vTexto = Texto.Get(context);

                string CadenaArchivo = "";

                if (File.Exists(vPathFile))
                {
                    //Leer TXT
                    using (StreamReader sr = File.OpenText(vPathFile))
                    {
                        string s = "";
                        while ((s = sr.ReadLine()) != null)
                        {
                            CadenaArchivo = CadenaArchivo + s;
                        }

                    }

                    CadenaArchivo = CadenaArchivo + Environment.NewLine + vTexto;

                    // escribir todo
                    using (StreamWriter sw = File.CreateText(vPathFile))
                    {
                        sw.WriteLine(CadenaArchivo);
                    }
                }
                else
                {
                    Console.WriteLine("El archivo No existe");
                    Result.Set(context, false);
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
    /// Descripción: Leer un Archivo Txt
    /// Autor: Carlos Cardona
    /// Fecha Ultima Versión: 2020-06-19
    /// </summary>
    [Description("Leer un archivo tipo TXT")]
    public class ReadTxt : CodeActivity
    {

        [Category("In")]
        [Description("Ruta completa del archivo")]
        [DisplayName("Ruta del archivo")]
        [RequiredArgument]
        public InArgument<string> PathFile { get; set; }

        [Category("Out")]
        [Description("Texto del archivo")]
        [RequiredArgument]
        public OutArgument<string> Texto { get; set; }

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
                String CadenaArchivo = "";

                if (File.Exists(vPathFile))
                {
                    using (StreamReader sr = File.OpenText(vPathFile))
                    {
                        string s = "";
                        while ((s = sr.ReadLine()) != null)
                        {
                            CadenaArchivo = CadenaArchivo + s;

                        }

                    }
                }
                else
                {
                    Texto.Set(context, null);
                    Result.Set(context, false);
                }

                Texto.Set(context, CadenaArchivo);
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