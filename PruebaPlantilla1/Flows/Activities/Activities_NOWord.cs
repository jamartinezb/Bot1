using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Activities;
using System.ComponentModel;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System.IO;
using System.IO.Packaging;
using System.Xml.Linq;

// Open-Xml-PowerTools Library
// More information: https://github.com/OfficeDev/Open-Xml-PowerTools
// License: https://github.com/OfficeDev/Open-Xml-PowerTools/blob/vNext/LICENSE
// Support: 
using OpenXmlPowerTools;

namespace MicrosoftWord
{
    /// <summary>
    /// Descripción: Convertir archivo word a pdf
    /// Autor: ???
    /// Fecha Ultima Versión: 2020-07-22
    /// </summary>
    [Description("Covnertir archivo Microsoft word a un archivo PDF")]
    public class SaveToPDF : CodeActivity
    {
        [Category("Input")]
        [DisplayName("Ruta Completa Archivo Word")]
        [Description("Ruta completa del archivo que se quiere convertir.")]
        [RequiredArgument]
        public InArgument<String> docxPath { get; set; }

        [Category("Input")]
        [DisplayName("Ruta Completa Archivo PDF")]
        [Description("Ruta completa del archivo que se quiere convertir.")]
        [RequiredArgument]
        public InArgument<String> pdfPath { get; set; }

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
                String vStrdocxPath = docxPath.Get(context);
                String vStrpdfPath = pdfPath.Get(context);

                var source = Package.Open(vStrdocxPath);
                var document = WordprocessingDocument.Open(source);
                HtmlConverterSettings settings = new HtmlConverterSettings();
                XElement html = HtmlConverter.ConvertToHtml(document, settings);

                //Console.WriteLine(html.ToString());
                var writer = File.CreateText(vStrpdfPath);
                writer.WriteLine(html.ToString());
                writer.Dispose();
                //Console.ReadLine();
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
      };
           
}
