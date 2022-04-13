using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

using System.Activities;
using System.ComponentModel;

using System.IO;
using System.Text;
using System.Data;


namespace CSV
{
    /// <summary>
    /// En construcción
    /// </summary>
    [Description("Insertar datos adicionales a un archivo tipo CSV")]
    public class AppendToCSV : CodeActivity
    {


        protected override void Execute(CodeActivityContext context)
        {
            throw new NotImplementedException();
        }
    }

    /// <summary>
    /// Descripción: Leer CSV
    /// Autor: Carlos Cardona
    /// Fecha Ultima Versión: 2020-06-16
    /// </summary>
    [Description("Leer un archivo tipo CSV")]
    public class ReadCSV : CodeActivity
    {
        [Category("Input")]
        [DisplayName("Ruta del archivo")]
        [Description("Ruta completa del archivo que se quiere leer.")]
        [RequiredArgument]
        public InArgument<String> PathFile { get; set; }

        [Category("Options")]
        [DisplayName("Delimitador")]
        [Description("Delimitador utilizado para separar cada columna, en csv normalmente es ';' ")]
        [RequiredArgument]
        public InArgument<String> Delimiter { get; set; }

        [Category("Options")]
        [DisplayName("Codificación Archivo (Encoding)")]
        [RequiredArgument]
        [Description("Escriba el número o el nombre de codificación del archivo CSV. Para más información https://docs.microsoft.com/en-us/dotnet/api/system.text.encoding?view=netframework-4.8")]
        public InArgument<String> Encoding { get; set; }
        //UTF-8 => 65001

        [Category("Options")]
        [DisplayName("Incluir nombres columnas")]
        [RequiredArgument]
        [Description("Si desea incluir los nombres de las columnas del CSV.")]
        public InArgument<Boolean> IncludeColumnNames { get; set; }

        [Category("Output")]
        [RequiredArgument]
        [Description("Data table donde se almacenara el resultado de la lectura.")]
        public OutArgument<DataTable> DataTable { get; set; }

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
                String vStrPathFile = PathFile.Get(context);
                Char[] vArrChrDelimiter = Delimiter.Get(context).ToCharArray();
                String vStrEncoding = Encoding.Get(context);
                Boolean vBooIncludeColumnNames = IncludeColumnNames.Get(context);

                // Creamos la variable de salida
                DataTable vDtbResult = new System.Data.DataTable();

                // Declarar variable de codificación
                Encoding vEncEncoding;

                // Encontrar la codificación del archivo

                if (String.IsNullOrWhiteSpace(vStrEncoding))
                {
                    vEncEncoding = FindEncoding(vStrPathFile);
                }
                else
                {
                    try
                    {
                        vEncEncoding = System.Text.Encoding.GetEncoding(vStrEncoding.ToLower());
                    }
                    catch (Exception)
                    {
                        vEncEncoding = System.Text.Encoding.Default;
                    }
                }

                using (StreamReader sr = new StreamReader(vStrPathFile, vEncEncoding))
                {
                    // Leer línea y separar por caracter delimitador
                    String[] headers = sr.ReadLine().Split(vArrChrDelimiter);

                    // Verificar si se incluyen nombres de columnas
                    if (vBooIncludeColumnNames)
                    {
                        foreach (var StrHeader in headers)
                        {
                            vDtbResult.Columns.Add(StrHeader);
                        }
                    }
                    else
                    {
                        // En caso de no tener encabezados se agregarán columnas y luego se agrega esta primera línea como datos
                        for (int i = 0; i < headers.Count(); i++)
                        {
                            vDtbResult.Columns.Add(new DataColumn());
                        }

                        DataRow dr = vDtbResult.NewRow();
                        for (int i = 0; i < headers.Length; i++)
                        {
                            dr[i] = headers[i];
                        }
                        vDtbResult.Rows.Add(dr);
                        dr = null;
                    }

                    // Leer el resto de filas
                    while (!sr.EndOfStream)
                    {
                        // Leer siguiente línea y separar por caracter delimitador
                        String[] rows = sr.ReadLine().Split(vArrChrDelimiter);
                        // Crear nueva fila de datos
                        DataRow dr = vDtbResult.NewRow();
                        for (int i = 0; i < headers.Length; i++)
                        {
                            dr[i] = rows[i];
                        }
                        // Agregar fila de datos a datatable y eliminar la fila creada
                        vDtbResult.Rows.Add(dr);
                        dr = null;
                    }
                }

                // Asignar salida
                DataTable.Set(context, vDtbResult);
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

        public static Encoding FindEncoding(String fileName)
        {
            // Pendiente crear función
            if (!File.Exists(fileName))
            return System.Text.Encoding.Default;

            byte[] byteOrderMark = new byte[4];
            using (FileStream file = new FileStream(fileName, FileMode.Open))
            {
                try 
	            {	        
		            file.Read(byteOrderMark, 0, 4);
	            }
	            catch (Exception)
	            {
                    Console.WriteLine("Hubo un error intentanto obtener la codificación del archivo " + fileName);
                    return System.Text.Encoding.Default;
	            } finally   
                {
                    file.Close();
                }
            }

            System.Text.Encoding encoding = (System.Text.Encoding) null;
            if (byteOrderMark == null || byteOrderMark.Length == 0)
            return System.Text.Encoding.Default;
            switch (byteOrderMark[0])
            {
            case 239:
                if (byteOrderMark.Length >= 3 && byteOrderMark[1] == (byte) 187 && byteOrderMark[2] == (byte) 191)
                {
                encoding = System.Text.Encoding.UTF8;
                break;
                }
                break;
            case 254:
                if (byteOrderMark[1] == byte.MaxValue)
                {
                encoding = System.Text.Encoding.BigEndianUnicode;
                break;
                }
                break;
            case byte.MaxValue:
                if (byteOrderMark[1] == (byte) 254)
                {
                encoding = System.Text.Encoding.Unicode;
                break;
                }
                break;
            default:
                encoding = System.Text.Encoding.UTF8;
                break;
            }
            return encoding;
        }
    }

    /// <summary>
    /// Descripción: Escribir CSV
    /// Autor: Carlos Cardona
    /// Fecha Ultima Versión: 2020-06-16
    /// </summary>
    [Description("Crear un archivo tipo CSV")]
    public class WriteCSV : CodeActivity
    {
        public WriteCSV()
        {
            Encoding = "65001";
            AddHeaders = true;
        }

        [Category("File")]
        [DisplayName("Ruta del archivo")]
        [Description("Ruta completa donde se creara el CSV.")]
        [RequiredArgument]
        public InArgument<String> PathFile { get; set; }

        [Category("Input")]
        [DisplayName("Data table base")]
        [Description("Data table base que se escribira en el CSV")]
        [RequiredArgument]
        public InArgument<System.Data.DataTable> DataTable { get; set; }

        [Category("Input")]
        [DisplayName("Delimitador")]
        [Description("Delimitador utilizado para separar cada columna, en csv normalmente es ';' ")]
        [RequiredArgument]
        public InArgument<String> Delimiter { get; set; }

        [Category("Options")]
        [DisplayName("AddHeaders")]
        [Description("Escribir los headers del Data Base?")]
        [RequiredArgument]
        public InArgument<Boolean> AddHeaders { get; set; }

        [Category("Options")]
        [DisplayName("Codificación Archivo (Encoding)")]
        [Description("Escriba el número o el nombre de codificación del archivo CSV. Para más información https://docs.microsoft.com/en-us/dotnet/api/system.text.encoding?view=netframework-4.8")]
        [RequiredArgument]
        public InArgument<String> Encoding { get; set; }
        //UTF-8 => 65001

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
                String vStrPathFile = PathFile.Get(context);
                DataTable vDtbDataTable = DataTable.Get(context);
                Char[] vArrChrDelimiter = Delimiter.Get(context).ToCharArray();
                Boolean vBooAddHeaders = AddHeaders.Get(context);
                String vStrEncoding = Encoding.Get(context);

                // Declarar variable de codificación
                Encoding vEncEncoding;

                // Verificar si se especificó codificado del archivo
                if (String.IsNullOrWhiteSpace(vStrEncoding))
                {
                    vEncEncoding = System.Text.Encoding.Default;
                }
                else
                {
                    try
                    {
                        vEncEncoding = System.Text.Encoding.GetEncoding(vStrEncoding.ToLower());
                    }
                    catch (Exception)
                    {
                        vEncEncoding = System.Text.Encoding.Default;
                    }
                }
                // Verificar si existe el directorio
                if (!Directory.Exists(Path.GetDirectoryName(vStrPathFile)))
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(vStrPathFile));
                }
                // Iniciar StreamWriter
                StreamWriter sw = new StreamWriter(vStrPathFile, false, vEncEncoding);

                // Verificar si se agregarán nombres de columnas
                if (vBooAddHeaders)
                {
                    for (int i = 0; i < vDtbDataTable.Columns.Count; i++)
                    {
                        sw.Write(vDtbDataTable.Columns[i].ToString().Trim());
                        if (i < vDtbDataTable.Columns.Count - 1)
                        {
                            sw.Write(vArrChrDelimiter);
                        }
                    }
                    sw.Write(sw.NewLine);
                }

                // Escribir las filas
                foreach (DataRow dr in vDtbDataTable.Rows)
                {
                    for (int i = 0; i < vDtbDataTable.Columns.Count; i++)
                    {
                        String vStrValue = dr[i].ToString().Trim();
                        if (vStrValue.Contains(vArrChrDelimiter.ToString()))
                        {
                            vStrValue = String.Format("\"{0}\"", vStrValue);
                            sw.Write(vStrValue);
                        }
                        else
                        {
                            sw.Write(dr[i].ToString().Trim());
                        }

                        if (i < vDtbDataTable.Columns.Count - 1)
                        {
                            sw.Write(vArrChrDelimiter);
                        }
                    }
                    sw.Write(sw.NewLine);
                }

                sw.Close();
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
