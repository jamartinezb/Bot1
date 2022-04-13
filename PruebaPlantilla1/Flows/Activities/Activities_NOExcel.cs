using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Activities;
using System.ComponentModel;

using System.Data;

using System.IO;

// Librería EPPlus
// Más información: https://github.com/JanKallman/EPPlus
// No recomendables al no soportar documentos Excel 1997-2003 (.xls)
//using OfficeOpenXml;
//using OfficeOpenXml.Packaging;

// Librería de ExcelLibrary
// Más información: https://code.google.com/archive/p/excellibrary
// Reemplazado por NPOI
// Usada para la lectura de archivos de Excel 1997-2003
//using ExcelLibrary;

// Librería de NPOI
// Más información: https://archive.codeplex.com/?p=npoi
// Support xls, xlsx, docx.
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.SS.Util;

namespace Excel_NoOffice
{
    [Description("Insertar en Excel, un datatable. No requiere Excel instalado")]
    public class AppendRange : CodeActivity
    {
        public AppendRange()
        {
            SheetName = "Sheet1";
        }

        [Category("Input")]
        [Description("DataTable a insertar")]
        [RequiredArgument]
        public InArgument<System.Data.DataTable> DataTable { get; set; }

        [Category("Input")]
        [Description("Nombre de la hoja en la que se insertara")]
        [RequiredArgument]
        [DefaultValue(true)]
        public InArgument<String> SheetName { get; set; }

        [Category("Input")]
        [Description("Ruta completa del documento Excel.")]
        [RequiredArgument]
        public InArgument<String> WorkbookPath { get; set; }

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
                String vStrSheetName = SheetName.Get(context);
                String vStrWorkbookPath = WorkbookPath.Get(context);
                System.Data.DataTable vDtbDataTable = DataTable.Get(context);

                // Declarar variables
                IWorkbook workbook;

                using (FileStream stream = new FileStream(vStrWorkbookPath, FileMode.Open, FileAccess.ReadWrite))
                {
                    switch (Path.GetExtension(vStrWorkbookPath).ToLower())
                    {
                        case ".xls": workbook = new HSSFWorkbook(stream); break;
                        case ".xlsx": workbook = new XSSFWorkbook(stream); break;
                        default: throw new Exception("This format is not supported");
                    }
                    stream.Close();
                }

                // Ubicar hoja condatos
                ISheet sheet = workbook.GetSheet(vStrSheetName);
                if (sheet == null)
                {

                    DataTable.Set(context, vDtbDataTable);
                    return;
                }
                // Obtener última fila con datos
                int rownum = sheet.LastRowNum;

                // Add rows
                foreach (DataRow dr in vDtbDataTable.Rows)
                {
                    // this creates a new row in the sheet
                    IRow row = sheet.CreateRow(rownum++);
                    int Cellnum = 0;

                    // Iterate for each column of the dr
                    foreach (Object item in dr.ItemArray)
                    {
                        ICell Cell = row.CreateCell(Cellnum++);
                        // this line creates a cell in the next column of that row 
                        Cell.SetCellValue((String)item.ToString());
                    }
                }

                using (FileStream stream = new FileStream(vStrWorkbookPath, FileMode.Open, FileAccess.ReadWrite))
                {
                    // Ending stream
                    workbook.Write(stream);
                    stream.Close();
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
    /// Descripción: Leer Rango
    /// Autor: Carlos Cardona
    /// Fecha Ultima Versión: 2020-06-26
    /// </summary>
    [Description("Leer rango de un Excel y almacenar en un data table. No requiere Excel instalado")]
    public class ReadRange : CodeActivity
    {
        //Parametros iniciales
        public ReadRange()
        {
            Range = "A1:A2";
            SheetName = "Sheet1";
            Headers = true;
            ExceptionHandling = false;
        }

        [Category("Input")]
        [DisplayName("Rango a Capturar")]
        [Description("Rango a capturar Ej: A#:C#")]
        [DefaultValue(true)]
        public InArgument<String> Range { get; set; }

        [Category("Input")]
        [DisplayName("Nombre de la hoja")]
        [Description("Nombre de la hoja, si no se coloca tomara la primer hoja.")]
        [DefaultValue(true)]
        public InArgument<String> SheetName { get; set; }

        [Category("Input")]
        [DisplayName("Ruta Completa Archivo Excel")]
        [Description("Ruta completa del archivo Excel")]
        [RequiredArgument]
        public InArgument<String> WorkbookPath { get; set; }

        [Category("Options")]
        [DisplayName("Tiene encabezados?")]
        [Description("Tiene encabezados?")]
        [RequiredArgument]
        public Boolean Headers { get; set; }

        [Category("Output")]
        [DisplayName("Data Table Respuesta")]
        [Description("DataTable de salida")]
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
                String vStrRange = Range.Get(context);
                String vStrSheetName = SheetName.Get(context);
                String vStrWorkbookPath = WorkbookPath.Get(context);
                Boolean vBooAddHeaders = Headers;
                System.Data.DataTable vDtbDataTable = new DataTable();

                // Declarar variables
                IWorkbook workbook;
                CellRangeAddress vCellRange = null;

                // Verificar si el libro existe
                if (!File.Exists(vStrWorkbookPath))
                {
                    throw new Exception("Archivo especificado no existe.");
                }

                // Abrir el libro
                using (FileStream file = new FileStream(vStrWorkbookPath, FileMode.Open, FileAccess.Read))
                {
                    switch (Path.GetExtension(vStrWorkbookPath).ToLower())
                    {
                        case ".xls": workbook = new HSSFWorkbook(file); break; // HSSF is the POI Project's of the Excel '97(-2007) file format. 
                        case ".xlsx": workbook = new XSSFWorkbook(file); break; // XSSF is the POI Project's of the Excel 2007 OOXML (.xlsx) file format.
                        default: throw new Exception("This format is not supported");
                    }
                }

                // Ubicar hoja condatos
                ISheet sheet;
                if (vStrSheetName == null)
                {
                    sheet = workbook.GetSheet(workbook.GetSheetName(0));
                }
                else
                {
                    sheet = workbook.GetSheet(vStrSheetName);
                }

                // Verificar si la hoja existe
                if (sheet == null)
                {

                    DataTable.Set(context, vDtbDataTable);
                    return;
                }

                // Verificar si se trabaja con un rango
                if (!String.IsNullOrWhiteSpace(vStrRange))
                {
                    try
                    {
                        vCellRange = CellRangeAddress.ValueOf(vStrRange);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                        DataTable.Set(context, vDtbDataTable);
                        throw;
                    }
                }
                else
                {
                    if (sheet.GetRow(sheet.FirstRowNum).Cells.AsEnumerable().Where(c => c.CellType != CellType.Blank).Count() > 0)
                    {
                        string range = String.Format("{0}{1}:{2}{3}", CellReference.ConvertNumToColString(sheet.GetRow(sheet.FirstRowNum).FirstCellNum), sheet.FirstRowNum + 1, CellReference.ConvertNumToColString(sheet.GetRow(sheet.FirstRowNum).LastCellNum - 1), sheet.LastRowNum + 1);
                        vCellRange = CellRangeAddress.ValueOf(range);
                    }
                }

                // Validar que el rango especificado tenga datos en la primera fila
                if (sheet.GetRow(vCellRange.FirstRow) == null)
                {
                    throw new Exception("La primera fila de datos del rango especificado no contiene datos");
                }
                else if (sheet.GetRow(vCellRange.FirstRow).AsEnumerable().Where(c => c.CellType != CellType.Blank && vCellRange.IsInRange(c.RowIndex, c.ColumnIndex)).Count() == 0)
                {
                    throw new Exception("La primera fila de datos del rango especificado no contiene datos");
                }

                // Si la primera fila contiene datos de nombres de columnas
                if (vBooAddHeaders)
                {

                    foreach (var item in sheet.GetRow(vCellRange.FirstRow).Cells.Where(c => c.ColumnIndex >= sheet.GetRow(vCellRange.FirstRow).FirstCellNum && c.ColumnIndex <= sheet.GetRow(vCellRange.FirstRow).LastCellNum).ToArray())
                    {
                        // Agregar nombre de columna si la celda no es vacía
                        if (!String.IsNullOrWhiteSpace(item.ToString()))
                        {
                            vDtbDataTable.Columns.Add(item.ToString());
                        }
                    }
                }
                else
                {

                    foreach (var item in sheet.GetRow(vCellRange.FirstRow).Cells.Where(c => c.ColumnIndex >= sheet.GetRow(vCellRange.FirstRow).FirstCellNum && c.ColumnIndex <= sheet.GetRow(vCellRange.FirstRow).LastCellNum).ToArray())
                    {
                        // Agregar nombre de columna si la celda no es vacía
                        if (!String.IsNullOrWhiteSpace(item.ToString()))
                        {
                            vDtbDataTable.Columns.Add(new DataColumn());
                        }
                    }
                }

                // Leer el resto de filas
                // el ? es como un if (condicion) ? expresion1 : expresion2
                for (int row = (vBooAddHeaders ? vCellRange.FirstRow + 1 : vCellRange.FirstRow); row <= vCellRange.LastRow; row++)
                {
                    if (sheet.GetRow(row) != null) //null is when the row only contains empty cells 
                    {
                        DataRow vDrwNewRow = vDtbDataTable.NewRow();

                        //================================Cambio Manuel
                        DataTable dt = new DataTable();
                        DataRow dr = dt.NewRow();
                        int ss = sheet.GetRow(row).Cells.Count;

                        for (int i = 0; i < sheet.GetRow(row).Cells.Count; i++)
                        {
                            var tcell = sheet.GetRow(row).GetCell(i);
                            object valorCell = null;

                            if (tcell != null)
                                switch (tcell.CellType)
                                {
                                    case NPOI.SS.UserModel.CellType.Blank: valorCell = DBNull.Value; break;
                                    case NPOI.SS.UserModel.CellType.Boolean: valorCell = sheet.GetRow(row).GetCell(i).BooleanCellValue; break;
                                    case NPOI.SS.UserModel.CellType.String: valorCell = sheet.GetRow(row).GetCell(i).StringCellValue; break;
                                    case NPOI.SS.UserModel.CellType.Formula:
                                        //
                                        switch (tcell.CachedFormulaResultType)
                                        {
                                            case CellType.Blank: valorCell = DBNull.Value; break;
                                            case CellType.String: valorCell = sheet.GetRow(row).GetCell(i).StringCellValue; break;
                                            case CellType.Boolean: valorCell = sheet.GetRow(row).GetCell(i).BooleanCellValue; break;
                                            case CellType.Numeric:
                                                if (HSSFDateUtil.IsCellDateFormatted(sheet.GetRow(row).GetCell(i)))
                                                { valorCell = sheet.GetRow(row).GetCell(i).DateCellValue; }
                                                else { valorCell = sheet.GetRow(row).GetCell(i).NumericCellValue; }
                                                break;
                                        }
                                        // 
                                        break;
                                    case NPOI.SS.UserModel.CellType.Numeric:
                                        if (HSSFDateUtil.IsCellDateFormatted(sheet.GetRow(row).GetCell(i)))
                                        { valorCell = sheet.GetRow(row).GetCell(i).DateCellValue; }
                                                else { valorCell = sheet.GetRow(row).GetCell(i).NumericCellValue; 
                                        } 
                                        // 
                                        break;
                                    default: valorCell = sheet.GetRow(row).GetCell(i).StringCellValue; break;
                                }

                            vDrwNewRow[i] = valorCell;
                        }


                        //vDrwNewRow.ItemArray = sheet.GetRow(row).Cells.Select(c => c.ToString()).ToArray();
                        //vDrwNewRow.ItemArray = sheet.GetRow(row).Cells.Where(c => c.ColumnIndex >= vCellRange.FirstColumn && c.ColumnIndex <= vCellRange.LastColumn).ToArray();
                        vDtbDataTable.Rows.Add(vDrwNewRow);
                    }
                }

                // Asignar la variable de salida
                DataTable.Set(context, vDtbDataTable);
                Result.Set(context, true);

            }
            catch (Exception e)
            {
                string error = "Mensaje Error: " + e.Message + Environment.NewLine + "InnerException: " + e.InnerException;
                    What.Set(context, error);
                    Result.Set(context, false);

                Boolean vExceptionHandling = ExceptionHandling;
                if (vExceptionHandling  == false)
                {
                    throw e;
                }
                
            }
            
        }
    }

    [Description("Escribir un datable en Excel. No requiere Excel instalado")]
    public class WriteRange : CodeActivity
    {

        public WriteRange()
        {
            SheetName = "Sheet1";
            StartingCell = "A1";
            AddHeaders = true;
        }

        [Category("Destination")]
        [DisplayName("Nombre de la hoja")]
        [Description("Hoja donde se escribirá")]
        [RequiredArgument]
        [DefaultValue(true)]
        public InArgument<String> SheetName { get; set; }

        [Category("Destination")]
        [DisplayName("Celda inicial escribir")]
        [Description("Celda inicial donde se escribirá EJ: B2")]
        [RequiredArgument]
        [DefaultValue(true)]
        public InArgument<String> StartingCell { get; set; }

        [Category("Input")]
        [DisplayName("Data Table a escribir")]
        [Description("Data Table a escribir.")]
        [RequiredArgument]
        public InArgument<System.Data.DataTable> DataTable { get; set; }

        [Category("Input")]
        [DisplayName("Ruta Completa Archivo Excel")]
        [Description("Ruta donde se creara el archivo")]
        [RequiredArgument]
        public InArgument<String> WorkbookPath { get; set; }

        [Category("Options")]
        [DisplayName("Agregar encabezados?")]
        [Description("Agregar encabezados?")]
        [RequiredArgument]
        [DefaultValue(true)]
        public Boolean AddHeaders { get; set; }

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
                String vStrSheetName = SheetName.Get(context);
                String vStrWorkbookPath = WorkbookPath.Get(context);
                Boolean vBooAddHeaders = AddHeaders;
                System.Data.DataTable vDtbDataTable = DataTable.Get(context);
                String vStrStartingCell = StartingCell.Get(context);

                // Verificar si existe el directorio
                if (!Directory.Exists(Path.GetDirectoryName(vStrWorkbookPath)))
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(vStrWorkbookPath));
                }
                
                // Delete the file if it exists.
                if (File.Exists(vStrWorkbookPath))
                {
                    File.Delete(vStrWorkbookPath);
                }


                // Declarar variables
                IWorkbook workbook;
                switch (Path.GetExtension(vStrWorkbookPath).ToLower())
                {
                    case ".xls": workbook = new HSSFWorkbook(); break;
                    case ".xlsx": workbook = new XSSFWorkbook(); break;
                    default: throw new Exception("This format is not supported");
                }

                using (FileStream stream = new FileStream(vStrWorkbookPath, FileMode.Create, FileAccess.ReadWrite))
                {
                    ISheet sheet = workbook.CreateSheet(vStrSheetName);

                    CellRangeAddress vCellRange = null;

                    // Verificar si se trabaja con un rango
                    if (!String.IsNullOrWhiteSpace(vStrStartingCell))
                    {
                        try
                        {
                            vCellRange = CellRangeAddress.ValueOf(vStrStartingCell);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                            DataTable.Set(context, vDtbDataTable);
                            throw;
                        }
                    }
                    else
                    {
                        vCellRange = CellRangeAddress.ValueOf("A1");
                    }

                    int rownum = vCellRange.FirstRow; 
                    // Add headers
                    if (vBooAddHeaders)
                    {
                        IRow headerRow = sheet.CreateRow(rownum);

                        int Cellnum = vCellRange.FirstColumn;
                        foreach (DataColumn Col in vDtbDataTable.Columns)
                        {
                            ICell Cell = headerRow.CreateCell(Cellnum++);
                            // this line creates a cell in the next column of that row 
                            Cell.SetCellValue(Col.ColumnName); 
                        }
                        rownum++;
                    }

                    // Add rows
                    foreach (DataRow dr in vDtbDataTable.Rows)
                    {
                        // this creates a new row in the sheet
                        IRow row = sheet.CreateRow(rownum++);
                        int Cellnum = vCellRange.FirstColumn;
                        // Iterate for each column of the dr
                        foreach (Object item in dr.ItemArray)
                        {
                            ICell Cell = row.CreateCell(Cellnum++);
                            // this line creates a cell in the next column of that row 
                            Cell.SetCellValue((String)item.ToString()); 
                        }
                    }
                    // Ending stream
                    workbook.Write(stream);
                    stream.Close();
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

}

