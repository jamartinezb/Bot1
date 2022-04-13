using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Activities;
using System.ComponentModel;

using System.Data;

namespace DataTableActivities
{

    /// <summary>
    /// Descripción: Agregar Columna a un DataTable
    /// Autor: Carlos Cardona
    /// Fecha Ultima Versión: 2020-06-16
    /// </summary>
    [Description("Agregar una columna a un Data Table")]
    public class AddDataColumn : CodeActivity
    {
        public AddDataColumn()
        {
            DefaultValue = "";
            MaxLenght = 100;
            Unique = false;
        }

        public enum DType
        {
            String,
            Boolean,
            Int32,
            Object
        }

        [Category("Input")]
        [RequiredArgument]
        [Description("DataTable donde se insertara la columna.")]
        public InOutArgument<DataTable> DataTable { get; set; }

        [Category("Input Option 2")]
        [DisplayName("DataColumn")]
        [Description("Variable tipo DataColumn que se insertara a la tabla. Con esta variable de entrada solo se requiere la variable DataTable.")]
        public InArgument<DataColumn> Column { get; set; }

        [Category("Input Option 1")]
        [DisplayName("Nombre Columna")]
        [Description("Nombre de la nueva columna a insertar en la tabla.")]
        public InArgument<String> ColumnName { get; set; }

        [Category("Input Option 1")]
        [DisplayName("Tipo de datos en la Columna")]
        [Description("Tipo de dato que manejará la nueva columna.")]
        [DefaultValue(true)]
        public DType TypeOfData { get; set; }

        [Category("Input Option 1")]
        [DisplayName("Valor Por Defecto")]
        [Description("valor por defecto de la nueva columna.")]
        [DefaultValue(true)]
        public InArgument<String> DefaultValue { get; set; }

        [Category("Input Option 1")]
        [DisplayName("Tamaño maximo")]
        [Description("Tamaño maximo.")]
        [DefaultValue(true)]
        public InArgument<Int32> MaxLenght { get; set; }

        [Category("Input Option 1")]
        [Description("Define el parametro Unique de la columna nueva.")]
        [DefaultValue(true)]
        public InArgument<Boolean> Unique { get; set; }

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
                DataTable vDtb = DataTable.Get(context);
                String vColumnName = ColumnName.Get(context);
                String vDefaultValue = DefaultValue.Get(context);

                if (Column.Get(context) != null)
                {
                    // Si se agrega columna de datos se agrega directamente
                    vDtb.Columns.Add(Column.Get(context));
                }
                else
                {
                        //Validar existencia de columna
                        if (!vDtb.Columns.Contains(vColumnName))
                        {
                            // Se no se especifica columna de datos se crea una con los parámetros ingresados
                            DataColumn col = new DataColumn("Int32Col");
                            col.DataType = System.Type.GetType("System." + TypeOfData.ToString());
                            col.DefaultValue = vDefaultValue;
                            col.Unique = Unique.Get(context);

                            if (TypeOfData.ToString() != "Boolean")
                                col.MaxLength = MaxLenght.Get(context);

                            col.ColumnName = vColumnName;
                            vDtb.Columns.Add(col);
                        }
                }

                // Establecer salida
                DataTable.Set(context, vDtb);
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
    //Cuenta el numero de veces que hay un registro en una tabla y lo tipifica en una nueva
    //Ultima version 16/04/2020
    //Autora:July Rodriguez.
    /// </summary>
    [Description("Cuenta el numero de veces que hay un registro en una tabla y lo tipifica en una nueva")]
    public class AddDataColumnWithNumberOfRecordsbyRegister : CodeActivity
    {
        [Category("Input")]
        public InOutArgument<DataTable> DataTable { get; set; }

        [Category("Input")]
        [Description("Nombre de la nueva columna donde se tipifica el numero de registros de 'Columna' ")]
        public InArgument<String> ColumnaCalcular { get; set; }

        [Category("Input")]
        [Description("Nombre de la columna de la que se va hacer conteo")]
        public InArgument<Int32> Columna { get; set; }

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
                System.Data.DataTable vDtb = DataTable.Get(context);
                //List<OutParameterTransaccion> parameters = new List<OutParameterTransaccion>();

                var query = from row in vDtb.AsEnumerable()
                            group row by row[Columna.Get(context)] into solicitud
                            select new
                            {
                                Name = solicitud.Key,
                                CountOfClients = solicitud.Count()
                            };

                foreach (var solicitudes in query)
                {
                    vDtb.AsEnumerable().Where(row => row[Columna.Get(context)].ToString().Contains(solicitudes.Name.ToString())).ToList().ForEach(row => row.SetField(ColumnaCalcular.Get(context), solicitudes.CountOfClients));
                }

                DataTable.Set(context, vDtb);
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
    /// Descripción: Agregar fila nueva a un datatable
    /// Autor: Carlos Cardona
    /// Fecha Ultima Versión: 2020-06-16
    /// </summary>
    [Description("Agregar un data row a un Data Table")]
    public class AddDataRow : CodeActivity
    {
        [Category("Input")]
        [DisplayName("DataTable")]
        [Description("Tabla de datos donde se va a agregar el datarow")]
        public InOutArgument<DataTable> Destino { get; set; }

        [Category("Input")]
        [DisplayName("Datarow")]
        [Description("Datarow a agregar")]
        public InArgument<DataRow> InDataRow { get; set; }


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
                //Capturar Variables
                DataTable vDtbActividad = Destino.Get(context);

                vDtbActividad.Rows.Add(InDataRow.Get(context));
                // Escribir salidas
                Destino.Set(context, vDtbActividad);
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
    /// Descripción: Construir DataTable * En construcción
    /// Autor: N/A
    /// Fecha Ultima Versión: 2020-06-16
    /// </summary>
    [Description("Construir un Data Table. En construcción*")]
    public class BuildDataTable : CodeActivity
    {
        [Category("Output")]
        public OutArgument<DataTable> DataTable { get; set; }

        [Category("Input")]
        public InArgument<String> ColumnNames { get; set; }

        [Category("Input")]
        public InArgument<String> ColumnTypes { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            // Obtener argumentos
            System.Data.DataTable vDtb = DataTable.Get(context);

            DataTable.Set(context, vDtb);
        }
    }

    [Description("Limpiar un Data Table")]
    public class ClearTable : CodeActivity
    {
        [Category("Input")]
        [DisplayName("DataTable")]
        [Description("Table to modify")]
        public InOutArgument<DataTable> DTable { get; set; }

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
                //Capturar Variables
                DataTable vDtbActividad = DTable.Get(context);

                vDtbActividad.Rows.Clear();
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
    /// Descripción: Eliminar Columna de un DataTable basado en el nombre de la columna
    /// Autor: Carlos Cardona
    /// Fecha Ultima Versión: 2020-06-16
    /// </summary>
    [Description("Eliminar columna de un Data Table por nombre")]
    public class DeleteColumnByName : CodeActivity
    {

        [Category("Input")]
        [Description("Nombre de la Columna a eliminar")]
        [RequiredArgument]
        public InArgument<String> ColumnName { get; set; }

        [Category("Input")]
        [Description("DataTable a la que se le realizará la actividad")]
        [RequiredArgument]
        public InArgument<DataTable> DataTable { get; set; }

        [Category("Out")]
        [Description("DataTable de salida.")]
        [RequiredArgument]
        public OutArgument<DataTable> DataTableOut { get; set; }


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
                DataTable vDtb = DataTable.Get(context);
                string vColumnName = ColumnName.Get(context);

                vDtb.Columns.Remove(vColumnName);

                // Establecer salida
                DataTableOut.Set(context, vDtb);
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
    /// Descripción: Eliminar filas de una tabla basado en el valor de una columna
    /// Autor: Carlos Cardona
    /// Fecha Ultima Versión: 2020-06-16
    /// </summary>
    [Description("Eliminar filas de un Data Table basado en el valor en una columna")]
    public class DeleteRowsByDataColumn : CodeActivity
    {
        public DeleteRowsByDataColumn()
        {
            ColumnName = "";
            ColumnValue = "";
        }


        [Category("Input")]
        [Description("Nombre de la Columna a revisar")]
        public InArgument<String> ColumnName { get; set; }

        [Category("Input")]
        [Description("Las filas cuya columna tenga este valor seran eliminadas")]
        public InArgument<String> ColumnValue { get; set; }

        [Category("Input")]
        [Description("Data Table de entrada a la actividad")]
        [RequiredArgument]
        public InArgument<DataTable> DataTable { get; set; }

        [Category("Output")]
        [Description("Data Table con el resultado de la actividad.")]
        [RequiredArgument]
        public OutArgument<DataTable> DataTableOut { get; set; }


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
                DataTable vDtb = DataTable.Get(context);
                string vColumnName = ColumnName.Get(context);
                string vColumnValue = ColumnValue.Get(context);

                string filtro = vColumnName + " = '" + vColumnValue + "'";


                if (vDtb.Rows.Count > 0)
                {
                    DataRow[] result = vDtb.Select(filtro);

                    //Eliminnar 
                    foreach (DataRow row in result)
                    {
                        if (row[vColumnName].ToString().Trim().ToUpper().Contains(vColumnValue.ToUpper().Trim()))
                            vDtb.Rows.Remove(row);
                    }

                    // Establecer salida
                    DataTableOut.Set(context, vDtb);
                }

                Result.Set(context, true);

            }
            catch (Exception e)
            {
                string error = "Mensaje Error: " + e.Message + Environment.NewLine + "InnerException: " + e.InnerException;
                What.Set(context, error);
                Result.Set(context, false);
            }
        }
    }

    /// <summary>
    /// Descripción: Unificar DataTables
    /// Autor: Carlos Cardona
    /// Fecha Ultima Versión: 2020-06-16
    /// </summary>
    [Description("Unificar los registros de dos DataTable")]
    public class MergeDataTables : CodeActivity
    {
        [Category("Input")]
        [Description("DataTable 1, tabla a la que se le insertaran los datos de la tabla 2.")]
        public InOutArgument<DataTable> Tabla1 { get; set; }

        [Category("Input")]
        [Description("DataTable 2, tabla de la cual se copiaran los registros a insertar en la tabla 1.")]
        public InArgument<DataTable> Tabla2 { get; set; }


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
                //Capturar variables
                DataTable vDtbTablaDestino = Tabla1.Get(context);
                DataTable vDtbTablaOrigen = Tabla2.Get(context);

                vDtbTablaDestino.Merge(vDtbTablaOrigen);
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

    [Description("Imprimir un DataTable en la consola")]
    public class PrintConsole : CodeActivity
    {


        [Category("In")]
        [Description("Navegador a utilizar")]
        [RequiredArgument]
        public InArgument<DataTable> DtTable { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            try
            {
                DataTable vDtable = DtTable.Get(context);

                foreach (DataRow dataRow in vDtable.Rows)
                {
                    foreach (var item in dataRow.ItemArray)
                    {
                        Console.Write(item + ";");
                    }
                    Console.WriteLine("");
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Error al imprimir DataTable: " + e.Message + ". Detalles: " + e.InnerException);

            }

        }
    }
    
    [Description("Imprimir un DataTable en la consola")]
    public class PrintTableCMD : CodeActivity
    {
        [Category("Input")]
        [DisplayName("DataTable")]
        [Description("Table to print")]
        public InOutArgument<DataTable> Table { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            DataTable vDtbActividad = Table.Get(context);

            try
            {
                DataRow[] currentRows = vDtbActividad.Select(
                null, null, DataViewRowState.CurrentRows);

                if (currentRows.Length < 1)
                    Console.WriteLine("No Current Rows Found");
                else
                {
                    foreach (DataColumn column in vDtbActividad.Columns)
                    {
                        Console.Write("\t{0}", column.ColumnName);
                    }

                    Console.WriteLine("\tRowState");

                    foreach (DataRow row in currentRows)
                    {
                        foreach (DataColumn column in vDtbActividad.Columns)
                        {
                            Console.Write("\t{0}", row[column]);
                        }

                        Console.WriteLine("\t" + row.RowState);
                    }
                }

            }
            catch (Exception)
            {

                throw;
            }
        }
    }

    [Description("Reemplazar un texto por otro en todos los valores del DataTable")]
    public class ReplaceInTable : CodeActivity
    {
        [Category("Input")]
        [DisplayName("DataTable")]
        [Description("Table to modify")]
        public InOutArgument<DataTable> Table { get; set; }

        [Category("Input")]
        [DisplayName("Value To Replace")]
        [Description("Value To Replace")]
        public InArgument<String> InValueToReplace { get; set; }

        [Category("Input")]
        [DisplayName("New Value")]
        [Description("New Value")]
        public InArgument<String> InValueNew { get; set; }


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
                //Capturar Variables
                DataTable vDtbActividad = Table.Get(context);
                String ValueToReplace = InValueToReplace.Get(context);
                String ValueNew = InValueNew.Get(context);

                DataRow[] currentRows = vDtbActividad.Select(
                null, null, DataViewRowState.CurrentRows);

                if (currentRows.Length < 1)
                    Console.WriteLine("No Current Rows Found");
                else
                {

                    foreach (DataRow row in currentRows)
                    {
                        foreach (DataColumn column in vDtbActividad.Columns)
                        {
                            row[column] = row[column].ToString().Replace(ValueToReplace, ValueNew);
                            //Console.Write("\t{0}", row[column]);
                        }

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

    [Description("Definir el valor de una celda de un Data Table")]
    public class SetCell : CodeActivity
    {
        [Category("Input")]
        [DisplayName("DataTable")]
        [Description("Tabla de datos donde se va a agregar el datarow")]
        public InOutArgument<DataTable> Destino { get; set; }

        [Category("Input")]
        [DisplayName("Row")]
        [Description("row a editar")]
        public InArgument<Int32> InRow { get; set; }

        [Category("Input")]
        [DisplayName("Colum Name")]
        [Description("Nombre de la columna")]
        public InArgument<String> ColumnName { get; set; }

        [Category("Input")]
        [DisplayName("Value")]
        [Description("Valor")]
        public InArgument<String> InValue { get; set; }


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
                //Capturar variables
                DataTable vDtbActividad = Destino.Get(context);
                Int32 vRow = InRow.Get(context);
                String vColumName = ColumnName.Get(context);
                String vValue = InValue.Get(context);

                vDtbActividad.Rows[vRow][vColumName] = vValue;

                // Escribir salidas
                Destino.Set(context, vDtbActividad);
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

    [Description("Definir el valor de toda una columna de un Data Table")]
    public class SetColumn : CodeActivity
    {
        [Category("Input")]
        [DisplayName("DataTable")]
        [Description("Tabla de datos donde se va a agregar el datarow")]
        public InOutArgument<DataTable> Destino { get; set; }

        [Category("Input")]
        [DisplayName("Colum Name")]
        [Description("Nombre de la columna")]
        public InArgument<String> ColumnName { get; set; }

        [Category("Input")]
        [DisplayName("Value")]
        [Description("Valor")]
        public InArgument<String> InValue { get; set; }


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
            DataTable vDtbActividad = Destino.Get(context);
            String vColumName = ColumnName.Get(context);
            String vValue = InValue.Get(context);

            try
            {
                vDtbActividad.Columns[vColumName].Expression = vValue;

                // Escribir salidas
                Destino.Set(context, vDtbActividad);
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

    [Description("Limpiar espacios en todos los valores de un Data Table")]
    public class TrimInTable : CodeActivity
    {
        [Category("Input")]
        [DisplayName("DataTable")]
        [Description("Table to modify")]
        public InOutArgument<DataTable> Table { get; set; }


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
                //Capturar Variables
                DataTable vDtbActividad = Table.Get(context);

                //Capturar Rows
                DataRow[] currentRows = vDtbActividad.Select(
                null, null, DataViewRowState.CurrentRows);

                if (currentRows.Length < 1)
                    Console.WriteLine("No Current Rows Found");
                else
                {

                    foreach (DataRow row in currentRows)
                    {
                        foreach (DataColumn column in vDtbActividad.Columns)
                        {
                            row[column] = row[column].ToString().Trim();
                            //Console.Write("\t{0}", row[column]);
                        }

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
    /// Descripción: BuscarColumna en el data table
    /// Autor: N/A
    /// Fecha Ultima Versión: 2020-06-16
    /// </summary>
    [Description("Realiza el cruce de dos Data Table para Actualizar campos en una de ella")]
    public class UpdateFieldByCrossingDT : CodeActivity
    {
        [Category("Input")]
        [Description("Datatable")]
        public InOutArgument<DataTable> DataTable { get; set; }

        [Category("Input")]
        [Description("Nombre de la columna a actualizar")]
        public InArgument<String> ColumnaCalcular { get; set; }

        [Category("Input")]
        [Description("Nombre de la columna con valor a buscar")]
        public InArgument<String> ColumnaValor { get; set; }

        [Category("Filtro")]
        [Description("Nombre de la columna que se quiere filtrar para aplicar la búsqueda")]
        public InArgument<String> ColumnaFiltro { get; set; }

        [Category("Filtro")]
        [Description("Filtro que se desea aplicar, se utiliza el método Contains")]
        public InArgument<String> ValorFiltro { get; set; }

        [Category("Origen Datos")]
        [Description("DataTable origen")]
        public InArgument<DataTable> DataTableOrigen { get; set; }

        [Category("Origen Datos")]
        [Description("Columna a buscar")]
        public InArgument<String> ColumnaBusqueda { get; set; }

        [Category("Origen Datos")]
        [Description("Columna de retorno")]
        public InArgument<String> ColumnaRetorno { get; set; }

        [Category("Misc")]
        [Description("Valor No Encontrado")]
        public InArgument<String> ValorNoEncontrado { get; set; }


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
                System.Data.DataTable vDtb = DataTable.Get(context);

                // Calcular columna
                if (!String.IsNullOrWhiteSpace(ColumnaFiltro.Get(context)) && !String.IsNullOrWhiteSpace(ValorFiltro.Get(context)))
                {
                    vDtb.AsEnumerable().Where(row => row[ColumnaFiltro.Get(context)].ToString().Contains(ValorFiltro.Get(context))).ToList().ForEach(row => row.SetField(ColumnaCalcular.Get(context), Calcular(DataTableOrigen.Get(context), ColumnaRetorno.Get(context), ColumnaBusqueda.Get(context), row.Field<String>(ColumnaValor.Get(context)), ValorNoEncontrado.Get(context))));
                }
                else
                {
                    vDtb.AsEnumerable().ToList().ForEach(row => row.SetField(ColumnaCalcular.Get(context), Calcular(DataTableOrigen.Get(context), ColumnaRetorno.Get(context), ColumnaBusqueda.Get(context), row.Field<String>(ColumnaValor.Get(context)), ValorNoEncontrado.Get(context))));
                }

                DataTable.Set(context, vDtb);
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

        public String Calcular(DataTable vDtbDataTableOrigen, String vStrNombreColumnaRetorno, String vStrNombreColumnaBusqueda, String ValorBusqueda, String ValorNoEncontrado)
        {
            try
            {
                return vDtbDataTableOrigen.AsEnumerable().Where(row => row.Field<String>(vStrNombreColumnaBusqueda).ToString() == ValorBusqueda).FirstOrDefault()[vStrNombreColumnaRetorno].ToString();
            }
            catch (Exception)
            {
                return ValorNoEncontrado;
            }
        }
    }

}
