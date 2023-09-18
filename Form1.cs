using SpreadsheetLight;
using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using ETL_EXCEL_SQL.Modelo;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using DocumentFormat.OpenXml.Drawing;

namespace ETL_EXCEL_SQL
{
    public partial class Form1 : Form
    {
        [DllImport("user32")]
        static extern IntPtr GetSystemMenu(IntPtr hWnd, bool bRevert);
        [DllImport("user32")]
        static extern bool EnableMenuItem(IntPtr hMenu, uint uIDEnableItem, uint uEnable);
        const int MF_BYCOMMAND = 0;
        const int MF_DISABLED = 2;
        const int SC_CLOSE = 0xF060;

        string CadenaConexion = System.Configuration.ConfigurationManager.ConnectionStrings["acae"].ConnectionString;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            btnImportar.Enabled = false;

            var sm = GetSystemMenu(Handle, false);
            EnableMenuItem(sm, SC_CLOSE, MF_BYCOMMAND | MF_DISABLED);
        }


        private void btnBuscar_Click(object sender, EventArgs e)
        {
            string ruta;
            OpenFileDialog openFileDialog = new OpenFileDialog() { Filter = "Archivos de Excel|*.xlsx|Archivos de Excel 97-2003|*.xls" };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                ruta = openFileDialog.FileName;

                try
                {
                    using (SLDocument MultiHoja = new SLDocument(ruta)) 
                    {
                        var TotalHojas = MultiHoja.GetWorksheetNames(); //FUNCION PARA LEER EL NOMBRE DE LA HOJA QUE ESTAMOS CARGANDO

                        foreach (var NombreHoja in TotalHojas) // RECORRIENDO TODAS LAS HOJAS DEL ARCHIVO
                        {
                            using (SLDocument sl = new SLDocument(ruta, NombreHoja))
                            {
                                //VALIDANDO LOS CAMPOS PARA SELECCION DE DOCUMENTO CORRECTO
                                string Codigo = sl.GetCellValueAsString(4, 2).Trim(); //CELDA CON LA PALABRA "CODIGO"
                                string CodCliente = sl.GetCellValueAsString(6, 3).Trim();
                                string Cliente = sl.GetCellValueAsString(6, 4).Trim();
                                string Precio = sl.GetCellValueAsString(6, 7).Trim();


                                //IF PARA LA VALIDACION DEL ARCHIVO
                                if (Codigo != "COD PRODUCTO" || CodCliente != "COD CLIENTE" || Cliente != "CLIENTE" || Precio != "NUEVO PRECIO")
                                {
                                    MessageBox.Show("FORMATO DE DOCUMENTO EQUIVOCADO", "ERROR DE DOCUMENTO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                    ruta = string.Empty;
                                    txtIdProducto.Text = "";
                                    txtCodigo.Text = "";
                                    txtDescripcion.Text = "";
                                    btnImportar.Enabled = false;
                                    txtRuta.Text = "";
                                }
                                else
                                {
                                    btnImportar.Enabled = true;
                                    txtRuta.Text = ruta;

                                    string codigoProducto = sl.GetCellValueAsString(4, 3);
                                    string Descripcion = "";
                                    int idProducto = 0;

                                    using (SqlConnection Conexion = new SqlConnection(CadenaConexion))
                                    {
                                        try
                                        {
                                            Conexion.Open();

                                            // Búsqueda del ID del producto (consulta parametrizada)
                                            SqlCommand cmd = new SqlCommand($"SELECT id, descripcion FROM inventario WHERE codigo=@CodigoProducto", Conexion);
                                            cmd.CommandType = CommandType.Text;
                                            cmd.Parameters.AddWithValue("@CodigoProducto", codigoProducto);
                                            SqlDataReader reader = cmd.ExecuteReader();
                                            if (reader.HasRows)
                                            {
                                                while (reader.Read())
                                                {
                                                    idProducto = reader.GetInt32(reader.GetOrdinal("Id"));
                                                    Descripcion = reader.GetString(reader.GetOrdinal("Descripcion"));
                                                }
                                                reader.Close();
                                                txtIdProducto.Text = idProducto.ToString();
                                                txtCodigo.Text = codigoProducto;
                                                txtDescripcion.Text = Descripcion.ToString();

                                                //CARGANDO LA VISUALIZACION DEL ARCHIVO EN EL GRIDVIEW
                                                /*List<InforGrid> listaClientes = new List<InforGrid>();

                                                 int iRow = 7; // Variable para recorrer el archivo

                                                 while (!string.IsNullOrEmpty(sl.GetCellValueAsString(iRow, 4)))
                                                 {
                                                     string CodigoCliente = sl.GetCellValueAsString(iRow, 3).Trim();
                                                     string NombreCliente = sl.GetCellValueAsString(iRow, 4).Trim();
                                                     string PrecioUni = sl.GetCellValueAsString(iRow, 5).Trim();
                                                     string Precio_iva = sl.GetCellValueAsString(iRow, 7).Trim();

                                                     InforGrid cliente = new InforGrid
                                                     {
                                                         Codigo = CodigoCliente,
                                                         Cliente = NombreCliente,
                                                         PrecioUnitario = "$ " + PrecioUni,
                                                         PrecioEspecial = "$ " + Precio_iva
                                                     };

                                                     listaClientes.Add(cliente);

                                                     iRow++;
                                                 }
                                                 sl.CloseWithoutSaving();

                                                 dataGridView1.DataSource = listaClientes;
                                                 dataGridView1.Columns[0].HeaderText = "CODIGO CLIENTE";
                                                 dataGridView1.Columns[1].HeaderText = "CLIENTE";
                                                 dataGridView1.Columns[2].HeaderText = "PRECIO UNITARIO";
                                                 dataGridView1.Columns[3].HeaderText = "PRECIO ESPECIAL";

                                                 dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                                                 dataGridView1.AutoResizeColumns();
                                                 dataGridView1.DefaultCellStyle.Font = new System.Drawing.Font("TAHOMA", 11);*/
                                            }
                                            else
                                            {
                                                MessageBox.Show("EL CÓDIGO DEL PRODUCTO INGRESADO NO EXISTE", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Information);

                                                reader.Close();
                                                txtIdProducto.Text = "";
                                                txtCodigo.Text = "";
                                                txtDescripcion.Text = "";
                                                btnImportar.Enabled = false;
                                                txtRuta.Text = "";
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            // CATCH DEL SELECT DE INVENTARIO
                                            Console.WriteLine("Error: " + ex.Message);
                                        }
                                        finally
                                        {
                                            // Cerrar la conexión
                                            if (Conexion.State == ConnectionState.Open)
                                            {
                                                Conexion.Close();
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }catch (Exception ex)
                {
                    MessageBox.Show("EL ARCHIVO ESTÁ SIENDO UTILIZADO, POR FAVOR CIERRELO \nPARA REALIZAR LA IMPORTACIÓN", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void btnImportar_Click(object sender, EventArgs e)
        {
            DialogResult resultado = MessageBox.Show("¿Deseas Importar la Información?", "Confirmación", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
            if (resultado == DialogResult.OK)
            {
                string ruta = txtRuta.Text;
                using (SLDocument MultiHojas = new SLDocument(ruta))
                {
                    var TotalHojas = MultiHojas.GetWorksheetNames();
                    foreach (var NombreHojas in TotalHojas)
                    {
                        Console.WriteLine(NombreHojas);
                        using (SLDocument sl = new SLDocument(ruta, NombreHojas))
                        {
                            using (SqlConnection Conexion = new SqlConnection(CadenaConexion))
                            {
                                //CARGANDO LA INFORMACION DEL EXCEL EN UNA LISTA
                                List<InforGrid> listaClientes = new List<InforGrid>();
                                int iRow = 7; // Variable para recorrer el archivo

                                while (!string.IsNullOrEmpty(sl.GetCellValueAsString(iRow, 4)))
                                {
                                    string CodigoCliente = sl.GetCellValueAsString(iRow, 3).Trim();
                                    string NombreCliente = sl.GetCellValueAsString(iRow, 4).Trim();
                                    string PrecioUni = sl.GetCellValueAsString(iRow, 5).Trim();
                                    string Precio_iva_ = sl.GetCellValueAsString(iRow, 7).Trim();

                                    InforGrid cliente = new InforGrid
                                    {
                                        Codigo = CodigoCliente,
                                        Cliente = NombreCliente,
                                        PrecioUnitario = PrecioUni,
                                        PrecioEspecial = Precio_iva_
                                    };

                                    listaClientes.Add(cliente);

                                    iRow++;
                                }

                                //VARIABLES DE INSERCION
                                string codigoProducto = sl.GetCellValueAsString(4, 3);
                                string Descripcion = "";
                                int idProducto = 0;

                                decimal Precio = 0, Precio_iva = 0;
                                string Cliente = "";
                                int idCliente = 0;

                                try
                                {
                                    Conexion.Open();

                                    foreach (InforGrid cliente in listaClientes)
                                    {
                                        if (cliente.Codigo != "" && cliente.PrecioEspecial != "")
                                        {
                                            Precio_iva = decimal.Parse(cliente.PrecioEspecial);
                                            Precio = Math.Round(Precio_iva / 1.13M, 2);

                                            //SELECCIONAR EL ID DEL CLIENTE
                                            using (SqlCommand cmd = new SqlCommand($"SELECT id, cliente FROM clientes WHERE codigo=@Codigo", Conexion))
                                            {
                                                cmd.CommandType = CommandType.Text;
                                                cmd.Parameters.AddWithValue("@Codigo", cliente.Codigo.Trim());
                                                SqlDataReader reader = cmd.ExecuteReader();
                                                while (reader.Read())
                                                {
                                                    idCliente = reader.GetInt32(reader.GetOrdinal("Id"));
                                                    Cliente = reader.GetString(reader.GetOrdinal("Cliente"));
                                                }
                                                reader.Close();
                                            }

                                            //SELECCIONANDO EL ID DEL PRODUCTO
                                            using (SqlCommand cmdProducto = new SqlCommand($"SELECT id, descripcion FROM inventario WHERE codigo=@CodigoProducto", Conexion))
                                            {
                                                cmdProducto.CommandType = CommandType.Text;
                                                cmdProducto.Parameters.AddWithValue("@CodigoProducto", codigoProducto);
                                                SqlDataReader reader = cmdProducto.ExecuteReader();
                                                while (reader.Read())
                                                {
                                                    idProducto = reader.GetInt32(reader.GetOrdinal("Id"));
                                                    Descripcion = reader.GetString(reader.GetOrdinal("Descripcion"));
                                                }
                                                reader.Close();
                                            }

                                            //BUSCANDO ID PRODUCTO Y CLIENTE PARA VERIFICAR SI YA EXISTE EN TABLA CLIENTES_PRECIOS
                                            using (SqlCommand cmdSelect = new SqlCommand($"SELECT id_cliente, id_inventario FROM Clientes_precios WHERE id_cliente=@idCliente AND id_inventario=@idProducto", Conexion))
                                            {
                                                cmdSelect.CommandType = CommandType.Text;
                                                cmdSelect.Parameters.AddWithValue("@idCliente", idCliente);
                                                cmdSelect.Parameters.AddWithValue("@idProducto", idProducto);

                                                SqlDataReader rdr = cmdSelect.ExecuteReader();

                                                if (rdr.HasRows)
                                                {
                                                    rdr.Close();
                                                    //ACTUALIZANDO EL PRECIO AL CLIENTE EN LA TABLA CLIENTES_PRECIOS
                                                    using (SqlCommand update = new SqlCommand($"UPDATE Clientes_precios SET precio_p=@precio_p, precio_p_iva=@precio_p_iva WHERE id_cliente=@idCliente AND id_inventario=@idInventario", Conexion))
                                                    {
                                                        update.CommandType = CommandType.Text;
                                                        update.Parameters.AddWithValue("@precio_p", Precio);
                                                        update.Parameters.AddWithValue("@precio_p_iva", Precio_iva);
                                                        update.Parameters.AddWithValue("@idCliente", idCliente);
                                                        update.Parameters.AddWithValue("@idInventario", idProducto);

                                                        update.ExecuteNonQuery();
                                                    }
                                                }
                                                else
                                                {
                                                    rdr.Close();
                                                    //INSERTANDO LOS DATOS EN LA TABLA CLIENTES_PRECIOS
                                                    using (SqlCommand insrt = new SqlCommand($"INSERT INTO Clientes_precios(Id_cliente, Cliente, Id_inventario, Codigo_producto, Descripcion, precio_p, precio_p_iva) VALUES (@idCliente, @Cliente, @idInventario, @codigoProducto, @Descripcion, @precio_p, @precio_p_iva)", Conexion))
                                                    {
                                                        insrt.CommandType = CommandType.Text;
                                                        insrt.Parameters.AddWithValue("@idCliente", idCliente);
                                                        insrt.Parameters.AddWithValue("@Cliente", Cliente);
                                                        insrt.Parameters.AddWithValue("@idInventario", idProducto);
                                                        insrt.Parameters.AddWithValue("@codigoProducto", codigoProducto);
                                                        insrt.Parameters.AddWithValue("@Descripcion", Descripcion);
                                                        insrt.Parameters.AddWithValue("@precio_p", Precio);
                                                        insrt.Parameters.AddWithValue("@precio_p_iva", Precio_iva);

                                                        insrt.ExecuteNonQuery();
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    // Manejo de excepciones aquí
                                    Console.WriteLine("Error: " + ex.Message);
                                }
                                finally
                                {
                                    // Cerrar la conexión y otros recursos aquí
                                    if (Conexion.State == ConnectionState.Open)
                                    {
                                        Conexion.Close();
                                    }
                                    MessageBox.Show("INFORMACIÓN ALMACENADA CORRECTAMENTE \n "+ NombreHojas +"", "INFORMACIÓN", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                    btnImportar.Enabled = false;
                                }
                            }
                        }
                    }
                }
            }
        }

        private void btnSalir_Click(object sender, EventArgs e)
        {
            DialogResult resultado = MessageBox.Show("¿Deseas Salir de la Aplicación?", "Confirmación", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);

            if (resultado == DialogResult.OK)
            {
                Application.Exit();
            }
            
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            var sm = GetSystemMenu(Handle, false);
            EnableMenuItem(sm, SC_CLOSE, MF_BYCOMMAND | MF_DISABLED);
        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }
    }
}
