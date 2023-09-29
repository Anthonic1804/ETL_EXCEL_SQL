using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;

namespace ETL_EXCEL_SQL
{
    public class FuncionesController
    {
        private readonly SqlConnection Conexion;

        public FuncionesController(SqlConnection conexion)
        {
            Conexion = conexion ?? throw new ArgumentNullException(nameof(conexion));
        }

        // Función para insertar un nuevo precio
        public void InsertarPrecio(int idCliente, string Cliente, int idProducto, string codigoProducto, string Descripcion, decimal Precio, decimal Precio_iva)
        {
            using (SqlCommand insrt = Conexion.CreateCommand())
            {
                insrt.CommandType = CommandType.Text;
                insrt.CommandText = "INSERT INTO Clientes_precios(Id_cliente, Cliente, Id_inventario, Codigo_producto, Descripcion, precio_p, precio_p_iva) VALUES (@idCliente, @Cliente, @idProducto, @codigoProducto, @Descripcion, @Precio, @Precio_iva)";

                insrt.Parameters.AddWithValue("@idCliente", idCliente);
                insrt.Parameters.AddWithValue("@Cliente", Cliente);
                insrt.Parameters.AddWithValue("@idProducto", idProducto);
                insrt.Parameters.AddWithValue("@codigoProducto", codigoProducto);
                insrt.Parameters.AddWithValue("@Descripcion", Descripcion);
                insrt.Parameters.AddWithValue("@Precio", Precio);
                insrt.Parameters.AddWithValue("@Precio_iva", Precio_iva);

                insrt.ExecuteNonQuery();
            }
        }

        // Función para actualizar un precio existente
        public void ActualizarPrecio(decimal Precio, decimal Precio_iva, int idCliente, int idProducto)
        {
            using (SqlCommand update = Conexion.CreateCommand())
            {
                update.CommandType = CommandType.Text;
                update.CommandText = "UPDATE Clientes_precios SET precio_p=@Precio, precio_p_iva=@Precio_iva WHERE id_cliente=@idCliente AND id_inventario=@idProducto";

                update.Parameters.AddWithValue("@Precio", Precio);
                update.Parameters.AddWithValue("@Precio_iva", Precio_iva);
                update.Parameters.AddWithValue("@idCliente", idCliente);
                update.Parameters.AddWithValue("@idProducto", idProducto);

                update.ExecuteNonQuery();
            }
        }

        // Buscar ID producto y cliente para verificar si ya existe en la tabla Clientes_precios
        public bool BuscarClientePrecio(int idCliente, int idProducto)
        {
            using (SqlCommand cmdSelect = Conexion.CreateCommand())
            {
                cmdSelect.CommandType = CommandType.Text;
                cmdSelect.CommandText = "SELECT id_cliente, id_inventario FROM Clientes_precios WHERE id_cliente=@idCliente AND id_inventario=@idProducto";

                cmdSelect.Parameters.AddWithValue("@idCliente", idCliente);
                cmdSelect.Parameters.AddWithValue("@idProducto", idProducto);

                using (SqlDataReader reader = cmdSelect.ExecuteReader())
                {
                    return reader.Read();
                }
            }
        }
    }

}
