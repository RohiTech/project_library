using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;

/* Agregamos nuestro proveedor de Base de Datos */
using System.Data.SqlClient;

public class DALUsuario
{
    private SqlConnection coneccion;
    private SqlCommand comando;

	public DALUsuario()
	{
        /*
        Primero Incializamos La Coneccion.
        * Luego con el ConfigurationManager.ConnectionString 
        *      obtenemos acceso a los Datos de la Configuracion 
        *      connectionString determinada en nuestro 
        *      Web.Config estableciendo la cadena de coneccion.
        */
        coneccion = new SqlConnection(
               ConfigurationManager.ConnectionStrings[
                       "ConnectionString"].ConnectionString);

        //Instanciamos el Comando
        comando = new SqlCommand();

        //Establecemos el tipo de cadena de comando y 
        //lo establecemos como comando de Texto SQL
        comando.CommandType = CommandType.Text;
	}

    // Seleccionamos todos los Datos
    public DataSet SeleccionarDatos()
    {
        //Establecemos nuestra Sentencia SQL
        //SELECT: Seleccionamos los campos
        //ponemos los campos y decinmos extraiga la TablaUsuario
        comando.CommandText = "SELECT [ID], [Nombre]," +
            "[Apellidos], [Email] FROM [TablaUsuario]";

        //Indicamo la Coneccion que va tener esta instancia
        comando.Connection = coneccion;

        //Creamos nuestro DataSet
        DataSet mydataset = new DataSet();

        try
        {
            //Abrimos la coneccion
            coneccion.Open();

            //Instanciamos DataReader con Sentencia de Sql y coneccion.
            SqlDataAdapter da = new SqlDataAdapter(comando.CommandText, 
                                            comando.Connection);
            //Agregamos las fila al DataSet
            da.Fill(mydataset, "TablaUsuario");
        }
        catch (Exception e)
        { throw; }
        finally 
        { //Cerramos la Coneccion
            coneccion.Close(); 
        }
        return mydataset;
    }

    // Eliminar Datos 
    public void eliminarDato(int id)
    {
        /*
         * Eliminamos los datos donde el Campo ID sea igual 
         * al id(Indice del Usuario) que queramos eliminar
         */
        comando.CommandText = "DELETE FROM [TablaUsuario] " +
                        "WHERE (ID=" + id + ")";
        comando.Connection = coneccion;
        try
        {
            coneccion.Open();
            //Ejecutamos la sentencia SQL del comando establecido
            comando.ExecuteNonQuery();
        }
        catch (Exception ex)
        { throw; }
        finally { coneccion.Close(); }
    }

    //Insertar Datos
    public void insertarDato(int id, string nombre, string apellidos, string email)
    {
        /*
        * Insertamos a la TablaUsuario en la los Campos necesarios
        * los nuevos valores teniendo en cuanta el tipo del campo
        */
        comando.CommandText = "INSERT INTO [TablaUsuario]([ID], [Nombre], [Apellidos], [Email])" +
                    " VALUES (" + id + ", '" + nombre + "', '" + apellidos + "', '" + email + "')";
        comando.Connection = coneccion;

        try
        {
            coneccion.Open();
            //Ejecutamos la sentencia SQL del comando establecido
            comando.ExecuteNonQuery();
        }
        catch (Exception ex)
        { throw; }
        finally { coneccion.Close(); }
    }

    //Actualizar Datos
    public void actualizarDatos(int id, string nombre, string apellidos, string email)
    {
        /*
         * Actualizar la TablaUsuario en el Campo con el nuevo valor
         * Cuando el Campo ID sea igual al seleccionado.
         */
        comando.CommandText = "UPDATE [TablaUsuario] SET [Nombre] = '" + nombre +
                "', [Apellidos] = '" + apellidos + "', [Email] = '" + email +
                "' WHERE (ID = " + id + ")";

        comando.Connection = coneccion;

        try
        {
            coneccion.Open();
            //Ejecutamos la sentencia SQL del comando establecida
            comando.ExecuteNonQuery();
        }
        catch (Exception ex)
        { throw; }
        finally { coneccion.Close(); }
    }
}
