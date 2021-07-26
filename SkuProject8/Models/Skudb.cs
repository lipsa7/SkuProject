using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;

namespace SkuProject8.Models
{
    public class Skudb
    {
        SqlConnection con = new SqlConnection("Data Source=LAPTOP-L151761A\\SQLEXPRESS;Initial Catalog=code;Integrated Security=True");

        public string SaveRecord(Sku sku)
        {

            try
            {
            //    DataTable dt = new DataTable("Sku_Table2");
            ////SqlConnection conn = new SqlConnection(connString);
            //string query = "select * from Sku_Table2";
            //SqlCommand cmd = new SqlCommand(query, con);
            //con.Open();
            //SqlDataAdapter da = new SqlDataAdapter(cmd);
            //da.Fill(dt);
            //DataSet ds = new DataSet();

            //if (dt.Rows.Count != 0)
            //{
            //    dt.Clear();
            //}
            //con.Close();

            
                //SqlConnection con = new SqlConnection("Data Source=LAPTOP-L151761A\\SQLEXPRESS;Initial Catalog=code;Integrated Security=True");
                
                SqlCommand com = new SqlCommand("sp_sku_add", con);
                com.CommandType = CommandType.StoredProcedure;
                com.Parameters.AddWithValue("@Id", sku.Id);
                com.Parameters.AddWithValue("@Sku_Id", sku.Sku_Id);
                com.Parameters.AddWithValue("@Comments", sku.Comments);
                con.Open();
                com.ExecuteNonQuery();
                con.Close();
                return ("OK");
            }

            catch(Exception e)
            {
                if(con.State == ConnectionState.Open)
                {
                    con.Close();
                }

                return e.Message.ToString();
            }
        }
    }
}
