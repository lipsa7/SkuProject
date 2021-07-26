using ClosedXML.Excel;
using ExcelDataReader;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using SkuProject8.Models;
using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace SkuProject8.Controllers
{
    public class SkuController : Controller
    {
        [HttpGet]
        public IActionResult Index(List<Models.Sku> skus = null)
        {
            skus = skus == null ? new List<Models.Sku>() : skus;
            return View(skus);
        }

        [HttpPost]
        public IActionResult Index(IFormFile file, [FromServices] IWebHostEnvironment hostingEnvironment)
        {
            string fileName = $"{hostingEnvironment.WebRootPath}\\files\\file.FileName";
            using (FileStream fileStream = System.IO.File.Create(fileName))
            {
                file.CopyTo(fileStream);
                fileStream.Flush();
            }

            var skus = this.GetSkuList(file.FileName);
            //Skudb dbop = new Skudb();
            //try
            //{
            //    if (ModelState.IsValid)
            //    {
            //        string res = dbop.SaveRecord(sku);
            //        TempData["msg"] = res;
            //    }
            //}
            //catch (Exception ex)
            //{
            //    TempData["msg"] = ex.Message;
            //}
            return Index(skus);
        }

        private List<Models.Sku> GetSkuList(string fName)
        {
            List<Models.Sku> skus = new List<Models.Sku>();
            var fileName = $"{Directory.GetCurrentDirectory()}{@"\wwwroot\files"}" + "\\" + fName;
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using (var stream = System.IO.File.Open(fileName, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    while (reader.Read())
                    {
                        skus.Add(new Models.Sku()
                        {
                            Id = reader.GetValue(0).ToString(),
                            Sku_Id = reader.GetValue(1).ToString(),
                            Comments = reader.GetValue(2).ToString()

                        });
                    }
                }
            }

            SqlConnection con = new SqlConnection("Data Source=LAPTOP-L151761A\\SQLEXPRESS;Initial Catalog=code;Integrated Security=True");

            string query = "Delete from Sku_Table2";
            SqlCommand cmd = new SqlCommand(query, con);
            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();

            //DataTable dt = new DataTable("Sku_Table2");
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


            try
            {
                foreach(var sku in skus)
                {
                    SqlCommand com = new SqlCommand("sp_sku_add", con);
                    com.CommandType = CommandType.StoredProcedure;
                    com.Parameters.AddWithValue("@Id", sku.Id);
                    com.Parameters.AddWithValue("@Sku_Id", sku.Sku_Id);
                    com.Parameters.AddWithValue("@Comments", sku.Comments);
                    con.Open();
                    com.ExecuteNonQuery();
                    con.Close();
                    //return ("OK");
                }


            }
            catch (Exception e)
            {
                if (con.State == ConnectionState.Open)
                {
                    con.Close();
                }
            }

            //string query = "Delete from Sku_Table2";
            //SqlCommand cmd = new SqlCommand(query, con);
            //con.Open();
            //cmd.ExecuteNonQuery();
            //con.Close();

            //SqlConnection con = new SqlConnection("Data Source=LAPTOP-L151761A\\SQLEXPRESS;Initial Catalog=code;Integrated Security=True");

            string query1 = "Delete from Sku_TableFinal";
            SqlCommand cmd1 = new SqlCommand(query1, con);
            con.Open();
            cmd1.ExecuteNonQuery();
            con.Close();

            string query2 = "Insert into Sku_TableFinal (Id, Sku_Id, Comments, CommentsFirstLetter) select Id, Sku_Id, Comments," +
                "substring(Comments, 1,1) from Sku_Table2";
            SqlCommand cmd2 = new SqlCommand(query2, con);
            con.Open();
            cmd2.ExecuteNonQuery();
            con.Close();

            return skus;
        }



        private List<Models.SkuFinalList> skusFin = new List<Models.SkuFinalList>();


        public IActionResult Excel()
        {
            //    //cmd.CommandText = "SELECT * from Sku_TableFinal";
            //    //SqlDataReader dr = cmd.ExecuteReader();
            //    //while (dr.Read())
            //    //{
            //    //    skus.Add(dr.GetValue(0).ToString());
            //    //}

            using (SqlConnection cn = new SqlConnection("Data Source=LAPTOP-L151761A\\SQLEXPRESS;Initial Catalog=code;Integrated Security=True"))
            {

                cn.Open();
                SqlCommand sqlCommand = new SqlCommand("SELECT * FROM Sku_TableFinal", cn);
                SqlDataReader reader = sqlCommand.ExecuteReader();
                while (reader.Read())
                {
                    //Fruitee.add(reader["aID"], reader["bID"], reader["name"]) // ??? not sure what to put here  as add is not available
                    SkuFinalList s = new SkuFinalList();
                    s.Id = (string)reader["Id"];
                    s.Sku_Id = (string)reader["Sku_Id"];
                    s.Comments = (string)reader["Comments"];
                    s.CommentsFirstLetter = (string)reader["CommentsFirstLetter"];
                    skusFin.Add(s);
                }
                cn.Close();
            }


            //foreach (string s in skus)
            //{
            //    Console.WriteLine(s);
            //}
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("SkuFinal");
                var currentRow = 1;
                worksheet.Cell(currentRow, 1).Value = "Id";
                worksheet.Cell(currentRow, 2).Value = "Sku_Id";
                worksheet.Cell(currentRow, 3).Value = "Comments";
                worksheet.Cell(currentRow, 4).Value = "CommentsFirstLetter";
                foreach (var sku in skusFin)
                {
                    currentRow++;
                    worksheet.Cell(currentRow, 1).Value = sku.Id;
                    worksheet.Cell(currentRow, 2).Value = sku.Sku_Id;
                    worksheet.Cell(currentRow, 3).Value = sku.Comments;
                    worksheet.Cell(currentRow, 4).Value = sku.CommentsFirstLetter;


                }

                var stream = new MemoryStream();
                

                    workbook.SaveAs(stream);
                    stream.Position = 0;
                    var content = stream.ToArray();

                    return new FileStreamResult(
                        stream,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    { FileDownloadName = "SkuFinal.xlsx" };
                    // "SkuFinal.xlsx");
                
                
                
            }
        }

        //Skudb dbop = new Skudb();

        //public IActionResult Save()
        //{
        //    return View();
        //}

        //[HttpPost]
        //public IActionResult Save([Bind] Sku sku)
        //{

        //    try
        //    {
        //        if (ModelState.IsValid)
        //        {
        //            string res = dbop.SaveRecord(sku);
        //            TempData["msg"] = res;
        //        }
        //    }
        //    catch(Exception ex)
        //    {
        //        TempData["msg"] = ex.Message;
        //    }

        //    return View();
        //}
        //    public string SaveRecord([Bind] Models.Sku sku)
        //    {
        //        SqlConnection con = new SqlConnection("Data Source=LAPTOP-L151761A\\SQLEXPRESS;Initial Catalog=code;Integrated Security=True");

        //        var fileName = $"{Directory.GetCurrentDirectory()}{@"\wwwroot\files"}" + "\\" + fName;
        //        var skus = this.GetSkuList(fileName);
        //        //var skus = List<Models.Sku>();

        //        try
        //        {
        //            DataTable dt = new DataTable("Sku_Table2");
        //            //SqlConnection conn = new SqlConnection(connString);
        //            string query = "select * from Sku_Table2";
        //            SqlCommand cmd = new SqlCommand(query, con);
        //            con.Open();
        //            SqlDataAdapter da = new SqlDataAdapter(cmd);
        //            da.Fill(dt);
        //            DataSet ds = new DataSet();

        //            if (dt.Rows.Count != 0)
        //            {
        //                dt.Clear();
        //            }
        //            con.Close();

        //            foreach (var sku in Models.Sku)
        //            {
        //                //SqlConnection con = new SqlConnection("Data Source=LAPTOP-L151761A\\SQLEXPRESS;Initial Catalog=code;Integrated Security=True");
        //                SqlCommand com = new SqlCommand("sp_sku_add", con);
        //                com.CommandType = CommandType.StoredProcedure;
        //                com.Parameters.AddWithValue("@Id", sku.Id);
        //                com.Parameters.AddWithValue("@Sku_Id", sku.Sku_Id);
        //                com.Parameters.AddWithValue("@Comments", sku.Comments);
        //                con.Open();
        //                com.ExecuteNonQuery();
        //                con.Close();
        //            }
        //            return ("OK");
        //        }

        //        catch (Exception e)
        //        {
        //            if (con.State == ConnectionState.Open)
        //            {
        //                con.Close();
        //            }

        //            return e.Message.ToString();
        //        }

        //    }
        //}
    }
}

