using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Data.OleDb;
using System.Configuration;



namespace Excel_to_SQL_Data_Import
{
    public partial class ExcelToSQL : System.Web.UI.Page
    {
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ConnectionString"].ToString());
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            { }

        }

        protected void btnUploadExcel_Click(object sender, EventArgs e)
        {
            string folderPath = Server.MapPath("~/ExcelFile/");

            //Check whether Directory (Folder) exists.
            if (!Directory.Exists(folderPath))
            {
                //If Directory (Folder) does not exists. Create it.
                Directory.CreateDirectory(folderPath);
            }

            string SL;
            String Name;
            String Mobile;
            String Email;
            string path = Path.GetFileName(FileUploadExcel.FileName);
            path = path.Replace(" ", "");

            // if Excel File Exists Start-------
            int count = 0;
            Find:
            if (File.Exists(path))
            {
                path = path + "(" + count.ToString() + ").xlsx";
                count++;
                goto Find;
            }
            else
            {
                //Add your logic here
                FileUploadExcel.SaveAs(Server.MapPath("~/ExcelFile/") + path);
            }
            // if Excel File Exists END----------


            String ExcelPath = Server.MapPath("~/ExcelFile/") + path;

            //Excel 2007 Connection
            //OleDbConnection mycon = new OleDbConnection("Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " + ExcelPath + "; Extended Properties=Excel 8.0; Persist Security Info = False");

            //Excel 2013 Connection
            OleDbConnection mycon = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + ExcelPath + ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1;';");

            mycon.Open();
            OleDbCommand cmd = new OleDbCommand("select * from [Sheet1$]", mycon);
            OleDbDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                // Response.Write("<br/>"+dr[0].ToString());
                //rollno = Convert.ToInt32(dr[0].ToString());
                SL = dr[0].ToString();

                Name = dr[1].ToString();
                Mobile = dr[2].ToString();
                Email = dr[3].ToString();

                //Function Use
                SaveDataSQL(SL, Name, Mobile, Email);


            }
            ScriptManager.RegisterStartupScript(this, this.GetType(), "Message", "alert('Successfuly Imported Excel Sheet To SQL Server');", true);


        }

        private void SaveDataSQL(String SL1, String Name1, String Mobile1, String Email1)
        {
            String query = "insert into tbl_ExcelUploadData(SL,Name,Mobile,Email) values(" + SL1 + ",'" + Name1 + "','" + Mobile1 + "','" + Email1 + "')";
            
            SqlCommand cmd = new SqlCommand(query,conn);
            conn.Open();
            cmd.ExecuteNonQuery();
            conn.Close();
        }

        protected void lbDownload_Click(object sender, EventArgs e)
        {
            //FileName AND Extension to be downloaded.
            string fileName = "UploadExcelSheetDemo.xlsx";

            //Path of the File to be downloaded.
            string filePath = Server.MapPath(string.Format("~/ExcelDemo/{0}", fileName));

            //Content Type for excel files and Header.
            Response.ContentType = "application/vnd.ms-excel";
            Response.AppendHeader("Content-Disposition", "attachment; filename=" + fileName);

            //Writing the File to Response Stream.
            Response.WriteFile(filePath);

            //Flushing the Response.
            Response.Flush();
            Response.End();
        }
    }
}