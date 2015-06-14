using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using System.Data.Odbc;
using System.Data.OleDb;

namespace 推送器
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        string filePath = "";
        DataSet excelTable = null;
        private void button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            richTextBox1.Enabled = false;
            excelTable = ToDataTable(filePath);
            if (excelTable == null)
                return;
            foreach (DataTable dt in excelTable.Tables)
            {
                foreach (DataRow dr in dt.Rows)
                {
                  //  foreach (DataColumn dc in dt.Columns)
                 //   {
                         //richTextBox1.Text += dr[dc].ToString()+" ";
                    DataBaseData NewData = new DataBaseData()
                    {
                        Id = dr[0].ToString(),
                        Dt = dr[1].ToString(),
                        Sid = dr[2].ToString(),
                        Item = dr[3].ToString(),
                        Val = dr[4].ToString(),
                        Unit = dr[5].ToString(),
                    };
                    IsoDateTimeConverter iso = new IsoDateTimeConverter();
                    //  iso.DateTimeFormat = "yyyy-MM-dd HH:mm:ss";
                    String str = JsonConvert.SerializeObject(NewData, iso);
                    richTextBox1.Text += str;
                   // }
                    richTextBox1.Text += "\n";
                }
            }
            excelTable = null;
            button1.Enabled = true;
            richTextBox1.Enabled = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Multiselect = true;
            fileDialog.Title = "请选择文件";
            fileDialog.Filter = "xlsx(*.*)|*.xlsx";
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                filePath = fileDialog.FileName;
              //  MessageBox.Show("已选择文件:" + filePath, "选择文件提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                textBox_filePath.Text = filePath;
            }
        }

        public static DataSet ToDataTable(string filePath)
        {

            try  
            {  
                string strConn;
                strConn = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + filePath + ";" + ";Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1\"";  
                OleDbConnection OleConn = new OleDbConnection(strConn);  
                OleConn.Open();  
                String sql = "SELECT * FROM  [Sheet1$]";//可是更改Sheet名称，比如sheet2，等等   
  
                OleDbDataAdapter OleDaExcel = new OleDbDataAdapter(sql, OleConn);  
                DataSet OleDsExcle = new DataSet();  
                OleDaExcel.Fill(OleDsExcle, "Sheet1");  
                OleConn.Close();  
                return OleDsExcle;  
            }  
            catch (Exception err)  
            {  
                MessageBox.Show("数据绑定Excel失败!失败原因：" + err.Message, "提示信息",  
                    MessageBoxButtons.OK, MessageBoxIcon.Information);  
                return null;  
            }  

         }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void textBox_filePath_TextChanged(object sender, EventArgs e)
        {
            filePath = textBox_filePath.Text;
        }
    }

    public class DataBaseData
    {
        public string Id { get; set; }
        public string Dt { get; set; }
        public string Sid { get; set; }
        public string Item { get; set; }
        // public int Index { get; set; }
        public string Val { get; set; }
        public string Unit { get; set; }
      
    }
}
