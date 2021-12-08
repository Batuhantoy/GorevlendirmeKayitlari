using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace Görevlendirme_Kayıtları
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        string bagExcel = "Provider=Microsoft.Jet.Oledb.4.0;"+"Data Source=vt_gorevlendirme.xls"+"Extended Properties= Excel 8.0";
        string komut = "select * from [Şehir Dışı$]";
        private void button1_Click(object sender, EventArgs e)
        {

        }
        private void Form1_Load(object sender, EventArgs e)
        {
            kayitgetir();
        } 
        public void kayitgetir()
        {
            OleDbDataAdapter adp = new OleDbDataAdapter(komut,bagExcel);
            DataSet ds = new DataSet();
            adp.Fill(ds,"ExcelVerileri");
            dataGridView1.DataSource = ds.Tables["ExcelVerileri"].DefaultView;
        }
    }
}

