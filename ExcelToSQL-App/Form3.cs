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

namespace ExcelToSQL_App
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
        }

        OleDbConnection baglanti = new OleDbConnection("Provider=Microsoft.jet.Oledb.4.0;Data Source=Database1.mdb");
        OleDbDataAdapter da = new OleDbDataAdapter();
        DataSet ds = new DataSet();
        private void yenile(string tablo)
        {
            baglanti.Open();
            da = new OleDbDataAdapter("select * from "+tablo+"", baglanti);
            ds = new DataSet();
            da.Fill(ds, "veri");
            dataGridView1.DataSource = ds.Tables["veri"];
            baglanti.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OleDbDataAdapter da1 = new OleDbDataAdapter();
            if (comboBox1.Text != "" && comboBox2.Text != "")
            {
                if (textBox1.Text != "")
                {
                    baglanti.Open();
                    da1 = new OleDbDataAdapter("select * from "+comboBox1.Text+" where "+comboBox2.Text+" like '"+textBox1.Text+"%'", baglanti);
                    DataSet ds1 = new DataSet();
                    da1.Fill(ds1, "veri");
                    dataGridView1.DataSource = ds1.Tables["veri"];
                    baglanti.Close();
                }
                else
                {
                    MessageBox.Show("Aranacak Veri Boş");
                }
            }
            else
            {
                MessageBox.Show("Aranacak Veri ya da Aranacak Kategori Boş");
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.Text != "")
            {
                yenile(comboBox1.Text);
                //comboboxtan seçilen ilçenin tablosunu getirtme
                for (int i = 0; i < dataGridView1.Columns.Count; i++)
                {
                    comboBox2.Items.Add(dataGridView1.Columns[i].Name.ToString());
                }
            }
        }
    }
}
