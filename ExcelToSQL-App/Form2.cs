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
using ExcelApp = Microsoft.Office.Interop.Excel;

namespace ExcelToSQL_App
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string DosyaYolu;
            string DosyaAdi;
            DataTable dt;
            OpenFileDialog file = new OpenFileDialog();
            file.Filter = "Excel Dosyası | *.xls; *.xlsx; *.xlsm";
            if (file.ShowDialog() == DialogResult.OK)
            {
                DosyaYolu = file.FileName;// seçilen dosyanın tüm yolunu verir
                DosyaAdi = file.SafeFileName;// seçilen dosyanın adını verir.
                ExcelApp.Application excelApp = new ExcelApp.Application();
                if (excelApp == null)
                { //Excel Yüklümü Kontrolü Yapılmaktadır.
                    MessageBox.Show("Excel yüklü değil.");
                    return;
                }
                //Excel Dosyası Açılıyor.
                ExcelApp.Workbook excelBook = excelApp.Workbooks.Open(DosyaYolu);
                //Excel Dosyasının Sayfası Seçilir.
                ExcelApp._Worksheet excelSheet = excelBook.Sheets[1];
                //Excel Dosyasının ne kadar satır ve sütun kaplıyorsa tüm alanları alır.
                ExcelApp.Range excelRange = excelSheet.UsedRange;
                int satirSayisi = excelRange.Rows.Count; //Sayfanın satır sayısını alır.
                int sutunSayisi = excelRange.Columns.Count;//Sayfanın sütun sayısını alır.
                dt = ToDataTable(excelRange, satirSayisi, sutunSayisi);
                dataGridView1.DataSource = dt;
                dataGridView1.Refresh();
                //Okuduktan Sonra Excel Uygulamasını Kapatıyoruz.
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                label1.Text = DosyaAdi;
            }
            else
            {
                MessageBox.Show("Dosya Seçilemedi.");
            }
        }
        public DataTable ToDataTable(ExcelApp.Range range, int rows, int cols)
        {
            DataTable table = new DataTable();
            for (int i = 1; i <= rows; i++)
            {
                if (i == 1)
                { // ilk satırı Sutun Adları olarak kullanıldığından
                  // bunları Sutün Adları Olarak Kaydediyoruz.
                    for (int j = 1; j <= cols; j++)
                    {
                        //Sütunların içeriği boş mu kontrolü yapılmaktadır.
                        if (range.Cells[i, j] != null && range.Cells[i, j].Value2 != null)
                            table.Columns.Add(range.Cells[i, j].Value2.ToString());
                        else //Boş olduğunda Kaçınsı Sutünsa Adı veriliyor.
                            table.Columns.Add(j.ToString() + ".Sütun");
                    }
                    continue;
                }
                //Yukarıda Sütunlar eklendi
                // onun şemasına göre yeni bir satır oluşturuyoruz. 
                //Okunan verileri yan yana sıralamak için
                var yeniSatir = table.NewRow();
                for (int j = 1; j <= cols; j++)
                {
                    //Sütunların içeriği boş mu kontrolü yapılmaktadır.
                    if (range.Cells[i, j] != null && range.Cells[i, j].Value2 != null)
                        yeniSatir[j - 1] = range.Cells[i, j].Value2.ToString();
                    else // İçeriği boş hücrede hata vermesini önlemek için
                        yeniSatir[j - 1] = String.Empty;
                }
                table.Rows.Add(yeniSatir);
            }
            return table;
        }
        OleDbConnection baglanti = new OleDbConnection("Provider=Microsoft.jet.Oledb.4.0;Data Source=Database1.mdb");
        private void button3_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text != "")
            {
                baglanti.Open();
                for (int i = 0; i < (dataGridView1.Rows.Count) - 1; i++)
                {
                    OleDbCommand cmd = new OleDbCommand("insert into " + comboBox1.Text + " values (@p1,@p2,@p3)", baglanti);
                    cmd.Parameters.AddWithValue("@p1", dataGridView1.Rows[i].Cells[0].Value);
                    cmd.Parameters.AddWithValue("@p2", dataGridView1.Rows[i].Cells[1].Value);
                    cmd.Parameters.AddWithValue("@p3", dataGridView1.Rows[i].Cells[2].Value);
                    cmd.ExecuteNonQuery();
                    cmd.Parameters.Clear();
                    /*
                    MessageBox.Show((dataGridView1.Rows[i].Cells[0].Value).ToString());
                    MessageBox.Show((dataGridView1.Rows[i].Cells[1].Value).ToString());
                    MessageBox.Show((dataGridView1.Rows[i].Cells[2].Value).ToString());
                    */
                }
                baglanti.Close();
                MessageBox.Show("Başarıyla Kayıt Edildi !");
            }
            else
            {
                MessageBox.Show("Kayıt Edilecek İlçe Seçiniz !");
            }
        }
    }
}
