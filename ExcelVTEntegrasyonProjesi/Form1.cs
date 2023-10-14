using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelVTEntegrasyonProjesi
{
    public partial class Form1 : Form
    {
        SqlConnection baglanti = new SqlConnection(@"Data Source=LAPTOP-FB9086OP;Initial Catalog=ExcelVTEntegrasyonu;Integrated Security=True");

        public Form1()
        {
            InitializeComponent();
        }


        private void btnVTdenOku_Click(object sender, EventArgs e)
        {

            Microsoft.Office.Interop.Excel.Application excelUygulama = new Microsoft.Office.Interop.Excel.Application();
            excelUygulama.Visible = true;
            Microsoft.Office.Interop.Excel.Workbook workbook = excelUygulama.Workbooks.Add(System.Reflection.Missing.Value);
            Microsoft.Office.Interop.Excel.Worksheet sayfa1 = workbook.Sheets[1];

            string[] baslikler = { "Personel no", "Ad", "Soyad", "Semt", "Şehir" };
            Microsoft.Office.Interop.Excel.Range range;
            for (int i = 0; i < baslikler.Length; i++)
            {
                range = sayfa1.Cells[1, (1 + i)];
                range.Value2 = baslikler[i];
            }



            try
            {
                baglanti.Open();
                string sqlCumlesi = "SELECT PersonelNo, Ad, Soyad, Semt, Sehir FROM Personel ";
                SqlCommand sqlKomut = new SqlCommand(sqlCumlesi, baglanti);
                SqlDataReader sdr = sqlKomut.ExecuteReader();

                int satir = 2;
                while (sdr.Read())
                {
                    string pno = sdr[0].ToString();
                    string ad = sdr[1].ToString();
                    string soyad = sdr[2].ToString();
                    string semt = sdr[3].ToString();
                    string sehir = sdr[4].ToString();
                    richTextBox1.Text = richTextBox1.Text + pno + "  " + ad + "  " + soyad + "  " + semt + "  " + sehir + "\n";

                    range = sayfa1.Cells[satir, 1];
                    range.Value2 = pno;
                    range = sayfa1.Cells[satir, 2];
                    range.Value2 = ad;
                    range = sayfa1.Cells[satir, 3];
                    range.Value2 = soyad;
                    range = sayfa1.Cells[satir, 4];
                    range.Value2 = semt;
                    range = sayfa1.Cells[satir, 5];
                    range.Value2 = sehir;
                    satir++;
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show("SQL Query sırasında bir hata oluştu, Hata Kodu: SQLREAD01 \n" + ex.ToString());
            }
            finally
            {
                if(baglanti != null)
                baglanti.Close();
            }
        }
    }
}
