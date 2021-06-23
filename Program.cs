using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Collections;
using System.Net;
using System.IO;

namespace veriAnalizApp
{
    class Program
    {
        public static ArrayList veriListesi = new ArrayList(); // Excel dosyasındaki BP kısmının verileri
        public static string metinBelgesi = ""; // Metin belgesine yazdırılacak olan BP verilerinin çıktıları
        static void Main(string[] args)
        {
            ExcelVeriOkuma();
            VeriGonderme();
            DosyayaVeriYazma();
            Console.ReadKey();
            
        }

        public static void VeriGonderme()
        {
            string deger = "";
            for (int i = 0; i < veriListesi.Count; i++) // 5000 adet veriden kaç tanesini çekmek istiyorsanız belirtebilirsiniz.
            {
                string url = "https://www.uniprot.org/uniprot/" + veriListesi[i].ToString() + ".fasta";
                deger = Get(url);
                string[] lines = deger.Split('\n');
                deger = "";
                for (int j = 1; j < lines.Length; j++)
                {
                    deger += lines[j];
                }
                metinBelgesi += deger;
                metinBelgesi += "\n";
                Console.WriteLine(i + " İşlem Başarıyla Tamamlandı.");
            }
            Console.WriteLine("asdads");
        }

        public static void DosyayaVeriYazma()
        {
            StreamWriter sw = new StreamWriter(@"C:\Users\muham\Desktop\BPVeri.txt");
            sw.WriteLine(metinBelgesi);
            sw.Close();
        }


        public static string Get(string uri)
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(uri);
            request.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate;

            using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
            using (Stream stream = response.GetResponseStream())
            using (StreamReader reader = new StreamReader(stream))
            {
                return reader.ReadToEnd();
            }
            // Burada BP'deki her veri unitpro adresine gönderilip FASTA değeri alınıyor.
        }

        public static void ExcelVeriOkuma()
        {
            string con = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\muham\Desktop\Veri.xlsx;" + @"Extended Properties='Excel 8.0;HDR=Yes;'";
            OleDbConnection connection = new OleDbConnection(con);
            connection.Open();
            OleDbCommand command = new OleDbCommand("select * from [BP$]", connection);
            OleDbDataReader dr = command.ExecuteReader();
            while (dr.Read())
            {
                veriListesi.Add(dr[0]);
            }
            connection.Close();
        }
    }
}
