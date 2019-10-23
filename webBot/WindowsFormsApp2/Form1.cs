using System;
using System.Collections;
using System.Net;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;


namespace WindowsFormsApp2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
        
        private void button1_Click(object sender, EventArgs e)
        {
         
            string link = textBox1.Text;
            Uri url = new Uri(link);

            WebClient client = new WebClient();
            client.Encoding = Encoding.UTF8;

            string html = client.DownloadString(url);

            HtmlAgilityPack.HtmlDocument document = new HtmlAgilityPack.HtmlDocument();
            document.LoadHtml(html);
            
            var secilenhtml = @"//*[@id=""main_container""]";

            StringBuilder st = new StringBuilder();

            var secilenHtmlList = document.DocumentNode.SelectNodes(secilenhtml);

            string[] veriler = new string[1000];
            //ArrayList veriler = new ArrayList();
            int i = 0;
            foreach (var items in secilenHtmlList)
            {
                foreach (var innerItem in items.SelectNodes(textBox3.Text))
                {
                    foreach (var item in innerItem.SelectNodes(textBox4.Text))
                    {
                        var classValue = item.Attributes["class"] == null ? null : item.Attributes["class"].Value;
                        if (classValue == "product-title")
                        {
                            if (radioButton1.Checked) {
                                //st.AppendLine(item.InnerText);
                                // listView1.Items.Add(item.InnerText.ToString()); //cekilen verileri direk listview a tasir
                                //veriler[i].Replace("akü", "akü-");
                                veriler[i] = item.InnerText.Replace(textBox2.Text, textBox5.Text);
                                i++;
                            }
                            if (radioButton2.Checked) {
                                veriler[i] = item.InnerText;
                                i++;
                            }
                        }

                    }
                }

            }
            
            int j;
            for (j = 0; j < i; j++)
            {
                listView1.Items.Add(veriler[j]);
            }


        }

        private void button2_Click(object sender, EventArgs e)
        {
            listView1.Items.Clear();
        }

        private void button3_Click(object sender, EventArgs e)
        {
           
            Excel.Application xla = new Excel.Application();
            xla.Visible = true;
            Workbook wb = xla.Workbooks.Add(XlSheetType.xlWorksheet);

            Worksheet ws = (Worksheet)xla.ActiveSheet;
            int k = 1;
            int l = 1;
            foreach (ListViewItem item in listView1.Items)
            {
                ws.Cells[k, l] = item.Text.ToString();
                foreach (ListViewItem.ListViewSubItem subitem in item.SubItems)
                {
                    ws.Cells[k, l] = subitem.Text.ToString();
                    l++;
                }
                l = 1;
                k++;
            }
        }
    }
}
        