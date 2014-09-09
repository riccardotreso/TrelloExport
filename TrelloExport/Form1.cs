using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using TrelloNet;


namespace TrelloExport
{
    public partial class Form1 : Form
    {

        String appKey;
        String token;
        ITrello trello;
        public Form1()
        {
            InitializeComponent();

            // read user value
            if (File.Exists("data.txt"))
            {
                var s = File.ReadAllLines("data.txt");
                if (s.Length > 0)
                {
                    if (!string.IsNullOrEmpty(s[0]))
                        appKey = this.textBox1.Text = s[0];


                    if (s.Length > 1 && !string.IsNullOrEmpty(s[1]))
                        token = this.txtToken.Text = s[1];
                }
            }


        }


        private void btnExport_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(appKey) || appKey != this.textBox1.Text)
            {
                string[] tow = new string[] { this.textBox1.Text, token };
                File.WriteAllLines("data.txt", tow);
            }

            appKey = this.textBox1.Text;
            trello = new Trello(appKey);

            this.tabControl1.SelectedIndex = 1;


        }
        private void button2_Click(object sender, EventArgs e)
        {

            if (string.IsNullOrEmpty(token) || token != this.txtToken.Text)
            {
                string[] tow = new string[] { appKey, this.txtToken.Text };
                File.WriteAllLines("data.txt", tow);
            }

            token = this.txtToken.Text;
            trello.Authorize(token);


            this.tabControl1.SelectedIndex = 2;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.tabControl1.SelectedIndex = 0;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.tabControl1.SelectedIndex = 1;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            //DO Export
            Board BoardToExport;
            IEnumerable<Card> cards;
            try
            {
                BoardToExport = trello.Boards.WithId(this.txtBoardID.Text);
                cards = trello.Cards.ForBoard(BoardToExport, BoardCardFilter.Visible);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }

            // write excel here
            CreateExcelDoc excell_app = new CreateExcelDoc();
            //creates the main header
            excell_app.createHeaders(1, 1, "List name");
            excell_app.createHeaders(1, 2, "Name");
            excell_app.createHeaders(1, 3, "Description");
            excell_app.createHeaders(1, 4, "Closed");
            excell_app.createHeaders(1, 5, "Due Date");
            excell_app.createHeaders(1, 6, "URL");
            int i = 2;
            foreach (var c in cards) {
                var list = trello.Lists.WithId(c.IdList);
                excell_app.addData(i, 1, list.Name);
                excell_app.addData(i, 2, c.Name);
                excell_app.addData(i, 3, c.Desc);
                excell_app.addData(i, 4, c.Closed.ToString());
                excell_app.addData(i, 5, c.Due.HasValue ? c.Due.Value.ToString("dd/MM/yyyy"): string.Empty);
                excell_app.addData(i, 6, c.Url);
                ++i;

            }
            //excell_app.close();

        }

        private void btnGetToken_Click(object sender, EventArgs e)
        {
            
            var url = trello.GetAuthorizationUrl("Trello Export", Scope.ReadWrite);

            // open a web page browser
            ProcessStartInfo sInfo = new ProcessStartInfo(url.ToString());
            Process.Start(sInfo);
        }

        private void label4_Click(object sender, EventArgs e)
        {
            ProcessStartInfo sInfo = new ProcessStartInfo(label4.Text);
            Process.Start(sInfo);
        }


    }


    class CreateExcelDoc
    {
        private Excel.Application app = null;
        private Excel.Workbook workbook = null;
        private Excel.Worksheet worksheet = null;
        private Excel.Range workSheet_range = null;

        public CreateExcelDoc()
        {
            createDoc();
        }

        public void createDoc()
        {
            try
            {
                app = new Excel.Application();
                app.Visible = true;
                workbook = app.Workbooks.Add(1);
                worksheet = (Excel.Worksheet)workbook.Sheets[1];
            }
            catch (Exception e)
            {
                Console.Write("Error");
            }
            finally
            {
            }
        }

        public void close()
        {
            workbook.Save();
            workbook.Close();
        }

        public void createHeaders(int row, int col, string htext)
        {
            worksheet.Cells[row, col] = htext;
        }

        public void addData(int row, int col, string data)
        {
            worksheet.Cells[row, col] = data;
        }
    }
}
