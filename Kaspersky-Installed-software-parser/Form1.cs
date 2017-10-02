using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Kaspersky_Installed_software_parser
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            Set_white();
            Set_Bad();
        }
        private List<string> WhatToConvertT;
        private string Set_white()
        {
            FileInfo F = new FileInfo(BoxWhite.Text);
            if (F.Exists)
            {
                WhatToConvertT = new List<string>();
                Applications_white = File.ReadAllLines(BoxWhite.Text, Encoding.UTF8);
                for (int i = 0; i < Applications_white.Length; i++)
                {
                    if (Applications_white[i] != "")
                    {
                        WhatToConvertT.Add(Applications_white[i]);
                    }
                }

                File.WriteAllLines(BoxWhite.Text, ListToStringArray(WhatToConvertT));
                WhatToConvertT.Clear();
                return F.FullName;
            }
            else
            {
                return "";
            }
        }

        private string[] ListToStringArray(List<string> WhatToConvert)
        {
            string[] output = new string[0];
            try
            {
                output = new string[WhatToConvert.Count];
                for (int i = 0; i < WhatToConvert.Count; i++)
                    output[i] = WhatToConvert[i].ToString();
            }
            catch {  }
            return output;
        }
        private string Set_Bad()
        {
            FileInfo F = new FileInfo(BoxBad.Text);
            if (F.Exists)
            {
                WhatToConvertT = new List<string>();
                Applications_bad = File.ReadAllLines(BoxBad.Text, Encoding.UTF8);
                for (int i = 0; i < Applications_bad.Length; i++)
                {
                    if (Applications_bad[i] != "")
                    {
                        WhatToConvertT.Add(Applications_bad[i]);
                    }
                }

                File.WriteAllLines(BoxBad.Text, ListToStringArray(WhatToConvertT));
                WhatToConvertT.Clear();
                return F.FullName;
            }
            else
            {
                return "";
            }
        }

        private void GetHostEntry(object fff)
        {
            try
            {
                string[] str = Array.ConvertAll<object, string>((object[])fff, Convert.ToString);

                if (CheckIp(str[0]))
                {
                    string dnsname = "";
                    try
                    {
                        dnsname = Dns.GetHostEntry(str[0]).HostName;
                    }
                    catch (Exception ee) { dnsname = ee.Message; }
                    Invoke((ThreadStart)delegate { dataGridView1.Rows[Convert.ToInt32(str[1])].Cells[2].Value = dnsname; iter++; label7.Text = iter.ToString() + " / " + dataGridView1.Rows.Count; label7.Refresh(); });
                }
            }
            catch { }
        }

        private List<Programms> P = new List<Programms>();
        private string[] Applications_white;
        private string[] Applications_bad;
        private static bool CheckIp(string address)
        {
            try
            {
                string[] nums = address.Split('.');
                int useless;
                return nums.Length == 4 && nums.All(n => int.TryParse(n, out useless)) && nums.Select(int.Parse).All(n => n < 256);
            }
            catch { return false; }
        }

        private int iter = 0;
        private void Button1_Click(object sender, EventArgs e)
        {
            Set_white();
            Set_Bad();
            OpenFileDialog openFileDialog1 = new OpenFileDialog()
            {
                Filter = "XLS files(*.xls)|*.xls|TXT files(*.txt)|*.txt",
                CheckFileExists = true,
                InitialDirectory = Directory.GetCurrentDirectory()
            };
            if (openFileDialog1.ShowDialog() == DialogResult.OK && openFileDialog1.FileName.Contains(".xls", StringComparison.OrdinalIgnoreCase))
            {
                BoxApplications.Text = openFileDialog1.FileName;
                Excel.Worksheet Sheet = new Excel.Worksheet(); //какой-то рабочий лист Excel
                Excel.Application Excel_App = new Excel.Application(); //создаём приложение Excel
                Excel.Workbook Book = Excel_App.Workbooks.Open(openFileDialog1.FileName); //открываем наш файл (Рабочую книгу Excel) //xlApp.Visible = true; //сделать Excel видимым, но это не обязательно
                foreach (Excel.Worksheet worksheet in Book.Worksheets)
                {
                    Sheet = Book.Worksheets[worksheet.Name]; //присваиваем переменной iSht Лист1 или так xlSht = xlWB.ActiveSheet //активный лист

                    int rowNo = Sheet.UsedRange.Rows.Count;
                    object[,] array = Sheet.UsedRange.Value;
                    if (array != null && Applications_bad != null && Applications_white != null && Applications_bad.Length != 0 && Applications_white.Length != 0)
                    {
                        white.Tables.Add();
                        P.Clear();
                        P = new List<Programms>();
                        for (int i = 9; i < rowNo; i++)
                        {
                            P.Add(new Programms(array[i, 1].ToString(), array[i, 3].ToString()));
                            bool next = false;
                            if (!array[i, 1].Equals(""))
                            {
                                for (int j = 0; j < Applications_white.Length; j++)
                                {
                                    if (!Applications_white[j].Equals("") && array[i, 1] != null && array[i, 1].ToString().Contains(Applications_white[j], StringComparison.OrdinalIgnoreCase))
                                    {
                                        dataGridView1.Rows.Add(array[i, 1], array[i, 3], "...", "White");
                                        array[i, 1] = "";//заменить на удаление строки
                                        next = true;
                                        break;
                                    }
                                }
                                if (!next)
                                {
                                    for (int j = 0; j < Applications_bad.Length; j++)
                                    {
                                        if (!Applications_bad[j].Equals("") && array[i, 1] != null && array[i, 1].ToString().Contains(Applications_bad[j], StringComparison.OrdinalIgnoreCase))
                                        {
                                            dataGridView1.Rows.Add(array[i, 1], array[i, 3], "...", "Bad");
                                            array[i, 1] = "";//заменить на удаление строки
                                            next = true;
                                            break;
                                        }
                                    }
                                }
                                if (!next)
                                {
                                    dataGridView1.Rows.Add(array[i, 1], array[i, 3], "...", "Need request");
                                    array[i, 1] = "";
                                }
                            }
                        }

                        Thread rfgfsd = new Thread(new ParameterizedThreadStart(Refresh_datagridview));
                        rfgfsd.Start((object)(rowNo - 1));
                    }
                    else
                    {
                        if (Applications_bad == null)
                        {
                            label7.Text = "Bad applications = null";
                        }
                        if (Applications_white == null)
                        {
                            label7.Text = "White applications = null";
                        }
                    }
                }
                dataGridView1.Refresh();
                Process[] pProcess = Process.GetProcessesByName("Excel");// cleanup:
                for (int i = pProcess.Length; i >= 0; i--)
                {
                    try { pProcess[0].Kill(); } catch { }
                }
                if (Sheet != null)
                {
                    Marshal.FinalReleaseComObject(Sheet);
                    Sheet = null;
                }
                if (Book != null)
                {
                    Marshal.FinalReleaseComObject(Book);
                    Book = null;
                }
                if (Excel_App != null)
                {
                    Marshal.FinalReleaseComObject(Excel_App);
                    Excel_App = null;
                }
            }
        }

        private void Refresh_datagridview(object rowNo)
        {
            for (int j = 0; j < dataGridView1.Rows.Count; j++)
            {
                try
                {
                    Thread myThread = new Thread(new ParameterizedThreadStart(GetHostEntry));
                    string[] h = new string[3] { dataGridView1.Rows[j].Cells[1].Value.ToString(), j.ToString(), rowNo.ToString() };
                    myThread.Start((object)h);
                }
                catch { }
            }
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            openFileDialog2.InitialDirectory = Directory.GetCurrentDirectory();
            if (openFileDialog2.ShowDialog() == DialogResult.OK)
            {
                if (openFileDialog2.FileName.Contains(".txt", StringComparison.OrdinalIgnoreCase))
                {
                    BoxWhite.Text = openFileDialog2.FileName;
                    Set_white();
                }
            }
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            openFileDialog2.InitialDirectory = Directory.GetCurrentDirectory();
            if (openFileDialog2.ShowDialog() == DialogResult.OK)
            {
                if (openFileDialog2.FileName.Contains(".txt", StringComparison.OrdinalIgnoreCase))
                {
                    BoxBad.Text = openFileDialog2.FileName;
                    Set_Bad();
                }
            }
        }

        private void Button4_Click(object sender, EventArgs e)
        {
            var t = DateTime.Now;
            string file = t.Year + "." + t.Month + "." + t.Day + " " + t.Hour + "-" + t.Minute + "-" + t.Second+".csv";
            string[] content = new string[dataGridView1.SelectedRows.Count];
            if (dataGridView1.SelectedRows.Count == 0)
            {
                try
                {
                    for (int i = 0; i < dataGridView1.SelectedRows.Count; i++)
                    {
                        var row = dataGridView1.SelectedRows[i];
                        content[i] = row.Cells[0].Value + "," + row.Cells[1].Value + "," + row.Cells[2].Value + "," + row.Cells[3].Value;
                    }
                }
                catch { }
            }
            else
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    try
                    {
                        var row = dataGridView1.Rows[i];
                        content[i] = row.Cells[0].Value + "," + row.Cells[1].Value + "," + row.Cells[2].Value + "," + row.Cells[3].Value;
                    }
                    catch { }
                }
            }
            File.WriteAllLines(file, content, Encoding.UTF8);
            Process.Start(new ProcessStartInfo("explorer.exe", " /select, " + file));
        }

        private void ButtonBadADD_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count != 0)
            {
                string newsoft = "";
                for (int i = 0; i < dataGridView1.SelectedRows.Count; i++)
                {
                    try
                    {
                        var row = dataGridView1.SelectedRows[i];
                        BadRichTextBox.Text += row.Cells[0].Value + Environment.NewLine;
                        dataGridView1.Rows[i].Cells[3].Value = "Bad";
                        newsoft = row.Cells[0].Value.ToString();
                        break;
                    }
                    catch { }
                }
                dataGridView1.Refresh();
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    try
                    {
                        var row = dataGridView1.Rows[i];
                        if (row.Cells[0].Value.ToString().Equals(newsoft, StringComparison.OrdinalIgnoreCase))
                        {
                            BadRichTextBox.Text += dataGridView1.Rows[i].Cells[0].Value + Environment.NewLine;
                            dataGridView1.Rows[i].Cells[3].Value = "Bad";
                        }
                    }
                    catch { }
                }
                BadRichTextBox.Lines = BadRichTextBox.Lines.Distinct().ToArray();
                dataGridView1.CurrentCell = dataGridView1[0, dataGridView1.SelectedRows[0].Index + 1];
            }
        }

        private void ButtonWhiteADD_Click(object sender, EventArgs e)
        {
            string newsoft = "";
            if (dataGridView1.Rows.Count != 0)
            {
                for (int i = 0; i < dataGridView1.SelectedRows.Count; i++)
                {
                    try
                    {
                        var row = dataGridView1.SelectedRows[i];
                        WhiteRichTextBox.Text += row.Cells[0].Value + Environment.NewLine;
                        dataGridView1.Rows[i].Cells[3].Value = "White";
                        newsoft = row.Cells[0].Value.ToString();
                        break;
                    }
                    catch { }
                }
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    try
                    {
                        var row = dataGridView1.Rows[i];
                        if (row.Cells[0].Value.ToString().Equals(newsoft, StringComparison.OrdinalIgnoreCase))
                        {
                            WhiteRichTextBox.Text += row.Cells[0].Value + Environment.NewLine;
                            dataGridView1.Rows[i].Cells[3].Value = "White";
                        }
                    }
                    catch { }
                }
                WhiteRichTextBox.Lines = WhiteRichTextBox.Lines.Distinct().ToArray();
                dataGridView1.CurrentCell = dataGridView1[0, dataGridView1.SelectedRows[0].Index + 1];
            }
        }

        private void Button7_Click(object sender, EventArgs e)
        {
            if (WhiteRichTextBox.Lines.Length != 0)
            {
                File.AppendAllLines(BoxWhite.Text, WhiteRichTextBox.Lines, Encoding.UTF8);
                WhiteRichTextBox.Clear();
            }
            if (BadRichTextBox.Lines.Length != 0)
            {
                File.AppendAllLines(BoxBad.Text, BadRichTextBox.Lines, Encoding.UTF8);
                BadRichTextBox.Clear();
            }
        }

        private void LinkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start("mailto:sergiomarotco@gmail.com");
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Icon = Properties.Resources.icon;
        }

        private void LinkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start("https://icons8.com");
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                string fff = BadRichTextBox.Text;
                if (BadRichTextBox.Lines.Length != 0 || WhiteRichTextBox.Lines.Length != 0)
                {
                    string gh = string.Empty;
                    gh += "new White applications: " + WhiteRichTextBox.Lines.Length + Environment.NewLine;
                    gh += "new Bad applications: " + BadRichTextBox.Lines.Length;
                    DialogResult dialogResult = MessageBox.Show(gh, "Save changes?", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes)
                    {
                        if (WhiteRichTextBox.Lines.Length != 0)
                        {
                            File.AppendAllLines(BoxWhite.Text, WhiteRichTextBox.Lines, Encoding.UTF8);
                            WhiteRichTextBox.Clear();
                        }
                        if (BadRichTextBox.Lines.Length != 0)
                        {
                            File.AppendAllLines(BoxBad.Text, BadRichTextBox.Lines, Encoding.UTF8);
                            BadRichTextBox.Clear();
                        }
                    }
                }
                Process[] pProcess = Process.GetProcessesByName("Excel");// cleanup:
                for (int i = pProcess.Length; i >= 0; i--)
                {
                    try { pProcess[0].Kill(); } catch { }
                }
            }
            catch { }
        }

        private void DataGridView1_DoubleClick(object sender, EventArgs e)
        {
            Process.Start("https://www.google.ru/search?q=" + dataGridView1.SelectedCells[0].Value);
        }
    }

    internal class Programms
    {
        private readonly string IP;
        private readonly string SoftwareName;
        public Programms(string programm, string IP)
        {
            SoftwareName = programm;
            this.IP = IP;
        }
    }

    internal static class StringExtensions
    {
        public static bool Contains(this string source, string toCheck, StringComparison comp)
        {
            return source.IndexOf(toCheck, comp) >= 0;
        }
    }
}
