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
        /// <summary>
        /// Перезаписать список хороших слов
        /// </summary>
        /// <returns></returns>
        private void Set_white()
        {
            try
            {
                FileInfo F = new FileInfo(BoxWhite.Text);
                if (F.Exists)
                {
                    WhatToConvertT = new List<string>();
                    Applications_white = File.ReadAllLines(BoxWhite.Text, Encoding.UTF8);
                    for (int i = 0; i < Applications_white.Length; i++)
                    {
                        if (!String.IsNullOrEmpty(Applications_white[i]))
                        {
                            WhatToConvertT.Add(Applications_white[i]);
                        }
                    }

                    File.WriteAllLines(BoxWhite.Text, ListToStringArray(WhatToConvertT));
                    WhatToConvertT.Clear();
                }
                else
                {
                    F.Create();
                    Applications_white = new string[0];
                }
            }
            catch { }
        }
        /// <summary>
        /// List<string> to string[]
        /// </summary>
        /// <param name="WhatToConvert">List string для конвертирования в string array</param>
        /// <returns></returns>
        private string[] ListToStringArray(List<string> WhatToConvert)
        {
            string[] output = new string[0];
            try
            {
                output = new string[WhatToConvert.Count];
                for (int i = 0; i < WhatToConvert.Count; i++)
                    output[i] = WhatToConvert[i];
            }
            catch {  }
            return output;
        }
        /// <summary>
        /// Перезаписать список плохих слов
        /// </summary>
        /// <returns></returns>
        private void Set_Bad()
        {
            try
            {
                FileInfo F = new FileInfo(BoxBad.Text);
                if (F.Exists)
                {
                    WhatToConvertT = new List<string>();
                    Applications_bad = File.ReadAllLines(BoxBad.Text, Encoding.UTF8);
                    for (int i = 0; i < Applications_bad.Length; i++)
                    {
                        if (!String.IsNullOrEmpty(Applications_bad[i]))
                        {
                            WhatToConvertT.Add(Applications_bad[i]);
                        }
                    }

                    File.WriteAllLines(BoxBad.Text, ListToStringArray(WhatToConvertT));
                    WhatToConvertT.Clear();
                }
                else
                {
                    F.Create();
                    Applications_bad = new string[0];
                }
            }
            catch { }
        }
        /// <summary>
        /// Определить DNS имя по IP
        /// </summary>
        /// <param name="fff"></param>
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
                        dnsname = Dns.GetHostEntry(str[2]).HostName;
                    }
                    catch (Exception ee)
                    { dnsname = ee.Message; }
                    Invoke((ThreadStart)delegate { dataGridView1.Rows[Convert.ToInt32(str[1])].Cells[3].Value = dnsname; iter++; label7.Text = iter.ToString() + " / " + dataGridView1.Rows.Count; label7.Refresh(); });
                }
            }
            catch { }
        }

        private List<Programms> P = new List<Programms>();
        /// <summary>
        /// Список хороших программ
        /// </summary>
        private string[] Applications_white;
        /// <summary>
        /// Список плохих программ
        /// </summary>
        private string[] Applications_bad;
        /// <summary>
        /// ПРостая функция проверки корректности введенного IP
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        private static bool CheckIp(string address)
        {
            try
            {
                string[] nums = address.Split('.');
                int useless;
                return nums.Length == 4 && nums.All(n => int.TryParse(n, out useless)) && nums.Select(int.Parse).All(n => n < 256);
            }
            catch
            { return false; }
        }

        private int iter = 0;

        private void Button1_Click(object sender, EventArgs e)
        {
            Set_white();
            Set_Bad();
            dataGridView1.Rows.Clear();
            dataGridView1.Refresh();
            OpenFileDialog openFileDialog1 = new OpenFileDialog()
            {
                Filter = "XLS files(*.xls)|*.xls|TXT files(*.txt)|*.txt",
                CheckFileExists = true,
                InitialDirectory = Directory.GetCurrentDirectory()
            };
            if (openFileDialog1.ShowDialog() == DialogResult.OK && openFileDialog1.FileName.Contains(".xls", StringComparison.OrdinalIgnoreCase))
            {
                try
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
                        if (array != null && Applications_bad != null && Applications_white != null)
                        {
                            P.Clear();
                            P = new List<Programms>();
                            for (int i = 9; i < rowNo; i++)
                            {
                                P.Add(new Programms(array[i, 1].ToString(), array[i, 3].ToString()));
                                bool next = false;
                                if (!String.IsNullOrEmpty(array[i, 1].ToString()))
                                {
                                    for (int j = 0; j < Applications_white.Length; j++)
                                    {
                                        if (!String.IsNullOrEmpty(Applications_white[j]) && array[i, 1] != null && array[i, 1].ToString().Contains(Applications_white[j], StringComparison.OrdinalIgnoreCase))
                                        {
                                            if (checkBox2.Checked)
                                                dataGridView1.Rows.Add(array[i, 1], array[i, 2], array[i, 3], "...", "White");
                                            array[i, 1] = "";//заменить на удаление строки
                                            next = true;
                                            break;
                                        }
                                    }
                                    if (!next)
                                    {
                                        for (int j = 0; j < Applications_bad.Length; j++)
                                        {
                                            if (!String.IsNullOrEmpty(Applications_bad[j]) && array[i, 1] != null && array[i, 1].ToString().Contains(Applications_bad[j], StringComparison.OrdinalIgnoreCase))
                                            {
                                                dataGridView1.Rows.Add(array[i, 1], array[i, 2], array[i, 3], "...", "Bad");
                                                array[i, 1] = "";//заменить на удаление строки
                                                next = true;
                                                break;
                                            }
                                        }
                                    }
                                    if (!next)
                                    {
                                        dataGridView1.Rows.Add(array[i, 1], array[i, 2], array[i, 3], "...", "Need request");
                                        array[i, 1] = "";
                                    }
                                }
                            }
                            if (checkBox1.Checked)
                            {
                                Thread t = new Thread(new ParameterizedThreadStart(Refresh_datagridview));
                                t.Start((object)(rowNo - 1));
                            }
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
                catch { }
            }
            dataGridView1.Enabled = false;
            bool isneed = false;
            for (int d = 0; d < dataGridView1.RowCount; d++)
            {
                DataGridViewRow row = dataGridView1.Rows[d];
                if (row.Cells[4].Value.ToString().Equals("Need request"))
                {
                    isneed = true;
                    break;
                }
            }
            if (isneed)
            {
                for (int d = 0; d < dataGridView1.RowCount; d++)
                {
                    if (!dataGridView1.CurrentRow.Cells[4].Value.Equals("Need request"))
                    {
                        dataGridView1.CurrentCell = dataGridView1[0, d];
                        dataGridView1.Refresh();
                    }
                    else
                    {
                        dataGridView1.Enabled = true;
                        break;
                    }
                }
            }
            dataGridView1.Enabled = true;
        }
        /// <summary>
        /// DNSlookup IP ресурсов
        /// </summary>
        /// <param name="rowNo"></param>
        private void Refresh_datagridview(object rowNo)
        {
            for (int j = 0; j < dataGridView1.Rows.Count; j++)
            {
                try
                {
                    Thread myThread = new Thread(new ParameterizedThreadStart(GetHostEntry));
                    string[] h = new string[3] { dataGridView1.Rows[j].Cells[2].Value.ToString(), j.ToString(), rowNo.ToString() };
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
            string[] content = new string[dataGridView1.SelectedRows.Count];
            if (dataGridView1.SelectedRows.Count == 0)
            {
                try
                {
                    for (int i = 0; i < dataGridView1.SelectedRows.Count; i++)
                    {
                        DataGridViewRow row = dataGridView1.SelectedRows[i];
                        content[i] = row.Cells[0].Value + "," + row.Cells[1].Value + "," + row.Cells[2].Value + "," + row.Cells[3].Value + "," + row.Cells[4].Value;
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
                        DataGridViewRow row = dataGridView1.Rows[i];
                        content[i] = row.Cells[0].Value + "," + row.Cells[1].Value + "," + row.Cells[2].Value + "," + row.Cells[3].Value + "," + row.Cells[4].Value;
                    }
                    catch { }
                }
            }
            DateTime t = DateTime.Now;
            string filename = t.Year + "." + t.Month + "." + t.Day + " " + t.Hour + "-" + t.Minute + "-" + t.Second + ".csv";
            File.WriteAllLines(filename, content, Encoding.UTF8);
            Process.Start(new ProcessStartInfo("explorer.exe", " /select, " + filename));
        }
        /// <summary>
        /// Добавить плохое слово
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ButtonBadADD_Click(object sender, EventArgs e)
        {
            string newsoft = "";
            if (dataGridView1.Rows.Count != 0)
            {
                DataGridViewRow row = dataGridView1.SelectedRows[0];
                newsoft = row.Cells[0].Value.ToString();//фиксируем текущую программу
                BadRichTextBox.Text += newsoft + Environment.NewLine;//Добавить новое белое слово
                dataGridView1.Rows[row.Index].Cells[4].Value = "Bad";//Помечаем текущую программу в списке белой

                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    try
                    {
                        row = dataGridView1.Rows[i];
                        if (row.Cells[0].Value.ToString().Equals(newsoft, StringComparison.OrdinalIgnoreCase))//если какая либо из строк является такой же программой
                        {
                            //BadRichTextBox.Text += row.Cells[0].Value + Environment.NewLine;
                            dataGridView1.Rows[i].Cells[4].Value = "Bad";//помечаем такую же программу белой
                        }
                    }
                    catch { }
                }
                BadRichTextBox.Lines = BadRichTextBox.Lines.Distinct().ToArray();//удаляем дубли
                dataGridView1.Enabled = false;
                for (int d = dataGridView1.SelectedRows[0].Index; d < dataGridView1.RowCount; d++)
                {
                    if (!dataGridView1.CurrentRow.Cells[4].Value.Equals("Need request"))
                    {
                        dataGridView1.CurrentCell = dataGridView1[0, d];
                        //dataGridView1.Refresh();
                    }
                    else
                    {
                        break;
                    }
                }
                dataGridView1.Enabled = true;//разблокируем
            }
        }
        /// <summary>
        /// Добавить хорошее слово
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ButtonWhiteADD_Click(object sender, EventArgs e)
        {
            string newsoft = "";
            if (dataGridView1.Rows.Count != 0)
            {
                DataGridViewRow row = dataGridView1.SelectedRows[0];
                newsoft = row.Cells[0].Value.ToString();//фиксируем текущую программу
                WhiteRichTextBox.Text += newsoft + Environment.NewLine;//Добавить новое белое слово
                dataGridView1.Rows[row.Index].Cells[4].Value = "White";//Помечаем текущую программу в списке белой                
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    try
                    {
                        row = dataGridView1.Rows[i];
                        if (row.Cells[0].Value.ToString().Equals(newsoft, StringComparison.OrdinalIgnoreCase))//если какая либо из строк является такой же программой
                        {
                            //WhiteRichTextBox.Text += row.Cells[0].Value + Environment.NewLine;
                            dataGridView1.Rows[i].Cells[4].Value = "White";//помечаем такую же программу белой
                        }
                    }
                    catch { }
                }
                WhiteRichTextBox.Lines = WhiteRichTextBox.Lines.Distinct().ToArray();//удаляем дубли
                dataGridView1.Enabled = false;
                for (int d = dataGridView1.SelectedRows[0].Index; d < dataGridView1.RowCount; d++)
                {
                    if (!dataGridView1.CurrentRow.Cells[4].Value.Equals("Need request"))
                    {
                        dataGridView1.CurrentCell = dataGridView1[0, d];
                        //dataGridView1.Refresh();
                    }
                    else
                    {
                        break;
                    }
                }
                dataGridView1.Enabled = true;//разблокируем
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

        private Duplicates D;
        private void Form1_Load(object sender, EventArgs e)
        {
            Icon = Properties.Resources.icon;
            Set_white();
            Set_Bad();

            List<string> DublicatesWords = new List<string>();
            // Applications_white;
            //  Applications_bad;
            for (int w = 0; w < Applications_white.Length; w++)
            {
                for (int b = 0; b < Applications_bad.Length; b++)
                {
                    if (Applications_white[w].Equals(Applications_bad[b]))
                    {
                        DublicatesWords.Add(Applications_white[w]);
                        break;
                    }
                }
            }
            if (DublicatesWords.Count != 0)
            {
                this.Hide();
                D = new Duplicates(DublicatesWords,BoxBad.Text,BoxWhite.Text)
                {
                    Owner = this
                };
                D.ShowDialog();
                Set_white();
                Set_Bad();
                this.Show();
            }
        }

        private void LinkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start("https://icons8.com");
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
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
