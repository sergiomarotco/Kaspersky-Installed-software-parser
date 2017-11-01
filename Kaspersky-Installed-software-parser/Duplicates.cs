using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace Kaspersky_Installed_software_parser
{
    public partial class Duplicates : Form
    {
        /// <summary>
        /// Список слов одновременно и белых и черных
        /// </summary>
        private List<string> DublicatesWords;
        private object[] o;
        /// <summary>
        /// Список новых белых слов
        /// </summary>
        private List<string> NewWhite;
        /// <summary>
        /// Список новых черных слов
        /// </summary>
        private List<string> NewBad;
        /// <summary>
        /// Путь к файлу с черным списком
        /// </summary>
        private string BoxBadText = String.Empty;
        /// <summary>
        /// Путь к файлу с белым списком
        /// </summary>
        private string BoxWhiteText = String.Empty;

        internal Duplicates(List<string> DublicatesWords, string BoxBadText, string BoxWhiteText)
        {
            InitializeComponent();
            NewWhite = new List<string>();
            NewBad = new List<string>();

            this.BoxBadText = BoxBadText;
            this.BoxWhiteText = BoxWhiteText;
            this.DublicatesWords = DublicatesWords;

            o = new object[DublicatesWords.Count];
            for (int b = 0; b < this.DublicatesWords.Count; b++)
            {
                o[b] = this.DublicatesWords[b];
            }
            listBox1.Items.AddRange(o);
        }

        private void Duplicates_Load(object sender, EventArgs e)
        {
            Icon = Properties.Resources.icon;
        }

        private void Duplicates_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (listBox1.Items.Count != 0)
            {
                DialogResult = DialogResult.Abort;
            }
            else { DialogResult = DialogResult.OK; }
        }

        private void Button5_Click(object sender, EventArgs e)
        {
            try
            {
                if (listBox1.SelectedItems.Count != 0)
                {
                    for (int l = listBox1.SelectedItems.Count - 1; l >= 0; l--)
                    {
                        NewBad.Add(listBox1.SelectedItems[l].ToString());
                        listBox1.Items.RemoveAt(listBox1.SelectedIndices[l]);
                    }
                    string[] White = File.ReadAllLines(BoxWhiteText, Encoding.UTF8);
                    for (int w = 0; w < White.Length; w++)
                    {
                        for (int b = 0; b < NewBad.Count; b++)
                        {
                            if (!String.IsNullOrEmpty(White[w]))
                            {
                                if (White[w].Equals(NewBad[b]))
                                {
                                    White[w] = "";
                                }
                            }
                        }
                    }
                    List<string> content = new List<string>();//исправленный список хороших слов с вырезанными одновременно плохими словами из этого списка
                    for (int w = 0; w < White.Length; w++)
                    {
                        if (!String.IsNullOrEmpty(White[w]))
                        {
                            content.Add(White[w]);
                        }
                    }
                    content = RemoveDuplicate(content);
                    File.WriteAllLines(BoxWhiteText, content.ToArray());
                    listBox1.SelectedIndex = 0;
                }
            }
            catch { }
            if (listBox1.Items.Count == 0)
                DialogResult = DialogResult.OK;
        }
        /// <summary>
        /// Удалить дубликаты в List
        /// </summary>
        /// <param name="sourceList">List в котором удалить дубликаты</param>
        /// <returns></returns>
        private List<string> RemoveDuplicate(List<string> sourceList = null)
        {
            List<string> list = new List<string>();
            try
            {
                foreach (string item in sourceList)
                    if (!list.Contains(item)) list.Add(item);
            }
            catch { }
            return list;
        }

        private void Button6_Click(object sender, EventArgs e)
        {
            try
            {
                if (listBox1.SelectedItems.Count != 0)
                {
                    for (int l = listBox1.SelectedItems.Count - 1; l >= 0; l--)
                    {
                        NewWhite.Add(listBox1.SelectedItems[l].ToString());
                        listBox1.Items.RemoveAt(listBox1.SelectedIndices[l]);
                    }
                    string[] Bad = File.ReadAllLines(BoxBadText, Encoding.UTF8);
                    for (int b = 0; b < Bad.Length; b++)
                    {
                        for (int w = 0; w < NewWhite.Count; w++)
                        {
                            if (!String.IsNullOrEmpty(Bad[b]))
                            {
                                if (Bad[b].Equals(NewWhite[w]))
                                {
                                    Bad[b] = "";
                                }
                            }
                        }
                    }
                    List<string> content = new List<string>();//исправленный список плохих слов с вырезанными одновременно хорошими словами из этого списка
                    for (int w = 0; w < Bad.Length; w++)
                    {
                        if (!String.IsNullOrEmpty(Bad[w]))
                        {
                            content.Add(Bad[w]);
                        }
                    }
                    content = RemoveDuplicate(content);
                    File.WriteAllLines(BoxBadText, content.ToArray());
                    listBox1.SelectedIndex = 0;
                }
            }
            catch { }
            if(listBox1.Items.Count==0)
                DialogResult = DialogResult.OK;
        }
    }
}
