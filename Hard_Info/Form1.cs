using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Management;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Hard_Info
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void toolStripComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string key = string.Empty;
            switch (comboBox1.Text)
            {
                case "Материнская плата":
                    key = "Win32_BaseBoard";
                    break;
                case "Процессор":
                    key = "Win32_Processor";
                    break;
                case "Видеокарта":
                    key = "Win32_VideoController";
                    break;
                case "Чипсет":
                    key = "Win32_IDEController";
                    break;
                case "Батарея":
                    key = "Win32_Battery";
                    break;
                case "Биос":
                    key = "Win32_BIOS";
                    break;
                case "Оперативная память":
                    key = "Win32_PhysicalMemory";
                    break;
                case "Кэш":
                    key = "Win32_CacheMemory";
                    break;
                case "USB":
                    key = "Win32_USBController";
                    break;
                case "Диск":
                    key = "Win32_DiskDrive";
                    break;
                case "Логические диски":
                    key = "Win32_LogicalDisk";
                    break;
                case "Клавиатура":
                    key = "Win32_Keyboard";
                    break;
                case "Сеть":
                    key = "Win32_NetworkAdapter";
                    break;
                case "Пользователи":
                    key = "Win32_Account";
                    break;
            }
            GetInfo(key, listView1);
        }

        private void GetInfo(string key, ListView list)
        {
            list.Items.Clear();
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM " + key);
            try
            {
                foreach (ManagementObject obj in searcher.Get())
                {
                    ListViewGroup listViewGroup;
                    try
                    {
                        listViewGroup = list.Groups.Add(obj["Name"].ToString(), obj["Name"].ToString());
                    }
                    catch
                    {
                        listViewGroup = list.Groups.Add(obj.ToString(), obj.ToString());
                    }
                    if (obj.Properties.Count == 0)
                    {
                        MessageBox.Show("Не удалось получить информацию", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    foreach (PropertyData data in obj.Properties)
                    {
                        ListViewItem listItem = new ListViewItem(listViewGroup);
                        if (list.Items.Count % 2 != 0)
                            listItem.BackColor = Color.White;
                        else 
                            listItem.BackColor = Color.WhiteSmoke;
                        listItem.Text = data.Name;
                        if (data.Value != null && !string.IsNullOrEmpty(data.Value.ToString()))
                        {
                            switch (data.Value.GetType().ToString())
                            {
                                case "System.Sting[]":
                                    string[] strings = data.Value as string[];
                                    string strings2 = string.Empty;
                                    foreach (string s in strings)
                                    {
                                        strings2 += $"{s} ";
                                    }
                                    listItem.SubItems.Add(strings2);
                                    break;
                                case "System.UInt16[]":
                                    ushort[] ushorts = data.Value as ushort[];
                                    string str3 = string.Empty;
                                    foreach (ushort u in ushorts)
                                    {
                                        str3 += $"{u} ";
                                    }
                                    listItem.SubItems.Add(str3);
                                    break;
                                default:
                                    listItem.SubItems.Add(data.Value.ToString());
                                    break;
                            }
                            list.Items.Add(listItem);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка1", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            comboBox1.SelectedIndex = 0;
            comboBox1.DropDownStyle = ComboBoxStyle.DropDownList;
        }

        private void listView1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.C) // Если нажато Ctrl+C
            {
                // Логика для извлечения текста из ListView и помещения в буфер обмена
                string selectedText = $"{listView1.SelectedItems[0].SubItems[0].Text} {listView1.SelectedItems[0].SubItems[1].Text}";
                Clipboard.SetText(selectedText);
            }
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            new choice(listView1, comboBox1.Text).ShowDialog();
        }

        
    }
}
