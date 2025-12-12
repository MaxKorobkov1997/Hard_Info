using System;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using ListView = System.Windows.Forms.ListView;

namespace Hard_Info
{
    public partial class choice : Form
    {
        private ListView listView1;
        private string name;
        public choice(ListView listView, string names)
        {
            InitializeComponent();
            listView1 = listView;
            name = names;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ExportToExcel(listView1);
        }

        private void ExportToExcel(ListView listView)
        {
            try
            {
                int x = 3;
                Excel.Application exApp = new Excel.Application();
                exApp.Workbooks.Add();
                Excel.Worksheet ws = (Excel.Worksheet)exApp.ActiveSheet;
                ws.Cells[1, 1].Value = name;
                ws.Cells[2, 1].Value = listView.Columns[0].Text;
                ws.Cells[2, 2].Value = listView.Columns[1].Text;
                for (int i = 0; i < listView.Groups.Count; i++)
                {
                    ws.Cells[x, 1].Value = listView.Groups[i].ToString();
                    x++;
                    for (int j = 0; j < listView.Groups[i].Items.Count; j++)
                    {
                        try
                        {
                            ws.Cells[x, 1].Value = listView.Groups[i].Items[j].SubItems[0].Text;
                            ws.Cells[x, 2].Value = listView.Groups[i].Items[j].SubItems[1].Text;
                        }
                        catch { }
                        x++;
                    }
                    x++;
                }
                exApp.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка1", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            using (var tw = new StreamWriter($"{name}.txt"))
                for (int i = 0; i < listView1.Groups.Count; i++)
                {
                    tw.WriteLine(listView1.Groups[i].ToString());
                    for (int j = 0; j < listView1.Groups[i].Items.Count; j++)
                        try
                        {
                            tw.WriteLine(listView1.Groups[i].Items[j].SubItems[0].Text + " " + listView1.Groups[i].Items[j].SubItems[1].Text);
                        }
                        catch { }
                }
        }
            
    }
}
