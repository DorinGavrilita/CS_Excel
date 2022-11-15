using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace CS_Excel
{
   public partial class Form1: Form
   {
      public Form1()
      {
         InitializeComponent();
      }
      Excel.Application excelApp = null;
      Excel.Workbook wkbk = null;
      Excel.Worksheet sheet = null;

      private void button2_Click(object sender, EventArgs e)
      {
         try
         {

            excelApp = new Excel.Application();
            excelApp.Visible = true;

            wkbk = excelApp.Workbooks.Add();
            sheet = wkbk.Sheets.Add();
            sheet.Name = "Test Sheet";
            sheet.Cells[1, 1] = "Table 1 - Test info";
            for (int i = 0; i < dataGridView1.RowCount; i++)
               for (int j = 0; j < dataGridView1.ColumnCount; j++)
               {
                  if (dataGridView1.Rows[i].Cells[j].Value != null)
                     sheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
               }
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
               wkbk.SaveAs(saveFileDialog1.FileName);
            }
            if (printDialog1.ShowDialog() == DialogResult.OK)
            {
               wkbk.PrintOutEx(printDialog1.Tag);
            }

         }
         catch (Exception ex)
         {
            MessageBox.Show(ex.Message);
         }
         finally
         {
            excelApp.Quit();
            excelApp = null;
            wkbk = null;
            sheet = null;
         }
      }

      private void button3_Click(object sender, EventArgs e)
      {
         string[] alfa = new string[] { "a", "b", "c", "d", "e", "f", "g" };

         try
         {
            excelApp = new Excel.Application();
            // excelApp.Visible = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
               wkbk = excelApp.Workbooks.Open(openFileDialog1.FileName);
               sheet = excelApp.ActiveSheet;
               int row = Convert.ToInt32(textBox1.Text);
               int row1 = Convert.ToInt32(textBox3.Text);
               int col = Convert.ToInt32(textBox2.Text);
               int col1 = Convert.ToInt32(textBox4.Text);
               dataGridView1.Rows.Clear();
               // set the count of colums and rows that we need
               //.......
               dataGridView1.Rows.Add(row1 - row);
               Excel.Range rg = sheet.get_Range(alfa[col - 1] + row, alfa[col1 - 1] + row1);
               for (int i = 0; i < dataGridView1.RowCount; i++)
                  for (int j = 0; j < dataGridView1.ColumnCount; j++)
                  {
                     if (rg.Item[row + i, col + j].Value != null)
                        dataGridView1.Rows[i].Cells[j].Value = rg.Item[row + i, col + j].Value.ToString();
                  }
            }

         }
         catch (Exception ex)
         {
            MessageBox.Show(ex.Message);
         }
         finally
         {
            excelApp.Quit();
            excelApp = null;
            wkbk = null;
            sheet = null;
         }

      }
   }
}
