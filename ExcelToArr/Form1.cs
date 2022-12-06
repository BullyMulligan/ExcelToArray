using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
namespace ExcelToArr
{
    public partial class Form1 : Form
    {
        private string[,] excelTable;
        private int _totalRows = 0;
        private int _totalColomns = 0;
        public Form1()
        {
            InitializeComponent();
        }

        private void открытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult res = openFileDialog1.ShowDialog();
                if (res == DialogResult.OK)
                {
                    ExcelPackage excelFile = new ExcelPackage(new FileInfo(openFileDialog1.FileName));
                    ExcelWorksheet worksheet = excelFile.Workbook.Worksheets[0];
                    _totalRows = worksheet.Dimension.End.Row;
                    _totalColomns = worksheet.Dimension.Columns;
                    excelTable = new string[_totalRows, _totalColomns];
                    for (int rowIndex = 1; rowIndex <= _totalRows; rowIndex++)
                    {
                        //Считанная строка с первой ячейки до конца. Усли значение ячейки null, то возвращаем string(Empty), в противном случае позвращаем значение ячейки приведенное в string
                        IEnumerable<string> row = worksheet.Cells[rowIndex, 1,rowIndex, _totalColomns].Select(c => c.Value == null ?string.Empty:c.Value.ToString());
                        List<string> list = row.ToList<string>();//добавляем строку к спуску строк
                        for (int i = 0; i < list.Count; i++)
                        {
                            excelTable[rowIndex - 1, i] = Convert.ToString(list[i].Replace('.', ','));
                        }
                    }

                    for (int i = 0; i < _totalRows; i++)
                    {
                        for (int j = 0; j < _totalColomns; j++)
                        {
                            richTextBox1.Text += Convert.ToString(excelTable[i, j]+"    ");
                        }

                        richTextBox1.Text += "\n";
                    }
                }
                else
                {
                    throw new Exception("Файл не выбран");
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}