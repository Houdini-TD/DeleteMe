using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Windows.Forms;

namespace laba_10
{
    public partial class Form1 : Form
    {
        public int currentID = 1;

        public Form1()
        {
            InitializeComponent();
            dataGridView1.Columns.Add("ID", "ID");
            dataGridView1.Columns.Add("Product", "Товар");
            dataGridView1.Columns.Add("Quantity", "Количество");
            dataGridView1.Columns.Add("Price", "Цена");
            dataGridView1.Columns.Add("Total", "Сумма");

            dataGridView1.Columns["ID"].ReadOnly = true;
            dataGridView1.Columns["Total"].ReadOnly = true;

            dataGridView1.Columns["ID"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridView1.Columns["ID"].DefaultCellStyle.BackColor = System.Drawing.Color.LightGray;
            dataGridView1.Columns["Total"].DefaultCellStyle.BackColor = System.Drawing.Color.LightGray;

            dataGridView1.Rows.Add(1, "", "", "", "");

            dataGridView1.CellValueChanged += dataGridView1_CellValueChanged;
        }


        private void addRow_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Add(++currentID, "", "", "", "");
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && (e.ColumnIndex == 2 || e.ColumnIndex == 3))
            {
                string priceStr = dataGridView1.Rows[e.RowIndex].Cells[2].Value?.ToString();
                string quantityStr = dataGridView1.Rows[e.RowIndex].Cells[3].Value?.ToString();

                if (!string.IsNullOrEmpty(priceStr) && !string.IsNullOrEmpty(quantityStr))
                {
                    bool isPriceValid = int.TryParse(priceStr, out int price);
                    bool isQuantityValid = int.TryParse(quantityStr, out int quantity);

                    if (isPriceValid && isQuantityValid)
                    {
                        int total = price * quantity;
                        dataGridView1.Rows[e.RowIndex].Cells[4].Value = total;

                        changePrice();
                    }
                    else
                    {

                    }
                }
            }
        }

        private void changePrice()
        {
            int totalDocumentPrice = 0;

            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                totalDocumentPrice += int.Parse(dataGridView1.Rows[i].Cells[4].Value.ToString());

            res.Text = $"{totalDocumentPrice}, руб";
        }


        private void toExcel(object sender, EventArgs e)
        {
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = true;

            Excel.Workbook workbook = excelApp.Workbooks.Add();
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets[1];

            DateTime selectDate = dateTimePicker2.Value;

            if (selectDate != null)
            {
                string orderNumberText = order_number.Text;
                string formattedDate = selectDate.ToString("dd.MM.yyyy");

                string fullText = $"Расходная накладная № {orderNumberText} от {formattedDate}";
                worksheet.Cells[1, 1] = fullText;


                int orderNumberStart = fullText.IndexOf(orderNumberText);
                Excel.Range boldUnderlineRangeOrderNumber = (Excel.Range)worksheet.Cells[1, 1];
                boldUnderlineRangeOrderNumber.Characters[orderNumberStart + 1, orderNumberText.Length].Font.Bold = true;
                boldUnderlineRangeOrderNumber.Characters[orderNumberStart + 1, orderNumberText.Length].Font.Underline = true;

                worksheet.Cells[1, 1].Font.Name = "Times new roman";
                worksheet.Cells[1, 1].Font.Size = 14;
            }

            string PurchasertxtBox = provider_textBox.Text;
            worksheet.Cells[2, 1] = "Поставщик: ";


            Excel.Range boldUnderlineRange = (Excel.Range)worksheet.Cells[2, 1];
            boldUnderlineRange.Characters[boldUnderlineRange.Text.Length + 1, PurchasertxtBox.Length].Text = PurchasertxtBox;
            boldUnderlineRange.Characters[boldUnderlineRange.Text.Length - PurchasertxtBox.Length + 1, PurchasertxtBox.Length].Font.Bold = true;
            boldUnderlineRange.Characters[boldUnderlineRange.Text.Length - PurchasertxtBox.Length + 1, PurchasertxtBox.Length].Font.Underline = true;

            worksheet.Cells[2, 1].Font.Name = "Times new roman";
            worksheet.Cells[2, 1].Font.Size = 14;

            string ProvidertxtBox = buyer_textBox.Text;
            worksheet.Cells[3, 1] = "Покупатель: ";

            Excel.Range boldUnderlineRange1 = (Excel.Range)worksheet.Cells[3, 1];
            boldUnderlineRange1.Characters[boldUnderlineRange1.Text.Length + 1, ProvidertxtBox.Length].Text = ProvidertxtBox;
            boldUnderlineRange1.Characters[boldUnderlineRange1.Text.Length - ProvidertxtBox.Length + 1, ProvidertxtBox.Length].Font.Bold = true;
            boldUnderlineRange1.Characters[boldUnderlineRange1.Text.Length - ProvidertxtBox.Length + 1, ProvidertxtBox.Length].Font.Underline = true;

            worksheet.Cells[3, 1].Font.Name = "Times new roman";
            worksheet.Cells[3, 1].Font.Size = 14;

            int nRows = dataGridView1.Rows.Count;

            for (int i = 1; i <= 5; i++)
            {
                worksheet.Cells[4, i] = dataGridView1.Columns[i - 1].HeaderText;
            }

            for (int i = 2; i <= nRows; i++)
            {
                for (int j = 1; j <= 5; j++)
                {
                    var cellValue = dataGridView1.Rows[i - 2].Cells[j - 1].Value;

                    if (cellValue != null)
                    {
                        worksheet.Cells[i + 2, j] = cellValue.ToString();
                        worksheet.Cells[i + 2, j].Font.Bold = true;
                        worksheet.Cells[i + 2, j].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    }
                    else
                    {
                        worksheet.Cells[i + 2, j] = "N/A";
                    }
                }
            }


            Excel.Range totalCell = (Excel.Range)worksheet.Cells[nRows + 4, 1];
            totalCell.Value = $"Итого: {res.Text}";
            totalCell.Font.Name = "Times new roman";
            totalCell.Font.Size = 14;
            totalCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Документ Excel (*.xlsx)|*.xlsx";
            saveFileDialog.Title = "Сохранить файл Excel";
            saveFileDialog.FileName = "excelExample"; // Имя файла по умолчанию

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                workbook.SaveAs(saveFileDialog.FileName);
                workbook.Close();
                excelApp.Quit();
            }
        }
    }
}
