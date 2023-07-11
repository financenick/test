using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Tab;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;


namespace WF1
{
    public partial class Form1 : Form
    {
        private int totalDocuments;
        private int rkkDocuments;
        private int requestsDocuments;

        private string rkkFileName;
        private string requestsFileName;
        public Form1()
        {
            InitializeComponent();
        }

        private void browseRkkFileButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Текстовые файлы (*.txt)|*.txt";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                rkkFileName = openFileDialog.FileName;
                rkkFileTextBox.Text = rkkFileName;
            }
        }

        private void requestsFileTextBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void rkkFileTextBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void browseRequestsFileButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Текстовые файлы (*.txt)|*.txt";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                requestsFileName = openFileDialog.FileName;
                requestsFileTextBox.Text = requestsFileName;
            }
        }

        private Dictionary<string, int> CalculateItemCount(string fileName)
        {
            Dictionary<string, int> itemCounts = new Dictionary<string, int>();

            using (StreamReader reader = new StreamReader(fileName))
            {
                while (!reader.EndOfStream)
                {
                    string line = reader.ReadLine();
                    if (!string.IsNullOrEmpty(line))
                    {
                        string[] columns = line.Split('\t');
                        if (columns.Length >= 2)
                        {
                            string leader = columns[0];
                            string performers = columns[1];

                            string responsiblePerformer = (leader == "Климов Сергей Александрович") ? performers.Split(',')[0] : leader;

                            if (itemCounts.ContainsKey(responsiblePerformer))
                                itemCounts[responsiblePerformer]++;
                            else
                                itemCounts[responsiblePerformer] = 1;
                        }
                    }
                }
            }

            return itemCounts;
        }

        private Dictionary<string, int> CalculateTotalCounts(Dictionary<string, int> rkkCounts, Dictionary<string, int> requestsCounts)
        {
            Dictionary<string, int> totalCounts = new Dictionary<string, int>();

            foreach (string performer in rkkCounts.Keys)
            {
                int rkkCount = rkkCounts[performer];
                int requestsCount = requestsCounts.ContainsKey(performer) ? requestsCounts[performer] : 0;
                int totalCount = rkkCount + requestsCount;

                totalCounts[performer] = totalCount;
            }

            foreach (string performer in requestsCounts.Keys)
            {
                if (!totalCounts.ContainsKey(performer))
                {
                    int requestsCount = requestsCounts[performer];
                    totalCounts[performer] = requestsCount;
                }
            }

            return totalCounts;
        }

        private void calculateButton_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(rkkFileName) || string.IsNullOrEmpty(requestsFileName))
            {
                MessageBox.Show("Выберите файлы с данными.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                Dictionary<string, int> rkkCounts = CalculateItemCount(rkkFileName);
                Dictionary<string, int> requestsCounts = CalculateItemCount(requestsFileName);
                Dictionary<string, int> totalCounts = CalculateTotalCounts(rkkCounts, requestsCounts);

                //int totalDocuments = totalCounts.Values.Sum();
                //int rkkDocuments = rkkCounts.Values.Sum();
                //int requestsDocuments = requestsCounts.Values.Sum();

                totalDocuments = totalCounts.Values.Sum();
                rkkDocuments = rkkCounts.Values.Sum();
                requestsDocuments = requestsCounts.Values.Sum();

                MessageBox.Show($"Документов РКК: {rkkDocuments}\nДокументов обращений: {requestsDocuments}\nВсего документов: {totalDocuments}", "Результат", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // Создаем таблицу DataGridView с пятью колонками: номер, исполнитель, РКК, обращения, сумма
                DataGridView dataGridView = new DataGridView();
                dataGridView.Dock = DockStyle.Fill;
                dataGridView.AutoGenerateColumns = true;
                dataGridView.RowHeadersVisible = true;

                // Добавляем колонку с номерами
                dataGridView.Columns.Add(new DataGridViewTextBoxColumn()
                {
                    HeaderText = "№",
                    Width = 50,
                    Frozen = true,
                    ReadOnly = true,
                    AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                });

                // Добавляем колонку с именем исполнителя
                dataGridView.Columns.Add(new DataGridViewTextBoxColumn()
                {
                    Name = "performerColumn",
                    HeaderText = "Исполнитель"
                });

                // Добавляем колонку с данными по РКК
                DataGridViewTextBoxColumn rkkColumn = new DataGridViewTextBoxColumn();
                rkkColumn.DataPropertyName = "RKK";
                rkkColumn.HeaderText = "РКК";
                dataGridView.Columns.Add(rkkColumn);

                // Добавляем колонку с данными по обращениям
                DataGridViewTextBoxColumn requestsColumn = new DataGridViewTextBoxColumn();
                requestsColumn.DataPropertyName = "Requests";
                requestsColumn.HeaderText = "Обращения";
                dataGridView.Columns.Add(requestsColumn);

                // Добавляем колонку с суммой
                DataGridViewTextBoxColumn totalColumn = new DataGridViewTextBoxColumn();
                totalColumn.DataPropertyName = "Total";
                totalColumn.HeaderText = "Сумма";
                dataGridView.Columns.Add(totalColumn);

                // Заполняем таблицу данными
                int rowNumber = 1;
                foreach (string performer in rkkCounts.Keys)
                {
                    int rkkCount = rkkCounts[performer];
                    int requestsCount = requestsCounts.ContainsKey(performer) ? requestsCounts[performer] : 0;
                    int totalCount = totalCounts.ContainsKey(performer) ? totalCounts[performer] : 0;

                    dataGridView.Rows.Add(rowNumber, performer, rkkCount, requestsCount, totalCount);
                    rowNumber++;
                }

                // Фиксированная нумерация строк при сортировке
                dataGridView.CellFormatting += (cellSender, cellFormattingEventArgs) =>
                {
                    if (cellFormattingEventArgs.ColumnIndex == 0 && cellFormattingEventArgs.RowIndex >= 0)
                    {
                        cellFormattingEventArgs.Value = (cellFormattingEventArgs.RowIndex + 1).ToString();
                    }
                };

                // Создаем новую форму и отображаем таблицу на ней
                Form resultForm = new Form();
                resultForm.Width = 800;
                resultForm.Height = 800;
                resultForm.Text = "Результат";

                resultForm.Controls.Add(dataGridView);

                // Добавляем кнопку сохранения в файл
                Button saveButton = new Button();
                saveButton.Text = "Сохранить в файл";
                saveButton.Dock = DockStyle.Bottom;
                saveButton.Click += (s, ev) => SaveToWordFile(dataGridView);
                resultForm.Controls.Add(saveButton);

                resultForm.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при обработке файлов: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void SaveToWordFile(DataGridView dataGridView)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Документ Word (*.docx)|*.docx";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string fileName = saveFileDialog.FileName;

                using (WordprocessingDocument document = WordprocessingDocument.Create(fileName, WordprocessingDocumentType.Document))
                {
                    MainDocumentPart mainPart = document.AddMainDocumentPart();
                    mainPart.Document = new Document();
                    DocumentFormat.OpenXml.Wordprocessing.Body body = mainPart.Document.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Body());

                    // Добавление заголовка документа
                    Paragraph heading = new Paragraph(
                        new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
                        new Run(new DocumentFormat.OpenXml.Wordprocessing.Text("Справка о неисполненных документах и обращениях граждан"))
                    );
                    body.AppendChild(heading);

                    // Добавление информации о неисполненных документах и обращениях граждан
                    // Получение количества документов РКК
                    //int rkkDocuments = rkkCounts.Sum(pair => pair.Value);
                    //// Получение количества документов обращений
                    //int requestsDocuments = requestsCounts.Sum(pair => pair.Value);
                    //// Получение общего количества документов
                    //int totalDocuments = totalCounts.Sum(pair => pair.Value);

                    Paragraph info = new Paragraph(
                        new Run(new DocumentFormat.OpenXml.Wordprocessing.Text($"Не исполнено в срок {totalDocuments} документов, из них:")),
                        new Run(new DocumentFormat.OpenXml.Wordprocessing.Text($"\n- количество неисполненных входящих документов: {rkkDocuments}")),
                        new Run(new DocumentFormat.OpenXml.Wordprocessing.Text($"\n- количество неисполненных письменных обращений граждан: {requestsDocuments}"))
                    );
                    body.AppendChild(info);


                    // Создание таблицы в документе Word
                    Table table = new Table();

                    // Добавление стилей таблицы
                    TableProperties tableProperties = new TableProperties(
                        new TableBorders(
                            new TopBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 1 },
                            new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 1 },
                            new LeftBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 1 },
                            new RightBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 1 },
                            new InsideHorizontalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 1 },
                            new InsideVerticalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 1 }
                        )
                    );
                    table.AppendChild(tableProperties);

                    // Добавление заголовка таблицы
                    TableRow headerRow = new TableRow();

                    headerRow.AppendChild(new TableCell(new Paragraph(new Run(new DocumentFormat.OpenXml.Wordprocessing.Text("№"))))); // Добавление ячейки с заголовком номера строки

                    for (int i = 1; i < dataGridView.Columns.Count; i++)
                    {
                        DataGridViewColumn column = dataGridView.Columns[i];
                        TableCell headerCell = new TableCell(new Paragraph(new Run(new DocumentFormat.OpenXml.Wordprocessing.Text(column.HeaderText))));
                        headerRow.AppendChild(headerCell);
                    }

                    table.AppendChild(headerRow);

                    int rowNumber = 1; // Переменная счетчика для номеров строк

                    foreach (DataGridViewRow dataRow in dataGridView.Rows)
                    {
                        TableRow tableRow = new TableRow();
                        tableRow.AppendChild(new TableCell(new Paragraph(new Run(new DocumentFormat.OpenXml.Wordprocessing.Text(rowNumber.ToString()))))); // Добавление ячейки с номером строки

                        for (int i = 0; i < dataRow.Cells.Count; i++)
                        {
                            // Пропустить второй столбец
                            if (i != 0)
                            {
                                string cellValue = dataRow.Cells[i].Value?.ToString() ?? string.Empty;

                                TableCell tableCell = new TableCell(
                                    new Paragraph(
                                        new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
                                        new Run(
                                            new RunProperties(new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = "14" }),
                                            new DocumentFormat.OpenXml.Wordprocessing.Text(cellValue)
                                        )
                                    )
                                );
                                tableRow.AppendChild(tableCell);
                            }
                        }

                        table.AppendChild(tableRow);
                        rowNumber++; // Увеличение счетчика номеров строк
                    }

                    body.AppendChild(table);
                }

                MessageBox.Show("Файл успешно сохранен.", "Сохранение", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }


    }
}
