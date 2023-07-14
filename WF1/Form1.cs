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

        private Dictionary<string, int[]> CalculateDocumentCounts(string FileName, bool isRKK)
        {
            Dictionary<string, int[]> documentCounts = new Dictionary<string, int[]>();
            

            // Обработка файла РКК
            ProcessFile(FileName, documentCounts, true);

            // Обработка файла обращений
            ProcessFile(FileName, documentCounts, false);

            return documentCounts;
        }

        private void ProcessFile(string fileName, Dictionary<string, int[]> documentCounts, bool isRKK)
        {
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
                            string leader = columns[0].Trim();
                            string performers = columns[1].Trim();

                            // Приведение формата имени лидера
                            string[] leaderParts = leader.Split(' ');
                            string leaderFormatted = leaderParts[0] + " " + leaderParts[1][0] + "." + leaderParts[2][0] + ".";

                            // Определение ответственного исполнителя
                            string responsiblePerformer = (leaderFormatted.Equals("Климов С.А.")) ? performers.Split(';')[0].Trim() : leaderFormatted;

                           
                            responsiblePerformer = responsiblePerformer.Replace(" (Отв.)", "");

                            // Обновление счетчиков
                            if (documentCounts.ContainsKey(responsiblePerformer))
                            {
                                if (isRKK)
                                    documentCounts[responsiblePerformer][0]++; // РКК
                                else
                                    documentCounts[responsiblePerformer][1]++; // Обращения

                                documentCounts[responsiblePerformer][2]++; // Суммарно документов
                            }
                            else
                            {
                                documentCounts[responsiblePerformer] = new int[] { isRKK ? 1 : 0, isRKK ? 0 : 1, 1 };
                            }
                        }
                    }
                }
            }
        }

        private Dictionary<string, int> CalculateTotalCounts(Dictionary<string, int[]> rkkCounts, Dictionary<string, int[]> requestsCounts)
        {
            Dictionary<string, int> totalCounts = new Dictionary<string, int>();

            foreach (string performer in rkkCounts.Keys)
            {
                int rkkCount = rkkCounts[performer][0];
                int requestsCount = requestsCounts.ContainsKey(performer) ? requestsCounts[performer][1] : 0;
                int totalCount = rkkCount + requestsCount;

                totalCounts[performer] = totalCount;
            }

            foreach (string performer in requestsCounts.Keys)
            {
                if (!totalCounts.ContainsKey(performer))
                {
                    int requestsCount = requestsCounts[performer][1];
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
                Dictionary<string, int[]> rkkCounts = CalculateDocumentCounts(rkkFileName, true);
                Dictionary<string, int[]> requestsCounts = CalculateDocumentCounts(requestsFileName, false);
                Dictionary<string, int> totalCounts = CalculateTotalCounts(rkkCounts, requestsCounts);

                totalDocuments = totalCounts.Values.Sum();
                rkkDocuments = rkkCounts.Values.Sum(x => x[0]);
                requestsDocuments = requestsCounts.Values.Sum(x => x[1]);

                MessageBox.Show($"Документов РКК: {rkkDocuments}\nДокументов обращений: {requestsDocuments}\nВсего документов: {totalDocuments}", "Результат", MessageBoxButtons.OK, MessageBoxIcon.Information);

                
                DataGridView dataGridView = new DataGridView();
                dataGridView.Dock = DockStyle.Fill;
                dataGridView.AutoGenerateColumns = true;
                dataGridView.RowHeadersVisible = true;

                
                dataGridView.Columns.Add(new DataGridViewTextBoxColumn()
                {
                    HeaderText = "№",
                    Width = 50,
                    Frozen = true,
                    ReadOnly = true,
                    AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                });

             
                dataGridView.Columns.Add(new DataGridViewTextBoxColumn()
                {
                    Name = "performerColumn",
                    HeaderText = "Исполнитель"
                });

               
                DataGridViewTextBoxColumn rkkColumn = new DataGridViewTextBoxColumn();
                rkkColumn.DataPropertyName = "RKK";
                rkkColumn.HeaderText = "РКК";
                dataGridView.Columns.Add(rkkColumn);

                
                DataGridViewTextBoxColumn requestsColumn = new DataGridViewTextBoxColumn();
                requestsColumn.DataPropertyName = "Requests";
                requestsColumn.HeaderText = "Обращения";
                dataGridView.Columns.Add(requestsColumn);

              
                DataGridViewTextBoxColumn totalColumn = new DataGridViewTextBoxColumn();
                totalColumn.DataPropertyName = "Total";
                totalColumn.HeaderText = "Сумма";
                dataGridView.Columns.Add(totalColumn);

             
                int rowNumber = 1;
                foreach (string performer in totalCounts.Keys)
                {
                    int rkkCount = rkkCounts.ContainsKey(performer) ? rkkCounts[performer][0] : 0;
                    int requestsCount = requestsCounts.ContainsKey(performer) ? requestsCounts[performer][1] : 0;
                    int totalCount = totalCounts[performer];

                    dataGridView.Rows.Add(rowNumber, performer, rkkCount, requestsCount, totalCount);
                    rowNumber++;
                }

          
                dataGridView.CellFormatting += (cellSender, cellFormattingEventArgs) =>
                {
                    if (cellFormattingEventArgs.ColumnIndex == 0 && cellFormattingEventArgs.RowIndex >= 0)
                    {
                        cellFormattingEventArgs.Value = (cellFormattingEventArgs.RowIndex + 1).ToString();
                    }
                };

          
                Form resultForm = new Form();
                resultForm.Width = 800;
                resultForm.Height = 800;
                resultForm.Text = "Результат";

                resultForm.Controls.Add(dataGridView);

                //Кнопка сохранения в файл
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

                   
                    Paragraph heading = new Paragraph(
                        new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
                        new Run(new DocumentFormat.OpenXml.Wordprocessing.Text("Справка о неисполненных документах и обращениях граждан"))
                    );
                    body.AppendChild(heading);

                    
                    Paragraph info = new Paragraph(
                        new Run(new DocumentFormat.OpenXml.Wordprocessing.Text($"Не исполнено в срок {totalDocuments} документов, из них:")),
                        new Break(),
                        new Run(new RunProperties(new RunFonts { Ascii = "Arial" }), new DocumentFormat.OpenXml.Wordprocessing.Text($"\n- количество неисполненных входящих документов: {rkkDocuments}")),
                        new Break(),
                        new Run(new RunProperties(new RunFonts { Ascii = "Arial" }), new DocumentFormat.OpenXml.Wordprocessing.Text($"\n- количество неисполненных письменных обращений граждан: {requestsDocuments}"))
                    );

                    body.AppendChild(info);

  
                    Table table = new Table();

                  
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

                    
                    TableRow headerRow = new TableRow();

                    headerRow.AppendChild(new TableCell(new Paragraph(new Run(new DocumentFormat.OpenXml.Wordprocessing.Text("№"))))); // Добавление ячейки с заголовком номера строки

                    for (int i = 1; i < dataGridView.Columns.Count; i++)
                    {
                        DataGridViewColumn column = dataGridView.Columns[i];
                        TableCell headerCell = new TableCell(new Paragraph(new Run(new DocumentFormat.OpenXml.Wordprocessing.Text(column.HeaderText))));
                        headerRow.AppendChild(headerCell);
                    }

                    table.AppendChild(headerRow);

                    int rowNumber = 1; 

                    foreach (DataGridViewRow dataRow in dataGridView.Rows)
                    {
                        TableRow tableRow = new TableRow();
                        tableRow.AppendChild(new TableCell(new Paragraph(new Run(new DocumentFormat.OpenXml.Wordprocessing.Text(rowNumber.ToString()))))); // Добавление ячейки с номером строки

                        for (int i = 0; i < dataRow.Cells.Count; i++)
                        {
                           
                            if (i != 0)
                            {
                                string cellValue = dataRow.Cells[i].Value?.ToString() ?? string.Empty;

                                TableCell tableCell = new TableCell();

                                Paragraph paragraph = new Paragraph(
                                    new ParagraphProperties(
                                        new SpacingBetweenLines { After = "0" },
                                        new Justification { Val = JustificationValues.Center }
                                    ),
                                    new Run(
                                        new RunProperties(
                                            new RunFonts { Ascii = "Arial" },
                                            new FontSize { Val = "14" }
                                        ),
                                        new DocumentFormat.OpenXml.Wordprocessing.Text(cellValue)
                                    )
                                );

                                tableCell.AppendChild(paragraph);
                                tableRow.AppendChild(tableCell);
                            }
                        }

                        table.AppendChild(tableRow);
                        rowNumber++;
                    }

                    body.AppendChild(table);


                    // Дата запуска программы
                    DateTime currentDate = DateTime.Now;
                    string formattedDate = currentDate.ToString("dd.MM.yyyy");

                    Paragraph dateParagraph = new Paragraph(
                        new Run(new DocumentFormat.OpenXml.Wordprocessing.Text("Дата составления справки:\t" + formattedDate))
                    );

                    body.AppendChild(dateParagraph);

                }

                MessageBox.Show("Файл успешно сохранен.", "Сохранение", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }


    }
}
