using ClosedXML.Excel;
using System;
using System.IO;
using System.Windows.Forms;
using Xceed.Words.NET;

namespace BigFileDivisor
{
    public partial class Form1 : Form
    {
        private const string DefaultTextBoxText = "Rows per file";
        private FolderBrowserDialog folderBrowserDialog;

        public Form1()
        {
            InitializeComponent();
            textBox1.Text = DefaultTextBoxText;
            textBox1.ForeColor = System.Drawing.SystemColors.GrayText;
            textBox1.UseWaitCursor = false;

            textBox1.MouseEnter += TextBox1_MouseEnter;
            textBox1.MouseDown += TextBox1_MouseDown;

            // Inicializar o FolderBrowserDialog
            folderBrowserDialog = new FolderBrowserDialog();
            folderBrowserDialog.Description = "Select destiny folder";
            folderBrowserDialog.ShowNewFolderButton = true;
        }

        private void TextBox1_MouseEnter(object sender, EventArgs e)
        {
            if (!textBox1.Focused)
                textBox1.Cursor = Cursors.Default;
        }

        private void TextBox1_MouseDown(object sender, MouseEventArgs e)
        {
            textBox1.Cursor = Cursors.IBeam;
        }

        private void textBox1_MouseUp(object sender, MouseEventArgs e)
        {
            textBox1.Cursor = Cursors.IBeam;
        }

        private void TextBox1_TextChanged(object sender, EventArgs e)
        {
            textBox1.Text = "";
            textBox1.ForeColor = System.Drawing.SystemColors.WindowText;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string originalArchive = openFileDialog1.FileName;
                int rowsByArchive = int.Parse(textBox1.Text);

                if (Path.GetExtension(originalArchive) == ".xlsx")
                {
                    SplitExcel(originalArchive, rowsByArchive);
                }
                else if (Path.GetExtension(originalArchive) == ".txt")
                {
                    SplitTxt(originalArchive, rowsByArchive);
                }
                else if (Path.GetExtension(originalArchive) == ".docx")
                {
                    SplitDocx(originalArchive, rowsByArchive);
                }
                else
                {
                    MessageBox.Show("Unsupported file. Please, select a .xlsx, .txt or .docx file.");
                }
            }
        }

        private void SplitExcel(string originalArchive, int rowsByArchive)
        {
            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                string destinationFolder = folderBrowserDialog.SelectedPath;

                using (var workbook = new XLWorkbook(originalArchive))
                {
                    var worksheet = workbook.Worksheet(1);
                    var range = worksheet.RangeUsed();

                    int totalRows = range.RowCount();
                    int totalArchives = (int)Math.Ceiling((double)totalRows / rowsByArchive);

                    for (int i = 0; i < totalArchives; i++)
                    {
                        string archiveName = Path.Combine(destinationFolder, $"archive_{i + 1}.xlsx");

                        using (var newWorkbook = new XLWorkbook())
                        {
                            var newWorksheet = newWorkbook.Worksheets.Add("Sheet1");

                            int begin = i * rowsByArchive + 1;
                            int end = Math.Min(begin + rowsByArchive - 1, totalRows);

                            var rangeCopy = range.Range(begin, 1, end, range.ColumnCount());
                            rangeCopy.CopyTo(newWorksheet.Cell(1, 1));

                            newWorkbook.SaveAs(archiveName);
                        }
                    }

                    MessageBox.Show($"{totalArchives} XSLX archives were created. Saved in: {destinationFolder}");
                }
            }
        }

        private void SplitTxt(string originalArchive, int rowsByArchive)
        {
            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                string destinationFolder = folderBrowserDialog.SelectedPath;

                string[] rows = File.ReadAllLines(originalArchive);
                int totalRows = rows.Length;
                int totalArchives = (int)Math.Ceiling((double)totalRows / rowsByArchive);

                for (int i = 0; i < totalArchives; i++)
                {
                    string archiveName = Path.Combine(destinationFolder, $"archive_{i + 1}.txt");

                    int begin = i * rowsByArchive;
                    int end = Math.Min(begin + rowsByArchive, totalRows);

                    string[] selectedRows = new string[end - begin];
                    Array.Copy(rows, begin, selectedRows, 0, end - begin);

                    File.WriteAllLines(archiveName, selectedRows);
                }

                MessageBox.Show($"{totalArchives} TXT archives were created. Saved in: {destinationFolder}");
            }
        }

        private void SplitDocx(string originalArchive, int rowsByArchive)
        {
            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                string destinationFolder = folderBrowserDialog.SelectedPath;

                using (var document = DocX.Load(originalArchive))
                {
                    var paragraphs = document.Paragraphs;

                    int totalParagraphs = paragraphs.Count;
                    int totalArchives = (int)Math.Ceiling((double)totalParagraphs / rowsByArchive);

                    for (int i = 0; i < totalArchives; i++)
                    {
                        string archiveName = Path.Combine(destinationFolder, $"archive_{i + 1}.docx");

                        using (var newDocument = DocX.Create(archiveName))
                        {
                            int begin = i * rowsByArchive;
                            int end = Math.Min(begin + rowsByArchive, totalParagraphs);

                            for (int j = begin; j < end; j++)
                            {
                                var paragraph = paragraphs[j];
                                newDocument.InsertParagraph(paragraph);
                            }

                            newDocument.Save();
                        }
                    }

                    MessageBox.Show($"{totalArchives} DOCX archives were created. Saved in: {destinationFolder}");
                }
            }
        }
    }
}
