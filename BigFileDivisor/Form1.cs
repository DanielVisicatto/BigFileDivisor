using ClosedXML.Excel;
using System;
using System.IO;
using System.Windows.Forms;
using Xceed.Words.NET;

namespace BigFileDivisor
{
    public partial class Form1 : Form
    {
        private const string DefaultTextBoxText = "Numb. of rows";
        public Form1()
        {
            InitializeComponent();
            textBox1.Text = DefaultTextBoxText;
            textBox1.ForeColor = System.Drawing.SystemColors.GrayText;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            textBox1.Text = "";
            textBox1.ForeColor = System.Drawing.SystemColors.WindowText;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string arquivoOrigem = openFileDialog1.FileName;
                int linhasPorArquivo = int.Parse(textBox1.Text);

                if (Path.GetExtension(arquivoOrigem) == ".xlsx")
                {
                    SepararArquivoExcel(arquivoOrigem, linhasPorArquivo);
                }
                else if (Path.GetExtension(arquivoOrigem) == ".txt")
                {
                    SepararArquivoTexto(arquivoOrigem, linhasPorArquivo);
                }
                else if (Path.GetExtension(arquivoOrigem) == ".docx")
                {
                    SepararArquivoDocx(arquivoOrigem, linhasPorArquivo);
                }
                else
                {
                    MessageBox.Show("Formato de arquivo não suportado. Por favor, selecione um arquivo .xlsx, .txt ou .docx.");
                }
            }
        }

        private void SepararArquivoExcel(string arquivoOrigem, int linhasPorArquivo)
        {
            using (var workbook = new XLWorkbook(arquivoOrigem))
            {
                var worksheet = workbook.Worksheet(1);
                var range = worksheet.RangeUsed();

                int totalLinhas = range.RowCount();
                int totalArquivos = (int)Math.Ceiling((double)totalLinhas / linhasPorArquivo);

                for (int i = 0; i < totalArquivos; i++)
                {
                    string nomeArquivo = $"arquivo_{i + 1}.xlsx";

                    using (var newWorkbook = new XLWorkbook())
                    {
                        var newWorksheet = newWorkbook.Worksheets.Add("Sheet1");

                        int inicio = i * linhasPorArquivo + 1;
                        int fim = Math.Min(inicio + linhasPorArquivo - 1, totalLinhas);

                        var rangeCopy = range.Range(inicio, 1, fim, range.ColumnCount());
                        rangeCopy.CopyTo(newWorksheet.Cell(1, 1));

                        newWorkbook.SaveAs(nomeArquivo);
                    }
                }

                MessageBox.Show($"{totalArquivos} arquivos Excel foram gerados.");
            }
        }

        private void SepararArquivoTexto(string arquivoOrigem, int linhasPorArquivo)
        {
            string[] linhas = File.ReadAllLines(arquivoOrigem);
            int totalLinhas = linhas.Length;
            int totalArquivos = (int)Math.Ceiling((double)totalLinhas / linhasPorArquivo);

            for (int i = 0; i < totalArquivos; i++)
            {
                string nomeArquivo = $"arquivo_{i + 1}.txt";

                int inicio = i * linhasPorArquivo;
                int fim = Math.Min(inicio + linhasPorArquivo, totalLinhas);

                string[] linhasSelecionadas = new string[fim - inicio];
                Array.Copy(linhas, inicio, linhasSelecionadas, 0, fim - inicio);

                File.WriteAllLines(nomeArquivo, linhasSelecionadas);
            }

            MessageBox.Show($"{totalArquivos} arquivos de texto foram gerados.");
        }

        private void SepararArquivoDocx(string arquivoOrigem, int linhasPorArquivo)
        {
            textBox1.Click += (sender, e) =>
            {
                textBox1.Text = "";
                textBox1.ForeColor = System.Drawing.SystemColors.WindowText;
            };

            using (var document = DocX.Load(arquivoOrigem))
            {
                var paragraphs = document.Paragraphs;

                int totalParagrafos = paragraphs.Count;
                int totalArquivos = (int)Math.Ceiling((double)totalParagrafos / linhasPorArquivo);

                for (int i = 0; i < totalArquivos; i++)
                {
                    string nomeArquivo = $"arquivo_{i + 1}.docx";

                    using (var newDocument = DocX.Create(nomeArquivo))
                    {
                        int inicio = i * linhasPorArquivo;
                        int fim = Math.Min(inicio + linhasPorArquivo, totalParagrafos);

                        for (int j = inicio; j < fim; j++)
                        {
                            var paragraph = paragraphs[j];
                            newDocument.InsertParagraph(paragraph);
                        }

                        newDocument.Save();
                    }
                }

                MessageBox.Show($"{totalArquivos} arquivos DOCX foram gerados.");
            }
        }
    }
}
