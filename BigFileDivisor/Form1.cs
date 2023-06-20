using ClosedXML.Excel;
using System;
using System.Linq;
using System.Windows.Forms;

namespace BigFileDivisor
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string arquivoOrigem = openFileDialog1.FileName;
                int linhasPorArquivo = 10;

                SepararArquivo(arquivoOrigem, linhasPorArquivo);
            }
        }

        private void SepararArquivo(string arquivoOrigem, int linhasPorArquivo)
        {
            using (var workbookOrigem = new XLWorkbook(arquivoOrigem))
            {
                var worksheetOrigem = workbookOrigem.Worksheet(1);
                int totalLinhas = worksheetOrigem.RowsUsed().Count();
                int totalArquivos = totalLinhas / linhasPorArquivo;

                for (int i = 0; i < totalArquivos; i++)
                {
                    string nomeArquivo = $"arquivo_{i + 1}.xlsx";

                    using (var workbookDestino = new XLWorkbook())
                    {
                        var worksheetDestino = workbookDestino.Worksheets.Add("Sheet1");

                        int inicio = i * linhasPorArquivo;
                        int fim = inicio + linhasPorArquivo;

                        for (int linha = inicio; linha < fim; linha++)
                        {
                            var linhaOrigem = worksheetOrigem.Row(linha + 1);
                            var linhaDestino = worksheetDestino.Row(linha - inicio + 1);

                            linhaOrigem.CopyTo(linhaDestino);
                        }

                        workbookDestino.SaveAs(nomeArquivo);
                    }
                }

                MessageBox.Show($"{totalArquivos} arquivos Excel foram gerados.");
            }
        }
    }
}
