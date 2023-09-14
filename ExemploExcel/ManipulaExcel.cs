using ClosedXML.Excel;
using Microsoft.VisualBasic.ApplicationServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = ClosedXML.Excel;

namespace ExemploExcel
{
    internal class ManipulaExcel
    {

        public void Salvar(string nome, string email)
        {
            try {
                var pasta = new XLWorkbook("teste.xlsx");
                var planilha = pasta.Worksheets.First(w => w.Name == "Pla1");
                var itens = pasta.Worksheet(1).RowsUsed();
                int linha = planilha.Rows().Count();
                MessageBox.Show("Linhas " + linha);
                if (linha == 0)
                {
                    planilha.Cell(1,1).Value = "Nome";
                    planilha.Cell(1,2).Value = "Email";
                }

                planilha.Cell(linha + 1, 1).Value = nome;
                planilha.Cell(linha + 1, 2).Value = email;

                pasta.Save();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
           
        }

        public void LerArquivo()
        {
            var pasta = new XLWorkbook("teste.xlsx");
            var planilha = pasta.Worksheets.First(w => w.Name == "Pla1");
            var itens = pasta.Worksheet(1).RowsUsed();
            int linha = planilha.Rows().Count();
            for (int i = 1; i < linha; i++)
            {
                string nome = planilha.Cell(linha, 1).Value.ToString();

                //Adionar a lista
            }
        }
    }
}
