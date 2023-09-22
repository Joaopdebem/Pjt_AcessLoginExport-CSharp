using System;
using System.Linq;
using ClosedXML.Excel;

namespace LerPlanilhaExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            var xls = new XLWorkbook(@"C:\Users\jpdeb\OneDrive\Área de Trabalho\Aviso Encapsulados\Encapsulados.xlsm");
            var planilha = xls.Worksheets.First(w => w.Name == "75dias");
            var totalLinhas = planilha.Rows().Count();

            for (int l = 2; l <= totalLinhas; l++)
            {
                var status = planilha.Cell($"G{l}").Value.ToString();

                if (status == "SIM" || status == "NÃO")
                {
                    var linhaInteira = planilha.Row(l).CellsUsed().Select(c => c.Value.ToString());
                    Console.WriteLine(string.Join(" - ", linhaInteira));
                }
            }
        }
    }
}
