using System.IO;
using Importador.Excel.Abstracciones;

namespace Importador.Excel.Epplus
{
    public class ImportadorExcelEpPlus : IImportadorExcel
    {
        public IImplementacionImportadorExcel IniciarImportador(Stream stream, int? pagina)
        {
            return new ImplementacionImportadorExcelEpPlus(stream, pagina.GetValueOrDefault());
        }
    }
}