using System.IO;

namespace Importador.Excel.Abstracciones
{
    public interface IImportadorExcel
    {
        IImplementacionImportadorExcel IniciarImportador(Stream stream, int? pagina = null);
    }
}