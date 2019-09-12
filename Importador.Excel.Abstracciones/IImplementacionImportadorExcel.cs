using System;
using System.Collections.Generic;

namespace Importador.Excel.Abstracciones
{
    public interface IImplementacionImportadorExcel : IDisposable
    {
        int ObtenerNumeroFilas();
        object ObtenerValorCelda(Type tipo, int fila, int columna);
        List<MapeoColumnaPropiedad> ObtenerMapeoColumnasPropiedades();
    }
}