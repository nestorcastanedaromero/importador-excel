using System;
using System.Collections.Generic;
using Importador.Excel.Abstracciones.Modelos;

namespace Importador.Excel.Abstracciones
{
    public interface IImplementacionImportadorExcel : IDisposable
    {
        int ObtenerCantidadFilas();

        object ObtenerValorCelda(Type tipo, int fila, int columna);

        List<PropiedadColumna> ObtenerPropiedadesEncabezado();

        bool FilaEstaVacia(int numeroFila);

        string ObtenerRango(int fila, int columna);
    }
}