using System.Collections.Generic;

namespace Importador.Excel.Abstracciones.Modelos
{
    public class Error
    {
        public int Columna { get; private set; }
        public string Mensaje { get; private set; }

        public Error(int columna, string mensaje)
        {
            Columna = columna;
            Mensaje = mensaje;
        }
    }

    public class ImportacionDetalles<T> where T : new()
    {
        public int Fila { get; set; }
        public T Entidad { get; set; }
        public List<Error> Erores { get; set; }

        public ImportacionDetalles()
        {
            Erores = new List<Error>();
        }
    }
}