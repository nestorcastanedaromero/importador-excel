namespace Importador.Excel.Abstracciones.Modelos
{
    public class PropiedadColumna
    {
        public PropiedadColumna(int numeroColumna, string nombrePropiedad)
        {
            NumeroColumna = numeroColumna;
            NombrePropiedad = nombrePropiedad;
        }

        public int NumeroColumna { get; set; }
        public string NombrePropiedad { get; set; }
    }
}