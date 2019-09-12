namespace Importador.Excel.Epplus
{
    public class MapeoColumnaPropiedad
    {
        public MapeoColumnaPropiedad(int columna, string propiedad)
        {
            Columna = columna;
            Propiedad = propiedad;
        }

        public int Columna { get; set; }
        public string Propiedad { get; set; }
    }
}