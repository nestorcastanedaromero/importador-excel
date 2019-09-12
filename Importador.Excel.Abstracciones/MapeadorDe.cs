using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace Importador.Excel.Abstracciones
{
    public class MapeadorDe<T> where T : new()
    {
        private readonly IImportadorExcel _importadorExcel;

        public MapeadorDe(IImportadorExcel importadorExcel)
        {
            _importadorExcel = importadorExcel;
        }

        public List<T> Mapear(MemoryStream stream)
        {
            var datos = new List<T>();

            using (IImplementacionImportadorExcel importador = _importadorExcel.IniciarImportador(stream))
            {
                List<MapeoColumnaPropiedad> mapeadores = importador.ObtenerMapeoColumnasPropiedades();

                int cantidadFilas = importador.ObtenerNumeroFilas();

                for (var fila = 2; fila < cantidadFilas; fila++)
                {
                    PropertyInfo[] propiedadesDto = typeof(T).GetProperties();
                    var dtoPrueba = new T();
                    foreach (PropertyInfo propertyInfo in propiedadesDto)
                    {
                        MapeoColumnaPropiedad mapaPropiedad =
                            mapeadores.FirstOrDefault(mapa => mapa.Propiedad == propertyInfo.Name);
                        if (mapaPropiedad is null)
                            continue;
                        object valorCelda = importador.ObtenerValorCelda(propertyInfo.PropertyType, fila, mapaPropiedad.Columna);
                        propertyInfo.SetValue(dtoPrueba, valorCelda);
                    }

                    datos.Add(dtoPrueba);
                }
                return datos;
            }
        }
    }
}