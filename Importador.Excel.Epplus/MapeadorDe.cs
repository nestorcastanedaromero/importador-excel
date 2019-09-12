using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using OfficeOpenXml;

namespace Importador.Excel.Epplus
{
    public class MapeadorDe<T> where T : new()
    {
        public List<T> Mapear(MemoryStream stream)
        {
            List<T> datos = new List<T>();

            using (var package = new ExcelPackage(stream))
            {

                ExcelWorksheet hoja1 = package.Workbook.Worksheets[0];

                List<MapeoColumnaPropiedad> mapeadores = hoja1.Cells[1, 1, 1, hoja1.Dimension.End.Column].Select(fila =>
                {
                    if (string.IsNullOrEmpty(fila.Text))
                        throw new Exception($"El encabezado de la columna {fila.Address} no puede estar vacío");
                    return new MapeoColumnaPropiedad(fila.Start.Column, fila.Text);
                }).ToList();

                int cantidadFilas = hoja1.Dimension.Rows;

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
                        object valorCelda = ObtenerValorCelda(propertyInfo.PropertyType,
                            hoja1.Cells[fila, mapaPropiedad.Columna]);
                        propertyInfo.SetValue(dtoPrueba, valorCelda);
                    }

                    datos.Add(dtoPrueba);
                }
                return datos;
            }
        }

        private object ObtenerValorCelda(Type tipo, ExcelRange excelRange)
        {
            if (excelRange.Value == null)
            {
                return null;
            }
            if (tipo == typeof(int))
            {
                return excelRange.GetValue<int>();
            }
            if (tipo == typeof(double))
            {
                return excelRange.GetValue<double>();
            }
            if (tipo == typeof(DateTime))
            {
                return excelRange.GetValue<DateTime>();
            }
            return excelRange.GetValue<string>();
        }
    }
}