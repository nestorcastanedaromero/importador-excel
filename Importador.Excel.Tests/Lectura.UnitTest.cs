using System;
using System.Collections.Generic;
using System.IO;
using System.IO.MemoryMappedFiles;
using System.Linq;
using System.Reflection;
using Importador.Excel.Tests.Properties;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace Importador.Excel.Tests
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void DebeLeerArchivo()
        {
            List<DtoPrueba> datos =new List<DtoPrueba>();

            using (var stream = new MemoryStream(Resources.Prueba))
            {
                var mapeador = new MapeadorDe<DtoPrueba>();

                datos = mapeador.Mapear(stream);
            };

            Assert.AreEqual(datos.Count, 1);
            Assert.AreEqual(1, datos[0].Id);
            Assert.AreEqual("Juan", datos[0].Nombre);
        }

    }

    public class DtoPrueba
    {
        public int Id { get; set; }
        public string Nombre { get; set; }
    }

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

