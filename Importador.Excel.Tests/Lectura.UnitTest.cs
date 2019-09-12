using System;
using Importador.Excel.Abstracciones;
using Importador.Excel.Epplus;
using Importador.Excel.Tests.Properties;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.IO;

namespace Importador.Excel.Tests
{
    [TestClass]
    public class ImportadorExcelUnitTests
    {
        [TestMethod]
        public void DebeLeerArchivo()
        {
            List<DtoPrueba> datos;

            using (var stream = new MemoryStream(Resources.Prueba))
            {
                var mapeador = new MapeadorDe<DtoPrueba>(new ImportadorExcelEpPlus());

                datos = mapeador.Mapear(stream);
            }

            Assert.AreEqual(datos.Count, 1);
            Assert.AreEqual(1, datos[0].Id);
            Assert.AreEqual("Juan", datos[0].Nombre);
        }

        [TestMethod]
        public void Debe_Mapeador_ImportarFechasCorrectamente()
        {
            List<DtoPruebaFecha> datos;

            using (var stream = new MemoryStream(Resources.PruebaFechaCorrecta))
            {
                var mapeador = new MapeadorDe<DtoPruebaFecha>(new ImportadorExcelEpPlus());

                datos = mapeador.Mapear(stream);
            }

            Assert.AreEqual(datos.Count, 1);
            Assert.AreEqual(1, datos[0].Id);
            Assert.AreEqual("Pedro", datos[0].Nombre);
            Assert.AreEqual(new DateTime(2017, 7, 7), datos[0].Fecha);
        }

        [TestMethod]
        public void Debe_CuandoFechaIncorrecta_Retornar_ListaErrores()
        {
            List<DtoPruebaFecha> datos;

            using (var stream = new MemoryStream(Resources.PruebaFechaInCorrecta))
            {
                var mapeador = new MapeadorDe<DtoPruebaFecha>(new ImportadorExcelEpPlus());

                datos = mapeador.Mapear(stream);
            }

            Assert.AreEqual(datos.Count, 0);
            Assert.AreEqual(1, datos[0].Id);
            Assert.AreEqual("Pedro", datos[0].Nombre);
            Assert.AreEqual(new DateTime(2017, 7, 7), datos[0].Fecha);
        }

    }

    public class DtoPrueba
    {
        public int Id { get; set; }
        public string Nombre { get; set; }
    }

    public class DtoPruebaFecha : DtoPrueba
    {
        public DateTime Fecha { get; set; }
    }
}

