using System.Collections.Generic;
using System.IO;
using Importador.Excel.Epplus;
using Importador.Excel.Tests.Properties;
using Microsoft.VisualStudio.TestTools.UnitTesting;

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
}

