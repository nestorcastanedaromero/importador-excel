using Importador.Excel.Abstracciones;
using Importador.Excel.Epplus;
using Importador.Excel.Tests.Properties;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using Importador.Excel.Abstracciones.Modelos;

namespace Importador.Excel.Tests
{
    [TestClass]
    public class ImportadorExcelUnitTests
    {
        [TestMethod]
        public void DebeLeerArchivo_ConTodosLosTiposCorrectos()
        {
            List<ImportacionDetalles<DtoPruebaTodosLosFormatos>> datos;

            using (var stream = new MemoryStream(Resources.PruebaFormatos))
            {
                var mapeador = new MapeadorDe<DtoPruebaTodosLosFormatos>(new ImportadorExcelEpPlus());

                datos = mapeador.Mapear(stream);
            }

            Assert.AreEqual(datos.Count, 5);
            Assert.AreEqual(1, datos[0].Entidad.ValorEntero);
            Assert.AreEqual(new DateTime(2017, 1, 1), datos[0].Entidad.Fecha);
            Assert.AreEqual(1003.50, datos[0].Entidad.ValorFloat);
            Assert.AreEqual(1001.50, datos[0].Entidad.ValorDouble);
            Assert.AreEqual(1002.50m, datos[0].Entidad.ValorDecimal);
            Assert.AreEqual(false, datos[0].Entidad.ValorBool);
        }

        [TestMethod]
        public void DebeLeerArchivo()
        {
            List<ImportacionDetalles<DtoPrueba>> datos;

            using (var stream = new MemoryStream(Resources.Prueba))
            {
                var mapeador = new MapeadorDe<DtoPrueba>(new ImportadorExcelEpPlus());

                datos = mapeador.Mapear(stream);
            }

            Assert.AreEqual(datos.Count, 1);
            Assert.AreEqual(1, datos[0].Entidad.Id);
            Assert.AreEqual("Juan", datos[0].Entidad.Nombre);
        }

        [TestMethod]
        public void Debe_Mapeador_ImportarFechasCorrectamente()
        {
            List<ImportacionDetalles<DtoPruebaFecha>> datos;

            using (var stream = new MemoryStream(Resources.PruebaFechaCorrecta))
            {
                var mapeador = new MapeadorDe<DtoPruebaFecha>(new ImportadorExcelEpPlus());

                datos = mapeador.Mapear(stream);
            }

            Assert.AreEqual(datos.Count, 1);
            Assert.AreEqual(1, datos[0].Entidad.Id);
            Assert.AreEqual("Pedro", datos[0].Entidad.Nombre);
            Assert.AreEqual(new DateTime(2017, 7, 7), datos[0].Entidad.Fecha);
        }

        [TestMethod]
        public void Debe_CuandoFechaIncorrecta_Retornar_ListaErrores()
        {
            List<ImportacionDetalles<DtoPruebaFecha>> resultado;

            using (var stream = new MemoryStream(Resources.PruebaFechaInCorrecta))
            {
                var mapeador = new MapeadorDe<DtoPruebaFecha>(new ImportadorExcelEpPlus());

                resultado = mapeador.Mapear(stream);
            }

            Assert.AreEqual(1, resultado.Count);
            Assert.AreEqual(1, resultado[0].Erores.Count);
            StringAssert.Contains(resultado[0].Erores[0].Mensaje, "El valor ingresado en la celda 'Fecha'(C2) no es válido.");
        }

        [TestMethod]
        public void Debe_CuandoTengaNumeroComoTextos_ImportarlosBien()
        {
            List<ImportacionDetalles<DtoPruebaNumeroComoTexto>> resultado;

            using (var stream = new MemoryStream(Resources.PruebaImportarColumnaTexto))
            {
                var mapeador = new MapeadorDe<DtoPruebaNumeroComoTexto>(new ImportadorExcelEpPlus());

                resultado = mapeador.Mapear(stream);
            }

            Assert.AreEqual(1, resultado.Count);
            StringAssert.Contains(resultado[0].Entidad.CentroCosto, "0101");
        }

        [TestMethod]
        public void Debe_CuandoTengaNumeroConDecimales_ImportarlosBien()
        {
            List<ImportacionDetalles<DtoPruebaNumeroConDecimales>> resultado;

            using (var stream = new MemoryStream(Resources.PruebaNumeroConDecimales))
            {
                var mapeador = new MapeadorDe<DtoPruebaNumeroConDecimales>(new ImportadorExcelEpPlus());

                resultado = mapeador.Mapear(stream);
            }

            Assert.AreEqual(1, resultado.Count);
            Assert.AreEqual(resultado[0].Entidad.valor, 1.5);
        }

        [TestMethod]
        public void Debe_CuandoTengaNumeroConDecimalesQueTienenComa_ImportarlosBien()
        {
            List<ImportacionDetalles<DtoPruebaNumeroConDecimales>> resultado;

            using (var stream = new MemoryStream(Resources.PruebaNumeroConDecimales))
            {
                var mapeador = new MapeadorDe<DtoPruebaNumeroConDecimales>(new ImportadorExcelEpPlus());

                resultado = mapeador.Mapear(stream);
            }

            Assert.AreEqual(1, resultado.Count);
            Assert.AreEqual(resultado[0].Entidad.valor, 1.5);
        }

        [TestMethod]
        public void Debe_SiElTipoDellegadaEsDecimalYElTipoDeLaCeldaNoEsNumero_GenerarError()
        {
            List<ImportacionDetalles<DtoPruebaNumeroConDecimales>> resultado;

            using (var stream = new MemoryStream(Resources.PruebaNumeroComoTexto))
            {
                var mapeador = new MapeadorDe<DtoPruebaNumeroConDecimales>(new ImportadorExcelEpPlus());

                resultado = mapeador.Mapear(stream);
            }

            Assert.AreEqual("El valor ingresado en la celda 'valor'(C2) no es válido.", resultado[0].Erores[0].Mensaje);
        }

        [TestMethod]
        public void Debe_CuandoTengaTengaNotacionCientfica_ImportarlosBien()
        {
            List<ImportacionDetalles<DtoPruebaNumeroConDecimales>> resultado;

            using (var stream = new MemoryStream(Resources.PruebaNumeroConNotacionCientifica))
            {
                var mapeador = new MapeadorDe<DtoPruebaNumeroConDecimales>(new ImportadorExcelEpPlus());

                resultado = mapeador.Mapear(stream);
            }

            Assert.AreEqual(1, resultado.Count);
            Assert.AreEqual(12345678901234, resultado[0].Entidad.valor);
        }
    }

    public class DtoPrueba
    {
        public int Id { get; set; }
        public string Nombre { get; set; }
    }

    public class DtoPruebaTodosLosFormatos
    {
        public int ValorEntero { get; set; }
        public DateTime Fecha { get; set; }
        public float ValorFloat { get; set; }

        public double ValorDouble { get; set; }
        public decimal ValorDecimal { get; set; }
        public bool ValorBool { get; set; }
    }

    public class DtoPruebaFecha : DtoPrueba
    {
        public DateTime Fecha { get; set; }
    }

    public class DtoPruebaNumeroComoTexto : DtoPrueba
    {
        public string CentroCosto { get; set; }
    }

    public class DtoPruebaNumeroConDecimales : DtoPrueba
    {
        public double valor { get; set; }
    }

    public class DtoPruebaNumeroConNotacionCientifica : DtoPrueba
    {
        public double valor { get; set; }
    }
}