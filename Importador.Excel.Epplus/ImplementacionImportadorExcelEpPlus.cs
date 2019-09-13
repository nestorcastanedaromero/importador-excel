using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Importador.Excel.Abstracciones;
using Importador.Excel.Abstracciones.Modelos;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace Importador.Excel.Epplus
{
    public class ImplementacionImportadorExcelEpPlus : IImplementacionImportadorExcel
    {
        private readonly ExcelPackage _excelPackage;
        private readonly ExcelWorksheet _hoja;

        public ImplementacionImportadorExcelEpPlus(Stream stream, int pagina = 0)
        {
            _excelPackage = new ExcelPackage(stream);
            _hoja = _excelPackage.Workbook.Worksheets[pagina];
        }

        public int ObtenerCantidadFilas()
        {
            return _hoja.Dimension.Rows;
        }

        public string ObtenerRango(int fila, int columna)
        {
            return _hoja.SelectedRange[fila, columna].Address;
        }

        public object ObtenerValorCelda(Type tipo, int fila, int columna)
        {
            var excelRange = _hoja.Cells[fila, columna];
            if (excelRange.Value == null)
            {
                return null;
            }
            if (tipo == typeof(int))
            {
                return excelRange.GetValue<int>();
            }
            if (tipo == typeof(float))
            {
                ValidarNumero(excelRange);
                return excelRange.GetValue<float>();
            }
            if (tipo == typeof(double))
            {
                ValidarNumero(excelRange);
                return excelRange.GetValue<double>();
            }
            if (tipo == typeof(decimal))
            {
                ValidarNumero(excelRange);
                return excelRange.GetValue<decimal>();
            }
            if (tipo == typeof(bool))
            {
                return excelRange.GetValue<bool>();
            }
            if (tipo == typeof(DateTime))
            {
                return excelRange.GetValue<DateTime>();
            }
            return excelRange.GetValue<string>();
        }

        private static void ValidarNumero(ExcelRange excelRange)
        {
            if (!excelRange.Value.IsNumeric())
                throw new ArgumentException("No se puede convertir un valor no numérico en decimal");
        }

        public List<PropiedadColumna> ObtenerPropiedadesEncabezado()
        {
            return _hoja.Cells[1, 1, 1, _hoja.Dimension.End.Column].Select(fila =>
            {
                if (string.IsNullOrEmpty(fila.Text))
                    throw new Exception($"El encabezado de la columna {fila.Address} no puede estar vacío");
                return new PropiedadColumna(fila.Start.Column, fila.Text);
            }).ToList();
        }

        public bool FilaEstaVacia(int numeroFila)
        {
            var asas = ((object[,])_hoja.SelectedRange[numeroFila, _hoja.Dimension.Start.Column, numeroFila,
                     _hoja.Dimension.End.Column].Value);
            foreach (var result in asas)
            {
                if (result != null)
                    return false;
            }

            return true;
        }

        public void Dispose()
        {
            _excelPackage?.Dispose();
        }
    }
}