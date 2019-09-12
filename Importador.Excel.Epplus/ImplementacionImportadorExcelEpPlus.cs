using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Importador.Excel.Abstracciones;
using OfficeOpenXml;

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

        public int ObtenerNumeroFilas()
        {
            return _hoja.Dimension.Rows;
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

        public List<MapeoColumnaPropiedad> ObtenerMapeoColumnasPropiedades()
        {
            return _hoja.Cells[1, 1, 1, _hoja.Dimension.End.Column].Select(fila =>
            {
                if (string.IsNullOrEmpty(fila.Text))
                    throw new Exception($"El encabezado de la columna {fila.Address} no puede estar vacío");
                return new MapeoColumnaPropiedad(fila.Start.Column, fila.Text);
            }).ToList();
        }

        public void Dispose()
        {
            _excelPackage?.Dispose();
        }
    }
}