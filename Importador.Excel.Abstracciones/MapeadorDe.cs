using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Importador.Excel.Abstracciones.Modelos;

namespace Importador.Excel.Abstracciones
{
    public class MapeadorDe<T> where T : new()
    {
        private readonly IImportadorExcel _importadorExcel;

        public MapeadorDe(IImportadorExcel importadorExcel)
        {
            _importadorExcel = importadorExcel;
        }

        public List<ImportacionDetalles<T>> Mapear(MemoryStream stream)
        {
            var importacionDetalles = new List<ImportacionDetalles<T>>();

            using (IImplementacionImportadorExcel importador = _importadorExcel.IniciarImportador(stream))
            {
                List<PropiedadColumna> propiedadesEncabezado = importador.ObtenerPropiedadesEncabezado();

                int cantidadFilas = importador.ObtenerCantidadFilas();

                for (int fila = 2; fila <= cantidadFilas; fila++)
                {
                    if (importador.FilaEstaVacia(fila))
                        continue;
                    var detalle = GenerarDetalle(fila, propiedadesEncabezado, importador);
                    importacionDetalles.Add(detalle);
                }

                return importacionDetalles;
            }
        }

        private ImportacionDetalles<T> GenerarDetalle(int fila, List<PropiedadColumna> columnasEncabezado, IImplementacionImportadorExcel importador)
        {
            var detalle = new ImportacionDetalles<T> { Fila = fila };
            var entidad = new T();

            foreach (PropertyInfo informacionPropiedad in typeof(T).GetProperties())
            {
                PropiedadColumna columnaEncontrada =
                    columnasEncabezado.FirstOrDefault(columna => columna.NombrePropiedad == informacionPropiedad.Name);

                if (columnaEncontrada is null)
                    continue;

                try
                {
                    object valorCelda = importador.ObtenerValorCelda(informacionPropiedad.PropertyType, fila, columnaEncontrada.NumeroColumna);
                    informacionPropiedad.SetValue(entidad, valorCelda);
                }
                catch
                {
                    var rango = importador.ObtenerRango(fila, columnaEncontrada.NumeroColumna);
                    detalle.Erores.Add(new Error(columnaEncontrada.NumeroColumna, $"El valor ingresado en la celda '{columnaEncontrada.NombrePropiedad}'({rango}) no es válido."));
                }
            }

            if (detalle.Erores.Count == 0)
                detalle.Entidad = entidad;

            return detalle;
        }
    }
}