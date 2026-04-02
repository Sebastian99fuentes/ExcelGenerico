using ClosedXML.Excel;
using System.IO;
using System;
using System.Globalization;

namespace Web.Api.Endpoints.Todos;

internal sealed class Excel : IEndpoint
{
    public sealed class Factura
    {
        public string NumeroFactura { get; set; }
        public decimal Monto { get; set; }
        public string Sucursal { get; set; }
        public DateTime Fecha { get; set; }
        public string Cliente { get; set; }
    }

    public void MapEndpoint(IEndpointRouteBuilder app)
    {
        app.MapGet("todos/Excel", (
            ) =>
        {
            // Crear datos de prueba (1 objeto inventado)
            var listaFacturasPrueba = new List<Factura>
            {
                new Factura
                {
                    NumeroFactura = "F001-001",
                    Monto = 1500.50m,
                    Sucursal = "Casa Matriz",
                    Fecha = DateTime.Now, // Corrected to use DateTime directly
                    Cliente = "Juan Pérez"
                },
                 new Factura
                {
                    NumeroFactura = "F001-002",
                    Monto = 2500.75m,
                    Sucursal = "Sucursal Norte",
                    Fecha = DateTime.Now.AddDays(-1), // Corrected to use DateTime directly
                    Cliente = "María Gómez"
                 }
            };

            // Generar el Excel y obtener el Base64
            string base64Excel = GenerarExcelEnBase64(listaFacturasPrueba);

            // Retornar el resultado
            return Results.Ok(new
            {
                success = true,
                fileName = "reporte_facturas.xlsx",
                base64 = base64Excel
            });
        })
        .WithTags(Tags.Todos)
        .RequireAuthorization();
    }

    /// <summary>
    /// Genera el Excel personalizado y lo retorna como string en Base64
    /// </summary>
    /// <param name="listaFacturas">Lista de facturas a exportar</param>
    /// <returns>String en Base64 del archivo Excel</returns>
    private string GenerarExcelEnBase64(List<Factura> listaFacturas)
    {
        // Usar MemoryStream en lugar de guardar en disco
        using var memoryStream = new MemoryStream();
        // Crear el libro de trabajo y la hoja

        using (var workbook = new XLWorkbook())
        {
            IXLWorksheet worksheet = workbook.Worksheets.Add("ReporteFacturas");

            // --- A. INSERTAR EL LOGO (opcional, si no tienes logo comenta esta sección) ---
            // Para pruebas sin logo, puedes comentar esta parte
            // Si tienes un logo embebido como recurso, aquí podrías cargarlo
            // string rutaLogo = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logo.png");
            // if (File.Exists(rutaLogo))
            // {
            //     IXLPicture picture = worksheet.AddPicture(rutaLogo)
            //                            .MoveTo(worksheet.Cell("A1"))
            //                            .Scale(0.5);
            // }

            // --- B. CREAR EL TÍTULO (debajo del logo) ---
            int filaTitulo = 4;
            IXLCell celdaTitulo = worksheet.Cell(filaTitulo, 1);
            celdaTitulo.Value = "REPORTE DE FACTURACIÓN";
            celdaTitulo.Style.Font.Bold = true;
            celdaTitulo.Style.Font.FontSize = 18;
            celdaTitulo.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            // --- C. ESCRIBIR LOS DATOS Y APLICAR FORMATOS ---
            int filaInicioDatos = filaTitulo + 2;

            // Escribir los encabezados de las columnas
            worksheet.Cell(filaInicioDatos, 1).Value = "Número Factura";
            worksheet.Cell(filaInicioDatos, 2).Value = "Monto";
            worksheet.Cell(filaInicioDatos, 3).Value = "Sucursal";
            worksheet.Cell(filaInicioDatos, 4).Value = "Fecha";
            worksheet.Cell(filaInicioDatos, 5).Value = "Cliente";

            // Insertar los datos
            for (int i = 0; i < listaFacturas.Count; i++)
            {
                int filaActual = filaInicioDatos + 1 + i;
                worksheet.Cell(filaActual, 1).Value = listaFacturas[i].NumeroFactura;
                worksheet.Cell(filaActual, 2).Value = listaFacturas[i].Monto;
                worksheet.Cell(filaActual, 3).Value = listaFacturas[i].Sucursal;
                worksheet.Cell(filaActual, 4).Value = listaFacturas[i].Fecha.ToString("dd-MM-yyyy", CultureInfo.InvariantCulture); // Corrected to use IFormatProvider
                worksheet.Cell(filaActual, 5).Value = listaFacturas[i].Cliente;
            }

            // Obtener rangos después de insertar datos
#pragma warning disable CS8602 // Dereference of a possibly null reference.
            int ultimaFilaUsada = worksheet.LastRowUsed().RowNumber();
#pragma warning restore CS8602 // Dereference of a possibly null reference.
#pragma warning disable CS8602 // Dereference of a possibly null reference.
            int ultimaColumnaUsada = worksheet.LastColumnUsed().ColumnNumber();
#pragma warning restore CS8602 // Dereference of a possibly null reference.

            // Aplicar formato a cada columna (fondo verde y letra blanca)
            // Columna 2 (Monto)
            IXLRange rangoMonto = worksheet.Range(filaInicioDatos + 1, 2, ultimaFilaUsada, 2);
            rangoMonto.Style.Fill.BackgroundColor = XLColor.Green;
            rangoMonto.Style.Font.FontColor = XLColor.White;

            // Columna 3 (Sucursal)
            IXLRange rangoSucursal = worksheet.Range(filaInicioDatos + 1, 3, ultimaFilaUsada, 3);
            rangoSucursal.Style.Fill.BackgroundColor = XLColor.Green;
            rangoSucursal.Style.Font.FontColor = XLColor.White;

            // Columna 4 (Fecha)
            IXLRange rangoFecha = worksheet.Range(filaInicioDatos + 1, 4, ultimaFilaUsada, 4);
            rangoFecha.Style.Fill.BackgroundColor = XLColor.Green;
            rangoFecha.Style.Font.FontColor = XLColor.White;

            // Columna 5 (Cliente)
            IXLRange rangoCliente = worksheet.Range(filaInicioDatos + 1, 5, ultimaFilaUsada, 5);
            rangoCliente.Style.Fill.BackgroundColor = XLColor.Green;
            rangoCliente.Style.Font.FontColor = XLColor.White;

            // Dar formato a los encabezados
            IXLRange rangoEncabezados = worksheet.Range(filaInicioDatos, 1, filaInicioDatos, ultimaColumnaUsada);
            rangoEncabezados.Style.Font.Bold = true;
            rangoEncabezados.Style.Fill.BackgroundColor = XLColor.LightGray;

            // Centrar el título
            IXLRange rangoTitulo = worksheet.Range(filaTitulo, 1, filaTitulo, ultimaColumnaUsada);
            rangoTitulo.Merge();
            rangoTitulo.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            // Autoajustar el ancho de las columnas
            worksheet.Columns().AdjustToContents();

            // Guardar el workbook en el MemoryStream
            workbook.SaveAs(memoryStream);
        }
#pragma warning restore IDE0007 // Use implicit type

        // Convertir el MemoryStream a Base64
        byte[] bytesArchivo = memoryStream.ToArray();
        string base64String = Convert.ToBase64String(bytesArchivo);

        return base64String;
    }
}
