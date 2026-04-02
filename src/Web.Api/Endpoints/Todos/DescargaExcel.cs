using System.Globalization;
using System.Reflection;
using Application.Abstractions.Messaging;
using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;

namespace Web.Api.Endpoints.Todos;

internal sealed class DescargaExcel : IEndpoint
{
    // Clase Factura (asegúrate de que esté definida)
    public sealed class Factura
    {
        public string NumeroFactura { get; set; } = string.Empty;
        public decimal Monto { get; set; }
        public string Sucursal { get; set; } = string.Empty;
        public DateTime Fecha { get; set; }
        public string Cliente { get; set; } = string.Empty;
    }

    // Método auxiliar para generar bytes del Excel con LOGO
    private byte[] GenerarExcelBytes(List<Factura> listaFacturas)
    {
        using var memoryStream = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            IXLWorksheet worksheet = workbook.Worksheets.Add("ReporteFacturas");

            // --- A. INSERTAR EL LOGO (desde recurso embebido) ---
            byte[]? logoBytes = CargarLogoDesdeRecurso();
            int filaTitulo = 1;

            if (logoBytes != null && logoBytes.Length > 0)
            {
                try
                {
                    using var logoStream = new MemoryStream(logoBytes);

                    // Agregar imagen desde el stream
#pragma warning disable S1481 // Unused local variables should be removed
                    IXLPicture picture = worksheet.AddPicture(logoStream)
                                                   .MoveTo(worksheet.Cell("A1"))
                                                   .Scale(0.5); // Ajusta escala según necesites (0.5 = 50%)
#pragma warning restore S1481 // Unused local variables should be removed

                    // Opcional: Ajustar tamaño específico
                    // picture.WithSize(200, 100); // Ancho 200px, Alto 100px

                    // El logo ocupa espacio, ajustamos el título
                    filaTitulo = 5; // El título irá en la fila 5 después del logo (dejamos espacio)

                    // Ajustar altura de la fila del logo
                    worksheet.Row(1).Height = 80; // Ajusta según tamaño de tu logo
                }
                catch (Exception ex)
                {
                    // Si falla la inserción del logo, continuamos sin él
                    Console.WriteLine($"Error al insertar logo: {ex.Message}");
                    filaTitulo = 1;
                }
            }

            // --- B. CREAR EL TÍTULO ---
            IXLCell celdaTitulo = worksheet.Cell(filaTitulo, 1);
            celdaTitulo.Value = "REPORTE DE FACTURACIÓN";
            celdaTitulo.Style.Font.Bold = true;
            celdaTitulo.Style.Font.FontSize = 18;
            celdaTitulo.Style.Font.FontColor = XLColor.DarkBlue;

            int filaInicioDatos = filaTitulo + 2;

            // --- C. ENCABEZADOS DE COLUMNAS ---
            worksheet.Cell(filaInicioDatos, 1).Value = "Número Factura";
            worksheet.Cell(filaInicioDatos, 2).Value = "Monto";
            worksheet.Cell(filaInicioDatos, 3).Value = "Sucursal";
            worksheet.Cell(filaInicioDatos, 4).Value = "Fecha";
            worksheet.Cell(filaInicioDatos, 5).Value = "Cliente";

            // Aplicar formato a encabezados
            for (int col = 1; col <= 5; col++)
            {
                IXLCell headerCell = worksheet.Cell(filaInicioDatos, col);
                headerCell.Style.Font.Bold = true;
                headerCell.Style.Fill.BackgroundColor = XLColor.LightGray;
                headerCell.Style.Font.FontColor = XLColor.Black;
                headerCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                headerCell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            }

            // --- D. INSERTAR DATOS ---
            for (int i = 0; i < listaFacturas.Count; i++)
            {
                int filaActual = filaInicioDatos + 1 + i;
                worksheet.Cell(filaActual, 1).Value = listaFacturas[i].NumeroFactura;
                worksheet.Cell(filaActual, 2).Value = listaFacturas[i].Monto;
                worksheet.Cell(filaActual, 3).Value = listaFacturas[i].Sucursal;
                worksheet.Cell(filaActual, 4).Value = listaFacturas[i].Fecha.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                worksheet.Cell(filaActual, 5).Value = listaFacturas[i].Cliente;
            }

            // --- E. OBTENER ÚLTIMAS FILAS Y COLUMNAS USADAS (MANEJO SEGURO DE NULL) ---
            int ultimaFilaUsada = worksheet.LastRowUsed()?.RowNumber() ?? filaInicioDatos + listaFacturas.Count;
            int ultimaColumnaUsada = worksheet.LastColumnUsed()?.ColumnNumber() ?? 5;

            // --- F. APLICAR FORMATO VERDE A TODAS LAS CELDAS DE DATOS ---
            if (ultimaFilaUsada >= filaInicioDatos + 1)
            {
                IXLRange rangoDatos = worksheet.Range(filaInicioDatos + 1, 1, ultimaFilaUsada, ultimaColumnaUsada);
                rangoDatos.Style.Fill.BackgroundColor = XLColor.Green;
                rangoDatos.Style.Font.FontColor = XLColor.White;
                rangoDatos.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                rangoDatos.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            }

            // --- G. FORMATO MONEDA PARA COLUMNA MONTO ---
            if (ultimaFilaUsada >= filaInicioDatos + 1)
            {
                IXLRange rangoMonto = worksheet.Range(filaInicioDatos + 1, 2, ultimaFilaUsada, 2);
                rangoMonto.Style.NumberFormat.Format = "$#,##0.00";
            }

            // --- H. CENTRAR TÍTULO (fusionar celdas) ---
            IXLRange rangoTitulo = worksheet.Range(filaTitulo, 1, filaTitulo, ultimaColumnaUsada);
            rangoTitulo.Merge();
            rangoTitulo.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            rangoTitulo.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            // --- I. AJUSTAR ANCHO DE COLUMNAS ---
            worksheet.Columns().AdjustToContents();

            // Ajustar ancho mínimo para algunas columnas
            worksheet.Column(1).Width = 18; // Número Factura
            worksheet.Column(2).Width = 15; // Monto
            worksheet.Column(3).Width = 20; // Sucursal
            worksheet.Column(4).Width = 12; // Fecha
            worksheet.Column(5).Width = 25; // Cliente

            workbook.SaveAs(memoryStream);
        }

        return memoryStream.ToArray();
    }

    // Método auxiliar para cargar el logo desde recursos embebidos
    private byte[]? CargarLogoDesdeRecurso()
    {
        try
        {
            var assembly = Assembly.GetExecutingAssembly();
            string[] resourceNames = assembly.GetManifestResourceNames();

            // 🔍 CÓDIGO DE DEPURACIÓN - MUESTRA TODOS LOS RECURSOS
            Console.WriteLine("=== RECURSOS DISPONIBLES EN EL ENSAMBLAJE ===");
            foreach (string resource in resourceNames)
            {
                Console.WriteLine($"- {resource}");
            }
            Console.WriteLine("============================================");

            // Buscar el logo con diferentes posibles nombres
            string? logoResource = resourceNames.FirstOrDefault(r =>
                r.Contains("logo", StringComparison.OrdinalIgnoreCase) &&
                (r.EndsWith(".png", StringComparison.OrdinalIgnoreCase) ||
                 r.EndsWith(".jpg", StringComparison.OrdinalIgnoreCase) ||
                 r.EndsWith(".jpeg", StringComparison.OrdinalIgnoreCase)));

            // Si no encuentra, intenta con nombres específicos
            logoResource ??= resourceNames.FirstOrDefault(r =>
                r.EndsWith(".png", StringComparison.OrdinalIgnoreCase) ||
                r.EndsWith(".jpg", StringComparison.OrdinalIgnoreCase));

            if (logoResource != null)
            {
                Console.WriteLine($"✅ Logo encontrado: {logoResource}");

                using Stream? stream = assembly.GetManifestResourceStream(logoResource);
                if (stream != null)
                {
                    using var memoryStream = new MemoryStream();
                    stream.CopyTo(memoryStream);
                    byte[] imageBytes = memoryStream.ToArray();

                    Console.WriteLine($"📊 Tamaño del logo: {imageBytes.Length} bytes");

                    if (imageBytes.Length > 0)
                    {
                        return imageBytes;
                    }
                    else
                    {
                        Console.WriteLine("❌ El logo está vacío (0 bytes)");
                    }
                }
                else
                {
                    Console.WriteLine("❌ No se pudo obtener el stream del logo");
                }
            }
            else
            {
                Console.WriteLine("❌ No se encontró ningún archivo de logo");
            }

            return null;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ Error al cargar logo: {ex.Message}");
            Console.WriteLine($"StackTrace: {ex.StackTrace}");
            return null;
        }
    }

    public void MapEndpoint(IEndpointRouteBuilder app)
    {
        app.MapGet("todos/ExcelDescargar", () =>
        {
            try
            {
                // Datos de prueba con fechas en formato correcto
                var listaFacturasPrueba = new List<Factura>
                {
                    new Factura
                    {
                        NumeroFactura = "F001-001",
                        Monto = 1500.50m,
                        Sucursal = "Casa Matriz",
                        Fecha = DateTime.Now,
                        Cliente = "Juan Pérez"
                    },
                    new Factura
                    {
                        NumeroFactura = "F001-002",
                        Monto = 2500.75m,
                        Sucursal = "Sucursal Norte",
                        Fecha = DateTime.Now.AddDays(-1),
                        Cliente = "María García"
                    },
                    new Factura
                    {
                        NumeroFactura = "F001-003",
                        Monto = 800.00m,
                        Sucursal = "Sucursal Sur",
                        Fecha = DateTime.Now.AddDays(-2),
                        Cliente = "Carlos López"
                    }
                };

                byte[] archivoExcel = GenerarExcelBytes(listaFacturasPrueba);

                // Validar que el archivo no esté vacío
                if (archivoExcel == null || archivoExcel.Length == 0)
                {
                    return Results.BadRequest(new { error = "Error al generar el archivo Excel" });
                }

                // Retorna el archivo directamente para descarga automática
                return Results.File(
                    archivoExcel,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    $"reporte_facturas_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
                );
            }
            catch (Exception ex)
            {
                return Results.BadRequest(new { error = $"Error al generar reporte: {ex.Message}" });
            }
        })
        .WithTags(Tags.Todos); // Si requieres autenticación
    }
}




//public class ExcelServicio : IExcel
//{
//    public Task<byte[]> Convertir(ExcelDto excel)
//    {
//        // Generar el Excel y obtener el Base64
//        return Task.FromResult(GenerarExcelEnBase64(excel));
//    }

//    private static byte[] GenerarExcelEnBase64(ExcelDto excel)
//    {
//        using MemoryStream memoryStream = new MemoryStream();
//        using XLWorkbook workbook = new XLWorkbook();
//        IXLWorksheet worksheet = workbook.Worksheets.Add("Reporte");

//        int filaActual = 1;

//        // 1️⃣ TÍTULO 

//        int filaTitulo = 4;
//        IXLCell celdaTitulo = worksheet.Cell(filaTitulo, 1);
//        celdaTitulo.Value = "REPORTE DE FACTURACIÓN";
//        celdaTitulo.Style.Font.Bold = true;
//        celdaTitulo.Style.Font.FontSize = 18;
//        celdaTitulo.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

//        // 2️⃣ ENCABEZADOS AUTOMÁTICOS
//        List<string> encabezados = excel.Filas.First().Keys.ToList();

//        for (int col = 0; col < encabezados.Count; col++)
//        {
//            worksheet.Cell(filaActual, col + 1).Value = encabezados[col];
//            worksheet.Cell(filaActual, col + 1).Style.Font.Bold = true;
//            worksheet.Cell(filaActual, col + 1).Style.Fill.BackgroundColor =
//                XLColor.LightGray;
//        }

//        int filaDatos = filaActual + 1;

//        // 3️⃣ DATOS
//        for (int i = 0; i < excel.Filas.Count; i++)
//        {
//            var fila = excel.Filas[i];

//            for (int col = 0; col < encabezados.Count; col++)
//            {
//                var key = encabezados[col];
//                worksheet.Cell(filaDatos + i, col + 1).Value =
//                    fila.ContainsKey(key) ? fila[key] : string.Empty;
//            }
//        }

//        worksheet.Columns().AdjustToContents();
//        workbook.SaveAs(memoryStream);

//        return memoryStream.ToArray();
//    }
//}
