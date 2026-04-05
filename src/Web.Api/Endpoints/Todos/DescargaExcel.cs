using System.Globalization;
using System.Linq.Expressions;
using System.Reflection;
using Application.Abstractions.Messaging;
using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;

namespace Web.Api.Endpoints.Todos;

internal sealed class DescargaExcel : IEndpoint
{
    // ==================== CLASES DE PRUEBA ====================

    // Clase Factura (original)
    public sealed class Factura
    {
        public string NumeroFactura { get; set; } = string.Empty;
        public decimal Monto { get; set; }
        public string Sucursal { get; set; } = string.Empty;
        public DateTime Fecha { get; set; }
        public string Cliente { get; set; } = string.Empty;
        public string Estado { get; set; } = "Pagado";
    }

    // ✅ NUEVA CLASE: Cliente
    public sealed class Cliente
    {
        public int Id { get; set; }
        public string Nombre { get; set; } = string.Empty;
        public string Apellido { get; set; } = string.Empty;
        public string Email { get; set; } = string.Empty;
        public string Telefono { get; set; } = string.Empty;
        public string Ciudad { get; set; } = string.Empty;
        public DateTime FechaRegistro { get; set; }
        public bool Activo { get; set; }
        public decimal Deuda { get; set; }
    }

    // ✅ NUEVA CLASE: Producto
    public sealed class Producto
    {
        public int Codigo { get; set; }
        public string Nombre { get; set; } = string.Empty;
        public string Categoria { get; set; } = string.Empty;
        public decimal Precio { get; set; }
        public int Stock { get; set; }
        public DateTime FechaVencimiento { get; set; }
        public string Proveedor { get; set; } = string.Empty;
        public bool Disponible { get; set; }
    }

    // ✅ NUEVA CLASE: Empleado (para más pruebas)
    public sealed class Empleado
    {
        public string Cedula { get; set; } = string.Empty;
        public string Nombres { get; set; } = string.Empty;
        public string Apellidos { get; set; } = string.Empty;
        public string Cargo { get; set; } = string.Empty;
        public decimal Salario { get; set; }
        public DateTime FechaIngreso { get; set; }
        public string Departamento { get; set; } = string.Empty;
    }


    private byte[] GenerarExcel<T>(List<T> listaDatos, string titulo)
    {
        using var memoryStream = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            IXLWorksheet worksheet = workbook.Worksheets.Add("Reporte");

            //  INSERTAR EL LOGO 
            byte[]? logoBytes = CargarLogoDesdeRecurso();
            int filaTitulo = 1;

            if (logoBytes != null && logoBytes.Length > 0)
            {
                try
                {
                    using var logoStream = new MemoryStream(logoBytes);
                    IXLPicture picture = worksheet.AddPicture(logoStream)
                                                   .MoveTo(worksheet.Cell("A1"))
                                                   .Scale(0.5);
                    picture.WithSize(500, 100);
                    filaTitulo = 6;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error al insertar logo: {ex.Message}");
                    filaTitulo = 1;
                }
            }

            // --- B. CREAR EL TÍTULO ---
            IXLCell celdaTitulo = worksheet.Cell(filaTitulo, 1);
            celdaTitulo.Value = titulo;
            celdaTitulo.Style.Font.Bold = true;
            celdaTitulo.Style.Font.FontSize = 18;
            celdaTitulo.Style.Font.FontColor = XLColor.DarkBlue;

            int filaInicioDatos = filaTitulo + 2;

            // --- C. OBTENER PROPIEDADES DINÁMICAMENTE ---
            PropertyInfo[] propiedades = typeof(T).GetProperties();

            // --- D. ENCABEZADOS DINÁMICOS CON ESTILOS ---
            for (int col = 0; col < propiedades.Length; col++)
            {
                IXLCell cabezera = worksheet.Cell(filaInicioDatos, col + 1);
                cabezera.Value = propiedades[col].Name;
                cabezera.Style.Font.Bold = true;
                cabezera.Style.Fill.BackgroundColor = XLColor.Green;
                cabezera.Style.Font.FontColor = XLColor.White;
                cabezera.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                cabezera.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            }

            // --- E. INSERTAR DATOS DINÁMICAMENTE ---
            /*Crear getters compilados, Esto convierte tus propiedades en funciones rápidas*/
            Func<T, object>[] getters = [.. propiedades.Select(prop =>
            {
                ParameterExpression param = Expression.Parameter(typeof(T), "x");
                MemberExpression property = Expression.Property(param, prop);
                UnaryExpression convert = Expression.Convert(property, typeof(object));

                return Expression.Lambda<Func<T, object>>(convert, param).Compile();
            })];

            int fila = 0;

            foreach (T? item in listaDatos)
            {
                int filaActual = filaInicioDatos + 1 + fila;

                for (int col = 0; col < getters.Length; col++)
                {
                    object valor = getters[col](item);
                    IXLCell celda = worksheet.Cell(filaActual, col + 1);

                    if (valor == null)
                    {
                        celda.Value = "";
                        continue;
                    }

                    switch (valor)
                    {
                        case DateTime fecha:
                            celda.Value = fecha;
                            celda.Style.DateFormat.Format = "yyyy-MM-dd";
                            break;

                        case decimal dec:
                            celda.Value = dec;
                            celda.Style.NumberFormat.Format = "$#,##0.00";
                            break;

                        case bool bol:
                            celda.Value = bol ? "Si" : "No";
                            break;

                        default:
                            celda.Value = (XLCellValue)valor;
                            break;
                    }
                }

                fila++;
            }

            // --- F. APLICAR FORMATO A DATOS ---
            int ultimaFilaUsada = filaInicioDatos + listaDatos.Count;
            int ultimaColumnaUsada = propiedades.Length;

            if (listaDatos.Count > 0)
            {
                IXLRange rangoDatos = worksheet.Range(filaInicioDatos + 1, 1, ultimaFilaUsada, ultimaColumnaUsada);

                rangoDatos.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                rangoDatos.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            }

            // --- H. CENTRAR TÍTULO ---
            IXLRange rangoTitulo = worksheet.Range(filaTitulo, 1, filaTitulo, ultimaColumnaUsada);
            rangoTitulo.Merge();
            rangoTitulo.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            rangoTitulo.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            // --- I. AJUSTAR ANCHO DE COLUMNAS ---
            worksheet.Columns().AdjustToContents();


            workbook.SaveAs(memoryStream);
        }

        return memoryStream.ToArray();
    }

    // ==================== MÉTODO PARA CARGAR LOGO ====================

    private byte[]? CargarLogoDesdeRecurso()
    {
        try
        {
            var assembly = Assembly.GetExecutingAssembly();
            string[] resourceNames = assembly.GetManifestResourceNames();

            string? logoResource = resourceNames.FirstOrDefault(r =>
                r.Contains("logo", StringComparison.OrdinalIgnoreCase) &&
                r.EndsWith(".png", StringComparison.OrdinalIgnoreCase));

            if (logoResource != null)
            {
                using Stream? stream = assembly.GetManifestResourceStream(logoResource);
                if (stream != null)
                {
                    using var memoryStream = new MemoryStream();
                    stream.CopyTo(memoryStream);
                    return memoryStream.ToArray();
                }
            }
            return null;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error al cargar logo: {ex.Message}");
            return null;
        }
    }

    // ==================== ENDPOINTS DE PRUEBA ====================

    public void MapEndpoint(IEndpointRouteBuilder app)
    {
        // Endpoint  para Facturas
        app.MapGet("todos/ExcelDescargar", () =>
        {
            try
            {
                var listaFacturasPrueba = new List<Factura>
                {
                    new Factura
                    {
                        NumeroFactura = "F001-001",
                        Monto = 1500.50m,
                        Sucursal = "Casa Matriz",
                        Fecha = DateTime.Now,
                        Cliente = "Juan Pérez",
                        Estado = "Pagado"
                    },
                    new Factura
                    {
                        NumeroFactura = "F001-002",
                        Monto = 2500.75m,
                        Sucursal = "Sucursal Norte",
                        Fecha = DateTime.Now.AddDays(-1),
                        Cliente = "María García",
                        Estado = "Pendiente"
                    },
                    new Factura
                    {
                        NumeroFactura = "F001-003",
                        Monto = 800.00m,
                        Sucursal = "Sucursal Sur",
                        Fecha = DateTime.Now.AddDays(-2),
                        Cliente = "Carlos López",
                        Estado = "Pagado"
                    }
                };

                byte[] archivoExcel = GenerarExcel(listaFacturasPrueba, "REPORTE DE FACTURACIÓN");
                return Results.File(archivoExcel, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $"reporte_facturas_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
            }
            catch (Exception ex)
            {
                return Results.BadRequest(new { error = $"Error: {ex.Message}" });
            }
        }).WithTags(Tags.Todos);

        // Descargar reporte de CLIENTES
        app.MapGet("todos/ExcelClientes", () =>
        {
            try
            {
                var listaClientes = new List<Cliente>
                {
                    new Cliente
                    {
                        Id = 1,
                        Nombre = "Juan",
                        Apellido = "Pérez",
                        Email = "juan.perez@email.com",
                        Telefono = "0991234567",
                        Ciudad = "Quito",
                        FechaRegistro = DateTime.Now.AddMonths(-3),
                        Activo = true,
                        Deuda = 150.75m
                    },
                    new Cliente
                    {
                        Id = 2,
                        Nombre = "María",
                        Apellido = "Gómez",
                        Email = "maria.gomez@email.com",
                        Telefono = "0997654321",
                        Ciudad = "Guayaquil",
                        FechaRegistro = DateTime.Now.AddMonths(-6),
                        Activo = true,
                        Deuda = 0m
                    },
                    new Cliente
                    {
                        Id = 3,
                        Nombre = "Carlos",
                        Apellido = "López",
                        Email = "carlos.lopez@email.com",
                        Telefono = "0999876543",
                        Ciudad = "Cuenca",
                        FechaRegistro = DateTime.Now.AddMonths(-1),
                        Activo = false,
                        Deuda = 2500.00m
                    },
                    new Cliente
                    {
                        Id = 4,
                        Nombre = "Ana",
                        Apellido = "Martínez",
                        Email = "ana.martinez@email.com",
                        Telefono = "0994567890",
                        Ciudad = "Quito",
                        FechaRegistro = DateTime.Now.AddDays(-15),
                        Activo = true,
                        Deuda = 75.50m
                    }
                };

                byte[] archivoExcel = GenerarExcel(listaClientes, "REPORTE DE CLIENTES");
                return Results.File(archivoExcel, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $"reporte_clientes_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
            }
            catch (Exception ex)
            {
                return Results.BadRequest(new { error = $"Error: {ex.Message}" });
            }
        }).WithTags(Tags.Todos);

        // ✅ NUEVO ENDPOINT: Descargar reporte de PRODUCTOS
        app.MapGet("todos/ExcelProductos", () =>
        {
            try
            {
                var listaProductos = new List<Producto>
                {
                    new Producto
                    {
                        Codigo = 1001,
                        Nombre = "Laptop HP",
                        Categoria = "Electrónica",
                        Precio = 899.99m,
                        Stock = 15,
                        FechaVencimiento = DateTime.Now.AddYears(2),
                        Proveedor = "HP Inc.",
                        Disponible = true
                    },
                    new Producto
                    {
                        Codigo = 1002,
                        Nombre = "Mouse Logitech",
                        Categoria = "Periféricos",
                        Precio = 25.50m,
                        Stock = 50,
                        FechaVencimiento = DateTime.Now.AddYears(1),
                        Proveedor = "Logitech",
                        Disponible = true
                    },
                    new Producto
                    {
                        Codigo = 1003,
                        Nombre = "Monitor Samsung",
                        Categoria = "Electrónica",
                        Precio = 299.99m,
                        Stock = 8,
                        FechaVencimiento = DateTime.Now.AddYears(3),
                        Proveedor = "Samsung",
                        Disponible = true
                    },
                    new Producto
                    {
                        Codigo = 1004,
                        Nombre = "Teclado Mecánico",
                        Categoria = "Periféricos",
                        Precio = 75.00m,
                        Stock = 0,
                        FechaVencimiento = DateTime.Now.AddYears(1),
                        Proveedor = "Redragon",
                        Disponible = false
                    }
                };

                byte[] archivoExcel = GenerarExcel(listaProductos, "REPORTE DE PRODUCTOS");
                return Results.File(archivoExcel, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $"reporte_productos_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
            }
            catch (Exception ex)
            {
                return Results.BadRequest(new { error = $"Error: {ex.Message}" });
            }
        }).WithTags(Tags.Todos);

        // ✅ NUEVO ENDPOINT: Descargar reporte de EMPLEADOS
        app.MapGet("todos/ExcelEmpleados", () =>
        {
            try
            {
                var listaEmpleados = new List<Empleado>
                {
                    new Empleado
                    {
                        Cedula = "1234567890",
                        Nombres = "Juan Carlos",
                        Apellidos = "Pérez Rodríguez",
                        Cargo = "Gerente",
                        Salario = 2500.00m,
                        FechaIngreso = DateTime.Now.AddYears(-5),
                        Departamento = "Administración"
                    },
                    new Empleado
                    {
                        Cedula = "0987654321",
                        Nombres = "María Fernanda",
                        Apellidos = "Gómez López",
                        Cargo = "Vendedor",
                        Salario = 1200.00m,
                        FechaIngreso = DateTime.Now.AddYears(-2),
                        Departamento = "Ventas"
                    },
                    new Empleado
                    {
                        Cedula = "1122334455",
                        Nombres = "Carlos Andrés",
                        Apellidos = "Martínez Silva",
                        Cargo = "Desarrollador",
                        Salario = 1800.00m,
                        FechaIngreso = DateTime.Now.AddMonths(-8),
                        Departamento = "Tecnología"
                    }
                };

                byte[] archivoExcel = GenerarExcel(listaEmpleados, "REPORTE DE EMPLEADOS");
                return Results.File(archivoExcel, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $"reporte_empleados_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
            }
            catch (Exception ex)
            {
                return Results.BadRequest(new { error = $"Error: {ex.Message}" });
            }
        }).WithTags(Tags.Todos);
    }
}
