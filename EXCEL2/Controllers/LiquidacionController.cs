using DocumentFormat.OpenXml.Packaging;
using EXCEL2.Models;
using EXCEL2.Models.Listas;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.UI.WebControls;
using OfficeOpenXml.Style;
using System.Data.Entity;
namespace EXCEL2.Controllers
{
    public class LiquidacionController : Controller
    {

        private AuraPortal_BPMSEntities db = new AuraPortal_BPMSEntities();

        [HttpGet]
        public ActionResult GenerarArchivo(int idPanel)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            // Crear un nuevo paquete Excel
            byte[] bytesDelExcel = ObtenerDatosVarbinary(); // Reemplaza esto con tus propios datos

            // Crear un MemoryStream a partir de los bytes del Excel
            using (MemoryStream memoryStream = new MemoryStream(bytesDelExcel))
            {
                // Crear un paquete Excel a partir del MemoryStream
                using (ExcelPackage package = new ExcelPackage(memoryStream))
                {

                    // Obtener el contenido del paquete en un array de bytes
                    byte[] excelBytes = getDataExcel(package,idPanel).GetAsByteArray();

                    return Json(Actualizar(excelBytes));

                    // Devolver el archivo Excel al cliente como descarga
                    return File(excelBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "archivo_excel.xlsx");
                }
            }

        }

        public bool Actualizar(byte[] excelBytes)
        {
            var familia = db.AP_Dyn_Familias_26.OrderBy(x => x.C_Name).ToList().Last();
            var integrate = db.AP__DocIntegratedStorage.Where(x => x.ObjectTypeProcessId == 26).Where(x => x.ObjectTypeProcess == 10).Where(x => x.IntegrationObjectId == familia.ID).First();
            var integrateData = db.AP__DocIntegratedDataStorage.Where(x => x.AscDocId == integrate.ID).First();
            integrateData.Content = excelBytes;
            db.Entry(integrateData).State = EntityState.Modified;
            db.SaveChanges();


            return true;
        }

        private byte[] ObtenerDatosVarbinary()
        {
            var familia = db.AP_Dyn_Familias_26.OrderBy( x => x.C_Name).ToList().Last();
            var integrate = db.AP__DocIntegratedStorage.Where( x => x.ObjectTypeProcessId == 26).Where(x => x.ObjectTypeProcess == 10).Where(x => x.IntegrationObjectId == familia.ID).First();
            var integrateData = db.AP__DocIntegratedDataStorage.Where(x => x.AscDocId == integrate.ID).First();
            return  integrateData.Content;
        }


        
        public ExcelPackage getDataExcel(ExcelPackage package, int idPanel)
        {
            List<string> nombresDeHojas = new List<string>();
            foreach (var worksheet in package.Workbook.Worksheets)
            {
                nombresDeHojas.Add(worksheet.Name);
            }
            ExcelWorksheet hoja1 = package.Workbook.Worksheets[nombresDeHojas[0]];

            var filtro = db.Panel_1008.Where(x => x.C3_PR3_ID_PANEL == idPanel).First();


            var format = "MM-dd-yyyy HH:mm:ss";
            var fecha1 = DateTime.ParseExact(filtro.C3_A__PR3_FS_FECHA_INICIO.Value.ToString(format), format, new CultureInfo("en-US"));
            var fecha2 = DateTime.ParseExact(filtro.C3_A__PR3_FS_FECHA_FIN.Value.ToString(format), format, new CultureInfo("en-US"));


            var FechaInicio = fecha1.ToString("dd MMMM", new CultureInfo("es-ES"));
            var FechaFin = fecha2.ToString("dd MMMM yyyy", new CultureInfo("es-ES"));
            

            int filaActual = 6;
            int columnaIndex = 1;
            string valorCelda = "00";

            while (hoja1.Cells[filaActual, columnaIndex].Value != null)
            {
                // Obtener el valor de la celda actual
                valorCelda = (string)hoja1.Cells[filaActual, columnaIndex].Value;


                // Mover a la siguiente fila
                filaActual++;
            }


            int numeroEntero = (int)(filtro.C3_PR3_ITEM_TRANSFERENCIA);
            int nuevaFila = filaActual;


            var familia = db.AP_Dyn_Familias_26.OrderBy(x => x.C_Name).ToList().Last();
            var anio = Int32.Parse(familia.C_Name);
            List<ExcelMes> meses = mesesList(anio);
            List<Panel_1008_4700> lista;

            if(fecha1.Month == 1)
            {
                lista = db.Panel_1008_4700.Where(x => x.C_ElementID == idPanel).ToList();
            }
            else
            {
                 lista = db.Panel_1008_4700.Where(x => x.C_ElementID == idPanel).ToList();
            }
            


            var gruposPorMes = lista.GroupBy(d => new { Mes = d.C3_FE_FECHA_COMPROBANTE_RT.Value.Month , Anio= d.C3_FE_FECHA_COMPROBANTE_RT.Value.Year }).Select(g => new
            {
                Anio = g.Key.Anio,
                Mes = g.Key.Mes,
                Suma = g.Sum(d => d.C3_ND_PEAJE_TOTAL_RT)
            }).ToArray();

            

            

            var gruposPorAnio = lista.GroupBy(d => new {  Anio = d.C3_FE_FECHA_COMPROBANTE_RT.Value.Year })
                .Select(g => new
                {
                    Anio = g.Key.Anio,
                })
                .ToList();

            var ultimoMes = gruposPorMes.Last().Mes;
            var ultimoAnio  = gruposPorMes.Last().Anio;

            //int anioAnterior = 0;
            hoja1.InsertRow(nuevaFila, 1);
            hoja1.Cells["B" + nuevaFila].Value = "Del " + FechaInicio + " al " + FechaFin;
            hoja1.Cells["D" + nuevaFila].Value = filtro.C8_A__PR3_ABONO_TRANSITO_3_T_FECHA_Fecha_tranfs_min.Value.ToString("dd/MM/yyyy");
            hoja1.Cells["A" + nuevaFila].Value = numeroEntero.ToString("00");
            hoja1.Cells["A" + nuevaFila.ToString()].Style.Font.Bold = true; // Poner en negrita
            hoja1.Cells["A" + nuevaFila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // Centrar el contenido

            foreach (var grupo in gruposPorMes)
            {
                foreach (var item2 in meses)
                {
                    if (item2.MesNumero == grupo.Mes.ToString("00") && item2.Anio == grupo.Anio)
                    {
                        
                        
                        hoja1.Cells[item2.Posicion + nuevaFila.ToString()].Value = grupo.Suma;
                        hoja1.Cells[item2.Posicion + nuevaFila.ToString()].Style.Numberformat.Format = "\"S/\"#,##0.00";
                    }
                }
            }

            hoja1.Cells["E" + nuevaFila].Value = filtro.C3_A__PR3_TOTAL_TRANSFERIDO;//Total Transferido            
            hoja1.Cells["E" + nuevaFila].Style.Numberformat.Format = "\"S/\"#,##0.00";
            if(filtro.C8_A__PR3_OTROS_DESCUENTOS_3_T_MONTO_Total_monto != null)
            {
                hoja1.Cells["F" + nuevaFila].Value = -filtro.C8_A__PR3_OTROS_DESCUENTOS_3_T_MONTO_Total_monto;//Descuento
                hoja1.Cells["F" + nuevaFila].Style.Numberformat.Format = "\"S/\"#,##0.00";
            }
            

            //Hoja 2-------------------------

            ExcelWorksheet hoja2 = package.Workbook.Worksheets[nombresDeHojas[1]];

            int filaActualHoja2 = 6;
            int columnaIndexHoja2 = 1;
            string valorCeldaHoja2 = "";
            while (hoja2.Cells[filaActualHoja2, columnaIndexHoja2].Value != null)
            {
                // Obtener el valor de la celda actual
                valorCeldaHoja2 = (string)hoja2.Cells[filaActualHoja2, columnaIndexHoja2].Value;


                // Mover a la siguiente fila
                filaActualHoja2++;
            }
            int nuevaFilaHoja2 = filaActualHoja2;
            int numeroEnterohoja2 = (int)(filtro.C3_PR3_ITEM_TRANSFERENCIA);

            var gruposPorMesHoja2 = lista.GroupBy(d => new { Mes = d.C3_FE_FECHA_COMPROBANTE_RT.Value.Month, Anio = d.C3_FE_FECHA_COMPROBANTE_RT.Value.Year }).Select(g => new
            {
                Anio = g.Key.Anio,
                Mes = g.Key.Mes,
                Suma = g.Sum(d => d.C3_ND_DETRACCION_TOTAL_RT)
            });

            var gruposPorAnioHoja2 = lista.GroupBy(d => new { Anio = d.C3_FE_FECHA_COMPROBANTE_RT.Value.Year })
                .Select(g => new
                {
                    Anio = g.Key.Anio,
                })
                .ToList();

            hoja2.InsertRow(nuevaFilaHoja2, 1);
            hoja2.Cells["B" + nuevaFilaHoja2].Value = "Del " + FechaInicio + " al " + FechaFin;
            hoja2.Cells["D" + nuevaFilaHoja2].Value = filtro.C8_A__PR3_ABONO_DETRACCION_3_PR3_FECHA_DET_Fecha_tranfs_min.Value.ToString("dd/MM/yyyy");
            hoja2.Cells["A" + nuevaFilaHoja2].Value = numeroEnterohoja2.ToString("00");
            hoja2.Cells["A" + nuevaFilaHoja2.ToString()].Style.Font.Bold = true; // Poner en negrita
            hoja2.Cells["A" + nuevaFilaHoja2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // Centrar el contenido

            foreach (var grupo in gruposPorMesHoja2)
            {

                foreach (var item2 in meses)
                {
                    if (item2.MesNumero == grupo.Mes.ToString("00") && item2.Anio == grupo.Anio)
                    {

                        hoja2.Cells[item2.Posicion + nuevaFila.ToString()].Value = grupo.Suma;
                        hoja2.Cells[item2.Posicion + nuevaFilaHoja2.ToString()].Style.Numberformat.Format = "\"S/\"#,##0.00";

                    }
                }
            }
            hoja2.Cells["E" + nuevaFilaHoja2].Value = filtro.C3_A__PR3_MONTO_DETRACCIONES;//Total Transferido            
            hoja2.Cells["E" + nuevaFilaHoja2].Style.Numberformat.Format = "\"S/\"#,##0.00";
            if (filtro.C8_PR3_GC_OTROS_DESCUENTOS_DETRAC_3_T_MONTO_DET_Total_monto != null)
            {
                hoja2.Cells["F" + nuevaFilaHoja2].Value = -filtro.C8_PR3_GC_OTROS_DESCUENTOS_DETRAC_3_T_MONTO_DET_Total_monto;//Descuento
                hoja2.Cells["F" + nuevaFilaHoja2].Style.Numberformat.Format = "\"S/\"#,##0.00";
            }
            

            return package;
        }

        public List<ExcelMes> mesesList(int anio)
        {
            List<ExcelMes> meses = new List<ExcelMes>();

            var mes13 = new ExcelMes();
            mes13.MesNombre = "Enero";
            mes13.MesNumero = "01";
            mes13.Anio = anio+1;
            mes13.Posicion = "G";
            meses.Add(mes13);

            var mes12 = new ExcelMes();
            mes12.MesNombre = "Dicembre";
            mes12.MesNumero = "12";
            mes12.Anio = anio;
            mes12.Posicion = "H";
            meses.Add(mes12);

            var mes11 = new ExcelMes();
            mes11.MesNombre = "Noviembre";
            mes11.MesNumero = "11";
            mes11.Anio = anio;
            mes11.Posicion = "I";
            meses.Add(mes11);

            var mes10 = new ExcelMes();
            mes10.MesNombre = "Octubre";
            mes10.MesNumero = "10";
            mes10.Anio = anio;
            mes10.Posicion = "J";
            meses.Add(mes10);

            var mes9 = new ExcelMes();
            mes9.MesNombre = "Septiembre";
            mes9.MesNumero = "09";
            mes9.Anio = anio;
            mes9.Posicion = "K";
            meses.Add(mes9);

            var mes8 = new ExcelMes();
            mes8.MesNombre = "Agosto";
            mes8.MesNumero = "08";
            mes8.Anio = anio;
            mes8.Posicion = "L";
            meses.Add(mes8);

            var mes7 = new ExcelMes();

            mes7.MesNombre = "Julio";
            mes7.MesNumero = "07";
            mes7.Anio = anio;
            mes7.Posicion = "M";
            meses.Add(mes7);

            var mes6 = new ExcelMes();
            mes6.MesNombre = "Junio";
            mes6.MesNumero = "06";
            mes6.Anio = anio;
            mes6.Posicion = "N";
            meses.Add(mes6);

            var mes5 = new ExcelMes
            {
                MesNombre = "Mayo",
                MesNumero = "05",
                Anio = anio,
                Posicion = "O"
            };
            meses.Add(mes5);

            var mes4 = new ExcelMes
            {
                MesNombre = "Abril",
                MesNumero = "04",
                Anio = anio,
                Posicion = "P"
            };
            meses.Add(mes4);

            var mes3 = new ExcelMes
            {
                MesNombre = "Marzo",
                MesNumero = "03",
                Anio = anio,
                Posicion = "Q"
            };
            meses.Add(mes3);

            var mes2 = new ExcelMes
            {
                MesNombre = "Febrero",
                MesNumero = "02",
                Anio = anio,
                Posicion = "R"
            };
            meses.Add(mes2);

            var mes1 = new ExcelMes
            {
                MesNombre = "Enero",
                MesNumero = "01",
                Anio = anio,
                Posicion = "S"
            };
            meses.Add(mes1);

            var mes0 = new ExcelMes
            {
                MesNombre = "Diciembre",
                MesNumero = "12",
                Anio = anio-1,
                Posicion = "T"
            };
            meses.Add(mes0);

            return meses;
        }


        

    }
}
