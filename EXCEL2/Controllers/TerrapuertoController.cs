using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using EXCEL2.Models;
using EXCEL2.Models.Listas;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.UI.WebControls;

namespace EXCEL2.Controllers
{
    public class TerrapuertoController : Controller
    {
        private AuraPortal_BPMSEntities db = new AuraPortal_BPMSEntities();


        //trasferencia
        [HttpGet]
        public ActionResult GenerarArchivo(int idPanel)
        {
            //return Json(idPanel);

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
                    byte[] excelBytes = getDataExcel(package, idPanel).GetAsByteArray();

                    return Json(Actualizar(excelBytes));


                    // Devolver el archivo Excel al cliente como descarga
                    return File(excelBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "archivo_excel.xlsx");
                }
            }
        }

        [HttpGet]
        public ActionResult GenerarArchivoComiciones(int idPanel)
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
                    byte[] excelBytes = Comisiones(package, idPanel).GetAsByteArray();

                    return Json(Actualizar(excelBytes));
                    // Devolver el archivo Excel al cliente como descarga
                    return File(excelBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "archivo_excel.xlsx");
                }
            }
        }

        private byte[] ObtenerDatosVarbinary()
        {
            var familia = db.AP_Dyn_Familias_27.OrderBy(x => x.C_Name).ToList().Last();
            var integrate = db.AP__DocIntegratedStorage.Where(x => x.ObjectTypeProcessId == 27).Where(x => x.ObjectTypeProcess == 10).Where(x => x.IntegrationObjectId == familia.ID).First();
            var integrateData = db.AP__DocIntegratedDataStorage.Where(x => x.AscDocId == integrate.ID).First();
            return integrateData.Content;
        }

        public bool Actualizar(byte[] excelBytes)
        {
            var familia = db.AP_Dyn_Familias_27.OrderBy(x => x.C_Name).ToList().Last();
            var integrate = db.AP__DocIntegratedStorage.Where(x => x.ObjectTypeProcessId == 27).Where(x => x.ObjectTypeProcess == 10).Where(x => x.IntegrationObjectId == familia.ID).First();
            var integrateData = db.AP__DocIntegratedDataStorage.Where(x => x.AscDocId == integrate.ID).First();
            integrateData.Content = excelBytes;
             db.Entry(integrateData).State = EntityState.Modified;
            db.SaveChanges();

            
            return true;
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


            var format = "yyyy-MM-dd HH:mm:ss";
            var fecha1 = DateTime.ParseExact(filtro.C3_A__PR3_FS_FECHA_INICIO.Value.ToString(format), format, new CultureInfo("en-US"));
            var fecha2 = DateTime.ParseExact(filtro.C3_A__PR3_FS_FECHA_FIN.Value.ToString(format), format, new CultureInfo("en-US"));

            //return package;



			var FechaInicio = fecha1.ToString("dd MMMM", new CultureInfo("es-ES"));
            var FechaFin = fecha2.ToString("dd MMMM yyyy", new CultureInfo("es-ES"));


            int filaActual = 8;
            int columnaIndex = 1;

            
                while (hoja1.Cells[filaActual, columnaIndex].Value != null)
                {
                    // Mover a la siguiente fila
                    filaActual++;
                }
            


            int numeroEntero = (int)(filtro.C3_PR3_ITEM_TRANSFERENCIA);

            int nuevaFila = filaActual;
            var familia = db.AP_Dyn_Familias_27.OrderBy(x => x.C_Name).ToList().Last();
            
            List<ExcelMes> meses = mesesList(Int32.Parse(familia.C_Name));

            var lista = db.Panel_1008_4700.Where(x => x.C_ElementID == idPanel).ToList();


            var gruposPorMes = lista.GroupBy(d => new { Mes = d.C3_FE_FECHA_COMPROBANTE_RT.Value.Month, Anio = d.C3_FE_FECHA_COMPROBANTE_RT.Value.Year }).Select(g => new
            {
                Anio = g.Key.Anio,
                Mes = g.Key.Mes,
                Suma = g.Sum(d => d.C3_ND_PEAJE_TOTAL_RT)
            });

            var gruposPorAnio = lista.GroupBy(d => new { Anio = d.C3_FE_FECHA_COMPROBANTE_RT.Value.Year })
                .Select(g => new
                {
                    Anio = g.Key.Anio,
                })
                .ToList();
            var startCell = hoja1.Cells["A"+ (nuevaFila)];
            var endCell = hoja1.Cells["U"+(nuevaFila)];
            var range = hoja1.Cells[startCell.Address + ":" + endCell.Address];
            bool insertarFila = false;
            foreach (var cell in range)
            {
                // Verifica si la celda tiene contenido
                if (cell.Text != "")
                {
                    insertarFila = true;
                    break;
                }
            }
            if (insertarFila)
            {

            }
            else
            {
                hoja1.InsertRow(nuevaFila, 1);
            }
            
            hoja1.Cells["B" + nuevaFila].Value = "Del " + FechaInicio + " al " + FechaFin;
            hoja1.Cells["D" + nuevaFila].Value = filtro.C8_A__PR3_ABONO_TRANSITO_3_T_FECHA_Fecha_tranfs_min.Value.ToString("dd/MM/yyyy");
            hoja1.Cells["A" + nuevaFila].Value = numeroEntero.ToString("00");
            hoja1.Cells["A" + nuevaFila.ToString()].Style.Font.Bold = true; // Poner en negrita
            hoja1.Cells["A" + nuevaFila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // Centrar el contenido
            foreach (var item2 in meses)
            {
                if (item2.MesNumero == fecha1.Month.ToString("00") && item2.Anio == fecha1.Year)
                {
                    hoja1.Cells[item2.Posicion + nuevaFila.ToString()].Value = filtro.C8_A__PR3_LIQUIDACION_RESUMEN_DETALL_3_TL_MONTO_LIQ_Total_Monto;
                    hoja1.Cells[item2.Posicion + nuevaFila.ToString()].Style.Numberformat.Format = "\"S/\"#,##0.00";
                }


                if (filtro.C3_PR3_REGULARIZAR_MONTO != null)
                {
                    if (item2.MesNumero == filtro.C3_PR3_REGULARIZAR_MES.Value.ToString("00") &&  item2.Anio == filtro.C3_PR3_REGULARIZAR_AÑO)
                    {
                        //hoja1.InsertRow(nuevaFila+1, 1);
                        hoja1.Cells[item2.Posicion + (nuevaFila ).ToString()].Value = filtro.C3_PR3_REGULARIZAR_MONTO;
                        hoja1.Cells[item2.Posicion + (nuevaFila ).ToString()].Style.Numberformat.Format = "\"S/\"#,##0.00";
                    }
                    
                }
               
                
                    
            }

            hoja1.Cells["E" + nuevaFila].Value = filtro.C8_A__PR3_ABONO_TRANSITO_3_T_MONTO_BONO_T_Sumatoria;//Total Transferido            
            hoja1.Cells["E" + nuevaFila].Style.Numberformat.Format = "\"S/\"#,##0.00";
            if (filtro.C3_PR3_REGULARIZAR_OTROS_DESCUENTOS != null)
            {
                
                hoja1.Cells["F" + nuevaFila].Value = filtro.C3_PR3_REGULARIZAR_OTROS_DESCUENTOS;//Descuento
                hoja1.Cells["F" + nuevaFila].Style.Numberformat.Format = "\"S/\"#,##0.00";
            }
            
            return package;
        }

        public List<ExcelMes> mesesList(int anio)
        {
            List<ExcelMes> meses = new List<ExcelMes>();

            var mes12 = new ExcelMes();
            mes12.MesNombre = "Dicembre";
            mes12.MesNumero = "12";
            mes12.Anio = anio;
            mes12.Posicion = "G";
            meses.Add(mes12);

            var mes11 = new ExcelMes();
            mes11.MesNombre = "Noviembre";
            mes11.MesNumero = "11";
            mes11.Anio = anio;
            mes11.Posicion = "H";
            meses.Add(mes11);

            var mes10 = new ExcelMes();
            mes10.MesNombre = "Octubre";
            mes10.MesNumero = "10";
            mes10.Anio = anio;
            mes10.Posicion = "I";
            meses.Add(mes10);

            var mes9 = new ExcelMes();
            mes9.MesNombre = "Septiembre";
            mes9.MesNumero = "09";
            mes9.Anio = anio;
            mes9.Posicion = "J";
            meses.Add(mes9);

            var mes8 = new ExcelMes();
            mes8.MesNombre = "Agosto";
            mes8.MesNumero = "08";
            mes8.Anio = anio;
            mes8.Posicion = "K";
            meses.Add(mes8);

            var mes7 = new ExcelMes();

            mes7.MesNombre = "Julio";
            mes7.MesNumero = "07";
            mes7.Anio = anio;
            mes7.Posicion = "L";
            meses.Add(mes7);

            var mes6 = new ExcelMes();
            mes6.MesNombre = "Junio";
            mes6.MesNumero = "06";
            mes6.Anio = anio;
            mes6.Posicion = "M";
            meses.Add(mes6);

            var mes5 = new ExcelMes
            {
                MesNombre = "Mayo",
                MesNumero = "05",
                Anio = anio,
                Posicion = "N"
            };
            meses.Add(mes5);

            var mes4 = new ExcelMes
            {
                MesNombre = "Abril",
                MesNumero = "04",
                Anio = anio,
                Posicion = "O"
            };
            meses.Add(mes4);

            var mes3 = new ExcelMes
            {
                MesNombre = "Marzo",
                MesNumero = "03",
                Anio = anio,
                Posicion = "P"
            };
            meses.Add(mes3);

            var mes2 = new ExcelMes
            {
                MesNombre = "Febrero",
                MesNumero = "02",
                Anio = anio,
                Posicion = "Q"
            };
            meses.Add(mes2);

            var mes1 = new ExcelMes
            {
                MesNombre = "Enero",
                MesNumero = "01",
                Anio = anio ,
                Posicion = "R"
            };
            meses.Add(mes1);

            var mes0 = new ExcelMes
            {
                MesNombre = "Diciembre",
                MesNumero = "01",
                Anio = anio-1,
                Posicion = "S"
            };
            meses.Add(mes0);

            return meses;
        }

        public ExcelPackage Comisiones(ExcelPackage package, int idPanel)
        {

            List<string> nombresDeHojas = new List<string>();
            foreach (var worksheet in package.Workbook.Worksheets)
            {
                nombresDeHojas.Add(worksheet.Name);
            }
            ExcelWorksheet hoja1 = package.Workbook.Worksheets[nombresDeHojas[0]];

            int filaActual = 6;

            while (hoja1.Cells[filaActual, 1].Value != null)
            {
                filaActual++;
            }

            var startCell = hoja1.Cells["G" + (filaActual)];
            var endCell = hoja1.Cells["U" + (filaActual)];
            var range = hoja1.Cells[startCell.Address + ":" + endCell.Address];
            bool insertarFila = false;
            foreach (var cell in range)
            {
                // Verifica si la celda tiene contenido
                if (cell.Text != "")
                {
                    insertarFila = true;
                    break;
                }
            }

            if (insertarFila)
            {
                filaActual++;
            }
            var familia = db.AP_Dyn_Familias_27.OrderBy(x => x.C_Name).ToList().Last();

            List<ExcelMes> meses = mesesList(Int32.Parse(familia.C_Name));

            var Comision = db.Panel_1009.Where(x => x.C3_PR3_ID_PANEL == idPanel).ToList();

            foreach (var item in Comision)
            {
                foreach (var item2 in meses)
                {
                    if (item2.MesNumero == item.C3_PR3_MES_TRANSFERENCIA.Value.ToString("00"))
                    {
                        if (item2.Anio == item.C3_PR3_EN_AÑO)
                        {
                            hoja1.Cells[item2.Posicion + '6'].Value = item.C3_PR3_COM_FA_RECAUDACION;
                            hoja1.Cells[item2.Posicion + '6'].Style.Numberformat.Format = "\"S/\"#,##0.00";
                            //string formula_1 = "=SUM(" + item2.Posicion + "7:" + item2.Posicion.ToString() + (filaActual) + ")";
                            //hoja1.Cells[item2.Posicion + (filaActual + 2)].Formula = formula_1;
                            //hoja1.Cells[item2.Posicion + (filaActual + 3)].Formula = "=" + item2.Posicion + (filaActual + 2) + "-" + item2.Posicion + "6";
                        }
                    }
                }
            }

            return package;
        }


    }
}