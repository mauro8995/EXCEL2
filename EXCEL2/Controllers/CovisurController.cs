using EXCEL2.Models;
using EXCEL2.Models.Listas;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.UI.WebControls;

namespace EXCEL2.Controllers
{
    public class CovisurController : Controller
    {

        private AuraPortal_BPMSEntities db = new AuraPortal_BPMSEntities();

        [HttpGet]
        public ActionResult GenerarArchivoComisiones(int idPanel)
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

        [HttpGet]
        public ActionResult GenerarArchivoFactura(int idPanel)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            // Crear un nuevo paquete Excel
            byte[] bytesDelExcel = ObtenerDatosVarbinaryFactura(idPanel); // Reemplaza esto con tus propios datos

            // Crear un MemoryStream a partir de los bytes del Excel
            using (MemoryStream memoryStream = new MemoryStream(bytesDelExcel))
            {
                // Crear un paquete Excel a partir del MemoryStream
                using (ExcelPackage package = new ExcelPackage(memoryStream))
                {

                    // Obtener el contenido del paquete en un array de bytes
                    byte[] excelBytes = Facturas(package, idPanel).GetAsByteArray();

                    return Json(ActualizarFactura(excelBytes,idPanel));


                    // Devolver el archivo Excel al cliente como descarga
                    return File(excelBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "archivo_excel.xlsx");
                }
            }
        }
        [HttpGet]
        public ActionResult GenerarArchivoFacturaPaso2(int idPanel)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            // Crear un nuevo paquete Excel
            byte[] bytesDelExcel = ObtenerDatosVarbinaryFactura(idPanel); // Reemplaza esto con tus propios datos

            // Crear un MemoryStream a partir de los bytes del Excel
            using (MemoryStream memoryStream = new MemoryStream(bytesDelExcel))
            {
                // Crear un paquete Excel a partir del MemoryStream
                using (ExcelPackage package = new ExcelPackage(memoryStream))
                {

                    // Obtener el contenido del paquete en un array de bytes
                    byte[] excelBytes = Facturas_2(package, idPanel).GetAsByteArray();

                    return Json(ActualizarFactura(excelBytes,idPanel));


                    // Devolver el archivo Excel al cliente como descarga
                    return File(excelBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "archivo_excel.xlsx");
                }
            }
        }


        private byte[] ObtenerDatosVarbinary()
        {
            var familia = db.AP_Dyn_Familias_26.OrderBy(x => x.C_Name).ToList().Last();
            var integrate = db.AP__DocIntegratedStorage.Where(x => x.ObjectTypeProcessId == 26).Where(x => x.ObjectTypeProcess == 10).Where(x => x.IntegrationObjectId == familia.ID).First();
            var integrateData = db.AP__DocIntegratedDataStorage.Where(x => x.AscDocId == integrate.ID).First();
            return integrateData.Content;
        }

        private byte[] ObtenerDatosVarbinaryFactura(int idPanel)
        {
            var elemeto = db.Panel_1008.Where(x => x.ID == idPanel).First();
            var integrate = db.AP__DocIntegratedStorage.Where(x => x.IntegrationObjectId == elemeto.C_ElementID).Where(x => x.AscTermId == 2772).First();
            var integrateData = db.AP__DocIntegratedDataStorage.Where(x => x.AscDocId == integrate.ID).First();
            return integrateData.Content;
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

        public bool ActualizarFactura(byte[] excelBytes,int idPanel)
        {
            var elemeto = db.Panel_1008.Where(x => x.ID == idPanel).First();
            var integrate = db.AP__DocIntegratedStorage.Where(x => x.IntegrationObjectId == elemeto.C_ElementID).Where(x => x.AscTermId == 2772).First();
            var integrateData = db.AP__DocIntegratedDataStorage.Where(x => x.AscDocId == integrate.ID).First();
            integrateData.Content = excelBytes;
            db.Entry(integrateData).State = EntityState.Modified;
            db.SaveChanges();
            return true;

        }


        public List<ExcelMes> mesesList(int anio)
        {
            List<ExcelMes> meses = new List<ExcelMes>();

            var mes13 = new ExcelMes();
            mes13.MesNombre = "Enero";
            mes13.MesNumero = "01";
            mes13.Anio = anio + 1;
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
                Anio = anio - 1,
                Posicion = "T"
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

            var familia = db.AP_Dyn_Familias_26.OrderBy(x => x.C_Name).ToList().Last();
            var anio = Int32.Parse(familia.C_Name);
            List<ExcelMes> meses = mesesList(anio);

            var Comision = db.Panel_1009.Where(x => x.C3_PR3_ID_PANEL == idPanel).ToList();

            foreach (var item in Comision)
            {
                foreach (var item2 in meses)
                {
                    if (item2.MesNumero == item.C3_PR3_MES_TRANSFERENCIA.Value.ToString("00"))
                    {
                        if (DateTime.Now.Year == item.C3_PR3_EN_AÑO)
                        {
                            hoja1.Cells[item2.Posicion + '6'].Value = item.C3_PR3_COM_FA_RECAUDACION;
                            hoja1.Cells[item2.Posicion + (filaActual + 2)].Value = item.C3_PR3_COM_PA_RECAUDACION;
                            hoja1.Cells[item2.Posicion + (filaActual + 3)].Value = item.C3_PR3_COM_FINAL_DIFERENCIA;
                        }
                        else
                        {
                            hoja1.Cells["S" + '6'].Value = item.C3_PR3_COM_FINAL_DIFERENCIA;
                        }
                        hoja1.Cells[item2.Posicion + '6'].Style.Numberformat.Format = "\"S/\"#,##0.00";
                        hoja1.Cells[item2.Posicion + (filaActual + 2)].Style.Numberformat.Format = "\"S/\"#,##0.00";
                        hoja1.Cells[item2.Posicion + (filaActual + 3)].Style.Numberformat.Format = "\"S/\"#,##0.00";
                    }
                }
            }



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

            foreach (var item in Comision)
            {
                foreach (var item2 in meses)
                {
                    if (item2.MesNumero == item.C3_PR3_MES_TRANSFERENCIA.Value.ToString("00"))
                    {
                        if (DateTime.Now.Year == item.C3_PR3_EN_AÑO)
                        {
                            hoja2.Cells[item2.Posicion + '6'].Value = item.C8_GC_COMPROBANTES_TRANSITOS_3_ND_DETRACCION_TOTAL_RT_ND_TOTAL_DETRACCION;
                            hoja2.Cells[item2.Posicion + (filaActual + 2)].Value = item.C8_GC_COMPROBANTES_TRANSITOS_3_ND_DETRACCION_TOTAL_RT_ND_TOTAL_DETRACCION;
                            hoja2.Cells[item2.Posicion + (filaActual + 3)].Value = item.C3_PR3_COM_POR_TRANSFERIR_DET;
                        }
                        else
                        {
                            hoja2.Cells["S" + '6'].Value = item.C8_GC_COMPROBANTES_TRANSITOS_3_ND_DETRACCION_TOTAL_RT_ND_TOTAL_DETRACCION;
                        }
                        hoja2.Cells[item2.Posicion + '6'].Style.Numberformat.Format = "\"S/\"#,##0.00";
                        hoja2.Cells[item2.Posicion + (filaActual + 2)].Style.Numberformat.Format = "\"S/\"#,##0.00";
                        hoja2.Cells[item2.Posicion + (filaActual + 3)].Style.Numberformat.Format = "\"S/\"#,##0.00";
                    }
                }
            }

            return package;
        }

        public ExcelPackage Facturas(ExcelPackage package, int idPanel)
        {
            //decimal peajeTotal = decimal.Parse(700,01);
            var lista = db.Panel_1008_4700.Where(x => x.C_ElementID == idPanel).Where(x => x.C3_ND_PEAJE_TOTAL_RT > 700).Where(x => x.C3_TL_CUENTA_RT != "46841")
                .Where(x => x.C3_ND_DETRACCION_TOTAL_RT == 0 || x.C3_ND_DETRACCION_TOTAL_RT == null).ToList();

          

            List<string> nombresDeHojas = new List<string>();
            foreach (var worksheet in package.Workbook.Worksheets)
            {
                nombresDeHojas.Add(worksheet.Name);
            }
            ExcelWorksheet hoja1 = package.Workbook.Worksheets[nombresDeHojas[0]];

            int filaActual = 2;

            foreach (var item in lista)
            {
                hoja1.Cells["A" + filaActual].Value = item.C3_TL_SERIE_RT;
                hoja1.Cells["B" + filaActual].Value = item.C3_TL_CORRELATIVO_RT;
                hoja1.Cells["C" + filaActual].Value = item.C3_TL_TIPO_COMPROBANTE_RT;
                hoja1.Cells["D" + filaActual].Value = item.C3_TL_NOMBRE_CLIENTE_RT;
                hoja1.Cells["E" + filaActual].Value = item.C3_TL_CUENTA_RT;
                filaActual++;
            }

            return package;
        }

        public ExcelPackage Facturas_2(ExcelPackage package, int idPanel)
        {
            //decimal peajeTotal = decimal.Parse(700,01);
            var lista = db.Panel_1008_4700.Where(x => x.C_ElementID == idPanel).Where(x => x.C3_ND_PEAJE_TOTAL_RT > 700).Where(x => x.C3_TL_CUENTA_RT != "46841")
                .Where(x => x.C3_TL_PR3_OBSERVACION_Y == 2329)
                .Where(x => x.C3_ND_DETRACCION_TOTAL_RT == 0 || x.C3_ND_DETRACCION_TOTAL_RT == null).ToList();



            List<string> nombresDeHojas = new List<string>();
            foreach (var worksheet in package.Workbook.Worksheets)
            {
                nombresDeHojas.Add(worksheet.Name);
            }
            ExcelWorksheet hoja1 = package.Workbook.Worksheets[nombresDeHojas[0]];

            int filaActual = 2;

            foreach (var item in lista)
            {
                hoja1.Cells["A" + filaActual].Value = item.C3_TL_SERIE_RT;
                hoja1.Cells["B" + filaActual].Value = item.C3_TL_CORRELATIVO_RT;
                hoja1.Cells["C" + filaActual].Value = item.C3_TL_TIPO_COMPROBANTE_RT;
                hoja1.Cells["D" + filaActual].Value = item.C3_TL_NOMBRE_CLIENTE_RT;
                hoja1.Cells["E" + filaActual].Value = item.C3_TL_CUENTA_RT;
                hoja1.Cells["F" + filaActual].Value = item.C3_TL_DETALLE2_Y;
                filaActual++;
            }

            return package;
        }

    }
}