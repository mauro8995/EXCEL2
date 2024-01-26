using DocumentFormat.OpenXml.Spreadsheet;
using EXCEL2.Models;
using EXCEL2.Models.Listas;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Drawing;
using System.Globalization;
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
                    if (item2.MesNumero == item.C3_PR3_MES_TRANSFERENCIA.Value.ToString("00") && item2.Anio == anio)
                    {
                        
                            hoja1.Cells[item2.Posicion + '6'].Value = item.C3_PR3_COM_FA_RECAUDACION;
                        
                        hoja1.Cells[item2.Posicion + '6'].Style.Numberformat.Format = "\"S/\"#,##0.00";
                        //hoja1.Cells[item2.Posicion + (filaActual + 2)].Style.Numberformat.Format = "\"S/\"#,##0.00";
                        //hoja1.Cells[item2.Posicion + (filaActual + 3)].Style.Numberformat.Format = "\"S/\"#,##0.00";
                    }
                }
            }



            ExcelWorksheet hoja2 = package.Workbook.Worksheets[nombresDeHojas[1]];

            int filaActualHoja2 = 6;
            int columnaIndexHoja2 = 1;
            while (hoja2.Cells[filaActualHoja2, columnaIndexHoja2].Value != null)
            {
                
                filaActualHoja2++;
            }

            foreach (var item in Comision)
            {
                foreach (var item2 in meses)
                {
                    if (item2.MesNumero == item.C3_PR3_MES_TRANSFERENCIA.Value.ToString("00") && item2.Anio == anio)
                    {
                        
                            hoja2.Cells[item2.Posicion + '6'].Value = item.C8_GC_COMPROBANTES_TRANSITOS_3_ND_DETRACCION_TOTAL_RT_ND_TOTAL_DETRACCION;
                            //hoja2.Cells[item2.Posicion + (filaActual + 2)].Value = item.C8_GC_COMPROBANTES_TRANSITOS_3_ND_DETRACCION_TOTAL_RT_ND_TOTAL_DETRACCION;
                            //hoja2.Cells[item2.Posicion + (filaActual + 3)].Value = item.C3_PR3_COM_POR_TRANSFERIR_DET;
                        
                        hoja2.Cells[item2.Posicion + '6'].Style.Numberformat.Format = "\"S/\"#,##0.00";
                        //hoja2.Cells[item2.Posicion + (filaActual + 2)].Style.Numberformat.Format = "\"S/\"#,##0.00";
                        //hoja2.Cells[item2.Posicion + (filaActual + 3)].Style.Numberformat.Format = "\"S/\"#,##0.00";
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




        //Consolidado de comprovantes------------------------------

        [HttpGet]
		public ActionResult ConsolidadoComprobantes(int idPanel)
        {
			ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
			// Crear un nuevo paquete Excel
			byte[] bytesDelExcel = ObtenerExcelConsolidadoComprobantes(idPanel); // Reemplaza esto con tus propios datos

			// Crear un MemoryStream a partir de los bytes del Excel
			using (MemoryStream memoryStream = new MemoryStream(bytesDelExcel))
			{
				// Crear un paquete Excel a partir del MemoryStream
				using (ExcelPackage package = new ExcelPackage(memoryStream))
				{

					// Obtener el contenido del paquete en un array de bytes
					byte[] excelBytes = ConsolidadoComprobantesEdit(package, idPanel).GetAsByteArray();

					return Json(UpdateExcelConsolidadoComprobantes(idPanel,excelBytes));


					// Devolver el archivo Excel al cliente como descarga
					return File(excelBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "archivo_excel.xlsx");
				}
			}
		}

        public byte[] ObtenerExcelConsolidadoComprobantes(int idPanel)
        {
			var elemeto = db.Panel_1009.Where(x => x.ID == idPanel).First();
			var integrate = db.AP__DocIntegratedStorage.Where(x => x.IntegrationObjectId == elemeto.C_ElementID).Where(x => x.AscTermId == 251).First();
			var integrateData = db.AP__DocIntegratedDataStorage.Where(x => x.AscDocId == integrate.ID).First();
			return integrateData.Content;
		}

		public byte[] UpdateExcelConsolidadoComprobantes(int idPanel,byte[] excelBytes)
		{
			var elemeto = db.Panel_1009.Where(x => x.ID == idPanel).First();
			var integrate = db.AP__DocIntegratedStorage.Where(x => x.IntegrationObjectId == elemeto.C_ElementID).Where(x => x.AscTermId == 251).First();
			var integrateData = db.AP__DocIntegratedDataStorage.Where(x => x.AscDocId == integrate.ID).First();

			integrateData.Content = excelBytes;
			db.Entry(integrateData).State = EntityState.Modified;
			db.SaveChanges();
			return integrateData.Content;
		}

		public ExcelPackage ConsolidadoComprobantesEdit(ExcelPackage package, int idPanel)
        {
			List<string> nombresDeHojas = new List<string>();
			foreach (var worksheet in package.Workbook.Worksheets)
			{
				nombresDeHojas.Add(worksheet.Name);
			}
			ExcelWorksheet hoja1 = package.Workbook.Worksheets[nombresDeHojas[0]];

            int filaInicial = 4;

            var Comprobantes = db.Panel_1009_4916.Where(x => x.C_ElementID == idPanel).OrderBy(x => x.C3_TL_NRO_COMPROBANTE_CP).ThenBy(x => x.C3_FE_FECHA_EMISION_CP).ToList();


            foreach (var item in Comprobantes)
            {

				ExcelRange border = hoja1.Cells["A" + filaInicial + ":V" + filaInicial];

				// Agregar bordes al rango de celdas
				border.Style.Border.Top.Style = ExcelBorderStyle.Thin;
				border.Style.Border.Left.Style = ExcelBorderStyle.Thin;
				border.Style.Border.Right.Style = ExcelBorderStyle.Thin;
				border.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
				hoja1.Cells["A" + filaInicial].Value = item.C3_TL_RUC_CLIENTE_CP;
				hoja1.Cells["B" + filaInicial].Value = item.C3_TL_NOMBRE_CLIENTE_CP;
				hoja1.Cells["C" + filaInicial].Value = item.C3_TL_TIPO_DOCUMENTO_CP;
				hoja1.Cells["D" + filaInicial].Value = item.C3_TL_NRO_COMPROBANTE_CP;
				hoja1.Cells["E" + filaInicial].Value = item.C3_TL_ORDEN_COMPRA;
				hoja1.Cells["F" + filaInicial].Value = item.C3_TL_MONEDA_CP;
				hoja1.Cells["G" + filaInicial].Value = item.C3_TL_OP_GRAVADA;
				hoja1.Cells["H" + filaInicial].Value = item.C3_TL_OP_NO_GRAVADA;
				hoja1.Cells["I" + filaInicial].Value = item.C3_TL_IGV_CP;
				hoja1.Cells["J" + filaInicial].Value = item.C3_TL_OTROS_IMPUESTOS;
				hoja1.Cells["K" + filaInicial].Value = item.C3_TL_OTROS_CARGOS;
				hoja1.Cells["L" + filaInicial].Value = item.C3_ND_IMPORTE_TOTAL_CP;
				hoja1.Cells["L" + filaInicial].Style.Numberformat.Format = "\"S/\"#,##0.00";
				ExcelRange CentrarTotal = hoja1.Cells["L" + filaInicial];
				CentrarTotal.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
				CentrarTotal.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
				if (item.C3_FE_FECHA_EMISION_CP != null)
                {
					hoja1.Cells["M" + filaInicial].Value = item.C3_FE_FECHA_EMISION_CP.Value.ToString("dd/MM/yyyy");
				}
				if (item.C3_F_FECHA_VENCIMIENTO != null)
				{
					hoja1.Cells["M" + filaInicial].Value = item.C3_F_FECHA_VENCIMIENTO.Value.ToString("dd/MM/yyyy");
				}

				hoja1.Cells["O" + filaInicial].Value = item.C3_TL_ESTADO_DOC;
				hoja1.Cells["P" + filaInicial].Value = item.C3_TL_ESTADO_DOC_TRIBUTARIO_CP;
				hoja1.Cells["Q" + filaInicial].Value = item.C3_F_ENVIADO_DECLARAR_TEXTO;
				hoja1.Cells["R" + filaInicial].Value = item.C3_TL_DOCUMENTO_REFERENCIA_CP;
				hoja1.Cells["S" + filaInicial].Value = item.C3_TL_OBSERVACION_SUNAT;
				hoja1.Cells["T" + filaInicial].Value = item.C3_TL_HABILITADO;
				hoja1.Cells["U" + filaInicial].Value = item.C3_TL_TIPO_EMISION;
				hoja1.Cells["V" + filaInicial].Value = item.C3_TL_OBSERVACIONES;
                filaInicial++;
			}

			hoja1.Cells["L" + filaInicial].Value = Comprobantes.Sum(x=> x.C3_ND_IMPORTE_TOTAL_CP);
			hoja1.Cells["L" + filaInicial].Style.Numberformat.Format = "\"S/\"#,##0.00";


			ExcelRange range = hoja1.Cells["L" + filaInicial]; // Ejemplo: celda A1

			// Aplicar un color de fondo
			range.Style.Fill.PatternType = ExcelFillStyle.Solid;
            System.Drawing.Color hexColor = ColorTranslator.FromHtml("#99CCFF");
			range.Style.Fill.BackgroundColor.SetColor(hexColor);
			range.Style.Font.Bold = true;
			return package;
		}



      

	}
}