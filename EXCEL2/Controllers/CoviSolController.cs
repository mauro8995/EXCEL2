using DocumentFormat.OpenXml.Drawing.Diagrams;
using EXCEL2.Models;
using EXCEL2.Models.Listas;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace EXCEL2.Controllers
{
    public class CoviSolController : Controller
    {
		// GET: CoviSol
		private AuraPortal_BPMSEntities db = new AuraPortal_BPMSEntities();
		[HttpGet]
		public ActionResult Mesual(int idPanel)
		{
			ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
			// Crear un nuevo paquete Excel
			byte[] bytesDelExcel = ObtenerDatosExcelMesual(); // Reemplaza esto con tus propios datos
																				 // Crear un MemoryStream a partir de los bytes del Excel
			using (MemoryStream memoryStream = new MemoryStream(bytesDelExcel))
			{
				// Crear un paquete Excel a partir del MemoryStream
				using (ExcelPackage package = new ExcelPackage(memoryStream))
				{
					// Obtener el contenido del paquete en un array de bytes
					byte[] excelBytes = ExcelMesual(package, idPanel).GetAsByteArray();

					return Json(UpdateDatosExcelMesual(excelBytes));


					// Devolver el archivo Excel al cliente como descarga
					return File(excelBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "archivo_excel.xlsx");
				}
			}
		}

		public byte[] ObtenerDatosExcelMesual()
		{
			var familia = db.AP_Dyn_Familias_31.OrderBy(x => x.C_Name).ToList().Last();
			var integrate = db.AP__DocIntegratedStorage.Where(x => x.IntegrationObjectId == 1).Where(x => x.SourceIntegration == 8)
				.Where(x => x.ObjectTypeProcessId == 31).Where(x => x.ObjectTypeProcess == 10).First(); 
			var integrateData = db.AP__DocIntegratedDataStorage.Where(x => x.AscDocId == integrate.ID).First();
			return integrateData.Content;
		}

		public byte[] UpdateDatosExcelMesual(byte[] excelBytes)
		{
			var familia = db.AP_Dyn_Familias_31.OrderBy(x => x.C_Name).ToList().Last();
			var integrate = db.AP__DocIntegratedStorage.Where(x => x.IntegrationObjectId == 1).Where(x => x.SourceIntegration == 8)
				.Where(x => x.ObjectTypeProcessId == 31).Where(x => x.ObjectTypeProcess == 10).First();
			var integrateData = db.AP__DocIntegratedDataStorage.Where(x => x.AscDocId == integrate.ID).First();

			integrateData.Content = excelBytes;
			db.Entry(integrateData).State = EntityState.Modified;
			db.SaveChanges();
			return integrateData.Content;
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

		public ExcelPackage ExcelMesual(ExcelPackage package, int idPanel)
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

			var familia = db.AP_Dyn_Familias_31.OrderBy(x => x.C_Name).ToList().Last();
			var anio = Int32.Parse(familia.C_Name);
			List<ExcelMes> meses = mesesList(anio);




			var Peaje = db.Panel_1008.Where(x => x.C3_PR3_ID_PANEL == idPanel).ToList().First();
			

			var format = "yyyy-MM-dd HH:mm:ss";
			var fecha1 = DateTime.ParseExact(Peaje.C3_A__PR3_FS_FECHA_INICIO.Value.ToString(format), format, new CultureInfo("en-US"));
			var fecha2 = DateTime.ParseExact(Peaje.C3_A__PR3_FS_FECHA_FIN.Value.ToString(format), format, new CultureInfo("en-US"));


			var FechaInicio = fecha1.ToString("dd MMMM", new CultureInfo("es-ES"));
			var FechaFin = fecha2.ToString("dd MMMM yyyy", new CultureInfo("es-ES"));

			int numeroEntero = (int)(Peaje.C3_PR3_ITEM_TRANSFERENCIA);
			int nuevaFila = filaActual;
			

			//hoja1.Cells[item2.Posicion + filaActual].Value = Peaje.C3_PR3_COM_FA_RECAUDACION;

			var lista = db.Panel_1008_4700.Where(x => x.C_ElementID == idPanel).ToList();
			//int anioAnterior = 0;
			hoja1.InsertRow(nuevaFila, 1);
			hoja1.Cells["B" + nuevaFila].Value = "Del " + FechaInicio + " al " + FechaFin;
			hoja1.Cells["D" + nuevaFila].Value = Peaje.C8_A__PR3_ABONO_TRANSITO_3_T_FECHA_Fecha_tranfs_min.Value.ToString("dd/MM/yyyy");//Fecha de trasferencia
			hoja1.Cells["A" + nuevaFila].Value = numeroEntero.ToString("00");//item
			hoja1.Cells["A" + nuevaFila.ToString()].Style.Font.Bold = true; // Poner en negrita
			hoja1.Cells["A" + nuevaFila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // Centrar el contenido

			hoja1.Cells["E" + nuevaFila].Formula = "=sum(F" + nuevaFila + ":T" + nuevaFila + ")";
			hoja1.Cells["E" + filaActual].Style.Numberformat.Format = "\"S/\"#,##0.00";
			hoja1.Cells["F" + nuevaFila].Value = Peaje.C3_PR3_REGULARIZAR_OTROS_DESCUENTOS;
			hoja1.Cells["F" + filaActual].Style.Numberformat.Format = "\"S/\"#,##0.00";


			var lista_hoja = db.Panel_1008_4700.Where(x => x.C_ElementID == idPanel).ToList();

			var gruposPorMesPeaje = lista_hoja.GroupBy(d => new { Mes = d.C3_FE_FECHA_COMPROBANTE_RT.Value.Month, Anio = d.C3_FE_FECHA_COMPROBANTE_RT.Value.Year }).Select(g => new
			{
				Anio = g.Key.Anio,
				Mes = g.Key.Mes,
				Suma = g.Sum(d => d.C3_ND_PEAJE_TOTAL_RT)
			}).ToArray();

			foreach (var grupo in gruposPorMesPeaje)
			{
				foreach (var item2 in meses)
				{
					if (item2.MesNumero == grupo.Mes.ToString("00") && item2.Anio == anio)
					{

						hoja1.Cells[item2.Posicion + filaActual].Value = grupo.Suma;
						hoja1.Cells[item2.Posicion + filaActual].Style.Numberformat.Format = "\"S/\"#,##0.00";
					}
				}
			}

			if (Peaje.C3_PR3_REGULARIZAR_MES != null)
			{

				foreach (var item2 in meses)
				{
					if (item2.MesNumero == Peaje.C3_PR3_REGULARIZAR_MES.Value.ToString("00") && item2.Anio == anio)
					{

						hoja1.Cells[item2.Posicion + filaActual].Value = Peaje.C3_PR3_REGULARIZAR_MONTO;

						hoja1.Cells[item2.Posicion + filaActual].Style.Numberformat.Format = "\"S/\"#,##0.00";
					}
				}
			}

			


			ExcelWorksheet hoja2 = package.Workbook.Worksheets[nombresDeHojas[1]];


			int filaActualHoja2 = 6;
			int nuevaFilaHoja2 = filaActual;
			while (hoja2.Cells[filaActualHoja2, 1].Value != null)
			{
				filaActualHoja2++;
			}
			var gruposPorMesDetraccion = lista_hoja.GroupBy(d => new { Mes = d.C3_FE_FECHA_COMPROBANTE_RT.Value.Month, Anio = d.C3_FE_FECHA_COMPROBANTE_RT.Value.Year }).Select(g => new
			{
				Anio = g.Key.Anio,
				Mes = g.Key.Mes,
				Suma = g.Sum(d => d.C3_ND_DETRACCION_TOTAL_RT)
			}).ToArray();

			if (gruposPorMesDetraccion.Count() != 0)
			{
				hoja2.InsertRow(filaActualHoja2, 1);
				hoja2.Cells["B" + nuevaFilaHoja2].Value = "Del " + FechaInicio + " al " + FechaFin;
				hoja2.Cells["D" + nuevaFilaHoja2].Value = Peaje.C8_A__PR3_ABONO_DETRACCION_3_PR3_FECHA_DET_Fecha_tranfs_min.Value.ToString("dd/MM/yyyy");//Fecha de trasferencia
				hoja2.Cells["A" + nuevaFilaHoja2].Value = numeroEntero.ToString("00");//item
				hoja2.Cells["A" + nuevaFilaHoja2.ToString()].Style.Font.Bold = true; // Poner en negrita
				hoja2.Cells["A" + nuevaFilaHoja2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // Centrar el contenido

				hoja2.Cells["E" + nuevaFilaHoja2].Formula = "=sum(F" + nuevaFila + ":T" + nuevaFila + ")";
				hoja2.Cells["E" + nuevaFilaHoja2].Style.Numberformat.Format = "\"S/\"#,##0.00";
				hoja2.Cells["F" + nuevaFilaHoja2].Value = Peaje.C3_PR3_REGULARIZAR_OTROS_DESCUENTOS_DET;
				hoja2.Cells["F" + nuevaFilaHoja2].Style.Numberformat.Format = "\"S/\"#,##0.00";
			}


			

			foreach (var grupo in gruposPorMesDetraccion)
			{
				foreach (var item2 in meses)
				{
					if (item2.MesNumero == grupo.Mes.ToString("00") && item2.Anio == anio)
					{

						hoja2.Cells[item2.Posicion + filaActual].Value = grupo.Suma;
						hoja2.Cells[item2.Posicion + filaActual].Style.Numberformat.Format = "\"S/\"#,##0.00";
					}
				}
			}

			if (Peaje.C3_PR3_REGULARIZAR_MES_DET != null)
			{
				foreach (var item2 in meses)
				{
					if (item2.MesNumero == Peaje.C3_PR3_REGULARIZAR_MES_DET.Value.ToString("00") && Peaje.C3_PR3_REGULARIZAR_AÑO_DET == anio)
					{

						hoja2.Cells[item2.Posicion + nuevaFilaHoja2].Value = Peaje.C3_PR3_REGULARIZAR_MONTO_DET;

						hoja2.Cells[item2.Posicion + nuevaFilaHoja2].Style.Numberformat.Format = "\"S/\"#,##0.00";
					}
				}
			}


			return package;
		}


		[HttpGet]
		public ActionResult ConsolidadoComprobantes(int idPanel)
		{
			ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
			// Crear un nuevo paquete Excel
			byte[] bytesDelExcel = ObtenerDatosConsolidadoComprobantes(idPanel); // Reemplaza esto con tus propios datos
															  // Crear un MemoryStream a partir de los bytes del Excel
			using (MemoryStream memoryStream = new MemoryStream(bytesDelExcel))
			{
				// Crear un paquete Excel a partir del MemoryStream
				using (ExcelPackage package = new ExcelPackage(memoryStream))
				{
					// Obtener el contenido del paquete en un array de bytes
					byte[] excelBytes = ConsolidadoComprobantes(package, idPanel).GetAsByteArray();

					return Json(UpdateDatosConsolidadoComprobante(idPanel,excelBytes));


					// Devolver el archivo Excel al cliente como descarga
					return File(excelBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "archivo_excel.xlsx");
				}
			}
		}

		public byte[] ObtenerDatosConsolidadoComprobantes( int idPanel)
		{
			var familia = db.Panel_1009.Where(x => x.ID == idPanel).ToList().Last();
			var integrate = db.AP__DocIntegratedStorage.Where(x => x.IntegrationObjectId == familia.C_ElementID)
				.Where(x => x.SourceIntegration == 6)
				.Where(x => x.ObjectTypeProcessId == 4809).Where(x => x.ObjectTypeProcess == 7).First();
			var integrateData = db.AP__DocIntegratedDataStorage.Where(x => x.AscDocId == integrate.ID).First();
			return integrateData.Content;
		}

		public byte[] UpdateDatosConsolidadoComprobante(int idPanel, byte[] excelBytes)
		{
			var familia = db.Panel_1009.Where(x => x.ID == idPanel).ToList().Last();
			var integrate = db.AP__DocIntegratedStorage.Where(x => x.IntegrationObjectId == familia.C_ElementID)
				.Where(x => x.SourceIntegration == 6)
				.Where(x => x.ObjectTypeProcessId == 4809).Where(x => x.ObjectTypeProcess == 7).First();
			var integrateData = db.AP__DocIntegratedDataStorage.Where(x => x.AscDocId == integrate.ID).First();

			integrateData.Content = excelBytes;
			db.Entry(integrateData).State = EntityState.Modified;
			db.SaveChanges();
			return integrateData.Content; 
		}



		public ExcelPackage ConsolidadoComprobantes(ExcelPackage package, int idPanel)
		{

			List<string> nombresDeHojas = new List<string>();
			foreach (var worksheet in package.Workbook.Worksheets)
			{
				nombresDeHojas.Add(worksheet.Name);
			}
			ExcelWorksheet hoja1 = package.Workbook.Worksheets[nombresDeHojas[0]];

			int filaActual = 2;

			
			var lista = db.Panel_1009_4916.Where(x => x.C_ElementID == idPanel).ToList();

			foreach (var row in lista)
			{
				hoja1.Cells["A" + filaActual].Value = row.C3_TL_TIPO_DOCUMENTO_CP;
				hoja1.Cells["B" + filaActual].Value = row.C3_TL_NRO_COMPROBANTE_CP;
				hoja1.Cells["C" + filaActual].Value = row.C3_TL_RUC_CLIENTE_CP;
				hoja1.Cells["D" + filaActual].Value = row.C3_TL_NOMBRE_CLIENTE_CP;
				hoja1.Cells["E" + filaActual].Value = row.C3_FE_FECHA_EMISION_CP.Value.ToString("dd/MM/yyyy");
				hoja1.Cells["F" + filaActual].Value = row.C3_TL_MONEDA_CP;
				hoja1.Cells["G" + filaActual].Value = row.C3_TL_IGV_CP;
				hoja1.Cells["H" + filaActual].Value = row.C3_TL_OTROS_IMPUESTOS;
				hoja1.Cells["I" + filaActual].Value = row.C3_ND_IMPORTE_TOTAL_CP;
				hoja1.Cells["J" + filaActual].Value = row.C3_TL_TIPO_EMISION;
				hoja1.Cells["K" + filaActual].Value = row.C3_TL_ESTADO_DOC_TRIBUTARIO_CP;
				hoja1.Cells["L" + filaActual].Value = row.C3_TL_URL;
				hoja1.Cells["M" + filaActual].Value = row.C3_TL_DOCUMENTO_REFERENCIA_CP;
				hoja1.Cells["N" + filaActual].Value = row.C3_TL_MENSAJE; 
				hoja1.Cells["O" + filaActual].Value = row.C3_TL_OBSERVACIONES;
				filaActual++;
			}

			return package;
		}

		[HttpGet]
		public ActionResult MesualPaso2(int idPanel)
		{
			ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
			// Crear un nuevo paquete Excel
			byte[] bytesDelExcel = ObtenerDatosExcelMesual(); // Reemplaza esto con tus propios datos
															  // Crear un MemoryStream a partir de los bytes del Excel
			using (MemoryStream memoryStream = new MemoryStream(bytesDelExcel))
			{
				// Crear un paquete Excel a partir del MemoryStream
				using (ExcelPackage package = new ExcelPackage(memoryStream))
				{
					// Obtener el contenido del paquete en un array de bytes
					byte[] excelBytes = ExcelMesualPaso2(package, idPanel).GetAsByteArray();

					return Json(UpdateDatosExcelMesual(excelBytes));


					// Devolver el archivo Excel al cliente como descarga
					return File(excelBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "archivo_excel.xlsx");
				}
			}
		}

		public ExcelPackage ExcelMesualPaso2(ExcelPackage package, int idPanel) 
		{

			List<string> nombresDeHojas = new List<string>();
			foreach (var worksheet in package.Workbook.Worksheets)
			{
				nombresDeHojas.Add(worksheet.Name);
			}
			ExcelWorksheet hoja1 = package.Workbook.Worksheets[nombresDeHojas[0]];

			int filaActual = 6;
				
			var lista = db.Panel_1009.Where(x => x.C3_PR3_ID_PANEL == idPanel).First();
			int anio = (int)lista.C3_PR3_EN_AÑO;
			List<ExcelMes> meses = mesesList(anio);


            foreach (var item in meses)
            {
                
				if(item.MesNombre == lista.C3_PR3_COM_MES_TEXTO && item.Anio == lista.C3_PR3_EN_AÑO)
				{
					hoja1.Cells[item.Posicion + filaActual].Value = lista.C3_PR3_COM_FA_RECAUDACION;
					hoja1.Cells[item.Posicion + filaActual].Style.Numberformat.Format = "\"S/\"#,##0.00";
				}
            
            }

			ExcelWorksheet hoja2 = package.Workbook.Worksheets[nombresDeHojas[1]];

			foreach (var item in meses)
			{

				if (item.MesNombre == lista.C3_PR3_COM_MES_TEXTO && item.Anio == lista.C3_PR3_EN_AÑO)
				{
					hoja2.Cells[item.Posicion + filaActual].Value = lista.C8_GC_COMPROBANTES_TRANSITOS_3_ND_DETRACCION_TOTAL_RT_ND_TOTAL_DETRACCION;
					hoja2.Cells[item.Posicion + filaActual].Style.Numberformat.Format = "\"S/\"#,##0.00";
				}

			}

			return package;
		}

	}
}
