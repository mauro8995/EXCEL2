using EXCEL2.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace EXCEL2.Controllers
{
    public class CarguerosTerrestresController : Controller
    {
        // GET: CarguerosTerrestres
        public ActionResult Index()
        {
            return View();
        }

		private AuraPortal_BPMSEntities db = new AuraPortal_BPMSEntities();



		public ActionResult GenerarArchivoCarguerosPostpagoConDetraccion(int idPanel)
		{
			ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

			var filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "AGRICOLA_SANTA_AZUL.xlsx");
			// Crear un nuevo paquete Excel
			byte[] bytesDelExcel = ObtenerPostpagoConDetraccion(); // Reemplaza esto con tus propios datos

			// Crear un MemoryStream a partir de los bytes del Excel
			using (MemoryStream memoryStream = new MemoryStream(bytesDelExcel))
			{
				// Crear un paquete Excel a partir del MemoryStream
				using (ExcelPackage package = new ExcelPackage(memoryStream))
				{
					// Obtener el contenido del paquete en un array de bytes
					byte[] excelBytes = CarguerosPostpagoConDestraccion(package, idPanel).GetAsByteArray();

					//return Json(UpdateDatosVarbinaryVisa(excelBytes));


					// Devolver el archivo Excel al cliente como descarga
					return File(bytesDelExcel, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "archivo_excel.xlsx");
				}
			}
		}

		public ActionResult GenerarArchivoCarguerosPostpagoSinDetraccion(int idPanel)
		{
			ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

			var filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "AGRICOLA_SANTA_AZUL_SIN_DETRACCION.xlsx");
			// Crear un nuevo paquete Excel
			byte[] bytesDelExcel = ObtenerPostpagoSinDetraccion();// System.IO.File.ReadAllBytes(filePath); // Reemplaza esto con tus propios datos

			// Crear un MemoryStream a partir de los bytes del Excel
			using (MemoryStream memoryStream = new MemoryStream(bytesDelExcel))
			{
				// Crear un paquete Excel a partir del MemoryStream
				using (ExcelPackage package = new ExcelPackage(memoryStream))
				{
					// Obtener el contenido del paquete en un array de bytes
					byte[] excelBytes = CarguerosPostpagoSinDestraccion(package, idPanel).GetAsByteArray();

					//return Json(UpdateDatosVarbinaryVisa(excelBytes));


					// Devolver el archivo Excel al cliente como descarga
					return File(excelBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "archivo_excel.xlsx");
				}
			}
		}

		public ActionResult GenerarArchivoCarguerosPrepagoConDetraccion(int idPanel)//ya esta
		{
			ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

			var filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "AGRICOLA_SANTA_AZUL.xlsx");
			// Crear un nuevo paquete Excel
			byte[] bytesDelExcel = ObtenerPrepagoConDetraccion();//System.IO.File.ReadAllBytes(filePath); // Reemplaza esto con tus propios datos

			// Crear un MemoryStream a partir de los bytes del Excel
			using (MemoryStream memoryStream = new MemoryStream(bytesDelExcel))
			{
				// Crear un paquete Excel a partir del MemoryStream
				using (ExcelPackage package = new ExcelPackage(memoryStream))
				{
					// Obtener el contenido del paquete en un array de bytes
					byte[] excelBytes = CarguerosPrepagoConDetraccion(package, idPanel).GetAsByteArray();

					//return Json(UpdateDatosVarbinaryVisa(excelBytes));


					// Devolver el archivo Excel al cliente como descarga
					return File(excelBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "archivo_excel.xlsx");
				}
			}
		}

		public ActionResult GenerarArchivoCarguerosPrepagoSinDetraccion(int idPanel)//ya esta
		{
			ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

			var filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "AGRICOLA_SANTA_AZUL_SIN_DETRACCION.xlsx");
			// Crear un nuevo paquete Excel
			byte[] bytesDelExcel = ObtenerPrepagoSinDetraccion(); //System.IO.File.ReadAllBytes(filePath); // Reemplaza esto con tus propios datos

			// Crear un MemoryStream a partir de los bytes del Excel
			using (MemoryStream memoryStream = new MemoryStream(bytesDelExcel))
			{
				// Crear un paquete Excel a partir del MemoryStream
				using (ExcelPackage package = new ExcelPackage(memoryStream))
				{
					// Obtener el contenido del paquete en un array de bytes
					byte[] excelBytes = CarguerosPrepagoSinDetraccion(package, idPanel).GetAsByteArray();

					//return Json(UpdateDatosVarbinaryVisa(excelBytes));


					// Devolver el archivo Excel al cliente como descarga
					return File(excelBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "archivo_excel.xlsx");
				}
			}
		}


		public byte[] ObtenerPrepagoSinDetraccion()
		{
			var familia = db.AP_Dyn_Familias_30.ToList().Last();
			var integrate = db.AP__DocIntegratedStorage.Where(x => x.AscTermId == 8010).First();
			var integrateData = db.AP__DocIntegratedDataStorage.Where(x => x.AscDocId == integrate.ID).First();
			return integrateData.Content;

		}
		public byte[] ObtenerPrepagoConDetraccion()
		{
			var familia = db.AP_Dyn_Familias_30.ToList().Last();
			var integrate = db.AP__DocIntegratedStorage.Where(x => x.AscTermId == 8015).First();
			var integrateData = db.AP__DocIntegratedDataStorage.Where(x => x.AscDocId == integrate.ID).First();
			return integrateData.Content;

		}

		public byte[] ObtenerPostpagoSinDetraccion()
		{
			var familia = db.AP_Dyn_Familias_30.ToList().Last();
			var integrate = db.AP__DocIntegratedStorage.Where(x => x.AscTermId == 8011).First();
			var integrateData = db.AP__DocIntegratedDataStorage.Where(x => x.AscDocId == integrate.ID).First();
			return integrateData.Content;
		}
		public byte[] ObtenerPostpagoConDetraccion()
		{
			var familia = db.AP_Dyn_Familias_30.ToList().Last();
			var integrate = db.AP__DocIntegratedStorage.Where(x => x.AscTermId == 8016).First();
			var integrateData = db.AP__DocIntegratedDataStorage.Where(x => x.AscDocId == integrate.ID).First();
			return integrateData.Content;
		}

		public ExcelPackage CarguerosPrepagoConDetraccion(ExcelPackage package,int idPanel)
		{
			List<string> nombresDeHojas = new List<string>();
			foreach (var worksheet in package.Workbook.Worksheets)
			{
				nombresDeHojas.Add(worksheet.Name);
			}

			ExcelWorksheet hoja1 = package.Workbook.Worksheets[nombresDeHojas[0]];

			var fecha = db.Panel_2013.Where(x => x.ID == idPanel).First();

			var dataHoja1 = db.Panel_2013_8987.Where(x => x.C_ElementID == idPanel).ToList();

			int filaInicialPlacas = 2;


            foreach (var item in dataHoja1)
            {
				hoja1.Cells["A" + filaInicialPlacas].Value = item.C3_PRE_TL_PLACA;
				hoja1.Cells["B" + filaInicialPlacas].Value = item.C3_PRE_TL_CATEGORIA;
				hoja1.Cells["C" + filaInicialPlacas].Value = item.C3_PRE_TL_FABRICANTE;
				hoja1.Cells["D" + filaInicialPlacas].Value = item.C3_PRE_TL_MODELO;
				hoja1.Cells["E" + filaInicialPlacas].Value = item.C3_PRE_TL_COLOR;
				hoja1.Cells["F" + filaInicialPlacas].Value = item.C3_PRE_TL_ESTADO;


				ExcelRange border = hoja1.Cells["A" + filaInicialPlacas + ":F" + filaInicialPlacas];

				// Agregar bordes al rango de celdas
				border.Style.Border.Top.Style = ExcelBorderStyle.Thin;
				border.Style.Border.Left.Style = ExcelBorderStyle.Thin;
				border.Style.Border.Right.Style = ExcelBorderStyle.Thin;
				border.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

				ExcelRange TextCentrar = hoja1.Cells["A" + filaInicialPlacas + ":F" + filaInicialPlacas];

				TextCentrar.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
				TextCentrar.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

				filaInicialPlacas++;
			}
           


			ExcelWorksheet hoja2 = package.Workbook.Worksheets[nombresDeHojas[1]];


			hoja2.Cells["B8"].Value = fecha.C8_GC_PREPAGOS_FINALES_3_PR_EMPRESA_ULTIMA_EMPRESA_PREP;
			hoja2.Cells["B9"].Value = fecha.C3_FE_FECHA_DE_INICIO_DE_PERIODO + " - " + fecha.C3_FE_FECHA_DE_FIN_DE_PERIODO;


			hoja2.Cells["F3"].Value = "Saldo al " +fecha.C3_FE_FECHA_DE_INICIO_DE_PERIODO.Value.ToString("dd/MM/yyyy HH:mm:ss");
			hoja2.Cells["F4"].Value = "Recarga del " + fecha.C3_FE_FECHA_DE_INICIO_DE_PERIODO.Value.ToString("dd") + " al " + fecha.C3_FE_FECHA_DE_FIN_DE_PERIODO.Value.ToString("dd");
			hoja2.Cells["F5"].Value = "Consumo del " + fecha.C3_FE_FECHA_DE_INICIO_DE_PERIODO.Value.ToString("dd") + " al " + fecha.C3_FE_FECHA_DE_FIN_DE_PERIODO.Value.ToString("dd") + " Covisol";
			hoja2.Cells["F6"].Value = "Consumo del " + fecha.C3_FE_FECHA_DE_INICIO_DE_PERIODO.Value.ToString("dd") + " al " + fecha.C3_FE_FECHA_DE_FIN_DE_PERIODO.Value.ToString("dd") + " Coviperú";
			hoja2.Cells["F7"].Value = "Saldo al " + fecha.C3_FE_FECHA_DE_FIN_DE_PERIODO.Value.ToString("dd/MM/yyyy HH:mm:ss");

			hoja2.Cells["H3"].Value = fecha.C3_ND_SALDODELULTIMOREPORTE; 
			hoja2.Cells["H4"].Value = fecha.C8_PRE_COMPROBANTESREPORTES_3_PRE_MONTORECARGA_PRE_SUMAREGARGAS; 
			hoja2.Cells["H5"].Value = fecha.C8_GC_TRANSITOGENERALESSPREPAGO_3_ND_SALDOCOVISOL_PREPAGO_CONSUMOSCOVISOL_PRE;
			hoja2.Cells["H6"].Value = fecha.C8_GC_TRANSITOGENERALESSPREPAGO_3_ND_SALDOCOVISUR_PREPAGO_CONSUMOSCOVISUR_PRE;
			hoja2.Cells["H7"].Value = fecha.C8_GC_TRANSITOGENERALESSPREPAGO_3_ND_SALDOCOVIPERU_PREPAGO_CONSUMOSCOVIPERU_PRE;

			hoja2.Cells["I3"].Value = fecha.C3_ND_SALDODELULTIMOREPORTE;
			hoja2.Cells["I4"].Value = fecha.C8_PRE_COMPROBANTESREPORTES_3_PRE_MONTODETRACCION_PRE_SUMADETRACCIONES;
			hoja2.Cells["I5"].Value = fecha.C8_GC_TRANSITOGENERALESSPREPAGO_3_ND_CONSUDETRACCOVISO_SUMADETRACCOVISOL;
			hoja2.Cells["I6"].Value = fecha.C8_GC_TRANSITOGENERALESSPREPAGO_3_ND_CONSUDETRACCOVISUR_SUMADETRACCOVISUR;
			hoja2.Cells["I7"].Value = fecha.C8_GC_TRANSITOGENERALESSPREPAGO_3_ND_CONSUDETRACCOVIPERU_SUMADETRACCOVIPERU;

			int contadorHoja2 = 11;
			var panel = db.Panel_2013_8730.Where(x => x.C_ElementID == idPanel).ToList();
			int cantidadElementos = 1;
            foreach (var item in panel)
            {
				hoja2.InsertRow(contadorHoja2, 1);
				ExcelRange borderHoja2 = hoja2.Cells["A" + contadorHoja2 + ":L" + contadorHoja2];

				// Agregar bordes al rango de celdas
				borderHoja2.Style.Border.Top.Style = ExcelBorderStyle.Thin;
				borderHoja2.Style.Border.Left.Style = ExcelBorderStyle.Thin;
				borderHoja2.Style.Border.Right.Style = ExcelBorderStyle.Thin;
				borderHoja2.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

				ExcelRange TextCentrarHoja2 = hoja2.Cells["A" + contadorHoja2 + ":L" + contadorHoja2];

				TextCentrarHoja2.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
				TextCentrarHoja2.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

				hoja2.Cells["A" + contadorHoja2].Value = item.C3_FE_FECHAYHORATRANSITO_PRE.Value.ToString("dd/MM/yyyy HH:mm:ss");
				hoja2.Cells["B" + contadorHoja2].Value = item.C3_PR4_TL_ESTACION;
				hoja2.Cells["C" + contadorHoja2].Value = item.C3_PR4_TL_VIA;
				hoja2.Cells["D" + contadorHoja2].Value = item.C3_PR4_TL_PLACA;
				hoja2.Cells["E" + contadorHoja2].Value = item.C3_PR4_TL_CATEGORIA;
				hoja2.Cells["F" + contadorHoja2].Value = item.C3_PR4_ND_VALORPEAJE;
				hoja2.Cells["F" + contadorHoja2].Style.Numberformat.Format = "\"S/\"#,##0.00";
				hoja2.Cells["G" + contadorHoja2].Value = item.C3_ND_VALORDEDETRACCION;
				hoja2.Cells["G" + contadorHoja2].Style.Numberformat.Format = "\"S/\"#,##0.00";
				hoja2.Cells["H" + contadorHoja2].Value = item.C3_TL_NUMERODEDETRAC_PRE;
				hoja2.Cells["I" + contadorHoja2].Value = (item.C3_FE_FECHADECOMPROBANTE_PRE == null ? "": item.C3_FE_FECHADECOMPROBANTE_PRE.Value.ToString("dd/MM/yyyy HH:mm:ss")) ;
				hoja2.Cells["J" + contadorHoja2].Value = item.C3_PR4_TL_NUMERODECOMPROBANTE;
				hoja2.Cells["K" + contadorHoja2].Value = item.C3_PR4_TL_RUCCONSECION;
				hoja2.Cells["L" + contadorHoja2].Value = item.C3_PR4_TL_TIPODECOMPROBANTE;
				cantidadElementos++;
				contadorHoja2++;

			}

			hoja2.Cells["F" + contadorHoja2].Formula = "SUBTOTAL(9,F11:F"+ (contadorHoja2-1) + ")";
			hoja2.Cells["G" + contadorHoja2].Formula = "SUBTOTAL(9,F11:F" + (contadorHoja2 - 1) + ")";

			ExcelRange borderHoja2a = hoja2.Cells["F" + contadorHoja2];

			// Agregar bordes al rango de celdas
			borderHoja2a.Style.Border.Top.Style = ExcelBorderStyle.Thin;
			borderHoja2a.Style.Border.Left.Style = ExcelBorderStyle.Thin;
			borderHoja2a.Style.Border.Right.Style = ExcelBorderStyle.Thin;
			borderHoja2a.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;


			ExcelWorksheet hoja3 = package.Workbook.Worksheets[nombresDeHojas[2]];

			var dataHoja3 = db.Panel_2013_8925.Where(x => x.C_ElementID==idPanel).ToList();
			int contadorHoja3 = 3;
            foreach (var item in dataHoja3)
            {
				hoja3.InsertRow(contadorHoja3, 1);


				ExcelRange borderHoja3 = hoja3.Cells["B" + contadorHoja3 + ":F" + contadorHoja3];

				// Agregar bordes al rango de celdas
				borderHoja3.Style.Border.Top.Style = ExcelBorderStyle.Thin;
				borderHoja3.Style.Border.Left.Style = ExcelBorderStyle.Thin;
				borderHoja3.Style.Border.Right.Style = ExcelBorderStyle.Thin;
				borderHoja3.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

				ExcelRange TextCentrarHoja3 = hoja3.Cells["B" + contadorHoja3 + ":F" + contadorHoja3];

				TextCentrarHoja3.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
				TextCentrarHoja3.Style.VerticalAlignment = ExcelVerticalAlignment.Center;



				hoja3.Cells["B" + contadorHoja3].Value = item.C3_PRE_FECHARECARGA == null ? "" : item.C3_PRE_FECHARECARGA.Value.ToString("dd/MM/yyyy HH:mm:ss");
				hoja3.Cells["C" + contadorHoja3].Value = item.C3_PRE_MONTORECARGA;
				hoja3.Cells["C" + contadorHoja3].Style.Numberformat.Format = "\"S/\"#,##0.00";
				hoja3.Cells["D" + contadorHoja3].Value = item.C3_PRE_TL_RECIBO;
				hoja3.Cells["D" + contadorHoja3].Style.Numberformat.Format = "\"S/\"#,##0.00";
				hoja3.Cells["E" + contadorHoja3].Value = item.C3_PRE_MONTODETRACCION;
				hoja3.Cells["F" + contadorHoja3].Value = item.C3_PRE_RECIBODETRACCION;
				contadorHoja3++;
			}
			hoja3.Cells["C" + contadorHoja3].Formula = "SUM(C3:C" + (contadorHoja3 - 1) + ")";
			hoja3.Cells["E" + contadorHoja3].Formula = "SUM(E3:E" + (contadorHoja3 - 1) + ")";
			return package;
		}//--

		public ExcelPackage CarguerosPrepagoSinDetraccion(ExcelPackage package, int idPanel)
		{
			List<string> nombresDeHojas = new List<string>();
			foreach (var worksheet in package.Workbook.Worksheets)
			{
				nombresDeHojas.Add(worksheet.Name);
			}

			ExcelWorksheet hoja1 = package.Workbook.Worksheets[nombresDeHojas[0]];

			var fecha = db.Panel_2013.Where(x => x.ID == idPanel).First();

			var dataHoja1 = db.Panel_2013_8987.Where(x => x.C_ElementID == idPanel).ToList();

			int filaInicialPlacas = 2;


			foreach (var item in dataHoja1)
			{
				hoja1.Cells["A" + filaInicialPlacas].Value = item.C3_PRE_TL_PLACA;
				hoja1.Cells["B" + filaInicialPlacas].Value = item.C3_PRE_TL_CATEGORIA;
				hoja1.Cells["C" + filaInicialPlacas].Value = item.C3_PRE_TL_FABRICANTE;
				hoja1.Cells["D" + filaInicialPlacas].Value = item.C3_PRE_TL_MODELO;
				hoja1.Cells["E" + filaInicialPlacas].Value = item.C3_PRE_TL_COLOR;
				hoja1.Cells["F" + filaInicialPlacas].Value = item.C3_PRE_TL_ESTADO;


				ExcelRange border = hoja1.Cells["A" + filaInicialPlacas + ":F" + filaInicialPlacas];

				// Agregar bordes al rango de celdas
				border.Style.Border.Top.Style = ExcelBorderStyle.Thin;
				border.Style.Border.Left.Style = ExcelBorderStyle.Thin;
				border.Style.Border.Right.Style = ExcelBorderStyle.Thin;
				border.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

				ExcelRange TextCentrar = hoja1.Cells["A" + filaInicialPlacas + ":F" + filaInicialPlacas];

				TextCentrar.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
				TextCentrar.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

				filaInicialPlacas++;
			}



			ExcelWorksheet hoja2 = package.Workbook.Worksheets[nombresDeHojas[1]];


			hoja2.Cells["B7"].Value = fecha.C8_GC_PREPAGOS_FINALES_3_PR_EMPRESA_ULTIMA_EMPRESA_PREP;
			hoja2.Cells["B8"].Value = fecha.C3_FE_FECHA_DE_INICIO_DE_PERIODO + " - " + fecha.C3_FE_FECHA_DE_FIN_DE_PERIODO;


			hoja2.Cells["F3"].Value = "Saldo al " + fecha.C3_FE_FECHA_DE_INICIO_DE_PERIODO.Value.ToString("dd/MM/yyyy HH:mm:ss");
			hoja2.Cells["F4"].Value = "Recarga del " + fecha.C3_FE_FECHA_DE_INICIO_DE_PERIODO.Value.ToString("dd") + " al " + fecha.C3_FE_FECHA_DE_FIN_DE_PERIODO.Value.ToString("dd");
			hoja2.Cells["F5"].Value = "Consumo del " + fecha.C3_FE_FECHA_DE_INICIO_DE_PERIODO.Value.ToString("dd") + " al " + fecha.C3_FE_FECHA_DE_FIN_DE_PERIODO.Value.ToString("dd") + " Covisol";
			hoja2.Cells["F6"].Value = "Consumo del " + fecha.C3_FE_FECHA_DE_INICIO_DE_PERIODO.Value.ToString("dd") + " al " + fecha.C3_FE_FECHA_DE_FIN_DE_PERIODO.Value.ToString("dd") + " Coviperú";
			hoja2.Cells["F7"].Value = "Saldo al " + fecha.C3_FE_FECHA_DE_FIN_DE_PERIODO.Value.ToString("dd/MM/yyyy HH:mm:ss");

			hoja2.Cells["H3"].Value = fecha.C3_ND_SALDODELULTIMOREPORTE;
			hoja2.Cells["H4"].Value = fecha.C8_PRE_COMPROBANTESREPORTES_3_PRE_MONTORECARGA_PRE_SUMAREGARGAS;
			hoja2.Cells["H5"].Value = fecha.C8_GC_TRANSITOGENERALESSPREPAGO_3_ND_SALDOCOVISOL_PREPAGO_CONSUMOSCOVISOL_PRE;
			hoja2.Cells["H6"].Value = fecha.C8_GC_TRANSITOGENERALESSPREPAGO_3_ND_SALDOCOVISUR_PREPAGO_CONSUMOSCOVISUR_PRE;
			hoja2.Cells["H7"].Value = fecha.C8_GC_TRANSITOGENERALESSPREPAGO_3_ND_SALDOCOVIPERU_PREPAGO_CONSUMOSCOVIPERU_PRE;


			int contadorHoja2 = 10;
			var panel = db.Panel_2013_8730.Where(x => x.C_ElementID == idPanel).ToList();
			int cantidadElementos = 1;
			foreach (var item in panel)
			{
				hoja2.InsertRow(contadorHoja2, 1);
				ExcelRange borderHoja2 = hoja2.Cells["A" + contadorHoja2 + ":J" + contadorHoja2];

				// Agregar bordes al rango de celdas
				borderHoja2.Style.Border.Top.Style = ExcelBorderStyle.Thin;
				borderHoja2.Style.Border.Left.Style = ExcelBorderStyle.Thin;
				borderHoja2.Style.Border.Right.Style = ExcelBorderStyle.Thin;
				borderHoja2.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

				ExcelRange TextCentrarHoja2 = hoja2.Cells["A" + contadorHoja2 + ":J" + contadorHoja2];

				TextCentrarHoja2.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
				TextCentrarHoja2.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

				hoja2.Cells["A" + contadorHoja2].Value = item.C3_FE_FECHAYHORATRANSITO_PRE.Value.ToString("dd/MM/yyyy HH:mm:ss");
				hoja2.Cells["B" + contadorHoja2].Value = item.C3_PR4_TL_ESTACION;
				hoja2.Cells["C" + contadorHoja2].Value = item.C3_PR4_TL_VIA;
				hoja2.Cells["D" + contadorHoja2].Value = item.C3_PR4_TL_PLACA;
				hoja2.Cells["E" + contadorHoja2].Value = item.C3_PR4_TL_CATEGORIA;
				hoja2.Cells["F" + contadorHoja2].Value = item.C3_PR4_ND_VALORPEAJE;
				hoja2.Cells["F" + contadorHoja2].Style.Numberformat.Format = "\"S/\"#,##0.00";
				hoja2.Cells["G" + contadorHoja2].Value = (item.C3_FE_FECHADECOMPROBANTE_PRE == null ? "" : item.C3_FE_FECHADECOMPROBANTE_PRE.Value.ToString("dd/MM/yyyy HH:mm:ss"));
				hoja2.Cells["H" + contadorHoja2].Value = item.C3_PR4_TL_NUMERODECOMPROBANTE;
				hoja2.Cells["I" + contadorHoja2].Value = item.C3_PR4_TL_RUCCONSECION;
				hoja2.Cells["J" + contadorHoja2].Value = item.C3_PR4_TL_TIPODECOMPROBANTE;
				cantidadElementos++;
				contadorHoja2++;

			}

			hoja2.Cells["F" + contadorHoja2].Formula = "SUBTOTAL(9,F11:F" + (contadorHoja2 - 1) + ")";

			ExcelRange borderHoja2a = hoja2.Cells["F" + contadorHoja2];

			// Agregar bordes al rango de celdas
			borderHoja2a.Style.Border.Top.Style = ExcelBorderStyle.Thin;
			borderHoja2a.Style.Border.Left.Style = ExcelBorderStyle.Thin;
			borderHoja2a.Style.Border.Right.Style = ExcelBorderStyle.Thin;
			borderHoja2a.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;


			ExcelWorksheet hoja3 = package.Workbook.Worksheets[nombresDeHojas[2]];

			var dataHoja3 = db.Panel_2013_8925.Where(x => x.C_ElementID == idPanel).ToList();
			int contadorHoja3 = 3;
			foreach (var item in dataHoja3)
			{
				hoja3.InsertRow(contadorHoja3, 1);


				ExcelRange borderHoja3 = hoja3.Cells["B" + contadorHoja3 + ":D" + contadorHoja3];

				// Agregar bordes al rango de celdas
				borderHoja3.Style.Border.Top.Style = ExcelBorderStyle.Thin;
				borderHoja3.Style.Border.Left.Style = ExcelBorderStyle.Thin;
				borderHoja3.Style.Border.Right.Style = ExcelBorderStyle.Thin;
				borderHoja3.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

				ExcelRange TextCentrarHoja3 = hoja3.Cells["B" + contadorHoja3 + ":D" + contadorHoja3];

				TextCentrarHoja3.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
				TextCentrarHoja3.Style.VerticalAlignment = ExcelVerticalAlignment.Center;



				hoja3.Cells["B" + contadorHoja3].Value = item.C3_PRE_FECHARECARGA == null ? "" : item.C3_PRE_FECHARECARGA.Value.ToString("dd/MM/yyyy HH:mm:ss");
				hoja3.Cells["C" + contadorHoja3].Value = item.C3_PRE_MONTORECARGA;
				hoja3.Cells["C" + contadorHoja3].Style.Numberformat.Format = "\"S/\"#,##0.00";
				hoja3.Cells["D" + contadorHoja3].Value = item.C3_PRE_TL_RECIBO;
				contadorHoja3++;
			}
			hoja3.Cells["C" + contadorHoja3].Formula = "SUM(C3:C" + (contadorHoja3 - 1) + ")";
			return package;
		}//--


		public ExcelPackage CarguerosPostpagoConDestraccion(ExcelPackage package, int idPanel)
		{
			List<string> nombresDeHojas = new List<string>();
			foreach (var worksheet in package.Workbook.Worksheets)
			{
				nombresDeHojas.Add(worksheet.Name);
			}

			ExcelWorksheet hoja1 = package.Workbook.Worksheets[nombresDeHojas[0]];

			var fecha = db.Panel_2013.Where(x => x.ID == idPanel).First();

			var dataHoja1 = db.Panel_2013_9017.Where(x => x.C_ElementID == idPanel).ToList();

			int filaInicialPlacas = 2;


			foreach (var item in dataHoja1)
			{


				hoja1.Cells["A" + filaInicialPlacas].Value = item.C3_POST_TL_PLACA;
				hoja1.Cells["B" + filaInicialPlacas].Value = item.C3_POST_TL_EJES;
				hoja1.Cells["C" + filaInicialPlacas].Value = item.C3_POST_TL_FABRICANTE;
				hoja1.Cells["D" + filaInicialPlacas].Value = item.C3_POST_TL_MODELO;
				hoja1.Cells["E" + filaInicialPlacas].Value = item.C3_POST_TL_COLOR;
				hoja1.Cells["F" + filaInicialPlacas].Value = item.C3_POST_TL_ESTADO;


				ExcelRange border = hoja1.Cells["A" + filaInicialPlacas + ":F" + filaInicialPlacas];

				// Agregar bordes al rango de celdas
				border.Style.Border.Top.Style = ExcelBorderStyle.Thin;
				border.Style.Border.Left.Style = ExcelBorderStyle.Thin;
				border.Style.Border.Right.Style = ExcelBorderStyle.Thin;
				border.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

				ExcelRange TextCentrar = hoja1.Cells["A" + filaInicialPlacas + ":F" + filaInicialPlacas];

				TextCentrar.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
				TextCentrar.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

				filaInicialPlacas++;
			}



			ExcelWorksheet hoja2 = package.Workbook.Worksheets[nombresDeHojas[1]];


			hoja2.Cells["B8"].Value = fecha.C8_GC_POSTPAGOS_FINALES_3_PT_EMPRESA_ULTIMA_EMPRESA_POST;
			hoja2.Cells["B9"].Value = fecha.C3_FE_FECHA_DE_INICIO_DE_PERIODO + " - " + fecha.C3_FE_FECHA_DE_FIN_DE_PERIODO;


			//hoja2.Cells["G3"].Value = "Saldo al " + fecha.C3_FE_FECHA_DE_INICIO_DE_PERIODO.Value.ToString("dd/MM/yyyy HH:mm:ss");
			//hoja2.Cells["G4"].Value = "Recarga del " + fecha.C3_FE_FECHA_DE_INICIO_DE_PERIODO.Value.ToString("dd") + " al " + fecha.C3_FE_FECHA_DE_FIN_DE_PERIODO.Value.ToString("dd");
			//hoja2.Cells["G5"].Value = "Consumo del " + fecha.C3_FE_FECHA_DE_INICIO_DE_PERIODO.Value.ToString("dd") + " al " + fecha.C3_FE_FECHA_DE_FIN_DE_PERIODO.Value.ToString("dd") + " Covisol";
			//hoja2.Cells["G6"].Value = "Consumo del " + fecha.C3_FE_FECHA_DE_INICIO_DE_PERIODO.Value.ToString("dd") + " al " + fecha.C3_FE_FECHA_DE_FIN_DE_PERIODO.Value.ToString("dd") + " Coviperú";
			//hoja2.Cells["G7"].Value = "Saldo al " + fecha.C3_FE_FECHA_DE_FIN_DE_PERIODO.Value.ToString("dd/MM/yyyy HH:mm:ss");

			//hoja2.Cells["G3"].Value = fecha.C3_ND_SALDODELULTIMOREPORTE;
			//hoja2.Cells["G4"].Value = fecha.C8_PR4_GC_RECIBODECOMPROBANTES_3_POST_MONTORECARGA_POST_ND_SUMARECARGAS;
			hoja2.Cells["G3"].Value = fecha.C8_GC_TRANSITOSGENERALESPOSTPAGO_3_ND_CONSUMOSCOVISOL_POST_SUMADECONSUMOCOVISOL;
			hoja2.Cells["G4"].Value = fecha.C8_GC_TRANSITOSGENERALESPOSTPAGO_3_ND_CONSUMOSCOVISUR_POST_SUMADECONSUMOCOVISUR;
			hoja2.Cells["G5"].Value = fecha.C8_GC_TRANSITOSGENERALESPOSTPAGO_3_ND_CONSUMOSCOVIPERU_POST_SUMADECONSUMOCOVIPER;

			//hoja2.Cells["J3"].Value = fecha.C3_ND_SALDODELULTIMOREPORTE;
			//hoja2.Cells["J4"].Value = fecha.C8_PR4_GC_RECIBODECOMPROBANTES_3_POST_MONTODETRACCION_POST_ND_SUMADETRACCI;
			hoja2.Cells["H5"].Value = fecha.C8_GC_TRANSITOSGENERALESPOSTPAGO_3_ND_DETRACCOVISOL_POST_SUMADETRACCOVISOL;
			hoja2.Cells["H6"].Value = fecha.C8_GC_TRANSITOSGENERALESPOSTPAGO_3_ND_DETRACCOVISUR_POST_SUMADETRACCOVISUR;
			hoja2.Cells["H7"].Value = fecha.C8_GC_TRANSITOSGENERALESPOSTPAGO_3_ND_DETRACCOVIPERU_POST_SUMADETRACCOVIPERU;

			int contadorHoja2 = 11;
			var panel = db.Panel_2013_8765.Where(x => x.C_ElementID == idPanel).ToList();
			int cantidadElementos = 1;
			foreach (var item in panel)
			{
				hoja2.InsertRow(contadorHoja2, 1);
				ExcelRange borderHoja2 = hoja2.Cells["A" + contadorHoja2 + ":K" + contadorHoja2];

				// Agregar bordes al rango de celdas
				borderHoja2.Style.Border.Top.Style = ExcelBorderStyle.Thin;
				borderHoja2.Style.Border.Left.Style = ExcelBorderStyle.Thin;
				borderHoja2.Style.Border.Right.Style = ExcelBorderStyle.Thin;
				borderHoja2.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

				ExcelRange TextCentrarHoja2 = hoja2.Cells["A" + contadorHoja2 + ":K" + contadorHoja2];

				TextCentrarHoja2.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
				TextCentrarHoja2.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

				hoja2.Cells["A" + contadorHoja2].Value = item.C3_POST_FH_FECHAYHORATRANSITO.Value.ToString("dd/MM/yyyy HH:mm:ss");
				hoja2.Cells["B" + contadorHoja2].Value = item.C3_POST_TL_ESTACION;
				hoja2.Cells["C" + contadorHoja2].Value = item.C3_POST_TL_VIA;
				hoja2.Cells["D" + contadorHoja2].Value = item.C3_POST_TL_PLACCA;
				hoja2.Cells["E" + contadorHoja2].Value = item.C3_POST_TL_CATEGORIA;
				hoja2.Cells["F" + contadorHoja2].Value = item.C3_POST_ND_VALORPEAJE;
				hoja2.Cells["F" + contadorHoja2].Style.Numberformat.Format = "\"S/\"#,##0.00";
				hoja2.Cells["G" + contadorHoja2].Value = item.C3_POST_ND_VALORDETRACCION;
				hoja2.Cells["G" + contadorHoja2].Style.Numberformat.Format = "\"S/\"#,##0.00";
				hoja2.Cells["H" + contadorHoja2].Value = (item.C3_POST_FE_FECHACOMPROBANTE == null ? "" : item.C3_POST_FE_FECHACOMPROBANTE.Value.ToString("dd/MM/yyyy HH:mm:ss"));
				hoja2.Cells["I" + contadorHoja2].Value = item.C3_POST_TL_NUMERODECOMPROBANTE;
				hoja2.Cells["J" + contadorHoja2].Value = item.C3_POST_TL_RUCCONCESION;
				hoja2.Cells["K" + contadorHoja2].Value = item.C3_POST_TL_TIPODECOMPROBANTE;
				cantidadElementos++;
				contadorHoja2++;

			}

			hoja2.Cells["F" + contadorHoja2].Formula = "SUBTOTAL(9,F11:F" + (contadorHoja2 - 1) + ")";

			ExcelRange borderHoja2a = hoja2.Cells["F" + contadorHoja2];

			// Agregar bordes al rango de celdas
			borderHoja2a.Style.Border.Top.Style = ExcelBorderStyle.Thin;
			borderHoja2a.Style.Border.Left.Style = ExcelBorderStyle.Thin;
			borderHoja2a.Style.Border.Right.Style = ExcelBorderStyle.Thin;
			borderHoja2a.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;


			ExcelWorksheet hoja3 = package.Workbook.Worksheets[nombresDeHojas[2]];

			var dataHoja3 = db.Panel_2013_8954.Where(x => x.C_ElementID == idPanel).ToList();
			int contadorHoja3 = 3;
			foreach (var item in dataHoja3)
			{
				hoja3.InsertRow(contadorHoja3, 1);


				ExcelRange borderHoja3 = hoja3.Cells["B" + contadorHoja3 + ":D" + contadorHoja3];

				// Agregar bordes al rango de celdas
				borderHoja3.Style.Border.Top.Style = ExcelBorderStyle.Thin;
				borderHoja3.Style.Border.Left.Style = ExcelBorderStyle.Thin;
				borderHoja3.Style.Border.Right.Style = ExcelBorderStyle.Thin;
				borderHoja3.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

				ExcelRange TextCentrarHoja3 = hoja3.Cells["B" + contadorHoja3 + ":D" + contadorHoja3];

				TextCentrarHoja3.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
				TextCentrarHoja3.Style.VerticalAlignment = ExcelVerticalAlignment.Center;


				hoja3.Cells["B" + contadorHoja3].Value = item.C3_POST_MONTORECARGA;
				hoja3.Cells["B" + contadorHoja3].Style.Numberformat.Format = "\"S/\"#,##0.00";
				hoja3.Cells["C" + contadorHoja3].Value = item.C3_POST_MONTODETRACCION;
				hoja3.Cells["C" + contadorHoja3].Style.Numberformat.Format = "\"S/\"#,##0.00";
				hoja3.Cells["D" + contadorHoja3].Value = item.C3_POST_COMPROBANTERECARGA == null ? "" : item.C3_POST_COMPROBANTERECARGA.Value.ToString("dd/MM/yyyy HH:mm:ss");

				//hoja3.Cells["C" + contadorHoja3].Value = item.C3_POST_MONTORECARGA;
				//hoja3.Cells["C" + contadorHoja3].Style.Numberformat.Format = "\"S/\"#,##0.00";
				//hoja3.Cells["D" + contadorHoja3].Value = item.C3_POST_TL_RECIBO;
				contadorHoja3++;
			}
			//hoja3.Cells["C" + contadorHoja3].Formula = "SUM(C3:F" + (contadorHoja3 - 1) + ")";
			return package;
		}

		public ExcelPackage CarguerosPostpagoSinDestraccion(ExcelPackage package, int idPanel)
		{
			List<string> nombresDeHojas = new List<string>();
			foreach (var worksheet in package.Workbook.Worksheets)
			{
				nombresDeHojas.Add(worksheet.Name);
			}

			ExcelWorksheet hoja1 = package.Workbook.Worksheets[nombresDeHojas[0]];

			var fecha = db.Panel_2013.Where(x => x.ID == idPanel).First();

			var dataHoja1 = db.Panel_2013_9017.Where(x => x.C_ElementID == idPanel).ToList();

			int filaInicialPlacas = 2;


			foreach (var item in dataHoja1)
			{


				hoja1.Cells["A" + filaInicialPlacas].Value = item.C3_POST_TL_PLACA;
				hoja1.Cells["B" + filaInicialPlacas].Value = item.C3_POST_TL_EJES;
				hoja1.Cells["C" + filaInicialPlacas].Value = item.C3_POST_TL_FABRICANTE;
				hoja1.Cells["D" + filaInicialPlacas].Value = item.C3_POST_TL_MODELO;
				hoja1.Cells["E" + filaInicialPlacas].Value = item.C3_POST_TL_COLOR;
				hoja1.Cells["F" + filaInicialPlacas].Value = item.C3_POST_TL_ESTADO;


				ExcelRange border = hoja1.Cells["A" + filaInicialPlacas + ":F" + filaInicialPlacas];

				// Agregar bordes al rango de celdas
				border.Style.Border.Top.Style = ExcelBorderStyle.Thin;
				border.Style.Border.Left.Style = ExcelBorderStyle.Thin;
				border.Style.Border.Right.Style = ExcelBorderStyle.Thin;
				border.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

				ExcelRange TextCentrar = hoja1.Cells["A" + filaInicialPlacas + ":F" + filaInicialPlacas];

				TextCentrar.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
				TextCentrar.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

				filaInicialPlacas++;
			}



			ExcelWorksheet hoja2 = package.Workbook.Worksheets[nombresDeHojas[1]];


			hoja2.Cells["B7"].Value = fecha.C8_GC_POSTPAGOS_FINALES_3_PT_EMPRESA_ULTIMA_EMPRESA_POST;
			hoja2.Cells["B8"].Value = fecha.C3_FE_FECHA_DE_INICIO_DE_PERIODO + " - " + fecha.C3_FE_FECHA_DE_FIN_DE_PERIODO;


			//hoja2.Cells["F3"].Value = "Saldo al " + fecha.C3_FE_FECHA_DE_INICIO_DE_PERIODO.Value.ToString("dd/MM/yyyy HH:mm:ss");
			//hoja2.Cells["F4"].Value = "Recarga del " + fecha.C3_FE_FECHA_DE_INICIO_DE_PERIODO.Value.ToString("dd") + " al " + fecha.C3_FE_FECHA_DE_FIN_DE_PERIODO.Value.ToString("dd");
			//hoja2.Cells["F5"].Value = "Consumo del " + fecha.C3_FE_FECHA_DE_INICIO_DE_PERIODO.Value.ToString("dd") + " al " + fecha.C3_FE_FECHA_DE_FIN_DE_PERIODO.Value.ToString("dd") + " Covisol";
			//hoja2.Cells["F6"].Value = "Consumo del " + fecha.C3_FE_FECHA_DE_INICIO_DE_PERIODO.Value.ToString("dd") + " al " + fecha.C3_FE_FECHA_DE_FIN_DE_PERIODO.Value.ToString("dd") + " Coviperú";
			//hoja2.Cells["F7"].Value = "Saldo al " + fecha.C3_FE_FECHA_DE_FIN_DE_PERIODO.Value.ToString("dd/MM/yyyy HH:mm:ss");

			//hoja2.Cells["G3"].Value = fecha.C3_ND_SALDODELULTIMOREPORTE;
			//hoja2.Cells["G4"].Value = fecha.C8_PR4_GC_RECIBODECOMPROBANTES_3_POST_MONTORECARGA_POST_ND_SUMARECARGAS;
			hoja2.Cells["G3"].Value = fecha.C8_GC_TRANSITOSGENERALESPOSTPAGO_3_ND_CONSUMOSCOVISOL_POST_SUMADECONSUMOCOVISOL;
			hoja2.Cells["G4"].Value = fecha.C8_GC_TRANSITOSGENERALESPOSTPAGO_3_ND_CONSUMOSCOVISUR_POST_SUMADECONSUMOCOVISUR;
			hoja2.Cells["G5"].Value = fecha.C8_GC_TRANSITOSGENERALESPOSTPAGO_3_ND_CONSUMOSCOVIPERU_POST_SUMADECONSUMOCOVIPER;


			int contadorHoja2 = 11;
			var panel = db.Panel_2013_8765.Where(x => x.C_ElementID == idPanel).ToList();
			int cantidadElementos = 1;
			foreach (var item in panel)
			{
				hoja2.InsertRow(contadorHoja2, 1);
				ExcelRange borderHoja2 = hoja2.Cells["A" + contadorHoja2 + ":J" + contadorHoja2];

				// Agregar bordes al rango de celdas
				borderHoja2.Style.Border.Top.Style = ExcelBorderStyle.Thin;
				borderHoja2.Style.Border.Left.Style = ExcelBorderStyle.Thin;
				borderHoja2.Style.Border.Right.Style = ExcelBorderStyle.Thin;
				borderHoja2.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

				ExcelRange TextCentrarHoja2 = hoja2.Cells["A" + contadorHoja2 + ":J" + contadorHoja2];

				TextCentrarHoja2.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
				TextCentrarHoja2.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

				hoja2.Cells["A" + contadorHoja2].Value = item.C3_POST_FH_FECHAYHORATRANSITO.Value.ToString("dd/MM/yyyy HH:mm:ss");
				hoja2.Cells["B" + contadorHoja2].Value = item.C3_POST_TL_ESTACION;
				hoja2.Cells["C" + contadorHoja2].Value = item.C3_POST_TL_VIA;
				hoja2.Cells["D" + contadorHoja2].Value = item.C3_POST_TL_PLACCA;
				hoja2.Cells["E" + contadorHoja2].Value = item.C3_POST_TL_CATEGORIA;
				hoja2.Cells["F" + contadorHoja2].Value = item.C3_POST_ND_VALORPEAJE;
				hoja2.Cells["F" + contadorHoja2].Style.Numberformat.Format = "\"S/\"#,##0.00";
				hoja2.Cells["G" + contadorHoja2].Value = (item.C3_POST_FE_FECHACOMPROBANTE == null ? "" : item.C3_POST_FE_FECHACOMPROBANTE.Value.ToString("dd/MM/yyyy HH:mm:ss"));
				hoja2.Cells["H" + contadorHoja2].Value = item.C3_POST_TL_NUMERODECOMPROBANTE;
				hoja2.Cells["I" + contadorHoja2].Value = item.C3_POST_TL_RUCCONCESION;
				hoja2.Cells["J" + contadorHoja2].Value = item.C3_POST_TL_TIPODECOMPROBANTE;
				cantidadElementos++;
				contadorHoja2++;

			}

			hoja2.Cells["F" + contadorHoja2].Formula = "SUBTOTAL(9,F11:F" + (contadorHoja2 - 1) + ")";

			ExcelRange borderHoja2a = hoja2.Cells["F" + contadorHoja2];

			// Agregar bordes al rango de celdas
			borderHoja2a.Style.Border.Top.Style = ExcelBorderStyle.Thin;
			borderHoja2a.Style.Border.Left.Style = ExcelBorderStyle.Thin;
			borderHoja2a.Style.Border.Right.Style = ExcelBorderStyle.Thin;
			borderHoja2a.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;


			ExcelWorksheet hoja3 = package.Workbook.Worksheets[nombresDeHojas[2]];

			var dataHoja3 = db.Panel_2013_8954.Where(x => x.C_ElementID == idPanel).ToList();
			int contadorHoja3 = 3;
			foreach (var item in dataHoja3)
			{
				hoja3.InsertRow(contadorHoja3, 1);


				ExcelRange borderHoja3 = hoja3.Cells["B" + contadorHoja3 + ":E" + contadorHoja3];

				// Agregar bordes al rango de celdas
				borderHoja3.Style.Border.Top.Style = ExcelBorderStyle.Thin;
				borderHoja3.Style.Border.Left.Style = ExcelBorderStyle.Thin;
				borderHoja3.Style.Border.Right.Style = ExcelBorderStyle.Thin;
				borderHoja3.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

				ExcelRange TextCentrarHoja3 = hoja3.Cells["B" + contadorHoja3 + ":E" + contadorHoja3];

				TextCentrarHoja3.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
				TextCentrarHoja3.Style.VerticalAlignment = ExcelVerticalAlignment.Center;


				hoja3.Cells["B" + contadorHoja3].Value = item.C3_POST_MONTORECARGA;
				hoja3.Cells["B" + contadorHoja3].Style.Numberformat.Format = "\"S/\"#,##0.00";
				hoja3.Cells["C" + contadorHoja3].Value = item.C3_POST_COMPROBANTERECARGA == null ? "" : item.C3_POST_COMPROBANTERECARGA.Value.ToString("dd/MM/yyyy HH:mm:ss");
				hoja3.Cells["D" + contadorHoja3].Value = item.C3_POST_TL_RECIBO;
				hoja3.Cells["E" + contadorHoja3].Value = item.C3_POST_TL_CUENTAIDENTIFICADOR;
				
				contadorHoja3++;
			}
			//hoja3.Cells["C" + contadorHoja3].Formula = "SUM(C3:F" + (contadorHoja3 - 1) + ")";
			return package;
		}
	}
}