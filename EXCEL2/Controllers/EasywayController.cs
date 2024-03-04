using DocumentFormat.OpenXml.Office2019.Excel.RichData2;
using EXCEL2.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.UI.WebControls;
using WebGrease.Css.Ast.Selectors;

namespace EXCEL2.Controllers
{
    public class EasywayController : Controller
    {

		private AuraPortal_BPMSEntities db = new AuraPortal_BPMSEntities();
		// GET: Easyway
		public ActionResult Index()
        {
            return View();
        }

		[HttpGet]
        public ActionResult LiquidacionSemanal(int idPanel)
        {
			ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
			// Crear un nuevo paquete Excel
			byte[] bytesDelExcel = ObtenerDatosExcelLiquidacionSemanal(idPanel); // Reemplaza esto con tus propios datos
																		  // Crear un MemoryStream a partir de los bytes del Excel
			using (MemoryStream memoryStream = new MemoryStream(bytesDelExcel))
			{
				// Crear un paquete Excel a partir del MemoryStream
				using (ExcelPackage package = new ExcelPackage(memoryStream))
				{
					// Obtener el contenido del paquete en un array de bytes
					byte[] excelBytes = ExcelLiquidacionSemanal(package, idPanel).GetAsByteArray();

					//return Json(UpdateDatosExcelLiquidacionSemanal(excelBytes,idPanel));


					// Devolver el archivo Excel al cliente como descarga
					return File(excelBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "archivo_excel.xlsx");
				}
			}
        }


		public byte[] ObtenerDatosExcelLiquidacionSemanal(int idPanel)
		{
			var familia = db.Panel_1008.Where( x => x.ID == idPanel).ToList().Last();
			var integrate = db.AP__DocIntegratedStorage.Where(x => x.AscTermId == 4570).Where(x => x.IntegrationObjectId == familia.C_ElementID).Where(x => x.ObjectTypeProcessId == 4795).First();
			var integrateData = db.AP__DocIntegratedDataStorage.Where(x => x.AscDocId == integrate.ID).First();
			return integrateData.Content;
		}

		public byte[] UpdateDatosExcelLiquidacionSemanal(byte[] excelBytes, int idPanel)
		{
			var familia = db.Panel_1008.Where(x => x.ID == idPanel).ToList().Last();
			var integrate = db.AP__DocIntegratedStorage.Where(x => x.AscTermId == 4570).Where(x => x.IntegrationObjectId == familia.C_ElementID).Where(x => x.ObjectTypeProcessId == 4795).First();
			var integrateData = db.AP__DocIntegratedDataStorage.Where(x => x.AscDocId == integrate.ID).First();

			integrateData.Content = excelBytes;
			db.Entry(integrateData).State = EntityState.Modified;
			db.SaveChanges();
			return integrateData.Content;
		}

		public ExcelPackage ExcelLiquidacionSemanal(ExcelPackage package, int idPanel)
		{
			List<string> nombresDeHojas = new List<string>();
			foreach (var worksheet in package.Workbook.Worksheets)
			{
				nombresDeHojas.Add(worksheet.Name);
			}
			ExcelWorksheet hoja1 = package.Workbook.Worksheets[nombresDeHojas[0]];


			var data = db.Panel_1008_4700.Where(x => x.C_ElementID == idPanel).ToArray();

			int finalInicial = 2;

			var ordenado = data.OrderBy(item => item.C3_FE_FECHA_COMPROBANTE_RT)
						   .ThenBy(item => item.C3_TL_CUENTA_RT)
						   .ThenBy(item => item.C3_TL_NOMBRE_CLIENTE_RT);

			foreach (var item in ordenado)
            {
				hoja1.Cells["A" + finalInicial].Value = item.C3_FE_FECHA_COMPROBANTE_RT.Value.ToString("dd/MM/yyyy");
				int cuenta = Int32.Parse(item.C3_TL_CUENTA_RT);
				hoja1.Cells["B" + finalInicial].Value = cuenta.ToString("000000000");
				hoja1.Cells["C" + finalInicial].Value = item.C3_TL_NOMBRE_CLIENTE_RT;
				hoja1.Cells["D" + finalInicial].Value = item.C3_TL_RUC_CLIENTE_RT;
				hoja1.Cells["E" + finalInicial].Value = item.C3_TL_TIPO_COMPROBANTE_RT;
				hoja1.Cells["F" + finalInicial].Value = item.C3_TL_SERIE_RT;
				hoja1.Cells["G" + finalInicial].Value = item.C3_TL_CORRELATIVO_RT;
				hoja1.Cells["H" + finalInicial].Value = item.C3_ND_PEAJE_TOTAL_RT;
				hoja1.Cells["H" + finalInicial].Style.Numberformat.Format = "#,##0.00";
				hoja1.Cells["I" + finalInicial].Value = item.C3_ND_DETRACCION_TOTAL_RT;
				hoja1.Cells["A"+ finalInicial + ":I"+ finalInicial].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
				hoja1.Cells["C" + finalInicial].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
				ExcelRange border = hoja1.Cells["A" + finalInicial + ":I" + finalInicial];
				// Agregar bordes al rango de celdas
				border.Style.Border.Top.Style = ExcelBorderStyle.Thin;
				border.Style.Border.Left.Style = ExcelBorderStyle.Thin;
				border.Style.Border.Right.Style = ExcelBorderStyle.Thin;
				border.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
				finalInicial++;
			}

			//finalInicial++;
			hoja1.Cells["H" + (finalInicial)].Value = ordenado.Sum(x => x.C3_ND_PEAJE_TOTAL_RT);
			hoja1.Cells["H" + (finalInicial)].Style.Numberformat.Format = "\"S/\"#,##0.00";

			hoja1.Cells["I" + (finalInicial)].Value = ordenado.Sum(x => x.C3_ND_DETRACCION_TOTAL_RT);
			hoja1.Cells["I" + (finalInicial)].Style.Numberformat.Format = "\"S/\"#,##0.00";

			ExcelRange border2 = hoja1.Cells["H" + finalInicial + ":I" + finalInicial];
			// Agregar bordes al rango de celdas
			border2.Style.Border.Top.Style = ExcelBorderStyle.Thin;
			border2.Style.Border.Left.Style = ExcelBorderStyle.Thin;
			border2.Style.Border.Right.Style = ExcelBorderStyle.Thin;
			border2.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

			finalInicial = finalInicial + 2;

			var descuentoTransitos = db.Panel_1008_4170.Where(x => x.C_ElementID == idPanel).ToList();
			var abonoTransito = db.Panel_1008_4233.Where(x => x.C_ElementID == idPanel).ToList();


			var contadorDescuentosTransito = finalInicial;

			foreach (var item in descuentoTransitos)
            {
				hoja1.Cells["G" + (contadorDescuentosTransito)].Value = item.C3_T_Descripcion;
				hoja1.Cells["H" + (contadorDescuentosTransito)].Value = item.C3_T_MONTO;
				hoja1.Cells["H" + (contadorDescuentosTransito)].Style.Numberformat.Format = "\"S/\"#,##0.00";
				contadorDescuentosTransito++;
			}

			var contadorAbonoTransito = contadorDescuentosTransito;
			var contador = 1;
			foreach (var item in abonoTransito)
			{
				hoja1.Cells["F" + (contadorAbonoTransito)].Value = contador + "°";
				hoja1.Cells["G" + (contadorAbonoTransito)].Value = "Abono "+ item.C3_T_FECHA.Value.ToString("dd/MM");
				hoja1.Cells["H" + (contadorAbonoTransito)].Value = item.C3_T_MONTO_BONO_T;
				hoja1.Cells["H" + (contadorAbonoTransito)].Style.Numberformat.Format = "\"S/\"#,##0.00";
				contadorAbonoTransito++;
				contador++;
			}

			var DescuentosDetraccion = db.Panel_1008_4450.Where(x => x.C_ElementID == idPanel).ToList();
			var AbonoDetraccion = db.Panel_1008_4261.Where(x => x.C_ElementID == idPanel).ToList();

			var contadorDescuentosDetraccion = finalInicial;

			foreach (var item in DescuentosDetraccion)
            {
				hoja1.Cells["I" + (contadorDescuentosDetraccion)].Value = item.C3_T_MONTO_DET;
				hoja1.Cells["I" + (contadorDescuentosDetraccion)].Style.Numberformat.Format = "\"S/\"#,##0.00";
				hoja1.Cells["J" + (contadorDescuentosDetraccion)].Value = item.C3_T_DESCRIPCIÓN_DET;
				contadorDescuentosDetraccion++;
			}

			var contadorFinal = 0;

			if(contadorDescuentosDetraccion < contadorDescuentosTransito)
			{
				contadorFinal = contadorDescuentosTransito;
			}
			else
			{
				contadorFinal = contadorDescuentosDetraccion;
			}

            foreach (var item in AbonoDetraccion)
            {
				hoja1.Cells["I" + (contadorFinal)].Value = item.C3_PR3_MONTO;
				hoja1.Cells["I" + (contadorFinal)].Style.Numberformat.Format = "\"S/\"#,##0.00";
				hoja1.Cells["J" + (contadorFinal)].Value = "Abono " + item.C3_PR3_FECHA_DET.Value.ToString("dd/MM");
				contadorFinal++;

			}



            return package;
		}

		

	}
}