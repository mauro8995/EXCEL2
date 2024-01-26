using EXCEL2.Models.Listas;
using OfficeOpenXml.Style;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using EXCEL2.Models;

namespace EXCEL2.Controllers
{

	public class CoviPeruController : Controller
    {


		private AuraPortal_BPMSEntities db = new AuraPortal_BPMSEntities();


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

		//---------------------------------------------------- Excel con visa


		[HttpGet]
		public ActionResult GenerarArchivoVisa(int idPanel)
		{
			ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
			// Crear un nuevo paquete Excel
			byte[] bytesDelExcel = ObtenerDatosVarbinaryVisa(); // Reemplaza esto con tus propios datos

			// Crear un MemoryStream a partir de los bytes del Excel
			using (MemoryStream memoryStream = new MemoryStream(bytesDelExcel))
			{
				// Crear un paquete Excel a partir del MemoryStream
				using (ExcelPackage package = new ExcelPackage(memoryStream))
				{
					// Obtener el contenido del paquete en un array de bytes
					byte[] excelBytes = VisaExcel(package, idPanel).GetAsByteArray();

					return Json(UpdateDatosVarbinaryVisa(excelBytes));


					// Devolver el archivo Excel al cliente como descarga
					return File(excelBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "archivo_excel.xlsx");
				}
			}
		}


		public byte[] ObtenerDatosVarbinaryVisa()
		{
			var familia = db.AP_Dyn_Familias_30.ToList().Last();
			var integrate = db.AP__DocIntegratedStorage.Where(x => x.AscTermId == 3219).Where(x => x.IntegrationObjectId == familia.ID).Where(x => x.ObjectTypeProcessId == 30).First();
			var integrateData = db.AP__DocIntegratedDataStorage.Where(x => x.AscDocId == integrate.ID).First();
			return integrateData.Content;

		}

		public byte[] UpdateDatosVarbinaryVisa(byte[] excelBytes)
		{
			var familia = db.AP_Dyn_Familias_30.ToList().Last();
			var integrate = db.AP__DocIntegratedStorage.Where(x => x.AscTermId == 3219).Where(x => x.IntegrationObjectId == familia.ID).Where(x => x.ObjectTypeProcessId == 30).First();
			var integrateData = db.AP__DocIntegratedDataStorage.Where(x => x.AscDocId == integrate.ID).First();

			integrateData.Content = excelBytes;
			db.Entry(integrateData).State = EntityState.Modified;
			db.SaveChanges();
			return integrateData.Content;
		}


		public ExcelPackage VisaExcel(ExcelPackage package, int idPanel)
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


			var FechaInicio = fecha1.ToString("dd MMMM", new CultureInfo("es-ES"));
			var FechaFin = fecha2.ToString("dd MMMM yyyy", new CultureInfo("es-ES"));


			int filaActual = 6;
			int columnaIndex = 1;

			while (hoja1.Cells[filaActual, columnaIndex].Value != null)
			{
				// Mover a la siguiente fila
				filaActual++;
			}

			int numeroEntero = (int)(filtro.C3_PR3_ITEM_TRANSFERENCIA);
			int nuevaFila = filaActual;

			hoja1.InsertRow(nuevaFila, 1);
			hoja1.Cells["B" + nuevaFila].Value = "Del " + FechaInicio + " al " + FechaFin;
			hoja1.Cells["D" + nuevaFila].Value = filtro.C8_A__PR3_ABONO_TRANSITO_3_T_FECHA_Fecha_tranfs_min.Value.ToString("dd/MM/yyyy");//Fecha de trasferencia
			hoja1.Cells["A" + nuevaFila].Value = numeroEntero.ToString("00");//item
			hoja1.Cells["A" + nuevaFila.ToString()].Style.Font.Bold = true; // Poner en negrita
			hoja1.Cells["A" + nuevaFila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // Centrar el contenido


			if (filtro.C3_PR3_REGULARIZAR_OTROS_DESCUENTOS != null)
			{
				hoja1.Cells["F" + nuevaFila].Value = filtro.C3_PR3_REGULARIZAR_OTROS_DESCUENTOS;//Descuento
				hoja1.Cells["F" + nuevaFila].Style.Numberformat.Format = "\"S/\"#,##0.00";
			}

			var lista_hoja = db.Panel_1008_4700.Where(x => x.C_ElementID == idPanel).ToList();


			var familia = db.AP_Dyn_Familias_30.OrderBy(x => x.C_Name).ToList().Last();
			var anio = Int32.Parse(familia.C_Name);

			List<ExcelMes> meses = mesesList(anio);


			var gruposPorMes = lista_hoja.GroupBy(d => new { Mes = d.C3_FE_FECHA_COMPROBANTE_RT.Value.Month, Anio = d.C3_FE_FECHA_COMPROBANTE_RT.Value.Year }).Select(g => new
			{
				Anio = g.Key.Anio,
				Mes = g.Key.Mes,
				Suma = g.Sum(d => d.C3_ND_PEAJE_TOTAL_RT)
			}).ToArray();



			foreach (var grupo in gruposPorMes)
			{
				foreach (var item2 in meses)
				{
					if (item2.MesNumero == grupo.Mes.ToString("00") && item2.Anio == grupo.Anio)
					{
						hoja1.Cells[item2.Posicion + nuevaFila.ToString()].Value = grupo.Suma;
						hoja1.Cells[item2.Posicion + nuevaFila.ToString()].Style.Numberformat.Format = "\"S/\"#,##0.00";
						//hoja1.Cells[item2.Posicion + (nuevaFila+2)].Formula = "=sum("+ item2.Posicion + 8 + ":"+ item2.Posicion + (nuevaFila )+ ")";
					}
				}
			}

			hoja1.Cells["E" + nuevaFila].Formula = "=sum(F" + nuevaFila + ":T" + nuevaFila + ")";
			hoja1.Cells["E" + nuevaFila].Style.Font.Bold = true;
			hoja1.Cells["E" + nuevaFila].Style.Numberformat.Format = "\"S/\"#,##0.00";

			foreach (var item in meses)
			{
				if (filtro.C3_PR3_REGULARIZAR_AÑO == item.Anio && filtro.C3_PR3_REGULARIZAR_MES == Int32.Parse(item.MesNumero))
				{
					hoja1.Cells[item.Posicion + nuevaFila.ToString()].Value = filtro.C3_PR3_REGULARIZAR_MONTO;
				}
			}




			//-------------- hoja 2
			ExcelWorksheet hoja2 = package.Workbook.Worksheets[nombresDeHojas[1]];
			int filaActual_hoja2 = 6;
			int columnaIndex_hoja2 = 1;

			while (hoja2.Cells[filaActual_hoja2, columnaIndex_hoja2].Value != null)
			{
				// Mover a la siguiente fila
				filaActual_hoja2++;
			}

			int numeroEntero_hoja2 = (int)(filtro.C3_PR3_ITEM_TRANSFERENCIA);
			int nuevaFila_hoja2 = filaActual_hoja2;

			hoja2.InsertRow(nuevaFila_hoja2, 1);
			hoja2.Cells["B" + nuevaFila_hoja2].Value = "Del " + FechaInicio + " al " + FechaFin;
			hoja2.Cells["D" + nuevaFila_hoja2].Value = filtro.C8_A__PR3_ABONO_DETRACCION_3_PR3_FECHA_DET_Fecha_tranfs_min.Value.ToString("dd/MM/yyyy");//Fecha de trasferencia
			hoja2.Cells["A" + nuevaFila_hoja2].Value = numeroEntero_hoja2.ToString("00");//item
			hoja2.Cells["A" + nuevaFila_hoja2.ToString()].Style.Font.Bold = true; // Poner en negrita
			hoja2.Cells["A" + nuevaFila_hoja2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // Centrar el contenido


			if (filtro.C3_PR3_REGULARIZAR_OTROS_DESCUENTOS_DET != null)
			{
				hoja2.Cells["F" + nuevaFila_hoja2].Value = filtro.C3_PR3_REGULARIZAR_OTROS_DESCUENTOS_DET;//Descuento
				hoja2.Cells["F" + nuevaFila_hoja2].Style.Numberformat.Format = "\"S/\"#,##0.00";
			}



			var gruposPorMes_hoja2 = lista_hoja.GroupBy(d => new { Mes = d.C3_FE_FECHA_COMPROBANTE_RT.Value.Month, Anio = d.C3_FE_FECHA_COMPROBANTE_RT.Value.Year }).Select(g => new
			{
				Anio = g.Key.Anio,
				Mes = g.Key.Mes,
				Suma = g.Sum(d => d.C3_ND_DETRACCION_TOTAL_RT)
			}).ToArray();



			foreach (var grupo in gruposPorMes_hoja2)
			{
				foreach (var item2 in meses)
				{
					if (item2.MesNumero == grupo.Mes.ToString("00") && item2.Anio == grupo.Anio)
					{
						hoja2.Cells[item2.Posicion + nuevaFila_hoja2.ToString()].Value = grupo.Suma;
						hoja2.Cells[item2.Posicion + nuevaFila_hoja2.ToString()].Style.Numberformat.Format = "\"S/\"#,##0.00";
						//hoja1.Cells[item2.Posicion + (nuevaFila+2)].Formula = "=sum("+ item2.Posicion + 8 + ":"+ item2.Posicion + (nuevaFila )+ ")";
					}
				}
			}

			hoja2.Cells["E" + nuevaFila_hoja2].Formula = "=sum(F" + nuevaFila_hoja2 + ":T" + nuevaFila_hoja2 + ")";
			hoja2.Cells["E" + nuevaFila_hoja2].Style.Font.Bold = true;
			hoja2.Cells["E" + nuevaFila_hoja2].Style.Numberformat.Format = "\"S/\"#,##0.00";

			//-------------- hoja 2
			ExcelWorksheet hoja3 = package.Workbook.Worksheets[nombresDeHojas[2]];
			int filaActual_hoja3 = 6;
			int columnaIndex_hoja3 = 1;

			while (hoja3.Cells[filaActual_hoja3, columnaIndex_hoja3].Value != null)
			{
				// Mover a la siguiente fila
				filaActual_hoja3++;
			}

			int numeroEntero_hoja3 = (int)(filtro.C3_PR3_ITEM_TRANSFERENCIA);
			int nuevaFila_hoja3 = filaActual_hoja3;

			hoja3.InsertRow(nuevaFila_hoja3, 1);
			hoja3.Cells["B" + nuevaFila_hoja3].Value = "Del " + FechaInicio + " al " + FechaFin;
			hoja3.Cells["D" + nuevaFila_hoja3].Value = filtro.C8_A__PR3_ABONO_TRANSITO_3_T_FECHA_Fecha_tranfs_min.Value.ToString("dd/MM/yyyy");//Fecha de trasferencia
			hoja3.Cells["A" + nuevaFila_hoja3].Value = numeroEntero_hoja3.ToString("00");//item
			hoja3.Cells["A" + nuevaFila_hoja3.ToString()].Style.Font.Bold = true; // Poner en negrita
			hoja3.Cells["A" + nuevaFila_hoja3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // Centrar el contenido


			if (filtro.C3_PR3_REGULARIZAR_OTROS_DESCUENTOS_VISA != null)
			{
				hoja3.Cells["F" + nuevaFila_hoja3].Value = filtro.C3_PR3_REGULARIZAR_OTROS_DESCUENTOS_VISA;//Descuento
				hoja3.Cells["F" + nuevaFila_hoja3].Style.Numberformat.Format = "\"S/\"#,##0.00";
			}

			var lista_hoja_hoja_3 = db.Panel_1008_7074.Where(x => x.C_ElementID == idPanel).ToList();

			var gruposPorMes_hoja3 = lista_hoja_hoja_3.GroupBy(d => new { Mes = d.C3_FE_FECHA_VISA.Value.Month, Anio = d.C3_FE_FECHA_VISA.Value.Year }).Select(g => new
			{
				Anio = g.Key.Anio,
				Mes = g.Key.Mes,
				Suma = g.Sum(d => d.C3_ND_IMPORTE_VISA)
			}).ToArray();



			foreach (var grupo in gruposPorMes_hoja3)
			{
				foreach (var item2 in meses)
				{
					if (item2.MesNumero == grupo.Mes.ToString("00") && item2.Anio == grupo.Anio)
					{
						hoja3.Cells[item2.Posicion + nuevaFila_hoja3.ToString()].Value = grupo.Suma;
						hoja3.Cells[item2.Posicion + nuevaFila_hoja3.ToString()].Style.Numberformat.Format = "\"S/\"#,##0.00";
						//hoja1.Cells[item2.Posicion + (nuevaFila+2)].Formula = "=sum("+ item2.Posicion + 8 + ":"+ item2.Posicion + (nuevaFila )+ ")";
					}
				}
			}

			hoja3.Cells["E" + nuevaFila_hoja3].Formula = "=sum(F" + nuevaFila_hoja3 + ":T" + nuevaFila_hoja3 + ")";
			hoja3.Cells["E" + nuevaFila_hoja3].Style.Font.Bold = true;
			hoja3.Cells["E" + nuevaFila_hoja3].Style.Numberformat.Format = "\"S/\"#,##0.00";



			return package;

		}



		//---------------------------------------------------------------------------COMICIONES


		[HttpGet]
		public ActionResult GenerarArchivoComiciones(int idPanel)
		{
			ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
			// Crear un nuevo paquete Excel
			byte[] bytesDelExcel = ObtenerDatosVarbinaryVisa(); // Reemplaza esto con tus propios datos

			// Crear un MemoryStream a partir de los bytes del Excel
			using (MemoryStream memoryStream = new MemoryStream(bytesDelExcel))
			{
				// Crear un paquete Excel a partir del MemoryStream
				using (ExcelPackage package = new ExcelPackage(memoryStream))
				{

					// Obtener el contenido del paquete en un array de bytes
					byte[] excelBytes = Comisiones(package, idPanel).GetAsByteArray();

					return Json(UpdateDatosVarbinaryVisa(excelBytes));
					// Devolver el archivo Excel al cliente como descarga
					return File(excelBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "archivo_excel.xlsx");
				}
			}
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
			var familia = db.AP_Dyn_Familias_30.OrderBy(x => x.C_Name).ToList().Last();

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

			//------------------hoja 2

			ExcelWorksheet hoja2 = package.Workbook.Worksheets[nombresDeHojas[1]];

			int filaActual_hoja2 = 6;

			while (hoja1.Cells[filaActual_hoja2, 1].Value != null)
			{
				filaActual_hoja2++;
			}

			var startCellHoja2 = hoja2.Cells["G" + (filaActual_hoja2)];
			var endCell_hoja2 = hoja2.Cells["U" + (filaActual_hoja2)];
			var range_hoja2 = hoja2.Cells[startCell.Address + ":" + endCell.Address];
			bool insertarFila_hoja2 = false;
			foreach (var cell in range_hoja2)
			{
				// Verifica si la celda tiene contenido
				if (cell.Text != "")
				{
					insertarFila = true;
					break;
				}
			}

			if (insertarFila_hoja2)
			{
				filaActual_hoja2++;
			}


			//var Comision_hoja2 = db.Panel_1009.Where(x => x.C3_PR3_ID_PANEL == idPanel).ToList();

			foreach (var item in Comision)
			{
				foreach (var item2 in meses)
				{
					if (item2.MesNumero == item.C3_PR3_MES_TRANSFERENCIA.Value.ToString("00"))
					{
						if (item2.Anio == item.C3_PR3_EN_AÑO)
						{
							hoja2.Cells[item2.Posicion + '6'].Value = item.C8_GC_COMPROBANTES_TRANSITOS_3_ND_DETRACCION_TOTAL_RT_ND_TOTAL_DETRACCION;
							hoja2.Cells[item2.Posicion + '6'].Style.Numberformat.Format = "\"S/\"#,##0.00";
							//string formula_1 = "=SUM(" + item2.Posicion + "7:" + item2.Posicion.ToString() + (filaActual) + ")";
							//hoja1.Cells[item2.Posicion + (filaActual + 2)].Formula = formula_1;
							//hoja1.Cells[item2.Posicion + (filaActual + 3)].Formula = "=" + item2.Posicion + (filaActual + 2) + "-" + item2.Posicion + "6";
						}
					}
				}
			}

			//------------------hoja 3

			ExcelWorksheet hoja3 = package.Workbook.Worksheets[nombresDeHojas[2]];

			int filaActual_hoja3 = 6;

			while (hoja1.Cells[filaActual_hoja3, 1].Value != null)
			{
				filaActual_hoja3++;
			}

			var startCellHoja3 = hoja3.Cells["G" + (filaActual_hoja3)];
			var endCell_hoja3 = hoja3.Cells["U" + (filaActual_hoja3)];
			var range_hoja3 = hoja3.Cells[startCell.Address + ":" + endCell.Address];
			bool insertarFila_hoja3 = false;
			foreach (var cell in range_hoja2)
			{
				// Verifica si la celda tiene contenido
				if (cell.Text != "")
				{
					insertarFila = true;
					break;
				}
			}

			if (insertarFila_hoja3)
			{
				filaActual_hoja3++;
			}


			//var Comision_hoja2 = db.Panel_1009.Where(x => x.C3_PR3_ID_PANEL == idPanel).ToList();

			foreach (var item in Comision)
			{
				foreach (var item2 in meses)
				{
					if (item2.MesNumero == item.C3_PR3_MES_TRANSFERENCIA.Value.ToString("00"))
					{
						if (item2.Anio == item.C3_PR3_EN_AÑO)
						{
							hoja3.Cells[item2.Posicion + '6'].Value = item.C3_PR3_COM_FA_RECAUDACION_VISA;
							hoja3.Cells[item2.Posicion + '6'].Style.Numberformat.Format = "\"S/\"#,##0.00";
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