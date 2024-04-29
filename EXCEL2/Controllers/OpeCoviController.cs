using DocumentFormat.OpenXml.Spreadsheet;
using EXCEL2.Models;
using EXCEL2.Models.Listas;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Security.Cryptography.Xml;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using System.Web.UI.WebControls;

namespace EXCEL2.Controllers
{
    public class OpeCoviController : Controller
    {
		private AuraPortal_BPMSEntities db = new AuraPortal_BPMSEntities();
		// GET: OpeCovi

		[HttpGet]
		public ActionResult IngresarComprobantes(int idPanel,string Referencia)
		{
			List<TipoComprobantes> ListOb = new List<TipoComprobantes>();
			TipoComprobantes ob1 = new TipoComprobantes
			{
				serie = "FE01",
				modalidad = 2421,
				detraccion = 0,
				Nombre = "FACTURACION_LIMA_PREPAGO_SIN_DETRACCION"
			};
			ListOb.Add(ob1);

			TipoComprobantes ob2 = new TipoComprobantes
			{
				serie = "FE01",
				modalidad = 2421,
				detraccion = 1,
				Nombre = "FACTURACION_LIMA_PREPAGO_CON_DETRACCION"
			};
			ListOb.Add(ob2);

			TipoComprobantes ob3 = new TipoComprobantes
			{
				serie = "FE01",
				modalidad = 2422,
				detraccion = 0,
				Nombre = "FACTURACION_LIMA_POSTPAGO_SIN_DETRACCION"
			};
			ListOb.Add(ob3);

			TipoComprobantes ob4 = new TipoComprobantes
			{
				serie = "FE01",
				modalidad = 2422,
				detraccion = 1,
				Nombre = "FACTURACION_LIMA_POSTPAGO_CON_DETRACCION"
			};
			ListOb.Add(ob4);

			TipoComprobantes ob5 = new TipoComprobantes
			{
				serie = "FE02",
				modalidad = 2421,
				detraccion = 0,
				Nombre = "FACTURACION_CHILCA_PREPAGO_SIN_DETRACCION"
			};
			ListOb.Add(ob5);

			TipoComprobantes ob6 = new TipoComprobantes
			{
				serie = "FE02",
				modalidad = 2421,
				detraccion = 1,
				Nombre = "FACTURACION_CHILCA_PREPAGO_CON_DETRACCION"
			};
			ListOb.Add(ob6);

			TipoComprobantes ob7 = new TipoComprobantes
			{
				serie = "FE02",
				modalidad = 2422,
				detraccion = 0,
				Nombre = "FACTURACION_CHILCA_POSTPAGO_SIN_DETRACCION"
			};
			ListOb.Add(ob7);

			TipoComprobantes ob8 = new TipoComprobantes
			{
				serie = "FE02",
				modalidad = 2422,
				detraccion = 1,
				Nombre = "FACTURACION_CHILCA_POSTPAGO_CON_DETRACCION"
			};
			ListOb.Add(ob8);


			TipoComprobantes ob9 = new TipoComprobantes
			{
				serie = "FE12",
				modalidad = 2421,
				detraccion = 0,
				Nombre = "FACTURACION_TRUJILLO_PREPAGO_SIN_DETRACCION"
			};
			ListOb.Add(ob9);

			TipoComprobantes ob10 = new TipoComprobantes
			{
				serie = "FE12",
				modalidad = 2421,
				detraccion = 1,
				Nombre = "FACTURACION_TRUJILLO_PREPAGO_CON_DETRACCION"
			};
			ListOb.Add(ob10);

			TipoComprobantes ob11 = new TipoComprobantes
			{
				serie = "FE12",
				modalidad = 2422,
				detraccion = 0,
				Nombre = "FACTURACION_TRUJILLO_POSTPAGO_SIN_DETRACCION"
			};
			ListOb.Add(ob11);

			TipoComprobantes ob12 = new TipoComprobantes
			{
				serie = "FE12",
				modalidad = 2422,
				detraccion = 1,
				Nombre = "FACTURACION_TRUJILLO_POSTPAGO_COM_DETRACCION"
			};
			ListOb.Add(ob12);

			var stringjson = "";
            foreach (var item in ListOb)
            {
				var lista1 = db.Panel_2011_8800.Where(x => x.C3_PR3_FC_SERIE == item.serie)
				.Where(x => x.C3_PR3_FC_MODALIDAD == item.modalidad)
				.Where(x => x.C_ElementID == idPanel)
				.Where(x => x.C3_PR3_FC_DETRACCION == item.detraccion).ToList();
				if(lista1.Count > 0)
				{
					double cantida = (lista1.Count / 30);
					if((lista1.Count % 30) != 0)
					{
						cantida++;
					}
					int pagina = (int)Math.Round(cantida, 1);


                    for (int i = 1; i <= pagina; i++)
                    {
						string dataUrl = $"{Request.Url.Scheme}://{Request.Url.Authority}" + "/OpeCovi/Comprobantes?idPanel=" + idPanel + "&modalidad=" + item.modalidad + "&detraccion=" + item.detraccion + "&serie=" + item.serie + "&nombre=" + item.Nombre+ "&pagina="+i;
						stringjson = stringjson + @"
						{""Row"": [
                                {
                                    ""InternalName"": ""3_PR3_ARCHIVO"",
                                    ""Type"": ""Text"",
                                    ""TextValue"": """ + dataUrl + @"""
                                },
                                {
                                    ""InternalName"": ""3_PR3_FLAG"",
                                    ""Type"": ""Integer"",
                                    ""IntegerValue"": 0
                                },
                                {
                                    ""InternalName"": ""3_ESTADO_FAC"",
                                    ""Type"": ""SimpleSelection"",
                                    ""SimpleSelectionValue"": 2435
                                }
                            ] ";
						if(pagina == i)
						{
							stringjson = stringjson + "},";
						}
						else
						{
							stringjson = stringjson + "},";
						}

						
					}

					

				}
			
			}
			
			var Cadena = stringjson.Substring(0, stringjson.Length - 1);
			Task<string> task = insertDataWSAuraAsync(Referencia, Cadena);

			var data = new
			{
				Message = "task",
				Object = "Success"
			};
			return Json(data, JsonRequestBehavior.AllowGet);
		}
		public ActionResult Comprobantes(int idPanel, int modalidad, int detraccion, string serie,string nombre,int pagina)
		{
			ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
			byte[] bytesDelExcel = null;
			// Crear un nuevo paquete Excel
			if (detraccion == 1)
			{
				 bytesDelExcel = ObtenerDatosComprobantesCon(idPanel); // Reemplaza esto con tus propios datos
			}
			else
			{
				bytesDelExcel = ObtenerDatosComprobantesSin(idPanel); // Reemplaza esto con tus propios datos
			}
			
															  // Crear un MemoryStream a partir de los bytes del Excel
			using (MemoryStream memoryStream = new MemoryStream(bytesDelExcel))
			{
				// Crear un paquete Excel a partir del MemoryStream
				using (ExcelPackage package = new ExcelPackage(memoryStream))
				{
					// Obtener el contenido del paquete en un array de bytes
					byte[] excelBytes = ComprobantesCon(package, idPanel,modalidad, detraccion, serie,pagina).GetAsByteArray();

					//return Json(UpdateDatosExcelMesual(excelBytes));


					// Devolver el archivo Excel al cliente como descarga
					return File(excelBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", nombre+"_PARTE_"+ pagina+".xlsx");
				}
			}
		}

		public byte[] ObtenerDatosComprobantesCon(int idPanel)
		{
			var familia = db.Panel_2011.Where(x => x.ID == idPanel).ToList().Last();
			var integrate = db.AP__DocIntegratedStorage.Where(x => x.IntegrationObjectId == familia.C_ElementID).ToList();
			byte[] data = null;
			foreach (var item in integrate)
			{
				
				var integrateData = db.AP__DocIntegratedDataStorage.Where(x => x.AscDocId == item.ID).ToList();
                foreach (var item1 in integrateData)
                {
                    if(item.Name == "ARCHIVO_CON_DETRACCION.xlsx")
					{
						data = item1.Content;
						break;
					}
                }
                
			}
			return data;
			
		}

		public byte[] ObtenerDatosComprobantesSin(int idPanel)
		{
			var familia = db.Panel_2011.Where(x => x.ID == idPanel).ToList().Last();
			var integrate = db.AP__DocIntegratedStorage.Where(x => x.IntegrationObjectId == familia.C_ElementID).ToList();
			byte[] data = null;
			foreach (var item in integrate)
			{

				var integrateData = db.AP__DocIntegratedDataStorage.Where(x => x.AscDocId == item.ID).ToList();
				foreach (var item1 in integrateData)
				{
					if (item.Name == "ARCHIVO_SIN_DETRACCION.xlsx")
					{
						data = item1.Content;
						break;
					}
				}

			}
			return data;

		}

		public ExcelPackage ComprobantesCon(ExcelPackage package, int idPanel, int modalidad, int detraccion, string serie,int pagina)
		{
			List<string> nombresDeHojas = new List<string>();
			foreach (var worksheet in package.Workbook.Worksheets)
			{
				nombresDeHojas.Add(worksheet.Name);
			}
			ExcelWorksheet hoja1 = package.Workbook.Worksheets[nombresDeHojas[0]];
			int filaActual = 14;

			var lista1 = db.Panel_2011_8800.Where(x =>x.C3_PR3_FC_SERIE == serie)
				.Where(x => x.C3_PR3_FC_MODALIDAD == modalidad)
				.Where(x => x.C_ElementID == idPanel)
				.Where(x => x.C3_PR3_FC_DETRACCION == detraccion)
				.OrderBy(x => x.C3_PR3_FC_FECHA_EMISION)
				.Skip((pagina - 1) * 30)
				.Take(30)//30 elementos
				.ToList();
			int correlativo = 1;
            foreach (var item in lista1)
            {
				hoja1.Cells["A" + filaActual].Value = "01";
				hoja1.Cells["B" + filaActual].Value = correlativo;
				hoja1.Cells["E" + filaActual].Value = "PEN";
				hoja1.Cells["F" + filaActual].Value = "01";
				hoja1.Cells["G" + filaActual].Value = "0000";
				hoja1.Cells["H" + filaActual].Value = "CREDITO";
				hoja1.Cells["I" + filaActual].Formula = "=IF(AN" + filaActual + " > 700, (AN" + filaActual + "-(AN" + filaActual + "*0.12)), AN" + filaActual + ")";
				hoja1.Cells["I" + filaActual].Style.Numberformat.Format = "#,##0.00";
				hoja1.Cells["J" + filaActual].Value = 1;
				hoja1.Cells["K" + filaActual].Formula = "=+ROUND(I" + filaActual + ",2)";
				hoja1.Cells["M" + filaActual].Value = 6;
				hoja1.Cells["Q" + filaActual].Value = "022";
				hoja1.Cells["R" + filaActual].Value = 80161501;
				hoja1.Cells["S" + filaActual].Value = 1;
				hoja1.Cells["T" + filaActual].Value = "NIU";
				hoja1.Cells["U" + filaActual].Value = "SERVICIO ADMINISTRACION DE CUENTA";
				hoja1.Cells["W" + filaActual].Formula = "=V"+ filaActual + "*1.18";
				hoja1.Cells["W" + filaActual].Style.Numberformat.Format = "#,##0.00";
				hoja1.Cells["X" + filaActual].Value = 10;
				hoja1.Cells["Y" + filaActual].Formula = "=((S"+ filaActual + "*V"+ filaActual + ")-Z"+ filaActual + ")*18%";
				hoja1.Cells["Y" + filaActual].Style.Numberformat.Format = "#,##0.00";
				hoja1.Cells["Z" + filaActual].Value = 0;
				hoja1.Cells["Z" + filaActual].Style.Numberformat.Format = "#,##0.00";
				hoja1.Cells["AA" + filaActual].Formula = "=(S"+ filaActual + "*V"+ filaActual + ")-Z"+ filaActual;
				hoja1.Cells["AA" + filaActual].Style.Numberformat.Format = "#,##0.00";
				hoja1.Cells["AB" + filaActual].Value = 0;
				hoja1.Cells["AB" + filaActual].Style.Numberformat.Format = "#,##0.00";
				hoja1.Cells["AC" + filaActual].Value = 0;
				hoja1.Cells["AC" + filaActual].Style.Numberformat.Format = "#,##0.00";
				hoja1.Cells["AD" + filaActual].Formula = "=+AA"+ filaActual;
				hoja1.Cells["AD" + filaActual].Style.Numberformat.Format = "#,##0.00";
				hoja1.Cells["AE" + filaActual].Value = 0;
				hoja1.Cells["AE" + filaActual].Style.Numberformat.Format = "#,##0.00";
				hoja1.Cells["AF" + filaActual].Value = 0;
				hoja1.Cells["AF" + filaActual].Style.Numberformat.Format = "#,##0.00";
				hoja1.Cells["AG" + filaActual].Value = 0;
				hoja1.Cells["AG" + filaActual].Style.Numberformat.Format = "#,##0.00";
				hoja1.Cells["AH" + filaActual].Value = 0.18;
				hoja1.Cells["AI" + filaActual].Formula = "=AD"+ filaActual + "*AH"+ filaActual;
				hoja1.Cells["AI" + filaActual].Style.Numberformat.Format = "#,##0.00";
				hoja1.Cells["AJ" + filaActual].Value = 0;
				hoja1.Cells["AJ" + filaActual].Style.Numberformat.Format = "#,##0.00";
				hoja1.Cells["AK" + filaActual].Value = 0;
				hoja1.Cells["AK" + filaActual].Style.Numberformat.Format = "#,##0.00";
				hoja1.Cells["AL" + filaActual].Value = 0;
				hoja1.Cells["AL" + filaActual].Style.Numberformat.Format = "#,##0.00";
				hoja1.Cells["AM" + filaActual].Value = 0;
				hoja1.Cells["AM" + filaActual].Style.Numberformat.Format = "#,##0.00";
				hoja1.Cells["AN" + filaActual].Formula = "=+W"+ filaActual;
				hoja1.Cells["AN" + filaActual].Style.Numberformat.Format = "#,##0.00";
				ExcelRange border = null;
				if (detraccion == 1)
				{
					hoja1.Cells["BA" + filaActual].Value = "00005170494";
					hoja1.Cells["BB" + filaActual].Value = "037";
					hoja1.Cells["BC" + filaActual].Value = 12;
					hoja1.Cells["BD" + filaActual].Value = 1;
					hoja1.Cells["BD" + filaActual].Style.Numberformat.Format = "0.00";
					hoja1.Cells["BE" + filaActual].Formula = "=AN" + filaActual + "*BD" + filaActual + "*BC" + filaActual + "/100";
					hoja1.Cells["BE" + filaActual].Style.Numberformat.Format = "#,##0.00";
					border = hoja1.Cells["A" + filaActual + ":BE" + filaActual];
				}
				else
				{
					border = hoja1.Cells["A" + filaActual + ":AN" + filaActual];
				}

				

				
				// Agregar bordes al rango de celdas
				border.Style.Border.Top.Style = ExcelBorderStyle.Thin;
				border.Style.Border.Left.Style = ExcelBorderStyle.Thin;
				border.Style.Border.Right.Style = ExcelBorderStyle.Thin;
				border.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

				string hexrosado = "#febca7";
				ExcelRange rosado = hoja1.Cells["A" + filaActual + ":AN" + filaActual];
				System.Drawing.Color color = System.Drawing.ColorTranslator.FromHtml(hexrosado);
				rosado.Style.Fill.PatternType = ExcelFillStyle.Solid;
				rosado.Style.Fill.BackgroundColor.SetColor(color);


				hoja1.Cells["A" + filaActual + ":BE" + filaActual].Style.Font.Name = "Arial";
				hoja1.Cells["A" + filaActual + ":AZ" + filaActual].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


				hoja1.Cells["C" + filaActual].Value = item.C3_PR3_FC_FECHA_EMISION.Value.ToString("yyyy-MM-dd");
				hoja1.Cells["D" + filaActual].Value = item.C3_PR3_FC_FECHA_VENCIMIENTO.Value.ToString("yyyy-MM-dd");
				hoja1.Cells["L" + filaActual].Value = item.C3_PR3_FC_FECHA_VENCIMIENTO.Value.ToString("yyyy-MM-dd");
				hoja1.Cells["N" + filaActual].Value = item.C3_PR3_FC_RUC;
				hoja1.Cells["O" + filaActual].Value = item.C3_PR3_FC_RAZON_SOCIAL;
				hoja1.Cells["P" + filaActual].Value = item.C3_PR3__FC_DIRECCION;
				hoja1.Cells["V" + filaActual].Value = item.C3_PR3_FC_BASE_FAC;
				hoja1.Cells["V" + filaActual].Style.Numberformat.Format = "#,##0.00";
				hoja1.Cells["AU" + filaActual].Value = item.C3_PR3_FC_CORREO;
				hoja1.Cells["AY" + filaActual].Value = item.C3_PR3_FC_CUENTA;
				filaActual++;
				correlativo++;
			}
            return package;
		}



		public async Task<string> insertDataWSAuraAsync(string referencia, string urlData)
		{
			var jsonData = @"
        {
            ""Reference"": """ + referencia + @""",
            ""PanelFields"": [
                {
                    ""InternalName"": ""6_PR3_FC_ARCHIVOS_FACTURACION"",
                    ""Type"": ""FieldsGroup"",
                    ""FieldsGroupValue"": [
                        " + urlData + @"
                    ]
                }
            ]
        }";

			// URL de destino
			string urlbase = ConfigurationManager.AppSettings["AuraUrl"];
			var url = urlbase+"/WebService/Custom/RestMessage.svc/PR3_GEN_EXCEL";

			// Crear cliente HttpClient
			using (var httpClient = new HttpClient())
			{
				try
				{
					// Configurar el encabezado personalizado
					httpClient.DefaultRequestHeaders.Add("UserData", ConfigurationManager.AppSettings["AuraUser"]);
					httpClient.DefaultRequestHeaders.Add("Apikey", ConfigurationManager.AppSettings["AuraKey"]);
					// Convertir JSON a StringContent
					var content = new StringContent(jsonData, Encoding.UTF8, "application/json");

					// Realizar la solicitud POST y obtener la respuesta
					var response = await httpClient.PostAsync(url, content);

					// Leer la respuesta
					var responseContent = await response.Content.ReadAsStringAsync();
					var dd = "121";
					// Imprimir la respuesta del servidor
					Console.WriteLine(responseContent);
					return responseContent;
				}
				catch (Exception ex)
				{
					// Manejar errores
					Console.WriteLine($"Error: {ex.Message}");
				}



			}
			return "12";
		}

	}
}