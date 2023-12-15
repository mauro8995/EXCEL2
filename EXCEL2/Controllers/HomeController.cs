using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using EXCEL2.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Data.Entity.Core.Common.CommandTrees.ExpressionBuilder;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Web;
using System.Web.Mvc;
using System.Web.UI;
using System.Web.UI.WebControls;
using HorizontalAlignmentValues = DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues;

namespace EXCEL2.Controllers
{
    public class HomeController : Controller
    {
        private AuraPortal_BPMSEntities db = new AuraPortal_BPMSEntities();
        public ActionResult Index()
        {
            return View();
        }

        [HttpGet]
        public ActionResult ExcelFacturaDownload(int idPanel)
        {
            //factura

       
                // Código que utiliza SpreadsheetLight.dll
                if (db.Panel_5.Any(x => x.C3_PR3_TAG_ID_PANEL == idPanel))
                {

                }
                else
                {
                    return Json("No existe el correlativo de id panel", JsonRequestBehavior.AllowGet);
                }

                var stream = new System.IO.MemoryStream();
                byte[] excelBytes =  ExcelFacturaArchivo(idPanel).GetAsByteArray();


                

                Panel_5 panel = db.Panel_5.Where(x => x.C3_PR3_TAG_ID_PANEL == idPanel).FirstOrDefault();

                var NombreArchivo = "TAG_COVISOL_F_" + panel.C3_PR3_TAG_MES_TEXTO + "_" + panel.C3_PR3_TAG_AÑO + ".xlsx";

                return File(excelBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", NombreArchivo);
       
        }

        [HttpGet]
        public ActionResult ExcelBoletaGet(int idPanel)
        {
            //factura

            int id = idPanel;

            if (db.Panel_5.Any(x => x.C3_PR3_TAG_ID_PANEL == id))
            {

            }
            else
            {
                return Json("No existe el correlativo de id panel", JsonRequestBehavior.AllowGet);
            }

            MemoryStream ms = new MemoryStream();

            var mdoc = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory);

            Panel_5 panel = db.Panel_5.Where(x => x.C3_PR3_TAG_ID_PANEL == id).FirstOrDefault();

            var NombreArchivo = "/TAG_COBISOL_F_" + panel.C3_PR3_TAG_MES_TEXTO + "_" + panel.C3_PR3_TAG_AÑO + ".xlsx";

            var ruta = mdoc + NombreArchivo;

            //ExcelFacturaArchivo(id).SaveAs(ruta);

            panel.C3_PR3_LINK_BOLETA = $"{Request.Url.Scheme}://{Request.Url.Authority}" + "/home/ExcelBoletaDownload?idPanel=" + id;

            db.Entry(panel).State = EntityState.Modified;
            db.SaveChanges();

            var v = new { Amount = "Excel generado", ruta = panel.C3_PR3_LINK_FACTURA };


            return Json(v, JsonRequestBehavior.AllowGet);
        }


        [HttpGet]
        public ActionResult ExcelFacturaGet(int idPanel)
        {
            //factura
        
            int id = idPanel;

            if (db.Panel_5.Any(x => x.C3_PR3_TAG_ID_PANEL == id))
            {

            }
            else
            {
                return Json("No existe el correlativo de id panel", JsonRequestBehavior.AllowGet);
            }

            MemoryStream ms = new MemoryStream();

            var mdoc = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory);

            Panel_5 panel = db.Panel_5.Where(x => x.C3_PR3_TAG_ID_PANEL == id).FirstOrDefault();

            var NombreArchivo = "/TAG_COBISOL_F_" + panel.C3_PR3_TAG_MES_TEXTO + "_" + panel.C3_PR3_TAG_AÑO + ".xlsx";

            var ruta = mdoc + NombreArchivo;


            panel.C3_PR3_LINK_FACTURA = $"{Request.Url.Scheme}://{Request.Url.Authority}" + "/home/ExcelFacturaDownload?idPanel=" + id;

            db.Entry(panel).State = EntityState.Modified;
            db.SaveChanges();

            var v = new { Amount = "Excel generado", ruta = panel.C3_PR3_LINK_FACTURA };


            return Json(v, JsonRequestBehavior.AllowGet);
        }


        [HttpGet]
        public FileResult ExcelBoletaDownload(int idPanel)
        {
            if (db.Panel_5.Any(x => x.C3_PR3_TAG_ID_PANEL == idPanel))
            {

            }
            else
            {
              // return Json("No existe el correlativo de id panel", JsonRequestBehavior.AllowGet);
            }

           
            var stream = new System.IO.MemoryStream();
            byte[] excelBytes = ExcelBoletaArchivo(idPanel).GetAsByteArray();

            Panel_5 panel = db.Panel_5.Where(x => x.C3_PR3_TAG_ID_PANEL == idPanel).FirstOrDefault();

            var NombreArchivo = "TAG_COVISOL_B_" + panel.C3_PR3_TAG_MES_TEXTO + "_" + panel.C3_PR3_TAG_AÑO + ".xlsx";


            return File(excelBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", NombreArchivo);
        }

        [HttpPost]
        public ActionResult ExcelBoleta(int idPanel)
        {
            if (db.Panel_5.Any(x => x.C3_PR3_TAG_ID_PANEL == idPanel))
            {

            }
            else
            {
                return Json("No existe el correlativo de id panel", JsonRequestBehavior.AllowGet);
            }

            Panel_5 panel = db.Panel_5.Where(x => x.C3_PR3_TAG_ID_PANEL == idPanel).FirstOrDefault();

            MemoryStream ms = new MemoryStream();
            var mdoc = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory);

            var ruta = mdoc + "/TAG_COBISOL_B_" + panel.C3_PR3_TAG_MES_TEXTO + "_" + panel.C3_PR3_TAG_AÑO + ".xlsx";

            ExcelBoletaArchivo(idPanel).SaveAs(ruta);

            panel.C3_PR3_LINK_BOLETA = ruta;

            db.Entry(panel).State = EntityState.Modified;
            db.SaveChanges();

            var v = new { Amount = "Excel generado", ruta = ruta };


            return Json(v, JsonRequestBehavior.AllowGet);
        }


        public ExcelPackage ExcelFacturaArchivo(int idPanel)
        {
            var mdoc = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            ExcelPackage package = new ExcelPackage(mdoc + "/FACTURAS.xlsx");

            List<string> nombresDeHojas = new List<string>();
            foreach (var worksheet in package.Workbook.Worksheets)
            {
                nombresDeHojas.Add(worksheet.Name);
            }
            ExcelWorksheet hoja1 = package.Workbook.Worksheets[nombresDeHojas[0]];

            var infoCorrelativo = db.Panel_5.Where(x => x.C3_PR3_TAG_ID_PANEL == idPanel).First();
            var dataFactura = db.Panel_5_2664.Where(x => x.C3_TIPO_COMPROBANTE == "2").Where(x => x.C_ElementID == idPanel).ToList();


            hoja1.Cells["C10"].Value = infoCorrelativo.C3_PR3_TAG_CORRELATIVO_FACTURA.ToString();
            int contador = 0;
            int filaInicialFactura = 14;
            foreach (var item in dataFactura)
            {
                contador++;
                ExcelRange border = hoja1.Cells["A" + filaInicialFactura+ ":BE" + filaInicialFactura];

                // Agregar bordes al rango de celdas
                border.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                border.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                border.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                border.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                ExcelRange TextCentrar = hoja1.Cells["A" + filaInicialFactura + ":H" + filaInicialFactura];

                TextCentrar.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                TextCentrar.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                ExcelRange TextCentrar2 = hoja1.Cells["J" + filaInicialFactura + ":N" + filaInicialFactura];
                TextCentrar2.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                TextCentrar2.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                ExcelRange TextCentrar3 = hoja1.Cells["Q" + filaInicialFactura + ":T" + filaInicialFactura];
                TextCentrar3.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                TextCentrar3.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                ExcelRange TextCentrar4 = hoja1.Cells["X" + filaInicialFactura];
                TextCentrar4.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                TextCentrar4.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                ExcelRange TextCentrar5 = hoja1.Cells["AH" + filaInicialFactura];
                TextCentrar5.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                TextCentrar5.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                ExcelRange TextCentrar6 = hoja1.Cells["AL" + filaInicialFactura];
                TextCentrar6.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                TextCentrar6.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                ExcelRange TextCentrar7 = hoja1.Cells["AQ" + filaInicialFactura];
                TextCentrar7.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                TextCentrar7.Style.VerticalAlignment = ExcelVerticalAlignment.Center;


                hoja1.Cells["A" + filaInicialFactura].Value = "01";
                hoja1.Cells["B" + filaInicialFactura].Value = contador;
                hoja1.Cells["E" + filaInicialFactura].Value = "PEN";
                hoja1.Cells["F" + filaInicialFactura].Value = "01";
                hoja1.Cells["G" + filaInicialFactura].Value = "0000";
                hoja1.Cells["H" + filaInicialFactura].Value = "CONTADO";
                hoja1.Cells["M" + filaInicialFactura].Value = "6";
                hoja1.Cells["Q" + filaInicialFactura].Value = "10";
                hoja1.Cells["R" + filaInicialFactura].Value = "32151910";
                hoja1.Cells["T" + filaInicialFactura].Value = "NIU";
                hoja1.Cells["U" + filaInicialFactura].Value = "TAG - ENTREGA A TITULO GRATUITO";
                hoja1.Cells["V" + filaInicialFactura].Value = RedondearYFormatear(25.423);
                hoja1.Cells["V" + filaInicialFactura].Style.Numberformat.Format = "0.00";
                hoja1.Cells["W" + filaInicialFactura].Formula = "=V"+filaInicialFactura.ToString();
                hoja1.Cells["X" + filaInicialFactura].Value = "13";
                hoja1.Cells["Y" + filaInicialFactura].Formula = "=((S" + filaInicialFactura.ToString() + "*V" + filaInicialFactura.ToString() + ")-Z" + filaInicialFactura.ToString() + ")*18%";
                hoja1.Cells["Z" + filaInicialFactura].Value = 0;
                hoja1.Cells["Z" + filaInicialFactura].Style.Numberformat.Format = "0.00";
                hoja1.Cells["AA" + filaInicialFactura].Formula = "=(S" + filaInicialFactura.ToString() + "*V" + filaInicialFactura.ToString() + ")-Z" + filaInicialFactura.ToString();
                hoja1.Cells["AB" + filaInicialFactura].Value = 0;
                hoja1.Cells["AB" + filaInicialFactura].Style.Numberformat.Format = "0.00";
                hoja1.Cells["AC" + filaInicialFactura].Value = 0;
                hoja1.Cells["AC" + filaInicialFactura].Style.Numberformat.Format = "0.00";
                hoja1.Cells["AD" + filaInicialFactura].Value = 0;
                hoja1.Cells["AD" + filaInicialFactura].Style.Numberformat.Format = "0.00";
                hoja1.Cells["AE" + filaInicialFactura].Value = 0;
                hoja1.Cells["AE" + filaInicialFactura].Style.Numberformat.Format = "0.00";
                hoja1.Cells["AF" + filaInicialFactura].Value = 0;
                hoja1.Cells["AF" + filaInicialFactura].Style.Numberformat.Format = "0.00";
                hoja1.Cells["AG" + filaInicialFactura].Formula = "AA" + filaInicialFactura.ToString();
                hoja1.Cells["AH" + filaInicialFactura].Value = 0.18;
                hoja1.Cells["AH" + filaInicialFactura].Style.Numberformat.Format = "0.00";
                hoja1.Cells["AI" + filaInicialFactura].Formula = "=AD" + filaInicialFactura.ToString() + "*AH" + filaInicialFactura.ToString();
                hoja1.Cells["AJ" + filaInicialFactura].Value = 0;
                hoja1.Cells["AJ" + filaInicialFactura].Style.Numberformat.Format = "0.00";
                hoja1.Cells["AK" + filaInicialFactura].Value = 0;
                hoja1.Cells["AK" + filaInicialFactura].Style.Numberformat.Format = "0.00";
                hoja1.Cells["AL" + filaInicialFactura].Value = 0;
                hoja1.Cells["AL" + filaInicialFactura].Style.Numberformat.Format = "0.00";
                hoja1.Cells["AM" + filaInicialFactura].Value = 0;
                hoja1.Cells["AM" + filaInicialFactura].Style.Numberformat.Format = "0.00";
                hoja1.Cells["AN" + filaInicialFactura].Formula = "=AD" + filaInicialFactura.ToString() + "+AE" + filaInicialFactura.ToString() + "+AF" + filaInicialFactura.ToString() + "+AI" + filaInicialFactura.ToString() + "+AK" + filaInicialFactura.ToString() + "+AM" + filaInicialFactura.ToString();

                string dateString = item.C3_FECHA_EMISION; 
                string format = "dd/MM/yyyy";
                DateTime dateTime = DateTime.ParseExact(dateString, format, System.Globalization.CultureInfo.InvariantCulture);
                hoja1.Cells["C" + filaInicialFactura].Value = dateTime.ToString("yyyy-MM-dd");//fecha DE EMICION
                hoja1.Cells["N" + filaInicialFactura].Value = item.C3_NUM_DOCUMENTO;//Numero de documento
                hoja1.Cells["O" + filaInicialFactura].Value = item.C3_RAZON_SOCIAL;//Razon social
                hoja1.Cells["P" + filaInicialFactura].Value = item.C3_DIRECCION;//Direccion fisica
                hoja1.Cells["S" + filaInicialFactura].Value = item.C3_CANTIDAD;//Cantidad
                //hoja1.Cells["AS" + filaInicialFactura].Value = item.;//Cantidad
                hoja1.Cells["AU" + filaInicialFactura].Value = item.C3_EMAIL;//Correo
                hoja1.Cells["AY" + filaInicialFactura].Value = item.C3_CUENTA;//Texto adicional

                filaInicialFactura++;
            }

            return package;
        }

        public ExcelPackage ExcelBoletaArchivo(int idPanel)
        {
            var mdoc = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var infoCorrelativo = db.Panel_5.Where(x => x.C3_PR3_TAG_ID_PANEL == idPanel).First();
            //Boleta
            var dataBoleta = db.Panel_5_2664.Where(x => x.C3_TIPO_COMPROBANTE == "1").Where(x => x.C_ElementID == idPanel).ToList();
            var mdocBoleta = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory);

            ExcelPackage package = new ExcelPackage(mdoc + "/BOLETAS.xlsx");

            List<string> nombresDeHojas = new List<string>();
            foreach (var worksheet in package.Workbook.Worksheets)
            {
                nombresDeHojas.Add(worksheet.Name);
            }
            ExcelWorksheet hoja1 = package.Workbook.Worksheets[nombresDeHojas[0]];

            hoja1.Cells["C9"].Value = infoCorrelativo.C3_PR3_TAG_CORRELATIVO.ToString();

            int filaInicialFactura = 13;
            int contador = 0;
            foreach (var item in dataBoleta)
            {
                contador++;
                ExcelRange border = hoja1.Cells["A" + filaInicialFactura + ":BE" + filaInicialFactura];

                // Agregar bordes al rango de celdas
                border.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                border.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                border.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                border.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                ExcelRange TextCentrar = hoja1.Cells["A" + filaInicialFactura + ":H" + filaInicialFactura];

                TextCentrar.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                TextCentrar.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                ExcelRange TextCentrar2 = hoja1.Cells["J" + filaInicialFactura + ":N" + filaInicialFactura];
                TextCentrar2.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                TextCentrar2.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                ExcelRange TextCentrar3 = hoja1.Cells["Q" + filaInicialFactura + ":T" + filaInicialFactura];
                TextCentrar3.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                TextCentrar3.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                ExcelRange TextCentrar4 = hoja1.Cells["X" + filaInicialFactura];
                TextCentrar4.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                TextCentrar4.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                ExcelRange TextCentrar5 = hoja1.Cells["AH" + filaInicialFactura];
                TextCentrar5.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                TextCentrar5.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                ExcelRange TextCentrar6 = hoja1.Cells["AL" + filaInicialFactura];
                TextCentrar6.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                TextCentrar6.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                ExcelRange TextCentrar7 = hoja1.Cells["AQ" + filaInicialFactura];
                TextCentrar7.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                TextCentrar7.Style.VerticalAlignment = ExcelVerticalAlignment.Center;


                hoja1.Cells["A" + filaInicialFactura].Value = "03";
                hoja1.Cells["B" + filaInicialFactura].Value = contador;
                hoja1.Cells["E" + filaInicialFactura].Value = "PEN";
                hoja1.Cells["F" + filaInicialFactura].Value = "01";
                hoja1.Cells["G" + filaInicialFactura].Value = "0000";
                hoja1.Cells["H" + filaInicialFactura].Value = "CONTADO";
                hoja1.Cells["M" + filaInicialFactura].Value = "1";
                hoja1.Cells["Q" + filaInicialFactura].Value = "10";
                hoja1.Cells["R" + filaInicialFactura].Value = "32151910";
                hoja1.Cells["T" + filaInicialFactura].Value = "NIU";
                hoja1.Cells["U" + filaInicialFactura].Value = "TAG - ENTREGA A TITULO GRATUITO";
                hoja1.Cells["V" + filaInicialFactura].Value = RedondearYFormatear(25.423);
                hoja1.Cells["V" + filaInicialFactura].Style.Numberformat.Format = "0.00";
                hoja1.Cells["W" + filaInicialFactura].Formula = "=V" + filaInicialFactura.ToString();
                hoja1.Cells["X" + filaInicialFactura].Value = "13";
                hoja1.Cells["Y" + filaInicialFactura].Formula = "=((S" + filaInicialFactura.ToString() + "*V" + filaInicialFactura.ToString() + ")-Z" + filaInicialFactura.ToString() + ")*18%";
                hoja1.Cells["Z" + filaInicialFactura].Value = 0;
                hoja1.Cells["Z" + filaInicialFactura].Style.Numberformat.Format = "0.00";
                hoja1.Cells["AA" + filaInicialFactura].Formula = "=(S" + filaInicialFactura.ToString() + "*V" + filaInicialFactura.ToString() + ")-Z" + filaInicialFactura.ToString();
                hoja1.Cells["AB" + filaInicialFactura].Value = 0;
                hoja1.Cells["AB" + filaInicialFactura].Style.Numberformat.Format = "0.00";
                hoja1.Cells["AC" + filaInicialFactura].Value = 0;
                hoja1.Cells["AC" + filaInicialFactura].Style.Numberformat.Format = "0.00";
                hoja1.Cells["AD" + filaInicialFactura].Value = 0;
                hoja1.Cells["AD" + filaInicialFactura].Style.Numberformat.Format = "0.00";
                hoja1.Cells["AE" + filaInicialFactura].Value = 0;
                hoja1.Cells["AE" + filaInicialFactura].Style.Numberformat.Format = "0.00";
                hoja1.Cells["AF" + filaInicialFactura].Value = 0;
                hoja1.Cells["AF" + filaInicialFactura].Style.Numberformat.Format = "0.00";
                hoja1.Cells["AG" + filaInicialFactura].Formula = "AA" + filaInicialFactura.ToString();
                hoja1.Cells["AH" + filaInicialFactura].Value = 0.18;
                hoja1.Cells["AH" + filaInicialFactura].Style.Numberformat.Format = "0.00";
                hoja1.Cells["AI" + filaInicialFactura].Formula = "=AD" + filaInicialFactura.ToString() + "*AH" + filaInicialFactura.ToString();
                hoja1.Cells["AJ" + filaInicialFactura].Value = 0;
                hoja1.Cells["AJ" + filaInicialFactura].Style.Numberformat.Format = "0.00";
                hoja1.Cells["AK" + filaInicialFactura].Value = 0;
                hoja1.Cells["AK" + filaInicialFactura].Style.Numberformat.Format = "0.00";
                hoja1.Cells["AL" + filaInicialFactura].Value = 0;
                hoja1.Cells["AL" + filaInicialFactura].Style.Numberformat.Format = "0.00";
                hoja1.Cells["AM" + filaInicialFactura].Value = 0;
                hoja1.Cells["AM" + filaInicialFactura].Style.Numberformat.Format = "0.00";
                hoja1.Cells["AN" + filaInicialFactura].Formula = "=AD" + filaInicialFactura.ToString() + "+AE" + filaInicialFactura.ToString() + "+AF" + filaInicialFactura.ToString() + "+AI" + filaInicialFactura.ToString() + "+AK" + filaInicialFactura.ToString() + "+AM" + filaInicialFactura.ToString();


                string dateString = item.C3_FECHA_EMISION;
                string format = "dd/MM/yyyy";
                DateTime dateTime = DateTime.ParseExact(dateString, format, System.Globalization.CultureInfo.InvariantCulture);
                hoja1.Cells["C" + filaInicialFactura].Value = dateTime.ToString("yyyy-MM-dd");//fecha DE EMICION
                hoja1.Cells["N" + filaInicialFactura].Value = item.C3_NUM_DOCUMENTO;//Numero de documento
                hoja1.Cells["O" + filaInicialFactura].Value = item.C3_RAZON_SOCIAL;//Razon social
                hoja1.Cells["P" + filaInicialFactura].Value = item.C3_DIRECCION;//Direccion fisica
                hoja1.Cells["S" + filaInicialFactura].Value = item.C3_CANTIDAD;//Cantidad
                //hoja1.Cells["AS" + filaInicialFactura].Value = item.;//Cantidad
                hoja1.Cells["AU" + filaInicialFactura].Value = item.C3_EMAIL;//Correo
                hoja1.Cells["AY" + filaInicialFactura].Value = item.C3_CUENTA;//Texto adicional

                filaInicialFactura++;
            }
            return package;
        }

        static double RedondearYFormatear(double numero)
        {
            // Redondea el número a dos decimales
            double numeroRedondeado = Math.Round(numero, 2);

            // Formatea el número con dos decimales y devuelve la cadena resultante
            return numeroRedondeado;
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}