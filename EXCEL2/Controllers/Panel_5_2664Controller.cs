using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using EXCEL2.Models;

namespace EXCEL2.Controllers
{
    public class Panel_5_2664Controller : Controller
    {
        private AuraPortal_BPMSEntities db = new AuraPortal_BPMSEntities();

        // GET: Panel_5_2664
        public ActionResult Index()
        {
            return View(db.Panel_5_2664.ToList());
        }

        // GET: Panel_5_2664/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Panel_5_2664 panel_5_2664 = db.Panel_5_2664.Find(id);
            if (panel_5_2664 == null)
            {
                return HttpNotFound();
            }
            return View(panel_5_2664);
        }

        // GET: Panel_5_2664/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: Panel_5_2664/Create
        // Para protegerse de ataques de publicación excesiva, habilite las propiedades específicas a las que quiere enlazarse. Para obtener 
        // más detalles, vea https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "ID,C_ElementID,Creado,IdCreador,TipoCreador,Modificado,IdModificador,TipoModificador,C_ExecutedScript,C_Register_creation_ProcessID,C_Origin_Register_ID,C_RegisterStatus,C_Register_Origin_ID,C_Register_Origin_Type,C_Numerator,C3_NUM_DOCUMENTO,C3_RAZON_SOCIAL,C3_EMAIL,C3_FECHA_EMISION,C3_CANTIDAD,C3_DIRECCION,C3_CUENTA,C3_TIPO_COMPROBANTE")] Panel_5_2664 panel_5_2664)
        {
            if (ModelState.IsValid)
            {
                db.Panel_5_2664.Add(panel_5_2664);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(panel_5_2664);
        }

        // GET: Panel_5_2664/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Panel_5_2664 panel_5_2664 = db.Panel_5_2664.Find(id);
            if (panel_5_2664 == null)
            {
                return HttpNotFound();
            }
            return View(panel_5_2664);
        }

        // POST: Panel_5_2664/Edit/5
        // Para protegerse de ataques de publicación excesiva, habilite las propiedades específicas a las que quiere enlazarse. Para obtener 
        // más detalles, vea https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "ID,C_ElementID,Creado,IdCreador,TipoCreador,Modificado,IdModificador,TipoModificador,C_ExecutedScript,C_Register_creation_ProcessID,C_Origin_Register_ID,C_RegisterStatus,C_Register_Origin_ID,C_Register_Origin_Type,C_Numerator,C3_NUM_DOCUMENTO,C3_RAZON_SOCIAL,C3_EMAIL,C3_FECHA_EMISION,C3_CANTIDAD,C3_DIRECCION,C3_CUENTA,C3_TIPO_COMPROBANTE")] Panel_5_2664 panel_5_2664)
        {
            if (ModelState.IsValid)
            {
                db.Entry(panel_5_2664).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(panel_5_2664);
        }

        // GET: Panel_5_2664/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Panel_5_2664 panel_5_2664 = db.Panel_5_2664.Find(id);
            if (panel_5_2664 == null)
            {
                return HttpNotFound();
            }
            return View(panel_5_2664);
        }

        // POST: Panel_5_2664/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            Panel_5_2664 panel_5_2664 = db.Panel_5_2664.Find(id);
            db.Panel_5_2664.Remove(panel_5_2664);
            db.SaveChanges();
            return RedirectToAction("Index");
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
