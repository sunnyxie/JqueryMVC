using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace JqueryMVC2.Controllers
{
    public class InvestingTaxController : Controller
    {
        // GET: InvestingTax
        public ActionResult Index()
        {
            return View();
        }

        // GET: InvestingTax/Details/5
        public ActionResult Details(int id)
        {
            return View();
        }

        // GET: InvestingTax/Create
        public ActionResult Create()
        {
            JqueryMVC2.Models.InvestingTax mo = new JqueryMVC2.Models.InvestingTax();
            return View("CreateView", mo);
        }

        // POST: InvestingTax/Create
        [HttpPost]
        public ActionResult Create(FormCollection collection)
        {
            try
            {
                // TODO: Add insert logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        // GET: InvestingTax/Edit/5
        public ActionResult Edit(int id)
        {
            return View();
        }

        // POST: InvestingTax/Edit/5
        [HttpPost]
        public ActionResult Edit(int id, FormCollection collection)
        {
            try
            {
                // TODO: Add update logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        // GET: InvestingTax/Delete/5
        public ActionResult Delete(int id)
        {
            return View();
        }

        // POST: InvestingTax/Delete/5
        [HttpPost]
        public ActionResult Delete(int id, FormCollection collection)
        {
            try
            {
                // TODO: Add delete logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }
    }
}
