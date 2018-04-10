﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using JqueryMVC2.Models;
using System.Text.RegularExpressions;

namespace JqueryMVC2.Controllers
{
    public class MoviesController : Controller
    {
        private ApplicationDbContext db = new ApplicationDbContext();

        // GET: Movies
        public ActionResult Index()
        {
            if (System.Runtime.Caching.MemoryCache.Default["TestID"] == null)
            {
                System.Runtime.Caching.MemoryCache.Default["TestID"] = "12345";
            }

            return View(db.Movies.ToList());
        }

        // GET: Movies/Details/5
        public ActionResult Details(int? id)
        {
            db.Movies.Add(new Movie { ID = 10, Name = "test 1", Score = 10 * 2 });
            db.Movies.Add(new Movie { ID = 20, Name = "test 2", Score = 20 * 2, DateOfBirth = System.DateTime.Today.Date });
            db.SaveChanges();

            NamedCaptureReuse();

            var Html1 = @"A<>>BN..>n>C";
             Html1 = System.Text.RegularExpressions.Regex.Replace(Html1, @"<(.|\n)*?>", string.Empty);
            Console.WriteLine("Resule " + Html1);

            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            else if (id == 13)
            {
                throw new OutOfMemoryException("For testing 2.");
            }

            Movie movie = db.Movies.Find(id);
            if (movie == null)
            {
                return HttpNotFound();
            }
            return View(movie);
        }

        public ActionResult Details2 (int? id)
        {
            db.Movies.Add(new Movie { ID = 20, Name = "test 2", Score = 20 * 2, DateOfBirth = System.DateTime.Today.Date });
            db.SaveChanges();

            if (id == null)
            {
                return View("Details2");
                // return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Movie movie = db.Movies.Find(id);
            if (movie == null)
            {
                return HttpNotFound();
            }
            return View(movie);
        }

        // GET: Movies/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: Movies/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "ID,Name,Score,PicturePath,DateOfBirth")] Movie movie)
        {
            if (ModelState.IsValid)
            {
                db.Movies.Add(movie);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(movie);
        }

        // GET: Movies/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Movie movie = db.Movies.Find(id);
            if (movie == null)
            {
                return HttpNotFound();
            }
            return View(movie);
        }

        // POST: Movies/Edit/5
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "ID,Name,Score,PicturePath,DateOfBirth")] Movie movie)
        {
            if (ModelState.IsValid)
            {
                db.Entry(movie).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(movie);
        }

        // GET: Movies/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Movie movie = db.Movies.Find(id);
            if (movie == null)
            {
                return HttpNotFound();
            }
            return View(movie);
        }

        // POST: Movies/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            Movie movie = db.Movies.Find(id);
            db.Movies.Remove(movie);
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

        protected void NamedCaptureReuse()
        {
            string outString = @"one:uno dos:two three:tres";
            string outPattern = @"(?<someword>\w+):(?<someword>\w+)";
            var outRegex = new Regex(outPattern);
            MatchCollection AllMatches = outRegex.Matches(outString);

            int matchNum = 1;

            foreach (Match SomeMatch in AllMatches)
            {
                matchNum++;
                Console.WriteLine(@"Groups[1].Value = " + SomeMatch.Groups[1].Value);
                Console.WriteLine(@"Groups[""someword""].Value = " + SomeMatch.Groups["someword"].Value);

                foreach (Capture someword in SomeMatch.Groups["someword"].Captures)
                {
                    Console.WriteLine("someword Capture: " + someword.Value);
                }
            }

            outString = @"Console.WriteLine()";
            outPattern = @"Write(?:Line)?";
            Match mat = Regex.Match(outString, outPattern);
            Console.WriteLine("1` " + mat.Value);


            outString = @"Console.WriteLineGogo()";
            outPattern = @"Write(?:Line)?\w+";
            mat = Regex.Match(outString, outPattern);
            Console.WriteLine("1` " + mat.Value);


            // \k< name >   Named backreference. Matches the value of a named expression.
            outString = @"deepdeep55";
            outPattern = @"(?<double>\w\w)\k<double>";
             mat = Regex.Match(outString, outPattern);
            Console.WriteLine("1` " + mat.Value);
            Console.WriteLine("Resule " + AllMatches.Count);
        }
    }
}
