using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Runtime.Remoting.Contexts;
using System.Web;
using System.Web.Mvc;
using System.Web.Security;
using Microsoft.Ajax.Utilities;
using MvcFirmaCagri.Models.Entity;

using OfficeOpenXml;

namespace MvcFirmaCagri.Controllers
{
    [Authorize]
    public class DefaultController : Controller
    {
        // GET: Default
        public ActionResult Index()
        {
            return View();
        }

        private DbisTakipEntities db = new DbisTakipEntities();              
        public ActionResult AktifCagrilar() 
        {
            var firmaMail = (string)Session["FIRMA_MAIL"];
            var firmaID = db.tbl_Firmalar.Where(x => x.FIRMA_MAIL == firmaMail).Select(x=> x.ID_FIRMA).FirstOrDefault();         

            var cagrilar = db.tbl_Cagri.Where(x=> x.CAGRI_DURUM == true && x.FIRMA_ID == firmaID).ToList();
            return View(cagrilar);
        }
        /*Aktif Cagrilar Export Excel*/
        public ActionResult AktifCagrilarExportToExcel()
        {
            var firmaMail = (string)Session["FIRMA_MAIL"];
            var firmaID = db.tbl_Firmalar.Where(x => x.FIRMA_MAIL == firmaMail).Select(x => x.ID_FIRMA).FirstOrDefault();
            var cagrilar = db.tbl_Cagri.Where(x => x.CAGRI_DURUM == true && x.FIRMA_ID == firmaID).ToList();

            var data = cagrilar.ToList(); 

            // Excel dosyası oluştur
            using (ExcelPackage excel = new ExcelPackage())
            {
                var workSheet = excel.Workbook.Worksheets.Add("Aktif Çağrılar");
                workSheet.Cells[1, 1].Value = "ID";
                workSheet.Cells[1, 2].Value = "Konu";
                workSheet.Cells[1, 3].Value = "Açıklama";
                workSheet.Cells[1, 4].Value = "Tarih";

                int rowStart = 2; // Başlangıç satırı (başlıkların altında)

                // Verileri Excel sayfasına ekle
                foreach (var item in data)
                {
                    workSheet.Cells[rowStart, 1].Value = item.ID_CAGRI;
                    workSheet.Cells[rowStart, 2].Value = item.KONU;
                    workSheet.Cells[rowStart, 3].Value = item.ACIKLAMA;
                    workSheet.Cells[rowStart, 4].Value = item.TARIH.ToString();
                    rowStart++;
                }

                // Excel dosyasını istemciye gönder
                var excelStream = new MemoryStream();
                excel.SaveAs(excelStream);
                var content = excelStream.ToArray();
                var fileName = $"AktifCagrilar_{DateTime.Now:yyyyMMdd}.xlsx"; // Dosya adı
                return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
        }

        public ActionResult PasifCagrilar() 
        {
            var firmaMail = (string)Session["FIRMA_MAIL"];
            var firmaID = db.tbl_Firmalar.Where(x => x.FIRMA_MAIL == firmaMail).Select(x => x.ID_FIRMA).FirstOrDefault();

            var cagrilar = db.tbl_Cagri.Where(x => x.CAGRI_DURUM == false && x.FIRMA_ID == firmaID).ToList();
            return View(cagrilar);
        }

        [HttpGet]
        public ActionResult YeniCagri() 
        {
            return View();
        }
        [HttpPost]
        public ActionResult YeniCagri(tbl_Cagri p) 
        {
            var firmaMail = (string)Session["FIRMA_MAIL"];
            var firmaID = db.tbl_Firmalar.Where(x => x.FIRMA_MAIL == firmaMail).Select(x => x.ID_FIRMA).FirstOrDefault();          

            p.CAGRI_DURUM = true;
            p.TARIH = DateTime.Parse(DateTime.Now.ToShortDateString());
            p.FIRMA_ID = firmaID;
            db.tbl_Cagri.Add(p);
            db.SaveChanges();
            return RedirectToAction("AktifCagrilar");
        }
        public ActionResult CagriDetay(int id) 
        {
            var cagri = db.tbl_CagriDetay.Where(x=> x.CAGRI_ID == id).ToList();
            return View(cagri);
        }
        
        public ActionResult CagriGetir(int id) 
        {
            var cagri = db.tbl_Cagri.Find(id);
            return View("CagriGetir",cagri);
        }

        public ActionResult CagriDuzenle(tbl_Cagri p) 
        {
            var cagri = db.tbl_Cagri.Find(p.ID_CAGRI);
            cagri.KONU = p.KONU;
            cagri.ACIKLAMA = p.ACIKLAMA;
            db.SaveChanges();
            return RedirectToAction("AktifCagrilar");
        }

        [HttpGet]
        public ActionResult ProfilDuzenle()
        {
            var firmaMail = (string)Session["FIRMA_MAIL"];
            var firmaID = db.tbl_Firmalar.Where(x => x.FIRMA_MAIL == firmaMail).Select(x => x.ID_FIRMA).FirstOrDefault();

            var profil = db.tbl_Firmalar.Where(x=> x.ID_FIRMA == firmaID).FirstOrDefault();
            return View(profil);
        }

        public ActionResult AnaSayfa() 
        {
            var firmaMail = (string)Session["FIRMA_MAIL"];
            var firmaID = db.tbl_Firmalar.Where(x => x.FIRMA_MAIL == firmaMail).Select(x => x.ID_FIRMA).FirstOrDefault();
            
            var toplamCagri = db.tbl_Cagri.Where(x=> x.FIRMA_ID == firmaID).Count();
            ViewBag.toplamCagri = toplamCagri;

            var toplamAktifCagri = db.tbl_Cagri.Where(x => x.FIRMA_ID == firmaID && x.CAGRI_DURUM == true).Count();
            ViewBag.toplamAktifCagri = toplamAktifCagri;

            var toplamPasifCagri = db.tbl_Cagri.Where(x => x.FIRMA_ID == firmaID && x.CAGRI_DURUM == false).Count();
            ViewBag.toplamPasifCagri = toplamPasifCagri;

            var firmaAdi = db.tbl_Firmalar.Where(x=> x.ID_FIRMA == firmaID).Select(x=> x.FIRMA_ADI).FirstOrDefault();
            ViewBag.firmaAdi = firmaAdi;

            var sektor =  db.tbl_Firmalar.Where(x => x.ID_FIRMA == firmaID).Select(x => x.FIRMA_SEKTOR).FirstOrDefault();
            ViewBag.sektor = sektor;

            var firmaGorsel = db.tbl_Firmalar.Where(x => x.ID_FIRMA == firmaID).Select(x => x.GORSEL).FirstOrDefault();
            ViewBag.firmaGorsel = firmaGorsel;

            return View();
        }

        public PartialViewResult Partial1() 
        {
            //true okunmamis mesajlar //false okunmus mesajlar
            var firmaMail = (string)Session["FIRMA_MAIL"];
            var firmaID = db.tbl_Firmalar.Where(x => x.FIRMA_MAIL == firmaMail).Select(x => x.ID_FIRMA).FirstOrDefault();
           
            var mesajlar = db.tbl_Mesajlar.Where(x => x.ALICI == firmaID && x.DURUM == true).ToList();
            var okunmamisMesajSayisi = db.tbl_Mesajlar.Where(x => x.ALICI == firmaID && x.DURUM == true).Count();
            ViewBag.okunmamisMesajSayisi = okunmamisMesajSayisi;
            
            return PartialView(mesajlar);
        }

        public PartialViewResult Partial2() 
        {
            var firmaMail = (string)Session["FIRMA_MAIL"];
            var firmaID = db.tbl_Firmalar.Where(x => x.FIRMA_MAIL == firmaMail).Select(x => x.ID_FIRMA).FirstOrDefault();
            var cagrilar = db.tbl_Cagri.Where(x => x.FIRMA_ID == firmaID && x.CAGRI_DURUM == true).ToList();
            var toplamAktifCagri = db.tbl_Cagri.Where(x => x.FIRMA_ID == firmaID && x.CAGRI_DURUM == true).Count();
            ViewBag.toplamAktifCagri = toplamAktifCagri;

            return PartialView(cagrilar);
        }

        public ActionResult LogOut() 
        {
            FormsAuthentication.SignOut();
            return RedirectToAction("Index","Login");
        }

        public ActionResult mesajlariGoruntule() 
        {
            var firmaMail = (string)Session["FIRMA_MAIL"];
            var firmaID = db.tbl_Firmalar.Where(x => x.FIRMA_MAIL == firmaMail).Select(x => x.ID_FIRMA).FirstOrDefault();

            var mesajlar = db.tbl_Mesajlar.Where(x => x.ALICI == firmaID).ToList();
            return View(mesajlar);
        }
    }
}