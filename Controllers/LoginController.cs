using MvcFirmaCagri.Models.Entity;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Security;

namespace MvcFirmaCagri.Controllers
{
    public class LoginController : Controller
    {
        // GET: Login
        private DbisTakipEntities db = new DbisTakipEntities();
        public ActionResult Index()
        {
            //code
            return View();
        }
        [HttpPost]
        public ActionResult Index(tbl_Firmalar p)
        {
            var girisBilgileri = db.tbl_Firmalar.FirstOrDefault(x => x.FIRMA_MAIL == p.FIRMA_MAIL && x.FIRMA_SIFRE == p.FIRMA_SIFRE);

            if (girisBilgileri != null) 
            {
                FormsAuthentication.SetAuthCookie(girisBilgileri.FIRMA_MAIL, false);
                Session["FIRMA_MAIL"] = girisBilgileri.FIRMA_MAIL.ToString();
                return RedirectToAction("AktifCagrilar","Default");
            }
            else 
            {
                return RedirectToAction("Index");
            }
        }
    }
}