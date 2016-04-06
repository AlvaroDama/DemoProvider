using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using DemoProviderWeb.Models;

namespace DemoProviderWeb.Controllers
{
    public class HomeController : Controller
    {
        [SharePointContextFilter]
        public ActionResult Index()
        {
            User spUser = null;

            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);
            var data = new List<TelefonoViewModel>();

            using (var clientContext = spContext.CreateUserClientContextForSPHost())
            {
                if (clientContext != null)
                {
                    var telefonosList = clientContext.Web.Lists.GetByTitle("Telefonos");
                    clientContext.Load(telefonosList);
                    clientContext.ExecuteQuery();

                    var query = new CamlQuery();
                    var telefonosItems = telefonosList.GetItems(query);
                    clientContext.Load(telefonosItems);
                    clientContext.ExecuteQuery();

                    foreach (var item in telefonosItems)
                    {
                        data.Add(TelefonoViewModel.FromListItem(item));
                    }
                }
            }

            return View(data);
        }

        public ActionResult Add()
        {
            return View(new TelefonoViewModel());
        }

        [HttpPost]
        public ActionResult Add(TelefonoViewModel model)
        {
            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);
            
            using (var clientContext = spContext.CreateUserClientContextForSPHost())
            {
                if (clientContext != null)
                {
                    var telefonosList = clientContext.Web.Lists.GetByTitle("Telefonos");

                    var listCreationInfo = new ListItemCreationInformation();
                    var item = telefonosList.AddItem(listCreationInfo);
                    item["Title"] = model.Nombre;
                    item["Numero"] = model.Numero;
                    item.Update();
                    clientContext.ExecuteQuery();
                }
            }

            return RedirectToAction("Index", new {SPHostUrl = SharePointContext.GetSPHostUrl(HttpContext.Request).AbsoluteUri});
        }

        public ActionResult Delete(int id)
        {
            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);
            
            using (var clientContext = spContext.CreateUserClientContextForSPHost())
            {
                if (clientContext != null)
                {
                    var telefonosList = clientContext.Web.Lists.GetByTitle("Telefonos");
                    
                    var telefonosItem = telefonosList.GetItemById(id);
                    telefonosItem.DeleteObject();
                    clientContext.ExecuteQuery();
                }
            }
            return RedirectToAction("Index", new { SPHostUrl = SharePointContext.GetSPHostUrl(HttpContext.Request).AbsoluteUri });
        }

        public ActionResult Edit(int id)
        {
            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);
            TelefonoViewModel model = new TelefonoViewModel();

            using (var clientContext = spContext.CreateUserClientContextForSPHost())
            {
                if (clientContext != null)
                {
                    var telefonosList = clientContext.Web.Lists.GetByTitle("Telefonos");
                    
                    var item = telefonosList.GetItemById(id);
                    
                    clientContext.Load(item);
                    clientContext.ExecuteQuery();

                    model = TelefonoViewModel.FromListItem(item);
                }
            }

            return View(model);
        }

        [HttpPost]
        public ActionResult Edit(TelefonoViewModel model)
        {
            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);

            using (var clientContext = spContext.CreateUserClientContextForSPHost())
            {
                if (clientContext != null)
                {
                    var telefonosList = clientContext.Web.Lists.GetByTitle("Telefonos");

                    var item = telefonosList.GetItemById(model.Id);

                    item["Title"] = model.Nombre;
                    item["Numero"] = model.Numero;
                    item.Update();
                    clientContext.ExecuteQuery();
                }
            }
            return RedirectToAction("Index", new { SPHostUrl = SharePointContext.GetSPHostUrl(HttpContext.Request).AbsoluteUri });
        }

        public ActionResult About()
        {
            ViewBag.Message = "Addin MVC para SharePoint.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "[FAKE] Página de contacto";

            return View();
        }
    }
}
