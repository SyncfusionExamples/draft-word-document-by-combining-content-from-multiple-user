using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using DocumentEditorApp.Models;
using Microsoft.AspNetCore.Hosting;
using System.IO;
using EJ2DocumentEditor = Syncfusion.EJ2.DocumentEditor;

namespace DocumentEditorApp.Controllers
{
    public class HomeController : Controller
    {
       
        public IActionResult Index()
        {
          
            return View();
        }
        public IActionResult ButtonView(string userName)
        {
            ViewBag.userName = userName;
            return View();
        }
        public IActionResult Documenteditor(string serviceName,string  userName)
        {
            ViewBag.serviceName = serviceName;
            ViewBag.userName = userName;
            return View();
        }

    }
}
