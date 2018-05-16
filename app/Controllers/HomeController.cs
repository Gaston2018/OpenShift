using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using app.Models;
using Microsoft.Extensions.Options;
using Microsoft.AspNetCore.Http;
using System.Threading;
using System.IO;
using System.Data;
using System.Net.Mail;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;
using DocumentFormat.OpenXml.Packaging;
using System.Text.RegularExpressions;
using System.Text;
using System.Globalization;

namespace app.Controllers
{
    public class HomeController : Controller
    {
        private readonly MassiveMailAppSetting _massiveMailAppSetting;

        public HomeController(IOptions<MassiveMailAppSetting> massiveMailAppSetting)
        {
            _massiveMailAppSetting = massiveMailAppSetting.Value;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
