using MicrosoftGraphFilesUpload.Helpers;
using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

namespace MicrosoftGraphFilesUpload.Controllers
{
    public class FileUploadController : BaseController
    {
        // GET: FileUpload
        [Authorize]
        public async Task<ActionResult> Index()
        {
            return View();
        }


        [HttpPost]
        public async Task<ActionResult> Index(HttpPostedFileBase file, string txtsiteurl)
        {

            if (file.ContentLength > 0)
            {
                //var fileName = Path.GetFileName(file.FileName);
                //var path = Path.Combine(Server.MapPath("~/App_Data/uploads"), fileName);
                //file.SaveAs(path);

                Stream fileStream = file.InputStream;
                bool isLargeFile = false;
                if ((file.ContentLength / (1024 * 1024)) > 3)
                    isLargeFile = true;
                try
                {
                    DriveItem uploadedFile = await GraphHelper.UploadFileAsync(fileStream, file.FileName, isLargeFile, txtsiteurl);
                    TempData["FileInfo"] = uploadedFile.WebUrl;
                }
                catch(Exception ex)
                {
                    ViewBag.ErrorMessage = ex.Message;
                }
            }

            //return RedirectToAction("Index");
            return View();
        }


    }
}