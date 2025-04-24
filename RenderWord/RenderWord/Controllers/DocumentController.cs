using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Aspose.Words;
using Aspose.Words.Saving;
using System.IO;
using Aspose.Words.Replacing;

namespace RenderWord.Controllers
{
    public class DocumentController : Controller
    {
        [HttpGet]
        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public IActionResult Index(IFormFile uploadedFile, string fieldName, string fieldValue, float xCoordinate, float yCoordinate)
        {
            if (uploadedFile == null || uploadedFile.Length == 0)
            {
                return Content("No file selected.");
            }

            // Dua file Doc vao file wwwroot
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "uploads", uploadedFile.FileName);
            using (var stream = new FileStream(filePath, FileMode.Create))
            {
                uploadedFile.CopyTo(stream);
            }

            // Load file Word document
            var doc = new Document(filePath);

            // Them cac merged field 
            if (!string.IsNullOrEmpty(fieldName) && !string.IsNullOrEmpty(fieldValue))
            {
                // Thay the 1 FieldName hoac cu the bookmark
                ReplaceField(doc, fieldName, fieldValue);
            }

            // Luu dang kieu HTML
            using (var memoryStream = new MemoryStream())
            {
                var saveOptions = new HtmlSaveOptions
                {
                    ExportImagesAsBase64 = true,
                    SaveFormat = SaveFormat.Html
                };

                doc.Save(memoryStream, saveOptions);
                memoryStream.Position = 0;

                // Dua vao noi dung HTML vao file View
                using (var reader = new StreamReader(memoryStream))
                {
                    string htmlContent = reader.ReadToEnd();
                    ViewData["WordHtml"] = htmlContent;
                }
            }

            return View();
        }

        private void ReplaceField(Document doc, string fieldName, string fieldValue)
        {
            // Ensure we are using Aspose.Words.Range and not System.Range
            var range = doc.Range;

            // Find and replace the field (this can be a placeholder text or a bookmark)
            range.Replace(fieldName, fieldValue, new FindReplaceOptions());
        }
    }
}
