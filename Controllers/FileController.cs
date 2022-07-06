using DocumentsPdf.Filters;
using DocumentsPdf.Generators;
using Models;
using Serilog;
using System;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Hosting;
using System.Web.Http;

namespace DocumentsPdf.Controllers
{
    [RoutePrefix("file")]
    [ExceptionHandler]
    public class FileController : ApiController
    {
        private const string templatePath = "~/App_Data/FastReportTemplates/IndividualNotificationPdfTemplate.frx";

        private string IndividualNotificationAttachmentName => $"Уведомление об ограничении {DateTime.Now:d}.pdf";

        [Route("fl/notifications/attach")]
        [HttpPost]
        public async Task<IHttpActionResult> IndividualNotificationAttachment()
        {
            var provider = await Request.Content.ReadAsMultipartAsync();
            var metadata = await provider.Contents.Single(content => content.Headers.ContentDisposition.Name == "metadata").ReadAsAsync<NotificationMetadata>();
            var sign = await provider.Contents.Single(content => content.Headers.ContentDisposition.Name == "sign").ReadAsAsync<Sign>();
            var company = await provider.Contents.Single(content => content.Headers.ContentDisposition.Name == "company").ReadAsAsync<CompanyInfo>();

            var generator = new NotificationFileGenerator();
            var template = HostingEnvironment.MapPath(templatePath);
            var file = await generator.IndividualAsync(metadata, sign, company, template);
            var document = new FileDocumentView(IndividualNotificationAttachmentName, file);
            Log.Information("Сформирован файл: {FileName}", IndividualNotificationAttachmentName);

            return Ok(document);
        }
    }
}