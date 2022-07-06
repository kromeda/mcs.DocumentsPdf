using Models;
using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Web.Hosting;

namespace DocumentsPdf.Generators
{
    public class NotificationFileGenerator
    {
        private static readonly AppDomain domain;
        private static readonly CultureInfo ruCulture;
        private static object individualSyncRoot = new object();

        static NotificationFileGenerator()
        {
            domain = AppDomain.CurrentDomain;
            domain.Load(Assembly.ReflectionOnlyLoadFrom(HostingEnvironment.MapPath("~/App_Data/FastReport.dll")).GetName());
            ruCulture = new CultureInfo("en-US");
        }

        public async Task<byte[]> IndividualAsync(NotificationMetadata metadata, Sign sign, CompanyInfo company, string templatePath)
        {
            var fastRerportAssembley = domain.GetAssemblies().FirstOrDefault(ass => ass.GetName().Name == "FastReport");
            var reportType = fastRerportAssembley.GetType("FastReport.Report");
            dynamic report = Activator.CreateInstance(reportType);
            report.Load(templatePath);

            report.SetParameterValue("Notification.NotificationNumber", metadata.Notification.NotificationNumber);
            report.SetParameterValue("Notification.NotificationDate", metadata.Notification.NotificationDate.ToShortDateString());
            report.SetParameterValue("Notification.DebtSum", metadata.Notification.DebtSum.ToString(ruCulture));
            report.SetParameterValue("Notification.DebtDate", metadata.Notification.DebtDate.ToShortDateString());

            bool hasRecentPayment = metadata.Notification.PaymentSum.HasValue && metadata.Notification.PaymentDate.HasValue;

            if (hasRecentPayment)
            {
                string recentPayment = $" (учтена последняя " +
                    $"оплата {metadata.Notification.PaymentSum?.ToString(ruCulture)} рублей " +
                    $"от {metadata.Notification.PaymentDate?.ToShortDateString()} г).";

                report.SetParameterValue("Notification.RecentPayment", recentPayment);
            }

            report.SetParameterValue("Department.DepartmentBranch", metadata.Department.Branch);
            report.SetParameterValue("Department.DepartmentOffice", metadata.Department.Office);
            report.SetParameterValue("Department.DepartmentAddress", metadata.Department.Address);
            report.SetParameterValue("Department.DepartmentPhone", metadata.Department.Phone);

            report.SetParameterValue("Customer.LS", metadata.Customer.ContractNumber);
            report.SetParameterValue("Customer.CustomerAddress", metadata.Customer.DeliveryAddress);
            report.SetParameterValue("Customer.PointAddress", metadata.Customer.PointAddress);

            report.SetParameterValue("Organization.Title", company.Title);
            report.SetParameterValue("Organization.INN", company.BankAccount.INN);
            report.SetParameterValue("Organization.KPP", company.BankAccount.KPP);
            report.SetParameterValue("Organization.BIK", company.BankAccount.BIK);
            report.SetParameterValue("Organization.Bank", company.BankAccount.DepartmentTitle);
            report.SetParameterValue("Organization.BankAccountNumber", company.BankAccount.AccountNumber);

            report.SetParameterValue("Sign.Rank", sign.BossPosition);
            report.SetParameterValue("Sign.Credentials", sign.BossName);

            dynamic picture = report.FindObject("SignaturePicture");
            picture.Image = new Bitmap(new MemoryStream(sign.Facsimile));

            // Возможно, сконфигурировать объект Report таким образом, чтобы он не показывал прогресс выполнения.
            // Поскольку прогресс является разделяемым ресурсом, доступ к нему должен предоставляться однопоточно.
            lock (individualSyncRoot)
            {
                report.Prepare();
            }

            Type pdfExportType = fastRerportAssembley.GetType("FastReport.Export.Pdf.PDFExport");
            dynamic export = Activator.CreateInstance(pdfExportType);

            using (var ms = new MemoryStream())
            {
                export.Export(report, ms);
                await ms.FlushAsync();
                return ms.ToArray();
            }
        }
    }
}