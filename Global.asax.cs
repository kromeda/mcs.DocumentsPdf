using Microsoft.Win32;
using Serilog;
using Serilog.Events;
using System;
using System.Web;
using System.Web.Http;
using System.Web.Mvc;
using System.Web.Routing;

namespace DocumentsPdf
{
    public class WebApiApplication : HttpApplication
    {
        private readonly string seqHttpPath;
        private readonly string seqApiKey;

        public WebApiApplication()
        {
            RegistryView view = Environment.Is64BitOperatingSystem ? RegistryView.Registry64 : RegistryView.Registry32;
            RegistryKey localMachine = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view);
            seqHttpPath = localMachine.OpenSubKey(@"SOFTWARE\Avalanche").GetValue("SeqHttpPath") as string;
            seqApiKey = localMachine.OpenSubKey(@"SOFTWARE\Avalanche\DocumentsPdf").GetValue("SeqApiKey") as string;
        }

        protected void Application_Start()
        {
            AreaRegistration.RegisterAllAreas();
            GlobalConfiguration.Configure(WebApiConfig.Register);
            FilterConfig.RegisterGlobalFilters(GlobalFilters.Filters);
            RouteConfig.RegisterRoutes(RouteTable.Routes);

            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Override("Microsoft", LogEventLevel.Information)
                .Enrich.FromLogContext()
                .WriteTo.Seq(seqHttpPath, apiKey: seqApiKey)
                .CreateLogger();

            Log.Information("Запуск сервиса.");
        }
    }
}
