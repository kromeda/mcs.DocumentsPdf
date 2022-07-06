using System.Net.Http.Formatting;
using System.Web.Http;

namespace DocumentsPdf
{
    public static class WebApiConfig
    {
        public static void Register(HttpConfiguration config)
        {
            config.MapHttpAttributeRoutes();
            config.Formatters.Add(new JsonMediaTypeFormatter());
        }
    }
}
