using Serilog;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Http.Filters;

namespace DocumentsPdf.Filters
{
    public class ExceptionHandlerAttribute : FilterAttribute, IExceptionFilter, IFilter
    {
        public Task ExecuteExceptionFilterAsync(HttpActionExecutedContext actionExecutedContext, CancellationToken cancellationToken)
        {
            if (actionExecutedContext.Exception != null)
            {
                Log.Error(actionExecutedContext.Exception, "Ошибка сервиса.");

                actionExecutedContext.Response = actionExecutedContext.Request.CreateResponse(HttpStatusCode.InternalServerError, new
                {
                    Reason = "Ошибка сервиса.",
                    Description = "При выполнении запроса вызникла ошибка. Обратитесь к разработчику."
                }, "application/json");
            }

            return Task.CompletedTask;
        }
    }
}