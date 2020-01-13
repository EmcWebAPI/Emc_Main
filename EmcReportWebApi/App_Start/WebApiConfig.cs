using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Http;

namespace EmcReportWebApi
{
    public static class WebApiConfig
    {
        public static void Register(HttpConfiguration config)
        {
            // Web API 配置和服务

            // Web API 路由
            config.MapHttpAttributeRoutes();

            config.Routes.MapHttpRoute(
                name: "DefaultApi",
                routeTemplate: "{controller}/{action}",
                defaults: new { controller="Report", action="Get", id = RouteParameter.Optional }
            );
        }
    }
}
