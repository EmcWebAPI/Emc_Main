using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Http;
using System.Web.Http.Cors;

namespace EmcReportWebApi
{
    /// <summary>
    /// webapi配置
    /// </summary>
    public static class WebApiConfig
    {
        /// <summary>
        /// 注册信息
        /// </summary>
        /// <param name="config"></param>
        public static void Register(HttpConfiguration config)
        {
            //跨域配置
            config.EnableCors(new EnableCorsAttribute("*", "*", "*"));

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
