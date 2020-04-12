﻿using EmcReportWebApi.App_Start;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Http;
using System.Web.Routing;

namespace EmcReportWebApi
{
    public class WebApiApplication : System.Web.HttpApplication
    {
        protected void Application_Start()
        {
            //配置log
            log4net.Config.XmlConfigurator.Configure(new System.IO.FileInfo(Server.MapPath("~/Web.config")));
            
            GlobalConfiguration.Configure(WebApiConfig.Register);
            AutoFacConfig.InitAutoFac();

        }
    }
}
