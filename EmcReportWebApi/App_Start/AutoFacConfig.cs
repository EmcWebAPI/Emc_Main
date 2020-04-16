/***
 * autofac 4.9.4
 * autofac 4.3.1
 * */

using Autofac;
using Autofac.Integration.WebApi;
using EmcReportWebApi.Business;
using EmcReportWebApi.Business.Implement;
using EmcReportWebApi.Repository;
using EmcReportWebApi.Repository.Implement;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Web;
using System.Web.Http;

namespace EmcReportWebApi.App_Start
{
    public static class AutoFacConfig
    {
        public static void InitAutoFac() {
            var configuration = GlobalConfiguration.Configuration;

            var builder = new ContainerBuilder();
            builder.RegisterType<ReportImpl>().As<IReport>().AsImplementedInterfaces();
            builder.RegisterType<ReportStandardImpl>().As<IReportStandard>().AsImplementedInterfaces();
            builder.RegisterType<ReportStandardInfos>().As<IReportStandardInfos>().AsImplementedInterfaces();

            builder.RegisterApiControllers(Assembly.GetExecutingAssembly());
            
            IContainer container = builder.Build();
            //将依赖关系解析器设置为Autofac。
            var resolver = new AutofacWebApiDependencyResolver(container);
            configuration.DependencyResolver = resolver;
        }
    }
}