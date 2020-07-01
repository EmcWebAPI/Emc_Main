using Swashbuckle.Swagger;
using System.Linq;
using System.Web.Http.Description;

namespace EmcReportWebApi.Config
{

    /// <summary>
    /// swagger隐藏api特性
    /// </summary>
    [System.AttributeUsage(System.AttributeTargets.Method | System.AttributeTargets.Class)]
    public partial class HiddenApiAttribute : System.Attribute { }
    /// <summary>
    /// swagger隐藏api
    /// </summary>
    public class HiddenApiFilter : IDocumentFilter
    {
        /// <summary>
        /// 隐藏api应用
        /// </summary>
        public void Apply(SwaggerDocument swaggerDoc, SchemaRegistry schemaRegistry, IApiExplorer apiExplorer)
        {
            foreach (ApiDescription apiDescription in apiExplorer.ApiDescriptions)
            {
                if (Enumerable.OfType<HiddenApiAttribute>(apiDescription.GetControllerAndActionAttributes<HiddenApiAttribute>()).Any())
                {
                    string key = "/" + apiDescription.RelativePath;
                    if (key.Contains("?"))
                    {
                        int idx = key.IndexOf("?", System.StringComparison.Ordinal);
                        key = key.Substring(0, idx);
                    }
                    swaggerDoc.paths.Remove(key);
                }
            }
        }
    }
}