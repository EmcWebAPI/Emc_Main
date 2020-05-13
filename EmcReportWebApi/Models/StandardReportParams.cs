using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EmcReportWebApi.Models
{
    public class StandardReportParams
    {
        /// <summary>
        /// 原始记录(json字符串)
        /// </summary>
        public string OriginalRecord { get; set; }
        
        /// <summary>
        /// 解压文件请求路径
        /// </summary>
        public string ZipFilesUrl { get; set; }

        /// <summary>
        /// 回调函数请求路径  <br/>
        /// 
        /// 参数:<br/>
        /// reportId: 报告id<br/>
        /// contractId: 合同id<br/>
        /// status:状态 <br/>
        /// message:生成报告信息<br/>
        /// multipartFile:获取文件的key<br/>
        /// </summary>
        public string CallbackUrl { get; set; }

        /// <summary>
        /// 生成文件的json字符串                                        <br/>
        /// Json格式                                        <br/>
        ///  //首页内容                                                 <br/>
        ///  {                                                          <br/>
        ///  "firstPage": <br/>
        ///  <pre>...</pre>{ <br/> 
        /// <pre>...</pre><pre>...</pre>合同内容  ,<br/>
        ///<pre>...</pre><pre>...</pre> ReportCode//报告编号 push合同内容中 <br/> 
        /// <pre>...</pre>},                                                         <br/>
        ///     "standard":[],//标准json                                   <br/>
        ///     "attach":[],//附表json                                     <br/>
        ///     "yptp"://图片和说明                                       <br/>
        ///     <pre>...</pre>[                                                         <br/>
        ///     <pre>...</pre>        <pre>...</pre>{                                                         <br/>
        ///     <pre>...</pre>        <pre>...</pre>    <pre>...</pre>"fileName":"g",//图片名称                    <br/>
        ///     <pre>...</pre>        <pre>...</pre>    <pre>...</pre>"content":""//描述                             <br/>
        ///     <pre>...</pre>        <pre>...</pre>},  <pre>...</pre>                                                      <br/>
        ///     <pre>...</pre>        <pre>...</pre>{   <pre>...</pre>                                                      <br/>
        ///     <pre>...</pre>        <pre>...</pre>    <pre>...</pre>"fileName":"",                               <br/>
        ///     <pre>...</pre>        <pre>...</pre>    <pre>...</pre>"content":""                                   <br/>
        ///     <pre>...</pre>        <pre>...</pre>}                                                         <br/>
        ///     <pre>...</pre>]                                                          <br/>
        ///  }                                                    <br/>
        /// </summary>
        public JObject JsonObject { get; set; }
        
        //public string JsonStr { get; set; }
        
    }
}