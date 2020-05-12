using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EmcReportWebApi.Models
{
    public class StandardReportParams
    {
        /// <summary>
        /// 报告编号
        /// </summary>
        public string ReportId { get; set; }

        /// <summary>
        /// 合同编号
        /// </summary>
        public string ContractId { get; set; }

        /// <summary>
        /// 文件解压路径
        /// </summary>
        public string ZipFilesUrl { get; set; }

        /// <summary>
        /// 生成文件的json字符串                                        <br/>
        /// Json格式                                        <br/>
        ///  //首页内容                                                 <br/>
        ///  {                                                          <br/>
        ///  "firstPage": 
        ///  <pre>...</pre>{                                             <br/>
        ///       <pre>...</pre><pre>...</pre>"bgbh": "",//报告编号 标准报告新追加                      <br/>
        ///       <pre>...</pre><pre>...</pre>"main_wtf": "",    //委托方                              <br/>
        ///       <pre>...</pre><pre>...</pre>"main_ypmc": "",  //样品名称                             <br/>
        ///       <pre>...</pre><pre>...</pre>"main_xhgg": "",   //规格型号                            <br/>
        ///       <pre>...</pre><pre>...</pre>"main_jylb": "",  //检验类别                             <br/>
        ///       <pre>...</pre><pre>...</pre>"ypmc": "",      //样品名称                              <br/>
        ///       <pre>...</pre><pre>...</pre>"sb": "",        //商标                                  <br/>
        ///       <pre>...</pre><pre>...</pre>"wtf": "",       //委托方                                <br/>
        ///       <pre>...</pre><pre>...</pre>"wtfdz": "",     //委托方地址                            <br/>
        ///       <pre>...</pre><pre>...</pre>"scdw": "",   //生产单位                                 <br/>
        ///       <pre>...</pre><pre>...</pre>"sjdw": "",   //受检单位                                 <br/>
        ///       <pre>...</pre><pre>...</pre>"cydw": "",     //抽样单位                               <br/>
        ///       <pre>...</pre><pre>...</pre>"cydd": "",     //抽样地点                               <br/>
        ///       <pre>...</pre><pre>...</pre>"cyrq": "",     //抽样日期                               <br/>
        ///       <pre>...</pre><pre>...</pre>"dyrq": "",     //到样日期                               <br/>
        ///       <pre>...</pre><pre>...</pre>"jyxm": "",     //检验项目                               <br/>
        ///       <pre>...</pre><pre>...</pre>"jyyj": "",//检验依据                                    <br/>
        ///       <pre>...</pre><pre>...</pre>"jyjl": "",//检验结论                                    <br/>
        ///       <pre>...</pre><pre>...</pre>"bz": "。",     //备注                                   <br/>
        ///       <pre>...</pre><pre>...</pre>"ypbh": "",     //样品编号                               <br/>
        ///       <pre>...</pre><pre>...</pre>"xhgg": "",     //型号规格                               <br/>
        ///       <pre>...</pre><pre>...</pre>"jylb": "",     //检验类别                               <br/>
        ///       <pre>...</pre><pre>...</pre>"cpbhph": "",     //产品编号/批号                        <br/>
        ///       <pre>...</pre><pre>...</pre>"cydbh": "",     //抽样单编号                            <br/>
        ///       <pre>...</pre><pre>...</pre>"scrq": "",     //生产日期                               <br/>
        ///       <pre>...</pre><pre>...</pre>"ypsl": "",     //抽样数量                               <br/>
        ///       <pre>...</pre><pre>...</pre>"cyjs": "",     //抽样基数                               <br/>
        ///       <pre>...</pre><pre>...</pre>"jydd": "",     //检验地点                               <br/>
        ///       <pre>...</pre><pre>...</pre>"jyrq": "",     //检验日期                               <br/>
        ///       <pre>...</pre><pre>...</pre>"ypms": ""     //样品描述                                <br/>
        ///       <pre>...</pre><pre>...</pre>"xhgghqtsm": ""     //型号规格或其他说明                 <br/>
        ///     <pre>...</pre>},                                                         <br/>
        ///     <pre>...</pre>"standard":[],//标准json                                   <br/>
        ///     <pre>...</pre>"attach":[],//附表json                                     <br/>
        ///     <pre>...</pre>"yptp"://图片和说明                                       <br/>
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
        public string JsonStr { get; set; }
        
    }
}