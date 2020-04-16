using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EmcReportWebApi.Models.Repository
{
    public class ContractInfo
    {
        #region 报告用的
        public ContractData Data { get; set; }
        #endregion

        public int Status { get; set; }

        public string Message { get; set; }
    }

    public class ContractData
    {
        #region 报告用
        /// <summary>
        /// "main_wtf" 委托方
        /// </summary>
        public string ContractClient { get; set; }

        /// <summary>
        /// "Main_jylb" 检验类别
        /// </summary>
        public string DetectType { get; set; }

        /// <summary>
        /// "main_ypmc" 样品名称
        /// </summary>
        public string SampleName { get; set; }

        /// <summary>
        /// "main_xhgg" 规格型号
        /// </summary>
        public string SampleModelSpecification { get; set; }

        /// <summary>
        /// "wtfdz" 委托方地址
        /// </summary>
        public string AddressOrIdCard { get; set; }

        /// <summary>
        /// "scdw" 生产单位
        /// </summary>
        public string ManufactureCompany { get; set; }

        /// <summary>
        ///  "sjdw" 受检单位
        /// </summary>
        public string DetectCompany { get; set; }

        /// <summary>
        /// "cydw" 抽样单位
        /// </summary>
        public string SamplingCompany { get; set; }

        /// <summary>
        /// "cydd" 抽样地点
        /// </summary>
        public string SamplingAddress { get; set; }

        /// <summary>
        /// "cyrq" 抽样日期
        /// </summary>
        public string SamplingDate { get; set; }

        /// <summary>
        /// "dyrq" 到样日期
        /// </summary>
        public string SampleReceiptDate { get; set; }

        /// <summary>
        /// "jyxm" 检验项目
        /// </summary>
        public string Content { get; set; }

        /// <summary>
        /// "jyyj" 检验依据
        /// </summary>
        public string SampleTestBasis { get; set; }

        /// <summary>
        /// "bz" 备注
        /// </summary>
        public string TestRemark { get; set; }

        /// <summary>
        /// "ypbh" 样品编号
        /// </summary>
        public string SampleNumber { get; set; }

        /// <summary>
        ///  "scrq" 生产日期
        /// </summary>
        public string SampleProductionDate { get; set; }

        /// <summary>
        /// "ypsl" /抽样数量
        /// </summary>
        public string SampleQuantity { get; set; }

        /// <summary>
        /// "cyjs" 抽样基数
        /// </summary>
        public string SamplingBase { get; set; }

        /// <summary>
        /// "jydd" 检验地点
        /// </summary>
        public string AfterTreatmentMethod { get; set; }

        /// <summary>
        /// "ypms" 样品描述
        /// </summary>
        public string SampleTrademark { get; set; }

        #endregion
    }
}