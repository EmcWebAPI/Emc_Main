using System;

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
        /// bgbh 报告编号
        /// </summary>
        public string ReportCode { get; set; }

        /// <summary>
        /// "main_wtf" 委托方
        /// </summary>
        public string ContractClient { get; set; }

        /// <summary>
        /// "Main_jylb" 检验类别
        /// </summary>
        private string detectType;
        public string DetectType
        {
            get {
                if (detectType.Equals("GYJ", StringComparison.OrdinalIgnoreCase)&& !SampleNumber.Equals("")&&SampleNumber.Contains("-")) {

                    string[] stringSplit = SampleNumber.Split('-');
                    if (stringSplit.Length > 0)
                    {
                        string stringFirst = stringSplit[0];
                        detectType = stringFirst.Substring(stringFirst.Length - 4, 4)+ "年国家医疗器械抽检";
                    }
                }

                return detectType; }
            set { detectType = value; }
        }


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


        private string _samplingDate;
        /// <summary>
        /// "cyrq" 抽样日期
        /// </summary>
        public string SamplingDate
        {
            get
            {
                DateTime dtTime;
                if (DateTime.TryParse(_samplingDate, out dtTime))
                {
                    return dtTime.ToString("yyyy年MM月dd日");
                }
                else
                {
                    return _samplingDate;
                }
            }
            set { _samplingDate = value; }
        }

        private string _sampleReceiptDate;
        /// <summary>
        /// "dyrq" 到样日期
        /// </summary>
        public string SampleReceiptDate
        {
            get {
                DateTime dtTime;
                if (DateTime.TryParse(_sampleReceiptDate, out dtTime))
                {
                    return dtTime.ToString("yyyy年MM月dd日");
                }
                else
                {
                    return _sampleReceiptDate;
                }
             }
            set { _sampleReceiptDate = value; }
        }


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

        private string _sampleProductionDate;
        /// <summary>
        ///  "scrq" 生产日期
        /// </summary>
        public string SampleProductionDate
        {
            get {
                DateTime dtTime;
                if (DateTime.TryParse(_sampleProductionDate, out dtTime))
                {
                    return dtTime.ToString("yyyy年MM月dd日");
                }
                else
                {
                    return _sampleProductionDate;
                }
            }
            set { _sampleProductionDate = value; }
        }


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

        /// <summary>
        /// jyjl 检验结论
        /// </summary>
        public string Conclusion { get; set; }

        /// <summary>
        /// ypms 样品描述
        /// </summary>
        public string SampleDescription { get; set; }

        /// <summary>
        /// xhgghqtsm 型号规格及其他说明
        /// </summary>
        public string OtherDescription { get; set; }

        /// <summary>
        /// jyrq 检验日期
        /// </summary>
        public string InspectionDate { get; set; }
        #endregion
    }
}