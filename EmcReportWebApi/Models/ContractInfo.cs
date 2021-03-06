﻿using System;
using System.Linq;

namespace EmcReportWebApi.Models
{
    /// <summary>
    /// 合同信息
    /// </summary>
    public class ContractInfo
    {
        #region 报告用的
        /// <summary>
        /// 
        /// </summary>
        public ContractData Data { get; set; }
        #endregion
        /// <summary>
        /// 
        /// </summary>
        public int Status { get; set; }
        /// <summary>
        ///
        /// </summary>
        public string Message { get; set; }
    }

    /// <summary>
    /// 合同内容
    /// </summary>
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
        private string _detectType;
        /// <summary>
        /// 检验类别
        /// </summary>
        public string DetectType
        {
            get
            {
                if (_detectType.Equals("GYJ", StringComparison.OrdinalIgnoreCase) && !SampleNumber.Equals("") && SampleNumber.Contains("-"))
                {

                    string[] stringSplit = SampleNumber.Split('-');
                    if (stringSplit.Length > 0)
                    {
                        string stringFirst = stringSplit[0];

                        _detectType = stringFirst.Substring(stringFirst.Length - 4, 4) + "年国家医疗器械抽检"+ (ReinspectionRPT ?? "");
                    }
                    else
                    {
                        _detectType = !_detectType.Contains(ReinspectionRPT) ? _detectType + (ReinspectionRPT ?? "") : _detectType;
                    }
                }
                else
                {
                    _detectType = !_detectType.Contains(ReinspectionRPT)?_detectType + (ReinspectionRPT ?? ""):_detectType;
                }

                if (!string.IsNullOrEmpty(ReinspectionRPT))
                {
                    _detectType = "国家监督抽查检验复验";
                }

                return _detectType;
            }
            set => _detectType = value;
        }


        /// <summary>
        /// "main_ypmc" 样品名称
        /// </summary>
        public string SampleName { get; set; }

        /// <summary>
        /// 产品编号/批号
        /// </summary>
        private string sampleNameRPT;

        /// <summary>
        /// 产品编号/批号
        /// </summary>
        public string SampleNameRPT
        {
            get
            {
                return sampleNameRPT ?? SampleName;
            }
            set => sampleNameRPT = value;
        }
        
        /// <summary>
        /// "main_xhgg" 规格型号/批号
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
                    return dtTime.ToString("yyyy年M月d日");
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
            get
            {
                DateTime dtTime;
                if (DateTime.TryParse(_sampleReceiptDate, out dtTime))
                {
                    return dtTime.ToString("yyyy年M月d日");
                }
                else
                {
                    return _sampleReceiptDate;
                }
            }
            set => _sampleReceiptDate = value;
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

        //private string _sampleProductionDate;
        ///// <summary>
        /////  "scrq" 生产日期
        ///// </summary>
        //public string SampleProductionDate
        //{
        //    get
        //    {
        //        DateTime dtTime;
        //        if (DateTime.TryParse(_sampleProductionDate, out dtTime))
        //        {
        //            return dtTime.ToString("yyyy年M月d日");
        //        }
        //        else
        //        {
        //            return _sampleProductionDate;
        //        }
        //    }
        //    set { _sampleProductionDate = value; }
        //}

        /// <summary>
        ///  "scrq" 生产日期
        /// </summary>
        public string SampleProductionDate { get; set; }

        /// <summary>
        /// 样品单位
        /// </summary>
        public string SampleUnit { get; set; }

        private string _sampleQuantity;
        /// <summary>
        /// "ypsl" /抽样数量
        /// </summary>
        public string SampleQuantity
        {
            get => _sampleQuantity==null||_sampleQuantity.Equals("") || SampleUnit ==null|| SampleUnit.Equals("") ?_sampleQuantity: _sampleQuantity+SampleUnit;
            set => _sampleQuantity = value;
        }

        /// <summary>
        /// "cyjs" 抽样基数
        /// </summary>
        public string SamplingBase
        {
            get => _samplingBase;
            set => _samplingBase = value;
        }
        
        /// <summary>
        /// "jydd" 检验地点
        /// </summary>
        public string AfterTreatmentMethod { get; set; }

        /// <summary>
        /// "sb" 商标
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

        /// <summary>
        /// zjgcs  主检工程师(检验)
        /// </summary>
        public string ChiefInspection { get; set; }


        private string _sampleAcquisitionMode;
        /// <summary>
        /// 抽样送样类型  1.送样syxz 2.抽样cyxz
        /// </summary>
        public string SampleAcquisitionMode
        {
            get
            {
                return _sampleAcquisitionMode;
            }

            set { _sampleAcquisitionMode = value; }
        }

        /// <summary>
        /// 送样
        /// </summary>

        private string _sampleAcquisitionModeSy;

        /// <summary>
        /// 送样
        /// </summary>
        public string SampleAcquisitionModeSy
        {
            get
            {
                if (_sampleAcquisitionMode.Equals("1"))
                {
                    _sampleAcquisitionModeSy = "送样（√）";
                }
                else
                {
                    _sampleAcquisitionModeSy = "送样（/）";
                }

                return _sampleAcquisitionModeSy;
            }
            set { _sampleAcquisitionModeSy = value; }
        }


        /// <summary>
        /// 抽样
        /// </summary>
        private string _sampleAcquisitionModeCy;

        /// <summary>
        /// 抽样
        /// </summary>
        public string SampleAcquisitionModeCy
        {
            get
            {
                _sampleAcquisitionModeCy = _sampleAcquisitionMode.Equals("1") ? "抽样（/）" : "抽样（√）";
                return _sampleAcquisitionModeCy;
            }
            set => _sampleAcquisitionModeCy = value;
        }

        /// <summary>
        /// 抽样单编号
        /// </summary>
        public string SamplingNumber
        {
            get => _samplingNumber.Length==13? $"{_samplingNumber.Substring(0, 2)}-{_samplingNumber.Substring(2, 2)}-{_samplingNumber.Substring(4, 5)}-{_samplingNumber.Substring(9, 2)}-{_samplingNumber.Substring(11, 2)}": _samplingNumber;
            set => _samplingNumber = value;
        }


        /// <summary>
        /// 型号规格
        /// </summary>
        private string sampleModelSpecificationRPT;

        /// <summary>
        /// 型号规格重写
        /// </summary>
        public string SampleModelSpecificationRPT
        {
            get
            {
                return sampleModelSpecificationRPT ?? SampleModelSpecification;

            }
            set { sampleModelSpecificationRPT = value; }
        }



        /// <summary>
        /// 产品编号/批号
        /// </summary>
        private string batchNumberRPT;

        private string _samplingBase;
        private string _samplingNumber;
        
        /// <summary>
        /// 产品编号/批号
        /// </summary>
        public string BatchNumberRPT
        {
            get
            {
                return batchNumberRPT ?? SampleModelSpecification;

            }
            set { batchNumberRPT = value; }
        }

        /// <summary>
        /// 复检
        /// </summary>
        public string ReinspectionRPT { get; set; }

        private string _sampleReceiptDateRpt;
        /// <summary>
        /// 实际到样日期
        /// </summary>
        public string SampleReceiptDateRPT
        {
            get => string.IsNullOrEmpty(_sampleReceiptDateRpt)?this._sampleReceiptDate: _sampleReceiptDateRpt;
            set => _sampleReceiptDateRpt = value;
        }
        /// <summary>
        /// 检验结果是否有加列
        /// </summary>
        public int? ColSpan { get; set; } = 0;

        /// <summary>
        /// 检验地点 jydd
        /// </summary>
        public string TestPoint
        {
            get=> string.IsNullOrEmpty(_testPoint) ? "本所实验室" : _testPoint;
            set => _testPoint = value;
        }

        private string _testPoint;

        #endregion
    }
}