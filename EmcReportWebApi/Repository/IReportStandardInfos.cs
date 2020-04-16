using EmcReportWebApi.Models.Repository;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EmcReportWebApi.Repository
{
    public interface IReportStandardInfos
    {
        /// <summary>
        /// 获取合同信息
        /// </summary>
        /// <returns></returns>
        ContractInfo GetContract(string contractId);
    }
}