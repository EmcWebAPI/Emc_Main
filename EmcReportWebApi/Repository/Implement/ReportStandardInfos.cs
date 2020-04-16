using Newtonsoft.Json.Linq;
using EmcReportWebApi.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using EmcReportWebApi.Models.Repository;
using System.Configuration;
using Newtonsoft.Json;

namespace EmcReportWebApi.Repository.Implement
{
    public class ReportStandardInfos:IReportStandardInfos
    {
        /// <summary>
        /// 获取合同信息
        /// </summary>
        /// <returns></returns>
        public ContractInfo GetContract(string contractId) {
            int datetime = int.Parse(DateTime.Now.ToString("yyyyMMdd"));
            string result = SyncHttpHelper.GetHttpResponse(string.Format("{0}?contractId={1}", ConfigurationManager.AppSettings["GetContractById"].ToString(),contractId), datetime);
            ContractInfo contract = JsonConvert.DeserializeObject<ContractInfo>(result);
            return contract;
        }
    }
}