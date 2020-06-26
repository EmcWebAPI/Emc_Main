using EmcReportWebApi.Config;
using EmcReportWebApi.Utils;
using EmcReportWebApi.Models;
using System;
using System.Diagnostics;
using System.Threading.Tasks;
using EmcReportWebApi.Business.ImplWordUtil;
using EmcReportWebApi.ReportComponent;
using EmcReportWebApi.ReportComponent.FirstPage;
using EmcReportWebApi.ReportComponent.Image;
using EmcReportWebApi.ReportComponent.ReviewTable;

namespace EmcReportWebApi.Business.Implement
{
    /// <summary>
    /// 报告实现类
    /// </summary>
    public class ReportImpl : ReportBase, IReport
    {
        /// <summary>
        /// 生成报告公共方法
        /// </summary>
        /// <param name="para"></param>
        /// <returns></returns>
        public ReportResult<string> CreateReport(ReportParams para)
        {
            Task<ReportResult<string>> task = new Task<ReportResult<string>>(() => CreateReportAsync(para));
            task.Start();
            ReportResult<string> result = task.Result;
            return result;
        }

        private ReportResult<string> CreateReportAsync(ReportParams para)
        {
            ReportResult<string> result;
            try
            {
                //线程池容量等待
                EmcConfig.SemLim.Wait();
                //计时
                TimerUtil tu = new TimerUtil(new Stopwatch());
                ReportInfo reportInfo = new ReportInfo(para);
                //生成报告
                string content = ReportJsonToWord(reportInfo);
                result = SetReportResult(string.Format(format: "报告生成成功,用时:" + tu.StopTimer()), true, content);
                EmcConfig.InfoLog.Info("报告:" + result.Content + ",信息:" + result.Message);
            }
            catch (Exception ex)
            {
                EmcConfig.ErrorLog.Error(ex.Message, ex);//设置错误信息
                result = SetReportResult($"报告生成失败,reportId:{para.ReportId},错误信息:{ex.Message}", false, "");
                return result;
            }
            finally
            {
                //保存参数用作排查bug
                SaveParams(para);
                EmcConfig.SemLim.Release();
            }
            return result;
        }

        /// <summary>
        /// Json格式转成word文件
        /// </summary>
        public string ReportJsonToWord(ReportInfo reportInfo)
        {
            //生成报告
            using (ReportHandleWord wordUtil = new ReportHandleWord(reportInfo.OutFileFullName, reportInfo.TemplateFileFullName))
            {
                //写首页内容
                ReportFirstPageAbstract reportFirstPage = reportInfo.ReportFirstPage;
                reportFirstPage.WriteFirstPage(wordUtil);

                //审查表信息(包含测试设备)
                ReviewTableInfoAbstract reviewTableInfo = reportInfo.ReviewTableInfo;
                reviewTableInfo.WriteReviewTableInfo(wordUtil);

                //实验数据
                var experimentBaseInfo = reportInfo.ExperimentInfo;
                experimentBaseInfo.WriteExperimentInfoAll(wordUtil);

                //识别标记和文件 从新文件中取
                ReviewTableInfoAbstract identityTableInfo = reportInfo.IdentityTableInfo;
                identityTableInfo.WriteReviewTableInfo(wordUtil);

                //样品图片
                ImageInfo imageInfo = reportInfo.ImageInfo;
                imageInfo.WriteImages(wordUtil);

                //替换页眉内容
                reportInfo.HandleReportHeader(wordUtil);

            }
            //删除中间件文件夹

            reportInfo.DeleteTemplateMiddleDirctory();

            return reportInfo.FileName;
        }
    }
}