using EmcReportWebApi.Business;
using EmcReportWebApi.Business.Implement;
using EmcReportWebApi.Common;
using EmcReportWebApi.Controllers;
using EmcReportWebApi.Repository;
using EmcReportWebApi.Repository.Implement;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using System.Web;

namespace EmcReportWebApi.App_Start
{
    public static class TimerConfig
    {

        public static void InitTimer() {
            System.Timers.Timer timer = new System.Timers.Timer();
            timer.Enabled = true;
            timer.Interval = 1000; //执行间隔时间,单位为毫秒; 这里实际间隔为10分钟  
            timer.Start();
            timer.Elapsed += new System.Timers.ElapsedEventHandler(TestTask);
        }

        private static List<Task> tasks = new List<Task>();

        private static void TestTask(object source, ElapsedEventArgs e)
        {
            //EmcConfig.KillWordProcess();

            while (EmcConfig.TaskQueue.Count > 0) {
                //TestTask2();
                if (tasks.Count <= 4) {
                    Guid guid = EmcConfig.TaskQueue.Dequeue();
                    Task<string> task = null;
                    try
                    {
                        task = new Task<string>(()=>TestTask2(1));
                        tasks.Add(task);
                        task.Start();
                        string result = task.Result;
                        
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                    finally
                    {
                        tasks.Remove(task);
                    }
                } 
            }
        }
        /// <summary>
        /// 测试和返回值 
        /// </summary>
        /// <param name="tt"></param>
        private static string TestTask2(int tt) {
            try
            {
                string reportFilesPath = FileUtil.CreateReportDirectory(string.Format("{0}Files\\ReportFiles", EmcConfig.CurrRoot));
                string reportZipFilesPath = string.Format("{0}Files\\ReportFiles\\Test\\{1}", EmcConfig.CurrRoot, "QT2019-3015.zip");
                //解压zip文件
                ZipFileHelper.DecompressionZip(reportZipFilesPath, reportFilesPath);
                IReportStandardInfos reportInfos = new ReportStandardInfos();
                IReportStandard _reportStandard = new ReportStandardImpl(reportInfos);
                string result = _reportStandard.JsonToWordStandard("QT2019-3015", jsonStr, reportFilesPath);

                return "12312312";
            }
            catch (Exception ex)
            {

                throw ex;
            }
           
        }

        private static string jsonStr = "{\"yptp\":[{\"fileName\":\"6da7e28a-e492-4a8a-8ac3-2bb03e05b132.jpg\",\"content\":\"外观\"},{\"fileName\":\"75aebfd8-6184-4678-bce9-8e38baeb8090.jpg\",\"content\":\"铭牌\"},{\"fileName\":\"6da7e28a-e492-4a8a-8ac3-2bb03e05b132.jpg\",\"content\":\"外观\"},{\"fileName\":\"75aebfd8-6184-4678-bce9-8e38baeb8090.jpg\",\"content\":\"铭牌\"}],\"attach\":[{\"col1\":\"报警状态\",\"col2\":\"报警类型\",\"col3\":\"指示灯颜色\",\"col4\":\"检验结果\",\"col5\":\"闪烁频率\",\"col6\":\"检验结果\",\"col7\":\"占空比\",\"col8\":\"检验结果\"},{\"col1\":\"断电报警\",\"col2\":\"高优先级\",\"col3\":\"红色\",\"col4-input\":\"红色\",\"col5\":\"1.4Hz~2.8Hz\",\"col6-input\":\"\",\"col7\":\"20%~60%\",\"col8-input\":\"\"},{\"col1\":\"肤温传感器断开\",\"col2\":\"高优先级\",\"col3\":\"红色\",\"col4-input\":\"红色\",\"col5\":\"1.4Hz~2.8Hz\",\"col6-input\":\"\",\"col7\":\"20%~60%\",\"col8-input\":\"\"}],\"firstPage\":{\"main_wtf\":\"国家药品监督管理局\",\"main_ypmc\":\"按产品标识\",\"main_xhgg\":\"按产品标识\",\"main_jylb\":\"2020年国家医疗器械抽检\",\"ypmc\":\"按产品标识\",\"sb\":\"\",\"wtf\":\"国家药品监督管理局\",\"wtfdz\":\"北京市西城区展览路北露园1号\",\"scdw\":\"按产品标识\",\"sjdw\":\"按抽样单上公章\",\"cydw\":\"按抽样单上公章\",\"cydd\":\"按抽样单上公章\",\"cyrq\":\"2020年*月*日\",\"dyrq\":\"2020年*月*日\",\"jyxm\":\"药监综械管〔2020〕*号文附件*《2020年国家医疗器械抽检(中央补助地方项目)产品检验方案》中“30200.无创自动测量血压计（电子血压计）”的检验项目\",\"jyyj\":\"药监综械管〔2020〕*号文附件*《2020年国家医疗器械抽检(中央补助地方项目)产品检验方案》中“30200.无创自动测量血压计（电子血压计）”的检验依据\",\"jyjl\":\"合格/不合格\",\"bz\":\"1）报告中的“——”表示此项不适用，报告中“/”表示此项空白。\",\"ypbh\":\"GYJ2020-****\",\"xhgg\":\"按产品标识\",\"jylb\":\"2020年国家医疗器械抽检\",\"cpbhph\":\"按产品标识形式、内容\",\"cydbh\":\"按产品标识形式、内容\",\"scrq\":\"按产品标识形式、内容\",\"ypsl\":\"按产品标识形式、内容\",\"cyjs\":\"按产品标识形式、内容\",\"jydd\":\"本所实验室\",\"jyrq\":\"2018年5月22日~2018年7月13日\",\"jydd\":\"本所实验室\",\"ypms\":\"1、被检样品封样完好。\r\n 2、本次检测包含下列部件：主机、袖带（根据实际情况填写）。 \",\"xhgghqtsm\":\"检测结果不包括不确定度的估算值。\"},\"standard\":[{\"itemId\":\"1\",\"idxNo\":\"8\",\"itemContent\":\"自动复位装置的选择\",\"itemPath\":\"1|\",\"comment\":\"12312312\",\"reMark\":\"44444\",\"list\":[{\"itemId\":\"2\",\"stdName\":\"GB9706.1-2007\",\"stdItmNo\":\"49\",\"itemPath\":\"1|2|\",\"list\":[{\"itemId\":\"3\",\"stdName\":\"GB9706.1-2007\",\"itemContent\":\"如果使用自动复位热断路器和过电流释放器,自动复位能保证安全.\",\"itemPath\":\"1|2|3|\",\"reference\":\"1\",\"list\":[]}]}]},{\"itemId\":\"4\",\"idxNo\":\"9\",\"itemContent\":\"电源中断后的复位\",\"itemPath\":\"4|\",\"comment\":\"12312312\",\"reMark\":\"44444\",\"list\":[{\"itemId\":\"5\",\"stdName\":\"GB9706.1-2007\",\"stdItmNo\":\"49\",\"itemPath\":\"4|5|\",\"list\":[{\"itemId\":\"6\",\"stdName\":\"GB9706.1-2007\",\"itemContent\":\"设备应设计成当电源供电中断后又恢复时,除预定功能中断外,不会发生安全方面危险.\",\"itemPath\":\"4|5|6|\",\"reference\":\"2\",\"list\":[]}]}]},{\"itemId\":\"7\",\"idxNo\":\"10\",\"itemContent\":\"指示器\",\"itemPath\":\"7|\",\"list\":[{\"itemId\":\"8\",\"stdName\":\"GB9706.1-2007\",\"stdItmNo\":\"57\",\"itemPath\":\"7|8|\",\"list\":[{\"itemId\":\"9\",\"stdName\":\"GB9706.1-2007\",\"itemContent\":\"除非对位于正常操作位置的操作者另有显而易见的指标,否则应安装指示灯,用于:\n----- 指示设备已通电.\",\"itemPath\":\"7|8|9|\",\"reference\":\"3\",\"list\":[]},{\"itemId\":\"10\",\"stdName\":\"GB9706.1-2007\",\"itemContent\":\"----设备装有不发光的电热器如会产生安全方面危险时,指示电热器已工作.\",\"itemPath\":\"7|8|10|\",\"reference\":\"4\",\"list\":[]},{\"itemId\":\"11\",\"stdName\":\"GB9706.1-2007\",\"itemContent\":\"----当输出电路意外的或长时间的工作可能引起安全方面危险时,指示处于输出状态.\",\"itemPath\":\"7|8|11|\",\"reference\":\"5\",\"list\":[]},{\"itemId\":\"12\",\"stdName\":\"GB9706.1-2007\",\"itemContent\":\"指示充电装置工作状态.\",\"itemPath\":\"7|8|12|\",\"reference\":\"6\",\"list\":[]}]}]},{\"itemId\":\"13\",\"idxNo\":\"7\",\"itemContent\":\"连续漏电流和患者辅助电流(正常工作温度下)\",\"itemPath\":\"13|\",\"list\":[{\"itemId\":\"14\",\"stdName\":\"GB9706.1-2007\",\"stdItmNo\":\"19\",\"itemPath\":\"13|14|\",\"list\":[{\"itemId\":\"15\",\"stdName\":\"GB9706.1-2007\",\"itemContent\":\"患者漏电流\",\"itemPath\":\"13|14|15|\",\"reference\":\"7\",\"list\":[{\"itemId\":\"16\",\"stdName\":\"GB9706.1-2007\",\"itemContent\":\"直流\",\"itemPath\":\"13|14|15|16|\",\"reference\":\"7\",\"list\":[{\"itemId\":\"17\",\"stdName\":\"GB9706.1-2007\",\"itemContent\":\"正常状态下≤0.01mA\",\"itemPath\":\"13|14|15|16|17|\",\"reference\":\"7\",\"list\":[],\"result\":\"测试检验结果1\"},{\"itemId\":\"18\",\"stdName\":\"GB9706.1-2007\",\"itemContent\":\"单一故障状态下≤0.05mA\",\"itemPath\":\"13|14|15|16|18|\",\"reference\":\"8\",\"result\":\"测试检验结果1\",\"list\":[]}]},{\"itemId\":\"19\",\"stdName\":\"GB9706.1-2007\",\"itemContent\":\"交流\",\"itemPath\":\"13|14|15|19|\",\"reference\":\"9\",\"list\":[{\"itemId\":\"20\",\"stdName\":\"GB9706.1-2007\",\"itemContent\":\"正常状态下≤0.1mA\",\"itemPath\":\"13|14|15|19|20|\",\"reference\":\"9\",\"list\":[]},{\"itemId\":\"21\",\"stdName\":\"GB9706.1-2007\",\"itemContent\":\"单一故障状态下≤0.5mA\",\"itemPath\":\"13|14|15|19|21|\",\"reference\":\"10\",\"list\":[]}]},{\"itemId\":\"22\",\"stdName\":\"GB9706.1-2007\",\"itemContent\":\"应用部分加压状态≤5mA\",\"itemPath\":\"13|14|15|22|\",\"reference\":\"11\",\"list\":[]},{\"itemId\":\"23\",\"stdName\":\"GB9706.1-2007\",\"itemContent\":\"信号输入/出部分加压状态≤{$1}mA\",\"itemPath\":\"13|14|15|23|\",\"reference\":\"12\",\"list\":[]}]},{\"itemId\":\"24\",\"stdName\":\"GB9706.1-2007\",\"itemContent\":\"患者辅助电流\n单位: mA\",\"itemPath\":\"13|14|24|\",\"reference\":\"13\",\"list\":[{\"itemId\":\"25\",\"stdName\":\"GB9706.1-2007\",\"itemContent\":\"直流\",\"itemPath\":\"13|14|24|25|\",\"reference\":\"13\",\"list\":[{\"itemId\":\"26\",\"stdName\":\"GB9706.1-2007\",\"itemContent\":\"正常状态下≤0.01mA\",\"itemPath\":\"13|14|24|25|26|\",\"reference\":\"13\",\"list\":[]},{\"itemId\":\"27\",\"stdName\":\"GB9706.1-2007\",\"itemContent\":\"单一故障状态下≤0.05mA\",\"itemPath\":\"13|14|24|25|27|\",\"reference\":\"14\",\"list\":[]}]},{\"itemId\":\"28\",\"stdName\":\"GB9706.1-2007\",\"itemContent\":\"交流\",\"itemPath\":\"13|14|24|28|\",\"reference\":\"15\",\"list\":[{\"itemId\":\"29\",\"stdName\":\"GB9706.1-2007\",\"itemContent\":\"正常状态下≤0.1mA\",\"itemPath\":\"13|14|24|28|29|\",\"reference\":\"15\",\"list\":[]},{\"itemId\":\"30\",\"stdName\":\"GB9706.1-2007\",\"itemContent\":\"单一故障状态下≤0.5mA\",\"itemPath\":\"13|14|24|28|30|\",\"reference\":\"16\",\"list\":[]}]}]}]}]}]}";
    }
}