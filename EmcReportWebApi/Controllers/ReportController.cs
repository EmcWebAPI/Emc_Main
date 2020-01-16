using EmcReportWebApi.Common;
using EmcReportWebApi.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web;
using System.Web.Http;

namespace EmcReportWebApi.Controllers
{
    public class ReportController : ApiController
    {
        private string jsonStr = "{\"firstPage\":{\"main_wtf\":\"飞利浦(中国)投资有限公司1\",\"main_ypmc\":\"病人监护仪1\",\"main_xhgg\":\"M8102A1\",\"main_jylb\":\"委托检验1\",\"ypmc\":\"病人监护仪\",\"sb\":\"\",\"wtf\":\"飞利浦（中国）投资有限公司\",\"wtfdz\":\"上海市静安区灵石路718号A幢\",\"scdw\":\"PhilipsMedizinSystemeBoeblingenGmbH\",\"sjdw\":\"PhilipsMedizinSystemeBoeblingenGmbH\",\"cydw\":\"\",\"cydd\":\"\",\"cyrq\":\"\",\"dyrq\":\"2018年5月8日\",\"jyxm\":\"YY0505全项目、YY0601中36、YY0667中36、YY0668中36、YY0783中36、YY0784中36\",\"jyyj\":\"YY0505-2012《医用电气设备第1-2部分：安全通用要求并列标准电磁兼容要求和试验》、YY0601-2009《医用电气设备呼吸气体监护仪的基本安全和主要性能专用要求》、YY0667-2008《医用电气设备第2-30部分：自动循环无创血压监护设备的安全和基本性能专用要求》、YY0668-2008《医用电气设备第2-49部分：多参数患者监护设备安全专用要求》、YY0783-2010《医用电气设备第2-34部分：有创血压监测设备的安全和基本性能专用要求》、YY0784-2010《医用电气设备医用脉搏血氧仪设备基本安全和主要性能专用要求》\",\"jyjl\":\"被检样品符合YY0505-2012标准要求、符合YY0601-2009标准第36章要求、符合YY0667-2008标准第36章要求、符合YY0668-2008标准第36章要求、符合YY0783-2008标准第36章要求、符合YY0784-2010标准第36章要求\",\"bz\":\"报告中“/”表示此项空白，“—”表示不适用。\",\"ypbh\":\"QW2018-0698\",\"xhgg\":\"M8102A\",\"jylb\":\"委托检验\",\"cpbhph\":\"DE65528125\",\"cydbh\":\"\",\"scrq\":\"2018-02-16\",\"ypsl\":\"1台\",\"cyjs\":\"\",\"jydd\":\"本所实验室\",\"jyrq\":\"2018年5月22日~2018年7月13日\",\"jydd\":\"本所实验室\",\"ypms\":\"见本报告第3页“1受检样品信息”。\",\"xhgghqtsm\":\"1.检测结果不包括不确定度的估算值。2.ECG附件有63个型号：M1631A、M1671A、M1984A、M1611A、M1968A、M1625A、M1639A、M1675A、M1602A、M1974A、M1601A、M1635A、M1678A、M1976A、M1672A、M1673A、M1533A、M1971A、M1973A、M1684A、M1613A、M1681A、M1558A、M1609A、M1683A、M1621A、M1674A、M1604A、M1685A、M1603A、M1619A、M1669A、M1645A、M1510A、M1500A、M1520A、M1979A、M1530A、M1557A、M1644A、M1605A、M1680A、M1537A、M1647A、M1532A、M1978A、M1615A、M1633A、M1668A、M1629A、M1663A、M1667A、M1623A、M1538A、M1665A、M1682A、M1540C、M1550C、M1560C、M1570C、989803170171、989803170181、989803143201。其电气原理和材料组成完全一致,   仅导联数与长度有所区别。本次检测了M1663A，M1978A，M1971A。SpO2附件有5个型号：M1192A、M1193A、M1194A、M1195A、M1196A，其电气原理和材料组成完全一致，仅长度和适用人群有所区别。本次检测了M1196A。CO2附件有17个型号：M2516A、M2761A、M2772A、M2751A、M2750A、M2745A、M2756A、M2757A、M2501A、M2768A、M2773A、M2741A、M2536A、M2746A、M2776A、M2777A、M1920A。其产品结构及原理均相同。本次检测了M2741A。温度探头有11个型号：21075A、21076A、21078A、M1837A、21091A、21093A、21094A、21095A、21090A、21082A、21082B。其电气原理和材料组成完全一致，仅长度和适用范围有所区别，本次检测了M21075A。袖带（含连接管）共有8个型号：M1571A、M1572A、M1573A、M1574A、M1575A、M1576A、M1598B、M1599B。其电气原理及材料组成完全一致，仅围度和连接管长度有所区别。本次检测了M1598B和M1574A。\",\"sjyp_ypmc\":\"病人监护仪\",\"sjyp_ypxh\":\"M8102A\",\"sjyp_ypbhph\":\"DE65528125\",\"sjyp_srdy\":\"AC100-240V\",\"sjyp_pl\":\"50/60Hz\",\"sjyp_edsrglhdl\":\"1.3-0.7A\",\"sjyp_dclx\":\"锂锰电池\",\"sjyp_gddy\":\"DC11.1V\",\"sjyp_ypcc\":\"199mm×146mm×89mm\"},\"ypgcList\":[{\"xh\":\"1\",\"bjmc\":\"主机\",\"bjfl\":\"\",\"xhbbh\":\"M8102A\",\"xlh\":\"DE65528125\",\"bz\":\"\"},{\"xh\":\"2\",\"bjmc\":\"模块\",\"bjfl\":\"\",\"xhbbh\":\"M3014A\",\"xlh\":\"DE45454454\",\"bz\":\"\"},{\"xh\":\"2\",\"bjmc\":\"模块\",\"bjfl\":\"\",\"xhbbh\":\"M3015B\",\"xlh\":\"DE45619953\",\"bz\":\"\"},{\"xh\":\"3\",\"bjmc\":\"外部电源配件\",\"bjfl\":\"\",\"xhbbh\":\"M8023A\",\"xlh\":\"DE21977324\",\"bz\":\"\"},{\"xh\":\"4\",\"bjmc\":\"锂电子电池\",\"bjfl\":\"\",\"xhbbh\":\"M4607A\",\"xlh\":\"\",\"bz\":\"\"},{\"xh\":\"5\",\"bjmc\":\"外接电池盒\",\"bjfl\":\"\",\"xhbbh\":\"865297/M4605A\",\"xlh\":\"865297：DE43610244\",\"bz\":\"\"},{\"xh\":\"6\",\"bjmc\":\"ECG附件\",\"bjfl\":\"3导联\",\"xhbbh\":\"M1631A，M1671A，M1611A，M1675A，M1601A，M1678A，M1672A，M1673A，M1613A，M1609A，M1674A，M1603A，M1619A，M1605A，M1615A，\",\"xlh\":\"\",\"bz\":\"M1663A，M1978A，M1971A为受检样品\"},{\"xh\":\"6\",\"bjmc\":\"ECG附件\",\"bjfl\":\"5导联\",\"xhbbh\":\"M1968A，M1625A，M1639A，M1974A，M1635A，M1971A，M1973A，M1621A，M1645A，M1647A，M1633A，M1629A，\",\"xlh\":\"\",\"bz\":\"M1663A，M1978A，M1971A为受检样品\"},{\"xh\":\"6\",\"bjmc\":\"ECG附件\",\"bjfl\":\"6导联\",\"xhbbh\":\"M1684A，M1681A，M1683A，M1685A， M1644A，M1680A，\",\"xlh\":\"\",\"bz\":\"M1663A，M1978A，M1971A为受检样品\"},{\"xh\":\"6\",\"bjmc\":\"ECG附件\",\"bjfl\":\"10导联\",\"xhbbh\":\"M1984A，M1602A，M1976A，M1533A，M1558A，M1604A，M1979A，M1557A，M1537A，M1532A，M1978A，M1663A，\",\"xlh\":\"\",\"bz\":\"M1663A，M1978A，M1971A为受检样品\"},{\"xh\":\"6\",\"bjmc\":\"ECG附件\",\"bjfl\":\"中继电缆\",\"xhbbh\":\"M1669A，M1665A，M1510A，M1520A，M1500A，M1530A，M1668A，M1540C，M1550C，M1560C，M1570C，989803170171， 989803170181\",\"xlh\":\"\",\"bz\":\"M1663A，M1978A，M1971A为受检样品\"},{\"xh\":\"7\",\"bjmc\":\"无创血压袖带\",\"bjfl\":\"\",\"xhbbh\":\"M1192A，M1193A，M1194A，M1195A，M1196A\",\"xlh\":\"\",\"bz\":\"M1196A为受检样品\"},{\"xh\":\"8\",\"bjmc\":\"血氧饱和度传感器\",\"bjfl\":\"\",\"xhbbh\":\"M1571A，M1572A，M1573A，M1574A，M1575A，M1576A\",\"xlh\":\"\",\"bz\":\"M1574A，M1599B为受检样品\"},{\"xh\":\"8\",\"bjmc\":\"血氧饱和度传感器\",\"bjfl\":\"\",\"xhbbh\":\"M1598B，M1599B\",\"xlh\":\"\",\"bz\":\"M1574A，M1599B为受检样品\"},{\"xh\":\"9\",\"bjmc\":\"温度探针\",\"bjfl\":\"\",\"xhbbh\":\"21075A，21076A，21078A，M1837A，21091A，21093A，21094A，21095A，21090A，\",\"xlh\":\"\",\"bz\":\"21075A为受检样品\"},{\"xh\":\"10\",\"bjmc\":\"二氧化碳附件\",\"bjfl\":\"\",\"xhbbh\":\"M2501A，M2516A，M2536A\",\"xlh\":\"\",\"bz\":\"M2741A为受检样品\"},{\"xh\":\"10\",\"bjmc\":\"二氧化碳附件\",\"bjfl\":\"\",\"xhbbh\":\"M2741A，M2745A，M2746A，M2750A，M2751A，M2756A，M2757A，M2761A，M2768A，M2772A，M2773A，M2776A，M2777A\",\"xlh\":\"\",\"bz\":\"M2741A为受检样品\"},{\"xh\":\"10\",\"bjmc\":\"二氧化碳附件\",\"bjfl\":\"\",\"xhbbh\":\"M1920A\",\"xlh\":\"\",\"bz\":\"M2741A为受检样品\"}],\"ypyxList\":[{\"msbh\":\"①\",\"msmc\":\"工作模式\",\"msms\":\"主机（心电+心率+呼吸+无创血压+脉搏血氧饱和度+主流/侧流二氧化碳）+外部电源配件+ M3015B（微流二氧化碳+有创血压+体温）\",\"bz\":\"网电源供电\"},{\"msbh\":\"②\",\"msmc\":\"工作模式\",\"msms\":\"主机（心电+心率+呼吸+无创血压+脉搏血氧饱和度+主流/侧流二氧化碳）+M3015B（微流二氧化碳+有创血压+体温）\",\"bz\":\"内部电源供电\"},{\"msbh\":\"③\",\"msmc\":\"工作模式\",\"msms\":\"主机（心电+心率+呼吸+无创血压+脉搏血氧饱和度+主流/侧流二氧化碳）+外部电源配件+ M3015B（微流二氧化碳+有创血压+体温）\",\"bz\":\"网电源供电\"},{\"msbh\":\"④\",\"msmc\":\"工作模式\",\"msms\":\"主机（心电+心率+呼吸+无创血压+脉搏血氧饱和度+主流/侧流二氧化碳）+ M3014A（有创血压+体温）\",\"bz\":\"内部电源供电\"}],\"connectionGraph\":[{\"content\":\"模式①、③\",\"graphFileName\":\"model1.jpg\"},{\"content\":\"模式②、④\",\"graphFileName\":\"model2.jpg\"}],\"ypdlList\":[{\"dlxh\":\"1\",\"dlmc\":\"电源线\",\"dlfl\":\"\",\"dlcd\":\"2.2\",\"sfpb\":\"否\",\"bz\":\" \"},{\"dlxh\":\"2\",\"dlmc\":\"ECG附件\",\"dlfl\":\"M1631A\",\"dlcd\":\"1.6\",\"sfpb\":\"否\",\"bz\":\" \"},{\"dlxh\":\"2\",\"dlmc\":\"ECG附件\",\"dlfl\":\"M1671A\",\"dlcd\":\"1.0\",\"sfpb\":\"是\",\"bz\":\" \"},{\"dlxh\":\"2\",\"dlmc\":\"ECG附件\",\"dlfl\":\"M1611A\",\"dlcd\":\"1.0\",\"sfpb\":\"是\",\"bz\":\" \"},{\"dlxh\":\"2\",\"dlmc\":\"ECG附件\",\"dlfl\":\"M1675A\",\"dlcd\":\"1.0\",\"sfpb\":\"是\",\"bz\":\" \"},{\"dlxh\":\"2\",\"dlmc\":\"ECG附件\",\"dlfl\":\"M1610A\",\"dlcd\":\"1.0\",\"sfpb\":\"是\",\"bz\":\" \"},{\"dlxh\":\"2\",\"dlmc\":\"ECG附件\",\"dlfl\":\"M1678A\",\"dlcd\":\"1.0\",\"sfpb\":\"是\",\"bz\":\" \"},{\"dlxh\":\"2\",\"dlmc\":\"ECG附件\",\"dlfl\":\"M1672A\",\"dlcd\":\"1.0\",\"sfpb\":\"是\",\"bz\":\" \"},{\"dlxh\":\"2\",\"dlmc\":\"ECG附件\",\"dlfl\":\"M1673A\",\"dlcd\":\"1.0\",\"sfpb\":\"是\",\"bz\":\" \"},{\"dlxh\":\"2\",\"dlmc\":\"ECG附件\",\"dlfl\":\"M1609A\",\"dlcd\":\"1.0\",\"sfpb\":\"是\",\"bz\":\" \"},{\"dlxh\":\"2\",\"dlmc\":\"ECG附件\",\"dlfl\":\"M1674A\",\"dlcd\":\"1.0\",\"sfpb\":\"是\",\"bz\":\" \"},{\"dlxh\":\"2\",\"dlmc\":\"ECG附件\",\"dlfl\":\"M1603A\",\"dlcd\":\"1.0\",\"sfpb\":\"是\",\"bz\":\" \"},{\"dlxh\":\"2\",\"dlmc\":\"ECG附件\",\"dlfl\":\"M1619A\",\"dlcd\":\"0.7\",\"sfpb\":\"是\",\"bz\":\" \"},{\"dlxh\":\"2\",\"dlmc\":\"ECG附件\",\"dlfl\":\"M1605A\",\"dlcd\":\"1.0\",\"sfpb\":\"是\",\"bz\":\" \"},{\"dlxh\":\"2\",\"dlmc\":\"ECG附件\",\"dlfl\":\"M1615A\",\"dlcd\":\"1.0\",\"sfpb\":\"是\",\"bz\":\" \"},{\"dlxh\":\"2\",\"dlmc\":\"ECG附件\",\"dlfl\":\"M1968A\",\"dlcd\":\"1.0\",\"sfpb\":\"是\",\"bz\":\" \"},{\"dlxh\":\"2\",\"dlmc\":\"ECG附件\",\"dlfl\":\"M1625A\",\"dlcd\":\"1.0\",\"sfpb\":\"是\",\"bz\":\" \"},{\"dlxh\":\"2\",\"dlmc\":\"ECG附件\",\"dlfl\":\"M1639A\",\"dlcd\":\"1.3\",\"sfpb\":\"是\",\"bz\":\" \"},{\"dlxh\":\"2\",\"dlmc\":\"ECG附件\",\"dlfl\":\"M1974A\",\"dlcd\":\"1.6\",\"sfpb\":\"是\",\"bz\":\" \"},{\"dlxh\":\"3\",\"dlmc\":\"血氧饱和度传感器\",\"dlfl\":\"M1192A\",\"dlcd\":\"1.5\",\"sfpb\":\"是\",\"bz\":\" \"},{\"dlxh\":\"3\",\"dlmc\":\"血氧饱和度传感器\",\"dlfl\":\"M1193A\",\"dlcd\":\"1.5\",\"sfpb\":\"是\",\"bz\":\" \"},{\"dlxh\":\"3\",\"dlmc\":\"血氧饱和度传感器\",\"dlfl\":\"M1194A\",\"dlcd\":\"1.5\",\"sfpb\":\"是\",\"bz\":\" \"},{\"dlxh\":\"3\",\"dlmc\":\"血氧饱和度传感器\",\"dlfl\":\"M1195A\",\"dlcd\":\"1.5\",\"sfpb\":\"是\",\"bz\":\" \"},{\"dlxh\":\"3\",\"dlmc\":\"血氧饱和度传感器\",\"dlfl\":\"M1196A\",\"dlcd\":\"1.5\",\"sfpb\":\"是\",\"bz\":\" \"},{\"dlxh\":\"3\",\"dlmc\":\"血氧饱和度传感器\",\"dlfl\":\"M1196A\",\"dlcd\":\"1.5\",\"sfpb\":\"是\",\"bz\":\" \"},{\"dlxh\":\"4\",\"dlmc\":\"温度适配线缆\",\"dlfl\":\"21082A\",\"dlcd\":\"3.0\",\"sfpb\":\"是\",\"bz\":\" \"},{\"dlxh\":\"4\",\"dlmc\":\"温度适配线缆\",\"dlfl\":\"21082B\",\"dlcd\":\"1.5\",\"sfpb\":\"是\",\"bz\":\" \"},{\"dlxh\":\"5\",\"dlmc\":\"二氧化碳附件\",\"dlfl\":\"M2501A\",\"dlcd\":\"3.0\",\"sfpb\":\"是\",\"bz\":\" \"},{\"dlxh\":\"5\",\"dlmc\":\"二氧化碳附件\",\"dlfl\":\"M2741A\",\"dlcd\":\"0.5\",\"sfpb\":\"是\",\"bz\":\" \"}],\"cssbList\":[{\"cssbxh\":\"1\",\"cssbbhxlh\":\"2-FW-11\",\"cssbmc\":\"测试接收机\",\"cssbzzs\":\"R&S\",\"cssbxhgg\":\"ESH2-Z5\",\"cssbxcjzrq\":\"2019.5.15\",\"cssbbz\":\" \"},{\"cssbxh\":\"2\",\"cssbbhxlh\":\"2-FW-12\",\"cssbmc\":\"人工电源网络\",\"cssbzzs\":\"R&S\",\"cssbxhgg\":\"ESCI\",\"cssbxcjzrq\":\"2019.5.15\",\"cssbbz\":\" \"},{\"cssbxh\":\"3\",\"cssbbhxlh\":\"2-FW-103\",\"cssbmc\":\"屏蔽室1\",\"cssbzzs\":\"ETS·LINDGREN\",\"cssbxhgg\":\"SR1\",\"cssbxcjzrq\":\"2019.4.14\",\"cssbbz\":\" \"},{\"cssbxh\":\"4\",\"cssbbhxlh\":\"2-FW-93\",\"cssbmc\":\"测试接收机\",\"cssbzzs\":\"R&S\",\"cssbxhgg\":\"ESU26\",\"cssbxcjzrq\":\"2019.5.15\",\"cssbbz\":\" \"},{\"cssbxh\":\"5\",\"cssbbhxlh\":\"2-FW-101\",\"cssbmc\":\"双锥复合对数周期天线\",\"cssbzzs\":\"SCHWARZBECK\",\"cssbxhgg\":\"VULB9163\",\"cssbxcjzrq\":\"2020.2.13\",\"cssbbz\":\" \"},{\"cssbxh\":\"6\",\"cssbbhxlh\":\"2-FW-102\",\"cssbmc\":\"10米法电波暗室\",\"cssbzzs\":\"ETS·LINDGREN\",\"cssbxhgg\":\"FACT10\",\"cssbxcjzrq\":\"2019.4.14\",\"cssbbz\":\" \"},{\"cssbxh\":\"7\",\"cssbbhxlh\":\"2-FW-102\",\"cssbmc\":\"控制室\",\"cssbzzs\":\"ETS·LINDGREN\",\"cssbxhgg\":\"CR\",\"cssbxcjzrq\":\"2019.4.14\",\"cssbbz\":\" \"},{\"cssbxh\":\"8\",\"cssbbhxlh\":\"2-FW-163\",\"cssbmc\":\"静电放电器\",\"cssbzzs\":\"EM TEST\",\"cssbxhgg\":\"dito\",\"cssbxcjzrq\":\"2019.1.22\",\"cssbbz\":\" \"},{\"cssbxh\":\"9\",\"cssbbhxlh\":\"2-FW-106\",\"cssbmc\":\"屏蔽室3\",\"cssbzzs\":\"ETS·LINDGREN\",\"cssbxhgg\":\"SR3\",\"cssbxcjzrq\":\"2019.4.14\",\"cssbbz\":\" \"},{\"cssbxh\":\"10\",\"cssbbhxlh\":\"2-FW-30\",\"cssbmc\":\"功率放大器\",\"cssbzzs\":\"BONN\",\"cssbxhgg\":\"BLWA0830-160/100/40D\",\"cssbxcjzrq\":\"\",\"cssbbz\":\" \"},{\"cssbxh\":\"11\",\"cssbbhxlh\":\"2-FW-34\",\"cssbmc\":\"场强表\",\"cssbzzs\":\"AR\",\"cssbxhgg\":\"FL7006/Kit M1\",\"cssbxcjzrq\":\"2019.9.11\",\"cssbbz\":\" \"},{\"cssbxh\":\"12\",\"cssbbhxlh\":\"2-FW-100\",\"cssbmc\":\"信号发生器\",\"cssbzzs\":\"R&S\",\"cssbxhgg\":\"SMB100A-B106\",\"cssbxcjzrq\":\"2019.5.15\",\"cssbbz\":\" \"}],\"fzsbList\":[{\"fzsbxh\":\"1\",\"fzsbbhxlh\":\"1-EI-20\",\"fzsbmc\":\"高频电刀分析仪\",\"fzsbsccj\":\"FLUKE\",\"fzsbxhgg\":\"QA-ESⅡ\",\"fzsbxcjzrq\":\"2018.9.8\",\"fzsbbz\":\" \"},{\"fzsbxh\":\"2\",\"fzsbbhxlh\":\"1-EV-236\",\"fzsbmc\":\"生命体征模拟仪\",\"fzsbsccj\":\"FLUKE\",\"fzsbxhgg\":\"Prosim 8\",\"fzsbxcjzrq\":\"2018.9.6\",\"fzsbbz\":\" \"}],\"experiment\":[{\"name\":\"传导发射实验\",\"syjg\":\"符合\",\"jyrq\":\"2020-01-01\",\"wd\":\"100\",\"xdsd\":\"65\",\"dqyl\":\"100\",\"sysj\":[{\"sygdy\":\"AC220V 50Hz\",\"syplfw\":\"0.15MHz~30MHz\",\"ypyxms\":\"① \",\"rtf\":[{\"name\":\"ZC2018-128  生物安全柜 模式1 CE L.rtf\"},{\"name\":\"ZC2018-128  生物安全柜 模式1 CE N.rtf\"}]},{\"sygdy\":\"AC220V 50Hz\",\"syplfw\":\"0.15MHz~30MHz\",\"ypyxms\":\"② \",\"rtf\":[{\"name\":\"ZC2018-128  生物安全柜 模式1 CE L - 副本.rtf\"},{\"name\":\"ZC2018-128  生物安全柜 模式1 CE N - 副本.rtf\"}]}],\"syljt\":[{\"name\":\"image1.jpg\",content:\"\"}],\"sybzt\":[{\"name\":\"image2.jpg\",content:\"①③\"},{\"name\":\"image2.jpg\",content:\"②④\"}]},{\"name\":\"辐射发射试验\",\"syjg\":\"符合\",\"jyrq\":\"2020-01-01\",\"wd\":\"100\",\"xdsd\":\"65\",\"dqyl\":\"100\",\"sysj\":[{\"sygdy\":\"AC220V 50Hz\",\"syplfw\":\"0.15MHz~30MHz\",\"ypyxms\":\"① \",\"rtf\":[{\"name\":\"QW2018-2065 冷冻射频肿瘤治疗系统   RE.rtf\"}]},{\"sygdy\":\"AC220V 50Hz\",\"syplfw\":\"0.15MHz~30MHz\",\"ypyxms\":\"② \",\"rtf\":[{\"name\":\"QW2018-2065 冷冻射频肿瘤治疗系统   RE.rtf\"}]}],\"syljt\":[{\"name\":\"reljt.jpg\",content:\"\"}],\"sybzt\":[{\"name\":\"re1.jpg\",content:\"①③\"},{\"name\":\"re2.jpg\",content:\"②④\"}]}]}";

        [HttpGet]
        public IEnumerable<string> Get()
        {
            return new string[] { "Emc", "生成报告" };
        }

        /// <summary>
        /// 上传文件
        /// </summary>
        [HttpPost]
        public ReportResult<string> UploadFiles()
        {
            ReportResult<string> result = new ReportResult<string>();
            HttpFileCollection filelist = HttpContext.Current.Request.Files;
            string currRoot = AppDomain.CurrentDomain.BaseDirectory;
            if (filelist != null && filelist.Count > 0)
            {
                for (int i = 0; i < filelist.Count; i++)
                {
                    try
                    {
                        HttpPostedFile file = filelist[i];
                        string filename = file.FileName;
                        if (filename.Equals(""))
                        {
                            MyTools.ErrorLog.Error("上传失败:上传的文件信息不存在！");
                            result = SetReportResult<string>("下载失败:上传的文件信息不存在！", false, "");
                        }
                        string extendName = MyTools.FilterExtendName(filename);
                        string filePath = currRoot + "\\Files\\Upload\\";
                        string forceName = "";
                        //判断上传的文件
                        switch (extendName)
                        {
                            case ".jpg":
                            case ".png":
                                filePath = currRoot + "\\Files\\Upload\\Image\\";
                                forceName = "image";
                                break;
                            case ".rtf":
                                filePath = currRoot + "\\Files\\Upload\\Rtf\\";
                                forceName = "rtf";
                                break;
                            default:
                                filePath = currRoot + "\\Files\\Upload\\";
                                forceName = "upload";
                                break;
                        }
                        string templateFileName = forceName + DateTime.Now.ToString("yyyyMMddHHmmssfff") + extendName;

                        DirectoryInfo di = new DirectoryInfo(filePath);
                        if (!di.Exists) { di.Create(); }

                        file.SaveAs(filePath + templateFileName);
                        MyTools.InfoLog.Info(result);
                        result = SetReportResult<string>(string.Format("上传成功:{0}", filename), true, templateFileName);
                    }
                    catch (Exception ex)
                    {
                        MyTools.ErrorLog.Error(ex.Message, ex);
                        result = SetReportResult<string>(string.Format("上传文件写入失败：{0}", ex.Message), false, "");
                    }
                }
            }
            else
            {
                MyTools.ErrorLog.Error("上传失败:上传的文件信息不存在！");
                result = SetReportResult<string>("下载失败:上传的文件信息不存在！", false, "");
            }

            return result;
        }

        /// <summary>
        /// 下载文件
        /// </summary>
        [HttpPost]
        public IHttpActionResult DownloadFiles(ReportParams para)
        {
            string fileName = para.FileName;
            string currRoot = AppDomain.CurrentDomain.BaseDirectory;
            try
            {
                var browser = String.Empty;
                if (HttpContext.Current.Request.UserAgent != null)
                {
                    browser = HttpContext.Current.Request.UserAgent.ToUpper();
                }
                string extendName = MyTools.FilterExtendName(fileName);
                string fileFullName = "";
                //判断上传的文件
                switch (extendName)
                {
                    case ".jpg":
                    case ".png":
                        fileFullName = GetImagePath(fileName);
                        break;
                    case ".rtf":
                        fileFullName = GetRtfPath(fileName);
                        break;
                    default:
                        fileFullName = GetWordPath(fileName);
                        break;
                }
                HttpResponseMessage httpResponseMessage = new HttpResponseMessage(HttpStatusCode.OK);
                FileStream fileStream = File.OpenRead(fileFullName);
                httpResponseMessage.Content = new StreamContent(fileStream);
                httpResponseMessage.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
                httpResponseMessage.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
                {
                    FileName =
                        browser.Contains("FIREFOX")
                            ? Path.GetFileName(fileFullName)
                            : HttpUtility.UrlEncode(Path.GetFileName(fileFullName))
                    //FileName = HttpUtility.UrlEncode(Path.GetFileName(filePath))
                };
                MyTools.InfoLog.Info("下载成功:" + fileName);
                return ResponseMessage(httpResponseMessage);
            }
            catch (Exception ex)
            {
                MyTools.ErrorLog.Error(ex.Message, ex);
                throw ex;
            }
        }


        [HttpGet]
        public string Get2()
        {
            Stopwatch sw = new Stopwatch();

            sw.Start();
            MyTools.KillWordProcess();
            string result = JsonToWord(jsonStr);
            //string result = "";
            sw.Stop();
            double time1 = (double)sw.ElapsedMilliseconds / 1000;
            return result + ":" + time1.ToString();
        }

        [HttpGet]
        public IHttpActionResult CreateReportAndDownLoad()
        {
            Stopwatch sw = new Stopwatch();
            sw.Start();
            MyTools.KillWordProcess();
            string result = JsonToWord(jsonStr);
            result = GetWordPath(result);
            //string result = "";
            sw.Stop();
            double time1 = (double)sw.ElapsedMilliseconds / 1000;
            return DownloadFiles(result);
        }

        [HttpPost]
        public string CreateReport(ReportParams para)
        {
            string jsonStr = para.JsonStr;
            string result = "创建成功";
            Stopwatch sw = new Stopwatch();

            sw.Start();
            MyTools.KillWordProcess();
            try
            {
                result = JsonToWord(jsonStr);
            }
            catch (Exception ex)
            {

                throw ex;
            }
            sw.Stop();
            double time1 = (double)sw.ElapsedMilliseconds / 1000;
            return "创建成功" + result + ":" + time1.ToString();
        }



        #region 私有方法

        private ReportResult<T> SetReportResult<T>(string message, bool submitResult, object content)
        {
            Type type = content.GetType();
            ReportResult<T> reportResult = new ReportResult<T>();
            reportResult.Message = message;
            reportResult.SumbitResult = submitResult;
            reportResult.Content = content;
            return reportResult;
        }

        private string GetImagePath(string fileName)
        {
            string currRoot = AppDomain.CurrentDomain.BaseDirectory;
            string imageFullFileName = string.Format(@"{0}\Files\Upload\Image\{1}", currRoot, fileName);
            return imageFullFileName;
        }

        private string GetRtfPath(string fileName)
        {
            string currRoot = AppDomain.CurrentDomain.BaseDirectory;
            string imageFullFileName = string.Format(@"{0}\Files\Upload\Rtf\{1}", currRoot, fileName);
            return imageFullFileName;
        }

        private string GetWordPath(string fileName)
        {
            string currRoot = AppDomain.CurrentDomain.BaseDirectory;
            string imageFullFileName = string.Format(@"{0}\Files\OutPut\{1}", currRoot, fileName);
            return imageFullFileName;
        }

        private string GetTemplatePath(string fileName)
        {
            string currRoot = AppDomain.CurrentDomain.BaseDirectory;
            string imageFullFileName = string.Format(@"{0}\Files\ExperimentTemplate\{1}", currRoot, fileName);
            return imageFullFileName;
        }


        #region 生成报表方法
        private string JsonToWord(string jsonStr)
        {
            //解析json字符串
            JObject mainObj = (JObject)JsonConvert.DeserializeObject(jsonStr);
            string currRoot = AppDomain.CurrentDomain.BaseDirectory;
            string outfileName = string.Format("report{0}.docx", MyTools.GetTimestamp(DateTime.Now));//输出文件名称
            string outfilePth = string.Format(@"{0}\Files\OutPut\{1}", currRoot, outfileName);//输出文件路径
            string filePath = string.Format(@"{0}\Files\{1}", currRoot, ConfigurationManager.AppSettings["TemplateName"].ToString());//模板文件

            string middleDir = currRoot + "\\Files\\TemplateMiddleware\\" + DateTime.Now.ToString("yyyyMMddhhmmss");
            filePath = CreateTemplateMiddle(middleDir,"template", filePath);
            //生成报告
            using (WordUtil wordUtil = new WordUtil(outfilePth, filePath))
            {
                ////首页内容 object
                ////受检样品描述 object
                //JObject firstPage = (JObject)mainObj["firstPage"];
                //result = InsertContentToWord(wordUtil, firstPage);
                //if (!result.Equals("保存成功"))
                //{
                //    return result;
                //}

                ////////样品构成 list
                //JArray ypgcList = (JArray)mainObj["ypgcList"];
                //result = InsertListIntoTable(wordUtil, ypgcList, 2, "ypgclist");
                //if (!result.Equals("保存成功"))
                //{
                //    return result;
                //}

                //////样品连接图 图片
                //JArray graphList = (JArray)mainObj["connectionGraph"];
                //InsertImageToWord(wordUtil, graphList, "connectionGraph");

                //////样品运行模式 list
                //JArray ypyxList = (JArray)mainObj["ypyxList"];
                //result = InsertListIntoTable(wordUtil, ypyxList, 1, "ypyxlist", false);
                //if (!result.Equals("保存成功"))
                //{
                //    return result;
                //}

                //////样品电缆 list
                //JArray ypdlList = (JArray)mainObj["ypdlList"];
                //result = InsertListIntoTable(wordUtil, ypdlList, 2, "ypdllist");
                //if (!result.Equals("保存成功"))
                //{
                //    return result;
                //}

                ////测试设备list
                //JArray cssbList = (JArray)mainObj["cssbList"];
                //result = InsertListIntoTable(wordUtil, cssbList, 1, "cssblist");
                //if (!result.Equals("保存成功"))
                //{
                //    return result;
                //}

                ////辅助设备 list
                //JArray fzsbList = (JArray)mainObj["fzsbList"];
                //result = InsertListIntoTable(wordUtil, fzsbList, 1, "fzsblist");
                //if (!result.Equals("保存成功"))
                //{
                //    return result;
                //}

                //实验结果概述
                //JArray experimentalResult = (JArray)mainObj["experimentalResult"];
                //result = InsertListIntoTableByTitle(wordUtil, experimentalResult, "experimentalResult");

                //实验数据
                JArray experiment = (JArray)mainObj["experiment"];
                string newBookmark = "experiment";
                foreach (JObject item in experiment)
                {
                    if (item["name"].ToString().Equals("传导发射实验"))
                        newBookmark = SetConductedEmission(wordUtil, item, newBookmark, "CE",middleDir);
                    else if (item["name"].ToString().Equals("辐射发射试验"))
                        newBookmark = SetConductedEmission(wordUtil, item, newBookmark, "RE", middleDir);
                }
            }
            //删除中间件文件夹
            //this.DelectDir(middleDir);

            return outfileName;
        }

        //设置首页内容
        private string InsertContentToWord(WordUtil wordUtil, JObject jo1)
        {
            foreach (var item in jo1)
            {
                wordUtil.InsertContentToWord(item.Value.ToString(), item.Key);
            }
            return "保存成功";
        }

        private string InsertListIntoTable(WordUtil wordUtil, JArray array, int mergeColumn, string bookmark, bool isNeedNumber = true)
        {
            List<string> list = JarrayToList(array);

            string result = wordUtil.InsertListToTable(list, bookmark, mergeColumn, isNeedNumber);

            return result;
        }
        ////实验结果概述
        //private string InsertListIntoTableByTitle(WordUtil wordUtil, JArray array, string bookmark)
        //{
        //    Dictionary<string, List<string>> dic = JarrayToDic(array);
        //    string result = wordUtil.InsertListIntoTableByTitle(dic, bookmark,false);

        //    return result;
        //}


        private void InsertImageToWord(WordUtil wordUtil, JArray array, string bookmark)
        {
            List<string> list = new List<string>();
            foreach (JObject item in array)
            {
                string jTemp = "";
                int iTemp = 0;
                foreach (var item2 in item)
                {
                    iTemp++;
                    string tempValue = item2.Value.ToString();
                    if (iTemp == 2)
                    {
                        tempValue = GetImagePath(tempValue);
                    }
                    if (iTemp != item.Count)
                        jTemp += (tempValue + ",");
                    else
                        jTemp += tempValue;
                }
                list.Add(jTemp);
            }

            wordUtil.InsertImageToWord(list, bookmark);

        }

        private List<string> JarrayToList(JArray array)
        {
            List<string> list = new List<string>();

            foreach (JObject item in array)
            {
                string jTemp = "";
                int iTemp = 0;
                foreach (var item2 in item)
                {
                    iTemp++;
                    if (iTemp != item.Count)
                        jTemp += (item2.Value + ",");
                    else
                        jTemp += item2.Value;
                }
                list.Add(jTemp);
            }

            return list;
        }

        private Dictionary<string, List<string>> JarrayToDic(JArray array)
        {
            Dictionary<string, List<string>> dic = new Dictionary<string, List<string>>();
            foreach (JObject main in array)
            {
                string title = main["title"].ToString();
                List<string> list = new List<string>();
                JArray resultArray = (JArray)main["resultList"];
                foreach (JObject item in resultArray)
                {
                    string jTemp = "";
                    int iTemp = 0;
                    foreach (var item2 in item)
                    {
                        iTemp++;
                        if (iTemp != item.Count)
                            jTemp += (item2.Value + ",");
                        else
                            jTemp += item2.Value;
                    }
                    list.Add(jTemp);
                }
                dic.Add(title, list);

            }
            return dic;
        }
        #endregion

        #region 实验数据
        /// <summary>
        /// 传导发射实验
        /// </summary>
        /// <returns></returns>
        private string SetConductedEmission(WordUtil wordUtil, JObject jObject, string bookmark, string rtfType,string middleDir)
        {
            string templateName = jObject["name"].ToString();
            string templateFullPath = CreateTemplateMiddle(middleDir,"experiment",GetTemplatePath(templateName + ".docx"));
            string sysjTemplateFilePath = CreateTemplateMiddle(middleDir,"sysj",GetTemplatePath("RTFTemplate.docx"));

            foreach (var item in jObject)
            {
                if (!item.Key.Equals("sysj") && !item.Key.Equals("name") && !item.Key.Equals("syljt") && !item.Key.Equals("sybzt"))
                    wordUtil.InsertContentInBookmark(templateFullPath, item.Value.ToString(), item.Key, false);
            }

            JArray sysj = (JArray)jObject["sysj"];
            RtfTableInfo rtfTableInfo = MyTools.RtfTableInfos.Where(p => p.RtfType == rtfType).FirstOrDefault();

            int startIndex = rtfTableInfo.StartIndex;
            Dictionary<int, string> dic = rtfTableInfo.ColumnInfoDic;
            string rtfbookmark = rtfTableInfo.Bookmark;

            RtfPictureInfo rtfPictureInfo = MyTools.RtfPictureInfos.Where(p => p.RtfType == rtfType).FirstOrDefault();
            int imageStartIndex = rtfPictureInfo.StartIndex;
            string imageBookmark = rtfPictureInfo.Bookmark;


            int i = 0;
            foreach (JObject item in sysj)
            {
                //插入实验数据信息 (画表格)

                List<string> contentList = new List<string>();
                contentList.Add("试验供电电源："+item["sygdy"].ToString());
                contentList.Add("试验频率范围："+item["syplfw"].ToString());
                contentList.Add("样品运行模式："+item["ypyxms"].ToString());

                wordUtil.CreateTableToWord(sysjTemplateFilePath, contentList, "sysj", false, i != 0);

                JArray rtf = (JArray)item["rtf"];
                int rtfCount = rtf.Count;
                int j = 0;

                foreach (JObject rtfObj in (JArray)item["rtf"])
                {
                    //需要画表格和插入rtf内容
                    wordUtil.CopyOtherFileTableForColByTableIndex(sysjTemplateFilePath, GetRtfPath(rtfObj["name"].ToString()), startIndex, dic, rtfbookmark, false, true, false);

                    wordUtil.CopyOtherFilePictureToWord(sysjTemplateFilePath, GetRtfPath(rtfObj["name"].ToString()), imageStartIndex, imageBookmark, false, true, j == rtfCount - 1);
                    j++;
                }

                //在最后添加分页符
                //wordUtil.InsertBreakForPage(sysjTemplateFilePath, false);
                i++;
            }

            wordUtil.CopyOtherFileContentToWord(sysjTemplateFilePath, templateFullPath, "sysj", true);

            //插入图片
            JArray syljt = (JArray)jObject["syljt"];
            List<string> list = new List<string>();

            foreach (JObject item in syljt)
            {
                list.Add(GetImagePath(item["name"].ToString()) + "," + item["content"].ToString());
            }

            wordUtil.InsertImageToTemplate(templateFullPath, list, "syljt", false);

            JArray sybzt = (JArray)jObject["sybzt"];
            list = new List<string>();
            foreach (JObject item in sybzt)
            {
                list.Add(GetImagePath(item["name"].ToString()) + "," + item["content"].ToString());
            }
            wordUtil.InsertImageToTemplate(templateFullPath, list, "sybzt", false);

            string result = wordUtil.CopyOtherFileContentToWordReturnBookmark(templateFullPath, bookmark);


            return result;
        }
        #endregion

        //生成报告直接下载
        private IHttpActionResult DownloadFiles(string fileFullName)
        {
            try
            {
                var browser = HttpContext.Current.Request.UserAgent.ToUpper();
                HttpResponseMessage httpResponseMessage = new HttpResponseMessage(HttpStatusCode.OK);
                FileStream fileStream = File.OpenRead(fileFullName);
                httpResponseMessage.Content = new StreamContent(fileStream);
                httpResponseMessage.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
                httpResponseMessage.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
                {
                    FileName =
                        browser.Contains("FIREFOX")
                            ? Path.GetFileName(fileFullName)
                            : HttpUtility.UrlEncode(Path.GetFileName(fileFullName))
                    //FileName = HttpUtility.UrlEncode(Path.GetFileName(filePath))
                };
                MyTools.InfoLog.Info("下载成功" + fileFullName);
                return ResponseMessage(httpResponseMessage);
            }
            catch (Exception ex)
            {
                MyTools.ErrorLog.Error(ex.Message, ex);
                throw ex;
            }
        }

        //创建模板中间件
        private string CreateTemplateMiddle(string dir,string template, string filePath)
        {

            string dateStr = DateTime.Now.ToString("yyyyMMddhhmmss");
            string fileName = template + dateStr + ".docx";
            DirectoryInfo di = new DirectoryInfo(dir);
            if (!di.Exists) { di.Create(); }

            string htmlpath = dir+"\\" + fileName;
            FileInfo file = new FileInfo(filePath);
            if (File.Exists(filePath))
            {
                file.CopyTo(htmlpath);
                return htmlpath;
            }
            else
            {
                return "模板不存在";
            }

        }

        public void DelectDir(string srcPath)
        {
            try
            {
                DirectoryInfo dir = new DirectoryInfo(srcPath);
                FileSystemInfo[] fileinfo = dir.GetFileSystemInfos();  //返回目录中所有文件和子目录
                foreach (FileSystemInfo i in fileinfo)
                {
                    if (i is DirectoryInfo)            //判断是否文件夹
                    {
                        DirectoryInfo subdir = new DirectoryInfo(i.FullName);
                        subdir.Delete(true);          //删除子目录和文件
                    }
                    else
                    {
                        //如果 使用了 streamreader 在删除前 必须先关闭流 ，否则无法删除 sr.close();
                        File.Delete(i.FullName);      //删除指定文件
                    }
                }
            }
            catch (Exception e)
            {
                throw;
            }
        }

        #region 测试
        private string JsonStrToJObject()
        {
            string jsonStr = "{\"FirstPage\":[{\"wtf\":\"飞利浦(中国)投资有限公司1\",\"ypmc\":\"病人监护仪1\",\"xhgg\":\"M8102A1\",\"jylb\":\"委托检验1\"},{\"wtf\":\"飞利浦(中国)投资有限公司2\",\"ypmc\":\"病人监护仪2111\",\"xhgg\":\"M8102A211\",\"jylb\":\"委托检验2111\"},{\"wtf\":\"飞利浦(中国)投资有限公司2\",\"ypmc\":\"病人监护仪2111\",\"xhgg\":\"M8102A211\",\"jylb\":\"委托检验2111\"},{\"wtf\":\"飞利浦(中国)投资有限公司1\",\"ypmc\":\"病人监护仪1\",\"xhgg\":\"M8102A1\",\"jylb\":\"委托检验1\"}]}";
            JObject jo1 = (JObject)JsonConvert.DeserializeObject(jsonStr);
            JArray firstPage = JArray.Parse(jo1["FirstPage"].ToString());
            //首页内容
            foreach (JObject item in firstPage)
            {
                //wordUtil.InsertContentInBookmark(item.Value.ToString(), item.Name);
            }

            string currRoot = AppDomain.CurrentDomain.BaseDirectory;
            string currDateStr = MyTools.GetTimestamp(DateTime.Now);
            string outfilePth = string.Format(@"{0}\Files\OutPut\output{1}.docx", currRoot, currDateStr);
            string filePath = string.Format(@"{0}\Files\{1}", currRoot, "国医检(磁)字QW2018第698号模板改造.docx");
            using (WordUtil wordUtil = new WordUtil(outfilePth, filePath))
            {
                //首页内容
                foreach (JArray item in firstPage)
                {
                    //wordUtil.InsertContentInBookmark(item.Value.ToString(), item.Name);
                }
            }
            return "转化成功";
        }

        private string InsertListIntoTable()
        {
            string jsonStr = "{\"FirstPage\":[{\"wtf\":\"飞利浦(中国)投资有限公司1\",\"ypmc\":\"病人监护仪1\",\"xhgg\":\"M8102A1\",\"jylb\":\"委托检验1\",\"t1\":\"t1\",\"t2\":\"t2\"},{\"wtf\":\"飞利浦(中国)投资有限公司2\",\"ypmc\":\"病人监护仪2111\",\"xhgg\":\"M8102A211\",\"jylb\":\"委托检验2111\",\"t1\":\"t1\",\"t2\":\"t2\"},{\"wtf\":\"飞利浦(中国)投资有限公司2\",\"ypmc\":\"病人监护仪2111\",\"xhgg\":\"M8102A211\",\"jylb\":\"委托检验2111\",\"t1\":\"t1\",\"t2\":\"t2\"},{\"wtf\":\"飞利浦(中国)投资有限公司1\",\"ypmc\":\"病人监护仪2111\",\"xhgg\":\"M8102A1\",\"jylb\":\"委托检验1\",\"t1\":\"t1\",\"t2\":\"t2\"},{\"wtf\":\"飞利浦(中国)投资有限公司1\",\"ypmc\":\"病人监护仪1\",\"xhgg\":\"M8102A1\",\"jylb\":\"委托检验1\",\"t1\":\"t1\",\"t2\":\"t2\"},{\"wtf\":\"飞利浦(中国)投资有限公司2\",\"ypmc\":\"病人监护仪2111\",\"xhgg\":\"M8102A211\",\"jylb\":\"委托检验2111\",\"t1\":\"t1\",\"t2\":\"t2\"},{\"wtf\":\"飞利浦(中国)投资有限公司2\",\"ypmc\":\"病人监护仪2111\",\"xhgg\":\"M8102A211\",\"jylb\":\"委托检验2111\",\"t1\":\"t1\",\"t2\":\"t2\"},{\"wtf\":\"飞利浦(中国)投资有限公司1\",\"ypmc\":\"病人监护仪1\",\"xhgg\":\"M8102A1\",\"jylb\":\"委托检验1\",\"t1\":\"t1\",\"t2\":\"t2\"},{\"wtf\":\"飞利浦(中国)投资有限公司1\",\"ypmc\":\"病人监护仪1\",\"xhgg\":\"M8102A1\",\"jylb\":\"委托检验1\",\"t1\":\"t1\",\"t2\":\"t2\"},{\"wtf\":\"飞利浦(中国)投资有限公司2\",\"ypmc\":\"病人监护仪2111\",\"xhgg\":\"M8102A211\",\"jylb\":\"委托检验2111\",\"t1\":\"t1\",\"t2\":\"t2\"},{\"wtf\":\"飞利浦(中国)投资有限公司2\",\"ypmc\":\"病人监护仪2111\",\"xhgg\":\"M8102A211\",\"jylb\":\"委托检验2111\",\"t1\":\"t1\",\"t2\":\"t2\"},{\"wtf\":\"飞利浦(中国)投资有限公司1\",\"ypmc\":\"病人监护仪1\",\"xhgg\":\"M8102A1\",\"jylb\":\"委托检验1\",\"t1\":\"t1\",\"t2\":\"t2\"}]}";
            JObject jo1 = (JObject)JsonConvert.DeserializeObject(jsonStr);
            JArray firstPage = (JArray)jo1["FirstPage"];

            List<string> list = new List<string>();

            foreach (JObject item in firstPage)
            {
                string jTemp = "";
                int iTemp = 0;
                foreach (var item2 in item)
                {
                    iTemp++;
                    if (iTemp != item.Count)
                        jTemp += (item2.Value + ",");
                    else
                        jTemp += item2.Value;
                }
                list.Add(jTemp);
            }

            string currRoot = AppDomain.CurrentDomain.BaseDirectory;
            string currDateStr = MyTools.GetTimestamp(DateTime.Now);
            string outfilePth = string.Format(@"{0}\Files\OutPut\output{1}.docx", currRoot, currDateStr);
            string filePath = string.Format(@"{0}\Files\{1}", currRoot, "TestListToTable.docx");
            using (WordUtil wordUtil = new WordUtil(outfilePth, filePath))
            {
                wordUtil.InsertListToTable(list, "bookmark41", 2);
            }

            return "保存成功";
        }

        private string InsertRtfIntoReport(string fileName, string htmlstr)
        {
            string currRoot = AppDomain.CurrentDomain.BaseDirectory;
            string currDateStr = MyTools.GetTimestamp(DateTime.Now);
            string outfilePth = string.Format(@"{0}\Files\OutPut\output{1}.docx", currRoot, currDateStr);
            string filePath = string.Format(@"{0}\Files\{1}", currRoot, fileName);

            MyTools.KillWordProcess();

            //string htmlfilePath = string.Format(@"{0}\Files\Html\{1}", currRoot, "testhtml.html");
            string htmlfilePath = CreateHtmlFile(htmlstr);

            string result = "创建成功";

            using (WordUtil wordUtil = new WordUtil(outfilePth, filePath))
            {

                wordUtil.CopyOtherFileContentToWord(htmlfilePath, "bookmark1");

                //获取文件中的table插入到当前文件
                string rtfFileName = "ZC2018-128  生物安全柜 模式1 CE L.Rtf";
                string rtfFullName = string.Format(@"{0}\Files\检测设备产出文档\{1}", currRoot, rtfFileName);

                RtfTableInfo rtfTableInfo = MyTools.RtfTableInfos.Where(p => rtfFullName.Contains(p.RtfType)).FirstOrDefault();

                if (rtfTableInfo == null)
                {
                    throw new Exception("rtf配置文件未找到(" + rtfFullName + ")相关文件信息");
                }

                int startIndex = rtfTableInfo.StartIndex;
                Dictionary<int, string> dic = rtfTableInfo.ColumnInfoDic;
                string bookmark = rtfTableInfo.Bookmark;

                wordUtil.CopyOtherFileTableForColByTableIndex(rtfFullName, startIndex, dic, bookmark, false);

                RtfPictureInfo rtfPictureInfo = MyTools.RtfPictureInfos.Where(p => rtfFullName.Contains(p.RtfType)).FirstOrDefault();
                startIndex = rtfPictureInfo.StartIndex;
                bookmark = rtfPictureInfo.Bookmark;

                wordUtil.CopyOtherFilePictureToWord(rtfFullName, startIndex, bookmark, false, true);
            }

            MyTools.InfoLog.Info(result);
            MyTools.ErrorLog.Error("创建失败");
            return result;
        }

        private string CreateHtmlFile(string htmlStr)
        {
            string dateStr = DateTime.Now.ToString("yyyyMMddhhmmss");
            string currRoot = AppDomain.CurrentDomain.BaseDirectory;
            string htmlpath = currRoot + "Files\\Html\\reportHtml" + dateStr + ".html";
            FileStream fs = new FileStream(htmlpath, FileMode.Create);
            StreamWriter sw = new StreamWriter(fs);
            sw.Write(htmlStr);
            sw.Close();
            sw.Dispose();
            fs.Close();
            fs.Dispose();
            return htmlpath;
        }
        #endregion
        #endregion

    }
}
