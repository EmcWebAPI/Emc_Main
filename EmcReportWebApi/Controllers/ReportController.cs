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
        private string jsonStr = "{\"scbWord\":\"0505 电磁兼容资料审查表0802.docx\",\"firstPage\":{\"main_wtf\":\"飞利浦(中国)投资有限公司1\",\"main_ypmc\":\"病人监护仪1\",\"main_xhgg\":\"M8102A1\",\"main_jylb\":\"委托检验1\",\"ypmc\":\"病人监护仪\",\"sb\":\"\",\"wtf\":\"飞利浦（中国）投资有限公司\",\"wtfdz\":\"上海市静安区灵石路718号A幢\",\"scdw\":\"PhilipsMedizinSystemeBoeblingenGmbH\",\"sjdw\":\"PhilipsMedizinSystemeBoeblingenGmbH\",\"cydw\":\"\",\"cydd\":\"\",\"cyrq\":\"\",\"dyrq\":\"2018年5月8日\",\"jyxm\":\"YY0505全项目、YY0601中36、YY0667中36、YY0668中36、YY0783中36、YY0784中36\",\"jyyj\":\"YY0505-2012《医用电气设备第1-2部分：安全通用要求并列标准电磁兼容要求和试验》、YY0601-2009《医用电气设备呼吸气体监护仪的基本安全和主要性能专用要求》、YY0667-2008《医用电气设备第2-30部分：自动循环无创血压监护设备的安全和基本性能专用要求》、YY0668-2008《医用电气设备第2-49部分：多参数患者监护设备安全专用要求》、YY0783-2010《医用电气设备第2-34部分：有创血压监测设备的安全和基本性能专用要求》、YY0784-2010《医用电气设备医用脉搏血氧仪设备基本安全和主要性能专用要求》\",\"jyjl\":\"被检样品符合YY0505-2012标准要求、符合YY0601-2009标准第36章要求、符合YY0667-2008标准第36章要求、符合YY0668-2008标准第36章要求、符合YY0783-2008标准第36章要求、符合YY0784-2010标准第36章要求\",\"bz\":\"报告中“/”表示此项空白，“—”表示不适用。\",\"ypbh\":\"QW2018-0698\",\"xhgg\":\"M8102A\",\"jylb\":\"委托检验\",\"cpbhph\":\"DE65528125\",\"cydbh\":\"\",\"scrq\":\"2018-02-16\",\"ypsl\":\"1台\",\"cyjs\":\"\",\"jydd\":\"本所实验室\",\"jyrq\":\"2018年5月22日~2018年7月13日\",\"jydd\":\"本所实验室\",\"ypms\":\"见本报告第3页“1受检样品信息”。\",\"xhgghqtsm\":\"1.检测结果不包括不确定度的估算值。2.ECG附件有63个型号：M1631A、M1671A、M1984A、M1611A、M1968A、M1625A、M1639A、M1675A、M1602A、M1974A、M1601A、M1635A、M1678A、M1976A、M1672A、M1673A、M1533A、M1971A、M1973A、M1684A、M1613A、M1681A、M1558A、M1609A、M1683A、M1621A、M1674A、M1604A、M1685A、M1603A、M1619A、M1669A、M1645A、M1510A、M1500A、M1520A、M1979A、M1530A、M1557A、M1644A、M1605A、M1680A、M1537A、M1647A、M1532A、M1978A、M1615A、M1633A、M1668A、M1629A、M1663A、M1667A、M1623A、M1538A、M1665A、M1682A、M1540C、M1550C、M1560C、M1570C、989803170171、989803170181、989803143201。其电气原理和材料组成完全一致,   仅导联数与长度有所区别。本次检测了M1663A，M1978A，M1971A。SpO2附件有5个型号：M1192A、M1193A、M1194A、M1195A、M1196A，其电气原理和材料组成完全一致，仅长度和适用人群有所区别。本次检测了M1196A。CO2附件有17个型号：M2516A、M2761A、M2772A、M2751A、M2750A、M2745A、M2756A、M2757A、M2501A、M2768A、M2773A、M2741A、M2536A、M2746A、M2776A、M2777A、M1920A。其产品结构及原理均相同。本次检测了M2741A。温度探头有11个型号：21075A、21076A、21078A、M1837A、21091A、21093A、21094A、21095A、21090A、21082A、21082B。其电气原理和材料组成完全一致，仅长度和适用范围有所区别，本次检测了M21075A。袖带（含连接管）共有8个型号：M1571A、M1572A、M1573A、M1574A、M1575A、M1576A、M1598B、M1599B。其电气原理及材料组成完全一致，仅围度和连接管长度有所区别。本次检测了M1598B和M1574A。\"},\"cssbList\":[{\"cssbxh\":\"1\",\"cssbbhxlh\":\"2-FW-11\",\"cssbmc\":\"测试接收机\",\"cssbzzs\":\"R&S\",\"cssbxhgg\":\"ESH2-Z5\",\"cssbxcjzrq\":\"2019.5.15\",\"cssbbz\":\" \"},{\"cssbxh\":\"2\",\"cssbbhxlh\":\"2-FW-12\",\"cssbmc\":\"人工电源网络\",\"cssbzzs\":\"R&S\",\"cssbxhgg\":\"ESCI\",\"cssbxcjzrq\":\"2019.5.15\",\"cssbbz\":\" \"},{\"cssbxh\":\"3\",\"cssbbhxlh\":\"2-FW-103\",\"cssbmc\":\"屏蔽室1\",\"cssbzzs\":\"ETS·LINDGREN\",\"cssbxhgg\":\"SR1\",\"cssbxcjzrq\":\"2019.4.14\",\"cssbbz\":\" \"},{\"cssbxh\":\"4\",\"cssbbhxlh\":\"2-FW-93\",\"cssbmc\":\"测试接收机\",\"cssbzzs\":\"R&S\",\"cssbxhgg\":\"ESU26\",\"cssbxcjzrq\":\"2019.5.15\",\"cssbbz\":\" \"},{\"cssbxh\":\"5\",\"cssbbhxlh\":\"2-FW-101\",\"cssbmc\":\"双锥复合对数周期天线\",\"cssbzzs\":\"SCHWARZBECK\",\"cssbxhgg\":\"VULB9163\",\"cssbxcjzrq\":\"2020.2.13\",\"cssbbz\":\" \"},{\"cssbxh\":\"6\",\"cssbbhxlh\":\"2-FW-102\",\"cssbmc\":\"10米法电波暗室\",\"cssbzzs\":\"ETS·LINDGREN\",\"cssbxhgg\":\"FACT10\",\"cssbxcjzrq\":\"2019.4.14\",\"cssbbz\":\" \"},{\"cssbxh\":\"7\",\"cssbbhxlh\":\"2-FW-102\",\"cssbmc\":\"控制室\",\"cssbzzs\":\"ETS·LINDGREN\",\"cssbxhgg\":\"CR\",\"cssbxcjzrq\":\"2019.4.14\",\"cssbbz\":\" \"},{\"cssbxh\":\"8\",\"cssbbhxlh\":\"2-FW-163\",\"cssbmc\":\"静电放电器\",\"cssbzzs\":\"EM TEST\",\"cssbxhgg\":\"dito\",\"cssbxcjzrq\":\"2019.1.22\",\"cssbbz\":\" \"},{\"cssbxh\":\"9\",\"cssbbhxlh\":\"2-FW-106\",\"cssbmc\":\"屏蔽室3\",\"cssbzzs\":\"ETS·LINDGREN\",\"cssbxhgg\":\"SR3\",\"cssbxcjzrq\":\"2019.4.14\",\"cssbbz\":\" \"},{\"cssbxh\":\"10\",\"cssbbhxlh\":\"2-FW-30\",\"cssbmc\":\"功率放大器\",\"cssbzzs\":\"BONN\",\"cssbxhgg\":\"BLWA0830-160/100/40D\",\"cssbxcjzrq\":\"\",\"cssbbz\":\" \"},{\"cssbxh\":\"11\",\"cssbbhxlh\":\"2-FW-34\",\"cssbmc\":\"场强表\",\"cssbzzs\":\"AR\",\"cssbxhgg\":\"FL7006/Kit M1\",\"cssbxcjzrq\":\"2019.9.11\",\"cssbbz\":\" \"},{\"cssbxh\":\"12\",\"cssbbhxlh\":\"2-FW-100\",\"cssbmc\":\"信号发生器\",\"cssbzzs\":\"R&S\",\"cssbxhgg\":\"SMB100A-B106\",\"cssbxcjzrq\":\"2019.5.15\",\"cssbbz\":\" \"}],\"experiment\":[{\"name\":\"传导发射实验\",\"syjg\":\"符合\",\"jyrq\":\"2020-01-01\",\"wd\":\"100\",\"xdsd\":\"65\",\"dqyl\":\"100\",\"sysj\":[{\"sygdy\":\"AC220V 50Hz\",\"syplfw\":\"0.15MHz~30MHz\",\"ypyxms\":\"① \",\"rtf\":[{\"name\":\"ZC2018-128  生物安全柜 模式1 CE L.rtf\"},{\"name\":\"ZC2018-128  生物安全柜 模式1 CE N.rtf\"}]},{\"sygdy\":\"AC220V 50Hz\",\"syplfw\":\"0.15MHz~30MHz\",\"ypyxms\":\"② \",\"rtf\":[{\"name\":\"ZC2018-128  生物安全柜 模式1 CE L - 副本.rtf\"},{\"name\":\"ZC2018-128  生物安全柜 模式1 CE N - 副本.rtf\"}]}],\"syljt\":[{\"name\":\"image1.jpg\",content:\"\"}],\"sybzt\":[{\"name\":\"image2.jpg\",content:\"①③\"},{\"name\":\"image2.jpg\",content:\"②④\"}]},{\"name\":\"辐射发射试验\",\"syjg\":\"符合\",\"jyrq\":\"2020-01-01\",\"wd\":\"100\",\"xdsd\":\"65\",\"dqyl\":\"100\",\"sysj\":[{\"sygdy\":\"AC220V 50Hz\",\"syplfw\":\"0.15MHz~30MHz\",\"ypyxms\":\"① \",\"rtf\":[{\"name\":\"QW2018-2065 冷冻射频肿瘤治疗系统   RE.rtf\"}]},{\"sygdy\":\"AC220V 50Hz\",\"syplfw\":\"0.15MHz~30MHz\",\"ypyxms\":\"② \",\"rtf\":[{\"name\":\"QW2018-2065 冷冻射频肿瘤治疗系统   RE.rtf\"}]}],\"syljt\":[{\"name\":\"reljt.jpg\",content:\"\"}],\"sybzt\":[{\"name\":\"re1.jpg\",content:\"①③\"},{\"name\":\"re2.jpg\",content:\"②④\"}]},{\"name\":\"静电放电\",\"syjg\":\"符合\",\"jyrq\":\"2020-01-01\",\"wd\":\"100\",\"xdsd\":\"65\",\"dqyl\":\"100\",\"sysj\":[{\"sygdy\":\"AC220V 50Hz\",\"syplfw\":\"0.15MHz~30MHz\",\"ypyxms\":\"① \",\"html\":[{\"table\":\"<html> <head>     <meta charset='utf-8'> 	<style> 		table,th,td{ 			border:1px solid #000000; 			border-spacing:0; 			border-collapse:collapse; 		} 		.center{ 			text-align:center; 		} 	</style> </head> <body> <table class='custom-table white' id='test_electrostatic_data' menu='true' style='width:100%;'><tbody id='test_electrostatic_data'> <tr class='sample-line'>     <th colspan='11' rowspan='1' class='whole-line'>试验数据     </th> </tr> <tr class='sample-line'>     <td colspan='3' rowspan='1' class='table-label'>放电电压（kV）</td>     <td class='center'>+2</td>     <td class='center'>-2</td>     <td class='center'>+4</td>     <td class='center'>-4</td>     <td class='center'>+6</td>     <td class='center'>-6</td>     <td class='center'>+8</td>     <td class='center'>-8</td>      </tr> <tr class='sample-line'>     <td colspan='2' rowspan='3' deepth='1' class='right-click table-label'>空气放电</td>     <td deepth='2' colspan='1' rowspan='1'>test</td>     <td deepth='3' colspan='1' rowspan='1' class='center value-switching' >√</td>     <td deepth='4' colspan='1' rowspan='1' class='center value-switching' >√</td>     <td deepth='5' colspan='1' rowspan='1' class='center value-switching' >√</td>     <td deepth='6' colspan='1' rowspan='1' class='center value-switching' >√</td>     <td deepth='7' colspan='1' rowspan='1' class='center value-switching' >√</td>     <td deepth='8' colspan='1' rowspan='1' class='center value-switching' >√</td>     <td deepth='9' colspan='1' rowspan='1' class='center value-switching' >√</td>     <td deepth='10' colspan='1' rowspan='1' class='center value-switching'>√</td>      </tr> <tr class='sample-line'>     <td deepth='2' colspan='1' rowspan='1'>test</td>     <td deepth='3' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='4' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='5' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='6' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='7' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='8' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='9' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='10' colspan='1' rowspan='1' class='center value-switching'>√</td>      </tr> <tr class='sample-line'>     <td deepth='2' colspan='1' rowspan='1'>test</td>     <td deepth='3' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='4' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='5' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='6' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='7' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='8' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='9' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='10' colspan='1' rowspan='1' class='center value-switching'>√</td>      </tr> <tr class='sample-line'>     <td colspan='1' rowspan='4' deepth='1' class='table-label'>接触放电</td>     <td colspan='1' rowspan='2' deepth='2' class='right-click table-label'>直接</td>     <td deepth='3' colspan='1' rowspan='1'>test</td>     <td deepth='4' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='5' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='6' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='7' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='8' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='9' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='10' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='11' colspan='1' rowspan='1' class='center value-switching'>√</td>      </tr> <tr class='sample-line'>     <td deepth='3' colspan='1' rowspan='1'>test</td>     <td deepth='4' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='5' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='6' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='7' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='8' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='9' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='10' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='11' colspan='1' rowspan='1' class='center value-switching'>√</td>      </tr> <tr class='sample-line'>     <td colspan='1' rowspan='2' deepth='2' class='right-click table-label'>间接</td>     <td deepth='3' colspan='1' rowspan='1'>HCP</td>     <td deepth='4' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='5' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='6' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='7' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='8' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='9' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='10' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='11' colspan='1' rowspan='1' class='center value-switching'>√</td>      </tr> <tr class='sample-line'>     <td deepth='3' colspan='1' rowspan='1'>VCP</td>     <td deepth='4' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='5' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='6' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='7' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='8' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='9' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='10' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='11' colspan='1' rowspan='1' class='center value-switching'>√</td>      </tr> <tr class='sample-line'>     <td colspan='3' rowspan='1' class='table-label'>异常现象描述</td>     <td colspan='8' rowspan='1'>这是一个异常现象</td> </tr> <tr class='sample-line'>     <td colspan='3' rowspan='1' class='table-label'>备注</td>     <td colspan='8' rowspan='1'>这是一条备注</td> </tr> </tbody>  </table></body></html>\"}]}],\"syljt\":[{\"name\":\"reljt.jpg\",content:\"\"}],\"sybzt\":[{\"name\":\"re1.jpg\",content:\"①③\"},{\"name\":\"re2.jpg\",content:\"②④\"}]}]}";

        private string jsonStr1 = "{\"scbWord\":\"5b838ac0-8ea1-4b13-9322-138d32534845.docx\",\"firstPage\":{\"main_wtf\":\"飞利浦(中国)投资有限公司1\",\"main_ypmc\":\"病人监护仪1\",\"main_xhgg\":\"M8102A1\",\"main_jylb\":\"委托检验1\",\"ypmc\":\"病人监护仪\",\"sb\":\"\",\"wtf\":\"飞利浦（中国）投资有限公司\",\"wtfdz\":\"上海市静安区灵石路718号A幢\",\"scdw\":\"PhilipsMedizinSystemeBoeblingenGmbH\",\"sjdw\":\"PhilipsMedizinSystemeBoeblingenGmbH\",\"cydw\":\"\",\"cydd\":\"\",\"cyrq\":\"\",\"dyrq\":\"2018年5月8日\",\"jyxm\":\"YY0505全项目、YY0601中36、YY0667中36、YY0668中36、YY0783中36、YY0784中36\",\"jyyj\":\"YY0505-2012《医用电气设备第1-2部分：安全通用要求并列标准电磁兼容要求和试验》、YY0601-2009《医用电气设备呼吸气体监护仪的基本安全和主要性能专用要求》、YY0667-2008《医用电气设备第2-30部分：自动循环无创血压监护设备的安全和基本性能专用要求》、YY0668-2008《医用电气设备第2-49部分：多参数患者监护设备安全专用要求》、YY0783-2010《医用电气设备第2-34部分：有创血压监测设备的安全和基本性能专用要求》、YY0784-2010《医用电气设备医用脉搏血氧仪设备基本安全和主要性能专用要求》\",\"jyjl\":\"被检样品符合YY0505-2012标准要求、符合YY0601-2009标准第36章要求、符合YY0667-2008标准第36章要求、符合YY0668-2008标准第36章要求、符合YY0783-2008标准第36章要求、符合YY0784-2010标准第36章要求\",\"bz\":\"报告中“/”表示此项空白，“—”表示不适用。\",\"ypbh\":\"QW2018-0698\",\"xhgg\":\"M8102A\",\"jylb\":\"委托检验\",\"cpbhph\":\"DE65528125\",\"cydbh\":\"\",\"scrq\":\"2018-02-16\",\"ypsl\":\"1台\",\"cyjs\":\"\",\"jydd\":\"本所实验室\",\"jyrq\":\"2018年5月22日~2018年7月13日\",\"jydd\":\"本所实验室\",\"ypms\":\"见本报告第3页“1受检样品信息”。\",\"xhgghqtsm\":\"1.检测结果不包括不确定度的估算值。2.ECG附件有63个型号：M1631A、M1671A、M1984A、M1611A、M1968A、M1625A、M1639A、M1675A、M1602A、M1974A、M1601A、M1635A、M1678A、M1976A、M1672A、M1673A、M1533A、M1971A、M1973A、M1684A、M1613A、M1681A、M1558A、M1609A、M1683A、M1621A、M1674A、M1604A、M1685A、M1603A、M1619A、M1669A、M1645A、M1510A、M1500A、M1520A、M1979A、M1530A、M1557A、M1644A、M1605A、M1680A、M1537A、M1647A、M1532A、M1978A、M1615A、M1633A、M1668A、M1629A、M1663A、M1667A、M1623A、M1538A、M1665A、M1682A、M1540C、M1550C、M1560C、M1570C、989803170171、989803170181、989803143201。其电气原理和材料组成完全一致,   仅导联数与长度有所区别。本次检测了M1663A，M1978A，M1971A。SpO2附件有5个型号：M1192A、M1193A、M1194A、M1195A、M1196A，其电气原理和材料组成完全一致，仅长度和适用人群有所区别。本次检测了M1196A。CO2附件有17个型号：M2516A、M2761A、M2772A、M2751A、M2750A、M2745A、M2756A、M2757A、M2501A、M2768A、M2773A、M2741A、M2536A、M2746A、M2776A、M2777A、M1920A。其产品结构及原理均相同。本次检测了M2741A。温度探头有11个型号：21075A、21076A、21078A、M1837A、21091A、21093A、21094A、21095A、21090A、21082A、21082B。其电气原理和材料组成完全一致，仅长度和适用范围有所区别，本次检测了M21075A。袖带（含连接管）共有8个型号：M1571A、M1572A、M1573A、M1574A、M1575A、M1576A、M1598B、M1599B。其电气原理及材料组成完全一致，仅围度和连接管长度有所区别。本次检测了M1598B和M1574A。\"},\"cssbList\":[{\"cssbxh\":\"1\",\"cssbbhxlh\":\"2-FW-11\",\"cssbmc\":\"测试接收机\",\"cssbzzs\":\"R&S\",\"cssbxhgg\":\"ESH2-Z5\",\"cssbxcjzrq\":\"2019.5.15\",\"cssbbz\":\" \"},{\"cssbxh\":\"2\",\"cssbbhxlh\":\"2-FW-12\",\"cssbmc\":\"人工电源网络\",\"cssbzzs\":\"R&S\",\"cssbxhgg\":\"ESCI\",\"cssbxcjzrq\":\"2019.5.15\",\"cssbbz\":\" \"},{\"cssbxh\":\"3\",\"cssbbhxlh\":\"2-FW-103\",\"cssbmc\":\"屏蔽室1\",\"cssbzzs\":\"ETS·LINDGREN\",\"cssbxhgg\":\"SR1\",\"cssbxcjzrq\":\"2019.4.14\",\"cssbbz\":\" \"},{\"cssbxh\":\"4\",\"cssbbhxlh\":\"2-FW-93\",\"cssbmc\":\"测试接收机\",\"cssbzzs\":\"R&S\",\"cssbxhgg\":\"ESU26\",\"cssbxcjzrq\":\"2019.5.15\",\"cssbbz\":\" \"},{\"cssbxh\":\"5\",\"cssbbhxlh\":\"2-FW-101\",\"cssbmc\":\"双锥复合对数周期天线\",\"cssbzzs\":\"SCHWARZBECK\",\"cssbxhgg\":\"VULB9163\",\"cssbxcjzrq\":\"2020.2.13\",\"cssbbz\":\" \"},{\"cssbxh\":\"6\",\"cssbbhxlh\":\"2-FW-102\",\"cssbmc\":\"10米法电波暗室\",\"cssbzzs\":\"ETS·LINDGREN\",\"cssbxhgg\":\"FACT10\",\"cssbxcjzrq\":\"2019.4.14\",\"cssbbz\":\" \"},{\"cssbxh\":\"7\",\"cssbbhxlh\":\"2-FW-102\",\"cssbmc\":\"控制室\",\"cssbzzs\":\"ETS·LINDGREN\",\"cssbxhgg\":\"CR\",\"cssbxcjzrq\":\"2019.4.14\",\"cssbbz\":\" \"},{\"cssbxh\":\"8\",\"cssbbhxlh\":\"2-FW-163\",\"cssbmc\":\"静电放电器\",\"cssbzzs\":\"EM TEST\",\"cssbxhgg\":\"dito\",\"cssbxcjzrq\":\"2019.1.22\",\"cssbbz\":\" \"},{\"cssbxh\":\"9\",\"cssbbhxlh\":\"2-FW-106\",\"cssbmc\":\"屏蔽室3\",\"cssbzzs\":\"ETS·LINDGREN\",\"cssbxhgg\":\"SR3\",\"cssbxcjzrq\":\"2019.4.14\",\"cssbbz\":\" \"},{\"cssbxh\":\"10\",\"cssbbhxlh\":\"2-FW-30\",\"cssbmc\":\"功率放大器\",\"cssbzzs\":\"BONN\",\"cssbxhgg\":\"BLWA0830-160/100/40D\",\"cssbxcjzrq\":\"\",\"cssbbz\":\" \"},{\"cssbxh\":\"11\",\"cssbbhxlh\":\"2-FW-34\",\"cssbmc\":\"场强表\",\"cssbzzs\":\"AR\",\"cssbxhgg\":\"FL7006/Kit M1\",\"cssbxcjzrq\":\"2019.9.11\",\"cssbbz\":\" \"},{\"cssbxh\":\"12\",\"cssbbhxlh\":\"2-FW-100\",\"cssbmc\":\"信号发生器\",\"cssbzzs\":\"R&S\",\"cssbxhgg\":\"SMB100A-B106\",\"cssbxcjzrq\":\"2019.5.15\",\"cssbbz\":\" \"}],\"experiment\":[{\"name\":\"传导发射实验\",\"syjg\":\"符合\",\"jyrq\":\"2020-01-01\",\"wd\":\"100\",\"xdsd\":\"65\",\"dqyl\":\"100\",\"sysj\":[{\"sygdy\":\"AC220V 50Hz\",\"syplfw\":\"0.15MHz~30MHz\",\"ypyxms\":\"① \",\"mccfpl\":\"AC220V 60Hz\",\"rtf\":[{\"name\":\"54b5f48c-ae87-4321-ba2a-1fa50c4e4411.Rtf\"},{\"name\":\"9d8a5174-1159-4db9-9f12-d4fb051c735c.Rtf\"}]},{\"sygdy\":\"AC220V 50Hz\",\"syplfw\":\"0.15MHz~30MHz\",\"ypyxms\":\"② \",\"rtf\":[{\"name\":\"093dd9de-e1c9-4c1d-952e-cd3f372ae14b.Rtf\"},{\"name\":\"72a42c7c-66fc-4395-a13a-1b9191b1b8c5.Rtf\"}]}],\"syljt\":[{\"name\":\"78b4cd30-4255-4051-948a-45da4bdeb7a2.jpg\",content:\"\"}],\"sybzt\":[{\"name\":\"6da7e28a-e492-4a8a-8ac3-2bb03e05b132.jpg\",content:\"①③\"},{\"name\":\"c3030f9b-7cc6-4a74-910a-54c415364694.jpg\",content:\"②④\"}]},{\"name\":\"辐射发射试验\",\"syjg\":\"符合\",\"jyrq\":\"2020-01-01\",\"wd\":\"100\",\"xdsd\":\"65\",\"dqyl\":\"100\",\"sysj\":[{\"sygdy\":\"AC220V 50Hz\",\"syplfw\":\"0.15MHz~30MHz\",\"ypyxms\":\"① \",\"rtf\":[{\"name\":\"857dd599-e88a-423a-a158-80befbbd4506.Rtf\"}]},{\"sygdy\":\"AC220V 50Hz\",\"syplfw\":\"0.15MHz~30MHz\",\"ypyxms\":\"② \",\"rtf\":[{\"name\":\"857dd599-e88a-423a-a158-80befbbd4506.Rtf\"}]}],\"syljt\":[{\"name\":\"98339fa7-3cf5-4434-aa13-a0eb1848099e.jpg\",content:\"\"}],\"sybzt\":[{\"name\":\"c3030f9b-7cc6-4a74-910a-54c415364694.jpg\",content:\"①③\"},{\"name\":\"c3030f9b-7cc6-4a74-910a-54c415364694.jpg\",content:\"②④\"}]},{\"name\":\"静电放电\",\"syjg\":\"符合\",\"jyrq\":\"2020-01-01\",\"wd\":\"100\",\"xdsd\":\"65\",\"dqyl\":\"100\",\"sysj\":[{\"sygdy\":\"AC220V 50Hz\",\"syplfw\":\"0.15MHz~30MHz\",\"ypyxms\":\"① \",\"html\":[{\"table\":\"<html> <head>     <meta charset='utf-8'> 	<style> 		table,th,td{ 			border:1px solid #000000; 			border-spacing:0; 			border-collapse:collapse; 		} 		.center{ 			text-align:center; 		} 	</style> </head> <body> <table class='custom-table white' id='test_electrostatic_data' menu='true' style='width:100%;'><tbody id='test_electrostatic_data'> <tr class='sample-line'>     <th colspan='11' rowspan='1' class='whole-line'>试验数据     </th> </tr> <tr class='sample-line'>     <td colspan='3' rowspan='1' class='table-label'>放电电压（kV）</td>     <td class='center'>+2</td>     <td class='center'>-2</td>     <td class='center'>+4</td>     <td class='center'>-4</td>     <td class='center'>+6</td>     <td class='center'>-6</td>     <td class='center'>+8</td>     <td class='center'>-8</td>      </tr> <tr class='sample-line'>     <td colspan='2' rowspan='3' deepth='1' class='right-click table-label'>空气放电</td>     <td deepth='2' colspan='1' rowspan='1'>test</td>     <td deepth='3' colspan='1' rowspan='1' class='center value-switching' >√</td>     <td deepth='4' colspan='1' rowspan='1' class='center value-switching' >√</td>     <td deepth='5' colspan='1' rowspan='1' class='center value-switching' >√</td>     <td deepth='6' colspan='1' rowspan='1' class='center value-switching' >√</td>     <td deepth='7' colspan='1' rowspan='1' class='center value-switching' >√</td>     <td deepth='8' colspan='1' rowspan='1' class='center value-switching' >√</td>     <td deepth='9' colspan='1' rowspan='1' class='center value-switching' >√</td>     <td deepth='10' colspan='1' rowspan='1' class='center value-switching'>√</td>      </tr> <tr class='sample-line'>     <td deepth='2' colspan='1' rowspan='1'>test</td>     <td deepth='3' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='4' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='5' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='6' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='7' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='8' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='9' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='10' colspan='1' rowspan='1' class='center value-switching'>√</td>      </tr> <tr class='sample-line'>     <td deepth='2' colspan='1' rowspan='1'>test</td>     <td deepth='3' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='4' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='5' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='6' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='7' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='8' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='9' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='10' colspan='1' rowspan='1' class='center value-switching'>√</td>      </tr> <tr class='sample-line'>     <td colspan='1' rowspan='4' deepth='1' class='table-label'>接触放电</td>     <td colspan='1' rowspan='2' deepth='2' class='right-click table-label'>直接</td>     <td deepth='3' colspan='1' rowspan='1'>test</td>     <td deepth='4' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='5' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='6' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='7' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='8' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='9' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='10' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='11' colspan='1' rowspan='1' class='center value-switching'>√</td>      </tr> <tr class='sample-line'>     <td deepth='3' colspan='1' rowspan='1'>test</td>     <td deepth='4' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='5' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='6' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='7' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='8' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='9' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='10' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='11' colspan='1' rowspan='1' class='center value-switching'>√</td>      </tr> <tr class='sample-line'>     <td colspan='1' rowspan='2' deepth='2' class='right-click table-label'>间接</td>     <td deepth='3' colspan='1' rowspan='1'>HCP</td>     <td deepth='4' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='5' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='6' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='7' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='8' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='9' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='10' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='11' colspan='1' rowspan='1' class='center value-switching'>√</td>      </tr> <tr class='sample-line'>     <td deepth='3' colspan='1' rowspan='1'>VCP</td>     <td deepth='4' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='5' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='6' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='7' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='8' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='9' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='10' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='11' colspan='1' rowspan='1' class='center value-switching'>√</td>      </tr> <tr class='sample-line'>     <td colspan='3' rowspan='1' class='table-label'>异常现象描述</td>     <td colspan='8' rowspan='1'>这是一个异常现象</td> </tr> <tr class='sample-line'>     <td colspan='3' rowspan='1' class='table-label'>备注</td>     <td colspan='8' rowspan='1'>这是一条备注</td> </tr> </tbody>  </table></body></html>\"}]}],\"syljt\":[{\"name\":\"98339fa7-3cf5-4434-aa13-a0eb1848099e.jpg\",content:\"\"}],\"sybzt\":[{\"name\":\"c3030f9b-7cc6-4a74-910a-54c415364694.jpg\",content:\"①③\"},{\"name\":\"c3030f9b-7cc6-4a74-910a-54c415364694.jpg\",content:\"②④\"}]}]}";


        private string _currRoot = AppDomain.CurrentDomain.BaseDirectory;

        [HttpGet]
        public IEnumerable<string> Get()
        {
            return new string[] { "Emc", "生成报告" };
        }

        /// <summary>
        /// 生成报告
        /// </summary>
        /// <param name="para">参数</param>
        /// <returns></returns>
        [HttpPost]
        public IHttpActionResult CreateReport(ReportParams para)
        {
            //string jsonStr = para.JsonStr;
            string reportId = para.ReportId;
            Stopwatch sw = new Stopwatch();
            sw.Start();

            ReportResult<string> result = new ReportResult<string>();
            try
            {
                //获取zip文件 
                string reportFilesPath = FileUtil.CreateReportDirectory(string.Format("{0}\\Files\\ReportFiles", _currRoot));
                string reportZipFilesPath = string.Format("{0}\\zip{1}.zip", reportFilesPath, DateTime.Now.ToString("yyyyMMddhhmmss"));
                string zipUrl = ConfigurationManager.AppSettings["ReportFilesUrl"].ToString() + reportId + "?timestamp=" + MyTools.GetTimestamp(DateTime.Now);
                if (para.ZipFilesUrl!=null&&!para.ZipFilesUrl.Equals("")) {
                    zipUrl = para.ZipFilesUrl;
                }
                byte[] fileBytes = SyncHttpHelper.GetHttpRespponseForFile(zipUrl, reportZipFilesPath,
                int.Parse(DateTime.Now.ToString("hhmmss")));
                if (fileBytes.Length <= 0)
                {
                    result = SetReportResult<string>("请求报告文件失败", false, para.ReportId.ToString());
                    MyTools.ErrorLog.Error(string.Format("请求报告失败,报告id:{0}", para.ReportId));
                    return Json<ReportResult<string>>(result);
                }
                //解压zip文件
                ZipFileHelper.DecompressionZip(reportZipFilesPath, reportFilesPath);
                //生成报告
                string content = JsonToWord(para.JsonStr.Equals("")?jsonStr1: para.JsonStr, reportFilesPath);
                sw.Stop();
                double time1 = (double)sw.ElapsedMilliseconds / 1000;
                result = SetReportResult<string>(string.Format("报告生成成功,用时:" + time1.ToString()), true, content);
                MyTools.InfoLog.Info("报告:" + result.Content + ",信息:" + result.Message);

            }
            catch (Exception ex)
            {
                MyTools.ErrorLog.Error(ex.Message, ex);//设置错误信息
                result= SetReportResult<string>(string.Format("报告生成失败"), false, ex.Message);
                return Json<ReportResult<string>>(result);
            }

            return Json<ReportResult<string>>(result);
        }

        [HttpPost]
        public IHttpActionResult CreateReportTest(ReportParams para)
        {
            //string jsonStr = para.JsonStr;
            string reportId = para.ReportId;
            Stopwatch sw = new Stopwatch();
            sw.Start();

            ReportResult<string> result = new ReportResult<string>();
            try
            {
                //获取zip文件 
                string reportFilesPath = FileUtil.CreateReportDirectory(string.Format("{0}\\Files\\ReportFiles", _currRoot));
                string reportZipFilesPath = string.Format("{0}Files\\ReportFiles\\Test\\{1}", _currRoot, "QT2019-3015.zip");
                //解压zip文件
                ZipFileHelper.DecompressionZip(reportZipFilesPath, reportFilesPath);
                
                //生成报告
                string content = JsonToWord(para.JsonStr.Equals("") ? jsonStr1 : para.JsonStr, reportFilesPath);
                sw.Stop();
                double time1 = (double)sw.ElapsedMilliseconds / 1000;
                result = SetReportResult<string>(string.Format("报告生成成功,用时:" + time1.ToString()), true, content);
                MyTools.InfoLog.Info("报告:" + result.Content + ",信息:" + result.Message);
            }
            catch (Exception ex)
            {
                MyTools.ErrorLog.Error(ex.Message, ex);
                throw ex;
            }

            return Json<ReportResult<string>>(result);
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

        /// <summary>
        /// 下载文件
        /// </summary>
        [HttpGet]
        public IHttpActionResult DownloadFilesForGet(string fileName)
        {
            //string fileName = para.FileName;
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

        /// <summary>
        /// 测试
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public string Get2()
        {
            Stopwatch sw = new Stopwatch();

            sw.Start();
            MyTools.KillWordProcess();

            string reportFilesPath = FileUtil.CreateReportDirectory(string.Format("{0}\\Files\\ReportFiles", _currRoot));
            string reportZipFilesPath = string.Format("{0}Files\\ReportFiles\\Test\\{1}", _currRoot, "QT2019-3015.zip");
            //解压zip文件
            ZipFileHelper.DecompressionZip(reportZipFilesPath, reportFilesPath);

            string result = JsonToWord(jsonStr1, reportFilesPath);
            //string result = "";
            sw.Stop();
            double time1 = (double)sw.ElapsedMilliseconds / 1000;
            return result + ":" + time1.ToString();
        }

        /// <summary>
        /// 测试2
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public string Get4()
        {
            Stopwatch sw = new Stopwatch();

            sw.Start();
            MyTools.KillWordProcess();
            string result = "";
            //解压zip文件 
            string reportFilesPath = FileUtil.CreateReportDirectory(string.Format("{0}\\Files\\ReportFiles", _currRoot));
            byte[] fileBytes = SyncHttpHelper.GetHttpRespponseForFile(ConfigurationManager.AppSettings["ReportFilesUrl"].ToString() + "QT2019-3015", reportFilesPath,
            int.Parse(DateTime.Now.ToString("hhmmss")));
            if (fileBytes.Length <= 0)
            {
                result = "下载文件失败";
            }
            else
            {
                //生成报告
                result = JsonToWord(jsonStr, reportFilesPath);
            }

            //string result = "";
            sw.Stop();
            double time1 = (double)sw.ElapsedMilliseconds / 1000;
            return result + ":" + time1.ToString();
        }

        [HttpGet]
        public string Get3()
        {
            Stopwatch sw = new Stopwatch();
            sw.Start();
            string result = TestCreateReportDirectory("11");
            sw.Stop();
            double time1 = (double)sw.ElapsedMilliseconds / 1000;
            return result + ":" + time1.ToString();
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

        private string GetTestPath(string fileName)
        {
            string currRoot = AppDomain.CurrentDomain.BaseDirectory;
            string imageFullFileName = string.Format(@"{0}\Files\ReportFiles\Test\{1}", currRoot, fileName);
            return imageFullFileName;
        }
        
        #region 生成报表方法
        private string JsonToWord(string jsonStr, string reportFilesPath)
        {
            //解析json字符串
            JObject mainObj = (JObject)JsonConvert.DeserializeObject(jsonStr);
            string currRoot = AppDomain.CurrentDomain.BaseDirectory;
            string outfileName = string.Format("report{0}.docx", MyTools.GetTimestamp(DateTime.Now));//输出文件名称
            string outfilePth = string.Format(@"{0}\Files\OutPut\{1}", currRoot, outfileName);//输出文件路径
            string filePath = string.Format(@"{0}\Files\{1}", currRoot, ConfigurationManager.AppSettings["TemplateName"].ToString());//模板文件

            string middleDir = currRoot + "\\Files\\TemplateMiddleware\\" + DateTime.Now.ToString("yyyyMMddhhmmss");
            filePath = CreateTemplateMiddle(middleDir, "template", filePath);
            string result = "保存成功1";
            //生成报告
            using (WordUtil wordUtil = new WordUtil(outfilePth, filePath))
            {
                //审查表 //测试数据
                string scbWord = reportFilesPath+"\\" + (string)mainObj["scbWord"];

                //首页内容 object
                JObject firstPage = (JObject)mainObj["firstPage"];
                result = InsertContentToWord(wordUtil, firstPage);
                if (!result.Equals("保存成功"))
                {
                    return result;
                }

                //受检样品描述 object  sjypms (审查表)
                GetTableFromReview(wordUtil, "sjypms", scbWord, 3, false);

                //样品构成 list ypgcList (审查表)
                GetTableFromReview(wordUtil, "ypgcList", scbWord, 4, false);

                //样品连接图 图片 connectionGraph (审查表)
                GetImageFomReview(wordUtil, "connectionGraph",scbWord, false);

                //样品运行模式 list ypyxList (审查表)
                GetTableFromReview(wordUtil, "ypyxList", scbWord, 6, false);

                //样品电缆 list ypdlList (审查表)
                GetTableFromReview(wordUtil, "ypdlList", scbWord, 7, false);

                //测试设备list cssbList 不动
                JArray cssbList = (JArray)mainObj["cssbList"];
                result = InsertListIntoTable(wordUtil, cssbList, 1, "cssblist");
                if (!result.Equals("保存成功"))
                {
                    return result;
                }

                //辅助设备 list fzsbList (审查表)
                GetTableFromReview(wordUtil, "fzsbList", scbWord, 5, true);

                //实验数据
                JArray experiment = (JArray)mainObj["experiment"];
                string newBookmark = "experiment";
                foreach (JObject item in experiment)
                {
                    if (item["name"].ToString().Equals("传导发射实验")|| item["name"].ToString().Equals("传导发射"))
                        newBookmark = SetConductedEmission(wordUtil, item, newBookmark, "CE", middleDir, reportFilesPath);
                    else if (item["name"].ToString().Equals("辐射发射试验") || item["name"].ToString().Equals("辐射发射"))
                        newBookmark = SetConductedEmission(wordUtil, item, newBookmark, "RE", middleDir, reportFilesPath);
                    else if (item["name"].ToString().Equals("谐波失真"))
                        newBookmark = SetConductedEmission(wordUtil, item, newBookmark, "谐波", middleDir, reportFilesPath);
                    else if (item["name"].ToString().Equals("电压波动和闪烁"))
                        newBookmark = SetConductedEmission(wordUtil, item, newBookmark, "波动", middleDir, reportFilesPath);
                    else
                    {
                        newBookmark = SetOtherEmission(wordUtil, item, newBookmark, middleDir, reportFilesPath);
                    }
                }
            }
            //删除中间件文件夹
            DelectDir(middleDir);
            DelectDir(reportFilesPath);

            return outfileName;
        }

        //设置首页内容
        private string InsertContentToWord(WordUtil wordUtil, JObject jo1)
        {
            foreach (var item in jo1)
            {
                wordUtil.InsertContentToWordByBookmark(item.Value.ToString(), item.Key);
            }
            return "保存成功";
        }
        //测试工具
        private string InsertListIntoTable(WordUtil wordUtil, JArray array, int mergeColumn, string bookmark, bool isNeedNumber = true)
        {
            List<string> list = JarrayToList(array);

            string result = wordUtil.InsertListToTable(list, bookmark, mergeColumn, isNeedNumber);

            return result;
        }

        //从审查表中取table数据
        private void GetTableFromReview(WordUtil wordUtil, string bookmark, string scbWordPath, int tableIndex, bool isCloseTheFile)
        {
            wordUtil.CopyTableToWord(scbWordPath, bookmark, tableIndex, isCloseTheFile);
        }

        //从审查表中取连接图
        private void GetImageFomReview(WordUtil wordUtil, string bookmark, string scbWordPath, bool isCloseTheFile)
        {
            wordUtil.CopyImageToWord(scbWordPath, bookmark, isCloseTheFile);
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
        #endregion

        #region 实验数据
        /// <summary>
        /// 传导发射实验 辐射发射实验
        /// </summary>
        /// <returns></returns>
        private string SetConductedEmission(WordUtil wordUtil, JObject jObject, string bookmark, string rtfType, string middleDir,string reportFilesPath)
        {
            string templateName = jObject["name"].ToString();
            string templateFullPath = CreateTemplateMiddle(middleDir, "experiment", GetTemplatePath(templateName + ".docx"));
            string sysjTemplateFilePath = CreateTemplateMiddle(middleDir, "sysj", GetTemplatePath("RTFTemplate.docx"));

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
                if (item["sygdy"] != null)
                    contentList.Add("试验供电电源：" + item["sygdy"].ToString());
                if (item["syplfw"] != null)
                    contentList.Add("试验频率范围：" + item["syplfw"].ToString());
                if (item["ypyxms"] != null)
                    contentList.Add("样品运行模式：" + item["ypyxms"].ToString());
                if (item["mccfpl"] != null)
                    contentList.Add("脉冲重复频率（kHz）：" + item["mccfpl"].ToString());
                if (item["sycxsj"] != null)
                    contentList.Add("试验持续时间（s）：" + item["sycxsj"].ToString());
                if (item["cfpl"] != null)
                    contentList.Add("重复频率（s）：" + item["cfpl"].ToString());
                if (item["cs"] != null)
                    contentList.Add("次数（次）：" + item["cs"].ToString());
                if (item["sycfcs"] != null)
                    contentList.Add("试验重复次数（次）：" + item["sycfcs"].ToString());
                if (item["sysjjg"] != null)
                    contentList.Add("试验时间间隔（s）：" + item["sysjjg"].ToString());
                if (item["sypl"] != null)
                    contentList.Add("试验频率（Hz）：" + item["sypl"].ToString());

                wordUtil.CreateTableToWord(sysjTemplateFilePath, contentList, "sysj", false, i != 0);

                JArray rtf = (JArray)item["rtf"];
                int rtfCount = rtf.Count;
                int j = 0;

                foreach (JObject rtfObj in (JArray)item["rtf"])
                {
                    //需要画表格和插入rtf内容
                    wordUtil.CopyOtherFileTableForColByTableIndex(sysjTemplateFilePath, reportFilesPath+"\\"+ rtfObj["name"].ToString(), startIndex, dic, rtfbookmark, false, true, false);

                    wordUtil.CopyOtherFilePictureToWord(sysjTemplateFilePath, reportFilesPath + "\\" + rtfObj["name"].ToString(), imageStartIndex, imageBookmark, false, true, j == rtfCount - 1);
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
                list.Add(reportFilesPath + "\\" + item["name"].ToString() + "," + item["content"].ToString());
            }

            wordUtil.InsertImageToTemplate(templateFullPath, list, "syljt", false);

            JArray sybzt = (JArray)jObject["sybzt"];
            list = new List<string>();
            foreach (JObject item in sybzt)
            {
                list.Add(reportFilesPath + "\\" + item["name"].ToString() + "," + item["content"].ToString());
            }
            wordUtil.InsertImageToTemplate(templateFullPath, list, "sybzt", false);

            string result = wordUtil.CopyOtherFileContentToWordReturnBookmark(templateFullPath, bookmark);


            return result;
        }

        /// <summary>
        /// 其他带有html的实验
        /// </summary>
        /// <returns></returns>
        private string SetOtherEmission(WordUtil wordUtil, JObject jObject, string bookmark, string middleDir, string reportFilesPath)
        {
            string templateName = jObject["name"].ToString();
            string templateFullPath = CreateTemplateMiddle(middleDir, "experiment", GetTemplatePath(templateName + ".docx"));
            string sysjTemplateFilePath = CreateTemplateMiddle(middleDir, "sysj", GetTemplatePath("RTFTemplate.docx"));

            foreach (var item in jObject)
            {
                if (!item.Key.Equals("sysj") && !item.Key.Equals("name") && !item.Key.Equals("syljt") && !item.Key.Equals("sybzt"))
                    wordUtil.InsertContentInBookmark(templateFullPath, item.Value.ToString(), item.Key, false);
            }

            JArray sysj = (JArray)jObject["sysj"];

            int i = 0;
            foreach (var item in sysj)
            {
                List<string> contentList = new List<string>();
                if (item["sygdy"] != null)
                    contentList.Add("试验供电电源：" + item["sygdy"].ToString());
                if (item["syplfw"] != null)
                    contentList.Add("试验频率范围：" + item["syplfw"].ToString());
                if (item["ypyxms"] != null)
                    contentList.Add("样品运行模式：" + item["ypyxms"].ToString());
                if (item["mccfpl"] != null)
                    contentList.Add("脉冲重复频率（kHz）：" + item["mccfpl"].ToString());
                if (item["sycxsj"] != null)
                    contentList.Add("试验持续时间（s）：" + item["sycxsj"].ToString());
                if (item["cfpl"] != null)
                    contentList.Add("重复频率（s）：" + item["cfpl"].ToString());
                if (item["cs"] != null)
                    contentList.Add("次数（次）：" + item["cs"].ToString());
                if (item["sycfcs"] != null)
                    contentList.Add("试验重复次数（次）：" + item["sycfcs"].ToString());
                if (item["sysjjg"] != null)
                    contentList.Add("试验时间间隔（s）：" + item["sysjjg"].ToString());
                if (item["sypl"] != null)
                    contentList.Add("试验频率（Hz）：" + item["sypl"].ToString());

                wordUtil.CreateTableToWord(sysjTemplateFilePath, contentList, "sysj", false, i != 0);

                JArray html = (JArray)item["html"];
                int htmlCount = html.Count;
                int j = 0;

                foreach (JObject rtfObj in html)
                {
                    //生成html并将内容插入到模板中
                    string htmlstr = (string)rtfObj["table"];
                    string htmlfullname = CreateHtmlFile(htmlstr, middleDir);
                    wordUtil.CopyHtmlContentToTemplate(htmlfullname, sysjTemplateFilePath, "sysj",true,true,false);
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
                list.Add(reportFilesPath + "\\" + item["name"].ToString() + "," + item["content"].ToString());
            }

            wordUtil.InsertImageToTemplate(templateFullPath, list, "syljt", false);

            JArray sybzt = (JArray)jObject["sybzt"];
            list = new List<string>();
            foreach (JObject item in sybzt)
            {
                list.Add(reportFilesPath + "\\" + item["name"].ToString() + "," + item["content"].ToString());
            }
            wordUtil.InsertImageToTemplate(templateFullPath, list, "sybzt", false);

            string result = wordUtil.CopyOtherFileContentToWordReturnBookmark(templateFullPath, bookmark);


            return result;
        }
        #endregion

        /// <summary>
        /// 创建模板中间件
        /// </summary>
        private string CreateTemplateMiddle(string dir, string template, string filePath)
        {

            string dateStr = DateTime.Now.ToString("yyyyMMddhhmmss");
            string fileName = template + dateStr + ".docx";
            DirectoryInfo di = new DirectoryInfo(dir);
            if (!di.Exists) { di.Create(); }

            string htmlpath = dir + "\\" + fileName;
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

        //删除模板中间件
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
                Directory.Delete(srcPath);
            }
            catch (Exception)
            {
                throw;
            }
        }

        #region 测试
        //测试下载解压zip文件
        private string TestCreateReportDirectory(string reportId)
        {
            string filePath = string.Format("{0}\\ReportFiles", _currRoot);
            string createOutputDir = FileUtil.CreateReportDirectory(filePath);
            string reportFilesZip = string.Format("{0}\\{1}.docx", createOutputDir, DateTime.Now.ToString("yyyyMMddhhmmss"));
            //byte[] fileBytes = SyncHttpHelper.GetHttpRespponseForFile(@"http://192.168.30.124:8989/Report/CreateReportAndDownLoad", reportFilesZip, 
            SyncHttpHelper.PostHttpResponse(@"http://192.168.30.124:8989/report/CreateReport", "");
            //int.Parse(DateTime.Now.ToString("hhmmss")));
            //if (fileBytes.Length <= 0) {
            //    throw new Exception("下载文件失败");
            //}
            //解压文件
            ZipFileHelper.DecompressionZip(reportFilesZip, createOutputDir);
            return "解压成功";
        }

        private string CreateHtmlFile(string htmlStr, string dirPath)
        {
            string dateStr = DateTime.Now.ToString("yyyyMMddhhmmss");
            string htmlpath = dirPath + "\\reportHtml" + dateStr + ".html";
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
