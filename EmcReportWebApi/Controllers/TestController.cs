using EmcReportWebApi.Business;
using EmcReportWebApi.Business.Implement;
using EmcReportWebApi.Common;
using EmcReportWebApi.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Http;

namespace EmcReportWebApi.Controllers
{
    public class TestController : ApiController
    {

        #region 测试数据
        //0592778d-05d6-463f-a5cc-9d53ed11ec0c.docx
        private string jsonStr1 = "{\"scbWord\":\"5b838ac0-8ea1-4b13-9322-138d32534845.docx\",\"yptp\":[{\"fileName\":\"6da7e28a-e492-4a8a-8ac3-2bb03e05b132.jpg\",\"content\":\"外观\"},{\"fileName\":\"75aebfd8-6184-4678-bce9-8e38baeb8090.jpg\",\"content\":\"铭牌\"}],\"firstPage\":{\"main_wtf\":\"飞利浦(中国)投资有限公司\",\"main_ypmc\":\"病人监护仪\",\"main_xhgg\":\"M8102A1\",\"main_jylb\":\"委托检验\",\"ypmc\":\"病人监护仪\",\"sb\":\"\",\"wtf\":\"飞利浦（中国）投资有限公司\",\"wtfdz\":\"上海市静安区灵石路718号A幢\",\"scdw\":\"PhilipsMedizinSystemeBoeblingenGmbH\",\"sjdw\":\"PhilipsMedizinSystemeBoeblingenGmbH\",\"cydw\":\"\",\"cydd\":\"\",\"cyrq\":\"\",\"dyrq\":\"2018年5月8日\",\"jyxm\":\"YY0505全项目、YY0601中36、YY0667中36、YY0668中36、YY0783中36、YY0784中36\",\"jyyj\":\"YY0505-2012《医用电气设备第1-2部分：安全通用要求并列标准电磁兼容要求和试验》、YY0601-2009《医用电气设备呼吸气体监护仪的基本安全和主要性能专用要求》、YY0667-2008《医用电气设备第2-30部分：自动循环无创血压监护设备的安全和基本性能专用要求》、YY0668-2008《医用电气设备第2-49部分：多参数患者监护设备安全专用要求》、YY0783-2010《医用电气设备第2-34部分：有创血压监测设备的安全和基本性能专用要求》、YY0784-2010《医用电气设备医用脉搏血氧仪设备基本安全和主要性能专用要求》\",\"jyjl\":\"被检样品符合YY0505-2012标准要求、符合YY0601-2009标准第36章要求、符合YY0667-2008标准第36章要求、符合YY0668-2008标准第36章要求、符合YY0783-2008标准第36章要求、符合YY0784-2010标准第36章要求\",\"bz\":\"报告中“/”表示此项空白，“—”表示不适用。\",\"ypbh\":\"QW2018-0698\",\"xhgg\":\"M8102A\",\"jylb\":\"委托检验\",\"cpbhph\":\"DE65528125\",\"cydbh\":\"\",\"scrq\":\"2018-02-16\",\"ypsl\":\"1台\",\"cyjs\":\"\",\"jydd\":\"本所实验室\",\"jyrq\":\"2018年5月22日~2018年7月13日\",\"jydd\":\"本所实验室\",\"ypms\":\"见本报告第3页“1受检样品信息”。\",\"xhgghqtsm\":\"1.检测结果不包括不确定度的估算值。2.ECG附件有63个型号：M1631A、M1671A、M1984A、M1611A、M1968A、M1625A、M1639A、M1675A、M1602A、M1974A、M1601A、M1635A、M1678A、M1976A、M1672A、M1673A、M1533A、M1971A、M1973A、M1684A、M1613A、M1681A、M1558A、M1609A、M1683A、M1621A、M1674A、M1604A、M1685A、M1603A、M1619A、M1669A、M1645A、M1510A、M1500A、M1520A、M1979A、M1530A、M1557A、M1644A、M1605A、M1680A、M1537A、M1647A、M1532A、M1978A、M1615A、M1633A、M1668A、M1629A、M1663A、M1667A、M1623A、M1538A、M1665A、M1682A、M1540C、M1550C、M1560C、M1570C、989803170171、989803170181、989803143201。其电气原理和材料组成完全一致,   仅导联数与长度有所区别。本次检测了M1663A，M1978A，M1971A。SpO2附件有5个型号：M1192A、M1193A、M1194A、M1195A、M1196A，其电气原理和材料组成完全一致，仅长度和适用人群有所区别。本次检测了M1196A。CO2附件有17个型号：M2516A、M2761A、M2772A、M2751A、M2750A、M2745A、M2756A、M2757A、M2501A、M2768A、M2773A、M2741A、M2536A、M2746A、M2776A、M2777A、M1920A。其产品结构及原理均相同。本次检测了M2741A。温度探头有11个型号：21075A、21076A、21078A、M1837A、21091A、21093A、21094A、21095A、21090A、21082A、21082B。其电气原理和材料组成完全一致，仅长度和适用范围有所区别，本次检测了M21075A。袖带（含连接管）共有8个型号：M1571A、M1572A、M1573A、M1574A、M1575A、M1576A、M1598B、M1599B。其电气原理及材料组成完全一致，仅围度和连接管长度有所区别。本次检测了M1598B和M1574A。\"},\"cssbList\":[{\"cssbxh\":\"1\",\"cssbbhxlh\":\"2-FW-11\",\"cssbmc\":\"测试接收机\",\"cssbzzs\":\"R&S\",\"cssbxhgg\":\"ESH2-Z5\",\"cssbxcjzrq\":\"2019.5.15\",\"cssbbz\":\" \"},{\"cssbxh\":\"2\",\"cssbbhxlh\":\"2-FW-12\",\"cssbmc\":\"人工电源网络\",\"cssbzzs\":\"R&S\",\"cssbxhgg\":\"ESCI\",\"cssbxcjzrq\":\"2019.5.15\",\"cssbbz\":\" \"},{\"cssbxh\":\"3\",\"cssbbhxlh\":\"2-FW-103\",\"cssbmc\":\"屏蔽室1\",\"cssbzzs\":\"ETS·LINDGREN\",\"cssbxhgg\":\"SR1\",\"cssbxcjzrq\":\"2019.4.14\",\"cssbbz\":\" \"},{\"cssbxh\":\"4\",\"cssbbhxlh\":\"2-FW-93\",\"cssbmc\":\"测试接收机\",\"cssbzzs\":\"R&S\",\"cssbxhgg\":\"ESU26\",\"cssbxcjzrq\":\"2019.5.15\",\"cssbbz\":\" \"},{\"cssbxh\":\"5\",\"cssbbhxlh\":\"2-FW-101\",\"cssbmc\":\"双锥复合对数周期天线\",\"cssbzzs\":\"SCHWARZBECK\",\"cssbxhgg\":\"VULB9163\",\"cssbxcjzrq\":\"2020.2.13\",\"cssbbz\":\" \"},{\"cssbxh\":\"6\",\"cssbbhxlh\":\"2-FW-102\",\"cssbmc\":\"10米法电波暗室\",\"cssbzzs\":\"ETS·LINDGREN\",\"cssbxhgg\":\"FACT10\",\"cssbxcjzrq\":\"2019.4.14\",\"cssbbz\":\" \"},{\"cssbxh\":\"7\",\"cssbbhxlh\":\"2-FW-102\",\"cssbmc\":\"控制室\",\"cssbzzs\":\"ETS·LINDGREN\",\"cssbxhgg\":\"CR\",\"cssbxcjzrq\":\"2019.4.14\",\"cssbbz\":\" \"},{\"cssbxh\":\"8\",\"cssbbhxlh\":\"2-FW-163\",\"cssbmc\":\"静电放电器\",\"cssbzzs\":\"EM TEST\",\"cssbxhgg\":\"dito\",\"cssbxcjzrq\":\"2019.1.22\",\"cssbbz\":\" \"},{\"cssbxh\":\"9\",\"cssbbhxlh\":\"2-FW-106\",\"cssbmc\":\"屏蔽室3\",\"cssbzzs\":\"ETS·LINDGREN\",\"cssbxhgg\":\"SR3\",\"cssbxcjzrq\":\"2019.4.14\",\"cssbbz\":\" \"},{\"cssbxh\":\"10\",\"cssbbhxlh\":\"2-FW-30\",\"cssbmc\":\"功率放大器\",\"cssbzzs\":\"BONN\",\"cssbxhgg\":\"BLWA0830-160/100/40D\",\"cssbxcjzrq\":\"\",\"cssbbz\":\" \"},{\"cssbxh\":\"11\",\"cssbbhxlh\":\"2-FW-34\",\"cssbmc\":\"场强表\",\"cssbzzs\":\"AR\",\"cssbxhgg\":\"FL7006/Kit M1\",\"cssbxcjzrq\":\"2019.9.11\",\"cssbbz\":\" \"},{\"cssbxh\":\"12\",\"cssbbhxlh\":\"2-FW-100\",\"cssbmc\":\"信号发生器\",\"cssbzzs\":\"R&S\",\"cssbxhgg\":\"SMB100A-B106\",\"cssbxcjzrq\":\"2019.5.15\",\"cssbbz\":\" \"}],\"experiment\":[{\"name\":\"传导发射实验\",\"syjg\":\"符合\",\"jyrq\":\"2020-01-01\",\"wd\":\"100\",\"xdsd\":\"65\",\"dqyl\":\"100\",\"sysj\":[{\"sygdy\":\"AC220V 50Hz\",\"syplfw\":\"0.15MHz~30MHz\",\"ypyxms\":\"① \",\"mccfpl\":\"AC220V 60Hz\",\"rtf\":[{\"name\":\"54b5f48c-ae87-4321-ba2a-1fa50c4e4411.Rtf\"},{\"name\":\"9d8a5174-1159-4db9-9f12-d4fb051c735c.Rtf\"}]},{\"sygdy\":\"AC220V 50Hz\",\"syplfw\":\"0.15MHz~30MHz\",\"ypyxms\":\"② \",\"rtf\":[{\"name\":\"093dd9de-e1c9-4c1d-952e-cd3f372ae14b.Rtf\"},{\"name\":\"72a42c7c-66fc-4395-a13a-1b9191b1b8c5.Rtf\"}]}],\"syljt\":[{\"name\":\"78b4cd30-4255-4051-948a-45da4bdeb7a2.jpg\",content:\"\"}],\"sybzt\":[{\"name\":\"6da7e28a-e492-4a8a-8ac3-2bb03e05b132.jpg\",content:\"①③\"},{\"name\":\"c3030f9b-7cc6-4a74-910a-54c415364694.jpg\",content:\"②④\"}]},{\"name\":\"辐射发射试验\",\"syjg\":\"符合\",\"jyrq\":\"2020-01-01\",\"wd\":\"100\",\"xdsd\":\"65\",\"dqyl\":\"100\",\"sysj\":[{\"sygdy\":\"AC220V 50Hz\",\"syplfw\":\"0.15MHz~30MHz\",\"ypyxms\":\"① \",\"rtf\":[{\"name\":\"857dd599-e88a-423a-a158-80befbbd4506.Rtf\"}]},{\"sygdy\":\"AC220V 50Hz\",\"syplfw\":\"0.15MHz~30MHz\",\"ypyxms\":\"② \",\"rtf\":[{\"name\":\"857dd599-e88a-423a-a158-80befbbd4506.Rtf\"}]}],\"syljt\":[{\"name\":\"98339fa7-3cf5-4434-aa13-a0eb1848099e.jpg\",content:\"\"}],\"sybzt\":[{\"name\":\"c3030f9b-7cc6-4a74-910a-54c415364694.jpg\",content:\"①③\"},{\"name\":\"c3030f9b-7cc6-4a74-910a-54c415364694.jpg\",content:\"②④\"}]},{\"name\":\"谐波失真\",\"syjg\":\"符合\",\"jyrq\":\"2020-01-01\",\"wd\":\"100\",\"xdsd\":\"65\",\"dqyl\":\"100\",\"sysj\":[{\"sygdy\":\"AC220V 50Hz\",\"syplfw\":\"0.15MHz~30MHz\",\"ypyxms\":\"① \",\"rtf\":[{\"name\":\"3eef059b-bf10-4ce9-9129-d9c992d01bd9.rtf\"}]}],\"syljt\":[{\"name\":\"98339fa7-3cf5-4434-aa13-a0eb1848099e.jpg\",content:\"\"}],\"sybzt\":[{\"name\":\"c3030f9b-7cc6-4a74-910a-54c415364694.jpg\",content:\"①③\"},{\"name\":\"c3030f9b-7cc6-4a74-910a-54c415364694.jpg\",content:\"②④\"}]},{\"name\":\"电压波动和闪烁\",\"syjg\":\"符合\",\"jyrq\":\"2020-01-01\",\"wd\":\"100\",\"xdsd\":\"65\",\"dqyl\":\"100\",\"sysj\":[{\"sygdy\":\"AC220V 50Hz\",\"syplfw\":\"0.15MHz~30MHz\",\"ypyxms\":\"① \",\"rtf\":[{\"name\":\"6f55f711-cb0f-49a3-8b29-0c463fd15c2d.rtf\"}]}],\"syljt\":[{\"name\":\"98339fa7-3cf5-4434-aa13-a0eb1848099e.jpg\",content:\"\"}],\"sybzt\":[{\"name\":\"c3030f9b-7cc6-4a74-910a-54c415364694.jpg\",content:\"①③\"},{\"name\":\"c3030f9b-7cc6-4a74-910a-54c415364694.jpg\",content:\"②④\"}]},{\"name\":\"静电放电\",\"syjg\":\"符合\",\"jyrq\":\"2020-01-01\",\"wd\":\"100\",\"xdsd\":\"65\",\"dqyl\":\"100\",\"sysj\":[{\"sygdy\":\"AC220V 50Hz\",\"syplfw\":\"0.15MHz~30MHz\",\"ypyxms\":\"① \",\"html\":[{\"table\":\"<html> <head>     <meta charset='utf-8'> 	<style> 		table,th,td{ 			border:1px solid #000000; 			border-spacing:0; 			border-collapse:collapse; 		} 		.center{ 			text-align:center; 		} 	</style> </head> <body> <table class='custom-table white' id='test_electrostatic_data' menu='true' style='width:100%;'><tbody id='test_electrostatic_data'> <tr class='sample-line'>     <th colspan='11' rowspan='1' class='whole-line'>试验数据     </th> </tr> <tr class='sample-line'>     <td colspan='3' rowspan='1' class='table-label'>放电电压（kV）</td>     <td class='center'>+2</td>     <td class='center'>-2</td>     <td class='center'>+4</td>     <td class='center'>-4</td>     <td class='center'>+6</td>     <td class='center'>-6</td>     <td class='center'>+8</td>     <td class='center'>-8</td>      </tr> <tr class='sample-line'>     <td colspan='2' rowspan='3' deepth='1' class='right-click table-label'>空气放电</td>     <td deepth='2' colspan='1' rowspan='1'>test</td>     <td deepth='3' colspan='1' rowspan='1' class='center value-switching' >√</td>     <td deepth='4' colspan='1' rowspan='1' class='center value-switching' >√</td>     <td deepth='5' colspan='1' rowspan='1' class='center value-switching' >√</td>     <td deepth='6' colspan='1' rowspan='1' class='center value-switching' >√</td>     <td deepth='7' colspan='1' rowspan='1' class='center value-switching' >√</td>     <td deepth='8' colspan='1' rowspan='1' class='center value-switching' >√</td>     <td deepth='9' colspan='1' rowspan='1' class='center value-switching' >√</td>     <td deepth='10' colspan='1' rowspan='1' class='center value-switching'>√</td>      </tr> <tr class='sample-line'>     <td deepth='2' colspan='1' rowspan='1'>test</td>     <td deepth='3' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='4' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='5' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='6' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='7' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='8' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='9' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='10' colspan='1' rowspan='1' class='center value-switching'>√</td>      </tr> <tr class='sample-line'>     <td deepth='2' colspan='1' rowspan='1'>test</td>     <td deepth='3' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='4' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='5' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='6' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='7' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='8' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='9' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='10' colspan='1' rowspan='1' class='center value-switching'>√</td>      </tr> <tr class='sample-line'>     <td colspan='1' rowspan='4' deepth='1' class='table-label'>接触放电</td>     <td colspan='1' rowspan='2' deepth='2' class='right-click table-label'>直接</td>     <td deepth='3' colspan='1' rowspan='1'>test</td>     <td deepth='4' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='5' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='6' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='7' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='8' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='9' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='10' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='11' colspan='1' rowspan='1' class='center value-switching'>√</td>      </tr> <tr class='sample-line'>     <td deepth='3' colspan='1' rowspan='1'>test</td>     <td deepth='4' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='5' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='6' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='7' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='8' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='9' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='10' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='11' colspan='1' rowspan='1' class='center value-switching'>√</td>      </tr> <tr class='sample-line'>     <td colspan='1' rowspan='2' deepth='2' class='right-click table-label'>间接</td>     <td deepth='3' colspan='1' rowspan='1'>HCP</td>     <td deepth='4' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='5' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='6' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='7' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='8' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='9' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='10' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='11' colspan='1' rowspan='1' class='center value-switching'>√</td>      </tr> <tr class='sample-line'>     <td deepth='3' colspan='1' rowspan='1'>VCP</td>     <td deepth='4' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='5' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='6' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='7' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='8' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='9' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='10' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='11' colspan='1' rowspan='1' class='center value-switching'>√</td>      </tr> <tr class='sample-line'>     <td colspan='3' rowspan='1' class='table-label'>异常现象描述</td>     <td colspan='8' rowspan='1'>这是一个异常现象</td> </tr> <tr class='sample-line'>     <td colspan='3' rowspan='1' class='table-label'>备注</td>     <td colspan='8' rowspan='1'>这是一条备注</td> </tr> </tbody>  </table></body></html>\"}]}],\"syljt\":[{\"name\":\"98339fa7-3cf5-4434-aa13-a0eb1848099e.jpg\",content:\"\"}],\"sybzt\":[{\"name\":\"c3030f9b-7cc6-4a74-910a-54c415364694.jpg\",content:\"①③\"},{\"name\":\"c3030f9b-7cc6-4a74-910a-54c415364694.jpg\",content:\"②④\"}]},{\"name\":\"电压暂降/短时中断\",\"syjg\":\"null\",\"jyrq\":\"2020-03-09 16:47:38\",\"wd\":\"null\",\"xdsd\":\"null\",\"dqyl\":\"null\",\"sysj\":[{\"sysjTitle\":\"电压暂降\",\"sygdy\":\"null\",\"ypyxms\":\"null\",\"html\":[{\"table\":\"<html> <head>     <meta charset='utf-8'> 	<style> 		table,th,td{ 			border:1px solid #000000; 			border-spacing:0; 			border-collapse:collapse; 		} 		.center{ 			text-align:center; 		} 	</style> </head> <body> <table class='custom-table white' id='test_electrostatic_data' menu='true' style='width:100%;'><tbody id='test_electrostatic_data'> <tr class='sample-line'>     <th colspan='11' rowspan='1' class='whole-line'>试验数据     </th> </tr> <tr class='sample-line'>     <td colspan='3' rowspan='1' class='table-label'>放电电压（kV）</td>     <td class='center'>+2</td>     <td class='center'>-2</td>     <td class='center'>+4</td>     <td class='center'>-4</td>     <td class='center'>+6</td>     <td class='center'>-6</td>     <td class='center'>+8</td>     <td class='center'>-8</td>      </tr> <tr class='sample-line'>     <td colspan='2' rowspan='3' deepth='1' class='right-click table-label'>空气放电</td>     <td deepth='2' colspan='1' rowspan='1'>test</td>     <td deepth='3' colspan='1' rowspan='1' class='center value-switching' >√</td>     <td deepth='4' colspan='1' rowspan='1' class='center value-switching' >√</td>     <td deepth='5' colspan='1' rowspan='1' class='center value-switching' >√</td>     <td deepth='6' colspan='1' rowspan='1' class='center value-switching' >√</td>     <td deepth='7' colspan='1' rowspan='1' class='center value-switching' >√</td>     <td deepth='8' colspan='1' rowspan='1' class='center value-switching' >√</td>     <td deepth='9' colspan='1' rowspan='1' class='center value-switching' >√</td>     <td deepth='10' colspan='1' rowspan='1' class='center value-switching'>√</td>      </tr> <tr class='sample-line'>     <td deepth='2' colspan='1' rowspan='1'>test</td>     <td deepth='3' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='4' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='5' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='6' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='7' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='8' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='9' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='10' colspan='1' rowspan='1' class='center value-switching'>√</td>      </tr> <tr class='sample-line'>     <td deepth='2' colspan='1' rowspan='1'>test</td>     <td deepth='3' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='4' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='5' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='6' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='7' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='8' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='9' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='10' colspan='1' rowspan='1' class='center value-switching'>√</td>      </tr> <tr class='sample-line'>     <td colspan='1' rowspan='4' deepth='1' class='table-label'>接触放电</td>     <td colspan='1' rowspan='2' deepth='2' class='right-click table-label'>直接</td>     <td deepth='3' colspan='1' rowspan='1'>test</td>     <td deepth='4' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='5' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='6' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='7' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='8' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='9' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='10' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='11' colspan='1' rowspan='1' class='center value-switching'>√</td>      </tr> <tr class='sample-line'>     <td deepth='3' colspan='1' rowspan='1'>test</td>     <td deepth='4' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='5' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='6' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='7' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='8' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='9' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='10' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='11' colspan='1' rowspan='1' class='center value-switching'>√</td>      </tr> <tr class='sample-line'>     <td colspan='1' rowspan='2' deepth='2' class='right-click table-label'>间接</td>     <td deepth='3' colspan='1' rowspan='1'>HCP</td>     <td deepth='4' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='5' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='6' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='7' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='8' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='9' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='10' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='11' colspan='1' rowspan='1' class='center value-switching'>√</td>      </tr> <tr class='sample-line'>     <td deepth='3' colspan='1' rowspan='1'>VCP</td>     <td deepth='4' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='5' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='6' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='7' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='8' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='9' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='10' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='11' colspan='1' rowspan='1' class='center value-switching'>√</td>      </tr> <tr class='sample-line'>     <td colspan='3' rowspan='1' class='table-label'>异常现象描述</td>     <td colspan='8' rowspan='1'>这是一个异常现象</td> </tr> <tr class='sample-line'>     <td colspan='3' rowspan='1' class='table-label'>备注</td>     <td colspan='8' rowspan='1'>这是一条备注</td> </tr> </tbody>  </table></body></html>\"}]},{\"sysjTitle\":\"短时中断\",\"sygdy\":\"null\",\"ypyxms\":\"null\",\"html\":[{\"table\":\"<html> <head>     <meta charset='utf-8'> 	<style> 		table,th,td{ 			border:1px solid #000000; 			border-spacing:0; 			border-collapse:collapse; 		} 		.center{ 			text-align:center; 		} 	</style> </head> <body> <table class='custom-table white' id='test_electrostatic_data' menu='true' style='width:100%;'><tbody id='test_electrostatic_data'> <tr class='sample-line'>     <th colspan='11' rowspan='1' class='whole-line'>试验数据     </th> </tr> <tr class='sample-line'>     <td colspan='3' rowspan='1' class='table-label'>放电电压（kV）</td>     <td class='center'>+2</td>     <td class='center'>-2</td>     <td class='center'>+4</td>     <td class='center'>-4</td>     <td class='center'>+6</td>     <td class='center'>-6</td>     <td class='center'>+8</td>     <td class='center'>-8</td>      </tr> <tr class='sample-line'>     <td colspan='2' rowspan='3' deepth='1' class='right-click table-label'>空气放电</td>     <td deepth='2' colspan='1' rowspan='1'>test</td>     <td deepth='3' colspan='1' rowspan='1' class='center value-switching' >√</td>     <td deepth='4' colspan='1' rowspan='1' class='center value-switching' >√</td>     <td deepth='5' colspan='1' rowspan='1' class='center value-switching' >√</td>     <td deepth='6' colspan='1' rowspan='1' class='center value-switching' >√</td>     <td deepth='7' colspan='1' rowspan='1' class='center value-switching' >√</td>     <td deepth='8' colspan='1' rowspan='1' class='center value-switching' >√</td>     <td deepth='9' colspan='1' rowspan='1' class='center value-switching' >√</td>     <td deepth='10' colspan='1' rowspan='1' class='center value-switching'>√</td>      </tr> <tr class='sample-line'>     <td deepth='2' colspan='1' rowspan='1'>test</td>     <td deepth='3' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='4' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='5' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='6' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='7' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='8' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='9' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='10' colspan='1' rowspan='1' class='center value-switching'>√</td>      </tr> <tr class='sample-line'>     <td deepth='2' colspan='1' rowspan='1'>test</td>     <td deepth='3' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='4' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='5' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='6' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='7' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='8' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='9' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='10' colspan='1' rowspan='1' class='center value-switching'>√</td>      </tr> <tr class='sample-line'>     <td colspan='1' rowspan='4' deepth='1' class='table-label'>接触放电</td>     <td colspan='1' rowspan='2' deepth='2' class='right-click table-label'>直接</td>     <td deepth='3' colspan='1' rowspan='1'>test</td>     <td deepth='4' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='5' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='6' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='7' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='8' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='9' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='10' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='11' colspan='1' rowspan='1' class='center value-switching'>√</td>      </tr> <tr class='sample-line'>     <td deepth='3' colspan='1' rowspan='1'>test</td>     <td deepth='4' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='5' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='6' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='7' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='8' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='9' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='10' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='11' colspan='1' rowspan='1' class='center value-switching'>√</td>      </tr> <tr class='sample-line'>     <td colspan='1' rowspan='2' deepth='2' class='right-click table-label'>间接</td>     <td deepth='3' colspan='1' rowspan='1'>HCP</td>     <td deepth='4' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='5' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='6' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='7' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='8' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='9' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='10' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='11' colspan='1' rowspan='1' class='center value-switching'>√</td>      </tr> <tr class='sample-line'>     <td deepth='3' colspan='1' rowspan='1'>VCP</td>     <td deepth='4' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='5' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='6' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='7' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='8' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='9' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='10' colspan='1' rowspan='1' class='center value-switching'>√</td>     <td deepth='11' colspan='1' rowspan='1' class='center value-switching'>√</td>      </tr> <tr class='sample-line'>     <td colspan='3' rowspan='1' class='table-label'>异常现象描述</td>     <td colspan='8' rowspan='1'>这是一个异常现象</td> </tr> <tr class='sample-line'>     <td colspan='3' rowspan='1' class='table-label'>备注</td>     <td colspan='8' rowspan='1'>这是一条备注</td> </tr> </tbody>  </table></body></html>\"}]}],\"syljt\":[{\"name\":\"98339fa7-3cf5-4434-aa13-a0eb1848099e.jpg\",content:\"\"}],\"sybzt\":[{\"name\":\"c3030f9b-7cc6-4a74-910a-54c415364694.jpg\",content:\"①③\"},{\"name\":\"c3030f9b-7cc6-4a74-910a-54c415364694.jpg\",content:\"②④\"}]}]}";


        private string jsonStr = "{\"yptp\":[{\"fileName\":\"6da7e28a-e492-4a8a-8ac3-2bb03e05b132.jpg\",\"content\":\"外观\"},{\"fileName\":\"75aebfd8-6184-4678-bce9-8e38baeb8090.jpg\",\"content\":\"铭牌\"},{\"fileName\":\"6da7e28a-e492-4a8a-8ac3-2bb03e05b132.jpg\",\"content\":\"外观\"},{\"fileName\":\"75aebfd8-6184-4678-bce9-8e38baeb8090.jpg\",\"content\":\"铭牌\"}],\"attach\":[{\"col1\":\"报警状态\",\"col2\":\"报警类型\",\"col3\":\"指示灯颜色\",\"col4\":\"检验结果\",\"col5\":\"闪烁频率\",\"col6\":\"检验结果\",\"col7\":\"占空比\",\"col8\":\"检验结果\"},{\"col1\":\"断电报警\",\"col2\":\"高优先级\",\"col3\":\"红色\",\"col4-input\":\"红色\",\"col5\":\"1.4Hz~2.8Hz\",\"col6-input\":\"\",\"col7\":\"20%~60%\",\"col8-input\":\"\"},{\"col1\":\"肤温传感器断开\",\"col2\":\"高优先级\",\"col3\":\"红色\",\"col4-input\":\"红色\",\"col5\":\"1.4Hz~2.8Hz\",\"col6-input\":\"\",\"col7\":\"20%~60%\",\"col8-input\":\"\"}],\"firstPage\":{\"main_wtf\":\"国家药品监督管理局\",\"main_ypmc\":\"按产品标识\",\"main_xhgg\":\"按产品标识\",\"main_jylb\":\"2020年国家医疗器械抽检\",\"ypmc\":\"按产品标识\",\"sb\":\"\",\"wtf\":\"国家药品监督管理局\",\"wtfdz\":\"北京市西城区展览路北露园1号\",\"scdw\":\"按产品标识\",\"sjdw\":\"按抽样单上公章\",\"cydw\":\"按抽样单上公章\",\"cydd\":\"按抽样单上公章\",\"cyrq\":\"2020年*月*日\",\"dyrq\":\"2020年*月*日\",\"jyxm\":\"药监综械管〔2020〕*号文附件*《2020年国家医疗器械抽检(中央补助地方项目)产品检验方案》中“30200.无创自动测量血压计（电子血压计）”的检验项目\",\"jyyj\":\"药监综械管〔2020〕*号文附件*《2020年国家医疗器械抽检(中央补助地方项目)产品检验方案》中“30200.无创自动测量血压计（电子血压计）”的检验依据\",\"jyjl\":\"合格/不合格\",\"bz\":\"1）报告中的“——”表示此项不适用，报告中“/”表示此项空白。\",\"ypbh\":\"GYJ2020-****\",\"xhgg\":\"按产品标识\",\"jylb\":\"2020年国家医疗器械抽检\",\"cpbhph\":\"按产品标识形式、内容\",\"cydbh\":\"按产品标识形式、内容\",\"scrq\":\"按产品标识形式、内容\",\"ypsl\":\"按产品标识形式、内容\",\"cyjs\":\"按产品标识形式、内容\",\"jydd\":\"本所实验室\",\"jyrq\":\"2018年5月22日~2018年7月13日\",\"jydd\":\"本所实验室\",\"ypms\":\"1、被检样品封样完好。\r\n 2、本次检测包含下列部件：主机、袖带（根据实际情况填写）。 \",\"xhgghqtsm\":\"检测结果不包括不确定度的估算值。\"},\"standard\":[{\"itemId\":\"1\",\"idxNo\":\"8\",\"itemContent\":\"自动复位装置的选择\",\"itemPath\":\"1|\",\"comment\":\"12312312\",\"reMark\":\"44444\",\"list\":[{\"itemId\":\"2\",\"stdName\":\"GB9706.1-2007\",\"stdItmNo\":\"49\",\"itemPath\":\"1|2|\",\"list\":[{\"itemId\":\"3\",\"stdName\":\"GB9706.1-2007\",\"itemContent\":\"如果使用自动复位热断路器和过电流释放器,自动复位能保证安全.\",\"itemPath\":\"1|2|3|\",\"reference\":\"1\",\"list\":[]}]}]},{\"itemId\":\"4\",\"idxNo\":\"9\",\"itemContent\":\"电源中断后的复位\",\"itemPath\":\"4|\",\"comment\":\"12312312\",\"reMark\":\"44444\",\"list\":[{\"itemId\":\"5\",\"stdName\":\"GB9706.1-2007\",\"stdItmNo\":\"49\",\"itemPath\":\"4|5|\",\"list\":[{\"itemId\":\"6\",\"stdName\":\"GB9706.1-2007\",\"itemContent\":\"设备应设计成当电源供电中断后又恢复时,除预定功能中断外,不会发生安全方面危险.\",\"itemPath\":\"4|5|6|\",\"reference\":\"2\",\"list\":[]}]}]},{\"itemId\":\"7\",\"idxNo\":\"10\",\"itemContent\":\"指示器\",\"itemPath\":\"7|\",\"list\":[{\"itemId\":\"8\",\"stdName\":\"GB9706.1-2007\",\"stdItmNo\":\"57\",\"itemPath\":\"7|8|\",\"list\":[{\"itemId\":\"9\",\"stdName\":\"GB9706.1-2007\",\"itemContent\":\"除非对位于正常操作位置的操作者另有显而易见的指标,否则应安装指示灯,用于:\n----- 指示设备已通电.\",\"itemPath\":\"7|8|9|\",\"reference\":\"3\",\"list\":[]},{\"itemId\":\"10\",\"stdName\":\"GB9706.1-2007\",\"itemContent\":\"----设备装有不发光的电热器如会产生安全方面危险时,指示电热器已工作.\",\"itemPath\":\"7|8|10|\",\"reference\":\"4\",\"list\":[]},{\"itemId\":\"11\",\"stdName\":\"GB9706.1-2007\",\"itemContent\":\"----当输出电路意外的或长时间的工作可能引起安全方面危险时,指示处于输出状态.\",\"itemPath\":\"7|8|11|\",\"reference\":\"5\",\"list\":[]},{\"itemId\":\"12\",\"stdName\":\"GB9706.1-2007\",\"itemContent\":\"指示充电装置工作状态.\",\"itemPath\":\"7|8|12|\",\"reference\":\"6\",\"list\":[]}]}]},{\"itemId\":\"13\",\"idxNo\":\"7\",\"itemContent\":\"连续漏电流和患者辅助电流(正常工作温度下)\",\"itemPath\":\"13|\",\"list\":[{\"itemId\":\"14\",\"stdName\":\"GB9706.1-2007\",\"stdItmNo\":\"19\",\"itemPath\":\"13|14|\",\"list\":[{\"itemId\":\"15\",\"stdName\":\"GB9706.1-2007\",\"itemContent\":\"患者漏电流\",\"itemPath\":\"13|14|15|\",\"reference\":\"7\",\"list\":[{\"itemId\":\"16\",\"stdName\":\"GB9706.1-2007\",\"itemContent\":\"直流\",\"itemPath\":\"13|14|15|16|\",\"reference\":\"7\",\"list\":[{\"itemId\":\"17\",\"stdName\":\"GB9706.1-2007\",\"itemContent\":\"正常状态下≤0.01mA\",\"itemPath\":\"13|14|15|16|17|\",\"reference\":\"7\",\"list\":[],\"result\":\"测试检验结果1\"},{\"itemId\":\"18\",\"stdName\":\"GB9706.1-2007\",\"itemContent\":\"单一故障状态下≤0.05mA\",\"itemPath\":\"13|14|15|16|18|\",\"reference\":\"8\",\"result\":\"测试检验结果1\",\"list\":[]}]},{\"itemId\":\"19\",\"stdName\":\"GB9706.1-2007\",\"itemContent\":\"交流\",\"itemPath\":\"13|14|15|19|\",\"reference\":\"9\",\"list\":[{\"itemId\":\"20\",\"stdName\":\"GB9706.1-2007\",\"itemContent\":\"正常状态下≤0.1mA\",\"itemPath\":\"13|14|15|19|20|\",\"reference\":\"9\",\"list\":[]},{\"itemId\":\"21\",\"stdName\":\"GB9706.1-2007\",\"itemContent\":\"单一故障状态下≤0.5mA\",\"itemPath\":\"13|14|15|19|21|\",\"reference\":\"10\",\"list\":[]}]},{\"itemId\":\"22\",\"stdName\":\"GB9706.1-2007\",\"itemContent\":\"应用部分加压状态≤5mA\",\"itemPath\":\"13|14|15|22|\",\"reference\":\"11\",\"list\":[]},{\"itemId\":\"23\",\"stdName\":\"GB9706.1-2007\",\"itemContent\":\"信号输入/出部分加压状态≤{$1}mA\",\"itemPath\":\"13|14|15|23|\",\"reference\":\"12\",\"list\":[]}]},{\"itemId\":\"24\",\"stdName\":\"GB9706.1-2007\",\"itemContent\":\"患者辅助电流\n单位: mA\",\"itemPath\":\"13|14|24|\",\"reference\":\"13\",\"list\":[{\"itemId\":\"25\",\"stdName\":\"GB9706.1-2007\",\"itemContent\":\"直流\",\"itemPath\":\"13|14|24|25|\",\"reference\":\"13\",\"list\":[{\"itemId\":\"26\",\"stdName\":\"GB9706.1-2007\",\"itemContent\":\"正常状态下≤0.01mA\",\"itemPath\":\"13|14|24|25|26|\",\"reference\":\"13\",\"list\":[]},{\"itemId\":\"27\",\"stdName\":\"GB9706.1-2007\",\"itemContent\":\"单一故障状态下≤0.05mA\",\"itemPath\":\"13|14|24|25|27|\",\"reference\":\"14\",\"list\":[]}]},{\"itemId\":\"28\",\"stdName\":\"GB9706.1-2007\",\"itemContent\":\"交流\",\"itemPath\":\"13|14|24|28|\",\"reference\":\"15\",\"list\":[{\"itemId\":\"29\",\"stdName\":\"GB9706.1-2007\",\"itemContent\":\"正常状态下≤0.1mA\",\"itemPath\":\"13|14|24|28|29|\",\"reference\":\"15\",\"list\":[]},{\"itemId\":\"30\",\"stdName\":\"GB9706.1-2007\",\"itemContent\":\"单一故障状态下≤0.5mA\",\"itemPath\":\"13|14|24|28|30|\",\"reference\":\"16\",\"list\":[]}]}]}]}]}]}";

        private IReport report = new ReportImpl();

        #endregion

        private IReport _report;
        private IReportStandard _reportStandard;

        public TestController(IReport report, IReportStandard reportStandard) {
            _report = report;
            _reportStandard = reportStandard;
        }

        [HttpPost]
        [CompressContentAttribute]
        public IHttpActionResult CreateReportTest1(ReportParams para)
        {
            Task<ReportResult<string>> task = new Task<ReportResult<string>>(()=>CreateReportTestAsync(para));
            task.Start();
            ReportResult<string> result = task.Result;
            return Json<ReportResult<string>>(result);
        }

        private ReportResult<string> CreateReportTestAsync(ReportParams para) {
            ReportResult<string> result = new ReportResult<string>();
            try
            {
                EmcConfig.SemLim.Wait();
                //string jsonStr = para.JsonStr;
                string reportId = para.ReportId;
                Stopwatch sw = new Stopwatch();
                sw.Start();
                //获取zip文件 
                string reportFilesPath = FileUtil.CreateReportDirectory(string.Format("{0}\\Files\\ReportFiles", EmcConfig.CurrRoot));
                string reportZipFilesPath = string.Format("{0}Files\\ReportFiles\\Test\\{1}", EmcConfig.CurrRoot, "QT2019-3015.zip");
                //解压zip文件
                ZipFileHelper.DecompressionZip(reportZipFilesPath, reportFilesPath);

                //生成报告
                string content = report.JsonToWord(reportId.Equals("") ? "QW2018-698" : reportId, para.JsonStr.Equals("") ? jsonStr1 : para.JsonStr, reportFilesPath);
                sw.Stop();
                double time1 = (double)sw.ElapsedMilliseconds / 1000;
                result.Message = string.Format("报告生成成功,用时:" + time1.ToString());
                result.SumbitResult = true;
                result.Content = content;
                EmcConfig.InfoLog.Info("报告:" + result.Content + ",信息:" + result.Message);

            }
            catch (Exception ex)
            {
                EmcConfig.ErrorLog.Error(ex.Message, ex);
                throw ex;
            }
            finally {
                EmcConfig.SemLim.Release();
            }
            return result;
        }

        /// <summary>
        /// 测试报告
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public string Test()
        {
            Stopwatch sw = new Stopwatch();

            sw.Start();
            EmcConfig.KillWordProcess();

            string reportFilesPath = FileUtil.CreateReportDirectory(string.Format("{0}\\Files\\ReportFiles", EmcConfig.CurrRoot));
            string reportZipFilesPath = string.Format("{0}Files\\ReportFiles\\Test\\{1}", EmcConfig.CurrRoot, "QT2019-3015.zip");
            //解压zip文件
            ZipFileHelper.DecompressionZip(reportZipFilesPath, reportFilesPath);

            string result = report.JsonToWord("QT2019-3015", jsonStr1, reportFilesPath);
            //string result = "";
            sw.Stop();
            double time1 = (double)sw.ElapsedMilliseconds / 1000;
            return result + ":" + time1.ToString();
        }

        /// <summary>
        /// 新测试标准
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public string Test2()
        {
            Stopwatch sw = new Stopwatch();

            sw.Start();
            EmcConfig.KillWordProcess();

            string reportFilesPath = FileUtil.CreateReportDirectory(string.Format("{0}Files\\ReportFiles", EmcConfig.CurrRoot));
            string reportZipFilesPath = string.Format("{0}Files\\ReportFiles\\Test\\{1}", EmcConfig.CurrRoot, "QT2019-3015.zip");
            //解压zip文件
            ZipFileHelper.DecompressionZip(reportZipFilesPath, reportFilesPath);

            string result = _reportStandard.JsonToWordStandardNew("QT2019-3015", "2c908aa86bfa3a96016bfa3d872a0002", reportFilesPath);
            //string result = "";
            sw.Stop();
            double time1 = (double)sw.ElapsedMilliseconds / 1000;
            return result + ":" + time1.ToString();
        }

        /// <summary>
        /// 测试标准
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public string Test3()
        {
            Stopwatch sw = new Stopwatch();

            sw.Start();
            EmcConfig.KillWordProcess();

            string reportFilesPath = FileUtil.CreateReportDirectory(string.Format("{0}Files\\ReportFiles", EmcConfig.CurrRoot));
            string reportZipFilesPath = string.Format("{0}Files\\ReportFiles\\Test\\{1}", EmcConfig.CurrRoot, "QT2019-3015.zip");
            //解压zip文件
            ZipFileHelper.DecompressionZip(reportZipFilesPath, reportFilesPath);
           
            string result = _reportStandard.JsonToWordStandard("QT2019-3015", jsonStr, reportFilesPath);
            //string result = "";
            sw.Stop();
            double time1 = (double)sw.ElapsedMilliseconds / 1000;
            return result + ":" + time1.ToString();
        }

        /// <summary>
        /// 测试异步
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public string Test4() {

            //for (int i = 0; i < 20; i++)
            //{
            //    EmcConfig.TaskQueue.Enqueue(Guid.NewGuid());
            //}

            string result = "";
            for (int i = 0; i < 10; i++)
            {
                Task<string> task = ReportTestAsync();
                result += task.Result;
            }

            //Task task = new Task(TestTask);
            //task.Start();
            return result;
        }

        private SemaphoreSlim semLim = new SemaphoreSlim(2);

        private async Task<string> ReportTestAsync() {
            var task = TestTask();
            Task.WaitAll(task);
            string result = await task;
            return result;
        }

        private Task<string> TestTask() {
            try
            {
                 
                return Task<string>.Run(() =>
                {
                    semLim.Wait();
                    string reportFilesPath = FileUtil.CreateReportDirectory(string.Format("{0}Files\\ReportFiles", EmcConfig.CurrRoot));
                    string reportZipFilesPath = string.Format("{0}Files\\ReportFiles\\Test\\{1}", EmcConfig.CurrRoot, "QT2019-3015.zip");
                    //解压zip文件
                    ZipFileHelper.DecompressionZip(reportZipFilesPath, reportFilesPath);

                    string result = _reportStandard.JsonToWordStandard("QT2019-3015", jsonStr, reportFilesPath);
                    semLim.Release();
                    return result;
                });
               // EmcConfig.KillWordProcess();

            }
            catch (Exception ex)
            {

                throw;
            }


        }
        
    }
}
