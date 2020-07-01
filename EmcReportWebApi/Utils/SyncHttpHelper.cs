using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;

namespace EmcReportWebApi.Utils
{
    /// <summary>
    /// httpClient帮助
    /// </summary>
    public class SyncHttpHelper
    {
        /// <summary>
        /// get请求下载文件
        /// </summary>
        /// <param name="url"></param>
        /// <param name="outFilePath"></param>
        /// <param name="Timeout"></param>
        public static byte[] GetHttpRespponseForFile(string url, string outFilePath, int Timeout)
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            request.Method = "GET";
            request.ContentType = "text/html;charset=UTF-8";
            request.UserAgent = null;
            request.Timeout = Timeout;

            byte[] fileBytes;
            try
            {
                using (WebResponse webRes = request.GetResponse())
                {
                    int length = (int)webRes.ContentLength;
                    HttpWebResponse response = webRes as HttpWebResponse;
                    Stream stream = response.GetResponseStream();

                    //读取到内存
                    MemoryStream stmMemory = new MemoryStream();
                    byte[] buffer = new byte[length];
                    int i;
                    //将字节逐个放入到Byte中
                    while ((i = stream.Read(buffer, 0, buffer.Length)) > 0)
                    {
                        stmMemory.Write(buffer, 0, i);
                    }
                    fileBytes = stmMemory.ToArray();//文件流Byte，需要文件流可直接return，不需要下面的保存代码
                    stmMemory.Close();

                    MemoryStream m = new MemoryStream(fileBytes);
                    FileStream fs = new FileStream(outFilePath, FileMode.OpenOrCreate);
                    m.WriteTo(fs);
                    m.Close();
                    fs.Close();
                }
            }
            catch (System.Exception ex)
            {
                throw ex;
            }
            return fileBytes;
        }

        /// Get请求
        /// 字符串
        public static string GetHttpResponse(string url, int Timeout)
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            request.Method = "GET";
            request.ContentType = "text/html;charset=UTF-8";
            request.UserAgent = null;
            request.Timeout = Timeout;

            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            Stream myResponseStream = response.GetResponseStream();
            StreamReader myStreamReader = new StreamReader(myResponseStream, Encoding.GetEncoding("utf-8"));
            string retString = myStreamReader.ReadToEnd();
            myStreamReader.Close();
            myResponseStream.Close();
            return retString;
        }
        ///Post请求
        ///字符串
        public static string PostHttpResponse(string url, string param)
        {

            HttpWebRequest WebReq = (HttpWebRequest)WebRequest.Create(url);
            WebReq.ContentType = "application/json";
            WebReq.Method = "Post";
            WebReq.ContentLength = Encoding.UTF8.GetByteCount(param);
            using (StreamWriter requestW = new StreamWriter(WebReq.GetRequestStream()))
            {
                requestW.Write(param);
            }
            string backstr = null;
            using (HttpWebResponse response = (HttpWebResponse)WebReq.GetResponse())
            {
                StreamReader sr = new StreamReader(response.GetResponseStream(), System.Text.Encoding.UTF8);
                backstr = sr.ReadToEnd();
            }

            return backstr;

        }

        ///post请求
        ///json对象
        public static string PostHttpResponseJson(string url, Dictionary<string, object> param)
        {
            HttpWebRequest request = WebRequest.Create(url) as HttpWebRequest; //创建请求
            CookieContainer cookieContainer = new CookieContainer();
            request.Method = "POST"; //请求方式为post
            request.ContentType = "application/json";
            JObject json = new JObject();
            if (param.Count != 0) //将参数添加到json对象中
            {
                foreach (var item in param)
                {
                    json.Add(item.Key, item.Value.ToString());
                }
            }
            string jsonstring = json.ToString();//获得参数的json字符串
            byte[] jsonbyte = Encoding.UTF8.GetBytes(jsonstring);
            Stream postStream = request.GetRequestStream();
            postStream.Write(jsonbyte, 0, jsonbyte.Length);
            postStream.Close();
            //发送请求并获取相应回应数据       
            HttpWebResponse res;
            try
            {
                res = (HttpWebResponse)request.GetResponse();
            }
            catch (WebException ex)
            {
                res = (HttpWebResponse)ex.Response;
            }
            StreamReader sr = new StreamReader(res.GetResponseStream(), Encoding.UTF8);
            string content = sr.ReadToEnd(); //获得响应字符串
            return content;
        }
    }
}