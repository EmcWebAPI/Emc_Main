using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace EmcReportWebApi.Common
{
    public static class MD5Helper
    {
        /// <summary>
        /// 用MD5加密字符串，可选择生成16位或者32位的加密字符串
        /// </summary>
        /// <param name="password">待加密的字符串</param>
        /// <param name="bit">位数，一般取值16 或 32</param>
        /// <returns>返回的加密后的字符串</returns>
        public static string MD5Encrypt(string password, int bit = 32)
        {
            using (MD5 mi = MD5.Create())
            {
                byte[] buffer = Encoding.GetEncoding("gb2312").GetBytes(password);
                //开始加密
                byte[] newBuffer = mi.ComputeHash(buffer);
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < newBuffer.Length; i++)
                {
                    sb.Append(newBuffer[i].ToString("x2"));
                }
                if (bit == 16)
                    return sb.ToString().Substring(8, 16);
                else if (bit == 32) return sb.ToString();//默认情况
                else return string.Empty;
            }
        }

        /// <summary>
        /// MD5流加密
        /// </summary>
        /// <param name="inputStream"></param>
        /// <returns></returns>
        public static string GenerateMD5(Stream inputStream)
        {
            using (MD5 mi = MD5.Create())
            {
                //开始加密
                byte[] newBuffer = mi.ComputeHash(inputStream);
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < newBuffer.Length; i++)
                {
                    sb.Append(newBuffer[i].ToString("x2"));
                }
                return sb.ToString();
            }
        }
    }
}