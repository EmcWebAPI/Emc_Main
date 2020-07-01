using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web;

namespace EmcReportWebApi.Config
{
    /// <summary>
    /// 压缩内容
    /// </summary>
    public class CompressContent : HttpContent
    {
        private readonly string _encodingType;
        private readonly HttpContent _originalContent;
        /// <summary>
        /// new
        /// </summary>
        /// <param name="content"></param>
        /// <param name="encodingType"></param>
        public CompressContent(HttpContent content, string encodingType = "gzip")
        {
            _originalContent = content;
            _encodingType = encodingType.ToLowerInvariant();
            Headers.ContentEncoding.Add(encodingType);
        }
        /// <summary>
        /// 判断长度
        /// </summary>
        /// <param name="length"></param>
        /// <returns></returns>
        protected override bool TryComputeLength(out long length)
        {
            length = -1;
            return false;
        }
        /// <summary>
        /// 序列化内存流
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="context"></param>
        /// <returns></returns>
        protected override Task SerializeToStreamAsync(Stream stream, TransportContext context)
        {
            Stream compressStream = null;
            switch (_encodingType)
            {
                case "gzip":
                    compressStream = new GZipStream(stream, CompressionMode.Compress, true);
                    break;
                case "deflate":
                    compressStream = new DeflateStream(stream, CompressionMode.Compress, true);
                    break;
                default:
                    compressStream = stream;
                    break;
            }
            return _originalContent.CopyToAsync(compressStream).ContinueWith(tsk =>
            {
                if (compressStream != null)
                {
                    compressStream.Dispose();
                }
            });
        }
    }
}