namespace RoleDomain.Common
{
    using System;
    using System.IO;
    using System.Net;
    using System.Text;
    using System.Web;

    public class HTTPGET
    {
        public static bool UrlIsExist(String url)
        {
            return UrlIsExist(url, 10000);
        }

        public static bool UrlIsExist(String url, int timeout)
        {
            System.Uri u = null;
            try
            {
                u = new Uri(url);
            }
            catch { return false; }
            bool isExist = false;
            System.Net.HttpWebRequest r = System.Net.HttpWebRequest.Create(u)
                                    as System.Net.HttpWebRequest;
            r.Method = "HEAD";
            try
            {
                System.Net.HttpWebResponse s = r.GetResponse() as System.Net.HttpWebResponse;
                r.Timeout = timeout;
                if (s.StatusCode == System.Net.HttpStatusCode.OK)
                {
                    isExist = true;
                }
            }
            catch (System.Net.WebException x)
            {
                try
                {
                    isExist = ((x.Response as System.Net.HttpWebResponse).StatusCode !=
                                 System.Net.HttpStatusCode.NotFound);
                }
                catch { isExist = (x.Status == System.Net.WebExceptionStatus.Success); }
            }
            return isExist;
        }  
        public static string GetWebContent(string Url)
        {
            return GetWebContent(Url, false, Encoding.GetEncoding("GB2312"));
        }

        public static string GetWebContent(string Url, bool urlfix)
        {
            return GetWebContent(Url, urlfix, Encoding.GetEncoding("GB2312"));
        }

        public static string GetWebContent(string Url, Encoding enc)
        {
            return GetWebContent(Url, false, Encoding.GetEncoding("GB2312"));
        }

        public static string GetWebContent(string Url, bool urlfix, Encoding enc)
        {
            string strResult = "";
            if (urlfix && (Url.IndexOf("http://") < 0))
            {
                Url = string.Concat(new object[] { "http://", HttpContext.Current.Request.Url.Host, ":", HttpContext.Current.Request.Url.Port, "/", Url });
            }
            try
            {
                HttpWebRequest request = (HttpWebRequest) WebRequest.Create(Url);
                request.Timeout = 0x7530;
                request.Headers.Set("Pragma", "no-cache");
                HttpWebResponse response = (HttpWebResponse) request.GetResponse();
                strResult = new StreamReader(response.GetResponseStream(), enc).ReadToEnd();
            }
            catch (Exception e)
            {
                if (urlfix)
                {
                    strResult = e.Message.ToString() + " err:" + Url;
                }
            }
            return strResult;
        }

        public static string PostWebContent(string PageUrl, string postStr)
        {
            return PostWebContent(PageUrl, postStr, Encoding.GetEncoding("gbk"));
        }

        public static string PostWebContent(string PageUrl, string postStr, Encoding enc)
        {
            byte[] postBin = enc.GetBytes(postStr);
            HttpWebRequest request = WebRequest.Create(PageUrl) as HttpWebRequest;
            request.Method = "post";
            request.KeepAlive = false;
            request.AllowAutoRedirect = false;
            request.ContentType = "application/x-www-form-urlencoded";
            request.ContentLength = postBin.Length;
            Stream ioStream = request.GetRequestStream();
            ioStream.Write(postBin, 0, postBin.Length);
            ioStream.Flush();
            ioStream.Close();
            HttpWebResponse response = request.GetResponse() as HttpWebResponse;
            StreamReader reader = new StreamReader(response.GetResponseStream(), enc);
            string r = reader.ReadToEnd();
            reader.Close();
            response.Close();
            return r;
        }
        public static string PostXML(string PageUrl, string postStr)
        {
            return PostXML(PageUrl, postStr, Encoding.GetEncoding("gbk"));
        }
        public static string PostXML(string PageUrl, string postStr, Encoding enc)
        {
            byte[] postBin = enc.GetBytes(postStr);
            HttpWebRequest request = WebRequest.Create(PageUrl) as HttpWebRequest;
            request.Method = "post";
            request.KeepAlive = false;
            request.AllowAutoRedirect = false;
            request.ContentType = "text/xml";
            request.ContentLength = postBin.Length;
            Stream ioStream = request.GetRequestStream();
            ioStream.Write(postBin, 0, postBin.Length);
            ioStream.Flush();
            ioStream.Close();
            HttpWebResponse response = request.GetResponse() as HttpWebResponse;
            StreamReader reader = new StreamReader(response.GetResponseStream(), enc);
            string r = reader.ReadToEnd();
            reader.Close();
            response.Close();
            return r;
        }
    }
}

