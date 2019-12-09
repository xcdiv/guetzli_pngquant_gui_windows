namespace RoleDomain.Common
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Drawing;
    using System.IO;
    using System.Net;
    using System.Net.Security;
    using System.Reflection;
    using System.Security.Cryptography.X509Certificates;
    using System.Text;
    using System.Web;
    using System.Web.UI;

    public class HTTPGET_UTF8 : Page
    {
        private static readonly string DefaultUserAgent = "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; SV1; .NET CLR 1.1.4322; .NET CLR 2.0.50727)";
        public static bool UrlIsExist(String url)
        {
            return UrlIsExist(url, 10000);
        }

        //仅检测链接头，不会获取链接的结果。所以速度很快，超时的时间单位为毫秒
        public static string GetWebStatusCode(string url, int timeout)
        {
            HttpWebRequest req = null;
            try
            {
                req = (HttpWebRequest)WebRequest.CreateDefault(new Uri(url));
                req.Method = "HEAD";  //这是关键        
                req.Timeout = timeout;
                HttpWebResponse res = (HttpWebResponse)req.GetResponse();
                return Convert.ToInt32(res.StatusCode).ToString();
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            finally
            {
                if (req != null)
                {
                    req.Abort();
                    req = null;
                }
            }

        }

        //需要注意的是如果你使用多线程。。C#默认同时只有4个网络线程，如需要破解此限制需要添加代码
        //ServicePointManager.DefaultConnectionLimit = 100;

        //此方法返回一个状态码。。状态码为200是为正常，异常时会返回错误信息。比如超时


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
        public static bool download_file(string _url, string _savepath)
        {
            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(_url);
            req.UserAgent = "netclient";
            WebResponse resp;
            try
            {
                resp = req.GetResponse();
            }
            catch (Exception err)
            {

                return false;
            }

            Stream stream = resp.GetResponseStream();

            FileStream fs = new FileStream(_savepath, FileMode.Create);

            byte[] nbytes = new byte[512];
            int nReadSize = 0;
            nReadSize = stream.Read(nbytes, 0, 512);
            while (nReadSize > 0)
            {
                fs.Write(nbytes, 0, nReadSize);
                nReadSize = stream.Read(nbytes, 0, 512);
            }


            fs.Close();
            stream.Close();
            resp.Close();
            return true;
        }



        private static bool CheckValidationResult(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors errors)
        {
            return true;
        }

        public static HttpWebResponse CreateGetHttpResponse(string url, int? timeout, string userAgent, CookieCollection cookies)
        {
            if (string.IsNullOrEmpty(url))
            {
                throw new ArgumentNullException("url");
            }
            HttpWebRequest request = WebRequest.Create(url) as HttpWebRequest;
            request.Method = "GET";
            request.UserAgent = DefaultUserAgent;
            if (!string.IsNullOrEmpty(userAgent))
            {
                request.UserAgent = userAgent;
            }
            if (timeout.HasValue)
            {
                request.Timeout = timeout.Value;
            }
            if (cookies != null)
            {
                request.CookieContainer = new CookieContainer();
                request.CookieContainer.Add(cookies);
            }
            return (request.GetResponse() as HttpWebResponse);
        }

        public static HttpWebResponse CreatePostHttpResponse(string url, IDictionary<string, string> parameters, int? timeout, string userAgent, Encoding requestEncoding, CookieCollection cookies)
        {
            if (string.IsNullOrEmpty(url))
            {
                throw new ArgumentNullException("url");
            }
            if (requestEncoding == null)
            {
                throw new ArgumentNullException("requestEncoding");
            }
            HttpWebRequest request = null;
            if (url.StartsWith("https", StringComparison.OrdinalIgnoreCase))
            {
                ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(HTTPGET_UTF8.CheckValidationResult);
                request = WebRequest.Create(url) as HttpWebRequest;
                request.ProtocolVersion = HttpVersion.Version10;
            }
            else
            {
                request = WebRequest.Create(url) as HttpWebRequest;
            }
            request.Method = "POST";
            request.ContentType = "application/x-www-form-urlencoded";
            if (!string.IsNullOrEmpty(userAgent))
            {
                request.UserAgent = userAgent;
            }
            else
            {
                request.UserAgent = DefaultUserAgent;
            }
            if (timeout.HasValue)
            {
                request.Timeout = timeout.Value;
            }
            if (cookies != null)
            {
                request.CookieContainer = new CookieContainer();
                request.CookieContainer.Add(cookies);
            }
            if ((parameters != null) && (parameters.Count != 0))
            {
                StringBuilder buffer = new StringBuilder();
                int i = 0;
                foreach (string key in parameters.Keys)
                {
                    if (i > 0)
                    {
                        buffer.AppendFormat("&{0}={1}", key, parameters[key]);
                    }
                    else
                    {
                        buffer.AppendFormat("{0}={1}", key, parameters[key]);
                    }
                    i++;
                }
                byte[] data = requestEncoding.GetBytes(buffer.ToString());
                using (Stream stream = request.GetRequestStream())
                {
                    stream.Write(data, 0, data.Length);
                }
            }
            return (request.GetResponse() as HttpWebResponse);
        }

        public static List<Cookie> GetAllCookies(CookieContainer cc)
        {
            List<Cookie> lstCookies = new List<Cookie>();
            HttpRequest request = HttpContext.Current.Request;
            Hashtable table = (Hashtable)cc.GetType().InvokeMember("m_domainTable", BindingFlags.GetField | BindingFlags.NonPublic | BindingFlags.Instance, null, cc, new object[0]);
            foreach (object pathList in table.Values)
            {
                SortedList lstCookieCol = (SortedList)pathList.GetType().InvokeMember("m_list", BindingFlags.GetField | BindingFlags.NonPublic | BindingFlags.Instance, null, pathList, new object[0]);
                foreach (CookieCollection colCookies in lstCookieCol.Values)
                {
                    foreach (Cookie c in colCookies)
                    {
                        lstCookies.Add(c);
                    }
                }
            }
            return lstCookies;
        }

        public static List<Cookie> GetAllCookiesWeb(CookieContainer cc)
        {
            List<Cookie> lstCookies = new List<Cookie>();
            HttpRequest request = HttpContext.Current.Request;
            Hashtable table = (Hashtable)cc.GetType().InvokeMember("m_domainTable", BindingFlags.GetField | BindingFlags.NonPublic | BindingFlags.Instance, null, cc, new object[0]);
            foreach (object pathList in table.Values)
            {
                SortedList lstCookieCol = (SortedList)pathList.GetType().InvokeMember("m_list", BindingFlags.GetField | BindingFlags.NonPublic | BindingFlags.Instance, null, pathList, new object[0]);
                foreach (CookieCollection colCookies in lstCookieCol.Values)
                {
                    foreach (Cookie c in colCookies)
                    {
                        if (c.Domain.ToLower() != request.Url.Host.ToString())
                        {
                            Cookie c_copy = c;
                            lstCookies.Add(c_copy);
                            c_copy.Domain = request.Url.Host.ToString();
                        }
                        lstCookies.Add(c);
                    }
                }
            }
            return lstCookies;
        }

        public static CookieContainer GetCurrent()
        {
            CookieContainer cc = new CookieContainer();
            try
            {
                if (HttpContext.Current.Session["usersession"] != null)
                {
                    cc = (CookieContainer)HttpContext.Current.Session["usersession"];
                }
            }
            catch
            {
            }
            return cc;
        }

        public static string GetWebContent(string Url)
        {
            return GetWebContent(Url, false, Encoding.GetEncoding("UTF-8"));
        }

        public static string GetWebContent(string Url, bool urlfix)
        {
            return GetWebContent(Url, urlfix, Encoding.GetEncoding("UTF-8"));
        }

        public static string GetWebContent(string PageUrl, string postStr)
        {
            return GetWebContent(PageUrl, postStr, Encoding.GetEncoding("UTF-8"));
        }

        public static string GetWebContent(string Url, Encoding enc)
        {
            return GetWebContent(Url, false, Encoding.GetEncoding("UTF-8"));
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
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(Url);
                request.Timeout = 0x7530;
                request.Headers.Set("Pragma", "no-cache");
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
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
        public static byte[] GetWebBin(string PageUrl, string postStr, Encoding enc)
        {
            byte[] postBin = enc.GetBytes(postStr);
            HttpWebRequest request = WebRequest.Create(PageUrl) as HttpWebRequest;
            request.Method = "get";
            request.UserAgent = "asynchlient/2.0";
            request.KeepAlive = false;
            request.Proxy = null;
            request.AllowAutoRedirect = false;
            request.ContentType = "application/x-www-form-urlencoded";
            request.ContentLength = postBin.Length;
            HttpWebResponse response = request.GetResponse() as HttpWebResponse;
            Stream responseStream = response.GetResponseStream();
            MemoryStream stmMemory = new MemoryStream();
            byte[] buffer = new byte[64 * 1024];
            int i;
            while ((i = responseStream.Read(buffer, 0, buffer.Length)) > 0)
            {
                stmMemory.Write(buffer, 0, i);
            }
            byte[] r = stmMemory.ToArray();
            stmMemory.Close();

            //StreamReader reader = new StreamReader(response.GetResponseStream(), enc);
            //string r = reader.ReadToEnd();
            //reader.Close();
            response.Close();
            return r;
        }
        public static string GetWebContent(string PageUrl, string postStr, Encoding enc)
        {
            byte[] postBin = enc.GetBytes(postStr);
            HttpWebRequest request = WebRequest.Create(PageUrl) as HttpWebRequest;
            request.Method = "get";
            request.UserAgent = "asynchlient/2.0";
            request.KeepAlive = false;
            request.Proxy = null;
            request.AllowAutoRedirect = false;
            request.ContentType = "application/x-www-form-urlencoded";
            request.ContentLength = postBin.Length;
            HttpWebResponse response = request.GetResponse() as HttpWebResponse;
            StreamReader reader = new StreamReader(response.GetResponseStream(), enc);
            string r = reader.ReadToEnd();
            reader.Close();
            response.Close();
            return r;
        }

        public string newPostWebContent(string PageUrl, string postStr, Encoding enc)
        {
            byte[] postBin = enc.GetBytes(postStr);
            HttpWebRequest request = WebRequest.Create(PageUrl) as HttpWebRequest;
            request.Method = "post";
            request.UserAgent = "asynchlient/2.0";
            request.KeepAlive = false;
            request.Proxy = null;
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

        public static string PostWebContent(string PageUrl, string postStr)
        {
            return PostWebContent(PageUrl, postStr, Encoding.GetEncoding("UTF-8"));
        }

        public static string PostWebContent(string PageUrl, string postStr, Encoding enc)
        {
            return PostWebContent(PageUrl, postStr, enc, 0x7530);
        }

        public static string PostWebContent(string PageUrl, string postStr, ref CookieContainer cookie)
        {
            return PostWebContent(PageUrl, postStr, Encoding.GetEncoding("UTF-8"), ref cookie);
        }

        public static string PostWebContent(string PageUrl, string postStr, Encoding enc, int Timeout)
        {
            HttpWebRequest request;
            if (PageUrl.StartsWith("https", StringComparison.OrdinalIgnoreCase))
            {
                request = WebRequest.Create(PageUrl) as HttpWebRequest;
                ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(CheckValidationResult);
                request.ProtocolVersion = HttpVersion.Version11;
                // 这里设置了协议类型。
                ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072;// SecurityProtocolType.Tls1.2; 
                request.KeepAlive = false;
                ServicePointManager.CheckCertificateRevocationList = true;
                ServicePointManager.DefaultConnectionLimit = 100;
                ServicePointManager.Expect100Continue = false;
            }
            else
            {

                request = (HttpWebRequest)WebRequest.Create(PageUrl);
            }


            byte[] postBin = enc.GetBytes(postStr);
            //  HttpWebRequest request = WebRequest.Create(PageUrl) as HttpWebRequest;
            request.Method = "post";
            request.UserAgent = "asynchlient/2.0";
            request.KeepAlive = false;
            request.Proxy = null;
            request.Timeout = Timeout;
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

        public static string PostWebContent(string PageUrl, string postStr, Encoding enc, ref CookieContainer cookie)
        {
            return PostWebContent(PageUrl, postStr, enc, 0x7530, ref cookie);
        }

        public static string PostWebContent(string PageUrl, string postStr, Encoding enc, int Timeout, ref CookieContainer cookie)
        {
            CookieContainer container = (cookie == null) ? new CookieContainer() : cookie;
            byte[] postBin = enc.GetBytes(postStr);
            HttpWebRequest request = WebRequest.Create(PageUrl) as HttpWebRequest;
            request.Method = "post";
            request.Timeout = Timeout;
            request.KeepAlive = false;
            request.Proxy = null;
            request.AllowAutoRedirect = false;
            request.ContentType = "application/x-www-form-urlencoded";
            request.ContentLength = postBin.Length;
            request.CookieContainer = container;
            Stream ioStream = request.GetRequestStream();
            ioStream.Write(postBin, 0, postBin.Length);
            ioStream.Flush();
            ioStream.Close();
            HttpWebResponse response = request.GetResponse() as HttpWebResponse;
            response.Cookies = container.GetCookies(request.RequestUri);
            cookie = container;
            StreamReader reader = new StreamReader(response.GetResponseStream(), enc);
            string r = reader.ReadToEnd();
            reader.Close();
            response.Close();
            return r;
        }

        public static void WriteCurrent(CookieContainer cc)
        {
            List<Cookie> cooklist = GetAllCookies(cc);
            foreach (Cookie c in cooklist)
            {
                HttpCookie cookie = new HttpCookie(c.Name)
                {
                    Domain = c.Domain,
                    Path = c.Path,
                    Secure = c.Secure,
                    Value = c.Value,
                    Expires = c.Expires
                };
                HttpContext.Current.Response.Cookies.Set(cookie);
            }
            HttpContext.Current.Session["usersession"] = cc;
        }

        public static void WriteCurrentCookieSyncDomain(CookieContainer cc)
        {
            List<Cookie> cooklist = GetAllCookies(cc);
            foreach (Cookie c in cooklist)
            {
                HttpCookie cookie = new HttpCookie(c.Name)
                {
                    Domain = c.Domain,
                    Path = c.Path,
                    Secure = c.Secure,
                    Value = c.Value,
                    Expires = c.Expires
                };
                HttpContext.Current.Response.Cookies.Set(cookie);
                string remote_url = HttpContext.Current.Request.Url.Host.ToString();
                if (c.Domain.ToLower() != remote_url.ToLower())
                {
                    HttpCookie cookie2 = new HttpCookie(c.Name)
                    {
                        Domain = remote_url,
                        Path = c.Path,
                        Secure = c.Secure,
                        Value = c.Value,
                        Expires = c.Expires
                    };
                    HttpContext.Current.Response.Cookies.Set(cookie2);
                }
            }
            HttpContext.Current.Session["usersession"] = cc;
        }
        public static string PostXML(string PageUrl, string postStr)
        {
            return PostXML(PageUrl, postStr, Encoding.GetEncoding("utf-8"));
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

        public static string GetPage(string posturl, string postData)
        {
            Stream outstream = null;
            Stream instream = null;
            StreamReader sr = null;
            HttpWebResponse response = null;
            HttpWebRequest request = null;
            Encoding encoding = Encoding.UTF8;
            byte[] data = encoding.GetBytes(postData);
            // 准备请求...
            try
            {
                // 设置参数
                request = WebRequest.Create(posturl) as HttpWebRequest;
                CookieContainer cookieContainer = new CookieContainer();
                request.CookieContainer = cookieContainer;
                request.AllowAutoRedirect = true;
                request.Method = "POST";
                request.ContentType = "application/x-www-form-urlencoded";
                request.ContentLength = data.Length;
                outstream = request.GetRequestStream();
                outstream.Write(data, 0, data.Length);
                outstream.Close();
                //发送请求并获取相应回应数据
                response = request.GetResponse() as HttpWebResponse;
                //直到request.GetResponse()程序才开始向目标网页发送Post请求
                instream = response.GetResponseStream();
                sr = new StreamReader(instream, encoding);
                //返回结果网页（html）代码
                string content = sr.ReadToEnd();
                string err = string.Empty;
                //Response.Write(content);
                return content;
            }
            catch (Exception ex)
            {
                //string err = ex.Message;
                //return string.Empty;

                return "err:" + ex.Message;
            }
        }

        public static string GetPage(string posturl)
        {
            Stream instream = null;
            StreamReader sr = null;
            HttpWebResponse response = null;
            HttpWebRequest request = null;
            Encoding encoding = Encoding.UTF8;
            // 准备请求...
            try
            {
                // 设置参数
                request = WebRequest.Create(posturl) as HttpWebRequest;
                CookieContainer cookieContainer = new CookieContainer();
                request.CookieContainer = cookieContainer;
                request.AllowAutoRedirect = true;
                request.Method = "GET";
                request.ContentType = "application/x-www-form-urlencoded";
                //发送请求并获取相应回应数据
                response = request.GetResponse() as HttpWebResponse;
                //直到request.GetResponse()程序才开始向目标网页发送Post请求
                instream = response.GetResponseStream();
                sr = new StreamReader(instream, encoding);
                //返回结果网页（html）代码
                string content = sr.ReadToEnd();
                string err = string.Empty;
                //Response.Write(content);
                return content;
            }
            catch (Exception ex)
            {
                //string err = ex.Message;
                //return string.Empty;

                return "err:" + ex.Message;
            }
        }


        public static Bitmap Get_img(string url)
        {
            Bitmap img = null;
            HttpWebRequest req;
            HttpWebResponse res = null;
            try
            {
                System.Uri httpUrl = new System.Uri(url);
                req = (HttpWebRequest)(WebRequest.Create(httpUrl));
                //req.Timeout = 180000; //设置超时值10秒
                //req.UserAgent = "XXXXX";
                //req.Accept = "XXXXXX";
                req.Method = "GET";
                res = (HttpWebResponse)(req.GetResponse());
                img = new Bitmap(res.GetResponseStream());//获取图片流                
                //img.Save(@"E:/" + DateTime.Now.ToFileTime().ToString() + ".png");//随机名
            }

            catch (Exception ex)
            {
                //string aa = ex.Message;
                throw ex;
            }
            finally
            {
                res.Close();
            }
            return img;
        }
    }
}

