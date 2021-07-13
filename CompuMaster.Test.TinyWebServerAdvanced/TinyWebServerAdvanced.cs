using System;
using System.Net;
using System.Text;
using System.Threading;

namespace CompuMaster.Test.Tools.TinyWebServerAdvanced
{
    /// <summary> 
    /// Listens for the specified request, and executes the given handler.
    /// 
    /// Example: 
    /// <code>
    /// var server = new WebServer(request => { return "<h1>Hello world!</h1>"; }, "http://localhost:8080/hello/");
    /// server.Run();
    /// ....
    /// server.Stop();
    /// </code>
    /// 
    /// Note - this code is adapted from http://codehosting.net/blog/BlogEngine/post/Simple-C-Web-Server.aspx
    /// The purpose of this library is more-or-less just to provide a Nuget package for this class.
    /// </summary>
    public class WebServer
    {
        private HttpListener _listener = new HttpListener();
        private readonly Func<HttpListenerRequest, string> _handler;
        private readonly System.Collections.Specialized.NameValueCollection _responseHeaders;

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Style", "IDE0016:throw-Ausdruck verwenden", Justification = "<Ausstehend>")]
        public WebServer(Func<HttpListenerRequest, string> handler, params string[] urls)
        {
            if (urls == null || urls.Length == 0)
                throw new ArgumentException("prefixes");
            if (handler == null)
                throw new ArgumentException("method");

            foreach (string s in urls)
                _listener.Prefixes.Add(s);

            _handler = handler;
            _listener.Start();
        }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Style", "IDE0016:throw-Ausdruck verwenden", Justification = "<Ausstehend>")]
        public WebServer(Func<HttpListenerRequest, string> handler, System.Collections.Specialized.NameValueCollection responseHeaders, params string[] urls)
        {
            if (urls == null || urls.Length == 0)
                throw new ArgumentException("prefixes");
            if (handler == null)
                throw new ArgumentException("method");

            foreach (string s in urls)
                _listener.Prefixes.Add(s);

            _handler = handler;
            _responseHeaders = responseHeaders;
            _listener.Start();
        }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Style", "IDE0019:Musterabgleich verwenden", Justification = "<Ausstehend>")]
        public void Run()
        {
            if (_listener == null) throw new InvalidOperationException("Server already closed");
            ThreadPool.QueueUserWorkItem(o =>
            {
                while (_listener != null && _listener.IsListening)
                {
                    HttpListenerContext httpContext;
                    try
                    {
                        httpContext = _listener.GetContext();
                    }
                    catch (HttpListenerException)
                    {
                        //Listener closed, but Linux/Mono doesn't stop to continue exection of this code block
                        return;
                    }
                    ThreadPool.QueueUserWorkItem(c =>
                    {
                        HttpListenerContext ctx = c as HttpListenerContext;
                        if (ctx != null)
                        {
                            try
                            {
                                var responseStr = _handler(ctx.Request);
                                var buf = Encoding.UTF8.GetBytes(responseStr);
                                foreach (string s in _responseHeaders)
                                    ctx.Response.Headers[s] = _responseHeaders[s];
                                ctx.Response.ContentLength64 = buf.Length;
                                ctx.Response.OutputStream.Write(buf, 0, buf.Length);
                            }
                            finally
                            {
                                ctx.Response.OutputStream.Close();
                            }
                        }
                    }, httpContext);
                }
            });
        }

        public void Stop()
        {
            if (_listener != null)
            {
                _listener.Stop();
                _listener.Close();
                _listener = null;
            }
        }

        ~WebServer()
        {
            Stop();
        }
    }
}