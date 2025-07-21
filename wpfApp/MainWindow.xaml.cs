using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Web.WebView2.Core;
using wpfApp.HelperClass;

namespace wpfApp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        public MainWindow()
        {
            InitializeComponent();
            StartWebview();
        }

        private async void StartWebview()
        {
            var env = await CoreWebView2Environment.CreateAsync(
                userDataFolder: Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "YourApp/WebView2Data")
            );

            await WebView21.EnsureCoreWebView2Async(env);

            WebView21.CoreWebView2.AddHostObjectToScript("FileHelper", new FileHelper());
            WebView21.CoreWebView2.AddHostObjectToScript("IpHelper", new IpHelper());

            string appPath = AppDomain.CurrentDomain.BaseDirectory;

            if (appPath.Contains("Debug"))
            {
                bool isDevServerRunning = await CheckDevServer("http://localhost:5173");
                if (!isDevServerRunning)
                {
                    MessageBox.Show("请先启动Vite开发服务器");
                    this.Close();
                }
                else
                {
                    WebView21.CoreWebView2.Settings.IsGeneralAutofillEnabled = false;
                    //    WebView21.CoreWebView2.Settings.AreBrowserAcceleratorKeysEnabled = false;
                    //    WebView21.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
                    //    WebView21.CoreWebView2.Settings.AreDefaultScriptDialogsEnabled = false;
                    //    WebView21.CoreWebView2.Settings.AreDevToolsEnabled = false;
                    //    WebView21.CoreWebView2.Settings.IsZoomControlEnabled = false;
                    //    WebView21.CoreWebView2.Settings.IsStatusBarEnabled = false;
                    WebView21.CoreWebView2.Navigate("http://localhost:5173");
                }
            }
            else
            {
                WebView21.CoreWebView2.Settings.IsGeneralAutofillEnabled = false;
                WebView21.CoreWebView2.Settings.AreBrowserAcceleratorKeysEnabled = false;
                WebView21.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
                WebView21.CoreWebView2.Settings.AreDefaultScriptDialogsEnabled = false;
                WebView21.CoreWebView2.Settings.AreDevToolsEnabled = false;
                WebView21.CoreWebView2.Settings.IsZoomControlEnabled = false;
                WebView21.CoreWebView2.Settings.IsStatusBarEnabled = false;
                var wwwrootPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "wwwroot");

                // var exePath = Path.Combine($"{AppDomain.CurrentDomain.BaseDirectory}", "wwwroot", "index.html");
                WebView21.CoreWebView2.SetVirtualHostNameToFolderMapping(
                    "app.example", wwwrootPath,
                    CoreWebView2HostResourceAccessKind.Allow
                );
                // 等待 blazor 启动完成
                WebView21.Source = new Uri("http://app.example/index.html");

                //await Task.Delay(1500);
                //WebView21.CoreWebView2.Navigate("http://localhost:5000");
            }
        }

        // 检查 Vite 开发服务器是否运行
        private async Task<bool> CheckDevServer(string url)
        {
            try
            {
                using (var client = new HttpClient())
                {
                    client.Timeout = TimeSpan.FromSeconds(1); // 快速检测
                    var response = await client.GetAsync(url);
                    return response.IsSuccessStatusCode;
                }
            }
            catch
            {
                return false;
            }
        }
    }
}