using System.Collections.Generic;
using System.Diagnostics;
using System.Net;
using System.Net.Sockets;
using System.Runtime.InteropServices;

namespace wpfApp.HelperClass
{
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ComVisible(true)]
    public class IpHelper
    {
        public string[] ShowMyLocalIp()
        {
            List<string> ipList = new List<string>();

            // 获取主机名
            string hostName = Dns.GetHostName();

            // 获取所有IP地址信息
            IPHostEntry hostEntry = Dns.GetHostEntry(hostName);

            foreach (IPAddress ip in hostEntry.AddressList)
            {
                // 只获取IPv4地址
                if (ip.AddressFamily == AddressFamily.InterNetwork)
                {
                    ipList.Add(ip.ToString());
                }
            }

            return ipList.ToArray();

        }
    }
}