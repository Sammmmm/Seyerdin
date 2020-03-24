using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Registry.UI {
    static class Actions {
        public static async Task<ObservableCollection<ServerInfo>> LoadServersList() {
            var servers = await ServerDataRetriever.FetchServerInfos();

            return servers;
        }

        public static async Task PingServers(ObservableCollection<ServerInfo> servers) {
            await ServerDataRetriever.PingServers(servers);
        }

        public static void LaunchServer(ServerInfo info) {
            if (info != null) {
                var arguments = $"-ip {info.IP} -port {info.Port} {(!string.IsNullOrWhiteSpace(info.UniqueId) ? $"-serverid {info.UniqueId}" : "")} {(info.HasCustomClasses ? "-cclasses" : "")} {(info.HasCustomClasses ? "-cskilldata" : "")}";

                var processPath = "Seyerdin.exe";
#if DEBUG
                processPath = "..\\..\\..\\client\\" + processPath;
#endif

                var processInfo = new ProcessStartInfo(processPath, arguments.Trim());
                var process = Process.Start(processInfo);

                Environment.Exit(0);
            }
        }
    }
}
