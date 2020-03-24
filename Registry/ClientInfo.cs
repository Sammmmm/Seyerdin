using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Registry {
    static class ClientInfo {

        public static int ClientVersion;

        public static async Task<int> FetchClientVersion() {
            var arguments = $"-ver";

            var processPath = "Seyerdin.exe";
#if DEBUG
            processPath = "..\\..\\..\\client\\" + processPath;
#endif

            ProcessStartInfo processInfo = new ProcessStartInfo(processPath, arguments.Trim());
            processInfo.RedirectStandardOutput = true;
            processInfo.UseShellExecute = false;

            Process process = new Process();
            process.StartInfo = processInfo;

            process.Start();

            process.WaitForExit();
            if (process.ExitCode < 0) {
                throw new FileLoadException("Could not launch Seyerdin.exe");
            }

            ClientVersion = int.Parse(await process.StandardOutput.ReadToEndAsync());

            return ClientVersion;
        }
    }
}
