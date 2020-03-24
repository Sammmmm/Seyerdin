using Microsoft.VisualBasic.FileIO;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Threading.Tasks;

namespace Registry.UI {
    static class ServerDataRetriever {
        public static async Task<ObservableCollection<ServerInfo>> FetchServerInfos() {
            var serverInfos = new List<ServerInfo>();

            var serverIds = new HashSet<string>();

            string URL = "https://docs.google.com/spreadsheets/d/1cTPl401pwbp3JznpKWU-AHIVPvjV4K0cGfnA6fmAWoI/gviz/tq?tqx=out:csv&sheet=Servers";
            var request = (HttpWebRequest)WebRequest.Create(URL);
            var response = await request.GetResponseAsync() as HttpWebResponse;

            using (var responseStream = response.GetResponseStream()) {
                using (var parser = new TextFieldParser(responseStream)) {
                    parser.TextFieldType = FieldType.Delimited;
                    parser.SetDelimiters(",");

                    string[] fields = parser.ReadFields(); // header row

                    while (!parser.EndOfData) {
                        //Process row
                        fields = parser.ReadFields();
                        foreach (string field in fields) {
                            string serverUid = fields[3].Substring(0, Math.Min(fields[3].Length, 10)).Trim();

                            if (!serverIds.Contains(serverUid)) { // somebody didn't use a unique serverid, so I'm ignoring it, sorry folks
                                serverIds.Add(serverUid);

                                int.TryParse(fields[2], out int port);
                                bool.TryParse(fields[5], out bool hasCustomSkillData);
                                bool.TryParse(fields[6], out bool hasCustomClasses);

                                if (!int.TryParse(fields[7], out int priority)) {
                                    priority = 10000;
                                }

                                serverInfos.Add(new ServerInfo() {
                                    Name = fields[0].Trim(),
                                    IP = fields[1].Trim(),
                                    Port = port,
                                    UniqueId = serverUid,
                                    Description = fields[4].Trim(),
                                    HasCustomSkilldata = hasCustomSkillData,
                                    HasCustomClasses = hasCustomClasses,
                                    Priority = priority,
                                });
                            }
                        }
                    }
                }
            }

            return new ObservableCollection<ServerInfo>(serverInfos.OrderBy(server=>server.Priority));
        }

        public static async Task PingServers(ObservableCollection<ServerInfo> serverInfos) {
            await Task.WhenAll(serverInfos.Select(PingServer));
        }

        private static async Task PingServer(ServerInfo server) {
            try {
                var ip = (await Dns.GetHostAddressesAsync(server.IP)).FirstOrDefault();

                if (ip != null) {
                    using (var client = new TcpClient()) {
                        await client.ConnectAsync(ip, server.Port);
                        byte[] buffer = new byte[256];
                        var watch = System.Diagnostics.Stopwatch.StartNew();
                        using (var connectionStream = client.GetStream()) {

                            SendPacket(connectionStream, new byte[] { (byte)5 });

                            var response = await connectionStream.ReadAsync(buffer, 0, 256);

                            server.UserCount = buffer[0];

                            server.Version = buffer[1] * 16777216 + buffer[2] * 65536 + buffer[3] * 256 + buffer[4];
                        }
                        watch.Stop();
                        server.Ping = watch.ElapsedMilliseconds.ToString();
                    }
                }
            }
            catch {
                // something went wrong
            }
        }

        private static async void SendPacket(NetworkStream stream, byte[] toSend) {
            byte packetsSent = 1;
            
            int calc = 0;
            byte calc2 = 0;

            foreach(char c in toSend) {
                calc += c + 7;
            }

            calc = (byte)(calc % 256);
            calc2 = (byte)(calc ^ packetsSent); //xor
            calc2 = (byte)~calc2; //not

            int finalPacketLength = toSend.Length+1;

            var bytesToSend = new List<byte>();

            bytesToSend.Add((byte)(finalPacketLength / 256));
            bytesToSend.Add((byte)(finalPacketLength % 256));
            bytesToSend.AddRange(toSend);
            bytesToSend.Add((byte)calc2);

            await stream.WriteAsync(bytesToSend.ToArray(), 0, bytesToSend.Count);
        }
    }
}
