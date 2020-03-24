using System;
using System.Collections.Generic;
using System.Text;

namespace Registry.UI {
    public class ServerInfo : BindableState {
        public string Name { get; set; }
        public string Description { get; set; }
        public string IP { get; set; }
        public string UniqueId { get; set; }
        public int Port { get; set; }
        public bool HasCustomSkilldata { get; set; }
        public bool HasCustomClasses { get; set; }
        public int Priority { get; set; }

        public string NameTruncated { get { return Name.Length > 25 ? Name.Substring(0, 25) + "..." : Name; } }
        public string DescriptionTruncated { get { return Description.Length > 50 ? Description.Substring(0, 50) + "..." : Description; } }
        public string DescriptionVisibility {  get { return Description.Trim().Length > 0 ? "Visible" : "Collapsed"; } }

        public int? UserCount {
            get { return Get<int?>(); }
            set {
                Set<int?>(value);
            }
        }

        public int? Version {
            get { return Get<int?>(); }
            set {
                Set<int?>(value);
                OnPropertyChanged("VersionMatch");
                OnPropertyChanged("IsOnline");
            }
        }

        public string Ping {
            get { return Get<string>() ?? "offline"; }
            set {
                Set<string>(value);
                OnPropertyChanged("ShouldHide");
                OnPropertyChanged("IsOnline");
            }
        }

        public bool VersionMatch { 
            get { return Version == ClientInfo.ClientVersion; }
        }

        public bool IsOnline {
            get { return Ping != "offline"; }
        }

        public bool ShouldHide { get { return Ping == "offline" && Priority > 0; } }
    }
}
