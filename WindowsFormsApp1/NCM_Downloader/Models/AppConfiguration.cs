using System;

namespace NCM_Downloader.Models
{
    public class AppConfiguration
    {
        public string LoadLogPath { get; set; } = string.Empty;
        public int FirstClientIndex { get; set; }
        public int NumClients { get; set; }
        public string LmHostsPath { get; set; } = string.Empty;
        public string Username { get; set; } = "SDI";
        public string Password { get; set; } = "A02460";
        public bool AutoRdp { get; set; } = true;
        public int RdpWidth { get; set; } = 1024;
        public int RdpHeight { get; set; } = 768;
        public int RdpLeft { get; set; } = 0;
        public int RdpTop { get; set; } = 300;
        public int ParallelDownloads { get; set; } = 5;
        public string SourcePath { get; set; } = string.Empty;
        public string DestinationPath { get; set; } = string.Empty;
        public string McpPath { get; set; } = string.Empty;
        public string VpnPassword { get; set; } = string.Empty;
        public bool CheckBox2 { get; set; } // TODO: Rename to meaningful name after analysis
        public bool IncludeNonClients { get; set; }
        public bool KillKey { get; set; }
        public bool UseRemotes { get; set; }
    }
}
