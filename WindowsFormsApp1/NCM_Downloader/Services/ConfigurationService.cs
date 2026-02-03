using System;
using System.IO;
using System.Linq;
using NCM_Downloader.Models;
using Encryption; // Assuming StringCipher is here

namespace NCM_Downloader.Services
{
    public class ConfigurationService
    {
        private readonly string _configPath;
        private const string PassPhrase = "Hush, little baby, don't say a word " +
            "And never mind that noise you heard " +
            "It's just the beasts under your bed " +
            "In your closet, in your head";

        public ConfigurationService(string startupPath)
        {
            if (File.Exists(Path.Combine(startupPath, "NCM_Downloader.ini")))
            {
                _configPath = Path.Combine(startupPath, "NCM_Downloader.ini");
            }
            else if (File.Exists(Path.Combine(startupPath, "\\NCM_Downloader.exe.ini")))
            {
                _configPath = Path.Combine(startupPath, "\\NCM_Downloader.exe.ini");
            }
            else
            {
                _configPath = Path.Combine(startupPath, "NCM_Downloader.ini"); // Default
            }
        }

        public AppConfiguration Load()
        {
            var config = new AppConfiguration();
            if (!File.Exists(_configPath)) return config;

            var lines = File.ReadLines(_configPath).ToList();
            if (lines.Count == 0) return config;

            try
            {
                config.LoadLogPath = GetLine(lines, 0);
                config.FirstClientIndex = ParseInt(GetLine(lines, 1));
                config.NumClients = ParseInt(GetLine(lines, 2));
                config.LmHostsPath = GetLine(lines, 3);
                config.Username = GetLine(lines, 4);
                
                string encPass = GetLine(lines, 5);
                if (!string.IsNullOrEmpty(encPass))
                    config.Password = StringCipher.Decrypt(encPass, PassPhrase);

                config.AutoRdp = ParseBool(GetLine(lines, 6));
                config.RdpWidth = ParseInt(GetLine(lines, 7));
                config.RdpHeight = ParseInt(GetLine(lines, 8));
                config.RdpLeft = ParseInt(GetLine(lines, 9));
                config.RdpTop = ParseInt(GetLine(lines, 10));
                config.ParallelDownloads = ParseInt(GetLine(lines, 11));
                config.SourcePath = GetLine(lines, 12);
                config.DestinationPath = GetLine(lines, 13);
                config.McpPath = GetLine(lines, 14);

                string encVpn = GetLine(lines, 15);
                if (!string.IsNullOrEmpty(encVpn))
                    config.VpnPassword = StringCipher.Decrypt(encVpn, PassPhrase);

                config.CheckBox2 = ParseBool(GetLine(lines, 16));
                config.IncludeNonClients = ParseBool(GetLine(lines, 17));
                config.KillKey = ParseBool(GetLine(lines, 18));
                config.UseRemotes = ParseBool(GetLine(lines, 19));
            }
            catch (Exception)
            {
                // Log error or return partial config
            }

            return config;
        }

        public void Save(AppConfiguration config)
        {
            using (var writer = File.CreateText(_configPath))
            {
                writer.WriteLine(config.LoadLogPath);
                writer.WriteLine(config.FirstClientIndex);
                writer.WriteLine(config.NumClients);
                writer.WriteLine(config.LmHostsPath);
                writer.WriteLine(config.Username);
                writer.WriteLine(StringCipher.Encrypt(config.Password, PassPhrase));
                writer.WriteLine(config.AutoRdp);
                writer.WriteLine(config.RdpWidth);
                writer.WriteLine(config.RdpHeight);
                writer.WriteLine(config.RdpLeft);
                writer.WriteLine(config.RdpTop);
                writer.WriteLine(config.ParallelDownloads);
                writer.WriteLine(config.SourcePath);
                writer.WriteLine(config.DestinationPath);
                writer.WriteLine(config.McpPath);
                writer.WriteLine(StringCipher.Encrypt(config.VpnPassword, PassPhrase));
                writer.WriteLine(config.CheckBox2);
                writer.WriteLine(config.IncludeNonClients);
                writer.WriteLine(config.KillKey);
                writer.WriteLine(config.UseRemotes);
                
                // Filler lines to match original format if needed, or clean up
                for(int i=0; i<8; i++) writer.WriteLine(""); 
                
                writer.WriteLine("Authored by Muresan Radu-Adrian (MURA02)");
            }
        }

        private string GetLine(System.Collections.Generic.List<string> lines, int index)
        {
            return index < lines.Count ? lines[index] : string.Empty;
        }

        private int ParseInt(string val)
        {
            return int.TryParse(val, out int result) ? result : 0;
        }

        private bool ParseBool(string val)
        {
            return bool.TryParse(val, out bool result) ? result : false;
        }
    }
}
