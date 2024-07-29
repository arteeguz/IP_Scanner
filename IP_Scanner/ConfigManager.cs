using System;
using System.IO;
using System.Xml;
using Newtonsoft.Json;

namespace IPProcessingTool
{
    public class ConfigManager
    {
        private static readonly string configPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "IPProcessingTool",
            "config.json"
        );

        public static Config LoadConfig()
        {
            if (File.Exists(configPath))
            {
                string json = File.ReadAllText(configPath);
                return JsonConvert.DeserializeObject<Config>(json);
            }
            return new Config();
        }

        public static void SaveConfig(Config config)
        {
            string directoryPath = Path.GetDirectoryName(configPath);
            if (!Directory.Exists(directoryPath))
            {
                Directory.CreateDirectory(directoryPath);
            }

            string json = JsonConvert.SerializeObject(config, Newtonsoft.Json.Formatting.Indented);
            File.WriteAllText(configPath, json);
        }
    }

    public class Config
    {
        public string OutputFilePath { get; set; } = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "output.csv");
        public int MaxConcurrentScans { get; set; } = Environment.ProcessorCount;
        public int ScanTimeout { get; set; } = 5000; // milliseconds
        public bool EnableDetailedLogging { get; set; } = false;
        public DataCollectionSettings DataCollectionSettings { get; set; } = new DataCollectionSettings();
    }

    [Flags]
    public enum DataCollectionOptions
    {
        None = 0,
        Hostname = 1 << 0,
        Timestamp = 1 << 1,
        Status = 1 << 2,
        LastLoggedUser = 1 << 3,
        MachineType = 1 << 4,
        MachineSKU = 1 << 5,
        InstalledCoreSoftware = 1 << 6,
        RAMSize = 1 << 7,
        WindowsVersion = 1 << 8
    }

    public class DataCollectionSettings
    {
        public DataCollectionOptions Options { get; set; }

        public DataCollectionSettings()
        {
            // Set default options
            Options = DataCollectionOptions.Hostname | DataCollectionOptions.Timestamp | DataCollectionOptions.Status;
        }

        public bool ShouldCollect(DataCollectionOptions option)
        {
            return Options.HasFlag(option);
        }
    }
}