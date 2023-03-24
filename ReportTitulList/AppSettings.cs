using YamlDotNet.Serialization;

namespace ReportTitulList;

public class AppSettings
{
    private const string SettingsFile = "appsettings.yml";

    public string FileName { get; set; } = "СД к отчетам";
    public string[] Years { get; set; } = {"2022-2023"};
    public int[] NumberOfSemesters { get; set; } = {1, 2};
    public string[] Disciplines { get; set; } = {"Устройства на основе ПЛИС"};
    public string[] Groups { get; set; } = {"СМ5-71"};
    public string[] WorkTypes { get; set; } = {"ДЗ", "ЛР"};

    public static AppSettings Read()
    {
        if (File.Exists(SettingsFile) is false)
            return new();

        IDeserializer deserializer = new DeserializerBuilder().Build();
        using StreamReader reader = new(SettingsFile);
        return deserializer.Deserialize<AppSettings>(reader);
    }

    public static void Save(AppSettings settings)
    {
        ISerializer serializer = new SerializerBuilder().Build();
        string yaml = serializer.Serialize(settings);
        File.WriteAllText(SettingsFile, yaml);
    }
}