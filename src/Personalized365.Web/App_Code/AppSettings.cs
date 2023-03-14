namespace Personalized365.Web
{
    public class AppSettings
    {
        public string TextAnalyticsCredential { get; set; }

        public string TextAnalyticsEndpoint { get; set; }

        public static AppSettings Load()
        {
            IConfiguration config = new ConfigurationBuilder()
                .AddJsonFile("appsettings.json", optional: false)
                .AddJsonFile($"appsettings.Development.json", optional: true)
                .Build();

            return config.GetRequiredSection("AppSettings").Get<AppSettings>()
                ?? throw new Exception("Could not load app settings. See README for configuration instructions.");
        }
    }
}
