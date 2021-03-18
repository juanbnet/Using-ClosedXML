using Microsoft.Extensions.Configuration;
using System;

namespace InfDiario
{
    public class Settings
    {
        IConfigurationRoot Configuration;
        public Settings()
        {
            var builder = new ConfigurationBuilder()
                  .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);
            Configuration = builder.Build();
        }

        public string getConnStr()
        {

            return Configuration["settings:connString"];

        }
      
    }
}
