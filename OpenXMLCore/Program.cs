using Microsoft.Extensions.Configuration;
using System.Reflection;

namespace OpenXMLCore
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string appsettingsFilePath = Directory.GetParent(AppDomain.CurrentDomain.BaseDirectory).Parent.Parent.Parent.FullName;

            IConfigurationRoot configuration = new ConfigurationBuilder()
                .SetBasePath(basePath: appsettingsFilePath)
                .AddJsonFile("appsettings.json", optional:false, reloadOnChange: true)
                .Build();

            Appsettings appSettings = configuration.GetSection("Appsettings").Get<Appsettings>();

            Console.WriteLine("please input the file name:");

            string fileName = Console.ReadLine() ?? "temp.docx";

            if (!fileName.EndsWith(".docx"))
            {
                fileName += ".docx";
            }

            appSettings.FilePath = Path.Combine(appSettings.FilePath, fileName);

            OpenXMLProcessor execute = new OpenXMLProcessor(appSettings.FilePath, appSettings.ImgPath);
            execute.RunProcess();
        }
    }
}
