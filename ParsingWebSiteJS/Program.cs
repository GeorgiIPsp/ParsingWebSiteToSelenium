using log4net;
using ParsingWebSite.Classes;


ILog log = LogManager.GetLogger(typeof(Program));
ProgramLogging.ConfigureFileLogging();


try
{
    log.Debug($"Время запуска программы: {DateTime.Now}");
    var dataTable = WebsiteParsing.parsinWebSiteToSelenium();
}
catch (Exception ex)
{
    log.Error($"Ошибка: {ex.Message}");
    log.Debug("Детали ошибки:", ex);
}


Console.WriteLine("Проверьте файл логов в папке 'Logs'");
Console.ReadKey();
