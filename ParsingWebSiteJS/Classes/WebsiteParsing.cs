using HtmlAgilityPack;
using log4net;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System.Collections.Concurrent;
using System.Data;
using System.Security.Policy;

namespace ParsingWebSite.Classes
{
    internal static class WebsiteParsing
    {
        private static readonly ILog log = LogManager.GetLogger(typeof(Program));
        /// <summary>
        /// Метод для создания веб-драйвера
        /// </summary>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        static IWebDriver CreateDriver()
        {
            try
            {
                ChromeOptions options = new ChromeOptions();
                // Нужно для того, чтобы не открывалось окно браузераа
                options.AddArgument("--headless");
                IWebDriver driver = new ChromeDriver(options);
                driver.Manage().Window.Maximize();
                Console.WriteLine("Selenium WebDriver для Chrome успешно инициализирован");
                log.Info("Selenium WebDriver для Chrome успешно инициализирован");
                return driver;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Проверьте установлен ли у вас браузер Chrome и проверьте установлена ли у вас вресия бразуера 143.0.7499.41. Подробнее об ошибке: " + ex);
                log.Error("Ошибка при создании драйвера Chrome", ex);
                throw new Exception("Проверьте установлен ли у вас браузер Chrome и проверьте установлена ли у вас вресия бразуера 143.0.7499.41. Подробнее об ошибке: ", ex);

            }
        }

        /// <summary>
        /// Получения данных с сайта с помощью Selenium
        /// </summary>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public static List<string> parsinWebSiteToSelenium()
        {
            IWebDriver driver = null;
            try
            {
                log.Info("Создание драйвера Chrome");
                driver = CreateDriver();
                log.Info("Установлен драйвер Chrome");
                Console.WriteLine("Ожидание загрузки данных на странице. Примерно 15-30 секунд...");
                string urls = "https://clientportal.jse.co.za/reports/delta-option-and-structured-option-trades";
                driver.Navigate().GoToUrl(urls);
                log.Info($"Переход на страницу {urls}");
                // Нужно для того, чтобы успели загрузиться данные на сайте, так как они загружаются динамически, а не хранятся на сайте
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(15));
                IWebElement table = driver.FindElement(By.XPath("//table[@id='tableTrades']"));
                
                // Для поиска заголовка
                var cols = table.FindElements(By.TagName("th"));

                log.Debug("Инициализация списка данных");
                List<string> dataLinks = new List<string>();
                int countCol = cols.Count;

                if (cols.Count() > 0)
                {
                    foreach (IWebElement col in cols)
                    {
                        dataLinks.Add(col.Text);
                    }
                    log.Info($"Найдено {cols.Count} заголовков таблицы");
                }
                else
                {
                    log.Warn("Заголовки таблицы не найдены");
                }

                // Для поиска строчек
                var dataRows = table.FindElements(By.XPath(".//tr/td"));
                bool dataTrue = false;
                if (dataRows.Count > 0)
                {
                    foreach (IWebElement row in dataRows)
                    {
                        if (!string.IsNullOrEmpty(row.Text))
                        {
                            dataLinks.Add(row.Text);
                            dataTrue = true;
                        }
                    }
                }
                else
                {
                    Console.WriteLine("Данных в таблице о сделках отсутствуют");
                    log.Warn("Данные в таблице о сделках отсутствуют");
                }

                log.Debug("Начало экспорта данных в CSV");
                try
                {

                    
                    if (dataTrue)
                    {
                        // Запуск метода из другого класса для сохранения результата в csv файл
                        WorkingWithCSVFile.ExportToCSV(dataLinks, countCol);
                        log.Info("Данные успешно получены и экспортированы");
                    }
                    else
                    {
                        log.Info("Данных нет, поэтому создание файла csv прервано!");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Ошибка при экспорте в CSV. Продробнее: ", ex);
                    log.Error("Ошибка при экспорте данных в CSV", ex);
                }

                return dataLinks;
            }
            catch (Exception ex)
            {
                log.Error("Критическая ошибка при парсинге веб-сайта", ex);
                throw new Exception("Ошибка при парсинге веб-сайта: ", ex);
            }
            finally
            {
                try
                {
                    if (driver != null)
                    {
                        driver.Quit();
                        log.Info("Драйвер Chrome успешно закрыт");
                    }
                }
                catch (Exception ex)
                {
                    log.Error("Ошибка при закрытии драйвера Chrome", ex);
                }
            }
        }
    }
}