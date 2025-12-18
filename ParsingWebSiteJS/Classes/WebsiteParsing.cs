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
            log.Info("Создание драйвера Chrome");

            try
            {
                ChromeOptions options = new ChromeOptions();
                log.Debug("ChromeOptions: headless режим");
                options.AddArgument("--headless");

                log.Info("Инициализация ChromeDriver");
                IWebDriver driver = new ChromeDriver(options);

                driver.Manage().Window.Maximize();

                Console.WriteLine("Selenium WebDriver для Chrome успешно инициализирован");
                log.Info("ChromeDriver успешно инициализирован");
                return driver;
            }
            catch (Exception ex)
            {
                log.Error($"Ошибка создания ChromeDriver. Детали: {ex.Message}");
                Console.WriteLine($"Ошибка создания ChromeDriver: {ex.Message}");
                throw new Exception("Проверьте установлен ли у вас браузер Chrome и проверьте установлена ли у вас версия браузера 143.0.7499.41.", ex);
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
            log.Info("Запуск парсинга с использованием Selenium");
            string urls = "https://clientportal.jse.co.za/reports/delta-option-and-structured-option-trades";
            try
            {
                driver = CreateDriver();
                log.Info("Драйвер Chrome установлен");

                
                log.Info($"Переход на страницу: {urls}");

                Console.WriteLine("Ожидание загрузки данных на странице. Примерно 15-30 секунд...");
                driver.Navigate().GoToUrl(urls);

                log.Info($"Страница загружена: {urls}");

                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(15));
                log.Debug("Ожидание загрузки таблицы (таймаут 15 сек)");

                IWebElement table = driver.FindElement(By.XPath("//table[@id='tableTrades']"));
                log.Info("Таблица tableTrades найдена");

                var cols = table.FindElements(By.TagName("th"));
                log.Debug($"Найдено элементов th: {cols.Count}");

                List<string> dataLinks = new List<string>();
                int countCol = cols.Count;

                if (cols.Count() > 0)
                {
                    log.Info($"Сбор заголовков таблицы: {cols.Count} колонок");
                    foreach (IWebElement col in cols)
                    {
                        dataLinks.Add(col.Text);
                    }
                    log.Info("Заголовки таблицы собраны");
                }
                else
                {
                    log.Warn("Заголовки таблицы не найдены");
                }

                var dataRows = table.FindElements(By.XPath(".//tr/td"));
                log.Debug($"Найдено элементов данных td: {dataRows.Count}");

                bool dataTrue = false;
                if (dataRows.Count > 0)
                {
                    log.Info($"Сбор данных из таблицы: {dataRows.Count} элементов");
                    foreach (IWebElement row in dataRows)
                    {
                        if (!string.IsNullOrEmpty(row.Text))
                        {
                            dataLinks.Add(row.Text);
                            dataTrue = true;
                        }
                    }
                    log.Info("Данные из таблицы собраны");
                }
                else
                {
                    log.Warn("Данные в таблице отсутствуют");
                    Console.WriteLine("Данных в таблице о сделках отсутствуют");
                }

                if (dataTrue)
                {
                    log.Info($"Экспорт данных в CSV. Элементов: {dataLinks.Count}, Колонок: {countCol}");
                    WorkingWithCSVFile.ExportToCSV(dataLinks, countCol);
                    log.Info("Данные успешно экспортированы в CSV");
                }
                else
                {
                    log.Info("Нет данных для экспорта в CSV");
                }

                log.Info($"Парсинг завершен. Получено элементов: {dataLinks.Count}");
                return dataLinks;
            }
            catch (WebDriverException ex)
            {
                log.Error($"Ошибка WebDriver при парсинге: {urls}", ex);
                throw new Exception($"Ошибка WebDriver: {ex.Message}", ex);
            }
            catch (TimeoutException ex)
            {
                log.Error($"Таймаут при ожидании элементов: {urls}", ex);
                throw new Exception($"Таймаут ожидания: {ex.Message}", ex);
            }
            catch (Exception ex)
            {
                log.Error($"Ошибка при парсинге сайта: {urls}", ex);
                throw new Exception($"Ошибка парсинга: {ex.Message}", ex);
            }
            finally
            {
                try
                {
                    if (driver != null)
                    {
                        log.Info("Закрытие драйвера Chrome");
                        driver.Quit();
                        log.Info("Драйвер Chrome закрыт");
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