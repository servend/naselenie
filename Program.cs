using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Globalization;
using System.Net.Http;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using System.Linq;

namespace CityPopulationFinder
{
    public class City
    {
        public string Name { get; set; }
        public double Latitude { get; set; }
        public double Longitude { get; set; }
        public int? Population { get; set; }
        public string DataSource { get; set; }
    }

    public class SearchStatistics
    {
        public int TotalCities { get; set; }
        public int FoundInWikidata { get; set; }
        public int FoundInWiderdataSearch { get; set; }
        public int FoundInOSM { get; set; }
        public int NotFound { get; set; }

        public void PrintStatistics()
        {
            Console.WriteLine("\nСтатистика поиска:");
            Console.WriteLine($"Всего населенных пунктов: {TotalCities}");
            Console.WriteLine($"Найдено в Wikidata по координатам: {FoundInWikidata}");
            Console.WriteLine($"Найдено в Wikidata по названию: {FoundInWiderdataSearch}");
            Console.WriteLine($"Найдено в OSM: {FoundInOSM}");
            Console.WriteLine($"Не найдено: {NotFound}");
        }
    }

    class Program
    {
        static async Task Main()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            try
            {
                Console.WriteLine("Начало работы программы...");
                var cities = ReadCitiesFromExcel(@"C:\Users\User\Desktop\Города.xlsx");
                Console.WriteLine($"Прочитано {cities.Count} населенных пунктов");

                var populationFinder = new PopulationFinder();
                await populationFinder.GetPopulationData(cities);

                await SaveResultsToExcel(cities, @"C:\Users\User\Desktop\Кусты_с_населением.xlsx");
                Console.WriteLine("Программа успешно завершена");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Критическая ошибка: {ex.Message}");
                Console.WriteLine($"Stack Trace: {ex.StackTrace}");
            }

            Console.WriteLine("Нажмите любую клавишу для завершения...");
            Console.ReadKey();
        }

        static List<City> ReadCitiesFromExcel(string filePath)
        {
            var cities = new List<City>();
            var culture = CultureInfo.InvariantCulture;

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                var rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    try
                    {
                        var longitudeCell = worksheet.Cells[row, 1].Value?.ToString();
                        var latitudeCell = worksheet.Cells[row, 2].Value?.ToString();
                        var name = worksheet.Cells[row, 3].Value?.ToString();

                        if (string.IsNullOrEmpty(name))
                            continue;

                        longitudeCell = longitudeCell?.Trim().Replace(",", ".");
                        latitudeCell = latitudeCell?.Trim().Replace(",", ".");

                        if (!double.TryParse(longitudeCell, NumberStyles.Any, culture, out double longitude))
                        {
                            Console.WriteLine($"Ошибка парсинга долготы в строке {row}: {longitudeCell}");
                            continue;
                        }

                        if (!double.TryParse(latitudeCell, NumberStyles.Any, culture, out double latitude))
                        {
                            Console.WriteLine($"Ошибка парсинга широты в строке {row}: {latitudeCell}");
                            continue;
                        }

                        cities.Add(new City
                        {
                            Name = name.Trim(),
                            Longitude = longitude,
                            Latitude = latitude
                        });

                        Console.WriteLine($"Успешно прочитано: {name} ({latitude}, {longitude})");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Ошибка чтения строки {row}: {ex.Message}");
                    }
                }
            }

            return cities;
        }

        static async Task SaveResultsToExcel(List<City> cities, string filePath)
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Результаты");

                worksheet.Cells[1, 1].Value = "Долгота";
                worksheet.Cells[1, 2].Value = "Широта";
                worksheet.Cells[1, 3].Value = "Название";
                worksheet.Cells[1, 4].Value = "Население";
                worksheet.Cells[1, 5].Value = "Источник данных";

                for (int i = 0; i < cities.Count; i++)
                {
                    worksheet.Cells[i + 2, 1].Value = cities[i].Longitude;
                    worksheet.Cells[i + 2, 2].Value = cities[i].Latitude;
                    worksheet.Cells[i + 2, 3].Value = cities[i].Name;
                    worksheet.Cells[i + 2, 4].Value = cities[i].Population;
                    worksheet.Cells[i + 2, 5].Value = cities[i].DataSource;

                    worksheet.Cells[i + 2, 1].Style.Numberformat.Format = "0.000000";
                    worksheet.Cells[i + 2, 2].Style.Numberformat.Format = "0.000000";
                    if (cities[i].Population.HasValue)
                    {
                        worksheet.Cells[i + 2, 4].Style.Numberformat.Format = "#,##0";
                    }
                }

                worksheet.Cells.AutoFitColumns();
                await package.SaveAsAsync(new FileInfo(filePath));
                Console.WriteLine($"Результаты сохранены в файл: {filePath}");
            }
        }
    }

    public class PopulationFinder
    {
        private static readonly HttpClient client = new HttpClient()
        {
            Timeout = TimeSpan.FromSeconds(30)
        };

        private const string WIKIDATA_ENDPOINT = "https://query.wikidata.org/sparql";
        private const string OSM_ENDPOINT = "https://overpass-api.de/api/interpreter";
        private readonly SearchStatistics statistics = new SearchStatistics();

        public async Task GetPopulationData(List<City> cities)
        {
            statistics.TotalCities = cities.Count;
            int processed = 0;

            foreach (var city in cities)
            {
                processed++;
                try
                {
                    // 1. Сначала пробуем Wikidata по координатам
                    var population = await GetPopulationFromWikidata(city);
                    if (population.HasValue)
                    {
                        city.Population = population;
                        city.DataSource = "Wikidata";
                        statistics.FoundInWikidata++;
                    }
                    else
                    {
                        // 2. Пробуем расширенный поиск в Wikidata
                        population = await GetPopulationFromWiderdataSearch(city);
                        if (population.HasValue)
                        {
                            city.Population = population;
                            city.DataSource = "Wikidata Search";
                            statistics.FoundInWiderdataSearch++;
                        }
                        else
                        {
                            // 3. Пробуем OSM
                            population = await GetPopulationFromOSM(city);
                            if (population.HasValue)
                            {
                                city.Population = population;
                                city.DataSource = "OSM";
                                statistics.FoundInOSM++;
                            }
                            else
                            {
                                statistics.NotFound++;
                            }
                        }
                    }

                    Console.WriteLine($"Обработано {processed}/{cities.Count}: {city.Name} - {(population.HasValue ? population.ToString() : "не найдено")}");
                    await Task.Delay(2000);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Ошибка при обработке {city.Name}: {ex.Message}");
                    LogToFile($"Ошибка при обработке {city.Name}: {ex.Message}");
                }
            }

            statistics.PrintStatistics();
        }

        private async Task<int?> GetPopulationFromWikidata(City city)
        {
            try
            {
                var query = $@"
                SELECT ?population WHERE {{
                  ?city wdt:P17 wd:Q159;
                        rdfs:label ""{city.Name}""@ru;
                        wdt:P625 ?coordinates;
                        wdt:P1082 ?population.
                  SERVICE wikibase:around {{ 
                    ?city wdt:P625 ?location . 
                    bd:serviceParam wikibase:center ""Point({city.Longitude} {city.Latitude})""^^geo:wktLiteral .
                    bd:serviceParam wikibase:radius ""5"" . 
                  }}
                }}
                ORDER BY DESC(?population)
                LIMIT 1";

                var request = new HttpRequestMessage(HttpMethod.Get,
                    $"{WIKIDATA_ENDPOINT}?query={Uri.EscapeDataString(query)}&format=json");

                request.Headers.Add("User-Agent", "CityPopulationBot/1.0");
                request.Headers.Add("Accept", "application/json");

                var response = await client.SendAsync(request);
                response.EnsureSuccessStatusCode();

                var content = await response.Content.ReadAsStringAsync();
                var json = JObject.Parse(content);

                var bindings = json["results"]["bindings"];
                if (bindings != null && bindings.HasValues)
                {
                    var populationValue = bindings[0]["population"]["value"].ToString();
                    if (int.TryParse(populationValue, out int pop))
                    {
                        return pop;
                    }
                }

                return null;
            }
            catch (Exception ex)
            {
                LogToFile($"Ошибка запроса к Wikidata для {city.Name}: {ex.Message}");
                return null;
            }
        }

        private async Task<int?> GetPopulationFromWiderdataSearch(City city)
        {
            try
            {
                var query = $@"
                SELECT ?city ?cityLabel ?population WHERE {{
                  ?city wdt:P17 wd:Q159;
                        wdt:P31 wd:Q16110;
                        wdt:P1082 ?population.
                  SERVICE wikibase:label {{ bd:serviceParam wikibase:language ""ru"". }}
                  FILTER(CONTAINS(LCASE(?cityLabel), LCASE(""{city.Name}""))).
                }}
                ORDER BY DESC(?population)
                LIMIT 1";

                var request = new HttpRequestMessage(HttpMethod.Get,
                    $"{WIKIDATA_ENDPOINT}?query={Uri.EscapeDataString(query)}&format=json");

                request.Headers.Add("User-Agent", "CityPopulationBot/1.0");
                request.Headers.Add("Accept", "application/json");

                var response = await client.SendAsync(request);
                response.EnsureSuccessStatusCode();

                var content = await response.Content.ReadAsStringAsync();
                var json = JObject.Parse(content);

                var bindings = json["results"]["bindings"];
                if (bindings != null && bindings.HasValues)
                {
                    var populationValue = bindings[0]["population"]["value"].ToString();
                    if (int.TryParse(populationValue, out int pop))
                    {
                        return pop;
                    }
                }

                return null;
            }
            catch (Exception ex)
            {
                LogToFile($"Ошибка расширенного поиска в Wikidata для {city.Name}: {ex.Message}");
                return null;
            }
        }

        private async Task<int?> GetPopulationFromOSM(City city)
        {
            try
            {
                var query = $@"[out:json];
area[name=""Россия""][admin_level=""2""]->.a;
(
  node(area.a)[place][name=""{city.Name}""];
  way(area.a)[place][name=""{city.Name}""];
  relation(area.a)[place][name=""{city.Name}""];
);
out body;";

                var content = new FormUrlEncodedContent(new[]
                {
                    new KeyValuePair<string, string>("data", query)
                });

                var response = await client.PostAsync(OSM_ENDPOINT, content);
                response.EnsureSuccessStatusCode();

                var jsonResponse = await response.Content.ReadAsStringAsync();
                var json = JObject.Parse(jsonResponse);

                foreach (var element in json["elements"])
                {
                    var tags = element["tags"];
                    if (tags != null && tags["population"] != null)
                    {
                        if (int.TryParse(tags["population"].ToString(), out int pop))
                        {
                            return pop;
                        }
                    }
                }

                return null;
            }
            catch (Exception ex)
            {
                LogToFile($"Ошибка запроса к OSM для {city.Name}: {ex.Message}");
                return null;
            }
        }

        private void LogToFile(string message)
        {
            try
            {
                File.AppendAllText("population_search.log", $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}\n");
            }
            catch { }
        }
    }
}
