using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using OfficeOpenXml;
using AEVWebClient.Models;
using System.Text.Json;

namespace AEVWebClient.Services
{
    public class FolderMonitorService : BackgroundService
    {
        private readonly ILogger<FolderMonitorService> _logger;
        private readonly FolderMonitorSettings _settings;
        private FileSystemWatcher _fileWatcher;
        private readonly TimeSpan _debounceInterval = TimeSpan.FromMilliseconds(500);
        private DateTime _lastReadTime;

        public FolderMonitorService(ILogger<FolderMonitorService> logger, IOptions<FolderMonitorSettings> options)
        {
            _logger = logger;
            _settings = options.Value;
        }

        protected override Task ExecuteAsync(CancellationToken stoppingToken)
        {
            _logger.LogInformation("FolderMonitorService is starting.");

            _fileWatcher = new FileSystemWatcher
            {
                Path = _settings.FolderPath,
                Filter = "*.*", // Listen for changes to all files
                NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.FileName
            };

            _fileWatcher.Changed += OnChanged;
            _fileWatcher.Created += OnChanged;
            _fileWatcher.EnableRaisingEvents = true;

            return Task.CompletedTask;
        }

        private void OnChanged(object sender, FileSystemEventArgs e)
        {
            // Ignore files starting with '~' (temporary files)
            if (Path.GetFileName(e.FullPath).StartsWith("~"))
            {
                _logger.LogInformation($"Ignoring temporary file: {e.FullPath}");
                return;
            }

            // Debounce logic to prevent multiple triggers
            if (DateTime.UtcNow - _lastReadTime < _debounceInterval)
                return;

            _lastReadTime = DateTime.UtcNow;

            _logger.LogInformation($"File change detected: {e.FullPath}");

            try
            {
                // Find the only .xlsx file in the directory (excluding temporary files)
                var xlsxFile = Directory.GetFiles(_settings.FolderPath, "*.xlsx")
                                         .FirstOrDefault(file => !Path.GetFileName(file).StartsWith("~"));

                if (xlsxFile == null)
                {
                    _logger.LogWarning("No valid .xlsx file found in the directory.");
                    return;
                }

                _logger.LogInformation($"Processing Excel file: {xlsxFile}");
                ParseExcelFile(xlsxFile);
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error processing file: {ex.Message}");
            }
        }


        private void ParseExcelFile(string filePath)
        {
            int maxRetries = 100;
            int delayMilliseconds = 500;

            for (int attempt = 1; attempt <= maxRetries; attempt++)
            {
                try
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    using (var package = new ExcelPackage(new FileInfo(filePath)))
                    {
                        var worksheet = package.Workbook.Worksheets["COMBINED"];
                        if (worksheet == null)
                        {
                            _logger.LogWarning("Worksheet 'COMBINED' not found.");
                            return;
                        }

                        int rowCount = worksheet.Dimension.Rows;

                        var scheduledUnits = new List<ScheduledUnit>();
                        for (int row = 2; row <= rowCount; row++)
                        {
                            var jobNumber = worksheet.Cells[row, 6].Text;
                            var customer = worksheet.Cells[row, 7].Text;

                            if (string.IsNullOrWhiteSpace(jobNumber) || string.IsNullOrWhiteSpace(customer))
                                continue;

                            var unit = new ScheduledUnit
                            {
                                StartDate = TryParseDate(worksheet.Cells[row, 1].Text),
                                ProjectedDeliveryDate = TryParseDate(worksheet.Cells[row, 2].Text),
                                StartPoint = worksheet.Cells[row, 3].Text,
                                ValueStream = worksheet.Cells[row, 4].Text,
                                WorkOrder = worksheet.Cells[row, 5].Text,
                                JobNumber = jobNumber,
                                Customer = customer,
                                Box = worksheet.Cells[row, 8].Text,
                                Chassis = worksheet.Cells[row, 9].Text,
                                Indicator = worksheet.Cells[row, 10].Text,
                                Complete = TryParseBool(worksheet.Cells[row, 11].Text),
                                FirstDayOfProdWeek = TryParseDate(worksheet.Cells[row, 12].Text),
                                DayAndNumber = worksheet.Cells[row, 13].Text,
                                LineOrder = worksheet.Cells[row, 14].Text,
                                BuildNumber = worksheet.Cells[row, 15].Text
                            };

                            scheduledUnits.Add(unit);
                        }

                        _logger.LogInformation($"Parsed {scheduledUnits.Count} valid scheduled units.");
                        var json = JsonSerializer.Serialize(scheduledUnits);
                        _logger.LogInformation("Scheduled Units JSON:\n" + json);
                        return; // Exit on success
                    }
                }
                catch (IOException ex) when (attempt < maxRetries)
                {
                    _logger.LogWarning($"Attempt {attempt}: File is locked or not ready. Retrying in {delayMilliseconds}ms...");
                    Thread.Sleep(delayMilliseconds);
                }
                catch (Exception ex)
                {
                    _logger.LogError($"Error reading Excel file: {ex.Message}");
                    return;
                }
            }

            _logger.LogError($"Failed to process file {filePath} after {maxRetries} attempts.");
        }


        private DateTime? TryParseDate(string value)
        {
            return DateTime.TryParse(value, out var date) ? date : (DateTime?)null;
        }

        private bool? TryParseBool(string value)
        {
            return bool.TryParse(value, out var result) ? result : (bool?)null;
        }


        public override void Dispose()
        {
            _fileWatcher?.Dispose();
            base.Dispose();
        }
    }
}
