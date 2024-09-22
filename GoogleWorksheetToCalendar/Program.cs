using CommandLine;
using GoogleServices.CustomServices;
using GoogleServices.GoogleAuthentication;
using GoogleServices.GoogleServices;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Threading.Tasks;

namespace GoogleWorksheetToCalendar
{
    internal class Program
    {
        private static void HandleParseError(IEnumerable<Error> errs)
        {
            //handle errors
        }

        private static void Main(string[] args)
        {
            var defaultCulture = new CultureInfo("en-GB");
            CultureInfo.DefaultThreadCurrentCulture = defaultCulture;
            CultureInfo.DefaultThreadCurrentUICulture = defaultCulture;

            try
            {
                Parser.Default.ParseArguments<Options>(args)
                  .WithParsed(RunWithOptions)
                  .WithNotParsed(HandleParseError);
                Console.WriteLine("Completed");
                Console.WriteLine("Press any key to continue...");
                Console.ReadKey();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine("Press any key to continue...");
                Console.ReadKey();
            }
        }

        private static async Task RunOnceAsync(Options options)
        {
            var googleCalendarService = new GoogleCalendarService();
            var googleCalendarsService = new GoogleCalendarsService();
            var googleSpreadsheetReadonlyService = new GoogleSpreadsheetReadonlyService();

            var googleAllScopesService = new GoogleAllScopesService();
            await GoogleOAuthAuthenticatorHelper.CreateAsync<GoogleAllScopesService>(googleAllScopesService);

            googleAllScopesService.ExecuteSomething();

            await GoogleOAuthAuthenticatorHelper.CreateAsync<Program>(googleSpreadsheetReadonlyService, googleCalendarService, googleCalendarsService);
            var calendar = await googleCalendarsService.CreateOrGetCalendarAsync(options.CalendarName, true);

            var customSpreadsheetService = new CustomSpreadsheetService(googleSpreadsheetReadonlyService, googleCalendarService);
            await customSpreadsheetService.WorksheetToCalendarAsync(options.SpreadsheetId, options.WorksheetName, calendar.Id, headerRowsCount: options.HeaderRowsCount);
        }

        private static void RunWithOptions(Options opts)
        {
            RunOnceAsync(opts).Wait();
        }

        private class Options
        {
            [Option('c', "calendar", Required = true, HelpText = "The name of the calendar to be produced, e.g. 'Arin's Holiday'")]
            public string CalendarName { get; set; }

            [Option('h', "headers", Required = false, HelpText = "The integer number of header rows in the spreadsheet, default=1")]
            public int HeaderRowsCount { get; set; } = 1;

            [Option('i', "spreadsheetid", Required = true, HelpText = "Google spreadsheet id from url, e.g '166KxWAwDKeMagoVh6RGdrc8BmzIaNmgM7i8W9IDCT7A'")]
            public string SpreadsheetId { get; set; }

            [Option('w', "worksheet", Required = true, HelpText = "The name of the worksheet from which the calendar is to be produced, e.g. 'ArinsHoliday'")]
            public string WorksheetName { get; set; }
        }
    }
}