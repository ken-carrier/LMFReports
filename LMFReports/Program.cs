using System;
using System.Collections.Generic;
using System.IO;
using LMFReports.Builder;
using LMFReports.Models;
using LMFReports.Models.SheetRequests;
using LMFReports.Services;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;

namespace LMFReports
{
    class Program
    {
            
        static void Main(string[] args)
        {
            Console.WriteLine("Hit enter to start");

            ConsoleKeyInfo keyPress = Console.ReadKey(intercept: true);
            while (keyPress.Key != ConsoleKey.Enter)
            {
                Console.Write(keyPress.KeyChar.ToString().ToUpper());

                keyPress = Console.ReadKey(intercept: true);
            }

            try
            {
                var services = new ServiceCollection();
                services.AddTransient<ThreeYearMembershipReportService>();
                services.AddSingleton<IWorkbookBuilder, WorkbookBuilder>();
                services.AddSingleton<Func<ReportServiceType, IReportService>>(serviceProvider => key =>
                {
                    switch (key)
                    {
                        case ReportServiceType.ThreeYearMembershipReportService:
                            return serviceProvider.GetService<ThreeYearMembershipReportService>();
                        default:
                            throw new KeyNotFoundException();
                    }
                });

                IServiceProvider serviceProvider = services.BuildServiceProvider();
                var workbookBuilder = serviceProvider.GetService<IWorkbookBuilder>();
                workbookBuilder.BuildWorkbooks();
            }
            catch (Exception ex)
            {
                Console.Write(ex.ToString());
                while (keyPress.Key != ConsoleKey.Enter)
                {
                    Console.Write(keyPress.KeyChar.ToString().ToUpper());

                    keyPress = Console.ReadKey(intercept: true);
                }
                
            }
            Console.WriteLine("Done");
        }
    }
}


