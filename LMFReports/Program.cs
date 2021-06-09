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
            var services = new ServiceCollection();
            services.AddTransient<Membership7447Service>();
            services.AddSingleton<IWorkbookBuilder,WorkbookBuilder>();
            services.AddSingleton<Func<ReportServiceType, IReportService>>(serviceProvider => key =>
            {
                switch (key)
                {
                    case ReportServiceType.Membership7447Service:
                        return serviceProvider.GetService<Membership7447Service>();
                    default:
                        throw new KeyNotFoundException();
                }
            });

            IServiceProvider serviceProvider = services.BuildServiceProvider();
            var workbookBuilder = serviceProvider.GetService<IWorkbookBuilder>();
            workbookBuilder.BuildWorkbooks();
        }

       
    }
}


