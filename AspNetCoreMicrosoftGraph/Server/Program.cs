﻿namespace AspNetCoreMicrosoftGraph.Server;

public class Program
{
    public static void Main(string[] args)
    {
        CreateHostBuilder(args).Build().Run();
    }

    public static IHostBuilder CreateHostBuilder(string[] args) =>
         Host.CreateDefaultBuilder(args)
             .ConfigureWebHostDefaults(webBuilder =>
             {
                 webBuilder
                     .ConfigureKestrel(options => options.AddServerHeader = false)
                     .UseStartup<Startup>();
             });
}
