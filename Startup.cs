using OfficeOpenXml;

public class Startup
{
    public void ConfigureServices(IServiceCollection services)
    {
        // Other service configurations

        // Set EPPlus license context
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }

    public void Configure(IApplicationBuilder app, IWebHostEnvironment env)
    {
        // Application configuration
    }
}
