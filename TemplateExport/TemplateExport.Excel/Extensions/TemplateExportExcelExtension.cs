using ExcelTemplateExport.Internals;
using Microsoft.Extensions.DependencyInjection;

namespace ExcelTemplateExport.Extensions;

public static class TemplateExportExcelExtension
{
    public static IServiceCollection AddTemplateExportExcel(this IServiceCollection services)
    {
        services.AddScoped<ITemplateExportExcel, TemplateExportExcel>();

        return services;
    }
}