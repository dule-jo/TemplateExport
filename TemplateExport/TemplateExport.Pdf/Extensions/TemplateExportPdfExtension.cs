using Microsoft.Extensions.DependencyInjection;
using TemplateExport.Pdf.Internals;

namespace TemplateExport.Pdf.Extensions;

public static class TemplateExportPdfExtension
{
    public static IServiceCollection AddTemplateExportPdf(this IServiceCollection services)
    {
        services.AddScoped<ITemplateExportPdf, TemplateExportPdf>();

        return services;
    }
}