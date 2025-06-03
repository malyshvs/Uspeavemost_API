
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;

using Uspevaemost_API.Models;
using Uspevaemost_API.Services;

namespace Uspevaemost_API
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var builder = WebApplication.CreateBuilder(args);

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;


            builder.Services.AddControllers();
            builder.Services.AddScoped<ReportService>();
            builder.Services.AddEndpointsApiExplorer();
            builder.Services.AddSwaggerGen();


            builder.WebHost.UseIISIntegration();


            var app = builder.Build();




            // Configure the HTTP request pipeline.
            if (app.Environment.IsDevelopment())
            {
                app.UseSwagger();
                app.UseSwaggerUI();
            }







            app.MapPost("/", (HttpContext context) =>
            {
                var user = context.User;

                if (user.Identity?.IsAuthenticated == true)
                {
                    return Results.Ok($"Добро пожаловать, {user.Identity.Name}!");
                }

                return Results.Unauthorized();
            });

     
           

            app.MapControllers();

            app.Run();
        }
    }
}
