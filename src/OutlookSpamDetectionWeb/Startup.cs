using System;
using System.Text.Json.Serialization;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.OpenApi.Models;

namespace OutlookSpamDetectionWeb
{
    public class Startup
    {
        public Startup(IConfiguration configuration)
        {
            Configuration = configuration;
        }

        public IConfiguration Configuration { get; }

        // This method gets called by the runtime. Use this method to add services to the container.
        public void ConfigureServices(IServiceCollection services)
        {
            services.AddControllers();
            this.ConfigureSwagger(services);
        }

        // This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
        public void Configure(IApplicationBuilder app, IWebHostEnvironment env)
        {
            app.UseSwagger();
            if (env.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
            }
            app.UseDefaultFiles();
            app.UseStaticFiles();
            app.UseHttpsRedirection();
            app.UseRouting();
            app.UseAuthorization();
            app.UseSwaggerUI(c =>
            {
                c.SwaggerEndpoint("/swagger/v1/swagger.json", "Outlook Spam Detection API V1");
            });

            app.UseEndpoints(endpoints =>
            {
                endpoints.MapControllers();
            });
        }

        private void ConfigureSwagger(IServiceCollection services)
        {
            services.AddSwaggerGen(c =>
            {
                c.SwaggerDoc("v1", new OpenApiInfo
                {
                    Version = "v1",
                    Title = "Outlook Spam Detection API",
                    Description = "Predict whether an email is spam",
                    Contact = new OpenApiContact
                    {
                        Name = "Juan Carlos Calderón Cuellar",
                        Email = "jc.calderon@outlook.com",
                        Url = new Uri("https://github.com/jc-calderon"),
                    },
                    License = new OpenApiLicense
                    {
                        Name = "MIT License",
                        Url = new Uri("https://github.com/jc-calderon/outlook-spam-detection/blob/master/LICENSE"),
                    }
                });
            });
        }
    }
}