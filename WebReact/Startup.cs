// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Authentication.OpenIdConnect;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.AspNetCore.SpaServices.ReactDevelopmentServer;
using Microsoft.Bot.Connector;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;

using ApplicationCore;
using ApplicationCore.Interfaces;
using Infrastructure.Identity;
using Infrastructure.Identity.Extensions;
using Infrastructure.GraphApi;
using Infrastructure.Services;
using WebReact.Interfaces;
using WebReact.Services;
using ApplicationCore.Services;
using Infrastructure.OfficeApi;
using WebReact.Helpers;

namespace WebReact
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
            // Add CORS exception to enable authentication in Teams Client
            //services.AddCors(options =>
            //{
            //    options.AddPolicy("AllowSpecificOrigins",
            //    builder =>
            //    {
            //        builder.WithOrigins("https://webreact20180403042343.azurewebsites.net", "https://login.microsoftonline.com");
            //    });

            //    options.AddPolicy("AllowAllMethods",
            //        builder =>
            //        {
            //            builder.WithOrigins("https://webreact20180403042343.azurewebsites.net")
            //                   .AllowAnyMethod();
            //        });

            //    options.AddPolicy("ExposeResponseHeaders",
            //        builder =>
            //        {
            //            builder.WithOrigins("https://webreact20180403042343.azurewebsites.net")
            //                   .WithExposedHeaders("X-Frame-Options");
            //        });
            //});

            // Add in-mem cache service
            services.AddMemoryCache();

            // Credentials for bot authentication
            var credentialProvider = new StaticCredentialProvider(
                Configuration.GetSection("ProposalManagement:" + MicrosoftAppCredentials.MicrosoftAppIdKey)?.Value,
                Configuration.GetSection("ProposalManagement:" + MicrosoftAppCredentials.MicrosoftAppPasswordKey)?.Value);

            services.AddSingleton(typeof(ICredentialProvider), credentialProvider);

            // Add authentication services
            services.AddAuthentication(sharedOptions =>
            {
                sharedOptions.DefaultScheme = JwtBearerDefaults.AuthenticationScheme;
                //sharedOptions.DefaultChallengeScheme = JwtBearerDefaults.AuthenticationScheme;

            })
            .AddAzureAdBearer(options => Configuration.Bind("AzureAd", options));

            //services.AddAuthentication(sharedOptions =>
            //{
            //    sharedOptions.DefaultScheme = JwtBearerDefaults.AuthenticationScheme;
            //    sharedOptions.DefaultChallengeScheme = JwtBearerDefaults.AuthenticationScheme;

            //})
            //.AddBotAuthentication(credentialProvider);

            // Add MVC services
            services.AddMvc(options =>
            {
                options.Filters.Add(typeof(TrustServiceUrlAttribute));
            })
            .SetCompatibilityVersion(CompatibilityVersion.Version_2_1);


            // This sample uses an in-memory cache for tokens and subscriptions. Production apps will typically use some method of persistent storage.
            services.AddMemoryCache();
			services.AddSession();

			// Register configuration options
			services.Configure<AppOptions>(Configuration.GetSection("ProposalManagement"));

			// Add application infrastructure services.
			services.AddSingleton<IGraphAuthProvider, GraphAuthProvider>(); // Auth provider for Graph client, must be singleton
			services.AddScoped<IGraphClientAppContext, GraphClientAppContext>();
			services.AddScoped<IGraphClientUserContext, GraphClientUserContext>();
			services.AddTransient<IUserContext, UserIdentityContext>();
            services.AddScoped<IWordParser, WordParser>();

            // Add core services
            services.AddScoped<IOpportunityFactory, OpportunityFactory>();
            services.AddScoped<IOpportunityRepository, OpportunityRepository>();
            services.AddScoped<IUserProfileRepository, UserProfileRepository>();
			services.AddScoped<IDocumentRepository, DocumentRepository>();
			services.AddScoped<IRegionRepository, RegionRepository>();
			services.AddScoped<IIndustryRepository, IndustryRepository>();
			services.AddScoped<ICategoryRepository, CategoryRepository>();
			services.AddScoped<IRoleMappingRepository,RoleMappingRepository>();
			services.AddScoped<INotificationRepository, NotificationRepository>();
			services.AddScoped<GraphSharePointAppService>();
			services.AddScoped<GraphSharePointUserService>();
			services.AddScoped<GraphTeamUserService>();
			services.AddScoped<GraphUserAppService>();

			// FrontEnd services
			services.AddScoped<IOpportunityService, OpportunityService>();
			services.AddScoped<IDocumentService, DocumentService>();
			services.AddScoped<IUserProfileService, UserProfileService>();
			services.AddScoped<IRegionService, RegionService>();
			services.AddScoped<IIndustryService, IndustryService>();
			services.AddScoped<IRoleMappingService, RoleMappingService>();
			
			services.AddScoped<ICategoryService, CategoryService>();
			services.AddScoped<INotificationService, NotificationService>();
			services.AddScoped<IContextService, ContextService>();
            services.AddScoped<UserProfileHelpers>();
            services.AddScoped<OpportunityHelpers>();
            services.AddScoped<CardNotificationService>();

            // In production, the React files will be served from this directory
            services.AddSpaStaticFiles(configuration =>
			{
				configuration.RootPath = "ClientApp/build";
			});

			// This sample uses an in-memory cache for tokens and subscriptions. Production apps will typically use some method of persistent storage.
			services.AddMemoryCache();
			services.AddSession();
        }

		// This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
		public void Configure(IApplicationBuilder app, IHostingEnvironment env, ILoggerFactory loggerFactory)
		{           
			// Add the console logger.
			loggerFactory.AddConsole(Configuration.GetSection("Logging"));
			loggerFactory.AddDebug();

			// Configure error handling middleware.
            if (env.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
            }
            else
            {
                app.UseExceptionHandler("/Home/Error");
                app.UseHsts();
            }

            app.UseHttpsRedirection();

            // Add CORS policies
            //app.UseCors("ExposeResponseHeaders");

            // Add static files to the request pipeline.
            app.UseStaticFiles();
			app.UseSpaStaticFiles();

			// Add session to the request pipeline
			app.UseSession();

			// Add authentication to the request pipeline
			app.UseAuthentication();

			// Configure MVC routes
			app.UseMvc(routes =>
			{
				routes.MapRoute(
					name: "default",
					template: "{controller}/{action=Index}/{id?}");
			});

			app.UseSpa(spa =>
			{
				spa.Options.SourcePath = "ClientApp";

				if (env.IsDevelopment())
				{
					spa.UseReactDevelopmentServer(npmScript: "start");
				}
			});
		}
	}
}
