using AspNetCoreMicrosoftGraph.Server.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Identity.Web;
using Microsoft.Identity.Web.UI;

namespace AspNetCoreMicrosoftGraph.Server;

public class Startup
{
    public Startup(IConfiguration configuration)
    {
        Configuration = configuration;
    }

    public IConfiguration Configuration { get; }

    public void ConfigureServices(IServiceCollection services)
    {
        services.AddScoped<MicrosoftGraphDelegatedClient>();
        services.AddScoped<EmailService>();
        services.AddScoped<TeamsService>();

        services.AddScoped<MicrosoftGraphApplicationClient>();
   
        services.AddAntiforgery(options =>
        {
            options.HeaderName = "X-XSRF-TOKEN";
            options.Cookie.Name = "__Host-X-XSRF-TOKEN";
            options.Cookie.SameSite = Microsoft.AspNetCore.Http.SameSiteMode.Strict;
            options.Cookie.SecurePolicy = Microsoft.AspNetCore.Http.CookieSecurePolicy.Always;
        });

        services.AddHttpClient();
        services.AddOptions();

        var scopes = Configuration.GetValue<string>("DownstreamApi:Scopes");
        string[]? initialScopes = scopes?.Split(' ');
        if(scopes == null) throw new ArgumentNullException(nameof(scopes));

        services.AddMicrosoftIdentityWebAppAuthentication(Configuration)
            .EnableTokenAcquisitionToCallDownstreamApi(initialScopes)
            .AddMicrosoftGraph("https://graph.microsoft.com/v1.0", initialScopes)
            .AddInMemoryTokenCaches();

        services.AddControllersWithViews(options =>
            options.Filters.Add(new AutoValidateAntiforgeryTokenAttribute()));

        services.AddRazorPages().AddMvcOptions(options =>
        {
            //var policy = new AuthorizationPolicyBuilder()
            //    .RequireAuthenticatedUser()
            //    .Build();
            //options.Filters.Add(new AuthorizeFilter(policy));
        }).AddMicrosoftIdentityUI();
    }

    public void Configure(IApplicationBuilder app, IWebHostEnvironment env)
    {
        if (env.IsDevelopment())
        {
            app.UseDeveloperExceptionPage();
            app.UseWebAssemblyDebugging();
        }
        else
        {
            app.UseExceptionHandler("/Error");
        }

        app.UseSecurityHeaders(
            SecurityHeadersDefinitions.GetHeaderPolicyCollection(env.IsDevelopment(),
                Configuration["AzureAd:Instance"]));

        app.UseHttpsRedirection();
        app.UseBlazorFrameworkFiles();
        app.UseStaticFiles();

        app.UseRouting();
        app.UseAuthentication();
        app.UseAuthorization();

        app.UseEndpoints(endpoints =>
        {
            endpoints.MapRazorPages();
            endpoints.MapControllers();
            endpoints.MapFallbackToPage("/_Host");
        });
    }
}
