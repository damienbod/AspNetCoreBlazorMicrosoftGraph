using AspNetCoreMicrosoftGraph.Client;
using AspNetCoreMicrosoftGraph.Server.Client.Services;
using Blazorise;
using Blazorise.Icons.FontAwesome;
using Blazorise.Icons.Material;
using Blazorise.Material;
using Microsoft.AspNetCore.Components.Authorization;
using Microsoft.AspNetCore.Components.WebAssembly.Hosting;
using Microsoft.Extensions.DependencyInjection.Extensions;

using System.Net.Http.Headers;

var builder = WebAssemblyHostBuilder.CreateDefault(args);
builder.Services.AddOptions();
builder.Services.AddAuthorizationCore();
builder.Services.TryAddSingleton<AuthenticationStateProvider, HostAuthenticationStateProvider>();
builder.Services.TryAddSingleton(sp => (HostAuthenticationStateProvider)sp.GetRequiredService<AuthenticationStateProvider>());
builder.Services.AddTransient<AuthorizedHandler>();

builder.Services
    .AddBlazorise(options => { options.Immediate = true; })
    .AddMaterialProviders()
    .AddMaterialIcons()
    .AddFontAwesomeIcons();

builder.RootComponents.Add<App>("#app");

builder.Services.AddHttpClient("default", client =>
{
    client.BaseAddress = new Uri(builder.HostEnvironment.BaseAddress);
    client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
});

builder.Services.AddHttpClient("authorizedClient", client =>
{
    client.BaseAddress = new Uri(builder.HostEnvironment.BaseAddress);
    client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
}).AddHttpMessageHandler<AuthorizedHandler>();

builder.Services.AddTransient(sp => sp.GetRequiredService<IHttpClientFactory>().CreateClient("default"));
builder.Services.AddTransient<IAntiforgeryHttpClientFactory, AntiforgeryHttpClientFactory>();

await builder.Build().RunAsync();
