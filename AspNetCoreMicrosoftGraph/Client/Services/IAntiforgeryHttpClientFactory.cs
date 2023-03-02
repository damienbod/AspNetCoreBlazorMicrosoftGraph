namespace AspNetCoreMicrosoftGraph.Server.Client.Services;

public interface IAntiforgeryHttpClientFactory
{
    Task<HttpClient> CreateClientAsync(string clientName = "authorizedClient");
}
