using System.Net.Http.Headers;

namespace NukiLockReport;

/// <summary>Result of one Nuki Web API call. Status 0 means the request never got a response.</summary>
public record ApiResult(int Status, string Body)
{
    public bool Ok => Status is >= 200 and < 300;
}

/// <summary>Read-only wrapper over https://api.nuki.io that never throws; every call returns status + body.</summary>
public sealed class NukiApi : IDisposable
{
    private readonly HttpClient _http;

    public NukiApi(string token)
    {
        _http = new HttpClient
        {
            BaseAddress = new Uri("https://api.nuki.io/"),
            Timeout = TimeSpan.FromSeconds(30)
        };
        _http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token.Trim());
        _http.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
    }

    public async Task<ApiResult> GetAsync(string path)
    {
        try
        {
            using var response = await _http.GetAsync(path);
            string text = await response.Content.ReadAsStringAsync();
            return new ApiResult((int)response.StatusCode, text);
        }
        catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException)
        {
            return new ApiResult(0, ex.Message);
        }
    }

    /// <summary>Plain-English meaning of the status codes the Nuki spec documents for these endpoints.</summary>
    public static string Explain(int status) => status switch
    {
        0 => "No response (network error or timeout).",
        401 => "Not authorized - the token is invalid, expired or revoked.",
        402 => "Account not paid - Nuki requires a paid subscription for this.",
        403 => "Forbidden - the token is missing a permission (check it in Nuki Web).",
        404 => "Not found - wrong smartlock ID.",
        _ when status >= 500 => "Nuki server error.",
        _ => ""
    };

    public void Dispose() => _http.Dispose();
}
