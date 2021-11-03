using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Web;

namespace MsGraph.Simple.Client {

  //-------------------------------------------------------------------------------------------------------------------
  //
  /// <summary>
  /// MS Graph Command
  /// </summary>
  //
  //-------------------------------------------------------------------------------------------------------------------

  public sealed class MsGraphCommand {
    #region Constants

    public const int MaximumPageSize = 999;

    #endregion Constants

    #region Private Data

    private static readonly Regex s_Version = new (@"^\/?v[0-9]+(?:\.[0-9]+)+");

    private string m_Version = "latest";

    #endregion Private Data

    #region Algorithm

    private string ActualVersion() {
      if ("latest".Equals(m_Version, StringComparison.OrdinalIgnoreCase))
        return "v1.0";
      else if ("beta".Equals(m_Version, StringComparison.OrdinalIgnoreCase))
        return "beta";
      else
        return $"v{m_Version}";
    }

    private string BuildAddress(string address) {
      // https://graph.microsoft.com/v1.0/

      if (string.IsNullOrWhiteSpace(address))
        return "";

      if (address.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
        return address;

      if (s_Version.IsMatch(address))
        return $"https://graph.microsoft.com/{address.TrimStart('/')}";

      if (address.StartsWith("beta/", StringComparison.OrdinalIgnoreCase))
        return $"https://graph.microsoft.com/{address}";

      if (address.StartsWith("/beta/", StringComparison.OrdinalIgnoreCase))
        return $"https://graph.microsoft.com{address}";

      if (address.StartsWith("/latest/", StringComparison.OrdinalIgnoreCase))
        address = address["/latest/".Length..]; // .Substring("/latest/".Length);
      else if (address.StartsWith("latest/", StringComparison.OrdinalIgnoreCase))
        address = address["latest/".Length..];  // .Substring("latest/".Length);

      if (address.StartsWith('/'))
        return $"https://graph.microsoft.com/{ActualVersion()}{address}";

      return $"https://graph.microsoft.com/{ActualVersion()}/{address}";
    }

    #endregion Algorithm

    #region Create

    /// <summary>
    /// Standard Constructor
    /// </summary>
    /// <param name="connection">Connection To use</param>
    public MsGraphCommand(MsGraphConnection connection) {
      Connection = connection ?? throw new ArgumentNullException(nameof(connection));
    }

    /// <summary>
    /// Standard Constructor
    /// </summary>
    /// <param name="connection">Connection To use</param>
    public MsGraphCommand(MsGraphConnection connection, string version)
      : this(connection) {
      Version = version;
    }

    #endregion Create

    #region Public

    /// <summary>
    /// Connection to use
    /// </summary>
    public MsGraphConnection Connection { get; }

    /// <summary>
    /// Version
    /// </summary>
    public string Version {
      get => m_Version;
      set {
        if (string.IsNullOrWhiteSpace(value))
          value = "latest";

        m_Version = "beta".Equals(value, StringComparison.OrdinalIgnoreCase) ||
                    "latest".Equals(value, StringComparison.OrdinalIgnoreCase) ||
                     Regex.IsMatch(value, @"[0-9]+(?:\.[0-9]+)+")
          ? value
          : throw new ArgumentException($"Version ({value}) is not supported", nameof(value));
      }
    }

    /// <summary>
    /// Perform Request
    /// </summary>
    public async Task<HttpResponseMessage> PerformRequestAsync(string address,
                                                               HttpRequestMessage request,
                                                               CancellationToken token) {
      if (request is null)
        throw new ArgumentNullException(nameof(request));

      string bearer = await Connection.AccessToken.ConfigureAwait(false);

      request.Headers.Add(HttpRequestHeader.Authorization.ToString(), $"Bearer {bearer}");

      if (request.RequestUri is null) {
        if (address is null)
          throw new ArgumentNullException(nameof(address));

        request.RequestUri = new Uri(BuildAddress(address));
      }

      return await MsGraphConnection.Client.SendAsync(request, token).ConfigureAwait(false);
    }

    /// <summary>
    /// Perform Request
    /// </summary>
    public async Task<HttpResponseMessage> PerformRequestAsync(string address,
                                                               HttpRequestMessage request) =>
      await PerformRequestAsync(address, request, CancellationToken.None).ConfigureAwait(false);

    /// <summary>
    /// Perform 
    /// </summary>
    public async Task<HttpResponseMessage> PerformStreamAsync(string address,
                                                              Stream stream,
                                                              HttpMethod method,
                                                              string header,
                                                              CancellationToken token) {
      if (address is null)
        throw new ArgumentNullException(nameof(address));
      if (stream is null)
        throw new ArgumentNullException(nameof(stream));

      address = BuildAddress(address);

      string bearer = await Connection.AccessToken.ConfigureAwait(false);

      using var req = new HttpRequestMessage {
        Method = method,
        RequestUri = new Uri(address),
        Headers = {
          { HttpRequestHeader.Authorization.ToString(), $"Bearer {bearer}" },
          { HttpRequestHeader.ContentType.ToString(), "application/octet-stream" },
        },

        Content = new StreamContent(stream)
      };

      req.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");

      if (!string.IsNullOrWhiteSpace(header))
        req.Headers.Add(HttpRequestHeader.Accept.ToString(), header?.Trim());

      return await MsGraphConnection.Client.SendAsync(req, token).ConfigureAwait(false);
    }

    /// <summary>
    /// Perform 
    /// </summary>
    public async Task<HttpResponseMessage> PerformStreamAsync(string address,
                                                              Stream stream,
                                                              HttpMethod method,
                                                              string header) =>
      await PerformStreamAsync(address, stream, method, header, default);

    /// <summary>
    /// Perform 
    /// </summary>
    public async Task<HttpResponseMessage> PerformStreamAsync(string address,
                                                              Stream stream,
                                                              HttpMethod method,
                                                              CancellationToken token) =>
      await PerformStreamAsync(address, stream, method, default, token);

    /// <summary>
    /// Perform 
    /// </summary>
    public async Task<HttpResponseMessage> PerformStreamAsync(string address,
                                                              Stream stream,
                                                              HttpMethod method) =>
      await PerformStreamAsync(address, stream, method, default, default);

    /// <summary>
    /// Perform 
    /// </summary>
    public async Task<HttpResponseMessage> PerformAsync(string address,
                                                        string query,
                                                        HttpMethod method,
                                                        string header,
                                                        CancellationToken token) {
      if (address is null)
        throw new ArgumentNullException(nameof(address));

      address = BuildAddress(address);

      header = header?.Trim() ?? "application/json";
      query = string.IsNullOrWhiteSpace(query) ? "{}" : query;

      string bearer = await Connection.AccessToken.ConfigureAwait(false);

      using var req = new HttpRequestMessage {
        Method = method,
        RequestUri = new Uri(address),
        Headers = {
          { HttpRequestHeader.Authorization.ToString(), $"Bearer {bearer}" },
          { HttpRequestHeader.Accept.ToString(), string.IsNullOrWhiteSpace(header) ? "application/json" : header},
        },
        Content = new StringContent(query, Encoding.UTF8, "application/json")
      };

      return await MsGraphConnection.Client.SendAsync(req, token).ConfigureAwait(false);
    }

    /// <summary>
    /// Perform 
    /// </summary>
    public async Task<HttpResponseMessage> PerformAsync(string address,
                                                        string query,
                                                        HttpMethod method,
                                                        string header) =>
      await PerformAsync(address, query, method, header, CancellationToken.None);

    /// <summary>
    /// Perform 
    /// </summary>
    public async Task<HttpResponseMessage> PerformAsync(string address,
                                                        string query,
                                                        HttpMethod method,
                                                        CancellationToken token) =>
      await PerformAsync(address, query, method, null, token);

    /// <summary>
    /// Perform 
    /// </summary>
    public async Task<HttpResponseMessage> PerformAsync(string address,
                                                        string query,
                                                        HttpMethod method) =>
      await PerformAsync(address, query, method, null, CancellationToken.None);

    /// <summary>
    /// Read JSON
    /// </summary>
    public async Task<JsonDocument> ReadJsonAsync(string address,
                                                  string query,
                                                  CancellationToken token) {
      if (address is null)
        throw new ArgumentNullException(nameof(address));

      address = BuildAddress(address);

      query = string.IsNullOrWhiteSpace(query) ? "" : query;

      string bearer = await Connection.AccessToken.ConfigureAwait(false);

      using var req = new HttpRequestMessage {
        Method = string.IsNullOrWhiteSpace(query) ? HttpMethod.Get : HttpMethod.Post,
        RequestUri = new Uri(address),
        Headers = {
          { HttpRequestHeader.Authorization.ToString(), $"Bearer {bearer}" },
          { HttpRequestHeader.Accept.ToString(), "application/json" },
        },
        Content = new StringContent(query, Encoding.UTF8, "application/json")
      };

      var response = await MsGraphConnection.Client.SendAsync(req, token).ConfigureAwait(false);

      if (!response.IsSuccessStatusCode)
        throw new InvalidOperationException($"{response.StatusCode} : {response.ReasonPhrase}");

      using Stream stream = await response.Content.ReadAsStreamAsync(token).ConfigureAwait(false);

      return await JsonDocument.ParseAsync(stream, default, token).ConfigureAwait(false);
    }

    /// <summary>
    /// Read JSON
    /// </summary>
    public async Task<JsonDocument> ReadJsonAsync(string address, string query) =>
      await ReadJsonAsync(address, query, CancellationToken.None);

    /// <summary>
    /// Read JSON
    /// </summary>
    public async Task<JsonDocument> ReadJsonAsync(string address, CancellationToken token) =>
      await ReadJsonAsync(address, "", token);

    /// <summary>
    /// Read JSON
    /// </summary>
    public async Task<JsonDocument> ReadJsonAsync(string address) =>
      await ReadJsonAsync(address, "", CancellationToken.None);

    /// <summary>
    /// Read JSON (paged)
    /// </summary>
    public async IAsyncEnumerable<JsonDocument> ReadJsonPagedAsync(string address,
                                                                   string query,
                                                                   int pageSize,
                                                                   [EnumeratorCancellation]
                                                                   CancellationToken token) {
      if (address is null)
        throw new ArgumentNullException(nameof(address));

      if (pageSize <= 0 || pageSize > MaximumPageSize)
        pageSize = MaximumPageSize;

      address = BuildAddress(address);

      if (HttpUtility.ParseQueryString(address).Count > 0)
        address += $"&$top={pageSize}";
      else
        address += $"?$top={pageSize}";

      query = string.IsNullOrWhiteSpace(query) ? "" : query;

      while (address is not null) {
        string bearer = await Connection.AccessToken.ConfigureAwait(false);

        using var req = new HttpRequestMessage {
          Method = string.IsNullOrWhiteSpace(query) ? HttpMethod.Get : HttpMethod.Post,
          RequestUri = new Uri(address),
          Headers = {
          { HttpRequestHeader.Authorization.ToString(), $"Bearer {bearer}" },
          { HttpRequestHeader.Accept.ToString(), "application/json" },
        },
          Content = new StringContent(query, Encoding.UTF8, "application/json")
        };

        var response = await MsGraphConnection.Client.SendAsync(req, token).ConfigureAwait(false);

        if (!response.IsSuccessStatusCode)
          throw new InvalidOperationException($"{response.StatusCode} : {response.ReasonPhrase}");

        using Stream stream = await response.Content.ReadAsStreamAsync(token).ConfigureAwait(false);

        var json = await JsonDocument.ParseAsync(stream, default, token).ConfigureAwait(false);

        if (json.RootElement.TryGetProperty("@odata.nextLink", out var next))
          address = next.GetString();
        else
          address = null;

        yield return json;
      }
    }

    /// <summary>
    /// Read JSON (paged)
    /// </summary>
    public async IAsyncEnumerable<JsonDocument> ReadJsonPagedAsync(string address,
                                                                   string query,
                                                                   int pageSize) {
      await foreach (var item in ReadJsonPagedAsync(address, query, pageSize, CancellationToken.None).ConfigureAwait(false))
        yield return item;
    }

    /// <summary>
    /// Read JSON (paged)
    /// </summary>
    public async IAsyncEnumerable<JsonDocument> ReadJsonPagedAsync(string address,
                                                                   int pageSize) {
      await foreach (var item in ReadJsonPagedAsync(address, null, pageSize, CancellationToken.None).ConfigureAwait(false))
        yield return item;
    }

    /// <summary>
    /// Read JSON (paged)
    /// </summary>
    public async IAsyncEnumerable<JsonDocument> ReadJsonPagedAsync(string address) {
      await foreach (var item in ReadJsonPagedAsync(address, null, MaximumPageSize, CancellationToken.None).ConfigureAwait(false))
        yield return item;
    }

    /// <summary>
    /// Read JSON (paged)
    /// </summary>
    public async IAsyncEnumerable<JsonDocument> ReadJsonPagedAsync(string address,
                                                                   string query) {
      await foreach (var item in ReadJsonPagedAsync(address, query, MaximumPageSize, CancellationToken.None).ConfigureAwait(false))
        yield return item;
    }

    /// <summary>
    /// Read JSON (paged)
    /// </summary>
    public async IAsyncEnumerable<JsonDocument> ReadJsonPagedAsync(string address,
                                                                   string query,
                                                                   [EnumeratorCancellation]
                                                                   CancellationToken token) {
      await foreach (var item in ReadJsonPagedAsync(address, query, MaximumPageSize, token).ConfigureAwait(false))
        yield return item;
    }

    /// <summary>
    /// Read JSON (paged)
    /// </summary>
    public async IAsyncEnumerable<JsonDocument> ReadJsonPagedAsync(string address,
                                                                   [EnumeratorCancellation]
                                                                   CancellationToken token) {
      await foreach (var item in ReadJsonPagedAsync(address, null, MaximumPageSize, token).ConfigureAwait(false))
        yield return item;
    }

    /// <summary>
    /// Read JSON (paged)
    /// </summary>
    public async IAsyncEnumerable<JsonDocument> ReadJsonPagedAsync(string address,
                                                                   int pageSize,
                                                                   [EnumeratorCancellation]
                                                                   CancellationToken token) {
      await foreach (var item in ReadJsonPagedAsync(address, null, pageSize, token).ConfigureAwait(false))
        yield return item;
    }

    #endregion Public
  }

}
