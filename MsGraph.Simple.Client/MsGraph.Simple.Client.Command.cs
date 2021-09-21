using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Linq;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace MsGraph.Simple.Client {

  //-------------------------------------------------------------------------------------------------------------------
  //
  /// <summary>
  /// MS Graph Command
  /// </summary>
  //
  //-------------------------------------------------------------------------------------------------------------------

  public sealed class MsGraphCommand {
    #region Algorithm

    private static string BuildAddress(string address) {
      // https://graph.microsoft.com/v1.0/

      if (string.IsNullOrWhiteSpace(address))
        return "";

      if (address.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
        return address;

      if (address.StartsWith("v1.0/", StringComparison.OrdinalIgnoreCase))
        return $"https://graph.microsoft.com/{address}";

      if (address.StartsWith("/v1.0/", StringComparison.OrdinalIgnoreCase))
        return $"https://graph.microsoft.com{address}";

      if (address.StartsWith("beta/", StringComparison.OrdinalIgnoreCase))
        return $"https://graph.microsoft.com/{address}";

      if (address.StartsWith("/beta/", StringComparison.OrdinalIgnoreCase))
        return $"https://graph.microsoft.com{address}";

      if (address.StartsWith('/'))
        return $"https://graph.microsoft.com/v1.0{address}";

      return $"https://graph.microsoft.com/v1.0/{address}";
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

    #endregion Create

    #region Public

    /// <summary>
    /// Connection to use
    /// </summary>
    public MsGraphConnection Connection { get; }

    /// <summary>
    /// Perform 
    /// </summary>
    public async Task<HttpResponseMessage> Perform(string address,
                                                   string query,
                                                   HttpMethod method,
                                                   CancellationToken token) {
      if (address is null)
        throw new ArgumentNullException(nameof(address));

      address = BuildAddress(address);

      query = string.IsNullOrWhiteSpace(query) ? "{}" : query;

      string bearer = await Connection.AccessToken.ConfigureAwait(false);

      using var req = new HttpRequestMessage {
        Method = method,
        RequestUri = new Uri(address),
        Headers = {
          { HttpRequestHeader.Authorization.ToString(), $"Bearer {bearer}" },
          { HttpRequestHeader.Accept.ToString(), "application/json" },
        },
        Content = new StringContent(query, Encoding.UTF8, "application/json")
      };

      return await MsGraphConnection.Client.SendAsync(req, token).ConfigureAwait(false);
    }

    /// <summary>
    /// Perform 
    /// </summary>
    public async Task<HttpResponseMessage> Perform(string address,
                                                   string query,
                                                   HttpMethod method) =>
      await Perform(address, query, method, CancellationToken.None);

    /// <summary>
    /// Read JSON
    /// </summary>
    public async Task<JsonDocument> ReadJson(string address, 
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
    public async Task<JsonDocument> ReadJson(string address, string query) =>
      await ReadJson(address, query, CancellationToken.None);

    /// <summary>
    /// Read JSON
    /// </summary>
    public async Task<JsonDocument> ReadJson(string address, CancellationToken token) =>
      await ReadJson(address, "", token);

    /// <summary>
    /// Read JSON
    /// </summary>
    public async Task<JsonDocument> ReadJson(string address) =>
      await ReadJson(address, "", CancellationToken.None);

    #endregion Public
  }

}
