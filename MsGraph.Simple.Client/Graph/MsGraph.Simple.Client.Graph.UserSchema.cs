using Microsoft.Graph;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text.Json;

using System.Threading;
using System.Threading.Tasks;

namespace MsGraph.Simple.Client.Graph {

  //-------------------------------------------------------------------------------------------------------------------
  //
  /// <summary>
  /// Schema Builder 
  /// </summary>
  //
  //-------------------------------------------------------------------------------------------------------------------

  public static class UserSchema {
    #region Public

    /// <summary>
    /// Create Extension if it doesn't exist
    /// </summary>
    public static async Task<bool> CreateExtensionAsync(this GraphServiceClient client,
                                                             string userId,
                                                             string extensionName,
                                                             IReadOnlyDictionary<string, object> fieldsAndValues,
                                                             CancellationToken token = default) {
      if (client is null)
        return false;

      if (fieldsAndValues is null || fieldsAndValues.Count <= 0)
        return false;

      if (string.IsNullOrEmpty(userId))
        return false;

      if (string.IsNullOrEmpty(extensionName))
        return false;

      Dictionary<string, object> doc = fieldsAndValues
        .ToDictionary(pair => pair.Key, pair => pair.Value);

      doc.TryAdd("@odata.type", "microsoft.graph.openTypeExtension");
      doc.TryAdd("extensionName", extensionName);

      string query = JsonSerializer.Serialize(doc);

      if (client.AuthenticationProvider is not MsGraphConnection connection)
        return false;

      var q = connection.CreateCommand();

      var result = await q
        .PerformAsync($"users/{userId}/extensions", query, HttpMethod.Post, token)
        .ConfigureAwait(false);

      return result.IsSuccessStatusCode;
    }

    /// <summary>
    /// Create Extension if it doesn't exist
    /// </summary>
    public static async Task<bool> DropExtensionAsync(this GraphServiceClient client,
                                                           string userId,
                                                           string extensionName,
                                                           CancellationToken token = default) {
      if (client is null)
        return false;

      if (string.IsNullOrEmpty(userId))
        return false;

      if (string.IsNullOrEmpty(extensionName))
        return false;

      string query = JsonSerializer.Serialize(new Dictionary<string, object>() {
        { "@odata.type", "microsoft.graph.openTypeExtension" },
        { "extensionName", userId },
      });

      if (client.AuthenticationProvider is not MsGraphConnection connection)
        return false;

      var q = connection.CreateCommand();

      var result = await q
        .PerformAsync($"users/{userId}/extensions/{extensionName}", query, HttpMethod.Delete, token)
        .ConfigureAwait(false);

      return result.IsSuccessStatusCode;
    }

    #endregion Public
  }

}
