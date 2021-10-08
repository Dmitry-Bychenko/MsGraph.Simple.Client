using Microsoft.Graph;

using MsGraph.Simple.Client.Graph.Storage;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;

namespace MsGraph.Simple.Client.Graph {

  //-------------------------------------------------------------------------------------------------------------------
  //
  /// <summary>
  /// Azure Active Directory User Extensions
  /// </summary>
  //
  //-------------------------------------------------------------------------------------------------------------------

  public static class AzureUserExtensions {
    #region Constant

    /// <summary>
    /// Null Element
    /// </summary>
    public static readonly JsonElement NullElement = JsonDocument.Parse("null").RootElement;

    #endregion Constant

    #region Public

    /// <summary>
    /// Hierarchy
    /// </summary>
    public static IEnumerable<User> Hierarchy(this User user) {
      for (User current = user; current is not null; current = current.Manager as User)
        yield return current;
    }

    /// <summary>
    /// Extension By Name (null when not found)
    /// </summary>
    public static Extension ExtensionByName(this User user, string name) {
      if (user is null)
        return null;

      if (string.IsNullOrEmpty(name))
        return null;

      if (user.Extensions is null)
        return null;

      foreach (var extension in user.Extensions)
        if (string.Equals(extension.Id, name, StringComparison.OrdinalIgnoreCase))
          return extension;

      return null;
    }

    /// <summary>
    /// Extension Value (ValueKind == JsonValueKind.Undefined when not found)
    /// </summary>
    public static JsonElement ExtensionValue(this User user, string extensionName, string valueName) {
      if (valueName is null)
        return NullElement;

      var extension = ExtensionByName(user, extensionName);

      if (extension is null)
        return NullElement;

      if (extension.AdditionalData.TryGetValue(valueName, out var value) && value is JsonElement result)
        return result;

      return NullElement;
    }

    /// <summary>
    /// Create Extension 
    /// </summary>
    public static async Task<bool> CreateExtension(this GraphUser user,
                                                  string name,
                                                  IDictionary<string, object> values) {
      if (user is null)
        return false;

      if (values is null || values.Count <= 0)
        return false;

      if (string.IsNullOrEmpty(user.User.Id))
        return false;

      Dictionary<string, object> doc = values
        .ToDictionary(pair => pair.Key, pair => pair.Value);

      doc.TryAdd("@odata.type", "microsoft.graph.openTypeExtension");
      doc.TryAdd("extensionName", name);

      string query = JsonSerializer.Serialize(doc);
      string address = $"users/{user.User.Id}/extensions";

      var q = new MsGraphCommand(user.Enterprise.Connection);

      var response = await q.PerformAsync(address, query, HttpMethod.Post);

      return response.IsSuccessStatusCode;
    }

    /// <summary>
    /// Create Extension 
    /// </summary>
    public static async Task<bool> CreateOrUpdateExtension(this GraphUser user,
                                                                string name,
                                                                IDictionary<string, object> values) {
      if (user is null)
        return false;

      if (values is null || values.Count <= 0)
        return false;

      if (string.IsNullOrEmpty(user.User.Id))
        return false;

      Dictionary<string, object> doc = values
        .ToDictionary(pair => pair.Key, pair => pair.Value);

      doc.TryAdd("@odata.type", "microsoft.graph.openTypeExtension");
      doc.TryAdd("extensionName", name);

      string query = JsonSerializer.Serialize(doc);
      string address = $"users/{user.User.Id}/extensions";

      var q = new MsGraphCommand(user.Enterprise.Connection);

      bool hasExt = user.User.ExtensionByName(name) is not null;

      var response = await q.PerformAsync(address, query, hasExt ? HttpMethod.Patch : HttpMethod.Post);

      return response.IsSuccessStatusCode;
    }

    /// <summary>
    /// Create Extension if it doesn't exist
    /// </summary>
    /*
    public static bool CreateExtension(this GraphServiceClient client,
                                            string userId,
                                            string name,
                                            IDictionary<string, object> values) {
      if (client is null)
        return false;

      if (values is null || values.Count <= 0)
        return false;

      if (string.IsNullOrEmpty(userId))
        return false;

      if (string.IsNullOrEmpty(name))
        return false;

      Dictionary<string, object> doc = values
        .ToDictionary(pair => pair.Key, pair => pair.Value);

      doc.TryAdd("@odata.type", "microsoft.graph.openTypeExtension");
      doc.TryAdd("extensionName", name);

      string query = JsonSerializer.Serialize(doc);



      //MsGraphCommand q = client.Co 

      //JsonDocument doc = new JsonDocument();

      //JsonSerializer.Serialize()

      //doc.

      return true;
    }

    */

    #endregion Public
  }

}
