using Microsoft.Graph;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;

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

    /*

    POST https://graph.microsoft.com/v1.0/me/extensions
    {
         "@odata.type": "microsoft.graph.openTypeExtension",
         "extensionName": "Nedra.HR.Buddy",

         "employeeBuddy": null,
         "levelUpDate": null,
         "welcomeLetter": false
    }

    POST https://graph.microsoft.com/v1.0/me/extensions
    {
         "@odata.type": "microsoft.graph.openTypeExtension",
         "extensionName": "Nedra.HR.Birthday",

         "employeeBirthday": "1974-07-08T00:00:00.000Z",
         "employeeGreetings": ["a", "b", "c" ]
    }


    */

    /// <summary>
    /// Create Extension if it doesn't exist
    /// </summary>
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



      //JsonDocument doc = new JsonDocument();

      //JsonSerializer.Serialize()

      //doc.

      return true;
    }

    #endregion Public
  }

}
