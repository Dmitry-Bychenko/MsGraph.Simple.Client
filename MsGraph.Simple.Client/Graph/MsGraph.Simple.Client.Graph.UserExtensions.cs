using Microsoft.Graph;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.Json;
using System.Threading;
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
    /// Extension By Name (null when not found)
    /// </summary>
    public static Extension ExtensionByName(this User user, string name) {
      if (user is null)
        return null;

      if (string.IsNullOrEmpty(name))
        return null;

      if (user.Extensions is null)
        return null;

      foreach (var ext in user.Extensions)
        if (string.Equals(ext.Id, name, StringComparison.OrdinalIgnoreCase))
          return ext;

      return null;
    }

    /// <summary>
    /// Extension Value (ValueKind == JsonValueKind.Undefined when not found)
    /// </summary>
    public static JsonElement ExtensionValue(this User user, string extensionName, string valueName) {
      if (valueName is null)
        return NullElement;

      var ext = ExtensionByName(user, extensionName);

      if (ext is null)
        return NullElement;

      if (ext.AdditionalData.TryGetValue(valueName, out var value) && value is JsonElement result)
        return result;

      return NullElement;
    }

    #endregion Public
  }

}
