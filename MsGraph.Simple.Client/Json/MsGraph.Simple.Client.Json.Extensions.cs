using System;
using System.Collections.Generic;
using System.Text.Json;

namespace MsGraph.Simple.Client.Json {

  //-------------------------------------------------------------------------------------------------------------------
  //
  /// <summary>
  /// JSON Element Extensions
  /// </summary>
  //
  //-------------------------------------------------------------------------------------------------------------------

  public static class JsonElementExtensions {
    #region Constant

    public static readonly JsonElement NullElement = JsonDocument.Parse("null").RootElement;

    #endregion Constant

    #region Public

    #region General

    /// <summary>
    /// Read
    /// </summary>
    /// <param name="parent">Parent Item</param>
    /// <param name="childrenNames">Names or indexes</param>
    public static JsonElement Read(this JsonElement parent, params string[] childrenNames) {
      if (childrenNames is null)
        throw new ArgumentNullException(nameof(childrenNames));

      JsonElement result = parent;

      foreach (string name in childrenNames)
        if (result.ValueKind == JsonValueKind.Object) {
          if (!result.TryGetProperty(name, out result))
            return NullElement;
        }
        else if (result.ValueKind == JsonValueKind.Array)
          if (int.TryParse(name, out int index) && index >= 0 && index < result.GetArrayLength())
            result = result[index];
          else
            return NullElement;
        else
          return NullElement;

      return result;
    }

    /// <summary>
    /// As Enumerable (Arrays only)
    /// </summary>
    /// <param name="parent">Parent Item</param>
    public static IEnumerable<JsonElement> AsEnumerable(this JsonElement parent) {
      if (parent.ValueKind != JsonValueKind.Array)
        yield break;

      using var en = parent.EnumerateArray();

      while (en.MoveNext())
        yield return en.Current;
    }

    /// <summary>
    /// Try Get Boolean
    /// </summary>
    public static bool TryGetBoolean(this JsonElement parent, out bool result) {
      if (parent.ValueKind == JsonValueKind.False) {
        result = false;

        return true;
      }

      if (parent.ValueKind == JsonValueKind.True) {
        result = true;

        return true;
      }

      result = default;

      return false;
    }

    #endregion General

    #region Reads

    /// <summary>
    /// Boolean Or Null
    /// </summary>
    public static bool? BooleanOrNull(this JsonElement item) {
      if (item.ValueKind == JsonValueKind.True)
        return true;
      if (item.ValueKind == JsonValueKind.False)
        return false;

      return null;
    }

    /// <summary>
    /// Sbyte Or Null
    /// </summary>
    public static sbyte? SbyteOrNull(this JsonElement item) {
      if (item.ValueKind != JsonValueKind.Number)
        return null;

      return !item.TryGetSByte(out var result) ? null : result;
    }

    /// <summary>
    /// Int16 Or Null
    /// </summary>
    public static short? Int16OrNull(this JsonElement item) {
      if (item.ValueKind != JsonValueKind.Number)
        return null;

      return !item.TryGetInt16(out var result) ? null : result;
    }

    /// <summary>
    /// Int32 Or Null
    /// </summary>
    public static int? Int32OrNull(this JsonElement item) {
      if (item.ValueKind != JsonValueKind.Number)
        return null;

      return !item.TryGetInt32(out var result) ? null : result;
    }

    /// <summary>
    /// Int64 Or Null
    /// </summary>
    public static long? Int64OrNull(this JsonElement item) {
      if (item.ValueKind != JsonValueKind.Number)
        return null;

      return !item.TryGetInt64(out var result) ? null : result;
    }

    /// <summary>
    /// Byte Or Null
    /// </summary>
    public static byte? ByteOrNull(this JsonElement item) {
      if (item.ValueKind != JsonValueKind.Number)
        return null;

      return !item.TryGetByte(out var result) ? null : result;
    }

    /// <summary>
    /// UInt16 Or Null
    /// </summary>
    public static ushort? UInt16OrNull(this JsonElement item) {
      if (item.ValueKind != JsonValueKind.Number)
        return null;

      return !item.TryGetUInt16(out var result) ? null : result;
    }

    /// <summary>
    /// UInt32 Or Null
    /// </summary>
    public static uint? UInt32OrNull(this JsonElement item) {
      if (item.ValueKind != JsonValueKind.Number)
        return null;

      return !item.TryGetUInt32(out var result) ? null : result;
    }

    /// <summary>
    /// UInt64 Or Null
    /// </summary>
    public static ulong? UInt64OrNull(this JsonElement item) {
      if (item.ValueKind != JsonValueKind.Number)
        return null;

      return !item.TryGetUInt64(out var result) ? null : result;
    }

    /// <summary>
    /// Single Or Null
    /// </summary>
    public static float? SingleOrNull(this JsonElement item) {
      if (item.ValueKind != JsonValueKind.Number)
        return null;

      return !item.TryGetSingle(out var result) ? null : result;
    }

    /// <summary>
    /// Double Or Null
    /// </summary>
    public static double? DoubleOrNull(this JsonElement item) {
      if (item.ValueKind != JsonValueKind.Number)
        return null;

      return !item.TryGetDouble(out var result) ? null : result;
    }

    /// <summary>
    /// Decimal Or Null
    /// </summary>
    public static Decimal? DecimalOrNull(this JsonElement item) {
      if (item.ValueKind != JsonValueKind.Number)
        return null;

      return !item.TryGetDecimal(out var result) ? null : result;
    }

    /// <summary>
    /// DateTime Or Null
    /// </summary>
    public static DateTime? DateTimeOrNull(this JsonElement item) {
      if (item.ValueKind != JsonValueKind.String)
        return null;

      return !item.TryGetDateTime(out var result) ? null : result;
    }

    /// <summary>
    /// DateTimeOffset Or Null
    /// </summary>
    public static DateTimeOffset? DateTimeOffsetOrNull(this JsonElement item) {
      if (item.ValueKind != JsonValueKind.String)
        return null;

      return !item.TryGetDateTimeOffset(out var result) ? null : result;
    }

    /// <summary>
    /// DateTimeOffset Or Null
    /// </summary>
    public static Guid? GuidOrNull(this JsonElement item) {
      if (item.ValueKind != JsonValueKind.String)
        return null;

      return !item.TryGetGuid(out var result) ? null : result;
    }

    /// <summary>
    /// String Or Null
    /// </summary>
    public static string StringOrNull(this JsonElement item) {
      if (item.ValueKind == JsonValueKind.String ||
          item.ValueKind == JsonValueKind.True ||
          item.ValueKind == JsonValueKind.False ||
          item.ValueKind == JsonValueKind.Number)
        return item.ToString();

      return null;
    }

    #endregion Reads

    #endregion Public
  }

}
