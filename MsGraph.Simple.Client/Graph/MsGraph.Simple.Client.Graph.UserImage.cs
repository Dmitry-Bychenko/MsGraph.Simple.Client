using Microsoft.Graph;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace MsGraph.Simple.Client.Json {

  //-------------------------------------------------------------------------------------------------------------------
  //
  /// <summary>
  /// User Images (as bytes)
  /// </summary>
  //
  //-------------------------------------------------------------------------------------------------------------------

  public static class UserImage {
    #region Public

    /// <summary>
    /// Get Image Data
    /// </summary>
    public static async Task<byte[]> ReadImageBytesAsync(this GraphServiceClient client,
                                                              string userId,
                                                              CancellationToken token = default) {
      if (client is null)
        throw new ArgumentNullException(nameof(client));

      token.ThrowIfCancellationRequested();

      if (string.IsNullOrEmpty(userId)) {
        var me = await client
          .Me
          .Request()
          .GetAsync(token)
          .ConfigureAwait(false);

        userId = me.Id;
      }

      using Stream stream = await client
        .Users[userId]
        .Photo
        .Content
        .Request()
        .GetAsync(token)
        .ConfigureAwait(false);

      if (stream is null)
        return Array.Empty<byte>();

      byte[] result = new byte[(int)(stream.Length)];

      await stream.ReadAsync(result.AsMemory(0, result.Length), token).ConfigureAwait(false);

      return result;
    }

    /// <summary>
    /// Get Image Data
    /// </summary>
    public static async Task<byte[]> ReadImageBytesAsync(this GraphServiceClient client,
                                                              CancellationToken token = default) =>
      await ReadImageBytesAsync(client, null, token).ConfigureAwait(false);

    /// <summary>
    /// Set Image Data
    /// </summary>
    public static async Task WriteImageByteAsync(this GraphServiceClient client,
                                                      string userId,
                                                      IEnumerable<byte> source,
                                                      CancellationToken token = default) {
      if (client is null)
        throw new ArgumentNullException(nameof(client));

      token.ThrowIfCancellationRequested();

      if (string.IsNullOrEmpty(userId)) {
        var me = await client
          .Me
          .Request()
          .GetAsync(token)
          .ConfigureAwait(false);

        userId = me.Id;
      }

      byte[] data = source is byte[] bt
        ? bt
        : source.ToArray();

      using var stream = new MemoryStream(data);

      await client
        .Users[userId]
        .Photo
        .Content
        .Request()
        .PutAsync(stream, token)
        .ConfigureAwait(false);
    }

    /// <summary>
    /// Get Image Stream
    /// </summary>
    public static async Task<Stream> ReadImageAsync(this GraphServiceClient client,
                                                         string userId,
                                                         CancellationToken token = default) {
      if (client is null)
        throw new ArgumentNullException(nameof(client));

      token.ThrowIfCancellationRequested();

      if (string.IsNullOrEmpty(userId)) {
        var me = await client
          .Me
          .Request()
          .GetAsync(token)
          .ConfigureAwait(false);

        userId = me.Id;
      }

      return await client
        .Users[userId]
        .Photo
        .Content
        .Request()
        .GetAsync(token)
        .ConfigureAwait(false);
    }

    /// <summary>
    /// Get Image Stream
    /// </summary>
    public static async Task<Stream> ReadImageAsync(this GraphServiceClient client,
                                                         CancellationToken token = default) =>
      await ReadImageAsync(client, null, token).ConfigureAwait(false);

    /// <summary>
    /// Set Image Data
    /// </summary>
    public static async Task WriteImageAsync(this GraphServiceClient client,
                                                  string userId,
                                                  Stream source,
                                                  CancellationToken token = default) {
      if (client is null)
        throw new ArgumentNullException(nameof(client));

      if (source is null)
        throw new ArgumentNullException(nameof(source));

      token.ThrowIfCancellationRequested();

      if (string.IsNullOrEmpty(userId)) {
        var me = await client
          .Me
          .Request()
          .GetAsync(token)
          .ConfigureAwait(false);

        userId = me.Id;
      }

      await client
        .Users[userId]
        .Photo
        .Content
        .Request()
        .PutAsync(source, token)
        .ConfigureAwait(false);
    }

    /// <summary>
    /// Set Image Data
    /// </summary>
    public static async Task WriteImageAsync(this GraphServiceClient client,
                                                  Stream source,
                                                  CancellationToken token = default) =>
      await WriteImageAsync(client, null, token).ConfigureAwait(false);

    #endregion Public
  }

}
