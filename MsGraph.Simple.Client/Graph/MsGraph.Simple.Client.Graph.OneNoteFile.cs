using Microsoft.Graph;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace MsGraph.Simple.Client.Graph {

  //-------------------------------------------------------------------------------------------------------------------
  //
  /// <summary>
  /// OneNote File
  /// </summary>
  //
  //-------------------------------------------------------------------------------------------------------------------

  public static class OneNoteFile {
    #region Public

    /// <summary>
    /// Delete File (in OneNote)
    /// </summary>
    public static async Task<bool> DeleteFileAsync(this GraphServiceClient client,
                                                        string userId,
                                                        string fileName,
                                                        CancellationToken token = default) {
      if (client is null)
        throw new ArgumentNullException(nameof(client));

      if (string.IsNullOrWhiteSpace(fileName))
        return false;

      if (string.IsNullOrEmpty(userId)) {
        var me = await client
          .Me
          .Request()
          .GetAsync(token)
          .ConfigureAwait(false);

        userId = me.Id;
      }

      try {
        await client
          .Users[userId]
          .Drive
          .Root
          .ItemWithPath(fileName)
          .Request()
          .DeleteAsync(token)
          .ConfigureAwait(false);

        return true;
      }
      catch (ServiceException) {
        return false;
      }
    }

    /// <summary>
    /// Delete Directory (in OneNote)
    /// </summary>
    public static async Task<bool> DeleteFileAsync(this GraphServiceClient client,
                                                        string fileName,
                                                        CancellationToken token = default) =>
      await DeleteFileAsync(client, null, fileName, token).ConfigureAwait(false);


    /// <summary>
    /// Read File
    /// </summary>
    public static async Task<byte[]> ReadAllBytes(this GraphServiceClient client,
                                                       string userId,
                                                       string fileName,
                                                       CancellationToken token = default) {
      if (client is null)
        throw new ArgumentNullException(nameof(client));

      if (string.IsNullOrEmpty(userId)) {
        var me = await client
          .Me
          .Request()
          .GetAsync(token)
          .ConfigureAwait(false);

        userId = me.Id;
      }

      token.ThrowIfCancellationRequested();

      if (string.IsNullOrWhiteSpace(fileName))
        return Array.Empty<byte>();

      try {
        var path = client
          .Users[userId]
          .Drive
          .Root;

        if (!string.IsNullOrEmpty(fileName))
          path = path.ItemWithPath(fileName);

        using var stream = await path
          .Content
          .Request()
          .GetAsync(token)
          .ConfigureAwait(false);

        byte[] data = new byte[stream.Length];

        for (int i = 0; i < data.Length; ++i)
          data[i] = (byte)(stream.ReadByte());

        return data;
      }
      catch (ServiceException) {
        return Array.Empty<byte>();
      }
    }

    /// <summary>
    /// Read File
    /// </summary>
    public static async Task<byte[]> ReadAllBytes(this GraphServiceClient client,
                                                       string fileName,
                                                       CancellationToken token = default) =>
      await ReadAllBytes(client, null, fileName, token).ConfigureAwait(false);

    /// <summary>
    /// Write file
    /// </summary>
    public static async Task<bool> WriteAllBytes(this GraphServiceClient client,
                                                      string userId,
                                                      string fileName,
                                                      IEnumerable<byte> data,
                                                      CancellationToken token = default) {
      if (client is null)
        throw new ArgumentNullException(nameof(client));

      if (string.IsNullOrEmpty(userId)) {
        var me = await client
          .Me
          .Request()
          .GetAsync(token)
          .ConfigureAwait(false);

        userId = me.Id;
      }

      if (string.IsNullOrWhiteSpace(fileName))
        return false;

      if (data is null)
        return false;

      string dir = System.IO.Path.GetDirectoryName(fileName);

      if (!string.IsNullOrWhiteSpace(dir)) {
        if (!await OneNoteDirectory.CreateDirectoryAsync(client, userId, dir, token).ConfigureAwait(false))
          return false;
      }

      token.ThrowIfCancellationRequested();

      try {
        using var stream = new MemoryStream(data.ToArray());

        var path = client
          .Users[userId]
          .Drive
          .Root;

        path = path.ItemWithPath(fileName);

        var result = await path
          .Content
          .Request()
          .PutAsync<DriveItem>(stream, token)
          .ConfigureAwait(false);

        return true;
      }
      catch (ServiceException) {
        return false;
      }
    }

    /// <summary>
    /// Write file
    /// </summary>
    public static async Task<bool> WriteAllBytes(this GraphServiceClient client,
                                                      string fileName,
                                                      IEnumerable<byte> data,
                                                      CancellationToken token = default) =>
      await WriteAllBytes(client, null, fileName, data, token).ConfigureAwait(false);

    /// <summary>
    /// Read All Text
    /// </summary>
    public static async Task<string> ReadAllText(this GraphServiceClient client,
                                                      string userId,
                                                      string fileName,
                                                      Encoding encoding = null,
                                                      CancellationToken token = default) {
      if (client is null)
        throw new ArgumentNullException(nameof(client));

      if (string.IsNullOrEmpty(userId)) {
        var me = await client
          .Me
          .Request()
          .GetAsync(token)
          .ConfigureAwait(false);

        userId = me.Id;
      }

      token.ThrowIfCancellationRequested();

      if (string.IsNullOrWhiteSpace(fileName))
        return "";

      encoding ??= Encoding.Default;

      try {
        var path = client
          .Users[userId]
          .Drive
          .Root;

        if (!string.IsNullOrEmpty(fileName))
          path = path.ItemWithPath(fileName);

        using var stream = await path
          .Content
          .Request()
          .GetAsync(token)
          .ConfigureAwait(false);

        using var reader = new StreamReader(stream, encoding);

        return await reader.ReadToEndAsync().ConfigureAwait(false);
      }
      catch (ServiceException) {
        return "";
      }
    }

    /// <summary>
    /// Read All Text
    /// </summary>
    public static async Task<string> ReadAllText(this GraphServiceClient client,
                                                      string fileName,
                                                      Encoding encoding = null,
                                                      CancellationToken token = default) =>
      await ReadAllText(client, null, fileName, encoding, token).ConfigureAwait(false);

    /// <summary>
    /// Read All Text
    /// </summary>
    public static async Task<string[]> ReadAllLines(this GraphServiceClient client,
                                                         string userId,
                                                         string fileName,
                                                         Encoding encoding = null,
                                                         CancellationToken token = default) {
      if (client is null)
        throw new ArgumentNullException(nameof(client));

      if (string.IsNullOrEmpty(userId)) {
        var me = await client
          .Me
          .Request()
          .GetAsync(token)
          .ConfigureAwait(false);

        userId = me.Id;
      }

      token.ThrowIfCancellationRequested();

      if (string.IsNullOrWhiteSpace(fileName))
        return Array.Empty<string>();

      encoding ??= Encoding.Default;

      try {
        var path = client
          .Users[userId]
          .Drive
          .Root;

        if (!string.IsNullOrEmpty(fileName))
          path = path.ItemWithPath(fileName);

        using var stream = await path
          .Content
          .Request()
          .GetAsync(token)
          .ConfigureAwait(false);

        using var reader = new StreamReader(stream, encoding);

        List<string> lines = new();

        for (string line = reader.ReadLine(); line is not null; line = reader.ReadLine()) {
          token.ThrowIfCancellationRequested();

          lines.Add(line);
        }

        return lines.ToArray();
      }
      catch (ServiceException) {
        return Array.Empty<string>();
      }
    }

    /// <summary>
    /// Read All Text
    /// </summary>
    public static async Task<string[]> ReadAllLines(this GraphServiceClient client,
                                                         string fileName,
                                                         Encoding encoding = null,
                                                         CancellationToken token = default) =>
      await ReadAllLines(client, null, fileName, encoding, token).ConfigureAwait(false);

    /// <summary>
    /// Read Lines
    /// </summary>
    public static async IAsyncEnumerable<string> ReadLines(this GraphServiceClient client,
                                                                string userId,
                                                                string fileName,
                                                                Encoding encoding = null,
                                                                [EnumeratorCancellation]
                                                                CancellationToken token = default) {
      if (client is null)
        throw new ArgumentNullException(nameof(client));

      if (string.IsNullOrEmpty(userId)) {
        var me = await client
          .Me
          .Request()
          .GetAsync(token)
          .ConfigureAwait(false);

        userId = me.Id;
      }

      token.ThrowIfCancellationRequested();

      if (string.IsNullOrWhiteSpace(fileName))
        yield break;

      encoding ??= Encoding.Default;

      Stream stream;

      try {
        var path = client
          .Users[userId]
          .Drive
          .Root;

        if (!string.IsNullOrEmpty(fileName))
          path = path.ItemWithPath(fileName);

        stream = await path
          .Content
          .Request()
          .GetAsync(token)
          .ConfigureAwait(false);
      }
      catch (ServiceException) {
        yield break;
      }

      using (stream) {
        using var reader = new StreamReader(stream, encoding);

        for (string line = reader.ReadLine(); line is not null; line = reader.ReadLine()) {
          token.ThrowIfCancellationRequested();

          yield return line;
        }
      }
    }

    /// <summary>
    /// Read Lines
    /// </summary>
    public static async IAsyncEnumerable<string> ReadLines(this GraphServiceClient client,
                                                                string fileName,
                                                                Encoding encoding = null,
                                                                [EnumeratorCancellation]
                                                                CancellationToken token = default) {
      await foreach (var item in ReadLines(client, null, fileName, encoding, token).ConfigureAwait(false))
        yield return item;
    }

    /// <summary>
    /// Write All Text
    /// </summary>
    public static async Task<bool> WriteAllText(this GraphServiceClient client,
                                                     string userId,
                                                     string fileName,
                                                     string text,
                                                     Encoding encoding = null,
                                                     CancellationToken token = default) {
      text ??= "";
      encoding ??= Encoding.Default;

      return await WriteAllBytes(client, userId, fileName, encoding.GetBytes(text), token).ConfigureAwait(false);
    }

    /// <summary>
    /// Write All Text
    /// </summary>
    public static async Task<bool> WriteAllText(this GraphServiceClient client,
                                                     string fileName,
                                                     string text,
                                                     Encoding encoding = null,
                                                     CancellationToken token = default) =>
      await WriteAllText(client, null, fileName, text, encoding, token).ConfigureAwait(false);

    /// <summary>
    /// Write All Text
    /// </summary>
    public static async Task<bool> WriteAllLines(this GraphServiceClient client,
                                                      string userId,
                                                      string fileName,
                                                      IEnumerable<string> lines,
                                                      Encoding encoding = null,
                                                      CancellationToken token = default) {
      if (lines is null)
        throw new ArgumentNullException(nameof(lines));

      return await WriteAllText(client, userId, fileName, string.Join(Environment.NewLine, lines), encoding, token).ConfigureAwait(false);
    }

    /// <summary>
    /// Write All Text
    /// </summary>
    public static async Task<bool> WriteAllLines(this GraphServiceClient client,
                                                      string fileName,
                                                      IEnumerable<string> lines,
                                                      Encoding encoding = null,
                                                      CancellationToken token = default) =>
      await WriteAllLines(client, null, fileName, lines, encoding, token).ConfigureAwait(false);

    #endregion Public
  }


}
