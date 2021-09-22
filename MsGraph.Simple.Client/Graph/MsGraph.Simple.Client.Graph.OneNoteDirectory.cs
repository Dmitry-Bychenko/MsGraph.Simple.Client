using Microsoft.Graph;

using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;

namespace MsGraph.Simple.Client.Graph {

  //-------------------------------------------------------------------------------------------------------------------
  //
  /// <summary>
  /// OneNote Directory
  /// </summary>
  //
  //-------------------------------------------------------------------------------------------------------------------

  public static class OneNoteDirectory {
    #region Public

    /// <summary>
    /// Create Directory (in OneNote)
    /// </summary>
    public static async Task<bool> CreateDirectoryAsync(this GraphServiceClient client,
                                                             string userId,
                                                             string directoryName,
                                                             CancellationToken token = default) {
      if (client is null)
        throw new ArgumentNullException(nameof(client));

      token.ThrowIfCancellationRequested();

      if (string.IsNullOrEmpty(directoryName))
        return false;

      if (string.IsNullOrEmpty(userId)) {
        var me = await client
          .Me
          .Request()
          .GetAsync(token)
          .ConfigureAwait(false);

        userId = me.Id;
      }

      string[] parts = directoryName.Split(
        new char[] { Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar }, StringSplitOptions.RemoveEmptyEntries);

      string currentPath = "";

      for (int i = 0; i < parts.Length; ++i) {
        var driveItem = new DriveItem {
          Name = parts[i],
          Folder = new Folder { },
          AdditionalData = new Dictionary<string, object>()  {
            {"@microsoft.graph.conflictBehavior", "replace"}
          }
        };

        try {
          var root = client
            .Users[userId]
            .Drive
            .Root;

          if (!string.IsNullOrEmpty(currentPath))
            root = root.ItemWithPath(currentPath);

          await root
            .Children
            .Request()
            .AddAsync(driveItem, token)
            .ConfigureAwait(false);

          currentPath = System.IO.Path.Combine(currentPath, parts[i]);
        }
        catch (ServiceException) {
          return false;
        }
      }

      return true;
    }

    /// <summary>
    /// Create Directory (in OneNote)
    /// </summary>
    public static async Task<bool> CreateDirectoryAsync(this GraphServiceClient client,
                                                             string directoryName,
                                                             CancellationToken token = default) =>
      await CreateDirectoryAsync(client, null, directoryName, token).ConfigureAwait(false);

    /// <summary>
    /// Delete Directory (in OneNote)
    /// </summary>
    public static async Task<bool> DeleteDirectoryAsync(this GraphServiceClient client,
                                                             string userId,
                                                             string directoryName,
                                                             CancellationToken token = default) {
      if (client is null)
        throw new ArgumentNullException(nameof(client));

      if (string.IsNullOrWhiteSpace(directoryName))
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
          .ItemWithPath(directoryName)
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
    public static async Task<bool> DeleteDirectoryAsync(this GraphServiceClient client,
                                                             string directoryName,
                                                             CancellationToken token = default) =>
      await DeleteDirectoryAsync(client, null, directoryName, token).ConfigureAwait(false);

    /// <summary>
    /// Enumerate Files
    /// </summary>
    public static async IAsyncEnumerable<string> EnumerateFilesAsync(this GraphServiceClient client,
                                                                          string userId,
                                                                          string path,
                                                                          Func<string, bool> filter = default,
                                                                          SearchOption options = default,
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

      var rootPath = client
          .Users[userId]
          .Drive
          .Root;

      Queue<string> agenda = new();

      agenda.Enqueue(path ?? "");

      while (agenda.Count > 0) {
        string currentPath = agenda.Dequeue();

        var currentRootPath = string.IsNullOrEmpty(currentPath)
          ? rootPath
          : rootPath.ItemWithPath(currentPath);

        IDriveItemChildrenCollectionPage data;

        try {
          data = await currentRootPath
            .Children
            .Request()
            .GetAsync(token)
            .ConfigureAwait(false);
        }
        catch (ServiceException) {
          yield break;
        }

        foreach (var item in data) {
          token.ThrowIfCancellationRequested();

          if (item.Folder is null) {
            if (filter is null || filter(item.Name))
              yield return Path.Combine(currentPath, item.Name);
          }
          else if (options == SearchOption.AllDirectories)
            agenda.Enqueue(Path.Combine(currentPath, item.Name));
        }
      }
    }

    /// <summary>
    /// Enumerate Files
    /// </summary>
    public static async IAsyncEnumerable<string> EnumerateFilesAsync(this GraphServiceClient client,
                                                                          string path,
                                                                          Func<string, bool> filter = default,
                                                                          SearchOption options = default,
                                                                          [EnumeratorCancellation]
                                                                          CancellationToken token = default) {
      await foreach (string file in EnumerateFilesAsync(client, null, path, filter, options, token).ConfigureAwait(false))
        yield return file;
    }

    /// <summary>
    /// Enumerate Directories
    /// </summary>
    public static async IAsyncEnumerable<string> EnumerateDirectoriesAsync(this GraphServiceClient client,
                                                                                string userId,
                                                                                string path,
                                                                                Func<string, bool> filter = default,
                                                                                SearchOption options = default,
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

      var rootPath = client
          .Users[userId]
          .Drive
          .Root;

      Queue<string> agenda = new();

      agenda.Enqueue(path ?? "");

      while (agenda.Count > 0) {
        string currentPath = agenda.Dequeue();

        var currentRootPath = string.IsNullOrEmpty(currentPath)
          ? rootPath
          : rootPath.ItemWithPath(currentPath);

        IDriveItemChildrenCollectionPage data;

        try {
          data = await currentRootPath
            .Children
            .Request()
            .GetAsync(token)
            .ConfigureAwait(false);
        }
        catch (ServiceException) {
          yield break;
        }

        foreach (var item in data) {
          token.ThrowIfCancellationRequested();

          if (item.Folder is not null) {
            if (options == SearchOption.AllDirectories)
              agenda.Enqueue(Path.Combine(currentPath, item.Name));

            if (filter is null || filter(item.Name))
              yield return Path.Combine(currentPath, item.Name);
          }
        }
      }
    }

    /// <summary>
    /// Enumerate Directories
    /// </summary>
    public static async IAsyncEnumerable<string> EnumerateDirectoriesAsync(this GraphServiceClient client,
                                                                                string path,
                                                                                Func<string, bool> filter = default,
                                                                                SearchOption options = default,
                                                                                [EnumeratorCancellation]
                                                                                CancellationToken token = default) {
      await foreach (string file in EnumerateDirectoriesAsync(client, null, path, filter, options, token).ConfigureAwait(false))
        yield return file;
    }

    #endregion Public
  }

}
