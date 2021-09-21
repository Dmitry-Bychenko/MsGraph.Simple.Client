using Microsoft.Graph;

using System;
using System.Collections.Generic;
using System.Linq;
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

      if (string.IsNullOrEmpty(userId)) {
        var me = await client
          .Me
          .Request()
          .GetAsync(token)
          .ConfigureAwait(false);

        userId = me.Id;
      }

      if (string.IsNullOrWhiteSpace(directoryName))
        return false;

      var driveItem = new DriveItem {
        Name = directoryName,
        Folder = new Folder { },
        AdditionalData = new Dictionary<string, object>()  {
          {"@microsoft.graph.conflictBehavior", "replace"}
        }
      };

      try {
        await client
          .Users[userId]
          .Drive
          .Root
          .Children
          .Request()
          .AddAsync(driveItem, token)
          .ConfigureAwait(false);

        return true;
      }
      catch (ServiceException) {
        return false;
      }
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

      if (string.IsNullOrWhiteSpace(fileName))
        return false;

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
    public static async Task<bool> DeleteDirectoryAsync(this GraphServiceClient client,
                                                             string fileName,
                                                             CancellationToken token = default) =>
      await DeleteDirectoryAsync(client, null, fileName, token).ConfigureAwait(false);

    /// <summary>
    /// Enumerate Files
    /// </summary>
    public static async IAsyncEnumerable<string> EnumerateFilesAsync(this GraphServiceClient client,
                                                                          string userId,
                                                                          string directory,
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

      var path = client
          .Users[userId]
          .Drive
          .Root;

      if (!string.IsNullOrEmpty(directory))
        path = path.ItemWithPath(directory);

      IDriveItemChildrenCollectionPage data = null;

      try {
        data = await path
          .Children
          .Request()
          .GetAsync(token)
          .ConfigureAwait(false);
      }
      catch (ServiceException) {
        yield break;
      }

      var items = data
        .Where(item => item.Folder == null)
        .Select(item => item.Name);

      foreach (var item in items) {
        token.ThrowIfCancellationRequested();

        yield return item;
      }
    }

    /// <summary>
    /// Enumerate Files
    /// </summary>
    public static async IAsyncEnumerable<string> EnumerateFilesAsync(this GraphServiceClient client,
                                                                          string directory,
                                                                          [EnumeratorCancellation]
                                                                          CancellationToken token = default) {
      await foreach (var item in EnumerateFilesAsync(client, null, directory, token).ConfigureAwait(false))
        yield return item;
    }

    /// <summary>
    /// Enumerate Directories
    /// </summary>
    public static async IAsyncEnumerable<string> EnumerateDirectoriesAsync(this GraphServiceClient client,
                                                                                string userId,
                                                                                string directory,
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

      var path = client
        .Users[userId]
        .Drive
        .Root;

      if (!string.IsNullOrEmpty(directory))
        path = path.ItemWithPath(directory);

      var data = await path
        .Children
        .Request()
        .GetAsync(token)
        .ConfigureAwait(false);

      var items = data
        .Where(item => item.Folder != null)
        .Select(item => item.Name);

      foreach (var item in items) {
        token.ThrowIfCancellationRequested();

        yield return item;
      }
    }

    /// <summary>
    /// Enumerate Directories
    /// </summary>
    public static async IAsyncEnumerable<string> EnumerateDirectoriesAsync(this GraphServiceClient client,
                                                                                string directory,
                                                                                [EnumeratorCancellation]
                                                                                CancellationToken token = default) {
      await foreach (var item in EnumerateDirectoriesAsync(client, null, directory, token).ConfigureAwait(false))
        yield return item;
    }

    #endregion Public
  }

}
