using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;

namespace TestForm {

  public static class Experiments {

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

      Queue<string> agenda = new Queue<string>();

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
            if (options.HasFlag(SearchOption.AllDirectories))
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

  }


}
