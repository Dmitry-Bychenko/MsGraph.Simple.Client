using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;

namespace MsGraph.Simple.Client.Graph {

  public static class MsGraphPaginationExtensions {
    #region Public

    public static async IAsyncEnumerable<T> EnumerateAsync<T>(this ICollectionPage<T> page,
                                                                  [EnumeratorCancellation]
                                                                   CancellationToken token) {
      if (page is null)
        throw new ArgumentNullException(nameof(page));

      while (page is not null) {
        token.ThrowIfCancellationRequested();

        foreach (T item in page) {
          token.ThrowIfCancellationRequested();

          yield return item;
        }

        token.ThrowIfCancellationRequested();

        var prop = page
          .GetType()
          .GetProperty("NextPageRequest", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);

        if (prop is null)
          break;

        var propValue = prop.GetValue(page, null);

        if (propValue is null)
          break;

        var method = propValue
          .GetType()
          .GetMethods(BindingFlags.Public | BindingFlags.Instance)
          .Where(m => m.Name == "GetAsync")
          .Where(m => m.GetParameters().Length == 1)
          .FirstOrDefault();

        if (method is null)
          break;

        Task<ICollectionPage<T>> task = method.Invoke(page, new object[] { token }) as Task<ICollectionPage<T>>;

        page = await task.ConfigureAwait(false);
      }
    }

    public static async IAsyncEnumerable<T> EnumerateAsync<T>(this ICollectionPage<T> page) {
      await foreach (var item in EnumerateAsync<T>(page, CancellationToken.None).ConfigureAwait(false))
        yield return item;
    }

    #endregion Public
  }
}
