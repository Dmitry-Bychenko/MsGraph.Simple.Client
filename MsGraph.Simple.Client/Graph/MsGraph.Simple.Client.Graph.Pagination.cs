using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;

namespace MsGraph.Simple.Client.Graph {

  //-------------------------------------------------------------------------------------------------------------------
  //
  /// <summary>
  /// Paged Requests
  /// </summary>
  //
  //-------------------------------------------------------------------------------------------------------------------

  public static class MsGraphPaginationExtensions {
    #region Public

    /// <summary>
    /// Pagination
    /// </summary>
    /// <example>
    /// <code>
    /// var data = client
    ///    .Users
    ///    .Request()
    ///    .Expand("Manager,Extensions")
    ///    .EnumerateAsync<User>();
    ///    
    /// await foreach (var user in data.ConfigureAwait(false)) {...}
    /// </code>
    /// </example>
    public static async IAsyncEnumerable<T> EnumerateAsync<T>(this IBaseRequest request,
                                                                   int pageSize = default,
                                                                   [EnumeratorCancellation]
                                                                   CancellationToken token = default) {
      if (request is null)
        throw new ArgumentNullException(nameof(request));

      pageSize = pageSize <= 0 || pageSize > MsGraphCommand.MaximumPageSize
        ? MsGraphCommand.MaximumPageSize
        : pageSize;

      var topMethod = request.GetType().GetMethod("Top");

      if (topMethod is null)
        yield break;

      var topPage = topMethod.Invoke(request, new object[] { pageSize });

      while (topPage is not null) {

        var getAsyncMethod = topPage.GetType().GetMethod("GetAsync");

        if (getAsyncMethod is null)
          yield break;

        var objTask = getAsyncMethod.Invoke(topPage, new object[] { token });


        await (objTask as Task).ConfigureAwait(false);

        var propResult = objTask.GetType().GetProperty("Result");

        if (propResult is null)
          yield break;

        if (propResult.GetValue(objTask) is not IEnumerable<T> taskResult)
          yield break;

        foreach (T item in taskResult) {
          token.ThrowIfCancellationRequested();

          yield return item;
        }

        var nextProperty = taskResult.GetType().GetProperty("NextPageRequest");

        if (nextProperty is null)
          yield break;

        topPage = nextProperty.GetValue(taskResult);
      }
    }

    #endregion Public
  }


}
