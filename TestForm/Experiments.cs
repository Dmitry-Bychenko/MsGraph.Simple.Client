using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

using Microsoft.Graph;

using MsGraph.Simple.Client;
using MsGraph.Simple.Client.Graph;

namespace TestForm {
  /*
  public static class Experiments {

    public static async IAsyncEnumerable<T> EnumerateAsync<T>(this IBaseRequest request,
                                                                   int pageSize = default,
                                                                   [EnumeratorCancellation]
                                                                   CancellationToken token = default) {
      if (request is null)
        throw new ArgumentNullException(nameof(request));

      pageSize = pageSize <= 0 || pageSize > 999
        ? 100
        : pageSize;

      var topMethod = request.GetType().GetMethod("Top");

      if (topMethod is null)
        yield break;

      var topPage = topMethod.Invoke(request, new object[] { pageSize });

      while (topPage is not null) {

        var getAsyncMethod = topPage.GetType().GetMethod("GetAsync");

        if (getAsyncMethod is null)
          yield break;

        var objTask = getAsyncMethod.Invoke(topPage, new object[] { token});
          
      
        await (objTask as Task).ConfigureAwait(false);

        var propResult = objTask.GetType().GetProperty("Result");

        if (propResult is null)
          yield break;

        var taskResult = propResult.GetValue(objTask) as IEnumerable<T>;

        if (taskResult is null)
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
  }
  */
}
