using Microsoft.Graph;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;

namespace MsGraph.Simple.Client.Graph.Storage {

  //-------------------------------------------------------------------------------------------------------------------
  //
  /// <summary>
  /// Graph User (wrapper on User)
  /// </summary>
  //
  //-------------------------------------------------------------------------------------------------------------------

  public sealed class GraphUser : IEquatable<GraphUser> {
    #region Create

    internal GraphUser(Enterprise enterprise, User user) {
      Enterprise = enterprise ?? throw new ArgumentNullException(nameof(enterprise));
      User = user ?? throw new ArgumentNullException(nameof(user));
    }

    #endregion Create

    #region Public

    /// <summary>
    /// User
    /// </summary>
    public User User { get; private set; }

    /// <summary>
    /// Enterprise
    /// </summary>
    public Enterprise Enterprise { get; internal set; }

    /// <summary>
    /// Azure AD Connection
    /// </summary>
    public MsGraphConnection Connection => Enterprise?.Connection;

    /// <summary>
    /// Graph Service Client
    /// </summary>
    public GraphServiceClient Client => Enterprise?.Client;

    /// <summary>
    /// Is Modified
    /// </summary>
    public bool IsModified { get; set; }

    /// <summary>
    /// Manager
    /// </summary>
    public GraphUser Manager => Enterprise.FindUser(User.Manager?.Id);

    /// <summary>
    /// Hierarchy
    /// </summary>
    public IEnumerable<GraphUser> Hierarchy {
      get {
        for (var item = this; item is not null; item = item.Manager)
          yield return item;
      }
    }

    /// <summary>
    /// Subordinate
    /// </summary>
    public IEnumerable<GraphUser> Subordinate => Enterprise
      .Users
      .Where(user => ReferenceEquals(user.Manager, this));

    /// <summary>
    /// Root Manager
    /// </summary>
    public GraphUser RootManager {
      get {
        GraphUser result = this;

        while (true) {
          GraphUser manager = result.Manager;

          if (manager is null)
            return result;

          result = manager;
        }
      }
    }

    /// <summary>
    /// Delete (from Azure AD as well)
    /// </summary>
    public async Task<bool> Delete() {
      if (User is null)
        return false;

      await Client
        .Users[User.Id]
        .Request()
        .DeleteAsync();

      User = null;

      Enterprise?.CoreRemove(this);

      Enterprise = null;

      return true;
    }

    /// <summary>
    /// Remove 
    /// </summary>
    public bool Remove() {
      if (User is null)
        return false;

      User = null;

      Enterprise?.CoreRemove(this);

      Enterprise = null;

      return true;
    }

    /// <summary>
    /// 
    /// </summary>
    public async Task WriteStreamAsync(string address,
                                       Stream stream,
                                       HttpMethod method,
                                       string header = default,
                                       CancellationToken token = default) {
      var q = Enterprise.Connection.CreateCommand();

      using var response = await q.PerformStreamAsync(address, stream, method, header, token);

      if (!response.IsSuccessStatusCode)
        throw new InvalidOperationException($"Code: {response.StatusCode} ({(int)(response.StatusCode)}); reason: {response.ReasonPhrase}");
    }

    /// <summary>
    /// To String (Display Name)
    /// </summary>
    public override string ToString() => User?.DisplayName ?? "?";

    #endregion Public

    #region IEquatable<GraphUser>

    /// <summary>
    /// Equals
    /// </summary>
    public bool Equals(GraphUser other) {
      if (other is null)
        return false;

      return string.Equals(User.Id, other?.User?.Id, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Equals
    /// </summary>
    public override bool Equals(object obj) => obj is GraphUser other && Equals(other);

    /// <summary>
    /// Hash Code
    /// </summary>
    public override int GetHashCode() => User?.Id is null
      ? -1
      : User.Id.GetHashCode(StringComparison.OrdinalIgnoreCase);

    #endregion IEquatable<GraphUser>
  }

}
