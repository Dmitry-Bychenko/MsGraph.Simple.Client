using Microsoft.Graph;

using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Threading.Tasks;

namespace MsGraph.Simple.Client.Graph.Storage {

  //-------------------------------------------------------------------------------------------------------------------
  //
  /// <summary>
  /// Enterprise (All Users)
  /// </summary>
  //
  //-------------------------------------------------------------------------------------------------------------------

  public sealed class Enterprise 
    : IReadOnlyList<GraphUser>, 
      IReadOnlyDictionary<string, GraphUser> {

    #region Private Data

    private static readonly IReadOnlyDictionary<string, bool> s_ExcludeProperties =
      new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase) {
        { "AboutMe", false },
        { "Birthday", false },
        { "DeviceEnrollmentLimit", false },
        { "HireDate", false },
        { "Interests", false },
        { "MailboxSettings", false },
        { "MySite", false },
        { "PastProjects", false },
        { "PreferredName", false },
        { "Responsibilities", false },
        { "Schools", false },
        { "Skills", false },
      };

    private readonly List<GraphUser> m_Users = new();

    private readonly Dictionary<string, GraphUser> m_UserDict = new(StringComparer.OrdinalIgnoreCase);

    private readonly ConcurrentDictionary<string, GraphUser> m_BookMarks = new (StringComparer.OrdinalIgnoreCase);

    #endregion Private Data

    #region Algorithm

    private static string AllFields() {
      var fields = typeof(User)
        .GetProperties()
        .Where(p => p.DeclaringType == typeof(User))
        .Select(p => p.Name)
        .Where(name => !s_ExcludeProperties.ContainsKey(name))
        .OrderBy(name => name, StringComparer.OrdinalIgnoreCase);

      return "Id," + string.Join(",", fields);
    }

    private async Task CoreLoadAll() {
      foreach (var user in m_Users)
        user.Enterprise = null;

      m_UserDict.Clear();
      m_Users.Clear();
      m_BookMarks.Clear();

      int pageSize = 999;

      var data = await Client
        .Users
        .Request()
        .Expand("Manager,Extensions")
        .Top(pageSize)
        .Select(AllFields())
        .GetAsync()
        .ConfigureAwait(false);

      while (true) {
        foreach (var user in data) {
          var item = new GraphUser(this, user);

          m_Users.Add(item);

          m_UserDict.TryAdd(item.User.Id.Trim(), item);
          m_UserDict.TryAdd(item.User.UserPrincipalName.Trim(), item);

          if (!string.IsNullOrWhiteSpace(item.User.DisplayName))
            m_UserDict.TryAdd(item.User.DisplayName.Trim(), item);

          if (!string.IsNullOrWhiteSpace(item.User.Mail))
            m_UserDict.TryAdd(item.User.Mail.Trim(), item);

          if (!string.IsNullOrWhiteSpace(item.User.EmployeeId))
            m_UserDict.TryAdd(item.User.EmployeeId.Trim(), item);
        }

        if (data.NextPageRequest is null)
          break;

        data = await data
          .NextPageRequest
          .GetAsync()
          .ConfigureAwait(false);
      }

      var me = await Client
        .Me
        .Request()
        .GetAsync()
        .ConfigureAwait(false);

      Me = m_UserDict[me.Id];

      m_Users.Sort((left, right) => string.Compare(left.User.DisplayName, right.User.DisplayName, StringComparison.OrdinalIgnoreCase));

      // Master Domain Computation
      foreach (var user in m_Users) {
        string name = user.User.UserPrincipalName;

        if (!string.IsNullOrEmpty(name)) {
          int p = name.IndexOf('@');

          if (p >= 0) {
            MasterDomain = name[(p + 1)..];

            break;
          }
        }
      }
    }

    private async Task CoreInitialize() {
      await Connection.ConnectAsync().ConfigureAwait(false);

      Client = await Connection.CreateGraphClientAsync();

      await CoreLoadAll().ConfigureAwait(false);
    }

    #endregion Algorithm

    #region Create

    private Enterprise(string connectionString) {
      Connection = new MsGraphConnection(connectionString);
    }

    /// <summary>
    /// Factory Method
    /// </summary>
    /// <param name="connectionString">Connection String to use</param>
    public static async Task<Enterprise> CreateAsync(string connectionString) {
      Enterprise result = new(connectionString);

      await result.CoreInitialize().ConfigureAwait(false);

      return result;
    }

    #endregion Create

    #region Public

    /// <summary>
    /// Azure AD Connection
    /// 
    /// Connection string
    /// 
    /// Tenant      - optional
    /// Application
    /// Redirect    - optional
    /// Login       - optional
    /// Password
    /// </summary>
    public MsGraphConnection Connection { get; }

    /// <summary>
    /// Graph Service Client
    /// </summary>
    public GraphServiceClient Client { get; private set; }

    /// <summary>
    /// Master Domain
    /// </summary>
    public string MasterDomain { get; private set; } = "";

    /// <summary>
    /// Me
    /// </summary>
    public GraphUser Me { get; private set; }

    /// <summary>
    /// Users
    /// </summary>
    public IReadOnlyList<GraphUser> Users => m_Users;

    /// <summary>
    /// Find User
    /// </summary>
    public GraphUser FindUser(string value) {
      if (string.IsNullOrWhiteSpace(value))
        return null;

      value = value.Trim();

      if (m_UserDict.TryGetValue(value, out var result))
        return result;

      if (m_BookMarks.TryGetValue(value, out result))
        return result;

      return null;
    }

    /// <summary>
    /// Load 
    /// </summary>
    public async Task LoadAsync() => await CoreLoadAll().ConfigureAwait(false); 

    /// <summary>
    /// Add Bookmark
    /// </summary>
    public bool AddBookmark(string mark, GraphUser user) {
      if (string.IsNullOrWhiteSpace(mark))
        return false;

      if (user is null || user.Enterprise != this)
        return false;

      return m_BookMarks.TryAdd(mark, user);
    }

    /// <summary>
    /// Delete Bookmark
    /// </summary>
    public bool DeleteBookmark(string mark) {
      if (string.IsNullOrWhiteSpace(mark))
        return false;

      return m_BookMarks.Remove(mark, out var _);
    }

    /// <summary>
    /// Clear Bookmarks
    /// </summary>
    public void ClearBookmarks() {
      m_BookMarks.Clear();
    }

    #endregion Public

    #region IReadOnlyList<GraphUser>

    /// <summary>
    /// Count
    /// </summary>
    public int Count => m_Users.Count;

    /// <summary>
    /// Graph user by id, mail, principal name
    /// </summary>
    public GraphUser this[string key] => FindUser(key);

    /// <summary>
    /// Indexer
    /// </summary>
    public GraphUser this[int index] => m_Users[index];

    /// <summary>
    /// Typed Enumerator
    /// </summary>
    public IEnumerator<GraphUser> GetEnumerator() => m_Users.GetEnumerator();

    /// <summary>
    /// Typed Enumerator
    /// </summary>
    IEnumerator IEnumerable.GetEnumerator() => m_Users.GetEnumerator();

    #endregion IReadOnlyList<GraphUser>

    #region IReadOnlyDictionary<string, GraphUser>

    /// <summary>
    /// Keys
    /// </summary>
    public IEnumerable<string> Keys => m_UserDict.Keys;

    /// <summary>
    /// Values
    /// </summary>
    public IEnumerable<GraphUser> Values => m_UserDict.Values;

    /// <summary>
    /// Contains Key
    /// </summary>
    public bool ContainsKey(string key) => FindUser(key) is not null;

    /// <summary>
    /// Try Get Value
    /// </summary>
    public bool TryGetValue(string key, [MaybeNullWhen(false)] out GraphUser value) {
      if (key is null) {
        value = null;

        return false;
      }

      return m_UserDict.TryGetValue(key, out value);
    }

    /// <summary>
    /// Enumerator
    /// </summary>
    IEnumerator<KeyValuePair<string, GraphUser>> IEnumerable<KeyValuePair<string, GraphUser>>.GetEnumerator() =>
      m_UserDict.GetEnumerator();

    #endregion IReadOnlyDictionary<string, GraphUser>
  }

}
