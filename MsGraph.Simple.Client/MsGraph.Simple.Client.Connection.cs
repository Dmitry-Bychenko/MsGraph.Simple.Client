using Microsoft.Graph;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace MsGraph.Simple.Client {

  //-------------------------------------------------------------------------------------------------------------------
  //
  /// <summary>
  /// Connect Event Args
  /// </summary>
  //
  //-------------------------------------------------------------------------------------------------------------------

  public sealed class ConnectingEventArgs : EventArgs {
    #region

    public ConnectingEventArgs(MsGraphConnection connection, string code, string url) {
      Connection = connection ?? throw new ArgumentNullException(nameof(connection));
      Code = code?.Trim() ?? throw new ArgumentNullException(nameof(code));
      Url = url?.Trim() ?? throw new ArgumentNullException(nameof(url));
    }

    #endregion

    #region Public

    /// <summary>
    /// Connection
    /// </summary>
    public MsGraphConnection Connection { get; }

    /// <summary>
    /// Code
    /// </summary>
    public string Code { get; }

    /// <summary>
    /// Url
    /// </summary>
    public string Url { get; }

    /// <summary>
    /// Show Authentication
    /// </summary>
    public bool ShowAuthentication() {
      if (Connection is not null)
        return false;

      if (string.IsNullOrWhiteSpace(Url))
        return false;

      using (System.Diagnostics.Process.Start(new ProcessStartInfo {
        FileName = Url,
        UseShellExecute = true
      })) { }

      return true;
    }

    #endregion Public
  }

  //-------------------------------------------------------------------------------------------------------------------
  //
  /// <summary>
  /// Microsoft Graph Simple Client
  /// </summary>
  //
  //-------------------------------------------------------------------------------------------------------------------

  public sealed class MsGraphConnection : IAuthenticationProvider, IEquatable<MsGraphConnection> {
    #region Private Data

    private object m_Sync = new object();

    private static readonly CookieContainer s_CookieContainer;

    private static readonly HttpClient s_HttpClient;

    private string m_ConnectionString = "";

    private List<string> m_Permissions = new();

    private AuthenticationResult m_AuthenticationResult;

    #endregion Private Data

    #region Algorithm

    private static (string code, string uri) ParseMessage(string message) {
      return (
        Regex.Match(message, @"code\s+(?<code>[A-Za-z0-9]+)").Groups["code"].Value,
        Regex.Match(message, @"https://[A-Za-z0-9/.:]+").Value
      );
    }

    private async Task<AuthenticationResult> GetAuthenticiationAsync(CancellationToken token = default) {
      token.ThrowIfCancellationRequested();

      AuthenticationResult result = null;

      if (ConfidentialApplication is not null && !Delegated) {
        try {
          result = await ConfidentialApplication
            .AcquireTokenForClient(Scope)
            .ExecuteAsync()
            .ConfigureAwait(false);

          UserAccount = result.Account;

          IsDelegated = false;

          return result;
        }
        catch (MsalUiRequiredException) {
          ;
        }
        catch (MsalServiceException) {
          ;
        }
        catch (Exception) {
          ;
        }
      }

      if (UserAccount is not null) {
        result = await Application
          .AcquireTokenSilent(Permissions, UserAccount)
          .ExecuteAsync(token)
          .ConfigureAwait(false);

        IsDelegated = true;

        return result;
      }

      if (!string.IsNullOrWhiteSpace(Login) &&
          !string.IsNullOrWhiteSpace(Password) &&
           Guid.TryParse(TenantId, out var _guid)) {
        try {
          System.Security.SecureString pwd = new();

          foreach (char c in Password)
            pwd.AppendChar(c);

          result = await Application
            .AcquireTokenByUsernamePassword(Permissions, Login, pwd)
            .ExecuteAsync(token)
            .ConfigureAwait(false);

          UserAccount = result.Account;

          IsDelegated = true;

          return result;
        }
        catch (TaskCanceledException) {; }
        catch (TimeoutException) {; }
        catch (MsalClientException) {; }
        catch (MsalUiRequiredException) {; }
      }

      var action = Connecting;

      if (action is not null) {
        try {
          result = await Application
            .AcquireTokenWithDeviceCode(
               Permissions,
               callback => {
                 var (code, uri) = ParseMessage(callback.Message);

                 ConnectingEventArgs args = new(this, code, uri);

                 action(this, args);

                 return Task.FromResult(0);
               })
            .ExecuteAsync(token)
            .ConfigureAwait(false);

          UserAccount = result.Account;
          IsDelegated = true;

          return result;
        }
        catch (TaskCanceledException) {; }
        catch (TimeoutException) {; }
        catch (MsalClientException) {; }
        catch (MsalUiRequiredException) {; }
      }

      try {
        SystemWebViewOptions options = new() {
        };

        result = await Application
          .AcquireTokenInteractive(Permissions)
          .WithAccount(null)
          .WithPrompt(Microsoft.Identity.Client.Prompt.SelectAccount)
          .WithSystemWebViewOptions(options)
          .ExecuteAsync(token)
          .ConfigureAwait(false);

        UserAccount = result.Account;
        IsDelegated = true;

        return result;
      }
      catch (TaskCanceledException) {; }
      catch (TimeoutException) {; }
      catch (MsalClientException) {; }
      catch (MsalUiRequiredException) {; }

      IsDelegated = false;

      return null;
    }

    // Access Token
    private async Task<string> GetAccessToken(CancellationToken token = default) {
      token.ThrowIfCancellationRequested();

      AuthenticationResult auth = null;

      Interlocked.Exchange(ref auth, m_AuthenticationResult);

      if (auth is not null) {
        if ((auth.ExpiresOn - DateTimeOffset.Now).TotalSeconds > Expired)
          return auth.AccessToken;

        Interlocked.Exchange(ref m_AuthenticationResult, null);
      }

      auth = await GetAuthenticiationAsync(token);

      if (auth is not null)
        Interlocked.Exchange(ref m_AuthenticationResult, auth);
     
      return auth?.AccessToken;
    }

    #endregion Algorithm

    #region Create

    static MsGraphConnection() {
      try {
        ServicePointManager.SecurityProtocol =
          SecurityProtocolType.Tls |
          SecurityProtocolType.Tls11 |
          SecurityProtocolType.Tls12;
      }
      catch (NotSupportedException) {
        ;
      }

      s_CookieContainer = new CookieContainer();

      var handler = new HttpClientHandler() {
        CookieContainer = s_CookieContainer,
        Credentials = CredentialCache.DefaultCredentials,
      };

      s_HttpClient = new HttpClient(handler) {
        Timeout = Timeout.InfiniteTimeSpan,
      };
    }

    /// <summary>
    /// Standard Constructor
    /// </summary>
    /// <param name="connectionString"></param>
    public MsGraphConnection(string connectionString) {
      if (connectionString is null)
        throw new ArgumentNullException(nameof(connectionString));

      ConnectionString = connectionString;
    }

    #endregion Create

    #region Public

    /// <summary>
    /// Show MS Graph Explorer
    /// </summary>
    public static void ShowGraphExplorer() {
      using (System.Diagnostics.Process.Start(new ProcessStartInfo {
        FileName = @"https://developer.microsoft.com/en-us/graph/graph-explorer",
        UseShellExecute = true
      })) { }
    }

    /// <summary>
    /// Show Azure Portal
    /// </summary>
    public static void ShowAzurePortal() {
      using (System.Diagnostics.Process.Start(new ProcessStartInfo {
        FileName = @"https://azure.microsoft.com/en-us/features/azure-portal",
        UseShellExecute = true
      })) { }
    }

    /// <summary>
    /// Http Client
    /// </summary>
    public static HttpClient Client => s_HttpClient;

    /// <summary>
    /// Connection String
    /// </summary>
    public string ConnectionString {
      get {
        return m_ConnectionString;
      }
      set {
        value = value ?? throw new ArgumentNullException(nameof(value));

        if (string.Equals(value, m_ConnectionString))
          return;

        m_ConnectionString = value;

        DbConnectionStringBuilder builder = new() {
          ConnectionString = value
        };

        TenantId = builder.TryGetValue("Tenant", out var v) ? v.ToString().Trim() : "common";
        ApplicationId = builder.TryGetValue("Application", out v) ? v.ToString().Trim() : "";
        RedirectionAddress = builder.TryGetValue("Redirect", out v) ? v.ToString().Trim() : "http://localhost";
        Login = builder.TryGetValue("Login", out v) ? v.ToString().Trim() : "";
        Password = builder.TryGetValue("Password", out v) ? v.ToString().Trim() : "";
        ClientSecret = builder.TryGetValue("ClientSecret", out v) ? v.ToString().Trim() : "";

        Delegated = builder.TryGetValue("Delegated", out v) && v is bool b ? b : false;

        Expired = builder.TryGetValue("Expired", out v) && v is int iv && iv > 0 ? iv : 30;

        string permissions = builder.TryGetValue("Permissions", out v) ? v.ToString() : "";

        m_Permissions = permissions
          .Split(',', ';', '\t', ' ')
          .Select(line => line.Trim())
          .Where(line => !string.IsNullOrWhiteSpace(line))
          .Distinct(StringComparer.OrdinalIgnoreCase)
          .OrderBy(line => line, StringComparer.OrdinalIgnoreCase)
          .ToList();

        Application = PublicClientApplicationBuilder
          .Create(ApplicationId)
          .WithAuthority(AzureCloudInstance.AzurePublic, TenantId)
          .WithRedirectUri(RedirectionAddress)
          .Build();

        // https://login.microsoftonline.com/eb1ed152-0000-0000-0000-32401f3f9abd

        if (!string.IsNullOrWhiteSpace(ClientSecret)) {
          ConfidentialApplication = ConfidentialClientApplicationBuilder
            .Create(ApplicationId)
            .WithClientSecret(ClientSecret)
            .WithAuthority(new Uri($"https://login.microsoftonline.com/{TenantId}"))
            .WithRedirectUri(RedirectionAddress)
            .Build();
        }

        UserAccount = null;
      }
    }

    /// <summary>
    /// Tenant
    /// </summary>
    public string TenantId { get; private set; } = "common";

    /// <summary>
    /// Application Id
    /// </summary>
    public string ApplicationId { get; private set; } = "";

    /// <summary>
    /// Redirection Address
    /// </summary>
    public string RedirectionAddress { get; private set; } = "";

    /// <summary>
    /// Login
    /// </summary>
    public string Login { get; private set; } = "";

    /// <summary>
    /// Password
    /// </summary>
    public string Password { get; private set; } = "";

    /// <summary>
    /// Client Secret
    /// </summary>
    public string ClientSecret { get; private set; } = "";

    /// <summary>
    /// Scope
    /// </summary>
    public static readonly IReadOnlyList<string> Scope = new List<string>() {
      "https://graph.microsoft.com/.default" };

    /// <summary>
    /// Grant Type
    /// </summary>
    public const string GrantType = "client_credentials";

    /// <summary>
    /// Is Delegated
    /// </summary>
    public bool IsDelegated { get; private set; }

    /// <summary>
    /// Permissions
    /// </summary>
    public IReadOnlyList<string> Permissions => m_Permissions;

    /// <summary>
    /// Connected
    /// </summary>
    public bool Connected => UserAccount is not null;

    /// <summary>
    /// Is Delegated
    /// </summary>
    public bool Delegated { get; private set; }

    /// <summary>
    /// Time to expire
    /// </summary>
    public int Expired { get; private set; } = 30;

    /// <summary>
    /// Connect Async
    /// 
    /// Connection String, parts:
    /// 
    /// Tenant       - optional
    /// ClientSecret - optional
    /// Application
    /// Redirect     - optional
    /// Login        - optional
    /// Password     - optional
    /// Delegated    - optional
    /// Expired      - optional
    /// </summary>
    public async Task<bool> ConnectAsync(CancellationToken token = default) {
      string bearer = await GetAccessToken(token).ConfigureAwait(false);

      return !string.IsNullOrEmpty(bearer);
    }

    /// <summary>
    /// Access Token
    /// </summary>
    public Task<string> AccessToken {
      get => GetAccessToken();
    }

    /// <summary>
    /// Create MS Graph Client
    /// </summary>
    public async Task<GraphServiceClient> CreateGraphClientAsync(CancellationToken token = default) {
      string bearer = await GetAccessToken(token).ConfigureAwait(false);

      if (string.IsNullOrWhiteSpace(bearer))
        throw new DataException("Not connected");

      return new GraphServiceClient(this);
    }

    /// <summary>
    /// Create Command
    /// </summary>
    public MsGraphCommand CreateCommand() => new(this);

    /// <summary>
    /// Create Command
    /// </summary>
    public MsGraphCommand CreateCommand(string version) => new(this, version);

    /// <summary>
    /// Connecting Event
    /// </summary>
    public event EventHandler<ConnectingEventArgs> Connecting;

    /// <summary>
    /// To String
    /// </summary>
    public override string ToString() => $"{(Connected ? "Connected" : "Not connected")}: {Login} @ {ApplicationId} ({TenantId})";

    #endregion Public

    #region IAuthenticationProvider

    /// <summary>
    /// Application (MSA Client)
    /// </summary>
    public IPublicClientApplication Application { get; private set; }

    /// <summary>
    /// Confidential Application (MSA Client)
    /// </summary>
    public IConfidentialClientApplication ConfidentialApplication { get; private set; }

    /// <summary>
    /// User Account
    /// </summary>
    public IAccount UserAccount { get; private set; }

    /// <summary>
    /// 
    /// </summary>
    public async Task AuthenticateRequestAsync(HttpRequestMessage request) {
      if (request is null)
        return;

      request.Headers.Authorization =
        new AuthenticationHeaderValue("bearer", await GetAccessToken().ConfigureAwait(false));
    }

    #endregion : IAuthenticationProvider

    #region IEquatable<MsGraphConnection>

    /// <summary>
    /// Equals
    /// </summary>
    public bool Equals(MsGraphConnection other) {
      if (other is null)
        return false;

      return string.Equals(TenantId, other.TenantId, StringComparison.OrdinalIgnoreCase) &&
             string.Equals(ApplicationId, other.ApplicationId, StringComparison.OrdinalIgnoreCase) &&
             string.Equals(RedirectionAddress, other.RedirectionAddress, StringComparison.OrdinalIgnoreCase) &&
             string.Equals(Login, other.Login, StringComparison.Ordinal) &&
             string.Equals(Password, other.Password, StringComparison.Ordinal) &&
             Permissions.Count == other.Permissions.Count &&
             Permissions.Zip(other.Permissions, (a, b) => string.Equals(a, b, StringComparison.Ordinal)).All(a => a);
    }

    /// <summary>
    /// Equals
    /// </summary>
    public override bool Equals(object obj) => obj is MsGraphConnection other && Equals(other);

    /// <summary>
    /// HashCode
    /// </summary>
    public override int GetHashCode() =>
      TenantId.GetHashCode(StringComparison.OrdinalIgnoreCase) ^
      ApplicationId.GetHashCode(StringComparison.OrdinalIgnoreCase) ^
      Login.GetHashCode(StringComparison.OrdinalIgnoreCase) ^
      Password.GetHashCode(StringComparison.OrdinalIgnoreCase) ^
      Permissions.Count;

    #endregion IEquatable<MsGraphConnection>
  }

}
