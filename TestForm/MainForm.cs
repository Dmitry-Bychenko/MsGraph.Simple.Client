using MsGraph.Simple.Client.Graph.Storage;
using MsGraph.Simple.Client.Json;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TestForm {

  // Correct:
  //   "Tenant=1da41803-dfa6-4727-8d04-dba93b9ea42d;Application=b502c859-aa0b-4895-890f-0d357491e96d;ClientSecret=Y5E7Q~tFzXSm2hnGKJr5tZfrPv5yeesfXNfk9;login=dmitry.bytchenko@7801676234.onmicrosoft.com;password=797Gnome797@;permissions=User.ReadBasic.All,User.Read,User.Read.All,User.ReadWrite.All,Files.Read,Files.ReadWrite,Files.ReadWrite.AppFolder";
  // Target:
  //   "Tenant=1b4a1891-24bd-451e-8548-48986af6f553;Application=a4f2dc33-a706-49fd-b56e-3018ca81f49d;ClientSecret=inQ7Q~lvvQU5Sddlfwqrwk38xMoNEzG716Xyy;login=sync_aad_1c@nedra.digital;password=Har15204;permissions=User.ReadBasic.All,User.Read,User.Read.All,User.ReadWrite.All,Files.Read,Files.ReadWrite,Files.ReadWrite.AppFolder";
  // Test:
  //   "Tenant=1b4a1891-24bd-451e-8548-48986af6f553;Application=a8b596a4-3183-4ba2-850e-f0f8f1b683ba;ClientSecret=yog7Q~jR1mhfRBvgvwLNlgJ2IZ-Gdoii3bA.p;login=sync_aad_1c@nedra.digital;password=Har15204;permissions=User.ReadBasic.All,User.Read,User.Read.All,User.ReadWrite.All,Files.Read,Files.ReadWrite,Files.ReadWrite.AppFolder";


  public partial class MainForm : Form {

    private static long Demo(long a, int n) {
      for (int i = 0; i < n; ++i)
        for (int j = 0; j < i; ++j)
          a = a * (i + j);

      return a;
    }

    public interface IBoolS {
      IEnumerable<bool> GetBools(int N);
    }

    public class Generator : IBoolS {
      public bool ValueToGenerate { get; init; }

      public IEnumerable<bool> GetBools(int N) => Enumerable.Repeat(ValueToGenerate, N);
    }

    class CustomType {
      public int X;
      public int Y;

      public override string ToString() => $"{X,2} : {Y,3}";
    }

    private static IEnumerable<CustomType> Fill(IEnumerable<CustomType> source) {
      int expected = 0; // let compiler be happy, we'll rewrite this value
      bool first = true;

      foreach (var item in source.OrderBy(x => x.X)) {
        if (!first && item.X > expected)
          while (item.X > expected)
            yield return new CustomType() { X = expected++, Y = 0 };

        first = false;
        expected = item.X + 1;

        yield return item;
      }
    }

    private async Task CorePerform() {
      // 2HZ7Q~X_R6-HGmHjA3zHLSE5TZIOj1qR7OAC1

      List<CustomType> demo = new List<CustomType>() {
        new CustomType() { X = 0, Y = 0 },
        new CustomType() { X = 1, Y = 10 },
        new CustomType() { X = 4, Y = 40 },
        new CustomType() { X = 5, Y = 50 },
        new CustomType() { X = 8, Y = 80},
        new CustomType() { X = 10, Y = 100 },
        new CustomType() { X = 10, Y = 101 },
        new CustomType() { X = 12, Y = 120 },
      };

      string report = string.Join(Environment.NewLine, Fill(demo));

      rtbMain.Text = report;


      return;

      string connectionString = "put it here";

      Enterprise users = await Enterprise.CreateAsync(connectionString);

      foreach (var user in users) {
        Stream stream = await user.Client.ReadImageAsync(user.User.Id);
      }

      rtbMain.Text = string.Join(Environment.NewLine, users
        .Users.Select(user => $"{user.User.Id} : {user.User.DisplayName} : {user.User.UserPrincipalName}"));
    }

    public MainForm() {
      InitializeComponent();
    }

    private void MainForm_Load(object sender, EventArgs e) {

    }

    private async void btnRun_Click(object sender, EventArgs e) {
      await CorePerform();
    }
  }
}
