using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Microsoft.Graph;

using MsGraph.Simple.Client;
using MsGraph.Simple.Client.Graph;

namespace TestForm {
  
  public partial class MainForm : Form {
    private async Task CorePerform() {
      string connectionString =
        "";

      MsGraphConnection conn = new (connectionString);

      var client = await conn.CreateGraphClient();

      List<string> list = new List<string>();

      var data = client
        .Users
        .Request()
        .Expand("Manager,Extensions")
        .EnumerateAsync<User>();

      await foreach(var item in data) {
        list.Add(item.DisplayName);
      }

      /*
      var data = await client
          .Users
          .Request()
          .Expand("Manager,Extensions")
          
          
          //.Select(select)
          .GetAsync()
          .ConfigureAwait(false);

      //IGraphServiceAgreementsCollectionRequest 

     

      await foreach(var item in data.EnumerateAsync<User>().ConfigureAwait(false)) {
        list.Add(item.DisplayName);
      }
      */

      rtbMain.Text = string.Join(Environment.NewLine, list);

      /*
      var user = await client
        .Me
        .Request()
        .GetAsync();
        

      rtbMain.Text = $"{user.DisplayName}";
      */
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
