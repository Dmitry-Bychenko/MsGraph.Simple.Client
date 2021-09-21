using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using MsGraph.Simple.Client;

namespace TestForm {
  
  public partial class MainForm : Form {
    private async Task CorePerform() {
      string connectionString = "***";
      
      MsGraphConnection conn = new (connectionString);

      var client = await conn.CreateGraphClient();

      var user = await client
        .Me
        .Request()
        .GetAsync();
        

      rtbMain.Text = $"{user.DisplayName}";
      
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
