﻿using MsGraph.Simple.Client;
using MsGraph.Simple.Client.Graph;
using MsGraph.Simple.Client.Graph.Storage;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TestForm {

  public partial class MainForm : Form {
    private async Task CorePerform() {
      string connectionString =
        "Connection String Here";
        ;

      Enterprise users = await Enterprise.CreateAsync(connectionString);

      rtbMain.Text = string.Join(Environment.NewLine, users
        .Users
        .Select(u => $"{u.User.DisplayName}"));

      /*
      MsGraphConnection conn = new(connectionString);

      var client = await conn.CreateGraphClientAsync();

      List<string> list = new();

      //var result = await OneNoteFile.DeleteFileAsync(client, @"abc\def\pqr.txt");

      //string text = await OneNoteFile.ReadAllText(client, @"abc/def/pqr.txt");

      var me = await client
        .Me
        .Request()
        .GetAsync();

      bool result = await UserSchema.DropExtensionAsync(
        client,
        me.Id,
        "HR.Russian.Names"
        );
      */

      /*
      var data = OneNoteDirectory.EnumerateFilesAsync(client, "", x => true, SearchOption.AllDirectories);

      await foreach (var item in data) {
        list.Add(item);
      }
      */

      //rtbMain.Text = result ? "OK" : "Failed";

      /*
      

      List<string> list = new List<string>();

      var data = client
        .Users
        .Request()
        .Expand("Manager,Extensions")
        .EnumerateAsync<User>();

      await foreach(var item in data) {
        list.Add(item.DisplayName);
      }
      */

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

      //rtbMain.Text = string.Join(Environment.NewLine, list);

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