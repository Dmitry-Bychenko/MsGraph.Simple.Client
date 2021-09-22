# MsGraph.Simple.Client

Simple MS Graph Client

```c#
using MsGraph.Simple.Client;
using MsGraph.Simple.Client.Graph;
using MsGraph.Simple.Client.Graph.Storage;

...

Enterprise users = await Enterprise.CreateAsync(connectionString);

Console.Write(string.Join(Environment.NewLine, users
  .Users
  .Select(u => $"{u.User.DisplayName}"));
```
