
namespace ExportReportToFile.Models {

  class PowerBiPermissionScopes {

    public static readonly string[] Default = new string[] {
          "https://analysis.windows.net/powerbi/api/.default"
    };

    public static readonly string[] ReadWorkspaceAssets = new string[] {
        "https://analysis.windows.net/powerbi/api/Dashboard.Read.All",
        "https://analysis.windows.net/powerbi/api/Report.Read.All",
        "https://analysis.windows.net/powerbi/api/Dataset.Read.All",
    };

    public static readonly string[] ReadUserWorkspaces = new string[] {
        "https://analysis.windows.net/powerbi/api/Group.Read.All"
    };

    public static readonly string[] ManageWorkspaceAssets = new string[] {
        "https://analysis.windows.net/powerbi/api/Content.Create",
        "https://analysis.windows.net/powerbi/api/Dashboard.ReadWrite.All",
        "https://analysis.windows.net/powerbi/api/Report.ReadWrite.All",
        "https://analysis.windows.net/powerbi/api/Dataset.ReadWrite.All",
        "https://analysis.windows.net/powerbi/api/Dataflow.ReadWrite.All",
        "https://analysis.windows.net/powerbi/api/Group.Read.All",
        "https://analysis.windows.net/powerbi/api/Workspace.ReadWrite.All"
    };

    public static readonly string[] ReadUserApps = new string[] {
        "https://analysis.windows.net/powerbi/api/App.Read.All"
    };

    public static readonly string[] TenantReadAll = new string[] {
        "https://analysis.windows.net/powerbi/api/Tenant.Read.All" // requires admin
    };

    public static readonly string[] TenantReadWriteAll = new string[] {
        "https://analysis.windows.net/powerbi/api/Tenant.ReadWrite.All" // requires admin
    };

    public static readonly string[] KitchenSink = new string[] {
        "https://analysis.windows.net/powerbi/api/Group.Read.All",
        "https://analysis.windows.net/powerbi/api/Content.Create",
        "https://analysis.windows.net/powerbi/api/Dashboard.ReadWrite.All",
        "https://analysis.windows.net/powerbi/api/Report.ReadWrite.All",
        "https://analysis.windows.net/powerbi/api/Dataset.ReadWrite.All",
        "https://analysis.windows.net/powerbi/api/Dataflow.ReadWrite.All",
        "https://analysis.windows.net/powerbi/api/Workspace.ReadWrite.All",
        "https://analysis.windows.net/powerbi/api/App.Read.All",
        "https://analysis.windows.net/powerbi/api/Capacity.ReadWrite.All",
        "https://analysis.windows.net/powerbi/api/Gateway.ReadWrite.All",
        "https://analysis.windows.net/powerbi/api/StorageAccount.ReadWrite.All",
        "https://analysis.windows.net/powerbi/api/Metadata.View_Any",
        "https://analysis.windows.net/powerbi/api/Data.Alter_Any"
    };

  }


}
