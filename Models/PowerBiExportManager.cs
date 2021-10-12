using System;
using System.Collections.Generic;
using Microsoft.PowerBI.Api;
using Microsoft.PowerBI.Api.Models;
using System.IO;
using System.Configuration;

namespace ExportReportToFile.Models {

  class PowerBiExportManager {

    // use this to export with identity of service principal
    private static PowerBIClient pbiClient = TokenManager.GetPowerBiAppOnlyClient();

    // use this to export with identity of a user
    // private readonly static string[] requiredScopes = PowerBiPermissionScopes.ManageWorkspaceAssets;
    // private static PowerBIClient pbiClient = TokenManager.GetPowerBiClient(requiredScopes);

    private static string ExportFolderPath = ConfigurationManager.AppSettings["export-folder-path"];

    public static void GetWorkspaces() {
      var workspaces = pbiClient.Groups.GetGroups().Value;
      foreach (var workspace in workspaces) {
        Console.WriteLine(workspace.Name);
      }
    }

    public static Group GetWorkspace(string WorkspaceName) {
      var workspaces = pbiClient.Groups.GetGroups().Value;
      foreach (var workspace in workspaces) {
        if (workspace.Name.Equals(WorkspaceName))
          return workspace;
      }
      return null;
    }

    public static Dataset GetDataset(Guid WorkspaceId, string DatasetName) {
      var datasets = pbiClient.Datasets.GetDatasetsInGroup(WorkspaceId).Value;
      foreach (var dataset in datasets) {
        if (dataset.Name.Equals(DatasetName)) {
          return dataset;
        }
      }
      return null;
    }

    public static Report GetReport(Guid WorkspaceId, string ReportName) {
      var reports = pbiClient.Reports.GetReportsInGroup(WorkspaceId).Value;
      foreach (var report in reports) {
        if (report.Name.Equals(ReportName)) {
          return report;
        }
      }
      return null;
    }

    public static void ExportPowerBIReport(Guid WorkspaceId, Guid ReportId, string ExportName, FileFormat ExportFileFormat) {

      Console.WriteLine("Exporting " + ExportName + " (Power BI report) to " + ExportFileFormat.ToString());
 
      var exportRequest = new ExportReportRequest {
        Format = ExportFileFormat
      };

      Export export = pbiClient.Reports.ExportToFileInGroup(WorkspaceId, ReportId, exportRequest);

      string exportId = export.Id;

      do {
        System.Threading.Thread.Sleep(10000);
        export = pbiClient.Reports.GetExportToFileStatusInGroup(WorkspaceId, ReportId, exportId);
        Console.WriteLine(" - Export status: " + export.PercentComplete.ToString() + "% complete");
      } while (export.Status != ExportState.Succeeded && export.Status != ExportState.Failed);

      if (export.Status == ExportState.Failed) {
        Console.WriteLine("Export failed!");
      }


      if (export.Status == ExportState.Succeeded) {
        string FileName = ExportName + export.ResourceFileExtension;

        string FilePath = ExportFolderPath + FileName;

        Console.WriteLine(" - Saving exported file to " + FilePath);
        var exportStream = pbiClient.Reports.GetFileOfExportToFileInGroup(WorkspaceId, ReportId, exportId);
        FileStream fileStream = File.Create(FilePath);
        exportStream.CopyTo(fileStream);
        fileStream.Close();
      }

      Console.WriteLine();

    }

    public static void ExportPaginatedReport(Guid WorkspaceId, Guid ReportId, string ExportName, FileFormat ExportFileFormat, string OutputFormat = "tif") {

      Console.WriteLine("Exporting " + ExportName + " (paginated report) to " + ((ExportFileFormat.Equals(FileFormat.IMAGE)) ? OutputFormat.ToUpper() : ExportFileFormat.ToString()) );

      var exportRequest = new ExportReportRequest {
        Format = ExportFileFormat,
        PaginatedReportConfiguration = new PaginatedReportExportConfiguration {
          FormatSettings = new Dictionary<string, string>() {
            { "OutputFormat", OutputFormat }
          }
        }
      };

      Export export = pbiClient.Reports.ExportToFileInGroup(WorkspaceId, ReportId, exportRequest);

      string exportId = export.Id;

      do {
        System.Threading.Thread.Sleep(10000);
        export = pbiClient.Reports.GetExportToFileStatusInGroup(WorkspaceId, ReportId, exportId);
        Console.WriteLine(" - Export status: " + export.PercentComplete.ToString() + "% complete");
      } while (export.Status != ExportState.Succeeded && export.Status != ExportState.Failed);

      if (export.Status == ExportState.Failed) {
        Console.WriteLine("Export failed!");
      }


      if (export.Status == ExportState.Succeeded) {
        string FileName = ExportName + export.ResourceFileExtension;

        string FilePath = ExportFolderPath + FileName;

        Console.WriteLine(" - Saving exported file to " + FilePath);
        var exportStream = pbiClient.Reports.GetFileOfExportToFileInGroup(WorkspaceId, ReportId, exportId);
        FileStream fileStream = File.Create(FilePath);
        exportStream.CopyTo(fileStream);
        fileStream.Close();
      }

      Console.WriteLine();

    }

    public static void ExportVisual(Guid WorkspaceId, Guid ReportId, string PageName, string VisualName) {

      var exportRequest = new ExportReportRequest {
        Format = FileFormat.PDF,
        PowerBIReportConfiguration = new PowerBIReportExportConfiguration {
          Pages = new List<ExportReportPage>() {
            new ExportReportPage{
              PageName=PageName,
              VisualName=VisualName
            }
          }
        }
      };

      Export export = pbiClient.Reports.ExportToFileInGroup(WorkspaceId, ReportId, exportRequest);

      string exportId = export.Id;

      do {
        System.Threading.Thread.Sleep(5000);
        export = pbiClient.Reports.GetExportToFileStatus(ReportId, exportId);
        Console.WriteLine("Getting export status - " + export.PercentComplete.ToString() + "% complete");
      } while (export.Status != ExportState.Succeeded && export.Status != ExportState.Failed);

      if (export.Status == ExportState.Failed) {
        Console.WriteLine("Export failed!");
      }

      if (export.Status == ExportState.Succeeded) {
        string fileName = ExportFolderPath + "Visual1" + export.ResourceFileExtension;
        Console.WriteLine("Saving exported file to " + fileName);
        var exportStream = pbiClient.Reports.GetFileOfExportToFile(WorkspaceId, ReportId, exportId);
        FileStream fileStream = File.Create(fileName);
        exportStream.CopyTo(fileStream);
        fileStream.Close();
      }

    }

  }

}
