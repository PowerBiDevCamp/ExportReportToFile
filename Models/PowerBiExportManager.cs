using System;
using System.Collections.Generic;
using Microsoft.PowerBI.Api;
using Microsoft.PowerBI.Api.Models;
using System.IO;
using System.Configuration;

namespace ExportReportToFile.Models {

  class PowerBiExportManager {

    private static PowerBIClient pbiClient = TokenManager.GetPowerBiAppOnlyClient();
 
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

    public static void GetReportPages(Guid WorkspaceId, string ReportName) {
      var reports = pbiClient.Reports.GetReportsInGroup(WorkspaceId).Value;
      foreach (var report in reports) {
        if (report.Name.Equals(ReportName)) {
          var reportPages = pbiClient.Reports.GetPagesInGroup(WorkspaceId, report.Id).Value;
          foreach (var page in reportPages) {
            Console.WriteLine(page.DisplayName);
            Console.WriteLine(page.Name);
          }
        }
      }
    }

    public static void ExportReport(Guid WorkspaceId, Guid ReportId, string ExportName, FileFormat ExportFileFormat) {

      // create export report request
      var exportRequest = new ExportReportRequest {
        Format = ExportFileFormat
      };

      // call ExportToFileInGroup to start async export job
      Export export = pbiClient.Reports.ExportToFileInGroup(WorkspaceId, ReportId, exportRequest);

      string exportId = export.Id;

      do { 
        // poll every 10 seconds to see if export job has completed
        System.Threading.Thread.Sleep(10000);
        export = pbiClient.Reports.GetExportToFileStatusInGroup(WorkspaceId, ReportId, exportId);

        // continue to poll until export job status is equal to Suceeded or Failed
      } while (export.Status != ExportState.Succeeded && export.Status != ExportState.Failed);

      if (export.Status == ExportState.Failed) {
        // deal with failure
        Console.WriteLine("Export failed!");
      }

      if (export.Status == ExportState.Succeeded) {
        // retreive file name extension from export object to construct file name
        string FileName = ExportName + export.ResourceFileExtension.ToLower();
        string FilePath = ExportFolderPath + FileName;
        
        // get export stream with file
        var exportStream = pbiClient.Reports.GetFileOfExportToFileInGroup(WorkspaceId, ReportId, exportId);

        // save exported file stream to local file
        FileStream fileStream = File.Create(FilePath);
        exportStream.CopyTo(fileStream);
        fileStream.Close();
      }
    }

    public static void ExportPowerBIReport(Guid WorkspaceId, Guid ReportId, string ExportName, FileFormat ExportFileFormat, string ExportFilter = "") {

      Console.WriteLine("Exporting " + ExportName + " (Power BI report) to " + ExportFileFormat.ToString());

      var exportRequest = new ExportReportRequest {
        Format = ExportFileFormat,
        PowerBIReportConfiguration = new PowerBIReportExportConfiguration {
          Settings = new ExportReportSettings {
            IncludeHiddenPages = false
          }
        }
      };

      if (ExportFilter != "") {
        exportRequest.PowerBIReportConfiguration.ReportLevelFilters =
          new List<ExportFilter>() {
            new ExportFilter { Filter = ExportFilter }
          };
      }
          
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
        string FileName = ExportName + export.ResourceFileExtension.ToLower();

        string FilePath = ExportFolderPath + FileName;

        Console.WriteLine(" - Saving exported file to " + FilePath);
        var exportStream = pbiClient.Reports.GetFileOfExportToFileInGroup(WorkspaceId, ReportId, exportId);
        FileStream fileStream = File.Create(FilePath);
        exportStream.CopyTo(fileStream);
        fileStream.Close();
      }

      Console.WriteLine();

    }

    public static void ExportPaginatedReport(Guid WorkspaceId, Guid ReportId, string ExportName, FileFormat ExportFileFormat, string OutputFormat = "pdf", List<ParameterValue> Parameters = null ) {

      Console.WriteLine("Exporting " + ExportName + " (paginated report) to " + ((ExportFileFormat.Equals(FileFormat.IMAGE)) ? OutputFormat.ToUpper() : ExportFileFormat.ToString()));

      var exportRequest = new ExportReportRequest {
        Format = ExportFileFormat,
        PaginatedReportConfiguration = new PaginatedReportExportConfiguration {
          ParameterValues = new List<ParameterValue>() {},
          FormatSettings = new Dictionary<string, string>() {
            { "OutputFormat", OutputFormat }
          }
        }
      };

      if(Parameters != null) {
        exportRequest.PaginatedReportConfiguration.ParameterValues = Parameters;
      }

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
        string FileName = ExportName + export.ResourceFileExtension.ToLower();

        string FilePath = ExportFolderPath + FileName;

        Console.WriteLine(" - Saving exported file to " + FilePath);
        var exportStream = pbiClient.Reports.GetFileOfExportToFileInGroup(WorkspaceId, ReportId, exportId);
        FileStream fileStream = File.Create(FilePath);
        exportStream.CopyTo(fileStream);
        fileStream.Close();
      }

      Console.WriteLine();

    }

  }

}
