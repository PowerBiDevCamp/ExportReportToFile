using ExportReportToFile.Models;
using Microsoft.PowerBI.Api.Models;
using System;

namespace ExportReportToFile {

  class Program {

    static void Main() {

      string workspaceName = "Export Demo";
      string powerBiReportName = "Customer Sales";
      string paginatedReportName = "Customer Sales Paginated";

      // get workspace info
      var workspace = PowerBiExportManager.GetWorkspace(workspaceName);

      // Export Power BI Report
      var powerBiReport = PowerBiExportManager.GetReport(workspace.Id, powerBiReportName);
      PowerBiExportManager.ExportPowerBIReport(workspace.Id, powerBiReport.Id, powerBiReportName, FileFormat.PDF);
      PowerBiExportManager.ExportPowerBIReport(workspace.Id, powerBiReport.Id, powerBiReportName, FileFormat.PPTX);
      PowerBiExportManager.ExportPowerBIReport(workspace.Id, powerBiReport.Id, powerBiReportName, FileFormat.PNG);

      // Export Pagnated Report
      var paginatedReport = PowerBiExportManager.GetReport(workspace.Id, paginatedReportName);
      PowerBiExportManager.ExportPaginatedReport(workspace.Id, paginatedReport.Id, paginatedReportName, FileFormat.PDF);
      PowerBiExportManager.ExportPaginatedReport(workspace.Id, paginatedReport.Id, paginatedReportName, FileFormat.PPTX);
      PowerBiExportManager.ExportPaginatedReport(workspace.Id, paginatedReport.Id, paginatedReportName, FileFormat.XLSX);
      PowerBiExportManager.ExportPaginatedReport(workspace.Id, paginatedReport.Id, paginatedReportName, FileFormat.CSV);
      PowerBiExportManager.ExportPaginatedReport(workspace.Id, paginatedReport.Id, paginatedReportName, FileFormat.XML);
      PowerBiExportManager.ExportPaginatedReport(workspace.Id, paginatedReport.Id, paginatedReportName, FileFormat.MHTML);
      PowerBiExportManager.ExportPaginatedReport(workspace.Id, paginatedReport.Id, paginatedReportName, FileFormat.IMAGE, "tif");
      PowerBiExportManager.ExportPaginatedReport(workspace.Id, paginatedReport.Id, paginatedReportName, FileFormat.IMAGE, "bmp");
      PowerBiExportManager.ExportPaginatedReport(workspace.Id, paginatedReport.Id, paginatedReportName, FileFormat.IMAGE, "emf");
      PowerBiExportManager.ExportPaginatedReport(workspace.Id, paginatedReport.Id, paginatedReportName, FileFormat.IMAGE, "gif");
      PowerBiExportManager.ExportPaginatedReport(workspace.Id, paginatedReport.Id, paginatedReportName, FileFormat.IMAGE, "jpg");
      PowerBiExportManager.ExportPaginatedReport(workspace.Id, paginatedReport.Id, paginatedReportName, FileFormat.IMAGE, "png");

    }
  }
}
