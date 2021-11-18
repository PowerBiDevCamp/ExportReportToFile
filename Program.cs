using ExportReportToFile.Models;
using Microsoft.PowerBI.Api.Models;
using System;
using System.Collections.Generic;

namespace ExportReportToFile {

  class Program {

    static void Main() {

      string workspaceName = "Export Demo";
      string powerBiReportName = "Customer Sales";
      string paginatedReportName = "Customer Sales By City";

      // get workspace info
      var workspace = PowerBiExportManager.GetWorkspace(workspaceName);

      // Export Power BI Report
      var powerBiReport = PowerBiExportManager.GetReport(workspace.Id, powerBiReportName);
      PowerBiExportManager.ExportPowerBIReport(workspace.Id, powerBiReport.Id, powerBiReportName, FileFormat.PDF);
      PowerBiExportManager.ExportPowerBIReport(workspace.Id, powerBiReport.Id, powerBiReportName, FileFormat.PPTX);
      PowerBiExportManager.ExportPowerBIReport(workspace.Id, powerBiReport.Id, powerBiReportName, FileFormat.PNG);


      string ReportName1 = "Customer Sales 2020 - Southwest States";
      string ReportFilter1 = "Customers/State in ('FL', 'GA', 'LA', 'TX') and Calendar/Year in (2020)";
      PowerBiExportManager.ExportPowerBIReport(workspace.Id, powerBiReport.Id, ReportName1, FileFormat.PDF, ReportFilter1);

      string ReportName2 = "Customer Sales 2020 - Western Region";
      string ReportFilter2 = "Customers/SalesRegion in ('Western Region') and Calendar/Year in (2020)";
      PowerBiExportManager.ExportPowerBIReport(workspace.Id, powerBiReport.Id, ReportName2, FileFormat.PDF, ReportFilter2);

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

      string PaginatedReportName1 = paginatedReportName + " for West Coast";
      var PaginatedReportParameters1 = new List<ParameterValue>() {
        new ParameterValue { Name="State", Value="CA" },
        new ParameterValue { Name="State", Value="OR" },
        new ParameterValue { Name="State", Value="WA" }
      };
      PowerBiExportManager.ExportPaginatedReport(workspace.Id, paginatedReport.Id, PaginatedReportName1, FileFormat.PDF, Parameters: PaginatedReportParameters1);

      string PaginatedReportName2 = paginatedReportName + " for New England";
      var PaginatedReportParameters2 = new List<ParameterValue>() {
        new ParameterValue { Name="State", Value="NH" },
        new ParameterValue { Name="State", Value="MA" },
        new ParameterValue { Name="State", Value="CT" },
        new ParameterValue { Name="State", Value="RI" },
        new ParameterValue { Name="State", Value="NY" }
      };
      PowerBiExportManager.ExportPaginatedReport(workspace.Id, paginatedReport.Id, PaginatedReportName2, FileFormat.PDF, Parameters: PaginatedReportParameters2);

    }
  }
}
