<!-- default badges list -->
![](https://img.shields.io/endpoint?url=https://codecentral.devexpress.com/api/v1/VersionRange/491023724/21.2.7%2B)
[![](https://img.shields.io/badge/Open_in_DevExpress_Support_Center-FF7200?style=flat-square&logo=DevExpress&logoColor=white)](https://supportcenter.devexpress.com/ticket/details/T1088492)
[![](https://img.shields.io/badge/📖_How_to_use_DevExpress_Examples-e9f6fc?style=flat-square)](https://docs.devexpress.com/GeneralInformation/403183)
<!-- default badges end -->
# Snap – Convert Your SNX Reports to REPX Files

As you may already know, the [WinForms Snap control](https://docs.devexpress.com/WindowsForms/11373/controls-and-libraries/snap) and [Snap Report API](https://docs.devexpress.com/OfficeFileAPI/15188/snap-report-api) are now in maintenance support mode. No new features or capabilities are incorporated into these products. We recommend that you use [DevExpress Reporting](https://docs.devexpress.com/XtraReports/2162/reporting) tool to generate, edit, print, and export your business reports/documents.

To help you migrate to DevExpress Reports, we created an application that allows you to convert your SNX report templates to REPX files.

## How to Convert Snap Reports to REPX Files

Download and run this application. Select an SNX file you need to convert to REPX and click **Convert**. The resulting report is opened in the [WinForms End-User Report Designer](https://docs.devexpress.com/XtraReports/8546/winforms-reporting/end-user-report-designer-for-winforms/gui/end-user-report-designer-with-a-ribbon-toolbar).

![Snap - SNX to REPX Converter](./images/snap-report-converter.png)

Handle the [ConfigureDataConnection](https://github.com/DevExpress-Examples/snap-convert-snx-to-repx/blob/21.2.7%2B/CS/SnxToRepx/MainForm.cs#L62) event if you need to update connection settings for your data source. Refer to the following help topic for details: [SqlDataSource.ConfigureDataConnection](https://docs.devexpress.com/CoreLibraries/DevExpress.DataAccess.Sql.SqlDataSource.ConfigureDataConnection).

## Limitations

Since Snap and DevExpress Reports use completely different formats and have incompatible feature sets, the conversion process has the following limitations:

* A Snap document can contain auxiliary controls (such as [checkboxes](https://docs.devexpress.com/WindowsForms/14803/controls-and-libraries/snap/graphical-user-interface/data-visualization-tools/check-box)) within formatted text. Since a DevExpress Report does not support this functionality, the converter uses textual representation for Snap content controls. This means that the resulting report contains the "True" and "False" strings instead of checkboxes.
* A Snap document can contain multiple sections with different page settings. A DevExpress Report cannot use different page settings in the same report, so only the first section's settings are applied during conversion.
* In Snap, different data sources can contain calculated fields that have the same name. To avoid potential conflicts in DevExpress Reports, we disabled the ability to add fields with the same name.
* Snap allows you to create documents with a multi-column layout. A DevExpress Report does not support this layout, so the converter creates a single-column report.
* [XRTableCell](https://docs.devexpress.com/XtraReports/DevExpress.XtraReports.UI.XRTableCell) and [XRLabel](https://docs.devexpress.com/XtraReports/DevExpress.XtraReports.UI.XRLabel) use expressions and HTML-style markup to build and format content. If these objects have nested controls, markup and expressions are ignored and only nested controls are visible. Snap allows you to place any content within a table cell (including multiple nested table levels). There is no generic solution to automatically convert this content to DevExpress Reports.
* Detail [Snap lists](https://docs.devexpress.com/WindowsForms/DevExpress.Snap.Core.API.SnapList) do not have the `DataSource` property specified. Instead, they use the parent's data source. DevExpress Reports require the `DataSource` property to be set. Moreover, the `DataMember` property should contain a full path to the data member (for example, DevExpress Reports use "Products.ProductOrders", while Snap uses "ProductOrders" in this case). You need to correct the `DataSource` and `DataMember` properties for each [DetailReportBand](https://docs.devexpress.com/XtraReports/DevExpress.XtraReports.UI.DetailReportBand) object after conversion.
* DevExpress Reports do not differentiate between first/primary/odd/even headers and footers. So, the converter processes only the primary header and footer of a Snap report.
* Snap allows you to place page fields at any position within formatted text. A DevExpress Report does not support this functionality, so all fields in headers and footers are converted to static values.
* Static text in a Snap report is converted to markup. Markup with different format settings starts with a new line.

If you encounter an issue while using our converter app, please [submit a ticket to our DevExpress Support Center](https://supportcenter.devexpress.com/ticket/create).
