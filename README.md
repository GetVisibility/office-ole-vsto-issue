# OLE / VSTO Issue Reproduction Project

This repository contains a minimal sample project demonstrating an issue
with the interaction between **Microsoft OLE** and **VSTO add-ins** in
**Microsoft Excel** and **Microsoft Word**.

The project is intended to reproduce a failure that occurs when
inserting an Excel Worksheet Object into Word and attempting to activate
it. This repository can be used for investigation by Microsoft support
engineers or developers facing similar behavior.

------------------------------------------------------------------------

## Project Structure

    /src
       ├── OLEIssue.sln
       ├── OLEIssue.Common
       ├── OLEIssue.ExcelAddin
       └── OLEIssue.WordAddin

The project is built using **Visual Studio 2022** and **VSTO Add-ins**.

------------------------------------------------------------------------

## Requirements

-   Visual Studio 2022
-   Installed Workloads:
    -   .NET desktop development
    -   Office/SharePoint development
-   Microsoft Office (Excel and Word) with VSTO add-in support

------------------------------------------------------------------------

## Steps to Reproduce the Issue

Follow the steps below to reliably reproduce the problem:

1.  Open the `OLEIssue.sln` solution in **Visual Studio 2022**.
2.  Build all projects (Build → Rebuild Solution).
3.  Start debugging the **OLEIssue.ExcelAddin** project.
4.  Start debugging the **OLEIssue.WordAddin** project.
5.  In Excel:
    -   Create a new workbook.
    -   Enter some sample values into a few cells.
6.  Select the cell range and copy it to the clipboard (Ctrl+C).
7.  Switch to Word:
    -   Open the **Paste Special...** dialog.
    -   Select **Microsoft Excel Worksheet Object** and insert it.
8.  After the object is inserted, double-click it to activate editing.
9.  Word displays an error --- this is the intended reproduction of the
    issue.

------------------------------------------------------------------------

## Expected Result

After inserting the Excel Worksheet Object and attempting to activate
it, Microsoft Word shows an error. The issue occurs only when VSTO
add-ins for both Word and Excel are running, demonstrating a conflict
between OLE embedding and VSTO add-in activation.

