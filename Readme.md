# [ARCHIVED] Sheet Switcher Task Pane Add-in Sample for Excel 2016

**Note:** This repo is archived and no longer actively maintained. Security vulnerabilities may exist in the project, or its dependencies. If you plan to reuse or run any code from this repo, be sure to perform appropriate security checks on the code or dependencies first. Do not use this project as the starting point of a production Office Add-in. Always start your production code by using the Office/SharePoint development workload in Visual Studio, or the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office), and follow security best practices as you develop the add-in. 
_Applies to: Excel 2016_

This task pane add-in provides a way to add new sheets to a workbook via the task pane and easily navigate to the different sheets in Excel 2016. It comes in two flavors: text editor and Visual Studio.

![Sheet Switcher Sample](images/SheetSwitcher_taskpane.PNG)

## Try it out
### Text editor version

The simplest way to deploy and test your add-in is to copy the files to a network share.

1.  Copy the files in the Text Editor folder and host them by using your favorite server.
2.  Edit the \<SourceLocation\> and \<Url\> elements of the manifest file (SheetSwitcherManifest.xml) so that it points to the hosted location from step 1. (i.e. https://localhost/SheetSwitcher/home.html)
3.  Copy the manifest (SheetSwitcherManifest.xml) to a network share (for example, \\\MyShare\SheetSwitcher).
4.  Add the share location that contains the manifest as a trusted app catalog in Excel.

    a.  Launch Excel and open a blank spreadsheet.

    b.  Choose the **File** tab, and then choose **Options**.

    c.  Choose **Trust Center**, and then choose the **Trust Center Settings** button.

    d.  Choose **Trusted App Catalogs**.

    e.  In the **Catalog Url** box, enter the path to the network share you created in step 1, and then choose **Add Catalog**.

   f.  Select the **Show in Menu** check box, and then choose **OK**. A message appears to inform you that your settings will be applied the next time you start Office.

5.  Test and run the add-in.

    a.  In the **Insert tab** in Excel 2016, choose **My Add-ins**.

    b.  In the **Office Add-ins** dialog box, choose **Shared Folder**.

    c.  Choose **Sheet Switcher Sample**>**Insert**. The add-in opens in a task pane to the right of the current worksheet, as shown in the following figure.

  ![Sheet Switcher Sample](images/SheetSwitcher_taskpane.PNG)

    d.  Click the **Add Sheets!** button. This will add fourteen new sheets to the spreadsheet. Click any of the sheet buttons rendered in the task pane to navigate to that sheet in the workbook.


### Visual Studio version
1.  Copy the project to a local folder and open the Excel-Add-in-Sheet-Switcher.sln in Visual Studio.
2.  Press F5 to build and deploy the sample add-in. Excel launches and the add-in opens in a task pane to the right of a blank worksheet, as shown in the following figure.

  ![Sheet Switcher Sample](images/SheetSwitcher_taskpane.PNG)

3. Click the **Add Sheets!** button. This will add fourteen new sheets to the spreadsheet. Click any of the sheet buttons rendered in the task pane to navigate to that sheet in the workbook.



### Learn more

The Excel JavaScript APIs have much more to offer you as you develop add-ins. The following are just a few of the available resources.

1.  [Excel Add-ins programming overview](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-programming-overview.md)
2.  [Snippet Explorer for Excel](http://officesnippetexplorer.azurewebsites.net/#/snippets/excel)
3.  [Excel Add-ins code samples](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-code-samples.md)
4.  [Excel Add-ins JavaScript API Reference](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-javascript-reference.md)
5.  [Build your first Excel Add-in](https://github.com/OfficeDev/office-js-docs/blob/master/excel/build-your-first-excel-add-in.md)
