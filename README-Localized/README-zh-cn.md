# <a name="sheet-switcher-task-pane-add-in-sample-for-excel-2016"></a>适用于 Excel 2016 的工作表切换程序任务窗格外接程序示例

_适用于：Excel 2016_

此任务窗格外接程序使你能够通过任务窗格向工作簿添加新工作表，并轻松地导航到 Excel 2016 中不同的工作表。它有两种类型：文本编辑器和 Visual Studio。

![工作表切换程序示例](../images/SheetSwitcher_taskpane.PNG)

## <a name="try-it-out"></a>尝试一下
### <a name="text-editor-version"></a>文本编辑器版本

部署和测试外接程序的最简单方法是将文件复制到网络共享中。

1.  复制文本编辑器文件夹中的文件，然后使用你常用的服务器托管这些文件。
2.  编辑清单文件 (SheetSwitcherManifest.xml) 中的 \<SourceLocation\> 和 \<Url\> 元素，使其指向第 1 步中的托管位置（即 https://localhost/SheetSwitcher/home.html）
3.  将清单文件 (SheetSwitcherManifest.xml) 复制到网络共享（例如，\\\MyShare\SheetSwitcher）中。
4.  在 Excel 中将内含清单的共享位置添加为受信任的应用目录。

    a.启动 Excel 并打开一个空白的电子表格。

    b.依次选择“**文件**”选项卡和“**选项**”。

    c.依次选择“**信任中心**”和“**信任中心设置**”按钮。

    d.选择“**受信任的应用程序目录**”。

    e.在“**目录 URL**”框中，输入你在第 1 步中创建的网络共享路径，然后选择“**添加目录**”。

   f.  选中“**显示在菜单中**”复选框，然后选择“**确定**”。此时，系统会显示一条消息，提醒你注意你的设置将在 Office 下次启动时应用。

5.  测试并运行外接程序。

    a.在 Excel 2016 的“**插入**”选项卡中，选择“**我的外接程序**”。

    b.在“**Office 外接程序**”对话框中，选择“**共享文件夹**”。

    c.依次选择“**工作表切换程序示例**”>“**插入**”。此时，系统会在当前工作表右侧的任务窗格中打开外接程序，如下图所示。

  ![工作表切换程序示例](../images/SheetSwitcher_taskpane.PNG)

    d.单击“**添加工作表!**”按钮。这将向电子表格添加十四个新工作表。单击任务窗格中呈现的任一工作表按钮，在工作簿中转到相应的工作表。


### <a name="visual-studio-version"></a>Visual Studio 版本
1.  将项目复制到本地文件夹，并在 Visual Studio 中打开 Excel-Add-in-Sheet-Switcher.sln。
2.  按 F5 生成并部署示例外接程序。Excel 启动并且外接程序会在空白工作簿右侧的任务窗格中打开，如下图所示。

  ![工作表切换程序示例](../images/SheetSwitcher_taskpane.PNG)

3. 单击“**添加工作表!**”按钮。这将向电子表格添加十四个新工作表。单击呈现在任务窗格中的任意一个工作表按钮，以导航到工作簿中的工作表。



### <a name="learn-more"></a>了解详细信息

在您开发外接程序时，Excel JavaScript API 可以提供更多功能。下面只是其中一些可用资源。

1.  [Excel 外接程序编程概述](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-programming-overview.md)
2.  [适用于 Excel 的代码段资源管理器](http://officesnippetexplorer.azurewebsites.net/#/snippets/excel)
3.  [Excel 外接程序代码示例](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-code-samples.md)
4.  [Excel 外接程序 JavaScript API 参考](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-javascript-reference.md)
5.  [生成你的第一个 Excel 外接程序](https://github.com/OfficeDev/office-js-docs/blob/master/excel/build-your-first-excel-add-in.md)