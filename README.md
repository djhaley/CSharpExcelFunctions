# CSharpExcelFunctions
*Very* minimal example in .Net 6 of using Blazor to quickly build custom functions for Excel in C#. It uses the default Blazor Server app, and some of the assets from the default Excel custom functions add-in created with yo office.

### The following items have been added to the default Blazor Server app
- wwwroot/excel folder: contains an assets folder with icons from the yo generated add-in, along with taskpane.css and taskpane.js also from the yo add-in.
  - The taskpane.js has a couple of additions: 1) two lines at the top to facilitate providing a JS reference to the C# object that will contain all the custom functions, and 2) an example of calling one of the custom functions in the JS code - "await add(1, 4)".
- Extensions/StringExtensions.cs: helper method for normalizing function names.
- Pages/Excel/TaskPane.razor: contains the HTML from the yo add-in, along with C# code that passes an instance of our custom functions code to JS, using the lines mentioned above in taskpane.js.
- Pages/Excel/Functions.cshtml: generates the JavaScript wrapper that Excel will use to call the custom functions.
- Pages/Excel/FunctionsJson.cshtml: generates the JSON that describes the custom functions for Excel.
- Pages/_Layout.cshtml: contains a block of code to add script/css/etc references if the page is the task pane. I was hoping to use HeadContent but it won't let you use &lt;script&gt; tags.
- Services/ExcelFunctionsService.cs: the actual code for the custom functions. Any method decorated with [JSInvokable] will be exposed to Excel.
- Shared/ExcelLayout.razor: a minimal layout for use by the task pane.
- CSharpExcelFunctions.xml: the manifest file that Excel will use.

### Directions for development use:
- Place the manifest file somewhere you can sideload it into Excel. [Directions for sideloading](https://docs.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins)
- Run the Blazor app using dotnet watch run in the root folder of the project. This will watch for file changes and do hot reloads.
- Insert the add-in into Excel. If all goes well you can now do "=CustomFn.Add(1, 2)", etc, and use the code in ExcelFunctionsService from Excel.
- To make changes to your C# code, simply edit your functions and save them. In Excel do a ctrl-alt-F9 to re-calculate everything and the new code should be applied.
- If you add or remove methods from ExcelFunctionsService, simply re-add the add-in in Excel and the changes will be available.
- The task pane is there simply because I'm planning on using it and it shows that you can use the functions in JavaScript also. It is possible to remove the ribbon command and the task pane and have a UI-less custom functions implementation.