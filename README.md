# Step by Step guide on how to create Visual Studio Extension

Nearly all text in this guide comes from [Starting to Develop Visual Studio Extensions](https://msdn.microsoft.com/en-us/library/bb166441.aspx)

### To start you'll need to [install Visual Studio SDK](https://msdn.microsoft.com/en-us/library/mt683786.aspx)

* Go to: Control Panel/Uninstall or change a program/Microsoft Visual Studio/Change/Modify and select Common Tools/Visual Studio Extensibility Tools.
![Control Panel](https://i-msdn.sec.s-msft.com/dynimg/IC846431.jpeg "control panel")
* Note that you must use the Visual Studio installer that matches your installed version of Visual Studio

### Creating Visual Studio Extension with a single command

* You can find the [VSIX](https://blogs.msdn.microsoft.com/quanto/2009/05/26/what-is-a-vsix/) project template in the New Project dialog under Visual C# / Extensibility. Name it **TestProject**.
* When the project opens, add a custom command item template named TestCommand.
In the Solution Explorer, right-click the project node and select Add / New Item.
In the Add New Item dialog, go to Visual C# / Extensibility and select Custom Command.
In the Name field at the bottom of the window, change the command file name to **TestCommand.cs**.

### Building and Testing the Extension

* Build and debug extension as you would any other application.
* An instance of the experimental instance should appear.
* You can test your command by executing **Invoke TestCommand** from **Tools** menu.

### How does it work

Have a look at **TestCommand.cs**
When your extension is loaded each command (a singleton) must be initialized

```cpp
public static void Initialize(Package package)
```

Upon creation new menu item is added

```cpp
OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
if (commandService != null)
{
    var menuCommandID = new CommandID(CommandSet, CommandId);
    var menuItem = new MenuCommand(this.MenuItemCallback, menuCommandID);
    commandService.AddCommand(menuItem);
}
```

And this is the method that gets executed

```cpp
private void MenuItemCallback(object sender, EventArgs e)
{
    string message = string.Format(CultureInfo.CurrentCulture, "Inside {0}.MenuItemCallback()", this.GetType().FullName);
    string title = "TestCommand";

    // Show a message box to prove we were here
    VsShellUtilities.ShowMessageBox(
        this.ServiceProvider,
        message,
        title,
        OLEMSGICON.OLEMSGICON_INFO,
        OLEMSGBUTTON.OLEMSGBUTTON_OK,
        OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
}
```

### How can I change where my command appears

Have a look at **TestCommandPackage.vsct**.
> This is the file that defines the actual layout and type of the commands. It is divided in different sections (e.g. command definition, command placement, ...), with each defining a specific set of properties. See the comment before each section for more details about how to use it.

You'll need to:

1. Create a group for your buttons
   * in our case it's `guid="guidTestCommandPackageCmdSet" id="MyMenuGroup"`
   * which is added to `guid="guidSHLMainMenu" id="IDM_VS_MENU_TOOLS"`
2. Create a button
   * for us it's `<Button guid="guidTestCommandPackageCmdSet" id="TestCommandId" priority="0x0100" type="Button">`
   * which is added to our group `<Parent guid="guidTestCommandPackageCmdSet" id="MyMenuGroup" />`
   * if you want you can set the icon `<Icon guid="guidImages" id="bmpPic1" />`
   * and change the text `<ButtonText>Invoke TestCommand</ButtonText>`
3. Please make sure that values defined in **Symbols** section matches those in your code
   * `<GuidSymbol name="guidTestCommandPackageCmdSet" value="{beafccd3-2ec9-4715-8d31-96d56e289bb3}">`
   matches
   `Guid CommandSet = new Guid("beafccd3-2ec9-4715-8d31-96d56e289bb3");`
   * `<IDSymbol name="TestCommandId" value="0x0100" />`
   matches
   `int CommandId = 0x0100;`
4. To change where command appears change the parent of **MyMenuGroup**
   * e.g. `<Parent guid="guidSHLMainMenu" id="cmdidShellWindowNavigate7" />`
   or `<Parent guid="guidSHLMainMenu" id="IDM_VS_CTXT_ITEMNODE" />`
   or `<Parent guid="guidSHLMainMenu" id="IDM_VS_CTXT_EZDOCWINTAB" />`

### What else can we do

To create an extension you have to provide your own package (**TestCommandPackage.cs** in our case) which derives from **Package**. One of the useful method you get is `protected virtual object GetService(Type serviceType);`
You can get an access to **DTE2** object which according to [msdn](https://msdn.microsoft.com/en-us/library/envdte80.dte2.aspx) is a
> The top-level object in the Visual Studio automation object model.

All you need to do is:
```cpp
DTE2 _dte = GetService(typeof(DTE)) as DTE2;
```

With that **DTE2** you can for example:
* access [tool windows](https://msdn.microsoft.com/en-us/library/envdte80.dte2.toolwindows.aspx) to get selected object from the Solution Explorer

```cpp
        public UIHierarchyItem GetSelectedItem()
        {
            var items = (Array)_dte.ToolWindows.SolutionExplorer.SelectedItems;
            foreach (UIHierarchyItem selItem in items)
            {
                return selItem;
            }
            return null;
        }
```

* access [active document](https://msdn.microsoft.com/en-us/library/0tkyf2yb.aspx)

```cpp
public Document GetActiveDocument()
{
    return _dte.ActiveDocument;
}
```

* get [solution](https://msdn.microsoft.com/en-us/library/envdte._solution.aspx) or [project](https://msdn.microsoft.com/en-us/library/envdte._solution.projects.aspx)

```cpp
public Project GetProject(string projectName)
{
    foreach (Project project in _dte.Solution.Projects)
    {
        if (project.Name == projectName)
            return project;
    }
    return null;
}
```

* add file to a project

```cpp
public static ProjectItem AddFileToProject(Project project, string path, string file)
{
    ProjectItems projectItems = project.ProjectItems;
    String[] paths = path.Split(new Char[] { '\\' });
    for (int i = 0; i < paths.Length - 1; ++i)
    {
        projectItems = projectItems.Item(paths[i]).ProjectItems;
    }
    return projectItems.AddFromFile(file);
}
```

* or [open a file](https://msdn.microsoft.com/en-us/library/envdte.itemoperations.openfile.aspx)

```cpp
public void OpenFile(string filepath)
{
    _dte.ItemOperations.OpenFile(filepath);
}
```

### How do I distribute my extension

When your extension is built **YourProject.vsix** can be found in `bin\Debug` or `bin\Release`.

### Some comments

* Finding parent for your menu item might be quite tricky. You can find all standard groups in `Microsoft Visual Studio 14.0\VSSDK\VisualStudioIntegration\Common\Inc\vsshlids.h`
* If this doesn't help you can [enable VSIP logging](https://blogs.msdn.microsoft.com/dr._ex/2007/04/17/using-enablevsiplogging-to-identify-menus-and-commands-with-vs-2005-sp1/)
```
[HKEY_CURRENT_USER\Software\Microsoft\VisualStudio\14.0\General]
“EnableVSIPLogging”=dword:00000001
```
When the DWORD registry value below is set to 1, CTRL+SHIFT is pressed, and you attempt to display a menu or execute a command, a VSDebug Message dialog will be displayed containing the GUID and ID of that menu or command.
* After a while your experimental instance can be full of old stuff. You can clean it up by cleaning `HKCU\Software\Microsoft\VisualStudio\14.0Exp and HKCU\Software\Microsoft\VisualStudio\14.0Exp_Config` as explainded in [The Experimental Instance](https://msdn.microsoft.com/en-us/library/bb166560.aspx)
