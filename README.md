# Add Any File

A Visual Studio extension for easily adding new files to any project. Simply hit Shift+F2 to create an empty file in the
selected folder or in the same folder as the selected file.

[![Build status](https://ci.appveyor.com/api/projects/status/252jpryc38qah37x?svg=true)](https://ci.appveyor.com/project/madskristensen/addanyfile)

Download the extension at the
[VS Gallery](http://visualstudiogallery.msdn.microsoft.com/3f820e99-6c0d-41db-aa74-a18d9623b1f3)
or get the
[nightly build](http://vsixgallery.com/extension/27dd9dea-6dd2-403e-929d-3ff20d896c5e/)

### Features

- Easily create any file with any file extension
- Create files starting with a dot like `.gitignore`
- Create deeper folder structures easily if required

![Add new file dialog](art/dialog.png)

### Show the dialog

A new button is added to the context menu in Solution Explorer.

![Add new file dialog](art/menu.png)

You can either click that button or use the keybord shortcut **Shift+F2**.

### Create folders

Create additional folders for your file by using forward-slash to
specify the structure.

For example, by typing **scripts/test.js** in the dialog, the
folder **scripts** is created if it doesn't exist and the file
**test.js** is then placed into it.