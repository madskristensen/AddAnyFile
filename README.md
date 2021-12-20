# Add Any File

[![Build](https://github.com/madskristensen/AddAnyFile/actions/workflows/build.yaml/badge.svg)](https://github.com/madskristensen/AddAnyFile/actions/workflows/build.yaml)
[![VS Marketplace](https://vsmarketplacebadges.dev/version-short/madskristensen.AddNewFile64.svg)](https://marketplace.visualstudio.com/items?itemName=MadsKristensen.AddNewFile64)
![Installs](https://img.shields.io/visual-studio-marketplace/i/madskristensen.AddNewFile?label=Installs&logo=visualstudio)
![Rating](https://vsmarketplacebadges.dev/rating-short/madskristensen.AddNewFile.svg)

A Visual Studio extension for easily adding new files to any project. Simply hit Shift+F2 to create an empty file in the
selected folder or in the same folder as the selected file.

### Features

- Easily create any file with any file extension
- Create files starting with a dot like `.gitignore`
- Create deeper folder structures easily if required
- Create folders when the entered name ends with a /

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

### Custom templates

Create a `.templates` folder at the root of your project.
The templates inside this folder will be used alongside the default ones.

#### Keywords
Inside the template those keywords can be used:
- `{itemname}`: The name of the file without the extension
- `{namespace}`: The namespace

#### Types of template
3 types of template are available:

##### Exact match
When creating the file `Dockerfile`, the extension will look for `dockerfile.txt` template.

##### Convention match or Partial match
If you create a template with the name `repository.txt`, then it will be used when creating a file ending with `Repository` (eg: DataRepository).

##### Extension match
When creating the file `Test.cs`, the extension will look for `cs.txt` template.
