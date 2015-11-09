﻿using System;
using System.ComponentModel.Design;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using EnvDTE;
using EnvDTE80;
using MadsKristensen.AddAnyFile.Templates;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace MadsKristensen.AddAnyFile
{
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [ProvideAutoLoad(UIContextGuids80.SolutionExists)]
    [InstalledProductRegistration("#110", "#112", Version, IconResourceID = 400)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(GuidList.guidAddAnyFilePkgString)]
    public sealed class AddAnyFilePackage : ExtensionPointPackage
    {
        private static DTE2 _dte;
        public const string Version = "1.9";
        private static TemplateMap _templates;
        private static readonly object _templateLock = new object();

        public static IServiceProvider ServiceProvider { get; private set; }

        protected override void Initialize()
        {
            _dte = GetService(typeof(DTE)) as DTE2;
            ServiceProvider = this;

            base.Initialize();

            OleMenuCommandService mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (null != mcs)
            {
                CommandID menuCommandID = new CommandID(GuidList.guidAddAnyFileCmdSet, (int)PkgCmdIDList.cmdidMyCommand);
                MenuCommand menuItem = new MenuCommand(MenuItemCallback, menuCommandID);
                mcs.AddCommand(menuItem);
            }
        }

        private void MenuItemCallback(object sender, EventArgs e)
        {
            UIHierarchyItem item = GetSelectedItem();

            if (item == null)
                return;

            string folder = FindFolder(item);

            if (string.IsNullOrEmpty(folder) || !Directory.Exists(folder))
                return;

            // See if the user has a valid selection on the Solution Tree and avoid prompting the user
            // for a file name.
            Project project = GetActiveProject();
            if (project == null)
                return;

            string input = PromptForFileName(folder).TrimStart('/', '\\').Replace("/", "\\");

            if (string.IsNullOrEmpty(input))
                return;

            TemplateMap templates = GetTemplateMap();

            string projectPath = Path.GetDirectoryName(project.FullName);
            string relativePath;
            if (folder.StartsWith(projectPath, StringComparison.OrdinalIgnoreCase) && folder.Length > projectPath.Length)
                relativePath = Path.Combine(folder.Substring(projectPath.Length + 1), input);
            else
                relativePath = input;

            try
            {
                var itemManager = new ProjectItemManager(_dte, templates);
                var creator = itemManager.GetCreator(folder, relativePath);
                creator.Create(project);

                SelectCurrentItem();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Cannot Add New File", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private static TemplateMap GetTemplateMap()
        {
            TemplateMap templates = null;
            if (_templates == null)
                lock (_templateLock)
                    if (_templates == null)
                    {
                        string path = Path.Combine(
                                        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                                        "VisualStudio.AddAnyFile",
                                        "Patterns.json");
                        if ((templates = TemplateMap.LoadFromFile(path)) == null)
                        {
                            templates = new TemplateMap();
                            templates.LoadDefaultMappings();
                            TemplateMap.WriteToFile(templates, path);
                        }
                        _templates = templates;
                    }
            return _templates;
        }

        private static string PromptForFileName(string folder)
        {
            DirectoryInfo dir = new DirectoryInfo(folder);
            FileNameDialog dialog = new FileNameDialog(dir.Name);
            var result = dialog.ShowDialog();
            return (result.HasValue && result.Value) ? dialog.Input : string.Empty;
        }

        private static string GetRelativePath(Project project, string fullName)
        {
            string projectPath = Path.GetDirectoryName(project.FullName);
            if (fullName.StartsWith(projectPath, StringComparison.OrdinalIgnoreCase))
                return fullName.Substring(projectPath.Length + 1);
            else
                return fullName;
        }

        private static string FindFolder(UIHierarchyItem item)
        {
            Window2 window = _dte.ActiveWindow as Window2;

            if (window != null && window.Type == vsWindowType.vsWindowTypeDocument)
            {
                // if a document is active, use the document's containing directory
                Document doc = _dte.ActiveDocument;
                if (doc != null && !string.IsNullOrEmpty(doc.FullName))
                {
                    ProjectItem docItem = _dte.Solution.FindProjectItem(doc.FullName);

                    if (docItem != null)
                    {
                        string fileName = docItem.Properties.Item("FullPath").Value.ToString();
                        if (File.Exists(fileName))
                            return Path.GetDirectoryName(fileName);
                    }
                }
            }

            string folder = null;

            ProjectItem projectItem = item.Object as ProjectItem;
            Project project = item.Object as Project;

            if (projectItem != null)
            {
                string fileName = projectItem.FileNames[0];

                if (File.Exists(fileName))
                {
                    folder = Path.GetDirectoryName(fileName);
                }
                else
                {
                    folder = fileName;
                }
            }
            else if (project != null)
            {
                Property prop = GetProjectRoot(project);

                if (prop != null)
                {
                    string value = prop.Value.ToString();

                    if (File.Exists(value))
                    {
                        folder = Path.GetDirectoryName(value);
                    }
                    else if (Directory.Exists(value))
                    {
                        folder = value;
                    }
                }
            }
            return folder;
        }

        private static UIHierarchyItem GetSelectedItem()
        {
            var items = (Array)_dte.ToolWindows.SolutionExplorer.SelectedItems;

            foreach (UIHierarchyItem selItem in items)
            {
                return selItem;
            }

            return null;
        }

        public static Project GetActiveProject()
        {
            try
            {
                Array activeSolutionProjects = _dte.ActiveSolutionProjects as Array;

                if (activeSolutionProjects != null && activeSolutionProjects.Length > 0)
                    return activeSolutionProjects.GetValue(0) as Project;
            }
            catch (Exception)
            {
                // Pass through and return null
            }

            return null;
        }

        private static void SelectCurrentItem()
        {
            if (_dte.Version == "11.0") // This errors in VS2012 for some reason.
                return;

            System.Threading.ThreadPool.QueueUserWorkItem((o) =>
            {
                try
                {
                    _dte.ExecuteCommand("View.TrackActivityInSolutionExplorer");
                    _dte.ExecuteCommand("View.TrackActivityInSolutionExplorer");
                }
                catch { /* Ignore any exceptions */ }
            });
        }

        public static Property GetProjectRoot(Project project)
        {
            Property prop;

            try
            {
                prop = project.Properties.Item("FullPath");
            }
            catch (ArgumentException)
            {
                try
                {
                    // MFC projects don't have FullPath, and there seems to be no way to query existence
                    prop = project.Properties.Item("ProjectDirectory");
                }
                catch (ArgumentException)
                {
                    // Installer projects have a ProjectPath.
                    prop = project.Properties.Item("ProjectPath");
                }
            }

            return prop;
        }

        public static void LogToOutputPane(string message)
        {
            EnvDTE.Window window = _dte.Windows.Item(EnvDTE.Constants.vsWindowKindOutput);
            OutputWindow outputWindow = (OutputWindow)window.Object;
            var outputPane = outputWindow.OutputWindowPanes.Cast<OutputWindowPane>().FirstOrDefault(p => p.Name == "Debug");
            if (outputPane != null)
            {
                outputPane.Activate();
                outputPane.OutputString(message);
            }
        }
    }
}
