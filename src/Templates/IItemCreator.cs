using EnvDTE;

namespace MadsKristensen.AddAnyFile.Templates
{
    interface IItemCreator
    {
        void Create(Project project);
    }
}