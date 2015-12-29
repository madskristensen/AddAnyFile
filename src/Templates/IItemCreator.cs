using EnvDTE;

namespace MadsKristensen.AddAnyFile.Templates
{
    interface IItemCreator
    {
        ItemInfo Create(Project project);
    }

    public class ItemInfo
    {
        public string FileName { get; set; }
        public string Extension { get; set; }

        public static readonly ItemInfo Empty = new ItemInfo();
    }
}