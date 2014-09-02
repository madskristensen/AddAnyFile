using System;

namespace MadsKristensen.AddAnyFile
{
    static class GuidList
    {
        public const string guidAddAnyFilePkgString = "27dd9dea-6dd2-403e-929d-3ff20d896c5e";
        public const string guidAddAnyFileCmdSetString = "32af8a17-bbbc-4c56-877e-fc6c6575a8cf";

        public static readonly Guid guidAddAnyFileCmdSet = new Guid(guidAddAnyFileCmdSetString);
    }

    static class PkgCmdIDList
    {
        public const uint cmdidMyCommand = 0x100;
    }
}