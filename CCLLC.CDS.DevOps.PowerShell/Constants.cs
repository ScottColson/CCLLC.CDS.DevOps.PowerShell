using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CCLLC.Cds.DevOps.PowerShell
{
    public static class Constants
    {
        public const string EntityLogicalName = "documenttemplate";

        public class Fields
        {
            public const string Id = "documenttemplateid";
            public const string Name = "name";
            public const string DocumentType = "documenttype";
            public const string AssociatedEntityTypeCode = "associatedentitytypecode";
            public const string Content = "content";
        }

        public class FileExtensions
        {
            public const string Word = ".docx";
            public const string Excel = ".xlsx";
        }
    }
}
