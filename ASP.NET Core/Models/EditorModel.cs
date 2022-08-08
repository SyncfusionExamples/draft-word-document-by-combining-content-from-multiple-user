using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace DocumentEditorApp.Models
{
    public class FilesPathInfo
    {
        public string text;
    }

    public class CustomParams
    {
        public string fileName
        {
            get;
            set;
        }
    }

    public class CustomParameter
    {
        public string fileName
        {
            get;
            set;
        }
        public string documentData
        {
            get;
            set;
        }
    }
}
