using System;
using System.Collections.Generic;
using System.Text;

namespace ScarfPupperBestPupper
{
    internal class TemplateDataConfig
    {
        public string Input { get; set; }
        public string OutputTemplate { get; set; }

        public string OutputDir { get; set; }

        /// <summary>
        ///  Data file references
        /// </summary>
        public DataFileInfo Data { get; set; }

        public AreaConfig[] Areas { get; set; }
    }

    public class DataFileInfo
    {
        public string File { get; set; }
        public string Range { get; set; }
    }

    public class AreaConfig
    {
        public string Range { get; set; }
    }
}
