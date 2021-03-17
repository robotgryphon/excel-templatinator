using HandlebarsDotNet;
using System;
using System.Collections.Generic;
using System.Text;

namespace ScarfPupperBestPupper
{
    internal class TemplateCompiled
    {
        internal HandlebarsTemplate<object, object> Template { get; set; }
        internal string Filename { get; set; }
    }
}
