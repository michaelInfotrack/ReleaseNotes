using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Pencil.Infrastructure.WordAutomation
{
    public class WordApplicationHelpers
    {    
        public static Microsoft.Office.Interop.Word.Application CreateNewWordInstance()
        {
            return new Microsoft.Office.Interop.Word.Application();
        }        
    }
}
