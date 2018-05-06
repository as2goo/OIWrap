using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OIWrap;
using System.Timers;

namespace OIWrap
{
    public class Word
    {
        public void Init()
        {
            var oBase = new Main();

            var oApp = new Microsoft.Office.Interop.Word.Application();

            oApp.Visible = true;
            oApp.Documents.Add();
            /// oBase.StringOut;
        }
    }
}
