using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OIWrap;
using System.Timers;
using Microsoft.Office.Interop;

namespace OIWrap
{
    public class WordC
    {
        public Microsoft.Office.Interop.Word.Application oApp { get; set; }

        public void Init()
        {
            var oBase = new Main();

            //var oApp = new Microsoft.Office.Interop.Word.Application();

            //oApp.Visible = true;

            this.Start();
            this.oApp.Documents.Add();
            
            /// oBase.StringOut;
        }

        public Boolean Start()
        {
            this.oApp = new Microsoft.Office.Interop.Word.Application();
            this.oApp.Visible = true;
            return true;
        }
    }
}
