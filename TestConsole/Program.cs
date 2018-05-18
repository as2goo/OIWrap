using System;
using OIWrap;

namespace TestConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            var oWordApp = new OIWrap.WordC();

            Console.Write("Starten von Microsoft Word...");
            oWordApp.Init();
            Console.WriteLine("erfolgreich!");
            Console.ReadKey();

        }
    }
}
