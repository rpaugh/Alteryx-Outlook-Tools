using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Setup
{
    class Program
    {
        static void Main(string[] args)
        {
            Process proc = new Process();
            proc.StartInfo.FileName = "G:\\Media\\Downloads\\print.bat";
            proc.Start();
        }
    }
}
