using GrandBattleSupport;
using System;
using System.Windows.Forms;

namespace GrandBattleSupport
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            ApplicationConfiguration.Initialize();
            Application.Run(new MainForm());
        }
    }
}