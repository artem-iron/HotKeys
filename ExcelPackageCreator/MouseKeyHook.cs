using Gma.System.MouseKeyHook;
using System;
using System.IO;
using System.Windows.Forms;

namespace ExcelPackageCreator
{
    internal class MouseKeyHook
    {
        private IKeyboardMouseEvents _globalHook;

        public void Subscribe()
        {
            if (_globalHook == null)
            {
                // Note: for the application hook, use the Hook.AppEvents() instead
                _globalHook = Hook.GlobalEvents();
                _globalHook.KeyPress += GlobalHookKeyPress;
            }
        }

        private static void GlobalHookKeyPress(object sender, KeyPressEventArgs e)
        {
            Console.WriteLine($"Key:{e.KeyChar}");

            ZipFileCreator.HandleZipCreation(e.KeyChar);
        }

        public void Unsubscribe()
        {
            if (_globalHook != null)
            {
                _globalHook.KeyPress -= GlobalHookKeyPress;
                _globalHook.Dispose();
            }
        }
    }
}
