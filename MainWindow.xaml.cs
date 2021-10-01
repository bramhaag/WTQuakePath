using System;
using System.Data;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using PInvoke;
using SHDocVw;
using Shell32;

namespace WTQuakePath
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        private const int WmKeydown = 0x0100;
        private const int WmKeyup = 0x0101;

        private const int KeyLWin = 0x5B;
        private const int KeyRWin = 0x5C;
        private const int KeyTilde = 0xC0;

        private static bool _winKeyDown;
        private static bool _tildeKeyDown;

        private static User32.SafeHookHandle _hookHandle = User32.SafeHookHandle.Null;

        public MainWindow()
        {
            InitializeComponent();
            _hookHandle = SetHook(HookCallback);
        }

        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
            _hookHandle.Close();
        }

        private static User32.SafeHookHandle SetHook(User32.WindowsHookDelegate proc)
        {
            using var curProcess = Process.GetCurrentProcess();
            using var curModule = curProcess.MainModule;

            if (curModule is null) throw new NoNullAllowedException("curModule cannot be null");

            return User32.SetWindowsHookEx(User32.WindowsHookType.WH_KEYBOARD_LL, proc,
                Kernel32.GetModuleHandle(curModule.ModuleName), 0);
        }

        private static int HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode < 0) return User32.CallNextHookEx(IntPtr.Zero, nCode, wParam, lParam);

            var keyDown = wParam == (IntPtr)WmKeydown;
            if (!keyDown && wParam != (IntPtr)WmKeyup) return User32.CallNextHookEx(IntPtr.Zero, nCode, wParam, lParam);

            var key = Marshal.ReadInt32(lParam);
            // Default case is handled, stop complaining Resharper
            // ReSharper disable once SwitchStatementHandlesSomeKnownEnumValuesWithDefault
            switch (key)
            {
                case KeyLWin:
                case KeyRWin:
                    _winKeyDown = keyDown;
                    break;
                case KeyTilde:
                    _tildeKeyDown = keyDown;
                    break;
                default:
                    return User32.CallNextHookEx(IntPtr.Zero, nCode, wParam, lParam);
            }

            if (!_winKeyDown || !_tildeKeyDown) return User32.CallNextHookEx(IntPtr.Zero, nCode, wParam, lParam);

            var handle = User32.GetForegroundWindow();

            // Workaround for COMException: "8001010d An outgoing call cannot be made since the application is dispatching an input-synchronous call."
            var path = Task.Factory.StartNew(() => GetPath(handle)).Result;

            if (path == null) return User32.CallNextHookEx(IntPtr.Zero, nCode, wParam, lParam);

            var startInfo = new ProcessStartInfo()
            {
                CreateNoWindow = true,
                FileName = "wt.exe",
                Arguments = $@"-w _quake -d ""{path}"""
            };

            var process = new Process()
            {
                StartInfo = startInfo
            };

            process.Start();

            return 1;
        }

        private static string GetPath(IntPtr handle)
        {
            var shellWindows = new ShellWindows();

            // LINQ does not make this code easier to read
            // ReSharper disable once ForeachCanBePartlyConvertedToQueryUsingAnotherGetEnumerator
            foreach (InternetExplorer window in shellWindows)
            {
                // Check if we have the correct window
                if (handle != new IntPtr(window.HWND)) continue;

                // Check if the window is an explorer window
                if (!(window.Document is IShellFolderViewDual2 shellWindow)) return null;

                var currentFolder = shellWindow.Folder.Items().Item();

                // Current folder can be null for some special folders (Like Desktop, although I did not experience this myself)
                // It can also be in the ::{GUID} format for other special folders, such as the "Quick Access" folder
                // Return null in these cases, as there is no viable workaround
                if (currentFolder == null || currentFolder.Path.StartsWith("::")) return null;

                return currentFolder.Path;
            }

            return null;
        }
    }
}