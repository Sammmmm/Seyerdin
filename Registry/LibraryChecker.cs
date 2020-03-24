using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Windows;

namespace Registry.UI {
    static class LibraryChecker {

        [DllImport("kernel32", CallingConvention = CallingConvention.StdCall)]
        extern static IntPtr LoadLibrary([MarshalAs(UnmanagedType.LPStr)]string lpLibFileName);

        [DllImport("kernel32", CallingConvention = CallingConvention.StdCall)]
        extern static Int32 FreeLibrary(IntPtr hLibModule);

        [DllImport("Kernel32.dll", CallingConvention = CallingConvention.StdCall)]
        static extern IntPtr GetProcAddress(IntPtr hModule, [MarshalAs(UnmanagedType.LPStr)] string lpProcName);

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        delegate UInt32 DllRegUnRegAPI();


        private static readonly List<string> LibrariesToCheck = new List<string>() { "dx8vb.dll", "dx7vb.dll", "SSubTmr6.dll", "MSSTDFMT.dll" };
        private static readonly List<string> LibrariesToCheckSystem = new List<string>() { "fmod.dll", "zlib.dll", "scilexer.dll","SCIVBX.ocx","vbaListView6.ocx" };

        public static bool CheckAssemblyRegistered(string assembly) {
            IntPtr hModuleDLL = IntPtr.Zero;
            try {
                hModuleDLL = LoadLibrary(assembly);
            }
            catch (Exception ex) {
                return false;
            }
            if (hModuleDLL != IntPtr.Zero) {
                FreeLibrary(hModuleDLL);
            }
            else {
                return false;
            }

            return true;
        }

        internal static void CheckLibraries() {
            var brokenLibs = new List<string>();
            var missingLibs = new List<string>();

            foreach (var libraryName in LibrariesToCheck) {
                if (!LibraryChecker.CheckAssemblyRegistered(libraryName)) {
                    brokenLibs.Add(libraryName);
                }
            }

            foreach (var libraryName in LibrariesToCheckSystem) {
                if (!LibraryChecker.CheckAssemblyRegistered(libraryName)) {
                    missingLibs.Add(libraryName);
                }
            }
            
            bool registerFailed = false;

            if (brokenLibs.Count > 0 || missingLibs.Count > 0) {
                var message = $@"The following libraries could not be loaded:
    - {string.Join("\n    - ", brokenLibs.Union(missingLibs))}
This will likely prevent Seyerdin from running correctly.
Would you like to try to register these libraries now?";

                if (MessageBox.Show(message, "Library Load Errors", MessageBoxButton.YesNo) == MessageBoxResult.Yes) {
                    foreach (var libraryName in brokenLibs) {
                        if (!LibraryChecker.RegisterLibrary(libraryName)) {
                            if (!IsRunAsAdmin()) {
                                if (MessageBox.Show(
                                    $"Could not register {libraryName}.  \nYou may need to run this application as an administrator,  Would you like to do that now?",
                                    "Register Library Errors",
                                    MessageBoxButton.YesNo) == MessageBoxResult.Yes) {
                                    LaunchAsAdmin();
                                }
                            }
                            registerFailed = true;
                        }
                    }

                    foreach (var libraryName in missingLibs) {
                        if (!LibraryChecker.CopyLibrary(libraryName)) {
                            if (!IsRunAsAdmin()) {
                                if (MessageBox.Show(
                                    $"Could not copy {libraryName}.  \nYou may need to run this application as an administrator,  Would you like to do that now?",
                                    "Copy Library Errors",
                                    MessageBoxButton.YesNo) == MessageBoxResult.Yes) {
                                    LaunchAsAdmin();
                                }
                            }
                            registerFailed = true;
                        }
                    }
                }
                else {
                    missingLibs.Clear();
                    brokenLibs.Clear();
                }
            }

            if (brokenLibs.Count > 0 || missingLibs.Count > 0) {
                if (registerFailed) {
                    MessageBox.Show("Could not register libraries.  You may need to do this yourself using regsvr32.");
                }
                else {
                    if (MessageBox.Show("Register libraries succeeded.  Consider extracting and installing the vb6 redistributable found in /dependencies if you continue to see errors.  \nYou may need to reboot, would you like to reboot now?", "Reboot?", MessageBoxButton.YesNo) == MessageBoxResult.Yes) {
#if !DEBUG
                        System.Diagnostics.Process.Start("shutdown.exe", "-r -t 0");
#endif
                    }
                }
            }
        }

        public static bool RegisterLibrary(string assembly) {
            var dependencyDir = "dependencies\\";
#if DEBUG
            dependencyDir = "..\\..\\..\\client\\" + dependencyDir;
#endif

            var depPath = dependencyDir + assembly;

            var systemPath = Environment.GetFolderPath(Environment.SpecialFolder.SystemX86) + "\\" + assembly;

            try {
                File.Copy(depPath, systemPath, true);
            }
            catch {
                return false;
            }

            if (depPath.EndsWith(".ocx")) {
                if (File.Exists(depPath.Replace(".ocx", ".oca"))) {
                    try {
                        File.Copy(depPath.Replace(".ocx", ".oca"), systemPath, true);
                    }
                    catch {
                        return false;
                    }
                }
            }

            IntPtr hModuleDLL = LoadLibrary(systemPath); ;

            if (hModuleDLL == IntPtr.Zero) {
                return false;
            }

            IntPtr pExportedFunction = IntPtr.Zero;
            pExportedFunction = GetProcAddress(hModuleDLL, "DllRegisterServer");

            // Obtain the delegate from the exported function, whether it be
            // DllRegisterServer() or DllUnregisterServer().
            var pDelegateRegUnReg =
              (DllRegUnRegAPI)(Marshal.GetDelegateForFunctionPointer(pExportedFunction, typeof(DllRegUnRegAPI)))
              as DllRegUnRegAPI;

            // Invoke the delegate.
            UInt32 hResult = pDelegateRegUnReg();

            FreeLibrary(hModuleDLL);
            hModuleDLL = IntPtr.Zero;

            return hResult == 0;
        }


        public static bool CopyLibrary(string assembly) {
            var dependencyDir = "dependencies\\";
#if DEBUG
            dependencyDir = "..\\..\\..\\client\\" + dependencyDir;
#endif

            var depPath = dependencyDir + assembly;

            var systemPath = Environment.GetFolderPath(Environment.SpecialFolder.SystemX86) + "\\" + assembly;

            try {
                File.Copy(depPath, systemPath, true);
            }
            catch (Exception ex) {
                return false;
            }

            if (depPath.EndsWith(".ocx")) {
                if (File.Exists(depPath.Replace(".ocx",".oca"))) {
                    try {
                        File.Copy(depPath.Replace(".ocx", ".oca"), systemPath.Replace(".ocx",".oca"), true);
                    }
                    catch {
                        return false;
                    }
                }
            }

            return true;
        }

        private static bool IsRunAsAdmin() {
            WindowsIdentity id = WindowsIdentity.GetCurrent();
            WindowsPrincipal principal = new WindowsPrincipal(id);

            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }

        private static bool LaunchAsAdmin() {
            ProcessStartInfo proc = new ProcessStartInfo();
            proc.UseShellExecute = true;
            proc.WorkingDirectory = Environment.CurrentDirectory;
            proc.FileName = Assembly.GetEntryAssembly().CodeBase;

            proc.Verb = "runas";

            try {
                Process.Start(proc);
                Environment.Exit(0);
                return true;
            }
            catch (Exception ex) {
                Console.WriteLine("This program must be run as an administrator! \n\n" + ex.ToString());
                return false;
            }
        }
    }
}
