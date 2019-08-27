using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using Microsoft.Win32;

namespace Spelling.Utilities
{
    internal class Language
    {
        internal bool IsAvailable { get; private set; }

        internal ushort Lcid { get; private set; }

        internal string LexiconFullPath { get; private set; }
        internal string LibraryFullPath { get; private set; }
        internal string Name { get; private set; }

        private const string Hklm = "HKEY_LOCAL_MACHINE\\";

        internal Language(string name, string library, string lexicon, ushort lcid)
        {
            Name = name;
            Lcid = lcid;

            var libraryPaths = Probe($"{library}_x64.dll", lexicon);
            if (libraryPaths != null)
            {
                LibraryFullPath = libraryPaths.Item1;
                LexiconFullPath = libraryPaths.Item2;

                if (LibraryFullPath != null && LexiconFullPath != null)
                {
                    IntPtr handle = KernalNativeMethods.LoadLibrary(LibraryFullPath);

                    if (handle == IntPtr.Zero)
                    {
                        IsAvailable = false;
                    }
                    else
                    {
                        IsAvailable = true;
                        if (!KernalNativeMethods.FreeLibrary(handle))
                        {
                            throw new Win32Exception();
                        }
                    }
                }
            }
        }

        private static Tuple<string, string> Probe(string spellingLibrary, string lexiconLibrary)
        {
            var pathsToOfficeProofingTools = GetPathsToOfficeProofingTools();
            if (pathsToOfficeProofingTools == null)
            {
                return null;
            }

            foreach (string pathToOfficeProofingTools in pathsToOfficeProofingTools)
            {
                string pathToSpellingLibrary = null;
                foreach (var libraryName in new[] { spellingLibrary, "msspell7.dll" })
                {
                    pathToSpellingLibrary = Path.Combine(pathToOfficeProofingTools, libraryName);
                    if (File.Exists(pathToSpellingLibrary))
                    {
                        break;
                    }
                }

                var pathToLexiconLibrary = Path.Combine(pathToOfficeProofingTools, lexiconLibrary);
                if (File.Exists(pathToSpellingLibrary) && File.Exists(pathToLexiconLibrary))
                {
                    return Tuple.Create(pathToSpellingLibrary, pathToLexiconLibrary);
                }
            }

            return null;
        }

        /// <summary>
        /// Gets a path to the Office 2010 proof directory. Returns string.Empty if the path could not be found.
        /// </summary>
        private static IEnumerable<string> GetPathsToOfficeProofingTools()
        {
            var proofDirectories = new List<string>();

            string[] fullInstallRootPaths = new[]
            {
                @"SOFTWARE\Microsoft\Office\14.0\Common\InstallRoot",
                @"SOFTWARE\Microsoft\Office\15.0\Common\InstallRoot",
                @"SOFTWARE\Microsoft\Office\16.0\Common\InstallRoot",
            };

            foreach (string possiblePath in fullInstallRootPaths)
            {
                var registryValue = Registry.GetValue(Hklm + possiblePath, "Path", null) as string;
                if (!string.IsNullOrEmpty(registryValue))
                {
                    var proofDirectory = Path.Combine(registryValue, @"Proof\");
                    if (Directory.Exists(proofDirectory))
                    {
                        proofDirectories.Add(proofDirectory);
                    }
                }
            }

            string[] clickToRunRootPaths = new[]
            {
                @"SOFTWARE\Microsoft\Office\14.0\ClickToRunStore\Applications",
                @"SOFTWARE\Microsoft\Office\15.0\ClickToRunStore\Applications",
                @"SOFTWARE\Microsoft\Office\16.0\ClickToRunStore\Applications",
            };

            foreach (string possiblePath in clickToRunRootPaths)
            {
                using var hklm = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64);
                using var key = hklm?.OpenSubKey(possiblePath);

                var value = key?.GetValue("Word", null) as string;
                if (!string.IsNullOrEmpty(value))
                {
                    var proofDirectory = Path.Combine(Path.GetDirectoryName(value), @"Proof\");
                    if (Directory.Exists(proofDirectory))
                    {
                        proofDirectories.Add(proofDirectory);
                    }
                }
            }

            return proofDirectories;
        }
    }
}
