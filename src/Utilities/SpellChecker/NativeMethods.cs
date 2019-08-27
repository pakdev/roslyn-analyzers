using System;
using System.Globalization;
using System.Runtime.InteropServices;

namespace Spelling.Utilities
{
    internal class NativeMethods
    {
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        internal struct ProofLexIn
        {
            internal string pwszLex;
            internal bool create;
            internal ProofLexType lxt;
            internal ushort lidExpected;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        internal struct ProofLexOut
        {
            internal string pwszCopyright;
            internal IntPtr lex;
            internal uint CchCopyright;
            internal uint version;
            internal bool readOnly;
            internal ushort lid;
        }

        [StructLayout(LayoutKind.Sequential)]
        internal struct ProofParams
        {
            internal uint VersionApi;
        }

        [StructLayout(LayoutKind.Sequential)]
        internal struct Ptec
        {
            internal uint Code;

            internal PtecMajor Major => ((PtecMajor)Code) & ((PtecMajor)0xff);

            internal PtecMinor Minor => (PtecMinor)(Code >> 0x10);

            internal bool Succeeded => Code == 0;

            public override string ToString()
            {
                string str = $"0x{Code.ToString("X", CultureInfo.InvariantCulture)} -- {Major}";
                if (Minor != 0)
                {
                    str = $"{str}:{Minor}";
                }

                return str;
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        internal struct SpellerSuggestion
        {
            internal unsafe char* pwsz;
            internal uint ichSugg;
            internal uint CchSugg;
            internal uint rating;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        internal struct Wsib
        {
            internal string pwsz;
            internal unsafe IntPtr* prglex;
            internal UIntPtr Cch;
            internal UIntPtr Clex;
            internal SpellerStates sstate;
            internal uint ichStart;
            internal UIntPtr CchUse;
        }

        [StructLayout(LayoutKind.Sequential)]
        internal struct Wsrb
        {
            internal unsafe char* pwsz;
            internal unsafe SpellerSuggestion* prgsugg;
            internal uint ichError;
            internal uint CchError;
            internal uint ichProcess;
            internal uint CchProcess;
            internal SpellerStatus sstat;
            internal uint csz;
            internal uint cszAlloc;
            internal uint CchMac;
            internal uint CchAlloc;
        }
    }
}
