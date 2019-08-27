using System;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.InteropServices;
using static Spelling.Utilities.NativeMethods;

namespace Spelling.Utilities
{
    internal abstract class SpellChecker : IDisposable
    {
        private SpellerAddUdr _addUdr;
        private SpellerCheck _check;
        private SpellerClearUdr _clearUdr;
        private ProofCloseLex _closeLex;
        private SpellerDelUdr _deleteUdr;
        private ProofOpenLex _openLex;
        private ProofTerminate _terminate;

        private IntPtr[] _lexicons;

        private IntPtr _id;
        private IntPtr _libraryHandle;
        private IntPtr _ignoredDictionary;

        private delegate NativeMethods.Ptec ProofCloseLex(IntPtr id, IntPtr lex, bool force);
        private delegate NativeMethods.Ptec ProofInit(out IntPtr pid, ref NativeMethods.ProofParams pxpar);
        private delegate NativeMethods.Ptec ProofOpenLex(
            IntPtr id, ref NativeMethods.ProofLexIn plxin, ref NativeMethods.ProofLexOut plxout);
        private delegate NativeMethods.Ptec ProofSetOptions(IntPtr id, uint iOptionSelect, uint iOptVal);
        private delegate NativeMethods.Ptec ProofTerminate(IntPtr id, bool fForce);
        private delegate NativeMethods.Ptec SpellerCheck(
            IntPtr sid, SpellerCommand scmd, ref NativeMethods.Wsib psib, ref NativeMethods.Wsrb psrb);
        private delegate NativeMethods.Ptec SpellerClearUdr(IntPtr sid, IntPtr lex);
        private delegate NativeMethods.Ptec SpellerDelUdr(
            IntPtr sid, IntPtr lex, [MarshalAs(UnmanagedType.LPTStr)] string delete);
        private delegate NativeMethods.Ptec SpellerAddUdr(
            IntPtr sid, IntPtr lex, [MarshalAs(UnmanagedType.LPTStr)] string add);

        private delegate IntPtr SpellerBuiltInUdr(IntPtr sid, ProofLexType lxt);

        protected SpellChecker(string path)
        {
            _libraryHandle = KernalNativeMethods.LoadLibrary(path);
            if (_libraryHandle == IntPtr.Zero)
            {
                throw new Win32Exception();
            }

            ProofInit proc = GetProc<ProofInit>(_libraryHandle, "SpellerInit");
            ProofSetOptions proofSetOptions = GetProc<ProofSetOptions>(_libraryHandle, "SpellerSetOptions");
            _terminate = GetProc<ProofTerminate>(_libraryHandle, "SpellerTerminate");
            _openLex = GetProc<ProofOpenLex>(_libraryHandle, "SpellerOpenLex");
            _closeLex = GetProc<ProofCloseLex>(_libraryHandle, "SpellerCloseLex");
            _check = GetProc<SpellerCheck>(_libraryHandle, "SpellerCheck");
            _addUdr = GetProc<SpellerAddUdr>(_libraryHandle, "SpellerAddUdr");
            _deleteUdr = GetProc<SpellerDelUdr>(_libraryHandle, "SpellerDelUdr");
            _clearUdr = GetProc<SpellerClearUdr>(_libraryHandle, "SpellerClearUdr");
            ProofParams pxpar = new ProofParams { VersionApi = 0x3000000 };
            CheckErrorCode(proc(out IntPtr ptr, ref pxpar));
            _id = ptr;
            CheckErrorCode(proofSetOptions(ptr, 0, 0x20006));
            InitIgnoreDictionary();
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            try
            {
                if (_lexicons != null)
                {
                    foreach (IntPtr ptr in _lexicons)
                    {
                        CheckErrorCode(_closeLex(_id, ptr, true));
                    }

                    _lexicons = null;
                }

                if (_id != IntPtr.Zero)
                {
                    CheckErrorCode(_terminate(_id, true));
                    _id = IntPtr.Zero;
                }

                if (_libraryHandle != IntPtr.Zero)
                {
                    if (!KernalNativeMethods.FreeLibrary(_libraryHandle))
                    {
                        throw new Win32Exception();
                    }

                    _libraryHandle = IntPtr.Zero;
                }
            }
            finally
            {
                if (disposing)
                {
                    _terminate = null;
                    _closeLex = null;
                    _openLex = null;
                    _check = null;
                    _addUdr = null;
                    _clearUdr = null;
                    _deleteUdr = null;
                }
            }
        }

        protected void AddIgnoredWord(string word)
        {
            CheckErrorCode(_addUdr(_id, _ignoredDictionary, word));
        }

        protected void AddLexicon(ushort lcid, string path)
        {
            var plxin = new ProofLexIn { pwszLex = path, lxt = ProofLexType.Main, lidExpected = lcid };
            var plxout = new ProofLexOut { CchCopyright = 0, readOnly = true };
            CheckErrorCode(_openLex(_id, ref plxin, ref plxout));
            AddLexicon(plxout.lex);
        }

        protected unsafe SpellerStatus CheckUnsafe(string word)
        {
            char* pwsz = stackalloc char[65];
            SpellerSuggestion* prgsugg =
                stackalloc SpellerSuggestion[checked(1 * sizeof(SpellerSuggestion) / sizeof(SpellerSuggestion))];

            fixed (IntPtr* lexicons2 = _lexicons)
            {
                Wsib wsib = default;
                wsib.pwsz = word;
                wsib.ichStart = 0u;
                wsib.Cch = (UIntPtr)((ulong)word.Length);
                wsib.CchUse = wsib.Cch;
                wsib.prglex = lexicons2;
                wsib.Clex = (UIntPtr)((ulong)_lexicons.Length);
                wsib.sstate = SpellerStates.StartsSentence;

                Wsrb wsrb = default;
                wsrb.pwsz = pwsz;
                wsrb.CchAlloc = 65u;
                wsrb.cszAlloc = 1u;
                wsrb.prgsugg = prgsugg;

                Ptec error;
                lock (this)
                {
                    error = _check(_id, SpellerCommand.VerifyBuffer, ref wsib, ref wsrb);
                }

                CheckErrorCode(error);
                return wsrb.sstat;
            }
        }

        protected void ClearIgnoredWords()
        {
            CheckErrorCode(_clearUdr(_id, _ignoredDictionary));
        }

        protected void RemoveIgnoredWord(string word)
        {
            CheckErrorCode(_deleteUdr(_id, _ignoredDictionary, word));
        }

        private static void CheckErrorCode(Ptec error)
        {
            if (!error.Succeeded)
            {
                throw new InvalidOperationException(
                    string.Format(
                        CultureInfo.CurrentCulture,
                        "Unexpected proofing tool error code: {0}.",
                        error));
            }
        }

        private static T GetProc<T>(IntPtr library, string procName) where T : class
        {
            IntPtr procAddress = KernalNativeMethods.GetProcAddress(library, procName);
            if (procAddress == IntPtr.Zero)
            {
                throw new Win32Exception();
            }

            return (T)((object)Marshal.GetDelegateForFunctionPointer(procAddress, typeof(T)));
        }

        private void AddLexicon(IntPtr lex)
        {
            IntPtr[] ptrArray;
            int length;
            if (_lexicons == null)
            {
                ptrArray = new IntPtr[1];
                length = 0;
            }
            else
            {
                ptrArray = new IntPtr[_lexicons.Length + 1];
                _lexicons.CopyTo(ptrArray, 0);
                length = _lexicons.Length;
            }

            ptrArray[length] = lex;
            _lexicons = ptrArray;
        }

        private void InitIgnoreDictionary()
        {
            SpellerBuiltInUdr proc = GetProc<SpellerBuiltInUdr>(_libraryHandle, "SpellerBuiltinUdr");
            _ignoredDictionary = proc(_id, ProofLexType.User);
            if (_ignoredDictionary == IntPtr.Zero)
            {
                throw new InvalidOperationException("Failed to get the ignored dictionary handle.");
            }
        }

        private enum SpellerCommand
        {
            Anagram = 7,
            Suggest = 3,
            SuggestMore = 4,
            VerifyBuffer = 2,
            VerifyBufferAutoReplace = 10,
            Wildcard = 6
        }
    }
}
