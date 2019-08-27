using System;

namespace Spelling.Utilities
{
    internal enum ProofLexType
    {
        ChangeAlways = 1,
        ChangeOnce = 0,
        Exclude = 3,
        IgnoreAlways = 2,
        Main = 4,
        Max = 6,
        SysUdr = 5,
        User = 2
    }

    internal enum PtecMajor
    {
        BufferTooSmall = 6,
        IoErrorMainLex = 3,
        IoErrorUserLex = 4,
        ModuleError = 2,
        ModuleNotLoaded = 8,
        NoErrors = 0,
        NotFound = 7,
        NotSupported = 5,
        OutOfMemory = 1
    }

    internal enum PtecMinor
    {
        EntryTooLong = 0x8f,
        FileCreate = 0x8a,
        FileOpenError = 0x92,
        FileRead = 0x88,
        FileShare = 0x8b,
        FileTooLargeError = 0x93,
        FileWrite = 0x89,
        InvalidCmd = 0x85,
        InvalidEntry = 0x8e,
        InvalidFormat = 0x86,
        InvalidId = 0x81,
        InvalidLanguage = 150,
        InvalidMainLex = 0x83,
        InvalidUserLex = 0x84,
        InvalidWsc = 130,
        MainLexCountExceeded = 0x90,
        ModuleAlreadyBusy = 0x80,
        ModuleNotTerminated = 140,
        OperNotMatchedUserLex = 0x87,
        ProtectModeOnly = 0x95,
        UserLexCountExceeded = 0x91,
        UserLexFull = 0x8d,
        UserLexReadOnly = 0x94
    }

    [Flags]
    internal enum SpellerStates
    {
        IsContinued = 1,
        IsEditedChange = 4,
        NoStateInfo = 0,
        StartsSentence = 2
    }

    internal enum SpellerStatus
    {
        NoErrors,
        UnknownInputWord,
        ReturningChangeAlways,
        ReturningChangeOnce,
        InvalidHyphenation,
        ErrorCapitalization,
        WordConsideredAbbreviation,
        HyphenChangesSpelling,
        NoMoreSuggestions,
        MoreInfoThanBufferCouldHold,
        NoSentenceStartCap,
        RepeatWord,
        ExtraSpaces,
        MissingSpace,
        InitialNumeral,
        NoErrorsUdHit,
        ReturningAutoReplace,
        ErrorAccent
    }
}
