using System;
using System.Linq;
using System.Collections.Generic;
using System.Globalization;

namespace Spelling.Utilities
{
    internal static class TextInfoExtensions
    {
        public static string ToTitleCase(this TextInfo textInfo, string word)
        {
            var lowerWord = textInfo.ToLower(word);
            return textInfo.ToUpper(lowerWord[0]) + new string(lowerWord.Skip(1).ToArray());
        }
    }

    /// <summary>
    /// A spell checker implemented with Microsoft Office components and only works for Windows.
    /// </summary>
    /// <remarks>
    /// Adapted from <see href="https://github.com/StyleCop/StyleCop/blob/master/Project/Src/StyleCop/Spelling/SpellChecker.cs"/>
    /// </remarks>
    internal sealed class WindowsSpellChecker : SpellChecker
    {
        internal const int MaximumTextLength = 0x40;

        // If additional langauges are eventually added, uncomment as necessary
        private static readonly Language[] Languages = new[]
        {
            //new Language("AR", "mssp7ar", "mssp7ar.lex", 0xc01),
            //new Language("CS", "mssp7cz", "mssp7cz.lex", 0x405),
            //new Language("DA", "mssp7da", "mssp7da.lex", 0x406),
            //new Language("DE", "mssp7ge", "mssp7geP.lex", 0x407),
            //new Language("EN", "mssp7en", "mssp7en.lex", 0x409),
            new Language("EN-US", "mssp7en", "mssp7en.lex", 0x409),
            //new Language("EN-GB", "mssp7en", "mssp7en.lex", 0x809),
            //new Language("EN-AU", "mssp7en", "mssp7en.lex", 0xc09),
            //new Language("EN-CA", "mssp7en", "mssp7en.lex", 0x1009),
            //new Language("EN-NZ", "mssp7en", "mssp7en.lex", 0x809),
            //new Language("ES", "mssp7es", "mssp7es.lex", 0xc0a),
            //new Language("FI", "mssp7fi", "mssp7fi.lex", 0x40b),
            //new Language("FR", "mssp7fr", "mssp7fr.lex", 0x40c),
            //new Language("GL", "mssp7gl", "mssp7gl.lex", 0x456),
            //new Language("HE", "mssp7hb", "mssp7hb.lex", 0x40d),
            //new Language("ID", "mssp7in", "mssp7in.lex", 0x421),
            //new Language("IT", "mssp7it", "mssp7it.lex", 0x410),
            //new Language("KO", "mssp7ko", "mssp7ko.lex", 0x412),
            //new Language("LT", "mssp7lt", "mssp7lt.lex", 0),
            //new Language("NL", "mssp7nl", "mssp7nl.lex", 0x413),
            //new Language("NB", "mssp7nb", "mssp7nb.lex", 0x414),
            //new Language("NN", "mssp7no", "mssp7no.lex", 0x814),
            //new Language("PL", "mssp7pl", "mssp7pl.lex", 0x415),
            //new Language("PT", "mssp7pt", "mssp7pt.lex", 0x816),
            //new Language("PT-BR", "mssp7pb", "mssp7pb.lex", 0x416),
            //new Language("RO", "mssp7ro", "mssp7ro.lex", 0),
            //new Language("RU", "mssp7ru", "mssp7ru.lex", 0x419),
            //new Language("SV", "mssp7sw", "mssp7sw.lex", 0x41d),
            //new Language("TR", "mssp7tr", "mssp7tr.lex", 0x41f),
            //new Language("UK", "mssp7ua", "mssp7ua.lex", 0x422)
        };

        private static readonly TextInfo UsaTextInfo = new CultureInfo("en-US").TextInfo;

        // The Languages array above needs to be initialized before this static executes.
        private static readonly Dictionary<string, Language> LanguageTable = BuildLanguageTable();

        private Dictionary<string, WordSpelling> _wordSpellingCache = new Dictionary<string, WordSpelling>();

        private WindowsSpellChecker(Language language) : base(language.LibraryFullPath)
        {
            AddLexicon(language.Lcid, language.LexiconFullPath);
        }

        public static WindowsSpellChecker FromCulture(CultureInfo culture)
        {
            if (culture == null)
            {
                throw new ArgumentNullException(nameof(culture));
            }

            if (culture.Equals(CultureInfo.InvariantCulture))
            {
                return null;
            }

            if (LanguageTable.TryGetValue(culture.Name.ToUpperInvariant(), out Language language) && language.IsAvailable)
            {
                return new WindowsSpellChecker(language);
            }

            return FromCulture(culture.Parent);
        }

        public WordSpelling Check(string text)
        {
            WordSpelling spelledCorrectly;
            if (text == null)
            {
                throw new ArgumentNullException(nameof(text));
            }

            if (text.Length == 0)
            {
                return WordSpelling.SpelledCorrectly;
            }

            lock (_wordSpellingCache)
            {
                if (_wordSpellingCache.TryGetValue(text, out spelledCorrectly))
                {
                    return spelledCorrectly;
                }

                var status = CheckUnsafe(text);
                if (status != SpellerStatus.NoErrors)
                {
                    status = CheckUnsafe(UsaTextInfo.ToTitleCase(text));
                }

                if (status == SpellerStatus.NoErrors)
                {
                    spelledCorrectly = WordSpelling.SpelledCorrectly;
                }
                else
                {
                    spelledCorrectly = WordSpelling.Unrecognized;
                }

                _wordSpellingCache[text] = spelledCorrectly;
            }

            return spelledCorrectly;
        }

        private static Dictionary<string, Language> BuildLanguageTable()
        {
            Dictionary<string, Language> dictionary = new Dictionary<string, Language>(Languages.Length, StringComparer.OrdinalIgnoreCase);
            foreach (Language language in Languages)
            {
                dictionary.Add(language.Name, language);
            }

            return dictionary;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _wordSpellingCache = null;
            }
            base.Dispose(disposing);
        }
    }
}
