using System;
using System.Linq;
using System.Collections.Immutable;
using System.Threading;
using Microsoft.CodeAnalysis;
using Spelling;
using Spelling.Utilities;
using System.IO;

namespace Text.Analyzers
{
    /// <summary>
    /// Service that interacts with the global SpellCheckerService to provide guidance on whether a word 
    /// is misspelled or should be named something else.
    /// </summary>
    /// <remarks>
    /// The naming service is created each time an analyzer fires
    /// </remarks>
    public sealed class NamingService
    {
        private readonly string _assemblyName;
        private readonly SpellCheckerService _spellCheckerService;

        public NamingService(string assemblyName)
        {
            _assemblyName = assemblyName;
            _spellCheckerService = SpellCheckerService.DefaultSpellCheckerService;
        }

        /// <summary>
        /// Boolean indicating if the spell checker is available or not.
        /// </summary>
        public bool SupportsSpelling => _spellCheckerService.SupportsSpelling;

        /// <summary>
        /// Searches for custom dictionaries within an analyzer's additional files and ensures they'll
        /// be used for the current assembly.
        /// </summary>
        /// <param name="additionalFiles">Files that are specifically added for analyzer consumption.</param>
        /// <param name="cancellationToken">Token that allows loading to be canceled.</param>
#pragma warning disable CA1801 // Review unused parameters - shouldn't be necessary (bug)
        public void LoadDictionaries(ImmutableArray<AdditionalText> additionalFiles, CancellationToken cancellationToken = default(CancellationToken))
        {
#pragma warning restore CA1801 // Review unused parameters
            var dictionaries =
                from file in additionalFiles
                let extension = Path.GetExtension(file.Path)?.ToUpperInvariant()
                where
                    (file.Path.IndexOf("dictionary", StringComparison.OrdinalIgnoreCase) >= 0 && extension == ".XML") ||
                    (file.Path.IndexOf("custom", StringComparison.OrdinalIgnoreCase) >= 0 && extension == ".DIC")
                select new DictionaryInfo(file.Path, file.GetText(cancellationToken).ToString(), extension);

            if (!cancellationToken.IsCancellationRequested)
            {
                _spellCheckerService.LoadDictionariesForAssembly(dictionaries, _assemblyName);
            }
        }

        /// <summary>
        /// Returns an value indicating whether the provided <paramref name="word"/> is correctly spelled.
        /// </summary>
        /// <param name="word"></param>
        /// <returns></returns>
        public WordSpelling CheckSpelling(string word)
        {
            return _spellCheckerService.CheckWordSpellingInAssembly(word, _assemblyName);
        }
    }
}
