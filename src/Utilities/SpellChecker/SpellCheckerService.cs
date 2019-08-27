using System;
using System.Linq;
using System.Collections.Generic;
using System.Globalization;
using System.Threading;

namespace Spelling.Utilities
{
    /// <summary>
    /// A global service that manages the <see cref="WindowsSpellChecker"/> and the custom dictionaries
    /// that apply to specific assemblies.
    /// </summary>
    /// <remarks>
    /// Adapted from <see href="https://github.com/StyleCop/StyleCop/blob/master/Project/Src/StyleCop/Spelling/NamingService.cs" />
    /// </remarks>
    public class SpellCheckerService : IDisposable
    {
        private static readonly Dictionary<string, SpellCheckerService> ServiceCache = new Dictionary<string, SpellCheckerService>();

        private static readonly object ServiceCacheLock = new object();

        private readonly IDictionary<string, CodeAnalysisDictionary> _pathToDictionaryMap;
        private readonly IDictionary<string, HashSet<string>> _assemblyNameToPathsMap;

        private static SpellCheckerService _defaultSpellCheckerService;

        private WindowsSpellChecker _spellChecker;

        private SpellCheckerService(CultureInfo culture)
        {
            Culture = culture;

            _assemblyNameToPathsMap = new Dictionary<string, HashSet<string>>();
            _pathToDictionaryMap = new Dictionary<string, CodeAnalysisDictionary>();
            _spellChecker = WindowsSpellChecker.FromCulture(culture);
        }

        public static SpellCheckerService DefaultSpellCheckerService => _defaultSpellCheckerService ??
            (_defaultSpellCheckerService = GetSpellCheckerService(CultureInfo.CurrentCulture));

        public CultureInfo Culture { get; }

        public bool SupportsSpelling => _spellChecker != null;

        /// <summary>
        /// Gets a spellchecker service for the specified <paramref name="culture"/>.
        /// </summary>
        /// <param name="culture">
        /// The culture to use.
        /// </param>
        /// <returns>
        /// The SpellCheckerService for the culture.
        /// </returns>
        public static SpellCheckerService GetSpellCheckerService(CultureInfo culture)
        {
            if (culture == null)
            {
                throw new ArgumentNullException(nameof(culture));
            }

            lock (ServiceCacheLock)
            {
                SpellCheckerService service;
                if (!ServiceCache.TryGetValue(culture.Name, out service))
                {
                    service = new SpellCheckerService(culture);
                    ServiceCache[culture.Name] = service;
                }

                return service;
            }
        }

        /// <summary>
        /// Parses custom dictionaries and associates the values with a specific <paramref name="assemblyName"/>.
        /// </summary>
        /// <param name="dictionaryInfos">
        /// An enumerable of objects containing custom dictionary path and contents.
        /// </param>
        /// <param name="assemblyName">
        /// Name of the assembly that should consult the dictionaries.
        /// </param>
        public void LoadDictionariesForAssembly(IEnumerable<DictionaryInfo> dictionaryInfos, string assemblyName)
        {
            // Find the assembly and clear it's associated paths. If the assembly was never added to the cache, add it.
            HashSet<string> paths;
            if (!_assemblyNameToPathsMap.TryGetValue(assemblyName, out paths))
            {
                paths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                _assemblyNameToPathsMap.Add(assemblyName, paths);
            }
            else
            {
                paths.Clear();
            }

            foreach (var dictionaryInfo in dictionaryInfos)
            {
                // Associate this custom dictionary with the current assembly
                paths.Add(dictionaryInfo.Path);

                CodeAnalysisDictionary dictionary;
                if (!_pathToDictionaryMap.TryGetValue(dictionaryInfo.Path, out dictionary))
                {
                    // If we haven't already created a CodeAnalysisDictionary from this path, do so now
                    _pathToDictionaryMap.Add(dictionaryInfo.Path, new CodeAnalysisDictionary(dictionaryInfo));
                }
                else
                {
                    // Otherwise, check to see if the file needs to be re-parsed.
                    dictionary.LoadWordsIfNecessary(dictionaryInfo);
                }
            }
        }

        /// <summary>
        /// Check spelling of the given <paramref name="word"/>.
        /// </summary>
        /// <param name="word">
        /// The word to check.
        /// </param>
        /// <param name="assemblyName">
        /// Name of the assembly used to find the appropriate CodeAnalysisDictionaries.
        /// </param>
        /// <returns>
        /// The Spelling.WordSpelling value for the given <paramref name="word"/>.
        /// </returns>
        public WordSpelling CheckWordSpellingInAssembly(string word, string assemblyName)
        {
            if (!SupportsSpelling)
            {
                throw new InvalidOperationException();
            }

            HashSet<string> paths;
            if (_assemblyNameToPathsMap.TryGetValue(assemblyName, out paths))
            {
                // This assembly has custom dictionaries
                foreach (var dictionary in paths.Select(path => _pathToDictionaryMap[path]))
                {
                    if (dictionary.RecognizedWords.Contains(word))
                    {
                        return WordSpelling.SpelledCorrectly;
                    }

                    if (dictionary.UnrecognizedWords.Contains(word))
                    {
                        return WordSpelling.Unrecognized;
                    }
                }
            }

            return _spellChecker.Check(word);
        }

        /// <summary>
        /// Cleans up all resources.
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            var spellChecker = Interlocked.Exchange(ref _spellChecker, null);
            spellChecker?.Dispose();
        }
    }
}
