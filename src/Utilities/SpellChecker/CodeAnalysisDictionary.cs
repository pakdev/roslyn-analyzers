using System;
using System.Linq;
using System.Collections.Generic;
using System.IO;
using System.Xml;

namespace Spelling.Utilities
{
    /// <summary>
    /// Parses and stores words/terms obtained from a custom dictionary.
    /// </summary>
    internal sealed class CodeAnalysisDictionary
    {
        public CodeAnalysisDictionary(DictionaryInfo dictionary)
        {
            RecognizedWords = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            UnrecognizedWords = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            CasingExceptions = new HashSet<string>();   // this _is_ case-sensitive
            DiscreteExceptions = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            DeprecatedAlternateWords = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            CompoundAlternateWords = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            LoadWordsIfNecessary(dictionary);
        }

        /// <summary>
        /// Latest date and time of the custom dictionary's last write time/creation time.
        /// </summary>
        public DateTime LastUpdated { get; private set; }

        /// <summary>
        /// A list of misspelled words that the spell checker will now ignore.
        /// </summary>
        /// <example>
        /// <Recognized>knokker</Recognized>
        /// </example>
        public HashSet<string> RecognizedWords { get; }

        /// <summary>
        /// A list of correctly spelled words that the spell checker will now report.
        /// </summary>
        /// <example>
        /// <Word>meth</Word>
        /// </example>
        public HashSet<string> UnrecognizedWords { get; }

        /// <summary>
        /// A list of acronyms that will be identified as spelled and cased correctly.
        /// </summary>
        /// <example>
        /// <Acronym>NESW</Acronym> <!-- North East South West -->
        /// </example>
        public HashSet<string> CasingExceptions { get; }

        /// <summary>
        /// A list of words that will be treated as single, discrete words when the term is checked by the casing rules for compound words.
        /// </summary>
        /// <example>
        /// <Term>checkbox</Term>
        /// </example>
        public HashSet<string> DiscreteExceptions { get; }

        /// <summary>
        /// A mapping of correctly-spelled, deprecated words to suggested alternatives (optional).
        /// </summary>
        /// <example>
        /// <Term PreferredAlternate="LogOn">login</Term>
        /// </example>
        public IDictionary<string, string> DeprecatedAlternateWords { get; }

        /// <summary>
        /// A mapping of single, discrete words to suggested compound alternatives.
        /// </summary>
        /// <remarks>
        /// The inner text term will automatically be added to the <see cref="DiscreteExceptions"/> list.
        /// </remarks>
        /// <example>
        /// <Term CompoundAlternate="CheckBox">checkbox</Term>
        /// </example>
        public IDictionary<string, string> CompoundAlternateWords { get; }

        /// <summary>
        /// Re-parses the dictionary if its path isn't rooted or it was updated from the last time it was parsed.
        /// </summary>
        /// <param name="dictionaryInfo">Object containing a custom dictionary path and contents.</param>
        public void LoadWordsIfNecessary(DictionaryInfo dictionaryInfo)
        {
            DateTime? updateDate = null;

            if (File.Exists(dictionaryInfo.Path))
            {
                var fileInfo = new FileInfo(dictionaryInfo.Path);
                updateDate = new[] { fileInfo.LastWriteTimeUtc, fileInfo.CreationTimeUtc }.Max();

                if (updateDate <= LastUpdated)
                {
                    return;
                }
            }

            LastUpdated = updateDate ?? DateTime.Now;

            // Clear out the word collections
            RecognizedWords.Clear();
            UnrecognizedWords.Clear();
            CasingExceptions.Clear();
            DiscreteExceptions.Clear();
            DeprecatedAlternateWords.Clear();
            CompoundAlternateWords.Clear();

            // Re-parse words
            if (dictionaryInfo.Extension == ".XML")
            {
                LoadWordsFromXml(dictionaryInfo.Contents);
            }
            else if (dictionaryInfo.Extension == ".DIC")
            {
                LoadWordsFromDic(dictionaryInfo.Contents);
            }
        }

        private void LoadWordsFromXml(string contents)
        {
            var document = new XmlDocument();
            try
            {
                using var reader = XmlReader.Create(new StringReader(contents));
                // var reader = new XmlTextReader(new StringReader(contents)) { DtdProcessing = DtdProcessing.Prohibit };
                document.Load(reader);
            }
            catch (XmlException)
            {
                // Silently stop parsing if there's an XML error. This could be changed so that a diagnostic is reported.
                return;
            }

            RecognizedWords.UnionWith(GetWordsFromXml(document, "/Dictionary/Words/Recognized/Word"));
            UnrecognizedWords.UnionWith(GetWordsFromXml(document, "/Dictionary/Words/Unrecognized/Word"));
            CasingExceptions.UnionWith(GetWordsFromXml(document, "/Dictionary/Acronyms/CasingExceptions/Acronym"));
            DiscreteExceptions.UnionWith(GetWordsFromXml(document, "/Dictionary/Words/DiscreteExceptions/Term"));

            foreach (var term in GetTermsFromXml(document, "/Dictionary/Words/Compound/Term", "CompoundAlternate"))
            {
                DeprecatedAlternateWords.Add(term.Item1, term.Item2);
            }

            foreach (var term in GetTermsFromXml(document, "/Dictionary/Words/Deprecated/Term", "PreferredAlternate"))
            {
                DiscreteExceptions.Add(term.Item1);
                CompoundAlternateWords.Add(term.Item1, term.Item2);
            }
        }

        private void LoadWordsFromDic(string contents)
        {
            using var reader = new StringReader(contents);

            string word;
            while ((word = reader.ReadLine()) != null)
            {
                if (!string.IsNullOrWhiteSpace(word.Trim()))
                {
                    RecognizedWords.Add(word);
                }
            }
        }

        private IEnumerable<string> GetWordsFromXml(XmlDocument document, string xPathQuery)
        {
            return GetTermsFromXml(document, xPathQuery, null).Select(word => word.Item1);
        }

        private static IEnumerable<Tuple<string, string>> GetTermsFromXml(
            XmlDocument document,
            string xPathQuery,
            string attributeName)
        {
            var xmlNodeList = document.SelectNodes(xPathQuery);
            if (xmlNodeList == null)
            {
                yield break;
            }

            foreach (XmlNode node in xmlNodeList)
            {
                var attributeValue = string.Empty;
                if (attributeName != null)
                {
                    var attribute = node.Attributes?[attributeName];
                    if (attribute != null)
                    {
                        attributeValue = attribute.Value;
                    }
                }

                if (string.IsNullOrWhiteSpace(node.InnerText.Trim()))
                {
                    continue;
                }

                yield return new Tuple<string, string>(node.InnerText, attributeValue);
            }
        }
    }
}
