namespace Spelling.Utilities
{
    /// <summary>
    /// Object that associates a dictionary's path with its contents.
    /// </summary>
    public sealed class DictionaryInfo
    {
        public DictionaryInfo(string path, string contents, string extension = null)
        {
            Path = path;
            Contents = contents;
            Extension = (extension ?? System.IO.Path.GetExtension(path))?.ToUpperInvariant();
        }

        public string Path { get; }

        public string Contents { get; }

        /// <summary>
        /// Either the provided extension or the extension extracted from <see cref="Path"/>, uppercase.
        /// </summary>
        public string Extension { get; }
    }
}
