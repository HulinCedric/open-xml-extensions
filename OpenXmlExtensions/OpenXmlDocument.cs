using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlPowerTools;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace OpenXmlDocument
{
    public class OpenXmlDocument : IDisposable
    {
        private static readonly string BeginMarkup = @"|{";

        private static readonly string EndMarkup = @"}|";

        private static readonly string EscapedBeginMarkup = @"\|\{";

        private static readonly string EscapedEndMarkup = @"\}\|";

        private readonly Stream stream;

        public OpenXmlDocument(string filePath) : this(File.Open(filePath, FileMode.Open))
        {
        }

        public OpenXmlDocument(Stream stream)
        {
            CheckValidityOfStream(stream);
            this.stream = stream;
        }

        public void Dispose() => stream.Dispose();

        public IEnumerable<string> GetAliasNames()
        {
            using (var wordDoc = WordprocessingDocument.Open(stream, false))
            {
                return GetAliases(wordDoc).Select(x => x.Val.ToString()).Distinct().ToList();
            }
        }

        public string GetContent()
        {
            using (var document = WordprocessingDocument.Open(stream, false))
            {
                return GetDocumentContent(document);
            }
        }

        public IEnumerable<string> GetTagNames(string escapedBeginMarkup, string escapedEndMarkup)
        {
            var regex = new Regex($"({escapedBeginMarkup})(.*?)({escapedEndMarkup})");
            using (var wordDoc = WordprocessingDocument.Open(stream, false))
            {
                return wordDoc.MainDocumentPart.Document.Body
                    .SelectMany(x => regex.Matches(x.InnerText))
                    .Select(m => m.Groups.ElementAt(2).Value).Distinct().ToList();
            }
        }

        public IEnumerable<string> GetTagNames()
            => GetTagNames(EscapedBeginMarkup, EscapedEndMarkup);

        public void SearchAndRemoveAliases(IEnumerable<string> aliasNamesToRemove)
        {
            if (aliasNamesToRemove == null || !aliasNamesToRemove.Any())
            {
                return;
            }

            using (var wordDoc = WordprocessingDocument.Open(stream, true))
            {
                var elementsToRemove = GetAncestorsElementsByAliasNames(wordDoc, aliasNamesToRemove);
                foreach (var elementToRemove in elementsToRemove)
                {
                    elementToRemove.Remove();
                }
            }
        }

        public void SearchAndReplaceTags(IDictionary<string, string> tagNamesToReplace)
        {
            if (tagNamesToReplace == null || !tagNamesToReplace.Any())
            {
                return;
            }

            using (var wordDoc = WordprocessingDocument.Open(stream, true))
            {
                var tagNamesWithMarkupToReplace = MapDictionnaryWithMarkupOnKey(tagNamesToReplace);
                foreach (var (tag, text) in tagNamesWithMarkupToReplace)
                {
                    TextReplacer.SearchAndReplace(wordDoc, tag, text, false);
                }
            }
        }

        private void CheckValidityOfStream(Stream stream)
        {
            using (var document = WordprocessingDocument.Open(stream, false)) { }
        }

        private IEnumerable<SdtAlias> GetAliases(WordprocessingDocument wordDoc)
            => wordDoc.MainDocumentPart.Document.Descendants<SdtAlias>();

        private IEnumerable<SdtAlias> GetAliasesByName(WordprocessingDocument wordDoc, IEnumerable<string> aliasNames)
            => from alias in GetAliases(wordDoc)
               join aliasName in aliasNames
               on alias.Val.Value equals aliasName
               select alias;

        private IEnumerable<SdtElement> GetAncestorsElementsByAliasNames(WordprocessingDocument wordDoc, IEnumerable<string> aliasNames)
            => from alias in GetAliasesByName(wordDoc, aliasNames).ToList()
               from element in alias.Ancestors<SdtElement>()
               select element;

        private string GetDocumentContent(WordprocessingDocument document)
        {
            using (var streamReader = new StreamReader(document.MainDocumentPart.GetStream()))
            {
                return streamReader.ReadToEnd();
            }
        }

        private IDictionary<string, string> MapDictionnaryWithMarkupOnKey(IDictionary<string, string> tagsToReplace)
            => tagsToReplace.ToDictionary(f => $"{BeginMarkup}{f.Key}{EndMarkup}", f => f.Value);
    }
}
