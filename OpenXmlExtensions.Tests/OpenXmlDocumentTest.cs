using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Xunit;

namespace OpenXmlDocument.Tests
{
    public class OpenXmlDocumentTest : IDisposable
    {
        public OpenXmlDocumentTest()
        {
            foreach (var (documentName, _, _) in GetDocumentsCharacteristic())
            {
                File.Copy($"Assets/{documentName}", documentName, true);
            }
        }

        public static IEnumerable<object[]> GetDocumentPaths()
        {
            foreach (var (documentName, _, _) in GetDocumentsCharacteristic())
            {
                yield return new object[] { documentName };
            }
        }

        public static IEnumerable<object[]> GetDocumentsPathWithTagsNumber()
        {
            foreach (var (documentName, tagsNumber, _) in GetDocumentsCharacteristic().Where(d => d.Item2 != 0))
            {
                yield return new object[] { documentName, tagsNumber };
            }
        }

        public static IEnumerable<object[]> GetDocumentsPathWithAliasNumber()
        {
            foreach (var (documentName, _, aliasNumber) in GetDocumentsCharacteristic())
            {
                yield return new object[] { documentName, aliasNumber };
            }
        }

        public void Dispose()
        {
            foreach (var (documentName, _, _) in GetDocumentsCharacteristic())
            {
                File.Delete($"{documentName}");
            }
        }

        [Theory]
        [MemberData(nameof(GetDocumentPaths))]
        public void ShouldGetContentFromFilePath(string documentPath)
        {
            // Assign
            string actualDocumentContent;

            // Act
            using (var document = new OpenXmlDocument(documentPath))
            {
                actualDocumentContent = document.GetContent();
            }

            // Assert
            Assert.NotEmpty(actualDocumentContent);
        }

        [Theory]
        [MemberData(nameof(GetDocumentPaths))]
        public void ShouldGetContentFromStream(string documentPath)
        {
            // Assign
            string actualDocumentContent;

            // Act
            using (var stream = File.Open(documentPath, FileMode.Open))
            {
                using (var document = new OpenXmlDocument(stream))
                {
                    actualDocumentContent = document.GetContent();
                }
            }

            // Assert
            Assert.NotEmpty(actualDocumentContent);
        }

        [Theory]
        [MemberData(nameof(GetDocumentsPathWithTagsNumber))]
        public void ShouldGetTagNamesForCorrespondingToMarkupFromFilePath(string documentPath, int tagsNumber)
        {
            // Arrange
            var escapedBeginMarkup = @"\|\{";
            var escapedEndMarkup = @"\}\|";
            var expectedTags = new List<string>()
            {
                "Title",
                "Subtitle"
            };
            IEnumerable<string> actualTagsToReplace;

            // Act
            using (var document = new OpenXmlDocument(documentPath))
            {
                actualTagsToReplace = document.GetTagNames(escapedBeginMarkup, escapedEndMarkup);
            }

            // Assert
            Assert.NotEmpty(actualTagsToReplace);
            Assert.Equal(tagsNumber, actualTagsToReplace.Count());
            Assert.All(expectedTags, expectedTag => Assert.Contains(expectedTag, actualTagsToReplace));
        }

        [Theory]
        [MemberData(nameof(GetDocumentsPathWithTagsNumber))]
        public void ShouldGetTagNamesForCorrespondingToMarkupFromStream(string documentPath, int tagsNumber)
        {
            // Arrange
            var escapedBeginMarkup = @"\|\{";
            var escapedEndMarkup = @"\}\|";
            var expectedTags = new List<string>()
            {
                "Title",
                "Subtitle"
            };
            IEnumerable<string> actualTagsToReplace;

            // Act
            using (var stream = File.Open(documentPath, FileMode.Open))
            {
                using (var document = new OpenXmlDocument(stream))
                {
                    actualTagsToReplace = document.GetTagNames(escapedBeginMarkup, escapedEndMarkup);
                }
            }

            // Assert
            Assert.NotEmpty(actualTagsToReplace);
            Assert.Equal(tagsNumber, actualTagsToReplace.Count());
            Assert.All(expectedTags, expectedTag => Assert.Contains(expectedTag, actualTagsToReplace));
        }

        [Theory]
        [MemberData(nameof(GetDocumentsPathWithTagsNumber))]
        public void ShouldGetTagNamesFromFilePath(string documentPath, int tagsNumber)
        {
            // Arrange
            var expectedTags = new List<string>()
            {
                "Title",
                "Subtitle"
            };
            IEnumerable<string> actualTagsToReplace;

            // Act
            using (var document = new OpenXmlDocument(documentPath))
            {
                actualTagsToReplace = document.GetTagNames();
            }

            // Assert
            Assert.NotEmpty(actualTagsToReplace);
            Assert.Equal(tagsNumber, actualTagsToReplace.Count());
            Assert.All(expectedTags, expectedTag => Assert.Contains(expectedTag, actualTagsToReplace));
        }

        [Theory]
        [MemberData(nameof(GetDocumentsPathWithTagsNumber))]
        public void ShouldGetTagNamesFromStream(string documentPath, int tagsNumber)
        {
            // Arrange
            var expectedTags = new List<string>()
            {
                "Title",
                "Subtitle"
            };
            IEnumerable<string> actualTagsToReplace;

            // Act
            using (var stream = File.Open(documentPath, FileMode.Open))
            {
                using (var document = new OpenXmlDocument(stream))
                {
                    actualTagsToReplace = document.GetTagNames();
                }
            }

            // Assert
            Assert.NotEmpty(actualTagsToReplace);
            Assert.Equal(tagsNumber, actualTagsToReplace.Count());
            Assert.All(expectedTags, expectedTag => Assert.Contains(expectedTag, actualTagsToReplace));
        }

        [Theory]
        [MemberData(nameof(GetDocumentsPathWithAliasNumber))]
        public void ShouldGetAliasNamesFromFilePath(string documentPath, int aliasNumber)
        {
            // Arrange
            IEnumerable<string> actualAliasNames;

            // Act
            using (var document = new OpenXmlDocument(documentPath))
            {
                actualAliasNames = document.GetAliasNames();
            }

            // Assert
            Assert.NotNull(actualAliasNames);
            Assert.Equal(aliasNumber, actualAliasNames.Count());
        }

        [Theory]
        [MemberData(nameof(GetDocumentPaths))]
        public void ShouldSearchAndReplaceTagNamesFromFilePath(string documentPath)
        {
            // Arrange
            var tags = new Dictionary<string, string>()
            {
                 { "Title", "Lorem ipsum dolor sit amet" },
                 { "Subtitle", "Proin blandit porta mauris nec convallis" }
            };
            string actualDocumentContent;

            // Act
            using (var document = new OpenXmlDocument(documentPath))
            {
                document.SearchAndReplaceTags(tags);
                actualDocumentContent = document.GetContent();
            }

            // Assert
            Assert.Contains("Lorem ipsum dolor sit amet", actualDocumentContent);
            Assert.Contains("Proin blandit porta mauris nec convallis", actualDocumentContent);
        }

        [Fact]
        public void ShouldSearchAndReplaceTagNames()
        {
            // Arrange
            var tags = new Dictionary<string, string>() { };
            var (documentPath, _, _) = GetDocumentsCharacteristic().First(d => d.Item2 == 0);
            string actualDocumentContent;

            // Act
            using (var document = new OpenXmlDocument(documentPath))
            {
                document.SearchAndReplaceTags(tags);
                actualDocumentContent = document.GetContent();
            }

            // Assert
            Assert.Contains("Lorem ipsum dolor sit amet", actualDocumentContent);
            Assert.Contains("Proin blandit porta mauris nec convallis", actualDocumentContent);
        }


        [Theory]
        [MemberData(nameof(GetDocumentPaths))]
        public void ShouldSearchAndReplaceTagNamesFromStream(string documentPath)
        {
            // Arrange
            var tags = new Dictionary<string, string>()
            {
                 { "Title", "Lorem ipsum dolor sit amet" },
                 { "Subtitle", "Proin blandit porta mauris nec convallis" }
            };
            string actualDocumentContent;

            // Act
            using (var stream = File.Open(documentPath, FileMode.Open))
            {
                using (var document = new OpenXmlDocument(stream))
                {
                    document.SearchAndReplaceTags(tags);
                    actualDocumentContent = document.GetContent();
                }
            }

            // Assert
            Assert.Contains("Lorem ipsum dolor sit amet", actualDocumentContent);
            Assert.Contains("Proin blandit porta mauris nec convallis", actualDocumentContent);
        }

        [Theory]
        [MemberData(nameof(GetDocumentPaths))]
        public void ShouldRemoveAllAliases(string documentPath)
        {
            // Arrange
            IEnumerable<string> actualAliasNames;

            // Act
            using (var document = new OpenXmlDocument(documentPath))
            {
                actualAliasNames = document.GetAliasNames();
                document.SearchAndRemoveAliases(actualAliasNames);
                actualAliasNames = document.GetAliasNames();
            }

            // Assert
            Assert.Empty(actualAliasNames);
        }

        [Fact]
        public void WhenEmptyFilePath_ShouldThrowsArgumentException()
        {
            // Assign
            var emptyFilePath = string.Empty;

            // Act
            Action act = () => new OpenXmlDocument(emptyFilePath);

            // Assert
            Assert.Throws<ArgumentException>(act);
        }

        [Fact]
        public void WhenInvalidStream_ShouldThrowsFileFormatException()
        {
            // Assign
            var invalidContent = Guid.NewGuid().ToString();
            var byteArray = Encoding.ASCII.GetBytes(invalidContent);
            var invalidSream = new MemoryStream(byteArray);

            // Act
            Action act = () => new OpenXmlDocument(invalidSream);

            // Assert
            Assert.Throws<FileFormatException>(act);
        }

        [Fact]
        public void WhenNullFilePath_ShouldThrowsArgumentNullException()
        {
            // Assign
            string nullFilePath = null;

            // Act
            Action act = () => new OpenXmlDocument(nullFilePath);

            // Assert
            Assert.Throws<ArgumentNullException>(act);
        }

        [Fact]
        public void WhenNullStream_ShouldThrowsArgumentNullException()
        {
            // Assign
            Stream nullStream = null;

            // Act
            Action act = () => new OpenXmlDocument(nullStream);

            // Assert
            Assert.Throws<ArgumentNullException>(act);
        }

        [Fact]
        public void WhenUnknownFilePath_ShouldThrowsFileNotFoundException()
        {
            // Assign
            var unknownFilePath = Guid.NewGuid().ToString();

            // Act
            Action act = () => new OpenXmlDocument(unknownFilePath);

            // Assert
            Assert.Throws<FileNotFoundException>(act);
        }

        private static IEnumerable<(string, int, int)> GetDocumentsCharacteristic()
        {
            yield return ("SampleWithAlias.docx", 0, 6);
            yield return ("SampleWithTags.docx", 2, 0);
            yield return ("Sample.docx", 0, 0);
        }
    }
}
