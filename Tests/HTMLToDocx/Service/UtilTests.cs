using System;
using System.Collections.Generic;
using System.IO;
using Moq;
using Xunit;
using Microsoft.Extensions.Logging;
using DocConverterFunctionApp;

namespace Tests
{
    public class UtilTests
    {
        [Fact]
        public void ValidateHtml_ValidHtml_ReturnsOuterHtml()
        {
            // Arrange
            var htmlContent = "<html><body><h1>Hello, World!</h1></body></html>";
            var mockLogger = new Mock<ILogger>();

            // Act
            var result = Util.ValidateHtml(htmlContent, mockLogger.Object);

            // Assert
            Assert.Equal(htmlContent, result);
        }

        [Fact]
        public void SimplifyHtml_ValidHtml_RemovesScriptAndStyleTags()
        {
            // Arrange
            var htmlContent = "<html><head><style>body { color: red; }</style></head><body><script>alert('Hi');</script><h1>Hello, World!</h1></body></html>";
            var expectedContent = "<html><head></head><body><h1>Hello, World!</h1></body></html>";
            var mockLogger = new Mock<ILogger>();

            // Act
            var result = Util.SimplifyHtml(htmlContent, mockLogger.Object);

            // Assert
            Assert.Equal(expectedContent, result);
        }

        [Fact]
        public void FindCommonBaseDirectory_ValidPaths_ReturnsCommonBaseDirectory()
        {
            // Arrange
            var basePath = Path.GetFullPath(Path.Combine("home", "user", "project"));
            var resourceFiles = new List<string>
            {
                Path.Combine(basePath, "file1.txt"),
                Path.Combine(basePath, "subdir", "file2.txt"),
                Path.Combine(basePath, "subdir2", "file3.txt")
            };
            var expectedCommonBase = basePath;
            var mockLogger = new Mock<ILogger>();

            // Act
            var result = Util.FindCommonBaseDirectory(resourceFiles);

            // Assert
            Assert.Equal(expectedCommonBase, result);
        }
    }
}