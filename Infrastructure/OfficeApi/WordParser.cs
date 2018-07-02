// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
using System.Linq;
using ApplicationCore.Entities;
using Newtonsoft.Json.Linq;
using System.Xml.Linq;
using ApplicationCore;
using Infrastructure.Services;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using ApplicationCore.Helpers.Exceptions;
using System.Threading.Tasks;
using ApplicationCore.Interfaces;

namespace Infrastructure.OfficeApi
{
    public class WordParser : BaseService<WordParser>, IWordParser
    {
        public WordParser(ILogger<WordParser> logger, IOptions<AppOptions> appOptions) : base(logger, appOptions)
        {
        }

        /// <summary>
        /// RetrieveTOC
        /// </summary>
        /// <param name="fileStream">stream containing the docx file contents</param>
        /// <returns>List of DocumentSection objects</returns>
        public Task<IList<DocumentSection>> RetrieveTOCAsync(Stream fileStream, string requestId = "")
        {
            _logger.LogInformation($"RequestId: {requestId} - RetrieveTOC called.");

            try
            {
                List<DocumentSection> documentSections = new List<DocumentSection>();

                XElement TOC = null;

                using (var document = WordprocessingDocument.Open(fileStream, false))
                {
                    string currentSection1Id = String.Empty;
                    string currentSection2Id = String.Empty;
                    string currentSection3Id = String.Empty;

                    var docPart = document.MainDocumentPart;
                    var doc = docPart.Document;

                    OpenXmlElement block = doc.Descendants<DocPartGallery>().
                        Where(b => b.Val.HasValue &&
                        (b.Val.Value == "Table of Contents")).FirstOrDefault();


                    if (block != null)
                    {
                        // Back up to the enclosing SdtBlock and return that XML.
                        while ((block != null) && (!(block is SdtBlock)))
                        {
                            block = block.Parent;
                        }
                        TOC = new XElement("TOC", block.OuterXml);
                    }


                    // Extract the Table of Contents section information and create the list
                    foreach (var tocPart in document.MainDocumentPart.Document.Body.Descendants<SdtContentBlock>().First())
                    {
                        // Locate each section and add them to the list
                        if (tocPart.InnerXml.Contains("TOC1"))
                        {
                            currentSection1Id = Guid.NewGuid().ToString();

                            // Create a new DocumentSection object and add it to the list
                            documentSections.Add(new DocumentSection
                            {
                                Id = currentSection1Id,
                                SubSectionId = String.Empty,  // TOC1 has no parent
                                DisplayName = tocPart.Descendants<Text>().ToArray()[0].InnerText,
                                LastModifiedDateTime = DateTimeOffset.MinValue,
                                Owner = new UserProfile
                                {
                                    Id = String.Empty,
                                    DisplayName = String.Empty,
                                    Fields = new UserProfileFields()
                                },
                                SectionStatus = ActionStatus.NotStarted
                            });
                        }
                        else if (tocPart.InnerXml.Contains("TOC2"))
                        {
                            currentSection2Id = Guid.NewGuid().ToString();

                            // Create a new DocumentSection object and add it to the list
                            documentSections.Add(new DocumentSection
                            {
                                Id = currentSection2Id,
                                SubSectionId = currentSection1Id,
                                DisplayName = tocPart.Descendants<Text>().ToArray()[0].InnerText,
                                LastModifiedDateTime = DateTimeOffset.MinValue,
                                Owner = new UserProfile
                                {
                                    Id = String.Empty,
                                    DisplayName = String.Empty,
                                    Fields = new UserProfileFields()
                                },
                                SectionStatus = ActionStatus.NotStarted
                            });
                        }
                        else if (tocPart.InnerXml.Contains("TOC3"))
                        {
                            currentSection3Id = Guid.NewGuid().ToString();

                            // Create a new DocumentSection object and add it to the list
                            documentSections.Add(new DocumentSection
                            {
                                Id = currentSection3Id,
                                SubSectionId = currentSection2Id,
                                DisplayName = tocPart.Descendants<Text>().ToArray()[0].InnerText,
                                LastModifiedDateTime = DateTimeOffset.MinValue,
                                Owner = new UserProfile
                                {
                                    Id = String.Empty,
                                    DisplayName = String.Empty,
                                    Fields = new UserProfileFields()
                                },
                                SectionStatus = ActionStatus.NotStarted
                            });
                        }
                    }
                }

                // Return the list of DocumentSections
                return Task.FromResult<IList<DocumentSection>>(documentSections);
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - RetrieveTOC Service Exception: {ex}");
                throw new ResponseException($"RequestId: {requestId} - RetrieveTOC Service Exception: {ex}");
            }
        }
    }
}