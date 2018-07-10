// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Newtonsoft.Json.Linq;
using SmartLink.Entity;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.Caching;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace SmartLink.Service
{
    public class DocumentService : IDocumentService
    {
        private readonly IConfigService _configService;

        public DocumentService(IConfigService configService)
        {
            _configService = configService;
        }

        protected string ClientID { get { return _configService.WebJobClientId; } }
        protected string Authority
        {
            get
            {
                return _configService.AzureAdInstance + _configService.AzureAdTenantId;
            }
        }

        protected string Resource
        {
            get
            {
                return _configService.SharePointUrl;
            }
        }
        protected string CertificatedPassword
        {
            get
            {
                return _configService.CertificatePassword;
            }
        }

        protected ObjectCache LocalCache
        {
            get
            {
                return MemoryCache.Default;
            }
        }

        public async Task<DocumentUpdateResult> UpdateBookmrkValue(string documentId, IEnumerable<DestinationPoint> destinationPoints, string value)
        {
            DocumentUpdateResult retValue = new DocumentUpdateResult() { IsSuccess = true };
            try
            {
                string authHeader = GetAuthorizationHeader();
                DocumentCheckResult documentResult = await GetDocumentUrlByID(new DocumentCheckResult() { DocumentId = documentId });
                if (documentResult.IsSuccess)
                {
                    FileContextInfo fileContextInfo = await GetFileContextInfo(authHeader, documentResult.DocumentUrl);
                    var stream = await GetFileStream(authHeader, documentResult.DocumentUrl, fileContextInfo);
                    stream.Seek(0, SeekOrigin.Begin);
                    UpdateStream(destinationPoints, value, stream, retValue);
                    await UploadStream(authHeader, documentResult.DocumentUrl, stream, fileContextInfo);
                    return retValue;
                }
                else
                {
                    retValue.IsSuccess = false;
                    retValue.Message.Add(documentResult.Message);
                }
            }
            catch (Exception ex)
            {
                retValue.IsSuccess = false;
                retValue.Message.Add(ex.ToString());
            }
            return retValue;
        }

        static private async Task UploadStream(string authHeader, string destinationFileName, Stream stream, FileContextInfo contextInfo)
        {
            var fileAbsolutePath = new Uri(destinationFileName).AbsolutePath;
            int index = fileAbsolutePath.LastIndexOf('/');
            var filePath = fileAbsolutePath.Substring(0, index);
            var fileName = fileAbsolutePath.Substring(index + 1);

            using (var client = new WebClient())
            {
                client.Headers.Add(HttpRequestHeader.Authorization, authHeader);
                client.Headers.Add(HttpRequestHeader.Accept, "application/json");

                var uploadUri = new Uri($"{contextInfo.WebUrl}/_api/web/GetFolderByServerRelativeUrl(@filePath)/Files/add(url='{fileName}',overwrite=true)?@filePath='{filePath}'");
                var returnValue = await client.UploadDataTaskAsync(uploadUri, "POST", (stream as MemoryStream).ToArray());
            }
        }
        static public void UpdateStream(IEnumerable<DestinationPoint> destinationPoints, string value, Stream stream, DocumentUpdateResult updateResult)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(stream, true))
            {
                foreach (var destinationPoint in destinationPoints)
                {
                    int foundTimes = 0;
                    var position = destinationPoint.RangeId;
                    string updatedTo = string.Empty;

                    if (destinationPoint.ReferencedSourcePoint.SourceType == SourceTypes.Point)
                    {
                        var formats = destinationPoint.CustomFormats.OrderBy(c => c.GroupOrderBy).ToList();
                        updatedTo = GetFormattedValue(value, formats, destinationPoint.DecimalPlace);
                        //Update value in sdtBlock
                        foundTimes += UpdateValueInSdtBlock(updatedTo, wordDoc, position);
                        //Update value in sdtCell
                        foundTimes += UpdateValueInSdtCell(updatedTo, wordDoc, position);
                        //Update value in sdtRun
                        foundTimes += UpdateValueInSdtRun(updatedTo, wordDoc, position);
                    }
                    else if (destinationPoint.ReferencedSourcePoint.SourceType == SourceTypes.Table)
                    {
                        JObject json = JObject.Parse(value);
                        if (destinationPoint.DestinationType == DestinationTypes.TableImage)
                        {
                            //Update table in sdtPic
                            updatedTo = json["image"].Value<string>();
                            foundTimes += UpdateValueInSdtPic(updatedTo, wordDoc, position);
                        }
                        else
                        {
                            //Update table in sdtTable
                            updatedTo = ((JObject)json["table"]).ToString();
                            foundTimes += UpdateValueInSdtTable(updatedTo, wordDoc, position);
                        }
                    }
                    else
                    {
                        //Update image in sdtPic
                        updatedTo = value;
                        foundTimes += UpdateValueInSdtPic(updatedTo, wordDoc, position);
                    }
                    updateResult.Message.Add($"Tag:{position} found {foundTimes}(updated to {updatedTo}).");
                }

                wordDoc.MainDocumentPart.Document.Save();
            }
        }
        static private int UpdateValueInSdtRun(string value, WordprocessingDocument wordDoc, string position)
        {
            int retValue = 0;
            var items = wordDoc.MainDocumentPart.Document.Descendants<SdtRun>().Where(
                o =>
                {
                    var tagedItem = o.SdtProperties.Elements<Tag>().FirstOrDefault();
                    if (tagedItem != null)
                    {
                        return tagedItem.Val == position;
                    }
                    return false;
                });
            foreach (var item in items)
            {
                var texts = item.Descendants<Text>();
                var textCount = texts.Count();
                if (textCount == 1)
                {
                    texts.First().Text = value;
                    retValue++;
                }
                else if (textCount > 1)
                {
                    texts.First().Text = value;
                    for (int i = 1; i < textCount; i++)
                    {
                        texts.ElementAt(i).Text = string.Empty;
                    }
                    /*
                    //remove all paragraphs from the content cell
                    item.SdtContentRun.RemoveAllChildren<Paragraph>();
                    //create a new paragraph containing a run and a text element
                    Paragraph newParagraph = new Paragraph();
                    Run newRun = new Run();
                    Text newText = new Text(value);
                    newRun.Append(newText);
                    newParagraph.Append(newRun);
                    item.Append(newParagraph);
                    */
                    retValue++;
                }
            }
            return retValue;
        }
        static private int UpdateValueInSdtBlock(string value, WordprocessingDocument wordDoc, string position)
        {
            int retValue = 0;
            var items = wordDoc.MainDocumentPart.Document.Descendants<SdtBlock>().Where(
                o =>
                {
                    var tagedItem = o.SdtProperties.Elements<Tag>().FirstOrDefault();
                    if (tagedItem != null)
                    {
                        return tagedItem.Val == position;
                    }
                    return false;
                });
            foreach (var item in items)
            {
                var texts = item.Descendants<Text>();
                var textCount = texts.Count();
                if (textCount == 1)
                {
                    texts.First().Text = value;
                    retValue++;
                }
                else if (textCount > 1)
                {
                    texts.First().Text = value;
                    for (int i = 1; i < textCount; i++)
                    {
                        texts.ElementAt(i).Text = string.Empty;
                    }
                    /*
                    //remove all paragraphs from the content cell
                    item.SdtContentBlock.RemoveAllChildren<Paragraph>();
                    //create a new paragraph containing a run and a text element
                    Paragraph newParagraph = new Paragraph();
                    Run newRun = new Run();
                    Text newText = new Text(value);
                    newRun.Append(newText);
                    newParagraph.Append(newRun);
                    item.Append(newParagraph);
                    */
                    retValue++;
                }
            }
            return retValue;
        }
        static private int UpdateValueInSdtCell(string value, WordprocessingDocument wordDoc, string position)
        {
            int retValue = 0;
            var items = wordDoc.MainDocumentPart.Document.Descendants<SdtCell>().Where(
                o =>
                {
                    var tagedItem = o.SdtProperties.Elements<Tag>().FirstOrDefault();
                    if (tagedItem != null)
                    {
                        return tagedItem.Val == position;
                    }
                    return false;
                });
            foreach (var item in items)
            {
                var tableCell = item.SdtContentCell.Elements<TableCell>().FirstOrDefault();
                if (tableCell != null)
                {
                    var texts = tableCell.Descendants<Text>();
                    var textCount = texts.Count();
                    if (textCount == 1)
                    {
                        texts.First().Text = value;
                        retValue++;
                    }
                    else if (textCount > 1)
                    {
                        texts.First().Text = value;
                        for (int i = 1; i < textCount; i++)
                        {
                            texts.ElementAt(i).Text = string.Empty;
                        }
                        /*
                        //remove all paragraphs from the content cell
                        tableCell.RemoveAllChildren<Paragraph>();
                        //create a new paragraph containing a run and a text element
                        Paragraph newParagraph = new Paragraph();
                        Run newRun = new Run();
                        Text newText = new Text(value);
                        newRun.Append(newText);
                        newParagraph.Append(newRun);
                        tableCell.Append(newParagraph);
                        */
                        retValue++;
                    }
                }
            }
            return retValue;
        }
        static private int UpdateValueInSdtPic(string value, WordprocessingDocument wordDoc, string position)
        {
            int retValue = 0;
            var items = wordDoc.MainDocumentPart.Document.Descendants<SdtElement>().Where(
                o =>
                {
                    var tagedItem = o.SdtProperties.Elements<Tag>().FirstOrDefault();
                    if (tagedItem != null)
                    {
                        return tagedItem.Val == position;
                    }
                    return false;
                });
            foreach (var item in items)
            {
                string embed = null;
                Drawing dr = item.Descendants<Drawing>().FirstOrDefault();
                if (dr != null)
                {
                    DocumentFormat.OpenXml.Drawing.Blip blip = dr.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault();
                    if (blip != null)
                        embed = blip.Embed;
                }

                if (embed != null)
                {
                    IdPartPair idpp = wordDoc.MainDocumentPart.Parts.Where(pa => pa.RelationshipId == embed).FirstOrDefault();
                    if (idpp != null)
                    {
                        ImagePart ip = (ImagePart)idpp.OpenXmlPart;
                        byte[] data = Convert.FromBase64String(value);
                        MemoryStream ms = new MemoryStream(data);
                        ip.FeedData(ms);
                    }
                }
            }
            return retValue;
        }
        static private int UpdateValueInSdtTable(string value, WordprocessingDocument wordDoc, string position)
        {
            JObject json = JObject.Parse(value);
            JArray values = (JArray)json["values"];
            JArray formats = (JArray)json["formats"];
            int index = 0;

            int retValue = 0;
            var items = wordDoc.MainDocumentPart.Document.Descendants<SdtElement>().Where(o =>
            {
                var tagedItem = o.SdtProperties.Elements<Tag>().FirstOrDefault();
                if (tagedItem != null)
                {
                    return tagedItem.Val == position;
                }
                return false;
            });
            foreach (var item in items)
            {
                //Find the table 
                var table = item.Descendants<Table>().FirstOrDefault();
                var content = item.Descendants<SdtContentBlock>().FirstOrDefault();
                if (table == null)
                {
                    //Create new table
                    content.RemoveAllChildren();
                    table = new Table(new TableProperties(new TableStyle() { Val = "TableGrid" }, new TableWidth() { Width = "auto" }));
                    content.Append(table);
                }
                else
                {
                    //Remove the previous table row
                    table.RemoveAllChildren<TableRow>();
                }
                //Add new table row with style
                foreach (JArray row in values)
                {
                    bool setRowHeight = false;
                    TableRow tableRow = new TableRow();

                    foreach (JValue cell in row)
                    {
                        var format = (JObject)formats.ElementAt(index);
                        if (!setRowHeight)
                        {
                            tableRow.Append(new TableRowProperties(new TableRowHeight() { Val = format["preferredHeight"] != null ? (format["preferredHeight"].Value<UInt32>() * 20) : 320 }));
                            setRowHeight = true;
                        }
                        string text = cell.Value.ToString();
                        string columnWidth = ((format["columnWidth"].Value<int>()) * 20).ToString();
                        var ha = format["horizontalAlignment"].Value<string>();
                        var va = format["verticalAlignment"].Value<string>();
                        var horizontalAlignment = ha == "Right" ? JustificationValues.Right : ha == "Centered" ? JustificationValues.Center : JustificationValues.Left;
                        var verticalAlignment = va == "Top" ? TableVerticalAlignmentValues.Top : va == "Center" ? TableVerticalAlignmentValues.Center : TableVerticalAlignmentValues.Bottom;
                        string shadingColor = format["shadingColor"].Value<string>();
                        var font = format["font"].Value<JObject>();
                        bool fontBold = font["bold"].Value<bool>();
                        string fontColor = font["color"].Value<string>();
                        bool fontItalic = font["italic"].Value<bool>();
                        string fontName = font["name"].Value<string>();
                        string fontSize = ((font["size"].Value<int>()) * 2).ToString();
                        string fu = font["underline"].Value<string>();
                        var fontUnderline = fu == "None" ? UnderlineValues.None : fu == "Single" ? UnderlineValues.Single : UnderlineValues.Double;

                        var border = format["border"].Value<JObject>();
                        var borderTop = border["top"].Value<JObject>();
                        string borderTopColor = (border["top"].Value<JObject>())["color"].Value<string>();
                        var borderTopValue = GetBorderValue((border["top"].Value<JObject>())["type"].Value<string>());
                        string borderBottomColor = (border["bottom"].Value<JObject>())["color"].Value<string>();
                        var borderBottomValue = GetBorderValue((border["bottom"].Value<JObject>())["type"].Value<string>());
                        string borderLeftColor = (border["left"].Value<JObject>())["color"].Value<string>();
                        var borderLeftValue = GetBorderValue((border["left"].Value<JObject>())["type"].Value<string>());
                        string borderRightColor = (border["right"].Value<JObject>())["color"].Value<string>();
                        var borderRightValue = GetBorderValue((border["right"].Value<JObject>())["type"].Value<string>());

                        TableCell tableCell = new TableCell();
                        tableCell.Append(new TableCellProperties(
                            new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = columnWidth },
                            new TableCellBorders(
                                new TopBorder() { Val = borderTopValue, Color = borderTopColor },
                                new BottomBorder() { Val = borderBottomValue, Color = borderBottomColor },
                                new LeftBorder() { Val = borderLeftValue, Color = borderLeftColor },
                                new RightBorder() { Val = borderRightValue, Color = borderRightColor }),
                            new Shading() { Val = ShadingPatternValues.Clear, Color = shadingColor, Fill = shadingColor },
                            new TableCellVerticalAlignment() { Val = verticalAlignment })
                            );

                        Paragraph tableParagraph = new Paragraph(new ParagraphProperties(new Justification() { Val = horizontalAlignment }));
                        Run tableRun = new Run(new RunProperties(
                            new RunFonts() { Ascii = fontName },
                            new Bold() { Val = fontBold },
                            new Italic() { Val = fontItalic },
                            new Color() { Val = fontColor },
                            new FontSize() { Val = fontSize },
                            new Underline() { Val = fontUnderline }));
                        Text tableText = new Text(text);
                        tableRun.Append(tableText);
                        tableParagraph.Append(tableRun);
                        tableCell.Append(tableParagraph);
                        tableRow.Append(tableCell);
                        index++;
                    }
                    table.Append(tableRow);
                }
            }
            return retValue;
        }


        static private async Task<Stream> GetFileStream(string authHeader, string destinationFileName, FileContextInfo fileContextInfo)
        {
            var filePath = new Uri(destinationFileName).AbsolutePath;
            var stream = new System.IO.MemoryStream();
            try
            {
                using (var client = new WebClient())
                {
                    client.Headers.Add(HttpRequestHeader.Authorization, authHeader);
                    var fileUri = new Uri(destinationFileName);
                    var downloadData = await client.DownloadDataTaskAsync($"{fileContextInfo.WebUrl}/_api/web/getfilebyserverrelativeurl('{filePath}')/$value");
                    await stream.WriteAsync(downloadData, 0, downloadData.Length);
                    return stream;
                }
            }
            catch (Exception ex)
            {
                stream.Dispose();
                throw ex;
            }
        }
        static private async Task<FileContextInfo> GetFileContextInfo(string authHeader, string destinationFileName)
        {
            var fileFolder = destinationFileName.Substring(0, destinationFileName.LastIndexOf('/'));
            using (var client = new WebClient())
            {
                client.Headers.Add(HttpRequestHeader.Authorization, authHeader);
                client.Headers.Add(HttpRequestHeader.Accept, "application/json");

                var fileInfo = await client.UploadStringTaskAsync($"{fileFolder}/_api/contextinfo", "POST", "");

                var jFileInfo = JObject.Parse(fileInfo);
                return new FileContextInfo()
                {
                    FormDigest = jFileInfo["FormDigestValue"].Value<string>(),
                    WebUrl = jFileInfo["WebFullUrl"].Value<string>()
                };
            }
        }

        private string GetAuthorizationHeader()
        {
            string authorizationHeader = LocalCache["AuthorizationHeader"] as string;
            lock (typeof(DocumentService))
            {

                if (string.IsNullOrWhiteSpace(authorizationHeader))
                {
                    AuthenticationContext authenticationContext = new AuthenticationContext(Authority, false);
                    var cert = GetCertificate();

                    //Try to use http://stackoverflow.com/questions/6392268/x509certificate-keyset-does-not-exist
                    using (cert.GetRSAPrivateKey()) { }
                    //switchest are important to work in webjob
                    ClientAssertionCertificate cac = new ClientAssertionCertificate(ClientID, cert);

                    var authenticationResult = authenticationContext.AcquireTokenAsync(Resource, cac).Result;
                    authorizationHeader = authenticationResult.CreateAuthorizationHeader();
                    LocalCache.Set("AuthorizationHeader", authorizationHeader, authenticationResult.ExpiresOn.AddMinutes(-2));
                }
            }
            return authorizationHeader;
        }
        private X509Certificate2 GetCertificate()
        {
            //read the certificate private key from the executing location
            //NOTE: This is a hack…Azure Key Vault is best approach
            var certPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            certPath = certPath.Substring(0, certPath.LastIndexOf('\\')) + "\\cert\\" + _configService.CertificateFile;
            var certfile = System.IO.File.OpenRead(certPath);
            var certificateBytes = new byte[certfile.Length];
            certfile.Read(certificateBytes, 0, (int)certfile.Length);
            return new X509Certificate2(
                certificateBytes,
                CertificatedPassword,
                X509KeyStorageFlags.Exportable |
                X509KeyStorageFlags.MachineKeySet |
                X509KeyStorageFlags.PersistKeySet);
        }

        static private string GetFormattedValue(string value, List<CustomFormat> formats, int? decimalPlace)
        {
            var originalValue = value;
            var hasDollar = originalValue.IndexOf("$") > -1;
            var hasComma = originalValue.IndexOf(",") > -1;
            var hasPercent = originalValue.IndexOf("%") > -1;
            var hasParenthesis = originalValue.IndexOf("(") > -1 && originalValue.IndexOf(")") > -1;
            var valueWithOutSymbol = originalValue.Replace("$", "").Replace(",", "").Replace("%", "").Replace("(", "").Replace(")", "");
            var isNumeric = Regex.IsMatch(valueWithOutSymbol, @"^[+-]?\d*[.]?\d*$");
            var isDate = IsDate(originalValue);

            foreach (var format in formats)
            {
                if (!format.IsDeleted)
                {
                    if (format.Name == "ConvertToHundreds")
                    {
                        if (isNumeric)
                        {
                            value = (Convert.ToDecimal(valueWithOutSymbol) / 100).ToString();
                            if (hasComma)
                            {
                                value = AddComma(value);
                            }
                            if (hasPercent)
                            {
                                value = value + "%";
                            }
                            if (hasParenthesis)
                            {
                                value = "(" + value + ")";
                            }
                            if (hasDollar)
                            {
                                value = "$" + value;
                            }
                        }
                    }
                    else if (format.Name == "ConvertToThousands")
                    {
                        if (isNumeric)
                        {
                            value = (Convert.ToDecimal(valueWithOutSymbol) / 1000).ToString();
                            if (hasComma)
                            {
                                value = AddComma(value);
                            }
                            if (hasPercent)
                            {
                                value = value + "%";
                            }
                            if (hasParenthesis)
                            {
                                value = "(" + value + ")";
                            }
                            if (hasDollar)
                            {
                                value = "$" + value;
                            }
                        }
                    }
                    else if (format.Name == "ConvertToMillions")
                    {
                        if (isNumeric)
                        {
                            value = (Convert.ToDecimal(valueWithOutSymbol) / 1000000).ToString();
                            if (hasComma)
                            {
                                value = AddComma(value);
                            }
                            if (hasPercent)
                            {
                                value = value + "%";
                            }
                            if (hasParenthesis)
                            {
                                value = "(" + value + ")";
                            }
                            if (hasDollar)
                            {
                                value = "$" + value;
                            }
                        }
                    }
                    else if (format.Name == "ConvertToBillions")
                    {
                        if (isNumeric)
                        {
                            value = (Convert.ToDecimal(valueWithOutSymbol) / 1000000000).ToString();
                            if (hasComma)
                            {
                                value = AddComma(value);
                            }
                            if (hasPercent)
                            {
                                value = value + "%";
                            }
                            if (hasParenthesis)
                            {
                                value = "(" + value + ")";
                            }
                            if (hasDollar)
                            {
                                value = "$" + value;
                            }
                        }
                    }
                    else if (format.Name == "ShowNegativesAsPositives")
                    {
                        var h = value.IndexOf("$") > -1;
                        var p = value.IndexOf("%") > -1;
                        var k = originalValue.IndexOf("(") > -1 && originalValue.IndexOf(")") > -1;
                        var hasHundred = value.IndexOf("hundred") > -1;
                        var hasThousand = value.IndexOf("thousand") > -1;
                        var hasMillion = value.IndexOf("million") > -1;
                        var hasBillion = value.IndexOf("billion") > -1;
                        value = value.Replace("$", "").Replace("-", "").Replace("%", "").Replace("(", "").Replace(")", "").Replace("hundred", "").Replace("thousand", "").Replace("million", "").Replace("billion", "").Trim();
                        if (p)
                        {
                            value = value + "%";
                        }
                        if (h)
                        {
                            value = "$" + value;
                        }
                        if (hasHundred)
                        {
                            value = value + " hundred";
                        }
                        else if (hasThousand)
                        {
                            value = value + " thousand";
                        }
                        else if (hasMillion)
                        {
                            value = value + " million";
                        }
                        else if (hasBillion)
                        {
                            value = value + " billion";
                        }
                    }
                    else if (format.Name == "IncludeHundredDescriptor")
                    {
                        if (isNumeric)
                        {
                            value = value + " hundred";
                        }
                    }
                    else if (format.Name == "IncludeThousandDescriptor")
                    {
                        if (isNumeric)
                        {
                            value = value + " thousand";
                        }
                    }
                    else if (format.Name == "IncludeMillionDescriptor")
                    {
                        if (isNumeric)
                        {
                            value = value + " million";
                        }
                    }
                    else if (format.Name == "IncludeBillionDescriptor")
                    {
                        if (isNumeric)
                        {
                            value = value + " billion";
                        }
                    }
                    else if (format.Name == "IncludeDollarSymbol")
                    {
                        if (value.IndexOf("$") == -1)
                        {
                            value = "$" + value;
                        }
                    }
                    else if (format.Name == "ExcludeDollarSymbol")
                    {
                        if (value.IndexOf("$") > -1)
                        {
                            value = value.Replace("$", "");
                        }
                    }
                    else if (format.Name == "DateShowLongDateFormat")
                    {
                        if (isDate)
                        {
                            var date = Convert.ToDateTime(value);
                            value = GetMonth(date.Month - 1) + " " + date.Day + ", " + date.Year;
                        }
                    }
                    else if (format.Name == "DateShowYearOnly")
                    {
                        if (isDate)
                        {
                            var date = Convert.ToDateTime(value);
                            value = date.Year.ToString();
                        }
                    }
                    else if (format.Name == "ConvertNegativeSymbolToParenthesis")
                    {
                        var h = value.IndexOf("$") > -1;
                        var pt = value.IndexOf("%") > -1;
                        var hasHundred = value.IndexOf("hundred") > -1;
                        var hasThousand = value.IndexOf("thousand") > -1;
                        var hasMillion = value.IndexOf("million") > -1;
                        var hasBillion = value.IndexOf("billion") > -1;
                        if (value.IndexOf("-") > -1)
                        {
                            value = value.Replace("$", "").Replace("-", "").Replace("%", "").Replace("(", "").Replace(")", "").Replace("hundred", "").Replace("thousand", "").Replace("million", "").Replace("billion", "").Trim();
                            if (h)
                            {
                                value = "$" + value;
                            }
                            value = "(" + value + ")";
                            if (pt)
                            {
                                value = value + "%";
                            }
                            if (hasHundred)
                            {
                                value = value + " hundred";
                            }
                            else if (hasThousand)
                            {
                                value = value + " thousand";
                            }
                            else if (hasMillion)
                            {
                                value = value + " million";
                            }
                            else if (hasBillion)
                            {
                                value = value + " billion";
                            }
                        }
                    }
                }
            }

            if (decimalPlace.HasValue)
            {
                var _hasComma = value.IndexOf(",") > -1;
                var _value = value.Replace("$", "").Replace(",", "").Replace("-", "").Replace("%", "").Replace("(", "").Replace(")", "").Replace("hundred", "").Replace("thousand", "").Replace("million", "").Replace("billion", "").Trim();
                if (Regex.IsMatch(_value, @"^[+-]?\d*[.]?\d*$"))
                {
                    var _format = "#0";
                    for (int i = 0; i < decimalPlace.Value; i++)
                    {
                        if (i == 0)
                        {
                            _format += ".";
                        }
                        _format += "0";
                    }
                    var _fromNumber = _hasComma ? AddComma(_value) : _value;
                    var _newNumber = Convert.ToDecimal(_value).ToString(_format);
                    var _toNumber = _hasComma ? AddComma(_newNumber) : _newNumber;
                    value = value.Replace(_fromNumber, _toNumber);
                }
            }
            return value;
        }

        static private bool IsDate(string value)
        {
            var flag = false;
            try
            {
                var date = Convert.ToDateTime(value);
                flag = date.Year > 0;
            }
            catch
            {
                flag = false;
            }
            return flag;
        }

        static private string AddComma(string value)
        {
            var decimalLength = value.Split('.').Length > 1 ? value.Split('.')[1].Length : 0;
            return Convert.ToDecimal(value).ToString("N" + decimalLength);
        }

        static private string GetMonth(int month)
        {
            var _m = new string[12] { "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December" };
            return _m[month];
        }

        static private BorderValues GetBorderValue(string borderValue)
        {
            BorderValues value = BorderValues.None;
            switch (borderValue)
            {
                case "Dotted":
                    value = BorderValues.Dotted; break;
                case "Dashed":
                    value = BorderValues.Dashed; break;
                case "Dot2Dashed":
                    value = BorderValues.DotDotDash; break;
                case "DotDashed":
                    value = BorderValues.DotDash; break;
                case "DashedSmall":
                    value = BorderValues.DashSmallGap; break;
                case "Single":
                    value = BorderValues.Single; break;
                case "DashDotStroked":
                    value = BorderValues.DashDotStroked; break;
                case "ThreeDEmboss":
                    value = BorderValues.ThreeDEmboss; break;
                case "Double":
                    value = BorderValues.Double; break;
                default:
                    value = BorderValues.None;
                    break;
            }
            return value;
        }

        #region Check document URL By ID

        public async Task<DocumentCheckResult> GetDocumentUrlByID(DocumentCheckResult result)
        {
            string url = string.Empty;
            try
            {
                string authHeader = GetAuthorizationHeader();
                using (var client = new WebClient())
                {
                    client.Headers.Add(HttpRequestHeader.Authorization, authHeader);
                    client.Headers.Add(HttpRequestHeader.Accept, "application/json");
                    var downloadString = await client.DownloadStringTaskAsync(new Uri(Resource + "/_api/search/query?querytext='DlcDocId:" + result.DocumentId + "'&t=" + (new Random()).Next()));
                    var json = JObject.Parse(downloadString);

                    foreach (JObject jO in (JArray)json["PrimaryQueryResult"]["RelevantResults"]["Table"]["Rows"])
                    {
                        foreach (JObject jA in (JArray)jO["Cells"])
                        {
                            if (jA["Key"].Value<string>() == "Path")
                            {
                                url = jA["Value"].Value<string>();
                                break;
                            }
                        }
                        if (!string.IsNullOrEmpty(url))
                        {
                            break;
                        }
                    }
                    result.IsSuccess = true;
                    if (!string.IsNullOrEmpty(url))
                    {
                        result.DocumentUrl = url;
                        result.IsDeleted = false;
                    }
                    else
                    {
                        result.IsDeleted = true;
                    }
                }
            }
            catch (Exception ex)
            {
                result.IsSuccess = false;
                result.Message = ex.Message;
            }
            return result;
        }

        #endregion
    }
}
