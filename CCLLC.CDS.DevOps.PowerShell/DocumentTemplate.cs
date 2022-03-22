
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.Xrm.Sdk;

namespace CCLLC.Cds.DevOps.PowerShell
{
    internal enum DocumentTypes { MicrosoftExcel = 1, MicrosoftWord = 2 }

    internal class DocumentTemplate
    {
        public Guid? Id { get; private set; }
        public string Name { get; private set; }
        public DocumentTypes? DocumentType { get; private set; }
        public byte[] FileContents { get; private set; }
        public string EntityTypeName { get; private set; }
        public int ObjectTypeCode { get; private set; }
        
        public DocumentTemplate(string filePath)
        {
            LoadFile(filePath);            
        }

        public DocumentTemplate(Entity record)
        {
            if(record.LogicalName != Constants.EntityLogicalName)
            {
                throw new Exception("Incorrect record type.");
            }

            Id = record.Id;
            DocumentType = (DocumentTypes)record.GetAttributeValue<OptionSetValue>(Constants.Fields.DocumentType)?.Value;
            EntityTypeName = record.GetAttributeValue<string>(Constants.Fields.AssociatedEntityTypeCode);
                      
            Name = record.GetAttributeValue<string>(Constants.Fields.Name);
            FileContents = Convert.FromBase64String(record.GetAttributeValue<string>(Constants.Fields.Content));

            if(DocumentType == DocumentTypes.MicrosoftWord 
                && FileContents != null)
            {
                var entityIdentifiers = ExtractEntityTypeAndCodeFromDocument();
                ObjectTypeCode = entityIdentifiers?.Item2 ?? 0;
            }
        }

        public void LoadFile(string filePath)
        {
            var extension = Path.GetExtension(filePath);
            switch (extension)
            {
                case Constants.FileExtensions.Word:
                    DocumentType = DocumentTypes.MicrosoftWord;
                    break;
                case Constants.FileExtensions.Excel:
                    DocumentType = DocumentTypes.MicrosoftExcel;
                    break;
                default:
                    throw new Exception("Unsupported document extension.");
            }


            Name = Path.GetFileNameWithoutExtension(filePath);

            FileContents = File.ReadAllBytes(filePath);

            if (DocumentType == DocumentTypes.MicrosoftWord
                && FileContents != null)
            {
                var entityIdentifiers = ExtractEntityTypeAndCodeFromDocument();

                EntityTypeName = entityIdentifiers?.Item1;
                ObjectTypeCode = entityIdentifiers?.Item2 ?? 0;

            }
        }

        public void SaveFile(string filePath)
        {
            var fileExtension = DocumentType == DocumentTypes.MicrosoftExcel ? Constants.FileExtensions.Excel : Constants.FileExtensions.Word;
            var filename = $"{Name}{fileExtension}";

            filename = Path.Combine(filePath, filename);

            using (var stream = new FileStream(filename, FileMode.OpenOrCreate, FileAccess.Write))
            {
                stream.Write(FileContents, 0, FileContents.Length);
                stream.Close();
            }
        }
               
        public Entity ToEntity()
        {
            var entity = new Entity("documenttemplate");

            if (Id != null)
            {
                entity.Id = Id ?? default(Guid);
                entity[Constants.Fields.Id] = Id;
            }

            entity[Constants.Fields.Name] = Name;
            entity[Constants.Fields.DocumentType] = new OptionSetValue((int)DocumentType);
            entity[Constants.Fields.AssociatedEntityTypeCode] = EntityTypeName;
            entity[Constants.Fields.Content] = Convert.ToBase64String(FileContents);

            return entity;

        }

        public void UpdateDocumentObjectTypeCodes(int newObjectTypeCode)
        {
            if (DocumentType != DocumentTypes.MicrosoftWord)
            {
                return;
            }

            using (var stream = new MemoryStream(FileContents))
            {
                var document = WordprocessingDocument.Open(stream, true);

                var mainPart = document.MainDocumentPart;

                var headerParts = mainPart?
                    .Parts
                    .Where(p => p.OpenXmlPart is HeaderPart)
                    .Select(p => p.OpenXmlPart);

                var footerParts = mainPart?
                    .Parts
                    .Where(p => p.OpenXmlPart is FooterPart)
                    .Select(p => p.OpenXmlPart);

                var customXmlPropertyParts = mainPart?
                    .Parts
                    .Where(p => p.OpenXmlPart is CustomXmlPart)
                    .Select(c => (c.OpenXmlPart as CustomXmlPart).CustomXmlPropertiesPart);

                UpdateDocumentPartObjectTypeCode(mainPart, newObjectTypeCode);
                UpdateDocumentPartObjectTypeCode(headerParts, newObjectTypeCode);
                UpdateDocumentPartObjectTypeCode(footerParts, newObjectTypeCode);
                UpdateDocumentPartObjectTypeCode(customXmlPropertyParts, newObjectTypeCode);

                document.Close();

                FileContents = stream.ToArray();
                ObjectTypeCode = newObjectTypeCode;
            }
        }

        private Tuple<string, int> ExtractEntityTypeAndCodeFromDocument()
        {
            const string rootPath = @"urn:microsoft-crm/document-template/";
            const string schemaNameSpace = @"{http://schemas.openxmlformats.org/officeDocument/2006/customXml}";

            using (var stream = new MemoryStream(FileContents))
            {
                var document = WordprocessingDocument.Open(stream, false);
                
                var customXmlPropertyParts = document
                    .MainDocumentPart?
                    .Parts
                    .Where(p => p.OpenXmlPart is CustomXmlPart)
                    .Select(c => (c.OpenXmlPart as CustomXmlPart).CustomXmlPropertiesPart);

                var dataMappingXml = customXmlPropertyParts?
                    .Where(p => p.RootElement?.InnerXml?.Contains(rootPath) ?? false)
                    .FirstOrDefault()?.RootElement.InnerXml;
                    
                if (dataMappingXml is null)
                {
                    return null;
                }                               
                
                using (var textStream = new StringReader(dataMappingXml))
                {
                    var xDoc = XDocument.Load(textStream);
                    var xAttribute = xDoc
                        .Descendants($"{schemaNameSpace}schemaRef")
                        .Attributes($"{schemaNameSpace}uri")
                        .Where<XAttribute>(a => a.Value.StartsWith(rootPath))
                        .FirstOrDefault();

                    var entityIdentifiers = xAttribute
                        .Value
                        .Substring(rootPath.Length)
                        .Split('/');

                    return new Tuple<string, int>(entityIdentifiers[0], int.Parse(entityIdentifiers[1]));                   
                }
            }
        }
        
        private void UpdateDocumentPartObjectTypeCode(IEnumerable<OpenXmlPart> parts, int newObjectTypeCode)
        {
            if (parts is null)
            {
                return;
            }

            foreach (var part in parts)
            {
                UpdateDocumentPartObjectTypeCode(part, newObjectTypeCode);
            }
        }

        private void UpdateDocumentPartObjectTypeCode(OpenXmlPart part, int newObjectTypeCode)
        {
            if (part is null)
            {
                return;
            }

            string documentText;
            using (var reader = new StreamReader(part.GetStream()))
            {
                documentText = reader.ReadToEnd();
            }

            documentText = documentText.Replace($"{EntityTypeName}/{ObjectTypeCode}/", $"{EntityTypeName}/{newObjectTypeCode}/");

            using (var writer = new StreamWriter(part.GetStream()))
            {
                writer.Write(documentText);
                writer.Flush();
            }
        }
    }
}
