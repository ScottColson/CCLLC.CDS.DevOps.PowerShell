using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management.Automation;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;


namespace CCLLC.Cds.DevOps.PowerShell.Cmdlets
{
    [Cmdlet(VerbsData.Import, "DocumentTemplates")]
    public class ImportDocumentTemplates : PSCmdlet
    {
        private static Dictionary<string, int> _typeCodes = new Dictionary<string, int>();

        [Parameter(Position = 0, Mandatory = true)]
        public IOrganizationService Conn { get; set; }

        [Parameter(Position = 1, Mandatory = true)]
        public string TemplateDirectory { get; set; }

        

        protected override void ProcessRecord()
        {
            var existingTemplates = LoadExistingTemplatesFromDynamics(Conn);

            WriteVerbose($"Found {existingTemplates.Count} existing templates in target environment.");

            WriteVerbose($"Loading templates from: {TemplateDirectory}");

            var filePaths = Directory.GetFiles(TemplateDirectory, $"*{Constants.FileExtensions.Word}");
            WriteVerbose($"Found {filePaths.Length} Word Templates in directory.");
            foreach (var path in filePaths)
            {
                ProcessFile(Conn, path, existingTemplates);
            }

            filePaths = Directory.GetFiles(TemplateDirectory, $"*{Constants.FileExtensions.Excel}");
            WriteVerbose($"Found {filePaths.Length} Excel Templates in directory.");
            foreach (var path in filePaths)
            {
                ProcessFile(Conn, path, existingTemplates);
            }
        }

        private void ProcessFile(IOrganizationService service, string path, IList<Entity> existingTemplates)
        {
            var name = Path.GetFileNameWithoutExtension(path);

            var existingTemplate = existingTemplates.Where(t => t.GetAttributeValue<string>(Constants.Fields.Name) == name).FirstOrDefault();

            if (existingTemplate is null)
            {
                WriteVerbose($"Adding {path} as a new template.");
                var template = new DocumentTemplate(path);

                if (template.DocumentType == DocumentTypes.MicrosoftWord)
                {
                    var targetTypeCode = GetTypeCodeFromDynamics(Conn, template.EntityTypeName);

                    if (targetTypeCode != template.ObjectTypeCode)
                    {
                        WriteVerbose($"Updating Document ObjectTypeCode to {targetTypeCode}.");
                        template.UpdateDocumentObjectTypeCodes(targetTypeCode);
                    }
                }

                var entity = template.ToEntity();
                Conn.Create(entity);
            }
            else
            {
                WriteVerbose($"Updating existing template with {path}.");
                var template = new DocumentTemplate(existingTemplate);
                template.LoadFile(path);

                if (template.DocumentType == DocumentTypes.MicrosoftWord)
                {
                    var targetTypeCode = GetTypeCodeFromDynamics(Conn, template.EntityTypeName);

                    if (targetTypeCode != template.ObjectTypeCode)
                    {
                        WriteVerbose($"Updating ObjectTypeCode to {targetTypeCode}.");
                        template.UpdateDocumentObjectTypeCodes(targetTypeCode);
                    }
                }

                var entity = template.ToEntity();
                Conn.Update(entity);
            }
        }

        private static IList<Entity> LoadExistingTemplatesFromDynamics(IOrganizationService service)
        {
            var qry = new QueryExpression
            {
                EntityName = Constants.EntityLogicalName,
                ColumnSet = new ColumnSet(Constants.Fields.Id, Constants.Fields.AssociatedEntityTypeCode, Constants.Fields.Content, Constants.Fields.DocumentType, Constants.Fields.Name),
            };

            var records = service.RetrieveMultiple(qry).Entities.ToList();
            return records;
        }

        private static int GetTypeCodeFromDynamics(IOrganizationService service, string entityName)
        {
            if (!_typeCodes.ContainsKey(entityName))
            {
                RetrieveEntityRequest request = new RetrieveEntityRequest();
                request.LogicalName = entityName;
                request.EntityFilters = Microsoft.Xrm.Sdk.Metadata.EntityFilters.Entity;

                RetrieveEntityResponse response = (RetrieveEntityResponse)service.Execute(request);
                _typeCodes.Add(entityName, response.EntityMetadata.ObjectTypeCode.Value);

            }

            return _typeCodes[entityName];
        }
    }
}