cd C:\repos\IRIS\iris_compliance\DocumentTemplateManager\bin\Debug;
Import-Module .\DocumentTemplateManager.dll -Force;
#$Conn = Get-CrmConnection -Interactive;

#Add-DocumentTemplates -Conn $Conn -TemplateDirectory  C:\repos\IRIS\iris_compliance\DeploymentPackage\CorrespondenceTemplates -Verbose

#$firstPackage = Get-CrmDataPackage -Conn $Conn -Fetches @("<fetch><entity name='account'><all-attributes/></entity></fetch>") -Verbose;
#Get-CrmDataPackage -Conn $Conn -Fetches @("<fetch><entity name='contact'><all-attributes/></entity></fetch>", "<fetch><entity name='category'><all-attributes/></entity></fetch>") -Identifiers @{ "contact" = @("firstname", "lastname"); "category" = @("categoryid") } -DisablePlugins @{ "contact" = $true } -DisablePluginsGlobally $true -Verbose `
#    | Add-FetchesToCrmDataPackage -Conn $Conn -Fetches @("<fetch><entity name='knowledgearticle'><all-attributes/></entity></fetch>") -Verbose `
#    | Merge-CrmDataPackage -AdditionalPackage $firstPackage -Verbose `
#	| Remove-CrmDataPackage -RemovePackage $firstPackage -Verbose `
#    | Export-CrmDataPackage -ZipPath $env:USERPROFILE\Downloads\testrun.zip -Verbose