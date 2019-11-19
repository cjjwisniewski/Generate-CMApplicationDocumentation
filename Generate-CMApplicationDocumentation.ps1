<#--------------------------
Name: Generate-CMApplicationDocumentation
Author: Cameron Wisniewski
Date: 1/2/19
Version: 1.22719
Comment: Parses information necessary to create SCCM Application documentation
Notes: Can't figure out how to query AppTaskSequenceDeployment,AppPersistClientCache yet
Link: N/A
--------------------------#>

[CmdletBinding()]
Param(
    [Parameter(Mandatory=$True)]
    [String[]]$ApplicationList,
    [Parameter(Mandatory=$True)]
    [String]$Destination,
    [String]$PathToTemplate = "$PSScriptRoot\Files\APP - _GenericTemplate.xlsx"
    )

#Set variables
$ErrorActionPreference = "SilentlyContinue"
$SiteCode = "CCI"
$SiteServer = "CCICUSSCCM1.us.crowncastle.com"
$CurrentPath = (Get-Location).Path

#Import SCCM cmdlet
Import-Module "${ENV:ProgramFiles(x86)}\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1"
Set-Location "$SiteCode`:"
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process

#Declare data table
$ApplicationData = New-Object System.Data.DataTable
$ApplicationData.Columns.Add("AppName","string") | Out-Null
$ApplicationData.Columns.Add("AppCIUniqueID","string") | Out-Null
$ApplicationData.Columns.Add("AppCreator","string") | Out-Null
$ApplicationData.Columns.Add("AppPublisher","string") | Out-Null
$ApplicationData.Columns.Add("AppSoftwareVersion","string") | Out-Null
$ApplicationData.Columns.Add("AppDatePublished","string") | Out-Null
$ApplicationData.Columns.Add("AppDateModified","string") | Out-Null
$ApplicationData.Columns.Add("AppComments","string") | Out-Null
$ApplicationData.Columns.Add("AppTaskSequenceDeployment","string") | Out-Null
$ApplicationData.Columns.Add("AppLocalizedDisplayName","string") | Out-Null
$ApplicationData.Columns.Add("AppFeatured","string") | Out-Null
$ApplicationData.Columns.Add("AppDeploymentTypeName","string") | Out-Null
$ApplicationData.Columns.Add("AppTechnology","string") | Out-Null
$ApplicationData.Columns.Add("AppContentLocation","string") | Out-Null
$ApplicationData.Columns.Add("AppPersistClientCache","string") | Out-Null
$ApplicationData.Columns.Add("AppPeerCache","string") | Out-Null
$ApplicationData.Columns.Add("AppFallbackDP","string") | Out-Null
$ApplicationData.Columns.Add("AppFastNetworkDownload","string") | Out-Null
$ApplicationData.Columns.Add("AppInstallationProgram","string") | Out-Null
$ApplicationData.Columns.Add("AppUninstallProgram","string") | Out-Null
$ApplicationData.Columns.Add("AppRunAs32Bit","string") | Out-Null
$ApplicationData.Columns.Add("AppProductCode","string") | Out-Null
$ApplicationData.Columns.Add("AppDetection","string") | Out-Null
$ApplicationData.Columns.Add("AppExecutionContext","string") | Out-Null
$ApplicationData.Columns.Add("AppLogonRequirement","string") | Out-Null
$ApplicationData.Columns.Add("AppInteractionMode","string") | Out-Null
$ApplicationData.Columns.Add("AppAllowInteraction","string") | Out-Null
$ApplicationData.Columns.Add("AppMaxExecuteTime","string") | Out-Null
$ApplicationData.Columns.Add("AppExecuteTime","string") | Out-Null
$ApplicationData.Columns.Add("AppPostInstallBehavior","string") | Out-Null
$ApplicationData.Columns.Add("AppRequirements","string") | Out-Null
$ApplicationData.Columns.Add("AppDependencies","string") | Out-Null

#Trim destination to stop save errors if it ends in \
$Destination = $Destination.TrimEnd("\")

$ApplicationList | ForEach-Object {
    #Query application by ApplicationList
    $Application = Get-CMApplication -Name $_

    #Loop though $Application to account for multiple returns on named search
    $Application | ForEach-Object {
        Write-Host "LOG: Retrieving information for"$_.LocalizedDisplayName

        #Query application XML
        [xml]$ApplicationXML = (Get-CMApplication -Name $_.LocalizedDisplayName).SDMPackageXML

        #region AppDetection
        if($ApplicationXML.AppMgmtDigest.DeploymentType.Installer.CustomData.EnhancedDetectionMethod.Settings.SimpleSetting.RegistryDiscoverySource.HasAttribute("Hive")){
            #Read registry hive
            $RegistryHiveXML = $ApplicationXML.AppMgmtDigest.DeploymentType.Installer.CustomData.EnhancedDetectionMethod.Settings.SimpleSetting.RegistryDiscoverySource.Hive

            #Switch registry hive, add new providers where necessary
            switch($RegistryHiveXML) {
                "HKEY_CLASSES_ROOT" {
                    $RegistryHive = "HKCR"
                    New-PSDrive -PSProvider Registry -Name HKCR -Root HKEY_CLASSES_ROOT
                }
                "HKEY_CURRENT_CONFIG" {
                    $RegistryHive = "HKCC"
                    New-PSDrive -PSProvider Registry -Name HKCC -Root HKEY_CURRENT_CONFIG
                }
                "HKEY_CURRENT_USER" {
                    $RegistryHive = "HKCU"
                }
                "HKEY_LOCAL_MACHINE" {
                    $RegistryHive = "HKLM"
                }
                "HKEY_USERS" {
                    $RegistryHive = "HKU"
                    New-PSDrive -PSProvider Registry -Name HKU -Root HKEY_USERS
                }
            }
                        
            #Read registry path
            $RegistryKeyPath = $ApplicationXML.AppMgmtDigest.DeploymentType.Installer.CustomData.EnhancedDetectionMethod.Settings.SimpleSetting.RegistryDiscoverySource.Key
            $RegistryKeyPath = "$RegistryHive`:\$RegistryKeyPath"

            #Read registry key value name
            $RegistryKeyName = $ApplicationXML.AppMgmtDigest.DeploymentType.Installer.CustomData.EnhancedDetectionMethod.Settings.SimpleSetting.RegistryDiscoverySource.ValueName

            #Read registry key value
            $RegistryKeyValue = $ApplicationXML.AppMgmtDigest.DeploymentType.Installer.CustomData.EnhancedDetectionMethod.Rule.Expression.Operands.ConstantValue.Value

            #Validate application
            $AppDetection = "$RegistryKeyPath\$RegistryKeyName = $RegistryKeyValue"
        }
        
        #Handle product code validations
        if($ApplicationXML.AppMgmtDigest.DeploymentType.Installer.CustomData.EnhancedDetectionMethod.Settings.MSI.HasAttribute("LogicalName")){
            #Read product code
            $AppDetection = $ApplicationXML.AppMgmtDigest.DeploymentType.Installer.CustomData.EnhancedDetectionMethod.Settings.MSI.ProductCode
        }

        #Handle file detection validation. This generally should not be used, and when it is should be a simple detection rather than a version check or similar, so I'm not going to write that.
        if($ApplicationXML.AppMgmtDigest.DeploymentType.Installer.CustomData.EnhancedDetectionMethod.Settings.File.HasAttribute("LogicalName")){
            #Read file path
            $DetectPath = $ApplicationXML.AppMgmtDigest.DeploymentType.Installer.CustomData.EnhancedDetectionMethod.Settings.File.Path
            #Read detect item
            $DetectItem = $ApplicationXML.AppMgmtDigest.DeploymentType.Installer.CustomData.EnhancedDetectionMethod.Settings.File.Filter
            #Build filepath
            $AppDetection = "$DetectPath\$DetectItem"
        }

        #Handle folder detection validation
        if($ApplicationXML.AppMgmtDigest.DeploymentType.Installer.CustomData.EnhancedDetectionMethod.Settings.Folder.HasAttribute("LogicalName")){
            #Read folder path
            $DetectPath = $ApplicationXML.AppMgmtDigest.DeploymentType.Installer.CustomData.EnhancedDetectionMethod.Settings.Folder.Path
            #Read detect item
            $DetectItem = $ApplicationXML.AppMgmtDigest.DeploymentType.Installer.CustomData.EnhancedDetectionMethod.Settings.Folder.Filter
            #Build filepath
            $AppDetection = "$DetectPath\$DetectItem"
        }
        #endregion AppDetection

        #region AppDependencies
        $AppDependencyApplicationReference = $ApplicationXML.AppMgmtDigest.DeploymentType.Dependencies.DeploymentTypeRule.DeploymentTypeExpression.Operands.DeploymentTypeIntentExpression.DeploymentTypeApplicationReference.LogicalName
        if($AppDependencyApplicationReference) {
            $AppDependencyApplicationReference | ForEach-Object {
                $DependencyApplication = Get-WmiObject -Namespace "root\SMS\site_$SiteCode" -Class "SMS_ApplicationLatest" -ComputerName $SiteServer -Filter "CI_UniqueID like '%$_%'"
                $AppDependencies += $DependencyApplication.LocalizedDisplayName+", "
            }
        }
        $AppDependencies = $AppDependencies.Substring(0,($AppDependencies.Length-2))
        #endregion AppDependencies

        #Add data to table
        $NewApp = $ApplicationData.NewRow()
        $NewApp.AppName = $_.LocalizedDisplayName
        $NewApp.AppCIUniqueID = $_.CI_UniqueID
        $NewApp.AppCreator = $_.CreatedBy
        $NewApp.AppPublisher = $_.Manufacturer
        $NewApp.AppSoftwareVersion = $_.SoftwareVersion
        $NewApp.AppDatePublished = $_.DateCreated
        $NewApp.AppDateModified = $_.DateLastModified
        $NewApp.AppComments = $_.LocalizedDescription
        #$NewApp.AppTaskSequenceDeployment = 
        $NewApp.AppLocalizedDisplayName = $_.LocalizedDisplayName
        $NewApp.AppFeatured = $_.Featured
        $NewApp.AppDeploymentTypeName = $ApplicationXML.AppMgmtDigest.DeploymentType.Title.'#text'
        $NewApp.AppTechnology = $ApplicationXML.AppMgmtDigest.DeploymentType.Installer.Technology
        $NewApp.AppContentLocation = $ApplicationXML.AppMgmtDigest.DeploymentType.Installer.Contents.Content.Location
        #$NewApp.AppPersistClientCache =
        $NewApp.AppPeerCache = $ApplicationXML.AppMgmtDigest.DeploymentType.Installer.Contents.Content.PeerCache
        $NewApp.AppFallbackDP = $ApplicationXML.AppMgmtDigest.DeploymentType.Installer.Contents.Content.FallbackToUnprotectedDP
        $NewApp.AppFastNetworkDownload = $ApplicationXML.AppMgmtDigest.DeploymentType.Installer.Contents.Content.OnFastNetwork
        $NewApp.AppInstallationProgram = $ApplicationXML.AppMgmtDigest.DeploymentType.Installer.InstallAction.Args.Arg[0].'#text'
        $NewApp.AppUninstallProgram = $ApplicationXML.AppMgmtDigest.DeploymentType.Installer.CustomData.UninstallCommandLine
        $NewApp.AppRunAs32Bit = $ApplicationXML.AppMgmtDigest.DeploymentType.Installer.InstallAction.Args.Arg[11].'#text'
        $NewApp.AppProductCode =  $ApplicationXML.AppMgmtDigest.DeploymentType.Installer.CustomData.SourceUpdateProductCode
        $NewApp.AppDetection = $AppDetection
        $NewApp.AppExecutionContext = $ApplicationXML.AppMgmtDigest.DeploymentType.Installer.InstallAction.Args.Arg[2].'#text'
        $NewApp.AppLogonRequirement = $_.LogonRequirement
        $NewApp.AppInteractionMode = $ApplicationXML.AppMgmtDigest.DeploymentType.Installer.InstallAction.Args.Arg[7].'#text'
        $NewApp.AppAllowInteraction = $ApplicationXML.AppMgmtDigest.DeploymentType.Installer.InstallAction.Args.Arg[5].'#text'
        $NewApp.AppMaxExecuteTime = $ApplicationXML.AppMgmtDigest.DeploymentType.Installer.InstallAction.Args.Arg[10].'#text'
        $NewApp.AppExecuteTime = $ApplicationXML.AppMgmtDigest.DeploymentType.Installer.InstallAction.Args.Arg[9].'#text'
        $NewApp.AppPostInstallBehavior = $ApplicationXML.AppMgmtDigest.DeploymentType.Installer.InstallAction.Args.Arg[8].'#text'
        $NewApp.AppRequirements = $ApplicationXML.AppMgmtDigest.DeploymentType.Requirements.Rule.Annotation.DisplayName.Text
        $NewApp.AppDependencies =  $AppDependencies
        $ApplicationData.Rows.Add($NewApp)

        #Import Excel document
        Copy-Item -Path $PathToTemplate -Destination "$ENV:TEMP\temp.xlsx"
        $ExcelProcess = New-Object -ComObject Excel.Application
        $ExcelProcess.Visible = $false
        $ExcelProcess.DisplayAlerts = $false
        $ExcelWorkbooks = $ExcelProcess.Workbooks.Open("$ENV:TEMP\temp.xlsx")
        $ExcelWorksheet = $ExcelWorkbooks.Worksheets.Item(1)

        #Write out documentation data
        $ExcelWorksheet.Cells(4,6) = Get-Date -Format d
        $ExcelWorksheet.Cells(4,8) = "1.0"
        $ExcelWorksheet.Cells(4,9) = $ENV:USERNAME
        $ExcelWorksheet.Cells(4,10) = "Original document"

        #Write out application data:
        #Application Name
        $ExcelWorksheet.Cells(2,4) = $NewApp.AppName

        #CI Unique ID:	
        $ExcelWorksheet.Cells(3,4) = $NewApp.AppCIUniqueID

        #Created by:
        $ExcelWorksheet.Cells(4,4) = $NewApp.AppCreator

        #Date
        $ExcelWorksheet.Cells(5,4) = $NewApp.AppDatePublished

        #Updated
        $ExcelWorksheet.Cells(6,4) = $NewApp.AppDateModified

        #Purpose
        $ExcelWorksheet.Cells(8,4) = $NewApp.AppComments

        #Name
        $ExcelWorksheet.Cells(11,4) = $NewApp.AppName

        #Publisher
        $ExcelWorksheet.Cells(12,4) = $NewApp.AppPublisher

        #Software Version
        $ExcelWorksheet.Cells(13,4) = $NewApp.AppSoftwareVersion

        #Date Published
        $ExcelWorksheet.Cells(14,4) = $NewApp.AppDatePublished

        #Allow this application to be published from the Installl Application task sequence action without being deployed
        #Cannot for the life of me figure out how this property is stored
        $ExcelWorksheet.Cells(15,4) = "Checked"

        #Localized Application Name
        $ExcelWorksheet.Cells(14,9) = $NewApp.AppName

        #Display this as a featured app and highlight it in the company portal
        switch($NewApp.AppFeatured){
            0 {$ExcelWorksheet.Cells(15,9) = "Unchecked"}
            1 {$ExcelWorksheet.Cells(15,9) = "Checked"}
        }

        #Deployment Type Title
        $ExcelWorksheet.Cells(19,2) = $NewApp.AppDeploymentTypeName

        #Technology
        switch($NewApp.AppTechnology){
            "MSI" {$ExcelWorksheet.Cells(20,4) = "Windows Installer (*.msi file)"}
            default {$ExcelWorksheet.Cells(20,4) = $NewApp.AppTechnology}
        }

        #Content location
        $ExcelWorksheet.Cells(21,4) = $NewApp.AppContentLocation

        #Persist in the client cache
        #Cannot for the life of me figure out how this property is stored
        $ExcelWorksheet.Cells(22,4) = "Unchecked"

        #Allow clients to share content with other clients on the same subnet:
        switch($NewApp.AppPeerCache){
            $true {$ExcelWorksheet.Cells(23,4) = "Checked"}
            $false {$ExcelWorksheet.Cells(23,4) = "Unchecked"}
            $null {$ExcelWorksheet.Cells(23,4) = "Unchecked"}
            default {$ExcelWorksheet.Cells(23,4) = $NewApp.AppPeerCache}
        }

        #Allow clients to use a distribution points from the default site boundary group
        switch($NewApp.AppFallbackDP){
            $true {$ExcelWorksheet.Cells(24,4) = "Checked"}
            $false {$ExcelWorksheet.Cells(24,4) = "Unchecked"}
            $null {$ExcelWorksheet.Cells(24,4) = "Unchecked"}
            default {$ExcelWorksheet.Cells(24,4) = $NewApp.AppFallbackDP}
        }

        #Deployment options
        switch($NewApp.AppFastNetworkDownload){
            "Download" {$ExcelWorksheet.Cells(25,4) = "Download content from distribution point and run locally"}
            $null {$ExcelWorksheet.Cells(25,4) = $NewApp.AppFastNetworkDownload}
            default {$ExcelWorksheet.Cells(25,4) = $NewApp.AppFastNetworkDownload}
        }

        #Installation program
        $ExcelWorksheet.Cells(26,4) = $NewApp.AppInstallationProgram

        #Uninstall program
        $ExcelWorksheet.Cells(27,4) = $NewApp.AppUninstallProgram

        #Run installation and uninstall program as 32-bit process on 64-bit clients:
        switch($NewApp.AppRunAs32Bit){
            $true {$ExcelWorksheet.Cells(28,4) = "Checked"}
            $false {$ExcelWorksheet.Cells(28,4) = "Unchecked"}
            $null {$ExcelWorksheet.Cells(28,4) = "Unchecked"}
            default {$ExcelWorksheet.Cells(28,4) = $NewApp.AppRunAs32Bit}
        }

        #Product code
        $ExcelWorksheet.Cells(29,4) = $NewApp.AppProductCode

        #Detection Method
        $ExcelWorksheet.Cells(30,4) = $NewApp.AppDetection

        #Installation behavior:
        switch($NewApp.AppExecutionContext){
            "System" {$ExcelWorksheet.Cells(34,4) = "Install for system"}
            "User" {$ExcelWorksheet.Cells(34,4) = "Install for user"}
            "Any" {$ExcelWorksheet.Cells(34,4) = "Install for system if resource is device; otherwise install for user"}
            default {$ExcelWorksheet.Cells(34,4) =  $NewApp.AppExecutionContext}
        }

        #Logon requirement
        switch($NewApp.AppLogonRequirement){
            0 {$ExcelWorksheet.Cells(35,4) = "Whether or not a user is logged on"}
            1 {$ExcelWorksheet.Cells(35,4) = "Only when a user is logged on"}
            2 {$ExcelWorksheet.Cells(35,4) = "Only when no user is logged on"}
            default {$ExcelWorksheet.Cells(35,4) = $NewApp.AppLogonRequirement}
        }

        #Installation program visibility
        $ExcelWorksheet.Cells(36,4) = $NewApp.AppInteractionMode

        #Allow users to interact with the program installation
        switch($NewApp.AppAllowInteraction){
            $true {$ExcelWorksheet.Cells(37,4) = "Checked"}
            $false {$ExcelWorksheet.Cells(37,4) = "Unchecked"}
            $null {$ExcelWorksheet.Cells(37,4) = "Unchecked"}
            default {$ExcelWorksheet.Cells(37,4) = $NewApp.AppAllowInteraction}
        }

        #Maximum allowed run time (minutes)
        $ExcelWorksheet.Cells(38,4) = $NewApp.AppMaxExecuteTime

        #Estimated installation time (minutes)
        $ExcelWorksheet.Cells(39,4) = $NewApp.AppExecuteTime

        #Should Cfg Mgr enforce specific behavior regardless of the applications's intended behavior?
        switch($NewApp.AppPostInstallBehavior){
            "NoAction" {$ExcelWorksheet.Cells(40,4) = "No specific action"}
            "BasedOnExitCode" {$ExcelWorksheet.Cells(40,4) = "Determine behavior based on return codes"}
            default {$ExcelWorksheet.Cells(40,4) = $NewApp.AppPostInstallBehavior}
        }

        #Requirements
        $ExcelWorksheet.Cells(41,4) = $NewApp.AppRequirements

        #Dependencies
        $ExcelWorksheet.Cells(45,4) = $NewApp.AppDependencies

        #Work on page 2
        $ExcelWorksheet = $ExcelWorkbooks.Worksheets.Item(3)
        $ExcelWorksheet.Cells(2,1) = Get-Date -Format d
        $ExcelWorksheet.Cells(2,2) = $ENV:USERNAME
        $ExcelWorksheet.Cells(2,3) = "Autogenerated by Generate-CMApplicationDocumentation.ps1. Please verify results."

        #Save and close Excel document
        $DocumentationFileName = $NewApp.AppName
        $DocumentationFilePath = ("$Destination\APP - $DocumentationFileName.xlsx").ToString()
        Write-Host "LOG: Saving documentation to $DocumentationFilePath"
        $ExcelWorkbooks.SaveAs($DocumentationFilePath)
        $ExcelProcess.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ExcelProcess) | Out-Null
        Get-Process | Where Name -like "*excel*" | Stop-Process -Force
        Remove-Item -Path "$ENV:TEMP\temp.xlsx" -Force

        #Clear variables
        Clear-Variable AppDependencies
    }
}

#Restore directory
Set-Location $CurrentPath