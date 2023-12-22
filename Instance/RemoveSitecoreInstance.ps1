[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
Function Get-FileName([System.Windows.Forms.TextBox]$ipBox) {
[Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
[System.Windows.Forms.Application]::EnableVisualStyles()
$browse = New-Object System.Windows.Forms.FolderBrowserDialog
$browse.RootFolder = [System.Environment+SpecialFolder]'MyComputer'
$browse.ShowNewFolderButton = $false
$browse.Description = "Choose a directory"
$loop = $true
while($loop)
{
    if ($browse.ShowDialog() -eq "OK")
    {
        $loop = $false
    }else
    {
        $res = [System.Windows.Forms.MessageBox]::Show("You clicked Cancel. Try again or exit script?", "Choose a directory", [System.Windows.Forms.MessageBoxButtons]::RetryCancel)
        if($res -eq "Cancel")
        {
            #End script
            return
        }
    }
}
$ipBox.Text = $browse.SelectedPath
$browse.SelectedPath
$browse.Dispose()
}

$objForm = New-Object System.Windows.Forms.Form 
$objForm.Text = "Sitecore Instance Uninstall"
$objForm.Size = New-Object System.Drawing.Size(900,800) 
$objForm.StartPosition = "CenterScreen"

$objForm.KeyPreview = $True
$objForm.Add_KeyDown({
    if ($_.KeyCode -eq "Enter" -or $_.KeyCode -eq "Escape"){
        $objForm.Close()
    }
})

## Site Name Input
$siteNameLabel = New-Object System.Windows.Forms.Label
$siteNameLabel.Location = New-Object System.Drawing.Size(10,20) 
$siteNameLabel.Size = New-Object System.Drawing.Size(280,20) 
$siteNameLabel.Text = "Site Name"
$objForm.Controls.Add($siteNameLabel) 

$siteNameBox = New-Object System.Windows.Forms.TextBox 
$siteNameBox.Location = New-Object System.Drawing.Size(10,40) 
$siteNameBox.Size = New-Object System.Drawing.Size(260,20)
$objForm.Controls.Add($siteNameBox)

## Service input Xconnect
$xConnectServiceNameLabel = New-Object System.Windows.Forms.Label
$xConnectServiceNameLabel.Location = New-Object System.Drawing.Size(10,60) 
$xConnectServiceNameLabel.Size = New-Object System.Drawing.Size(280,20) 
$xConnectServiceNameLabel.Text = "service name for - Sitecore XConnect Search Indexer"
$objForm.Controls.Add($xConnectServiceNameLabel) 

$xConnectServiceNameBox = New-Object System.Windows.Forms.TextBox 
$xConnectServiceNameBox.Location = New-Object System.Drawing.Size(10,80) 
$xConnectServiceNameBox.Size = New-Object System.Drawing.Size(260,20)
$objForm.Controls.Add($xConnectServiceNameBox)

## Service input Processing Engine
$processingEngineServiceLabel = New-Object System.Windows.Forms.Label
$processingEngineServiceLabel.Location = New-Object System.Drawing.Size(10,100) 
$processingEngineServiceLabel.Size = New-Object System.Drawing.Size(280,20) 
$processingEngineServiceLabel.Text = "service name for - Sitecore Processing Engine"
$objForm.Controls.Add($processingEngineServiceLabel)

$processingEngineServiceBox = New-Object System.Windows.Forms.TextBox 
$processingEngineServiceBox.Location = New-Object System.Drawing.Size(10,120) 
$processingEngineServiceBox.Size = New-Object System.Drawing.Size(260,20)
$objForm.Controls.Add($processingEngineServiceBox)

## Service input Marketing Automation
$marketingAutomationServiceLabel = New-Object System.Windows.Forms.Label
$marketingAutomationServiceLabel.Location = New-Object System.Drawing.Size(10,140) 
$marketingAutomationServiceLabel.Size = New-Object System.Drawing.Size(280,20) 
$marketingAutomationServiceLabel.Text = "service name for - Sitecore Sitecore Marketing Automation"
$objForm.Controls.Add($marketingAutomationServiceLabel)

$marketingAutomationServiceBox = New-Object System.Windows.Forms.TextBox 
$marketingAutomationServiceBox.Location = New-Object System.Drawing.Size(10,160) 
$marketingAutomationServiceBox.Size = New-Object System.Drawing.Size(260,20)
$objForm.Controls.Add($marketingAutomationServiceBox)

##DevSiteName
$devSiteNameLabel = New-Object System.Windows.Forms.Label
$devSiteNameLabel.Location = New-Object System.Drawing.Size(10,180) 
$devSiteNameLabel.Size = New-Object System.Drawing.Size(280,20) 
$devSiteNameLabel.Text = "Dev Site Name"
$objForm.Controls.Add($devSiteNameLabel)

$devSiteNameBox = New-Object System.Windows.Forms.TextBox 
$devSiteNameBox.Location = New-Object System.Drawing.Size(10,200) 
$devSiteNameBox.Size = New-Object System.Drawing.Size(260,20)
$objForm.Controls.Add($devSiteNameBox)

##Xconnect site name 

$xConnectSiteNameLabel = New-Object System.Windows.Forms.Label
$xConnectSiteNameLabel.Location = New-Object System.Drawing.Size(10,220) 
$xConnectSiteNameLabel.Size = New-Object System.Drawing.Size(280,20) 
$xConnectSiteNameLabel.Text = "Xconnect Site Name"
$objForm.Controls.Add($xConnectSiteNameLabel)

$xConnectSiteNameBox = New-Object System.Windows.Forms.TextBox 
$xConnectSiteNameBox.Location = New-Object System.Drawing.Size(10,240) 
$xConnectSiteNameBox.Size = New-Object System.Drawing.Size(260,20)
$objForm.Controls.Add($xConnectSiteNameBox)

##IdentityServer site name 

$identityServerSiteNameLabel = New-Object System.Windows.Forms.Label
$identityServerSiteNameLabel.Location = New-Object System.Drawing.Size(10,260) 
$identityServerSiteNameLabel.Size = New-Object System.Drawing.Size(280,20) 
$identityServerSiteNameLabel.Text = "Identityserver Site Name"
$objForm.Controls.Add($identityServerSiteNameLabel)

$identityServerSiteNameBox = New-Object System.Windows.Forms.TextBox 
$identityServerSiteNameBox.Location = New-Object System.Drawing.Size(10,280) 
$identityServerSiteNameBox.Size = New-Object System.Drawing.Size(260,20)
$objForm.Controls.Add($identityServerSiteNameBox)

##SQL Instance Name
$sqlInstanceNameLabel = New-Object System.Windows.Forms.Label
$sqlInstanceNameLabel.Location = New-Object System.Drawing.Size(10,300) 
$sqlInstanceNameLabel.Size = New-Object System.Drawing.Size(280,20) 
$sqlInstanceNameLabel.Text = "SQL Instance Name"
$objForm.Controls.Add($sqlInstanceNameLabel)

$sqlInstanceNameBox = New-Object System.Windows.Forms.TextBox 
$sqlInstanceNameBox.Location = New-Object System.Drawing.Size(10,320) 
$sqlInstanceNameBox.Size = New-Object System.Drawing.Size(260,20)
$objForm.Controls.Add($sqlInstanceNameBox)

##SQL Username
$sqlUsernameLabel = New-Object System.Windows.Forms.Label
$sqlUsernameLabel.Location = New-Object System.Drawing.Size(10,340) 
$sqlUsernameLabel.Size = New-Object System.Drawing.Size(280,20) 
$sqlUsernameLabel.Text = "SQL Username"
$objForm.Controls.Add($sqlUsernameLabel)

$sqlUsernameBox = New-Object System.Windows.Forms.TextBox 
$sqlUsernameBox.Location = New-Object System.Drawing.Size(10,360) 
$sqlUsernameBox.Size = New-Object System.Drawing.Size(260,20)
$objForm.Controls.Add($sqlUsernameBox)

##SQL Password
$sqlPasswordLabel = New-Object System.Windows.Forms.Label
$sqlPasswordLabel.Location = New-Object System.Drawing.Size(10,380) 
$sqlPasswordLabel.Size = New-Object System.Drawing.Size(280,20) 
$sqlPasswordLabel.Text = "SQL Password"
$objForm.Controls.Add($sqlPasswordLabel)

$sqlPasswordBox = New-Object Windows.Forms.MaskedTextBox
$sqlPasswordBox.Location = New-Object System.Drawing.Size(10,400) 
$sqlPasswordBox.Size = New-Object System.Drawing.Size(260,20)
$sqlPasswordBox.PasswordChar  = "*"
$objForm.Controls.Add($sqlPasswordBox)


## WebsitePhysicalRootPath
$websitePathLabel = New-Object System.Windows.Forms.Label
$websitePathLabel.Location = New-Object System.Drawing.Size(10,420) 
$websitePathLabel.Size = New-Object System.Drawing.Size(280,20) 
$websitePathLabel.Text = "Website Physical Root Path"
$objForm.Controls.Add($websitePathLabel)

$websitePathBox = New-Object System.Windows.Forms.TextBox 
$websitePathBox.Location = New-Object System.Drawing.Size(10,440) 
$websitePathBox.Size = New-Object System.Drawing.Size(260,20)
$objForm.Controls.Add($websitePathBox)
$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Location = New-Object System.Drawing.Size(150,460)
$browseButton.Size = New-Object System.Drawing.Size(75,23)
$browseButton.Text = "Browse"
$browseButton.Add_Click({Get-FileName $websitePathBox})
$objForm.Controls.Add($browseButton)

## SOLR Path
$solrPathLabel = New-Object System.Windows.Forms.Label
$solrPathLabel.Location = New-Object System.Drawing.Size(10,480) 
$solrPathLabel.Size = New-Object System.Drawing.Size(280,20) 
$solrPathLabel.Text = "SOLR Root Path"
$objForm.Controls.Add($solrPathLabel)

$solrPathBox = New-Object System.Windows.Forms.TextBox 
$solrPathBox.Location = New-Object System.Drawing.Size(10,500) 
$solrPathBox.Size = New-Object System.Drawing.Size(260,20)
$objForm.Controls.Add($solrPathBox)
$browseButtonSolr = New-Object System.Windows.Forms.Button
$browseButtonSolr.Location = New-Object System.Drawing.Size(150,520)
$browseButtonSolr.Size = New-Object System.Drawing.Size(75,23)
$browseButtonSolr.Text = "Browse"
$browseButtonSolr.Add_Click({Get-FileName $solrPathBox})
$objForm.Controls.Add($browseButtonSolr)

## Host File Path
$hostFilePathLabel = New-Object System.Windows.Forms.Label
$hostFilePathLabel.Location = New-Object System.Drawing.Size(10,540) 
$hostFilePathLabel.Size = New-Object System.Drawing.Size(280,20) 
$hostFilePathLabel.Text = "Host file Path"
$objForm.Controls.Add($hostFilePathLabel)

$hostFilePathBox = New-Object System.Windows.Forms.TextBox 
$hostFilePathBox.Location = New-Object System.Drawing.Size(10,560) 
$hostFilePathBox.Size = New-Object System.Drawing.Size(260,20)
$objForm.Controls.Add($hostFilePathBox)
$browseButtonHost = New-Object System.Windows.Forms.Button
$browseButtonHost.Location = New-Object System.Drawing.Size(150,580)
$browseButtonHost.Size = New-Object System.Drawing.Size(75,23)
$browseButtonHost.Text = "Browse"
$browseButtonHost.Add_Click({Get-FileName $hostFilePathBox})
$objForm.Controls.Add($browseButtonHost)

## OK BUTTON AND CANCEL BUTTON
$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Size(75,600)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = "OK"
$OKButton.Add_Click({$objForm.Close()})
$objForm.Controls.Add($OKButton)

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Size(150,600)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = "Cancel"
$CancelButton.Add_Click({$objForm.Close()})
$objForm.Controls.Add($CancelButton)


$objForm.Topmost = $True

$objForm.Add_Shown({$objForm.Activate()})
[void]$objForm.ShowDialog()

$siteName = $siteNameBox.Text
$SitecoreXConnectSearchIndexer = $xConnectServiceNameBox.Text
$SitecoreProcessingEngine = $processingEngineServiceBox.Text
$SitecoreMarketingAutomation = $marketingAutomationServiceBox.Text
$devSiteName = $devSiteNameBox.Text
$xconnectSiteName = $xConnectSiteNameBox.Text
$identityServerSiteName = $identityServerSiteNameBox.Text
$SQLInstanceName = $sqlInstanceNameBox.Text
$SQLUsername = $sqlUsernameBox.Text
$SQLPassword = $sqlPasswordBox.Text
$WebsitePhysicalRootPath = $websitePathBox.Text
$SolrPath = $solrPathBox.Text
$hostPath = $hostFilePathBox.Text+"\hosts"

#remove site and its dependencies
Write-host -foregroundcolor yellow "Removing Services" 
[System.Threading.Thread]::Sleep(1500)
if($SitecoreXConnectSearchIndexer){
	sc.exe delete $SitecoreXConnectSearchIndexer
	Write-Host -foregroundcolor Green 'Service removed' $SitecoreXConnectSearchIndexer
	[System.Threading.Thread]::Sleep(1500)
}if($SitecoreProcessingEngine){
	sc.exe delete $SitecoreProcessingEngine
	Write-Host -foregroundcolor Green 'Service removed' $SitecoreProcessingEngine
	[System.Threading.Thread]::Sleep(1500)
}if($SitecoreMarketingAutomation){
	sc.exe delete $SitecoreMarketingAutomation
	Write-Host -foregroundcolor Green 'Service removed' $SitecoreMarketingAutomation
	[System.Threading.Thread]::Sleep(1500)
}
Write-host -foregroundcolor green "Removed services for sites" 
[System.Threading.Thread]::Sleep(1500)

Write-host -foregroundcolor yellow "Removing App Pool for sites" 
[System.Threading.Thread]::Sleep(1500)
if($devSiteName){
	Remove-IISSite -Name $devSiteName
	Remove-WebAppPool -Name $devSiteName
	Write-Host -foregroundcolor Green $devSiteName ' -site and application pool removed successfully'
	[System.Threading.Thread]::Sleep(1500)
}if($xconnectSiteName ){
	Remove-IISSite -Name $xconnectSiteName 
	Remove-WebAppPool -Name $xconnectSiteName
	Write-Host -foregroundcolor Green $xconnectSiteName ' -site and application pool removed successfully'
	[System.Threading.Thread]::Sleep(1500)
}if($identityServerSiteName){
	Remove-IISSite -Name $identityServerSiteName
	Remove-WebAppPool -Name $identityServerSiteName
	Write-Host -foregroundcolor Green $identityServerSiteName ' -site and application removed successfully'
	[System.Threading.Thread]::Sleep(1500)
}
Write-Host -foregroundcolor Green "Removed app pool for sites!!" 
[System.Threading.Thread]::Sleep(1500)

Write-host -foregroundcolor yellow "Removing databases" 
[System.Threading.Thread]::Sleep(1500)
$DBListQuery = "select * from sys.databases where Name like '" + $siteName + "_%';"
#$SQLPasswords =[Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($SQLPassword))
$Output = Invoke-SqlCmd -ServerInstance $SQLInstanceName -username $SQLUsername -password $SQLPassword -Query $DBListQuery -ErrorAction Stop
Write-Host -foregroundcolor Green "SQL Connection Successful!!" 
[System.Threading.Thread]::Sleep(1500)

$DBList = invoke-sqlcmd -ServerInstance "$SQLInstanceName" -U "$SQLUsername" -P "$SQLPassword" -Query $DBListQuery 
ForEach($DB in $DBList) {
Write-host "Deleting Database " $DB.Name
$AlterQuery = "ALTER DATABASE [" + $DB.Name + "] SET SINGLE_USER;"
$DropQuery = "DROP DATABASE [" + $DB.Name + "];" 
Invoke-SqlCmd -ServerInstance "$SQLInstanceName" -U "$SQLUsername" -P "$SQLPassword" -Query $AlterQuery
Invoke-SqlCmd -ServerInstance "$SQLInstanceName" -U "$SQLUsername" -P "$SQLPassword" -Query $DropQuery
Write-host -foregroundcolor Green "Deleted Database " $DB.Name
[System.Threading.Thread]::Sleep(1500)
}
Write-host -foregroundcolor green "Removed Database" 
[System.Threading.Thread]::Sleep(1500)

Write-host -foregroundcolor yellow "Removing website folders from wwwroot" 
[System.Threading.Thread]::Sleep(1500)
$XConnectWebsitePhysicalPath = $WebsitePhysicalRootPath+"\"+$siteName+"xconnect.dev.local"
$SitecoreWebsitePhysicalPath = $WebsitePhysicalRootPath+"\"+$siteName+"sc.dev.local"
$IdentityWebsitePhysicalPath = $WebsitePhysicalRootPath+"\"+$siteName+"identityserver.dev.local"

if($XConnectWebsitePhysicalPath){
	Remove-Item -path $XConnectWebsitePhysicalPath\* -recurse 
	Remove-Item -path $XConnectWebsitePhysicalPath 
	Write-host -foregroundcolor Green $XConnectWebsitePhysicalPath " Deleted" 
	[System.Threading.Thread]::Sleep(1500) 
}
if($SitecoreWebsitePhysicalPath){
	Remove-Item -path $SitecoreWebsitePhysicalPath\* -recurse 
	Remove-Item -path $SitecoreWebsitePhysicalPath 
	Write-host -foregroundcolor Green $SitecoreWebsitePhysicalPath " Deleted" 
	[System.Threading.Thread]::Sleep(1500) 	
}
if($IdentityWebsitePhysicalPath){
	Remove-Item -path $IdentityWebsitePhysicalPath\* -recurse 
	Remove-Item -path $IdentityWebsitePhysicalPath 
	Write-host -foregroundcolor Green $IdentityWebsitePhysicalPath " Deleted" 
	[System.Threading.Thread]::Sleep(1500) 		
}
Write-host -foregroundcolor green "Removed website folders from wwwroot"
[System.Threading.Thread]::Sleep(1500)

Write-host -foregroundcolor yellow "Deleting SOLR Cores" 
[System.Threading.Thread]::Sleep(1500)
if($SolrPath){
	& "$SolrPath\bin\solr.cmd" delete -c ($siteName + "_core_index")
	& "$SolrPath\bin\solr.cmd" delete -c ($siteName + "_fxm_master_index")
	& "$SolrPath\bin\solr.cmd" delete -c ($siteName + "_fxm_web_index")
	& "$SolrPath\bin\solr.cmd" delete -c ($siteName + "_marketing_asset_index_master")
	& "$SolrPath\bin\solr.cmd" delete -c ($siteName + "_marketing_asset_index_web")
	& "$SolrPath\bin\solr.cmd" delete -c ($siteName + "_marketingdefinitions_master")
	& "$SolrPath\bin\solr.cmd" delete -c ($siteName + "_marketingdefinitions_web")
	& "$SolrPath\bin\solr.cmd" delete -c ($siteName + "_suggested_test_index")
	& "$SolrPath\bin\solr.cmd" delete -c ($siteName + "_testing_index")
	& "$SolrPath\bin\solr.cmd" delete -c ($siteName + "_web_index")
	& "$SolrPath\bin\solr.cmd" delete -c ($siteName + "_master_index")
	& "$SolrPath\bin\solr.cmd" delete -c ($siteName + "_xdb")
	& "$SolrPath\bin\solr.cmd" delete -c ($siteName + "_xdb_rebuild")
}
Write-host -foregroundcolor Green "SOLR Cores Deleted" 
[System.Threading.Thread]::Sleep(1500)

Write-host -foregroundcolor yellow "Deleting Host file entries" 
[System.Threading.Thread]::Sleep(1500)
Write-host "host file path" $hostPath
if($hostPath){
	(Get-Content -Path $hostPath) |
    Where-Object { -not $_.Contains($devSiteName ) } |
    Set-Content -Path $hostPath -Force
	
	(Get-Content -Path $hostPath) |
    Where-Object { -not $_.Contains($xconnectSiteName) } |
    Set-Content -Path $hostPath -Force
	
	(Get-Content -Path $hostPath) |
    Where-Object { -not $_.Contains($identityServerSiteName) } |
    Set-Content -Path $hostPath -Force
}
Write-host -foregroundcolor Green "Host File Entries deleted" 
[System.Threading.Thread]::Sleep(1500)
