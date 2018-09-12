<#
.SYNOPSIS
This script will create the DHCP options for an IP phone to use with SfB.
This is designed for Polycom Phones

.Role
Must have the ability to read the SfB Topology
Must have the ability to set scope options and create classes in DHCP

.DESCRIPTION
This script will read from SfB Topology the current pools.  It will then search AD for all of the DHCP Servers.
Once it has the DHCP Servers it will prompt you for which scope.
If on that servers the classes are not available, it will create the classes and then set the DHCP Options

.Notes
In case you are not using any kind of provisioning service for your device you can do the following
$EnableProvisioningPrompt = $false. If you set this to false it will not prompt and assume provisioning.
$Defaultoption160  = $null
$Defaultoption161 = $null

Basically these two options are used for provisioning services, like RPRM or Enoten
#>

#Null Everythung
Remove-Variable * -ErrorAction SilentlyContinue

################################################
#      Defualt Variables
#      Change Only in this Section
################################################

$EnableProvisioningPrompt = $True
$Defaultoption160 = "http://ha.usrprm.domain.com/phoneservice/configfiles" 
$Defaultoption161 = "http://ha.usrprm.domain.com/phoneservice/configfiles" 

################################################
#      Nothing to Change below here
################################################

#Select DHCP Server and Pool
do {
    try {$DHCPServer = Get-DhcpServerInDC | Out-GridView -Title "Choose DHCP Server" -PassThru -ErrorAction stop}
    catch {write-host "Could not get DHCP, but can continue without automatic configuration"}
}
while ($DHCPServer.count -gt 1)

if ($DHCPServer) {
    [Array]$DHCPScope = (Get-DhcpServerv4Scope -ComputerName $DHCPServer.DnsName |Select-Object ScopeID, Name | Out-GridView -PassThru -Title $DHCPServer).ScopeId.IPAddressToString
}

#Get Front End pools. If we cannot, exit
try {
    $pool = Get-CsService -Registrar | Select-Object PoolFqdn,SiteId | Out-GridView -Title "Which Front End Pool" -PassThru
}
catch {
    write-host "Could not get front end pool list, cannot continue"
    exit
}
$URLS = Get-CsService -Registrar -PoolFqdn $pool.PoolFqdn | ForEach-Object {Get-CsService -WebServer -PoolFqdn $_.PoolFQDN} | Select-Object InternalFQDN,PoolFQDN

#Select Time Zone
$Timezonelist = Get-TimeZone -ListAvailable | Select-Object BaseUtcOffset,ID | Sort-Object BaseUtcOffset | Out-GridView -PassThru
$timezoneselect = get-timezone $Timezonelist.ID
$timeoffsetbit = [convert]::tostring($timezoneselect.baseutcOffset.totalseconds,16)
if ($timeoffsetbit.Length -gt 7) {
    $timeoffset = "0x" + $timeoffsetbit.Substring($timeoffsetbit.Length -8,8).ToUpper()
}
else {
for ($i=0;$i -lt (8-$timeoffsetbit.Length);$i++) {$timepre += "0"}
    $timeoffset = "0x" + $timepre + $timeoffsetbit.Substring($timeoffsetbit.Length -$timeoffsetbit.Length,$timeoffsetbit.Length).ToUpper()
}

if (-not $dhcpserver.IPAddress.IPAddressToString) {
    $timeserver = read-host "No DHCP Server selected, please enter a time server: "
}
else {
    #Prompt for Time Server
    if (($timeserver = read-host -Prompt "Time Server: Use $($dhcpserver.IPAddress.IPAddressToString) or enter custom") -ne '') {
        Write-host $timeserver
    }
    else {
        $timeserver = $dhcpserver.IPAddress.IPAddressToString
    }
}

#Prompt if using provisioning
if ($EnableProvisioningPrompt -eq $true) {
    $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Provisioning is used."
    $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No",  "NO Provisioning used."
    $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
    $result = $host.ui.PromptForChoice("Polycom Provisioning", "Do you use Polycom Provisioning", $options, 0) 
    switch ($result)
        {
            0 {
                $Provisioning = $true

                if (($option160  = read-host -Prompt "Provisioning Server SIP: Use $Defaultoption160  as default or enter custom and hit enter") -ne '') {
                    Write-host $option160 
                }
                else {
                    $option160 = $Defaultoption160
                }
        
                if (($option161  = read-host -Prompt "Provisioning Server UC: Use $Defaultoption161  as default or enter custom and hit enter") -ne '') {
                    Write-host $option161
                }
                else {
                    $option161 = $Defaultoption161
                }
            }

            1 {$Provisioning = $false}

            default {
                write-host "No Option picked, exiting"
                exit
            }
        }
}

$PoolFqdn = $null
$InternalFqdn = $null
$UCIdentifier = $null
$URLScheme = $null
$WebServerPort = $null
$CertRelPath = $null

#Convert the Strings to Hex
Foreach ($char in ($URLS.PoolFqdn.toString().ToLower().ToCharArray())) {$PoolFqdn=$PoolFqdn+[System.String]::Format("{0:X}",[System.Convert]::ToUInt32($char))}
Foreach ($char in ($URLS.InternalFqdn.toString().ToLower().ToCharArray())) {$InternalFqdn=$InternalFqdn+[System.String]::Format("{0:X}",[System.Convert]::ToUInt32($char))}
Foreach ($char in ("MS-UC-Client".ToCharArray())) {$UCIdentifier=$UCIdentifier+[System.String]::Format("{0:X}",[System.Convert]::ToUInt32($char))}
Foreach ($char in ("https".ToCharArray())) {$URLScheme=$URLScheme+[System.String]::Format("{0:X}",[System.Convert]::ToUInt32($char))}
Foreach ($char in ("443".ToCharArray())) {$WebServerPort=$WebServerPort+[System.String]::Format("{0:X}",[System.Convert]::ToUInt32($char))}
Foreach ($char in ("/CertProv/CertProvisioningService.svc".ToCharArray())) {$CertRelPath=$CertRelPAth+[System.String]::Format("{0:X}",[System.Convert]::ToUInt32($char))}

#Write out all the settings
write-host "DHCP Server: " $DHCPServer.DnsName
Write-host "DHCP Scope: " $DHCPScope
write-host "-------"
write-host "SFB Pool Name and binary: " + $URLS.PoolFqdn.ToString().ToLower() + " : " + $PoolFqdn.TrimEnd(",")
write-host "SfB Web Services Name and Binary: " $URLS.InternalFqdn.ToString().ToLower() "  : "  $InternalFqdn.TrimEnd(",")
write-host "MS-UC-Client  : "  $UCIdentifier.TrimEnd(",")
write-host "https  : "  $URLScheme.TrimEnd(",")
write-host "443  : "  $WebServerPort.TrimEnd(",")
write-host "/CertProv/CertProvisioningService.svc  : "  $CertRelPath.TrimEnd(",")
write-host "-------"
write-host "Time Server: " $timeserver
Write-Host "Time Zone: " $timezoneselect
write-host "Time Offset Bit" $timeoffsetbit
write-host "Time Offset: " $timeoffset

##
# If no DHCP Server is selected just pring out the settings and exit
##
if (-not $DHCPServer) {
    write-host -ForegroundColor yellow "---Since No DHCP Server Selected here is output to import---"
    write-host -ForegroundColor yellow "---You can copy and paste these into a DHCP server---"
    write-host -ForegroundColor yellow "---If you are not using windows DHCP, good luck---"

$dchpclasses=@"
Add-DhcpServerv4Class -Name "MSUCClient" -Type Vendor -Description "UC Vendor Class Id" -Data "MS-UC-Client"
Add-DhcpServerv4OptionDefinition -VendorClass MSUCClient -OptionId 1 -Name UCIdentifier -Type BinaryData -Description "UC Identifier" 
Add-DhcpServerv4OptionDefinition -VendorClass MSUCClient -OptionId 2 -Name URLScheme -Type BinaryData -Description "URL Scheme"
Add-DhcpServerv4OptionDefinition -VendorClass MSUCClient -OptionId 3 -Name WebServerFqdn -Type BinaryData -Description "Web Server Fqdn" 
Add-DhcpServerv4OptionDefinition -VendorClass MSUCClient -OptionId 4 -Name WebServerPort -Type BinaryData -Description "Web Server Port"
Add-DhcpServerv4OptionDefinition -VendorClass MSUCClient -OptionId 5 -Name CertProvRelPath -Type BinaryData -Description "Cert Prov Relative Path" 
"@

$options=@"
Add-DhcpServerv4OptionDefinition -Name "UCSipServer" -Description "Sip Server Fqdn" -OptionId 120 -Type BinaryData
"@

if ($Provisioning -eq $true) {

$Provisioningoptions=@"
Add-DhcpServerv4OptionDefinition -Name "SIP Provisioning Server" -Description "Polycom SIP Provisioning Server" -OptionId 160 -Type String
Add-DhcpServerv4OptionDefinition -Name "SfB Provisioning Server" -Description "Polycom SfB Provisioning Server" -OptionId 161 -Type String
Set-DhcpServerv4OptionValue -scopeID `$ScopeID -OptionId 160 -Value $option160
Set-DhcpServerv4OptionValue -ScopeId `$ScopeID -OptionId 161 -Value $option161
"@

}

$scopeoptions =@"
Set-DhcpServerv4OptionValue -ScopeId `$ScopeID -VendorClass MSUCClient -OptionId 1 -Value $UCIdentifier
Set-DhcpServerv4OptionValue -ScopeId `$ScopeID -VendorClass MSUCClient -OptionId 2 -Value $URLScheme 
Set-DhcpServerv4OptionValue -ScopeId `$ScopeID -VendorClass MSUCClient -OptionId 3 -Value $InternalFqdn 
Set-DhcpServerv4OptionValue -ScopeId `$ScopeID -VendorClass MSUCClient -OptionId 4 -Value $WebServerPort
Set-DhcpServerv4OptionValue -ScopeId `$ScopeID -VendorClass MSUCClient -OptionId 5 -Value $CertRelPath
Set-DhcpServerv4OptionValue -ScopeId `$ScopeID -OptionId 120 -Value $PoolFqdn 
Set-DhcpServerv4OptionValue -ScopeId `$ScopeID -OptionId 4 -Value $timeserver
Set-DhcpServerv4OptionValue -ScopeId `$ScopeID -OptionId 2 -Value $timeoffset
"@


$outfile = [System.IO.Path]::GetTempFileName()
$outstring = "`$ScopeID=`"10.1.1.0`""
$outstring | out-file -FilePath $outfile
$outstring = $dchpclasses+$options+$Provisioningoptions+$scopeoptions
$outstring | out-file -FilePath $outfile -Append

notepad $outfile
exit
} #End that there was NO DHCP Server Found


if ((read-host -Prompt "Continue with configuration, press enter continue or any other key to stop") -ne '') {
Exit
}


#Write the actual data to DHCP

$PoolFqdn=[System.Text.Encoding]::ASCII.GetBytes($URLS.PoolFqdn.toString().ToLower())
$InternalFqdn = [System.Text.Encoding]::ASCII.GetBytes($URLS.InternalFqdn.toString().ToLower())
$UCIdentifier = [System.Text.Encoding]::ASCII.GetBytes("MS-UC-Client")
$URLScheme = [System.Text.Encoding]::ASCII.GetBytes("https")
$WebServerPort = [System.Text.Encoding]::ASCII.GetBytes("443")
$CertRelPath = [System.Text.Encoding]::ASCII.GetBytes("/CertProv/CertProvisioningService.svc")

[System.Text.Encoding]::ASCII.GetString($PoolFqdn)
[System.Text.Encoding]::ASCII.GetString($InternalFqdn)
[System.Text.Encoding]::ASCII.GetString($UCIdentifier)
[System.Text.Encoding]::ASCII.GetString($URLScheme)
[System.Text.Encoding]::ASCII.GetString($WebServerPort)
[System.Text.Encoding]::ASCII.GetString($CertRelPath)

#Create the vendor Class
$foundclass = $false
foreach ($name in (Get-DhcpServerv4Class -ComputerName $DHCPServer.DnsName -Type Vendor)) {
    if ($name.name.tolower() -like "msucclient") {
        $FoundClass = $true
    }
}

if ($foundclass -ne $true ) {
    Add-DhcpServerv4Class -Name "MSUCClient" -Type Vendor -Description "UC Vendor Class Id" -ComputerName $DHCPServer.DnsName -Data "MS-UC-Client"
    #Create the Vendor Class Options
    Add-DhcpServerv4OptionDefinition -VendorClass MSUCClient -OptionId 1 -Name UCIdentifier -Type BinaryData -ComputerName $DHCPServer.DnsName -Description "UC Identifier" 
    Add-DhcpServerv4OptionDefinition -VendorClass MSUCClient -OptionId 2 -Name URLScheme -Type BinaryData -ComputerName $DHCPServer.DnsName -Description "URL Scheme"
    Add-DhcpServerv4OptionDefinition -VendorClass MSUCClient -OptionId 3 -Name WebServerFqdn -Type BinaryData -ComputerName $DHCPServer.DnsName -Description "Web Server Fqdn" 
    Add-DhcpServerv4OptionDefinition -VendorClass MSUCClient -OptionId 4 -Name WebServerPort -Type BinaryData -ComputerName $DHCPServer.DnsName -Description "Web Server Port"
    Add-DhcpServerv4OptionDefinition -VendorClass MSUCClient -OptionId 5 -Name CertProvRelPath -Type BinaryData -ComputerName $DHCPServer.DnsName -Description "Cert Prov Relative Path" 
}

$error.Clear()
#Test Option 120
try {Get-DhcpServerv4OptionDefinition -ComputerName $DHCPServer.DnsName  -OptionId 120 -ErrorAction stop}
catch {
    $temperror = ($error[0].Exception).MessageID
}
if ($temperror -like "DHCP 20010") {
    Add-DhcpServerv4OptionDefinition -ComputerName $DHCPServer.DnsName -Name "UCSipServer" -Description "Sip Server Fqdn" -OptionId 120 -Type BinaryData
}
$error.Clear()
$temperror = $null
if ($Provisioning -eq $true) {
    #Test Provisioning Options
    try {Get-DhcpServerv4OptionDefinition -ComputerName $DHCPServer.DnsName -OptionId 160 -ErrorAction stop}
    catch {
        $temperror = ($error[0].Exception).MessageID
    }
    if ($temperror -like "DHCP 20010") {
        Add-DhcpServerv4OptionDefinition -ComputerName $DHCPServer.DnsName -Name "UCC SIP Provisioning Server" -Description "UCC SIP Provisioning Server" -OptionId 160 -Type String
        Add-DhcpServerv4OptionDefinition -ComputerName $DHCPServer.DnsName -Name "UCC SfB Provisioning Server" -Description "UCC SfB Provisioning Server" -OptionId 161 -Type String
    }
    $error.Clear()
}

#SfB Polycom Phones

    for ($i=0;$i -lt $DHCPScope.count; $i++) {
        Set-DhcpServerv4OptionValue -ScopeId $DHCPScope[$i] -VendorClass MSUCClient -OptionId 1 -Value $UCIdentifier -computername $DHCPServer.DnsName
        Set-DhcpServerv4OptionValue -ScopeId $DHCPScope[$i] -VendorClass MSUCClient -OptionId 2 -Value $URLScheme -computername $DHCPServer.DnsName
        Set-DhcpServerv4OptionValue -ScopeId $DHCPScope[$i] -VendorClass MSUCClient -OptionId 3 -Value $InternalFqdn -computername $DHCPServer.DnsName
        Set-DhcpServerv4OptionValue -ScopeId $DHCPScope[$i] -VendorClass MSUCClient -OptionId 4 -Value $WebServerPort -computername $DHCPServer.DnsName
        Set-DhcpServerv4OptionValue -ScopeId $DHCPScope[$i] -VendorClass MSUCClient -OptionId 5 -Value $CertRelPath  -computername $DHCPServer.DnsName
        Set-DhcpServerv4OptionValue -ScopeId $DHCPScope[$i] -OptionId 120 -Value $PoolFqdn  -computername $DHCPServer.DnsName
        #Time
        Set-DhcpServerv4OptionValue -ScopeId $DHCPScope[$i] -OptionId 4 -Value $timeserver -computername $DHCPServer.DnsName
        Set-DhcpServerv4OptionValue -ScopeId $DHCPScope[$i] -OptionId 2 -Value $timeoffset -computername $DHCPServer.DnsName
        #Polycom Specific
        if ($EnableProvisioningPrompt = $true) {
        Set-DhcpServerv4OptionValue -ScopeId $DHCPScope[$i] -OptionId 160 -Value $option160 -computername $DHCPServer.DnsName
        Set-DhcpServerv4OptionValue -ScopeId $DHCPScope[$i] -OptionId 161 -Value $option161 -computername $DHCPServer.DnsName
        }

    }


