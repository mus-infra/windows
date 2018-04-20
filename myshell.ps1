# Copyright (c) Microsoft Corporation. All rights reserved.  

#load hashtable of localized string
Import-LocalizedData -BindingVariable RemoteExchange_LocalizedStrings -FileName RemoteExchange.strings.psd1

## INCREASE WINDOW WIDTH #####################################################
function WidenWindow([int]$preferredWidth)
{
  $host.ui.rawui.BackgroundColor = "Black"
  $host.ui.rawui.ForegroundColor = "Green"
  [int]$maxAllowedWindowWidth = $host.ui.rawui.MaxPhysicalWindowSize.Width
  if ($preferredWidth -lt $maxAllowedWindowWidth)
  {
    # first, buffer size has to be set to windowsize or more
    # this operation does not usually fail
    $current=$host.ui.rawui.BufferSize
    $bufferWidth = $current.width
    if ($bufferWidth -lt $preferredWidth)
    {
      $current.width=$preferredWidth
      $host.ui.rawui.BufferSize=$current
    }
    # else not setting BufferSize as it is already larger
    
    # setting window size. As we are well within max limit, it won't throw exception.
    $current=$host.ui.rawui.WindowSize
    if ($current.width -lt $preferredWidth)
    {
      $current.width=$preferredWidth
      $host.ui.rawui.WindowSize=$current
    }
    #else not setting WindowSize as it is already larger
  }
}

WidenWindow(120)

## ALIASES ###################################################################

set-alias list       format-list 
set-alias table      format-table 

## Confirmation is enabled by default, uncommenting next line will disable it 
# $ConfirmPreference = "None"

## EXCHANGE VARIABLEs ########################################################

$global:exbin = (get-itemproperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup).MsiInstallPath + "bin\"
$global:exinstall = (get-itemproperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup).MsiInstallPath
$global:exscripts = (get-itemproperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup).MsiInstallPath + "scripts\"

## LOAD CONNECTION FUNCTIONS #################################################

. $global:exbin"CommonConnectFunctions.ps1"
. $global:exbin"ConnectFunctions.ps1"

$FormatEnumerationLimit = 16

## LOAD EXCHANGE EXTENDED TYPE INFORMATION ###################################

# loads powershell types file, parses out just the type names and returns an array of string
# it skips all template types as template parameter types individually are defined in types file
function GetTypeListFromXmlFile( [string] $typeFileName ) 
{
	$xmldata = [xml](Get-Content $typeFileName)
	$returnList = $xmldata.Types.Type | where { ($_.Name.StartsWith("Microsoft.Exchange") -and !$_.Name.Contains("[[")) } | foreach { $_.Name }
	return $returnList
}

$ConfigurationPath = join-path $global:exbin "Microsoft.Exchange.Configuration.ObjectModel.dll"
[System.Reflection.Assembly]::LoadFrom($ConfigurationPath) > $null

# Check if every single type from from Exchange.Types.ps1xml can be successfully loaded
$ManagementPath = join-path $global:exbin "Microsoft.Exchange.Management.dll"
$typeFilePath = join-path $global:exbin "exchange.types.ps1xml"
$typeListToCheck = GetTypeListFromXmlFile $typeFilePath
$typeLoadResult = [Microsoft.Exchange.Configuration.Tasks.TaskHelper]::TryLoadExchangeTypes($ManagementPath, $typeListToCheck)
# $typeListToCheck is a big list, release it to free up some memory
$typeListToCheck = $null


if (Get-ItemProperty HKLM:\Software\microsoft\ExchangeServer\v14\CentralAdmin -ea silentlycontinue)
{
	$CentralAdminPath = join-path $global:exbin "Microsoft.Exchange.Management.Powershell.CentralAdmin.dll"
	[Microsoft.Exchange.Configuration.Tasks.TaskHelper]::LoadExchangeAssemblyAndReferences($CentralAdminPath) > $null
}

# Register Assembly Resolver to handle generic types
[Microsoft.Exchange.Data.SerializationTypeConverter]::RegisterAssemblyResolver()


# Finally, load the types information
# We will load type information only if every single type from Exchange.Types.ps1xml can be successfully loaded
if ($typeLoadResult)
{
	Update-TypeData -PrependPath $typeFilePath
}
else
{
	# put a short warning message here that we are skipping type loading
	write-host -fore Yellow $RemoteExchange_LocalizedStrings.res_types_file_not_loaded
}

#load partial types
$partialTypeFile = join-path $global:exbin "Exchange.partial.Types.ps1xml"
Update-TypeData -PrependPath $partialTypeFile 

# If Central Admin cmdlets are installed, it loads the types information for those too
if (Get-ItemProperty HKLM:\Software\microsoft\ExchangeServer\v14\CentralAdmin -ea silentlycontinue)
{
	$typeFile = join-path $global:exbin "Exchange.CentralAdmin.Types.ps1xml"
	Update-TypeData -PrependPath $typeFile
}


## FUNCTIONs #################################################################

## returns all defined functions 

function functions
{ 
    if ( $args ) 
    { 
        foreach($functionName in $args )
        {
             get-childitem function:$functionName | 
                  foreach { "function " + $_.Name; "{" ; $_.Definition; "}" }
        }
    } 
    else 
    { 
        get-childitem function: | 
             foreach { "function " + $_.Name; "{" ; $_.Definition; "}" }
    } 
}

## only returns exchange commands 

function get-excommand
{
	if ($args[0] -eq $null)
	{
		get-command -module $global:importResults
	}
	else
	{
		get-command $args[0] | where { $_.module -eq $global:importResults }
	}
}


## only returns PowerShell commands 

function get-pscommand
{
	if ($args[0] -eq $null) 
	{
		get-command -pssnapin Microsoft.PowerShell* 
	}
	else 
	{
		get-command $args[0] | where { $_.PsSnapin -ilike 'Microsoft.PowerShell*' }	
	}
}

## prints the Exchange Banner in pretty colors 

function get-exbanner
{
	write-host "Welcome to Paul's Bitchin' PowerShell"

	write-host -no $RemoteExchange_LocalizedStrings.res_full_list
	write-host -no " "
	write-host -fore Yellow $RemoteExchange_LocalizedStrings.res_0003

	write-host -no $RemoteExchange_LocalizedStrings.res_only_exchange_cmdlets
	write-host -no " "
	write-host -fore Yellow $RemoteExchange_LocalizedStrings.res_0005

	write-host -no $RemoteExchange_LocalizedStrings.res_cmdlets_specific_role
	write-host -no " "
	write-host -fore Yellow $RemoteExchange_LocalizedStrings.res_0007

	write-host -no $RemoteExchange_LocalizedStrings.res_general_help
	write-host -no " "
	write-host -fore Yellow $RemoteExchange_LocalizedStrings.res_0009

	write-host -no $RemoteExchange_LocalizedStrings.res_help_for_cmdlet
	write-host -no " "
	write-host -fore Yellow $RemoteExchange_LocalizedStrings.res_0011

	write-host -no $RemoteExchange_LocalizedStrings.res_show_quick_ref
	write-host -no " "
	write-host -fore Yellow $RemoteExchange_LocalizedStrings.res_0013

	write-host -no $RemoteExchange_LocalizedStrings.res_team_blog
	write-host -no " "
	write-host -fore Yellow $RemoteExchange_LocalizedStrings.res_0015

	write-host -no $RemoteExchange_LocalizedStrings.res_show_full_output
	write-host -no " "
	write-host -fore Yellow $RemoteExchange_LocalizedStrings.res_0017
	
	write-host "Log in to a vCenter Server or ESX host:              " -NoNewLine
	write-host "Connect-VIServer" -foregroundcolor yellow
	write-host "To find out what commands are available, type:       " -NoNewLine
	write-host "Get-VICommand" -foregroundcolor yellow
	write-host "To show searchable help for all PowerCLI commands:   " -NoNewLine
	write-host "Get-PowerCLIHelp" -foregroundcolor yellow  
	write-host "Once you've connected, display all virtual machines: " -NoNewLine
	write-host "Get-VM" -foregroundcolor yellow
	write-host "If you need more help, visit the PowerCLI community: " -NoNewLine
	write-host "Get-PowerCLICommunity" -foregroundcolor yellow
}

## shows quickref guide

function quickref
{
    $exchculture = [System.Threading.Thread]::CurrentThread.CurrentUICulture
    $foundculture = $false
    while($exchculture -ne [Globalization.CultureInfo]::InvariantCulture -and !$foundculture)
    {
    	if ( test-path "$($global:exbin)\$($exchculture.Name)\exquick.htm" )
    	{
    		$foundculture = $true
    	}
    	else
    	{
    		$exchculture = $exchculture.Parent
    	}
    }

	if($foundculture -eq $true)
    {
	$exchculture = $exchculture.Name 
    }
    else 
    {
	$exchculture = 'en'
    } 

    if ( test-path "$($global:exbin)\$exchculture\exquick.htm" )
    {
	invoke-item $global:exbin\$exchculture\exquick.htm
    }
    else
    {
      ($RemoteExchange_LocalizedStrings.res_quickstart_not_found -f "$global:exbin\$exchculture\exquick.htm")
    } 
}

function get-exblog
{
       invoke-expression 'cmd /c start http://go.microsoft.com/fwlink/?LinkId=35786'
}

## FILTERS #################################################################
## Assembles a message and writes it to file from many sequential BinaryFileDataObject instances 
Filter AssembleMessage ([String] $Path) { Add-Content -Path:"$Path" -Encoding:"Byte" -Value:$_.FileData }

## now actually call the functions 

$host.ui.rawui.WindowTitle = "Paul's Bitchin' PowerShell"
$host.ui.rawui.BackgroundColor = "Black"
$host.ui.rawui.ForegroundColor = "Green"

get-exbanner 
get-tip
add-pssnapin VMWare*
add-pssnapin Quest*

#
# TIP: You can create your own customizations and put them in My Documents\WindowsPowerShell\profile.ps1
# Anything in profile.ps1 will then be run every time you start the shell. 
#


# SIG # Begin signature block
# MIIbGQYJKoZIhvcNAQcCoIIbCjCCGwYCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUquegTpjfFzKZtIZ3JzNt9TvU
# Sz2gghXyMIIEoDCCA4igAwIBAgIKYRr16gAAAAAAajANBgkqhkiG9w0BAQUFADB5
# MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVk
# bW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSMwIQYDVQQDExpN
# aWNyb3NvZnQgQ29kZSBTaWduaW5nIFBDQTAeFw0xMTExMDEyMjM5MTdaFw0xMzAy
# MDEyMjQ5MTdaMIGDMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
# MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
# MQ0wCwYDVQQLEwRNT1BSMR4wHAYDVQQDExVNaWNyb3NvZnQgQ29ycG9yYXRpb24w
# ggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQDDqR/PfCN/MR4GJYnddXm5
# z5NLYZK2lfLvqiWdd/NLWm1JkMzgMbimAjeHdK/yrKBglLjHTiX+h9hY0iBOLfE6
# ZS6SW6Zd5pV14DTlUCGcfTmXto5EI2YWpmUg4Dbrivqd4stgAfwqZMiHRRTxHsrN
# KKy65VdZJtzsxUpsmuYDGikyPwCeg6wlDYTM3W+2arst94Q6bWYx6DZw/4SSkPdA
# dp6ILkfWKxH3j+ASZSu8X+8V/PfsAWi3RQzuwASwDre9eGuujeRQ8TXingHS4etb
# cYJhISDz1MneHLgCRWVJvn61N4anzexa37h2IPwRE1H8+ipQqrQe0DqAvmPK3IFH
# AgMBAAGjggEdMIIBGTATBgNVHSUEDDAKBggrBgEFBQcDAzAdBgNVHQ4EFgQUAAOm
# 5aLEcaKCw492zSwNEuKdSigwDgYDVR0PAQH/BAQDAgeAMB8GA1UdIwQYMBaAFFdF
# dBxdsPbIQwXgjFQtjzKn/kiWMFYGA1UdHwRPME0wS6BJoEeGRWh0dHA6Ly9jcmwu
# bWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3RzL01pY0NvZFNpZ1BDQV8wOC0z
# MS0yMDEwLmNybDBaBggrBgEFBQcBAQROMEwwSgYIKwYBBQUHMAKGPmh0dHA6Ly93
# d3cubWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljQ29kU2lnUENBXzA4LTMxLTIw
# MTAuY3J0MA0GCSqGSIb3DQEBBQUAA4IBAQCQ9/h5kmnIj2uKYO58wa4+gThS9LrP
# mYzwLT0T9K72YfB1OE5Zxj8HQ/kHfMdT5JFi1qh2FHWUhlmyuhDCf2wVPxkVww4v
# fjnDz/5UJ1iUNWEHeW1RV7AS4epjcooWZuufOSozBDWLg94KXjG8nx3uNUUNXceX
# 3yrgnX86SfvjSEUy3zZtCW52VVWsNMV5XW4C1cyXifOoaH0U6ml7C1V9AozETTC8
# Yvd7peygkvAOKg6vV5spSM22IaXqHe/cCfWrYtYN7DVfa5nUsfB3Uvl36T9smFbA
# XDahTl4Q9Ix6EZcgIDEIeW5yFl8cMFeby3yiVfVwbHjsoUMgruywNYsYMIIEujCC
# A6KgAwIBAgIKYQUTNgAAAAAAGjANBgkqhkiG9w0BAQUFADB3MQswCQYDVQQGEwJV
# UzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UE
# ChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEwHwYDVQQDExhNaWNyb3NvZnQgVGlt
# ZS1TdGFtcCBQQ0EwHhcNMTEwNzI1MjA0MjE3WhcNMTIxMDI1MjA0MjE3WjCBszEL
# MAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1v
# bmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjENMAsGA1UECxMETU9Q
# UjEnMCUGA1UECxMebkNpcGhlciBEU0UgRVNOOjE1OUMtQTNGNy0yNTcwMSUwIwYD
# VQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNlMIIBIjANBgkqhkiG9w0B
# AQEFAAOCAQ8AMIIBCgKCAQEAnDSYGckJKWOZAhZ1qIhXfaG7qUES/GSRpdYFeL93
# 3OzmrrhQTsDjGr3tt/34IIpxOapyknKfignlE++RQe1hJWtRre6oQ7VhQiyd8h2x
# 0vy39Xujc3YTsyuj25RhgFWhD23d2OwW/4V/lp6IfwAujnokumidj8bK9JB5euGb
# 7wZdfvguw2oVnDwUL+fVlMgiG1HLqVWGIbda80ESOZ/wValOqiUrY/uRcjwPfMCW
# ctzBo8EIyt7FybXACl+lnAuqcgpdCkB9LpjQq7KIj4aA6H3RvlVr4FgsyDY/+eYR
# w/BDBYV4AxflLKcpfNPilRcAbNvcrTwZOgLgfWLUzvYdPQIDAQABo4IBCTCCAQUw
# HQYDVR0OBBYEFPaDiyCHEe6Dy9vehaLSaIY3YXSQMB8GA1UdIwQYMBaAFCM0+NlS
# RnAK7UD7dvuzK7DDNbMPMFQGA1UdHwRNMEswSaBHoEWGQ2h0dHA6Ly9jcmwubWlj
# cm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3RzL01pY3Jvc29mdFRpbWVTdGFtcFBD
# QS5jcmwwWAYIKwYBBQUHAQEETDBKMEgGCCsGAQUFBzAChjxodHRwOi8vd3d3Lm1p
# Y3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY3Jvc29mdFRpbWVTdGFtcFBDQS5jcnQw
# EwYDVR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZIhvcNAQEFBQADggEBAGL0BQ1P5xtr
# gudSDN95jKhVgTOX06TKyf6vSNt72m96KE/H0LeJ2NGmmcyRVgA7OOi3Mi/u+c9r
# 2Zje1gL1QlhSa47aQNwWoLPUvyYVy0hCzNP9tPrkRIlmD0IOXvcEnyNIW7SJQcTa
# bPg29D/CHhXfmEwAxLLs3l8BAUOcuELWIsiTmp7JpRhn/EeEHpFdm/J297GOch2A
# djw2EUbKfjpI86/jSfYXM427AGOCnFejVqfDbpCjPpW3/GTRXRjCCwFQY6f889GA
# noTjMjTdV5VAo21+2usuWgi0EAZeMskJ6TKCcRan+savZpiJ+dmetV8QI6N3gPJN
# 1igAclCFvOUwggYHMIID76ADAgECAgphFmg0AAAAAAAcMA0GCSqGSIb3DQEBBQUA
# MF8xEzARBgoJkiaJk/IsZAEZFgNjb20xGTAXBgoJkiaJk/IsZAEZFgltaWNyb3Nv
# ZnQxLTArBgNVBAMTJE1pY3Jvc29mdCBSb290IENlcnRpZmljYXRlIEF1dGhvcml0
# eTAeFw0wNzA0MDMxMjUzMDlaFw0yMTA0MDMxMzAzMDlaMHcxCzAJBgNVBAYTAlVT
# MRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQK
# ExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xITAfBgNVBAMTGE1pY3Jvc29mdCBUaW1l
# LVN0YW1wIFBDQTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAJ+hbLHf
# 20iSKnxrLhnhveLjxZlRI1Ctzt0YTiQP7tGn0UytdDAgEesH1VSVFUmUG0KSrphc
# MCbaAGvoe73siQcP9w4EmPCJzB/LMySHnfL0Zxws/HvniB3q506jocEjU8qN+kXP
# CdBer9CwQgSi+aZsk2fXKNxGU7CG0OUoRi4nrIZPVVIM5AMs+2qQkDBuh/NZMJ36
# ftaXs+ghl3740hPzCLdTbVK0RZCfSABKR2YRJylmqJfk0waBSqL5hKcRRxQJgp+E
# 7VV4/gGaHVAIhQAQMEbtt94jRrvELVSfrx54QTF3zJvfO4OToWECtR0Nsfz3m7IB
# ziJLVP/5BcPCIAsCAwEAAaOCAaswggGnMA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0O
# BBYEFCM0+NlSRnAK7UD7dvuzK7DDNbMPMAsGA1UdDwQEAwIBhjAQBgkrBgEEAYI3
# FQEEAwIBADCBmAYDVR0jBIGQMIGNgBQOrIJgQFYnl+UlE/wq4QpTlVnkpKFjpGEw
# XzETMBEGCgmSJomT8ixkARkWA2NvbTEZMBcGCgmSJomT8ixkARkWCW1pY3Jvc29m
# dDEtMCsGA1UEAxMkTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5
# ghB5rRahSqClrUxzWPQHEy5lMFAGA1UdHwRJMEcwRaBDoEGGP2h0dHA6Ly9jcmwu
# bWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3RzL21pY3Jvc29mdHJvb3RjZXJ0
# LmNybDBUBggrBgEFBQcBAQRIMEYwRAYIKwYBBQUHMAKGOGh0dHA6Ly93d3cubWlj
# cm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljcm9zb2Z0Um9vdENlcnQuY3J0MBMGA1Ud
# JQQMMAoGCCsGAQUFBwMIMA0GCSqGSIb3DQEBBQUAA4ICAQAQl4rDXANENt3ptK13
# 2855UU0BsS50cVttDBOrzr57j7gu1BKijG1iuFcCy04gE1CZ3XpA4le7r1iaHOEd
# AYasu3jyi9DsOwHu4r6PCgXIjUji8FMV3U+rkuTnjWrVgMHmlPIGL4UD6ZEqJCJw
# +/b85HiZLg33B+JwvBhOnY5rCnKVuKE5nGctxVEO6mJcPxaYiyA/4gcaMvnMMUp2
# MT0rcgvI6nA9/4UKE9/CCmGO8Ne4F+tOi3/FNSteo7/rvH0LQnvUU3Ih7jDKu3hl
# XFsBFwoUDtLaFJj1PLlmWLMtL+f5hYbMUVbonXCUbKw5TNT2eb+qGHpiKe+imyk0
# BncaYsk9Hm0fgvALxyy7z0Oz5fnsfbXjpKh0NbhOxXEjEiZ2CzxSjHFaRkMUvLOz
# sE1nyJ9C/4B5IYCeFTBm6EISXhrIniIh0EPpK+m79EjMLNTYMoBMJipIJF9a6lbv
# pt6Znco6b72BJ3QGEe52Ib+bgsEnVLaxaj2JoXZhtG6hE6a/qkfwEm/9ijJssv7f
# UciMI8lmvZ0dhxJkAj0tr1mPuOQh5bWwymO0eFQF1EEuUKyUsKV4q7OglnUa2ZKH
# E3UiLzKoCG6gW4wlv6DvhMoh1useT8ma7kng9wFlb4kLfchpyOZu6qeXzjEp/w7F
# W1zYTRuh2Povnj8uVRZryROj/TCCBoEwggRpoAMCAQICCmEVCCcAAAAAAAwwDQYJ
# KoZIhvcNAQEFBQAwXzETMBEGCgmSJomT8ixkARkWA2NvbTEZMBcGCgmSJomT8ixk
# ARkWCW1pY3Jvc29mdDEtMCsGA1UEAxMkTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNh
# dGUgQXV0aG9yaXR5MB4XDTA2MDEyNTIzMjIzMloXDTE3MDEyNTIzMzIzMloweTEL
# MAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1v
# bmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEjMCEGA1UEAxMaTWlj
# cm9zb2Z0IENvZGUgU2lnbmluZyBQQ0EwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAw
# ggEKAoIBAQCfjd+FN4yxBlZmNk7UCus2I5Eer6uNWOnEz8GfOgokxMTEXrDuFRTF
# +j6ZM2sZaXL0fAVf5ZklRNc1GYqQ3CiOkAzv1ZBhrd7cGHAtg8lvr4Us+N25uTD9
# cXgcg/3IqbmCZw16uMEJwrwWl1c/HJjTadcwkJCQjTAf2CbUnnuI2eIJ7ZdJResE
# UoF1e7i1IrguVrvXz6lOPAqDoqg6xa22AQ5qzyK0Ix9s1Sfnt37BtNUyrXklHEKG
# 4p2F9FfaG1kvLSaSKcWz14WjnmBalOZ7nHtegjRLbf/U7ifQotzRkAzOfQ4VfIis
# NMfAbJiESslEeWgo3yKDDbiKLEhh4v4RAgMBAAGjggIjMIICHzAQBgkrBgEEAYI3
# FQEEAwIBADAdBgNVHQ4EFgQUV0V0HF2w9shDBeCMVC2PMqf+SJYwCwYDVR0PBAQD
# AgHGMA8GA1UdEwEB/wQFMAMBAf8wgZgGA1UdIwSBkDCBjYAUDqyCYEBWJ5flJRP8
# KuEKU5VZ5KShY6RhMF8xEzARBgoJkiaJk/IsZAEZFgNjb20xGTAXBgoJkiaJk/Is
# ZAEZFgltaWNyb3NvZnQxLTArBgNVBAMTJE1pY3Jvc29mdCBSb290IENlcnRpZmlj
# YXRlIEF1dGhvcml0eYIQea0WoUqgpa1Mc1j0BxMuZTBQBgNVHR8ESTBHMEWgQ6BB
# hj9odHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2NybC9wcm9kdWN0cy9taWNy
# b3NvZnRyb290Y2VydC5jcmwwVAYIKwYBBQUHAQEESDBGMEQGCCsGAQUFBzAChjho
# dHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY3Jvc29mdFJvb3RD
# ZXJ0LmNydDB2BgNVHSAEbzBtMGsGCSsGAQQBgjcVLzBeMFwGCCsGAQUFBwICMFAe
# TgBDAG8AcAB5AHIAaQBnAGgAdAAgAKkAIAAyADAAMAA2ACAATQBpAGMAcgBvAHMA
# bwBmAHQAIABDAG8AcgBwAG8AcgBhAHQAaQBvAG4ALjATBgNVHSUEDDAKBggrBgEF
# BQcDAzANBgkqhkiG9w0BAQUFAAOCAgEAMLywIKRioKfvOSZhPdysxpnQhsQu9YMy
# ZV4iPpvWhvjotp/Ki9Y7dQuhkT5M3WR0jEnyiIwYZ2z+FWZGuDpGQpfIkTfUJLHn
# rNPqQRSDd9PJTwVfoxRSv5akLz5WWxB1zlPDzgVUabRlySSlD+EluBq5TeUCuVAe
# T7OYDB2VAu4iWa0iywV0CwRFewRZ4NgPs+tM+GDdwnie0bqfa/fz7n5EEUDSvbqb
# SxYIbqS+VeSmOBKjSPQcVXqKINF9/pHblI8vwntrpmSFT6PlLDQpXQu/9cc4L8Qg
# xFYx9mnOhfgKkezQ1q66OAUM625PTJwDKaqi/BigKQwNXFxWI1faHJYNyCY2wUTL
# 5eHmb4nnj+mYtXPTeOPtowE8dOVevGz2IYlnBeyXnbWx/a+m6XKlwzThL5/59Go5
# 4i0Eglv80JyufJ0R+ea1Uxl0ujlKOet9QrNKOzc9wkp7J5jn4k6bG0pUOGojN75q
# t0ju6kINSSSRjrcELpdv5OdFu49N/WDZ11nC2IDWYDR7t6GTIP6BuKqlXAnpig2+
# KE1+1+gP7WV40TFfuWbb30LnC8wCB43f/yAGo0VltLMyjS6R4k20qcn6vGsEDrKf
# 6p/epMkKlvSN99iYqPCFAghZpCCmLAsa8lIG7WnlZBgb4KOr3sp8FGFDuGX1NqNV
# EytnLE0bMEwxggSRMIIEjQIBATCBhzB5MQswCQYDVQQGEwJVUzETMBEGA1UECBMK
# V2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0
# IENvcnBvcmF0aW9uMSMwIQYDVQQDExpNaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBD
# QQIKYRr16gAAAAAAajAJBgUrDgMCGgUAoIG+MBkGCSqGSIb3DQEJAzEMBgorBgEE
# AYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJ
# BDEWBBQJuJy1aMlyYBrOieoYLjfyHYcaCDBeBgorBgEEAYI3AgEMMVAwTqAmgCQA
# UgBlAG0AbwB0AGUARQB4AGMAaABhAG4AZwBlAC4AcABzADGhJIAiaHR0cDovL3d3
# dy5taWNyb3NvZnQuY29tL2V4Y2hhbmdlIDANBgkqhkiG9w0BAQEFAASCAQCJ9XEP
# J8FQxrTqD9REXO/UFXVFtolPLeVXB+zfPt8KiI1VZZVfQpRiP9+cQTmXvl8pV80y
# xrNLQNwHV5wn60R3eXUwnszcJQkbPoeX4/SmUeIw6ppF8UBHwgsD2Or+ddLEf0c/
# V3ofOfdWVfMXq5UA2ygBNOt7hL4BqXJ/DXTq5E6bRYBpegIedaBWaMgl1xwvAQQL
# JLNZnlKsFQeBqNTPNlIZZXVekcxZt1g4rhIDA1Hqq/VCueIVzQOg+6Olj4a27bqd
# bFkoD+ttNT1MtU5qmeLPDSWsuYZLLeFntGX82FQMpOjzN0acFVFAxcb7IQOMio+4
# o/GzAeoMfJ+vHrryoYICHTCCAhkGCSqGSIb3DQEJBjGCAgowggIGAgEBMIGFMHcx
# CzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRt
# b25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xITAfBgNVBAMTGE1p
# Y3Jvc29mdCBUaW1lLVN0YW1wIFBDQQIKYQUTNgAAAAAAGjAHBgUrDgMCGqBdMBgG
# CSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTExMTEyMzIx
# MDYyMFowIwYJKoZIhvcNAQkEMRYEFMQ8uMyOzQuDPlQQXvSC6th0nrcQMA0GCSqG
# SIb3DQEBBQUABIIBAI0jNgNFxmH7ufkd2xi52CaOqounDj60Uvo58158Egzcd3Q7
# DW9Hqe2stJG7u8Rcn5Mhx1ptYhLHzqTX50FwfkguPzOSDoXU0dSOuwmxPXXt6Vgy
# vfI3Ot8p49wp+8jqCM4rMGplczAwWBciVkb3/4GqeTugBA3qfZnOov7P7iLYtHJh
# +NW/PQ+/ZuIpa1RX8raxSiYH9UeeM+Po7mREK8mB/Mf8Wkkh00hEGbM7wocpLyvu
# Fjanwx1xElwDBdhoj+Iz8IPA+YBJbu8cK3+1U4fbZQMavVoo3WibDi4+rr4loDKJ
# 568wE6KsqM5+J6TCtJM/YQeyoBWL3J6upjsl/qQ=
# SIG # End signature block
