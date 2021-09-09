<#***********************************************************************

Get-EXORecipientRate.ps1
Modified 2021/09/09
Last Modifier:  Jim Martin
Project Owner:  Jim Martin
Version: v1.0

    .SYNOPSIS
    This script aggregates message trace events hourly and generates a report with the top o365 senders and the number of recipients they sent messages to in the past x number of days/hours

    .DESCRIPTION

    .PARAMETER StartDate
     The StartDate parameter specifies the end date of the date range

    .PARAMETER EndDate
     The EndDate parameter specifies the end date of the date range. It is recommended to limit the start-end date range to a range of hours. i.e. an ~ 5 to 7 hours.

    .PARAMETER TimeoutAfter
     The TimeoutAfter parameter specifies the number of minutes before the script stop working.
     This is to make sure that the script does not run for infinity. The default value is 30 minutes.

     .PARAMETER Sender
      The Sender parameter specifies the sender to query for the number of recipients.
      
    .EXAMPLE

    $results = .\Get-EXORecipientRate -StartDate (Get-Date).AddHours(-24) -EndDate (Get-Date) -Sender jim@contoso.com
    Shows the number recipients that the sender sent over the past 24 hours

    $results = .\Get-EXORecipientRate -StartDate (Get-Date).AddDays(-24) -EndDate (Get-Date)
    Shows the top 10 senders and the number of recipients they sent messages to in the past 24 hours

    .OUTPUTS
     $results.TopSenders : hourly report for top recipients over the threshold
     $results.HourlyReport  : hourly aggregated message events without applying the threshold
     $results.GroupReport : list of distribution groups for further investigation if needed


***********************************************************************

Copyright (c) 2018 Microsoft Corporation. All rights reserved.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.

**********************************************************************​
#>

[CmdletBinding()]
param
(
    [Parameter(Mandatory=$true)]
    [DateTime]$StartDate,
    [Parameter(Mandatory=$true)]
    [DateTime]$EndDate,
    [Parameter(Mandatory=$false)]
    [int]$TimeoutAfter = 30,
    [Parameter(Mandatory=$false)]
    [string]$SenderAddress
)

#Main
if($StartDate -lt (Get-Date).AddDays(-10)) {
    Write-Warning "Message trace data is only available for the past 10 days. Please update the StartDate and try again."
    exit
}
$NumberOfHours = ($EndDate - $StartDate).TotalHours
$eventList = New-Object -TypeName "System.Collections.ArrayList"
$hourlyReport = New-Object -TypeName "System.Collections.ArrayList"
$groupReport = New-Object -TypeName "System.Collections.ArrayList"
$UserReport = New-Object -TypeName "System.Collections.ArrayList"

## Get message trace data using input from command
for($x=0; $x -le $NumberOfHours; $x++) {
    $EndDate = $StartDate.AddHours(1)
    $pc = ($x/$NumberOfHours)*100 
    Write-Progress -Activity "Getting message trace information up to $($EndDate)..." -Status "Please wait." -PercentComplete $pc
    [int]$page=1
    [DateTime]$timeout = (Get-Date).AddMinutes($TimeoutAfter)
    Do {
        $pageList = New-Object -TypeName "System.Collections.ArrayList"
        if($SenderAddress -notlike $null) {$pageList = Get-MessageTrace -StartDate $StartDate -EndDate $EndDate -Page $page -PageSize 5000 -SenderAddress $SenderAddress}
        else {$pageList = Get-MessageTrace -StartDate $StartDate -EndDate $EndDate -Page $page -PageSize 5000}
        $eventList += $pageList
        $page++
    }While ($pageList.count -eq 5000 -and (Get-Date) -lt $timeout)
    $StartDate = $EndDate
}

# Get hourly count for number of recipients by sender
$eventList = $eventList | Sort-Object SenderAddress, Received
$eventList.foreach(
    {
        $hourlyEvent = $hourlyReport[-1] #data is sorted to min get operations. we only need to compare with last element in the array
        if($_.Status -eq "Delivered" -or $_.Status -eq "Failed"){
            if ($hourlyEvent.SenderAddress -eq $_.SenderAddress -and $hourlyEvent.Hour -eq $_.Received.Hour ) {
                $hourlyEvent.RecipientCount+=1
            } else {
                $eventObj = New-Object PSObject -Property @{ Date=$_.Received.Date.ToString(“MM-dd-yyyy”); Hour=$_.Received.Hour;SenderAddress=$_.SenderAddress; RecipientCount=1; Status=$_.Status };
                [void]$hourlyReport.Add($eventObj)
            }
        }
    }
)

# Get hourly count for number of recipients by sender
$eventList = $eventList | Sort-Object SenderAddress, Received
$eventList.foreach(
    {
        if ($_.Status -eq "Expanded") {
            $eventObj = New-Object PSObject -Property @{ Date=$_.Received.DateTime; SenderAddress=$_.SenderAddress; Recipient=$_.RecipientAddress; Status=$_.Status; MessageTraceId=$_.MessageTraceId };
            [void]$groupReport.Add($eventObj)
        }
    }
)

# Get total number of recipients by sender
$hourlyReport = $hourlyReport | Sort-Object SenderAddress
if($hourlyReport.Count -gt 1) {
    $hourlyReport.ForEach( 
        {  $userEvent = $userReport[-1]
            if($userEvent.SenderAddress -eq $_.SenderAddress) {
                $userEvent.RecipientCount = $userEvent.RecipientCount+$_.RecipientCount
            } else {
                $eventObj = New-Object PSObject -Property @{ SenderAddress=$_.SenderAddress; RecipientCount=$_.RecipientCount };
                [void]$userReport.Add($eventObj)
            }
        }
    )
} else {
    $eventObj = New-Object PSObject -Property @{ SenderAddress=$hourlyReport.SenderAddress; RecipientCount=$hourlyReport.RecipientCount };
    [void]$userReport.Add($eventObj)
}

$props = [ordered]@{
    'TopSenders'      = $UserReport | Sort-Object RecipientCount -Descending | Select -First 10
    'HourlyReport'       = $hourlyReport | Sort-Object Date, Hour
    'GroupReport' = $groupReport | Sort-Object Date
#    'MessageTraceEvents' = $eventList
}
$results = New-Object -TypeName PSObject -Property $props;
return $results



# SIG # Begin signature block
# MIIFvQYJKoZIhvcNAQcCoIIFrjCCBaoCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCgx9JkImlVHhRH
# e+YEqHuIryGHKWw/9KeX8+bYh24SDqCCAzYwggMyMIICGqADAgECAhA8ATOaNhKD
# u0LkWaETEtc0MA0GCSqGSIb3DQEBCwUAMCAxHjAcBgNVBAMMFWptYXJ0aW5AbWlj
# cm9zb2Z0LmNvbTAeFw0yMTAzMjYxNjU5MDdaFw0yMjAzMjYxNzE5MDdaMCAxHjAc
# BgNVBAMMFWptYXJ0aW5AbWljcm9zb2Z0LmNvbTCCASIwDQYJKoZIhvcNAQEBBQAD
# ggEPADCCAQoCggEBAMSWhFMKzV8qMywbj1H6lg4h+cvR9CtxmQ1J3V9uf9+R2d9p
# laoDqCNS+q8wz+t+QffvmN2YbcsHrXp6O7bF+xYjuPtIurv8wM69RB/Uy1xvsUKD
# L/ZDQZ0zewMDLb5Nma7IYJCPYelHiSeO0jsyLXTnaOG0Rq633SUkuPv+C3N8GzVs
# KDnxozmHGYq/fdQEv9Bpci2DkRTtnHvuIreeqsg4lICeTIny8jMY4yC6caQkamzp
# GcJWWO0YZlTQOaTgHoVVnSZAvdJhzxIX2wqd0/VaVIbpN0HcPKtMrgXv0O2Bl4Lo
# tmZR7za7H6hamxaPYQHHyReFs2xM7hlVVWhnfpECAwEAAaNoMGYwDgYDVR0PAQH/
# BAQDAgeAMBMGA1UdJQQMMAoGCCsGAQUFBwMDMCAGA1UdEQQZMBeCFWptYXJ0aW5A
# bWljcm9zb2Z0LmNvbTAdBgNVHQ4EFgQUCB04A8myETdoRJU9zsScvFiRGYkwDQYJ
# KoZIhvcNAQELBQADggEBAEjsxpuXMBD72jWyft6pTxnOiTtzYykYjLTsh5cRQffc
# z0sz2y+jL2WxUuiwyqvzIEUjTd/BnCicqFC5WGT3UabGbGBEU5l8vDuXiNrnDf8j
# zZ3YXF0GLZkqYIZ7lUk7MulNbXFHxDwMFD0E7qNI+IfU4uaBllsQueUV2NPx4uHZ
# cqtX4ljWuC2+BNh09F4RqtYnocDwJn3W2gdQEAv1OQ3L6cG6N1MWMyHGq0SHQCLq
# QzAn5DpXfzCBAePRcquoAooSJBfZx1E6JeV26yw2sSnzGUz6UMRWERGPeECSTz3r
# 8bn3HwYoYcuV+3I7LzEiXOdg3dvXaMf69d13UhMMV1sxggHdMIIB2QIBATA0MCAx
# HjAcBgNVBAMMFWptYXJ0aW5AbWljcm9zb2Z0LmNvbQIQPAEzmjYSg7tC5FmhExLX
# NDANBglghkgBZQMEAgEFAKB8MBAGCisGAQQBgjcCAQwxAjAAMBkGCSqGSIb3DQEJ
# AzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8G
# CSqGSIb3DQEJBDEiBCAVXQVIFLLTSe1dQRIjDAmEcHOnO9MBvrQLzLxLaPArMzAN
# BgkqhkiG9w0BAQEFAASCAQAwuPVGdWALyrQqDgcZMuc+EoOK53J31+UJfrVA3gHy
# ufcnc2SrW1ZUYHBxthOMCPA069jWhyq81chs6baCnqGr71Rq5HDbYYpkXCD5PTtT
# ST+zqcpXHlirz8kIIyBL0UgEhUhxqmxC0YgmA+sfZNrUr8EwsVv4fpllXooOIiNi
# B0vswQGKSsW8EM0ujaudjewMqwPbCFP6zfckt3doicZB8NGUQPgr5TRTNyab7hMS
# aBrifCcHIBRxiJ010dP1wXeNIz9u/6Q8czohZgMkaaN8WTIX8zosSNM6SMkmYokE
# 0jyxshW7NIkZ6b3QqhOQKAe0Y7L/lUBPez9G7TL/MjuV
# SIG # End signature block
