$ErrorActionPreference = "SilentlyContinue"
#$RegPath = "HKLM:\SOFTWARE\Microsoft"
$Date = Get-Date
$LastCompliance = (Get-ItemProperty -Path $RegPath -ErrorAction SilentlyContinue).LastCompliance
$MissingUpdates = get-wmiobject -query "SELECT * FROM CCM_SoftwareUpdate" -namespace "ROOT\ccm\ClientSDK" | Where {$_.ExclusiveUpdate -eq $false} | select name
[int]$MissingCount = $MissingUpdates.Name.Count

# If we are not missing any updates. Set LastCompliance to todays date.
If ($MissingCount -eq 0) {
    # Set LastCompliance to todays date.
    #(New-ItemProperty -Path $RegPath -Name "LastCompliance" -Value $Date -PropertyType String -Force | Out-Null)
    Return $MissingCount
}
# The device has one or more updates pending.
Else {
    # Device has never been compliant.
    If ($LastCompliance -eq "Never" -or $LastCompliance -eq $null) {
        # Set LastCompliance to Never.
        (New-ItemProperty -Path $RegPath -Name "LastCompliance" -Value "Never" -PropertyType String -Force | Out-Null)
        # Refresh/reevaluate updates.
        (New-Object -ComObject Microsoft.CCM.UpdatesStore).RefreshServerComplianceState()
        ([wmiclass]'ROOT\ccm:SMS_Client').TriggerSchedule('{00000000-0000-0000-0000-000000000111}') | out-null
        # Return number of missing updates to baseline report.
        Return $MissingCount
    }
    # Device has not been compliant the last 20 days.
    ElseIf ($LastCompliance -lt ($Date).AddDays(-20)) {
            # Return amount of missing updates.
            Return $MissingCount
    }
    # Device has been compliant the last 20 days.
    Else {
            # Returning 0 to verify compliance (grace period).
            Return 0
    }
}