Set-ExecutionPolicy Bypass -Scope Process -Force
$officeVersions = @("12.0", "14.0", "15.0", "16.0")
ForEach ($officeVersion in $officeVersions) {
    $regPath = "HKCU:\Software\Microsoft\Office\$officeVersion"
    $wordRegPath = "$regPath\Word\Security"
    $excelRegPath = "$regPath\Excel\Security"
    $powerPointRegPath = "$regPath\PowerPoint\Security"
    $outlookRegPath = "$regPath\Outlook\Security"
    $publisherRegPath = "$regPath\Publisher\Security"
    $onenoteRegPath = "$regPath\OneNote\Security" 
    $appsRegPaths = @($wordRegPath, $excelRegPath, $powerPointRegPath, $outlookRegPath, $publisherRegPath, $onenoteRegPath)

    ForEach ($appRegPath in $appsRegPaths) {
        if (Test-Path $appRegPath) {
            New-ItemProperty -Path $appRegPath -Name "VBAWarnings" -Value 1 -PropertyType DWORD -Force
            New-ItemProperty -Path $appRegPath -Name "AccessVBOM" -Value 1 -PropertyType DWORD -Force
            New-ItemProperty -Path $appRegPath -Name "DisableAttachementsInPV" -Value 1 -PropertyType DWORD -Force
            New-ItemProperty -Path $appRegPath -Name "DisableInternetFilesInPV" -Value 1 -PropertyType DWORD -Force
            New-ItemProperty -Path $appRegPath -Name "DisableUnsafeLocationsInPV" -Value 1 -PropertyType DWORD -Force
            New-ItemProperty -Path $appRegPath -Name "BlockContentExecutionFromInternet" -Value 1 -PropertyType DWORD -Force
            New-ItemProperty -Path $appRegPath -Name "DisableSafeMode" -Value 1 -PropertyType DWORD -Force
            New-ItemProperty -Path $appRegPath -Name "LowRiskFileTypes" -Value ".*" -PropertyType String -Force
            New-ItemProperty -Path $appRegPath -Name "DisableTrustCenterUI" -Value 1 -PropertyType DWORD -Force
            New-ItemProperty -Path $appRegPath -Name "TrustBarNotificationOnLoad" -Value 0 -PropertyType DWORD -Force
            New-ItemProperty -Path $appRegPath -Name "TrustBarNotificationOnMacro" -Value 0 -PropertyType DWORD -Force
            New-ItemProperty -Path $appRegPath -Name "TrustAccessVBOM" -Value 0 -PropertyType DWORD -Force
            New-ItemProperty -Path $appRegPath -Name "EnableUnsafeFilesInPV" -Value 1 -PropertyType DWORD -Force
            New-ItemProperty -Path $appRegPath -Name "Disable<mark>Office</mark>FileValidation" -Value 1 -PropertyType DWORD -Force
            New-ItemProperty -Path $appRegPath -Name "DisableDocInspector" -Value 1 -PropertyType DWORD -Force
        }
    }
}
