#requires -Version 5.1
<#
  FaturaAktarim'i Windows Gorev Zamanlayicisi'na ekler: her N dakikada bir Gelen klasorunu kontrol eder.
    Zamanlayici_Kur.ps1                          -> 5 dakikada bir, siz oturum acikken calisir
    Zamanlayici_Kur.ps1 -Dakika 10               -> 10 dakikada bir
    Zamanlayici_Kur.ps1 -OturumKapaliykenDe      -> oturum kapaliyken de calisir (Windows sifrenizi sorar)
    Zamanlayici_Kur.ps1 -Kaldir                  -> gorevi siler
#>
param([int]$Dakika = 5, [switch]$OturumKapaliykenDe, [switch]$Kaldir)
$ErrorActionPreference = 'Stop'
$ad = 'FaturaAktarim'

if ($Kaldir) {
    Unregister-ScheduledTask -TaskName $ad -Confirm:$false
    Write-Host "Gorev silindi: $ad"; return
}

$script = Join-Path $PSScriptRoot 'FaturaAktarim.ps1'
$act = New-ScheduledTaskAction -Execute 'powershell.exe' `
        -Argument ('-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "{0}"' -f $script) `
        -WorkingDirectory $PSScriptRoot
$trg = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval (New-TimeSpan -Minutes $Dakika)
$set = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Hours 2)
$kullanici = [Security.Principal.WindowsIdentity]::GetCurrent().Name

if ($OturumKapaliykenDe) {
    $cred = Get-Credential -UserName $kullanici -Message 'Gorevin oturum kapaliyken calismasi icin Windows sifrenizi girin'
    Register-ScheduledTask -TaskName $ad -Action $act -Trigger $trg -Settings $set -Force `
        -User $cred.UserName -Password $cred.GetNetworkCredential().Password `
        -Description 'Gelen klasorundeki fatura ZIPlerini SQL tablosuna aktarir' | Out-Null
} else {
    Register-ScheduledTask -TaskName $ad -Action $act -Trigger $trg -Settings $set -Force -User $kullanici `
        -Description 'Gelen klasorundeki fatura ZIPlerini SQL tablosuna aktarir' | Out-Null
}
Write-Host "Gorev kuruldu: '$ad' - $Dakika dakikada bir, kullanici $kullanici"
Write-Host "Gorev Zamanlayici > Gorev Zamanlayici Kitapligi altinda gorebilirsiniz."
