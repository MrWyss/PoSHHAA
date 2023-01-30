#region Changes Here
$Language = "en-US"
$NarratorVoice = "Microsoft Hazel Desktop" 
#region

#region Configure Language
Import-LocalizedData -BindingVariable msgTable -UICulture $Language
#endregion

#region install secrets modules if not installed
if (-not (Get-Module Microsoft.PowerShell.SecretManagement -ListAvailable)) {
    Install-Module Microsoft.PowerShell.SecretManagement -Scope CurrentUser -Force
}
if (-not (Get-Module Microsoft.PowerShell.SecretStore -ListAvailable)) {
    Install-Module Microsoft.PowerShell.SecretStore -Scope CurrentUser -Force
}
#endregion


#region Check if secret is already save, if not create Secrete Store and save the token
$VaultName = "PoSHHAA"
$TokenSecretName = "PoSHHAAToken"
$BaseUrlSecretName = "PoSHHAABaseURL"

If (-not (Get-SecretVault -Name $VaultName -ErrorAction SilentlyContinue)) {
    Set-SecretStoreConfiguration -Scope CurrentUser -Authentication Password -PasswordTimeout 3600 -Confirm:$false
    Register-SecretVault -Name $VaultName -ModuleName Microsoft.PowerShell.SecretStore -DefaultVault
}

# Save bearer token into a secret in a vault
If (-not (Get-Secret -Name $TokenSecretName -ErrorAction SilentlyContinue)) {
    $HATokenSecret = Read-Host -assecurestring $msgTable.TokenPrompt
    Set-Secret -Name $TokenSecretName -Metadata @{Purpose = "Powershell Home Assistant Assist Bearer Token" } -Secret $HATokenSecret -Vault $VaultName
    #Get-SecretInfo -Name $TokenSecretName -Vault $VaultName | Select-Object Name, Metadata
    #Remove-Secret $TokenSecretName -Vault $VaultName
}

# Save home assistant url into a secret in a vault
If (-not (Get-Secret -Name $BaseUrlSecretName -ErrorAction SilentlyContinue)) {
    $HomeAssistantURLSecret = Read-Host $msgTable.URLPrompt
    Set-Secret -Name $BaseUrlSecretName -Metadata @{Purpose = "Powershell Home Assistant Assist BaseURL" } -Secret $($HomeAssistantURLSecret.TrimEnd('/')) -Vault $VaultName
    #Get-SecretInfo -Name $BaseUrlSecretName -Vault $VaultName | Select-Object Name, Metadata
    #Remove-Secret $BaseUrlSecretName -Vault $VaultName
}
#endregion

#region Setup rest endpoint
$Token = Get-Secret -Name $TokenSecretName -AsPlainText
$BaseUrl = Get-Secret -Name $BaseUrlSecretName -AsPlainText

$LoginParameters = @{
    Uri             = "$BaseUrl/api/conversation/process"
    SessionVariable = 'Session'
    Method          = 'POST'
    Headers         = @{
        Authorization = "Bearer " + $Token
    }
}
#endregion

#region Setup Narrator
Add-Type -AssemblyName System.Speech
$Narrator = New-Object -TypeName System.Speech.Synthesis.SpeechSynthesizer
$Narrator.SelectVoice($NarratorVoice)
#endregion

#region Functions
Function Talk() {
    Param(
        # Text to output and speak
        [Parameter(Mandatory = $true)]
        [String]$Text,
        # Do not speak only write
        [Parameter(Mandatory = $false)]
        [Switch]$Mute
    )

    Write-Host "$($msgTable.PoSHHAABotName): $Text" -ForegroundColor Blue
    if (-not $Mute) {
        $Narrator.Speak($Text)
    }
}

Function Chat() {
    Param(
        # Chat Prompt
        [Parameter(Mandatory = $true)]
        [String]$Intent
    )
    Try {
        $Body = @{
            "text"     = $Intent
            "language" = $Language
        }
        $Response = Invoke-WebRequest @LoginParameters -Body $Body
        Talk($Response)
    }
    Catch {
        Talk("$($msgTable.Error) $($_.Exception.Response.StatusCode.Value__)") -Mute
    }
    Finally {
        #$quit = $true
        #Talk("exiting")
    }
}

Function ChangeVoice() {
    $Narrator.GetInstalledVoices().VoiceInfo | Select-Object Name, Culture, Gender, Description | Format-Table
    $NewVoice = $(Write-Host "$($msgTable.NewVoiceNamePrompt): " -ForegroundColor DarkBlue -NoNewLine; Read-Host) 
    try {
        $Narrator.SelectVoice($NewVoice)
        Talk -Text "$($msgTable.WelcomeMsg) $($Narrator.Voice.Name)"
    }
    catch {
        Talk -Text $msgTable.NoVoiceInstalled
    }
}
Function ChangeLanguage() {
    $msgTable.NotImplemented
}
Function GetVersion() {
    $msgTable.NotImplemented
}

#endregion

$quit = $false
While ($quit -eq $false) {
    $Intent = $(Write-Host "$($msgTable.YouPrompt): " -ForegroundColor Green -NoNewLine; Read-Host) 

    Switch -regex ($Intent) {
        '^(exit|stop|quit|bye)$' { 
            $quit = $true
            Talk -Text $msgTable.Bye
        }
        '^(change\svoices|change\svoice|select\svoices|select\svoice|voices?|voice)$' {
            ChangeVoice
        }
        '^(change\slanguage|select\slanguage|language)$' {
            ChangeLanguage
        }
        '^(ha\sversion|ha\sinfo|check\sversion|version)$' {
            GetVersion
        }
        '^(help|info|\?|hello|)$' {
            "$($msgTable.HelpQuitMsg): exit, stop, quit"
            "$($msgTable.HelpChangeVoicesMsg): change voice"
            "$($msgTable.HelpCheckVersionMsg): ha version, ha info"
            "$($msgTable.HelpChangeLanguageMsg): change language, select language, language"
            "$($msgTable.HelpExampleMsg)"
        }
        Default {
            Chat -Intent $Intent
        }
    }
}