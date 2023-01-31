#region Configure Language
Import-LocalizedData -BindingVariable msgTable -UICulture $Global:Language
#endregion

#region Globals
$Global:LoginParameters = @{}
$Global:Language        = "en-US"
$Global:NarratorVoice   = "Microsoft David Desktop" 
$Global:YouPrompt = {"`r`n$(Get-Date -Format "yyyy-MM-dd hh:mm:ss")`r`n$($msgTable.YouPrompt): "}
$Global:YouPromptColorAndBehaviour = @{
    ForegroundColor = "Green"
    NoNewline = $true
}
$Global:BotPrompt = {"`r`n$(Get-Date -Format "yyyy-MM-dd hh:mm:ss")`r`n$($msgTable.PoSHHAABotName): "}
$Global:BotPromptColorAndBehaviour = @{
    ForegroundColor = "Blue"
    NoNewline = $true
}

#region Setup Narrator
Add-Type -AssemblyName System.Speech
$Global:Narrator = New-Object -TypeName System.Speech.Synthesis.SpeechSynthesizer
$Global:Narrator.SelectVoice($Global:NarratorVoice)
#endregion

#endregion

#region Check if secret is already save, if not create Secrete Store and save the token
Function Initialize-HAAChat {
    $VaultName          = "PoSHHAA"
    $TokenSecretName    = "PoSHHAAToken"
    $BaseUrlSecretName  = "PoSHHAABaseURL"
    
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
    $Token = Get-Secret -Name $TokenSecretName -AsPlainText
    $BaseUrl = Get-Secret -Name $BaseUrlSecretName -AsPlainText
    $Global:LoginParameters = @{
        Uri             = "$BaseUrl/api/conversation/process"
        SessionVariable = 'Session'
        Method          = 'POST'
        Headers         = @{
            Authorization = "Bearer " + $Token
        }
    }
}
#endregion

#region Functions
Function Ask() {
    $PromptPrefix = Invoke-Command -ScriptBlock $Global:YouPrompt
    return $(Write-Host $PromptPrefix @Global:YouPromptColorAndBehaviour; Read-Host) 
}

Function Talk() {
    Param(
        # Text to output and speak
        [Parameter(Mandatory = $true)]
        [String]$Text,
        # Do not speak only write
        [Parameter(Mandatory = $false)]
        [Switch]$Mute
    )

    $PromptPrefix = Invoke-Command -ScriptBlock $Global:BotPrompt
    Write-Host $PromptPrefix @Global:BotPromptColorAndBehaviour
    Write-Host $Text
    if (-not $Mute) {
        $Global:Narrator.Speak($Text)
    }
}

Function FindOut() {
    Param(
        # Chat Prompt
        [Parameter(Mandatory = $true)]
        [String]$Intent
    )
    Try {
        $Body = @{
            "text"     = $Intent
            "language" = $Global:Language
        }
        $Response = Invoke-WebRequest @Global:LoginParameters -Body $Body
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
    # Check Languages available in Windows and Home Assistant
    $msgTable.NotImplemented
}
Function GetVersion() {
    # Check Script Version and Home Assistant Versions
    Talk -Text "`r`nPoSHHA: $((get-module -Name PoSHHAA).Version.ToString())`r`nHome Assistant: $($msgTable.NotImplemented)" -Mute
}
#endregion

#region Help Messages
$HelpMessages = @(
    [pscustomobject]@{Commands = "help, [ENTER]"; HelpMessage = $msgTable.HelpMsg}
    [pscustomobject]@{Commands = "exit, stop, quit"; HelpMessage = $msgTable.HelpQuitMsg},
    [pscustomobject]@{Commands = "change voice"; HelpMessage = $msgTable.HelpChangeVoicesMsg},
    [pscustomobject]@{Commands = "ha version, ha info"; HelpMessage = $msgTable.HelpCheckVersionMsg},
    [pscustomobject]@{Commands = "change language, select language, language"; HelpMessage = $msgTable.HelpChangeLanguageMsg}
    [pscustomobject]@{Commands = $msgTable.HelpExampleCmd; HelpMessage = $msgTable.HelpExampleMsg}
)
#endregion

#region Start-HAAChat (main function)
Function Start-HAAChat {
    [Alias("haa")]
    param()

    Initialize-HAAChat
    $quit = $false
    While ($quit -eq $false) {
        $Intent = Ask
    
        Switch -regex ($Intent) {
            # exit, stop, quit and bye
            '^(exit|stop|quit|bye)$' { 
                $quit = $true
                Talk -Text $msgTable.Bye
            }
            # change voice, change voices, select voice, select voices, voice, voices
            '^((change|select|^)(\s|^)voice(s|))$' {
                ChangeVoice
            }
            # change language, select language, language
            '^((change|select|^)(\s|^)language)$' {
                ChangeLanguage
            }
            # ha version, info version, check version, version
            '^((ha|info|check|^)(\s|^)version)$' {
                GetVersion
            }
            # help, info, hello, ?, [ENTER]
            '^(help|info|hello|\?|)$' {
                Talk -Text $msgTable.Help -Mute
                $($HelpMessages | Format-Table)
            }
            Default {
                FindOut -Intent $Intent
            }
        }
    }
}
#endregion

#region Module Setup
Export-ModuleMember -Alias "haa" -Function Start-HAAChat
Export-ModuleMember -Function @(
	'Start-HAAChat'
	'Initialize-HAAChat'
)
#endregion