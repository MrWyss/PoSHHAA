# PoSHHAA

A simple PowerShell Module that implements a Home Assistant Assist ChatBot.

```text
2023-02-02 10:18:34
You: turn on office spots

2023-02-02 10:18:40
PoSHHAA: Turned on office spots
Devices: Office Spots
```

## Features

- Interactive Prompt
- Secure permanent token storage
- TTS and ability to change voices
- :construction: Multi language support

> :information_source: instead of typing you can also use Windows built-in STT **WIN + H**

## Install

### Create a local PS Repository

Create a directory e.g. ``C:\LocalPSRepo`` \
Register the Repository:

```powershell
Register-PSRepository -Name LocalPSRepo -SourceLocation 'C:\LocalPSRepo' -ScriptSourceLocation 'C:\LocalPSRepo' -InstallationPolicy Trusted
```

### Download and Publish Module

Download sources e.g. ``git clone git@github.com:MrWyss/PoSHHAA.git`` let's assume it clones the Github repository to ``C:\Git\PoSHHAA`` \
Publish module into local Repository

```powershell
Publish-Module -Path 'C:\Git\PoSHHAA\' -Repository LocalPSRepo -NuGetApiKey 'anything will do'
```

### Install Module from local Repository

Install required modules ``Microsoft.PowerShell.SecretManagement`` and ``Microsoft.PowerShell.SecretStore``

```powershell
Install-Module Microsoft.PowerShell.SecretManagement -Scope CurrentUser
Install-Module Microsoft.PowerShell.SecretStore -Scope CurrentUser
```

Install ``PoSHHAA`` module

```powershell
Install-Module PoSHHAA -Scope CurrentUser -Repository LocalPSRepo
Import-Module PoSHHAA
```

## Setup & Configure

The module exports two functions and one alias

```text
ModuleType Version    Name                                ExportedCommands
---------- -------    ----                                ----------------
Script     0.0.4      PoSHHAA                             {Initialize-HAAChat, Start-HAAChat, haa}
```

``Start-HAAChat``
You can run ``Initialize-HAAChat`` or ``Start-HAAChat`` with its alias ``haa``. \
Or just run ``haa``, as ``haa`` / ``Start-HAAChat`` calls ``Initialize-HAAChat``.

1. During the first run, the script creates a new secret vault ``PoSHHAA``, which requires a _master_ password.
2. Next the two secrets will be created ``PoSHHAABaseURL`` and ``PoSHHAAToken``. You will need to provide the values for these. Get them from your home assistant instance.

```text
A password is now required for the local store configuration.
To complete the change please provide new password.
Enter password:
*********
Enter password again for verification:
*********
Vault PoSHHAA requires a password.
Enter password:
*********
Home Assistant Token: *************.....
Home Assistant URL: https://yourinstanace.ui.nabu.casa/
```

## Using

### First run

Simply run ``Start-HAAChat`` or the alias ``haa``. \
From time to time (3600s), you'll be asked to provide the Vaults password.

```text
Vault PoSHHAA requires a password.
Enter password:
```

### Help

```text
2023-02-02 10:15:54
You: help

2023-02-02 10:15:58
PoSHHAA: Help

Commands                                   HelpMessage
--------                                   -----------
help, [ENTER]                              Shows this help
exit, stop, quit                           to quit
change voice                               to change voices
ha version, ha info                        to check version
change language, select language, language to change language
Turn on the living room lights             try: Turn on the living room lights

2023-02-02 10:15:58
You:
```

### Turn on some lights

```text
2023-02-02 10:18:34
You: turn on office spots

2023-02-02 10:18:40
PoSHHAA: Turned on office spots
Devices: Office Spots
```

## Uninstall

To uninstall run:

```powershell
Remove-Module PoSHHAA
Uninstall-Module PoSHHAA -force
```

> :information_source: The Vault and its secrets will remain. So does the local PS repository.

## TODO

- [ ] Implement Home Assistant version check
- [ ] Change Language
- [ ] More translations
- [ ] a way to turn off TTS
- [ ] more chat-like prompt
- [ ] Publish to PSGallery
