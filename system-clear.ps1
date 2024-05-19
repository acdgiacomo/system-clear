function Test-Administrator {
    $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Set-ExecutionPolicyRemoteSigned {
    Set-ExecutionPolicy RemoteSigned -Scope Process -Force
}

function Restore-ExecutionPolicy {
    param (
        [string]$originalPolicy
    )
    Set-ExecutionPolicy $originalPolicy -Scope Process -Force
}

function HasUnsavedNotepadDocuments {
    $notepadWindows = Get-Process notepad -ErrorAction SilentlyContinue | ForEach-Object {
        $_.MainWindowHandle
    }

    foreach ($windowHandle in $notepadWindows) {
        $windowTitle = (Get-Process -Id (Get-Process notepad -ErrorAction SilentlyContinue).Id).MainWindowTitle
        if ($windowTitle -match '\*') {
            return $true
        }
    }
    return $false
}

function HasUnsavedWordDocuments {
    $wordApp = New-Object -ComObject Word.Application -ErrorAction SilentlyContinue
    if ($null -ne $wordApp) {
        foreach ($doc in $wordApp.Documents) {
            if ($doc.Saved -eq $false) {
                $wordApp.Quit()
                return $true
            }
        }
        $wordApp.Quit()
    }
    return $false
}

function HasUnsavedExcelDocuments {
    $excelApp = New-Object -ComObject Excel.Application -ErrorAction SilentlyContinue
    if ($null -ne $excelApp) {
        foreach ($wb in $excelApp.Workbooks) {
            if ($wb.Saved -eq $false) {
                $excelApp.Quit()
                return $true
            }
        }
        $excelApp.Quit()
    }
    return $false
}

function HasUnsavedPowerPointDocuments {
    $powerPointApp = New-Object -ComObject PowerPoint.Application -ErrorAction SilentlyContinue
    if ($null -ne $powerPointApp) {
        foreach ($presentation in $powerPointApp.Presentations) {
            if ($presentation.Saved -eq $msoFalse) {
                $powerPointApp.Quit()
                return $true
            }
        }
        $powerPointApp.Quit()
    }
    return $false
}

function ClearTempFolder {
    param (
        [bool]$executeDirectly
    )
    if ($executeDirectly) {
        $response = 's'
    } else {
        $response = Read-Host "`nDeseja limpar a pasta Temp do usuário atual? Isso removerá arquivos temporários. (s/n)"
    }

    if ($response -eq 's') {
        $tempPath = [System.IO.Path]::GetTempPath()
        Write-Host "`nLimpando a pasta Temp: $tempPath"
        Remove-Item -Path "$tempPath*" -Force -Recurse -ErrorAction SilentlyContinue
        Write-Host "Pasta Temp limpa."
    } else {
        Write-Host "`nPasta Temp do usuário atual não será limpa."
    }
}

function ClearSystemTempFolder {
    param (
        [bool]$executeDirectly
    )
    if ($executeDirectly) {
        $response = 's'
    } else {
        $response = Read-Host "`nDeseja limpar a pasta Temp do sistema? Isso removerá arquivos temporários do sistema. (s/n)"
    }

    if ($response -eq 's') {
        $systemTempPath = "C:\Windows\Temp"
        Write-Host "`nLimpando a pasta Temp do sistema: $systemTempPath"
        Remove-Item -Path "$systemTempPath\*" -Force -Recurse -ErrorAction SilentlyContinue
        Write-Host "Pasta Temp do sistema limpa."
    } else {
        Write-Host "`nPasta Temp do sistema não será limpa."
    }
}

function EmptyRecycleBin {
    param (
        [bool]$executeDirectly
    )
    if ($executeDirectly) {
        $response = 's'
    } else {
        $response = Read-Host "`nDeseja esvaziar a Lixeira? Isso removerá permanentemente os arquivos na Lixeira. (s/n)"
    }

    if ($response -eq 's') {
        Write-Host "`nEsvaziando a Lixeira"
        Clear-RecycleBin -Force -ErrorAction SilentlyContinue
        Write-Host "Lixeira esvaziada."
    } else {
        Write-Host "`nLixeira não será esvaziada."
    }
}

function ClearBrowserCaches {
    param (
        [bool]$executeDirectly
    )

    $browsers = @(
        @{
            Name = "Google Chrome"
            Path = "$env:LocalAppData\Google\Chrome\User Data\Default\Cache"
        },
        @{
            Name = "Microsoft Edge"
            Path = "$env:LocalAppData\Microsoft\Edge\User Data\Default\Cache"
        },
        @{
            Name = "Mozilla Firefox"
            Path = "$env:LocalAppData\Mozilla\Firefox\Profiles"
            ProfileSubfolder = "cache2"
        },
        @{
            Name = "Opera"
            Path = "$env:LocalAppData\Opera Software\Opera Stable\Cache"
        },
        @{
            Name = "Opera GX"
            Path = "$env:LocalAppData\Opera Software\Opera GX Stable\Cache"
        },
        @{
            Name = "Brave"
            Path = "$env:LocalAppData\BraveSoftware\Brave-Browser\User Data\Default\Cache"
        }
    )

    if ($executeDirectly) {
        $response = 's'
    } else {
        $response = Read-Host "`nDeseja limpar o cache dos navegadores instalados? Isso removerá arquivos temporários da internet. (s/n)"
    }

    if ($response -eq 's') {
        foreach ($browser in $browsers) {
            if (Test-Path $browser.Path) {
                if ($browser.Name -eq "Mozilla Firefox") {
                    $firefoxProfiles = Get-ChildItem -Path $browser.Path -Directory
                    foreach ($profile in $firefoxProfiles) {
                        $cachePath = "$browser.Path\$profile\$($browser.ProfileSubfolder)"
                        if (Test-Path $cachePath) {
                            Write-Host "`nLimpando o cache do $($browser.Name) no perfil $($profile.Name)"
                            Remove-Item -Path "$cachePath\*" -Force -Recurse -ErrorAction SilentlyContinue
                            Write-Host "Cache do $($browser.Name) limpo para o perfil $($profile.Name)."
                        }
                    }
                } else {
                    Write-Host "`nLimpando o cache do $($browser.Name) em $($browser.Path)"
                    Remove-Item -Path "$($browser.Path)\*" -Force -Recurse -ErrorAction SilentlyContinue
                    Write-Host "Cache do $($browser.Name) limpo."
                }
            }
        }
    } else {
        Write-Host "`nCaches dos navegadores não serão limpos."
    }
}

function ClearWindowsUpdateTempFiles {
    param (
        [bool]$executeDirectly
    )
    if ($executeDirectly) {
        $response = 's'
    } else {
        $response = Read-Host "`nDeseja limpar os arquivos temporários do Windows Update? Isso removerá os arquivos baixados de atualizações. (s/n)"
    }

    if ($response -eq 's') {
        Write-Host "`nLimpando arquivos temporários do Windows Update"
        Remove-Item -Path "C:\Windows\SoftwareDistribution\Download\*" -Force -Recurse -ErrorAction SilentlyContinue
        Write-Host "Arquivos temporários do Windows Update limpos."
    } else {
        Write-Host "`nArquivos temporários do Windows Update não serão limpos."
    }
}

function ClearWindowsUpdateLogs {
    param (
        [bool]$executeDirectly
    )
    if ($executeDirectly) {
        $response = 's'
    } else {
        $response = Read-Host "`nDeseja limpar os logs do Windows Update? Isso removerá os registros de atualizações passadas. (s/n)"
    }

    if ($response -eq 's') {
        Write-Host "`nLimpando logs do Windows Update"
        Remove-Item -Path "C:\Windows\Logs\WindowsUpdate\*" -Force -Recurse -ErrorAction SilentlyContinue
        Write-Host "Logs do Windows Update limpos."
    } else {
        Write-Host "`nLogs do Windows Update não serão limpos."
    }
}

function DiskCleanup {
    param (
        [bool]$executeDirectly
    )
    if ($executeDirectly) {
        $response = 's'
    } else {
        $response = Read-Host "`nDeseja executar a limpeza de disco? Isso removerá arquivos desnecessários e temporários. (s/n)"
    }

    if ($response -eq 's') {
        Write-Host "`nExecutando limpeza de disco"
        Cleanmgr /sagerun:1
        Write-Host "Limpeza de disco executada."
    } else {
        Write-Host "`nLimpeza de disco não será executada."
    }
}

function DisableFastStartup {
    param (
        [bool]$executeDirectly
    )
    if ($executeDirectly) {
        $response = 's'
    } else {
        $response = Read-Host "`nDeseja desabilitar a Inicialização Rápida? Isso pode melhorar o desempenho do sistema. (s/n)"
    }

    if ($response -eq 's') {
        Write-Host "`nDesabilitando Inicialização Rápida..."
        powercfg -h off
        reg add "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Power" /v HiberbootEnabled /t REG_DWORD /d 0 /f
        Write-Host "Inicialização Rápida desabilitada.`n"
    } else {
        Write-Host "`nInicialização Rápida não será desabilitada."
    }
}

function PerformSystemCleanup {
    param (
        [bool]$executeDirectly
    )
    ClearTempFolder -executeDirectly $executeDirectly
    ClearSystemTempFolder -executeDirectly $executeDirectly
    EmptyRecycleBin -executeDirectly $executeDirectly
    ClearBrowserCaches -executeDirectly $executeDirectly
    ClearWindowsUpdateTempFiles -executeDirectly $executeDirectly
    ClearWindowsUpdateLogs -executeDirectly $executeDirectly
    DiskCleanup -executeDirectly $executeDirectly
    DisableFastStartup -executeDirectly $executeDirectly
}

if (-not (Test-Administrator)) {
    Write-Error "`nERRO: Este script precisa ser executado como administrador. Por favor, execute novamente como administrador."
    pause
    exit
}

$executeAllDirectly = $false
if ((Read-Host "`nDeseja executar todas as funções diretamente sem solicitar confirmação para cada uma? (s/n)") -eq 's') {
    $executeAllDirectly = $true
}

$originalPolicy = Get-ExecutionPolicy -Scope Process

Set-ExecutionPolicyRemoteSigned

try {
    PerformSystemCleanup -executeDirectly $executeAllDirectly

    if (HasUnsavedNotepadDocuments -or HasUnsavedWordDocuments -or HasUnsavedExcelDocuments -or HasUnsavedPowerPointDocuments) {
        Write-Warning "`nATENÇÃO: Existem documentos não salvos. Verifique antes de autorizar a reinicialização."
    }

    $confirmation = Read-Host "`nDeseja reiniciar o sistema agora? (s/n)"
} catch {
    Write-Error "`nOcorreu um erro inesperado. Por favor entrar em contato com o administrador [anaccdg]. Erro: $_"
} finally {
    Write-Host "`nRestaurando a política de execução..."
    Restore-ExecutionPolicy -originalPolicy $originalPolicy
    Write-Host "Política de execução restaurada."
}

if ($confirmation -eq 's') {
    Write-Host "`nReiniciando o sistema..."
    Restart-Computer -Force
} else {
    Write-Host "`nO sistema não será reiniciado."
}

pause