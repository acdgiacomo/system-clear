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