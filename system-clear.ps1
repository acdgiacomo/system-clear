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