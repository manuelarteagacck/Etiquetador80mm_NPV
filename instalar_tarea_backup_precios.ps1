param(
    [string]$TaskName = "NPV Backup Precios Venta",
    [string]$ExePath = "",
    [string]$PythonPath = "C:\Users\manuel.arteaga\AppData\Local\Programs\Python\Python314\python.exe"
)

$ErrorActionPreference = "Stop"

$ScriptPath = Join-Path $PSScriptRoot "backup_precios_venta.py"
$DefaultExePath = Join-Path $PSScriptRoot "BackupPreciosVenta.exe"

if (-not $ExePath) {
    $ExePath = $DefaultExePath
}

if (Test-Path -LiteralPath $ExePath) {
    $Execute = $ExePath
    $Argument = "--apply"
} else {
    if (-not (Test-Path -LiteralPath $ScriptPath)) {
        throw "No existe $ExePath ni el script Python: $ScriptPath"
    }

    if (-not (Test-Path -LiteralPath $PythonPath)) {
        $PythonCommand = Get-Command python -ErrorAction SilentlyContinue
        if ($null -eq $PythonCommand) {
            throw "No se encontro Python. Copia BackupPreciosVenta.exe junto a este instalador o indica -PythonPath."
        }
        $PythonPath = $PythonCommand.Source
    }

    $Execute = $PythonPath
    $Argument = "`"$ScriptPath`" --apply"
}

$Action = New-ScheduledTaskAction `
    -Execute $Execute `
    -Argument $Argument `
    -WorkingDirectory $PSScriptRoot

$Trigger = New-ScheduledTaskTrigger -Daily -At 21:00

$Settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -MultipleInstances IgnoreNew `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 30)

$Description = "Detecta precios futuros nuevos y reemplaza el respaldo diario NPV.dbo.NPVFDPreciosVentaBkp."

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $Action `
    -Trigger $Trigger `
    -Settings $Settings `
    -Description $Description `
    -Force | Out-Null

Write-Host "Tarea programada creada/actualizada: $TaskName"
Write-Host "Horario: diario 21:00"
Write-Host "Accion: $Execute $Argument"
