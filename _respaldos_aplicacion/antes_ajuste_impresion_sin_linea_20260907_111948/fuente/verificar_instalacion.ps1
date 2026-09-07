$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$installRoot = "D:\Etiquetador80mm"
$results = [ordered]@{}

$artifactPairs = [ordered]@{
    "Etiquetador80mm.exe" = @(
        (Join-Path $projectRoot ".installer_payload\Etiquetador80mm.exe"),
        (Join-Path $installRoot "Etiquetador80mm.exe")
    )
    "BackupPreciosVenta.exe" = @(
        (Join-Path $projectRoot ".installer_payload\BackupPreciosVenta.exe"),
        (Join-Path $installRoot "BackupPreciosVenta.exe")
    )
    "Video_Funcionamiento_Etiquetador80mm.mp4" = @(
        (Join-Path $projectRoot "presentacion_publica\Recorrido_Etiquetador80mm_CirculoK_v7_actualizado.mp4"),
        (Join-Path $installRoot "Video_Funcionamiento_Etiquetador80mm.mp4")
    )
    "Instalador_Etiquetador80mm.exe" = @(
        (Join-Path $projectRoot ".installer_build\Instalador_Etiquetador80mm.exe"),
        (Join-Path $installRoot "Instalador_Etiquetador80mm.exe")
    )
}

$artifactChecks = foreach ($name in $artifactPairs.Keys) {
    $source, $installed = $artifactPairs[$name]
    [pscustomobject]@{
        Name = $name
        Exists = Test-Path -LiteralPath $installed -PathType Leaf
        HashMatches = (
            (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash -eq
            (Get-FileHash -LiteralPath $installed -Algorithm SHA256).Hash
        )
        Bytes = (Get-Item -LiteralPath $installed).Length
    }
}
$results.Artifacts = @($artifactChecks)

$preservedConfig = foreach ($name in @("config.json", "config.key", "paramconf.json")) {
    $path = Join-Path $installRoot $name
    [pscustomobject]@{
        Name = $name
        Exists = Test-Path -LiteralPath $path -PathType Leaf
        Bytes = (Get-Item -LiteralPath $path).Length
    }
}
$results.Configuration = @($preservedConfig)

$shell = New-Object -ComObject WScript.Shell
$menuRoot = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Etiquetador80mm"
$menuChecks = foreach ($name in @(
    "Etiquetador80mm.lnk",
    "Video de funcionamiento.lnk",
    "Desinstalar Etiquetador80mm.lnk"
)) {
    $path = Join-Path $menuRoot $name
    $shortcut = $shell.CreateShortcut($path)
    [pscustomobject]@{
        Name = $name
        Exists = Test-Path -LiteralPath $path -PathType Leaf
        Target = $shortcut.TargetPath
        Arguments = $shortcut.Arguments
    }
}
$results.StartMenu = @($menuChecks)

$excludedProfiles = @("all users", "default user", "public")
$taskbarChecks = foreach ($profile in Get-ChildItem -LiteralPath "C:\Users" -Directory) {
    if ($profile.Name.ToLowerInvariant() -in $excludedProfiles) {
        continue
    }
    $path = Join-Path $profile.FullName "AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Etiquetador80mm.lnk"
    $exists = Test-Path -LiteralPath $path -PathType Leaf
    $target = ""
    if ($exists) {
        $target = $shell.CreateShortcut($path).TargetPath
    }
    [pscustomobject]@{
        Profile = $profile.Name
        Exists = $exists
        Target = $target
    }
}
$results.Taskbar = @($taskbarChecks)

$task = Get-ScheduledTask -TaskName "NPV Backup Precios Venta"
$taskInfo = Get-ScheduledTaskInfo -TaskName "NPV Backup Precios Venta"
$results.ScheduledTask = [pscustomobject]@{
    State = [string]$task.State
    User = $task.Principal.UserId
    Execute = $task.Actions[0].Execute
    Arguments = $task.Actions[0].Arguments
    NextRunTime = $taskInfo.NextRunTime
    LastTaskResult = $taskInfo.LastTaskResult
}

$outputPath = Join-Path $projectRoot ".installer_verification.json"
$results | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $outputPath -Encoding UTF8
