$ErrorActionPreference = "Stop"

$projectRoot = "D:\CCK\EXTRA\Documentos\GitHub\Etiquetador80mm_NPV"
$inputPptx = Join-Path $projectRoot "presentacion_publica\Recorrido_Etiquetador80mm_CirculoK.pptx"
$outputVideo = Join-Path $projectRoot ".presentation_work\Recorrido_Etiquetador80mm_CirculoK_v4_base.mp4"
$metadataPath = Join-Path $projectRoot ".presentation_work\audio_natural\timings.json"

if (Test-Path -LiteralPath $outputVideo) {
    throw "Ya existe el video base actualizado: $outputVideo"
}
if (-not (Test-Path -LiteralPath $metadataPath)) {
    throw "No existe la configuración de narración: $metadataPath"
}

$metadata = Get-Content -LiteralPath $metadataPath -Raw | ConvertFrom-Json
$durations = @($metadata.timings | ForEach-Object { [double]$_ })
$powerPoint = $null
$presentation = $null

try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $presentation = $powerPoint.Presentations.Open($inputPptx, $false, $false, $false)

    if ($presentation.Slides.Count -ne $durations.Count) {
        throw "Se esperaban $($durations.Count) diapositivas y se encontraron $($presentation.Slides.Count)."
    }

    for ($index = 1; $index -le $presentation.Slides.Count; $index++) {
        $transition = $presentation.Slides.Item($index).SlideShowTransition
        $transition.AdvanceOnClick = 0
        $transition.AdvanceOnTime = -1
        $transition.AdvanceTime = [single]$durations[$index - 1]
        try {
            $transition.EntryEffect = [Microsoft.Office.Interop.PowerPoint.PpEntryEffect]::ppEffectNone
            $transition.Duration = [single]0
        }
        catch {
            # Se conservan los tiempos aun si esta versión de PowerPoint
            # no expone la duración del efecto.
        }
    }

    $presentation.CreateVideo($outputVideo, $true, 5, 1080, 30, 85)

    $deadline = (Get-Date).AddMinutes(12)
    do {
        Start-Sleep -Seconds 3
        $status = [int]$presentation.CreateVideoStatus
        Write-Host "Estado del video base: $status"
        if ((Get-Date) -gt $deadline) {
            throw "La exportación del video base excedió el tiempo esperado."
        }
    } while ($status -eq 1 -or $status -eq 2)

    if ($status -ne 3) {
        throw "PowerPoint no pudo completar el video base. Estado final: $status"
    }

    Write-Host "VIDEO_BASE=$outputVideo"
}
finally {
    if ($null -ne $presentation) {
        $presentation.Close()
    }
    if ($null -ne $powerPoint) {
        $powerPoint.Quit()
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
