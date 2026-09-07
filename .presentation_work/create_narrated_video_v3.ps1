$ErrorActionPreference = "Stop"

$projectRoot = "D:\CCK\EXTRA\Documentos\GitHub\Etiquetador80mm_NPV"
$inputPptx = Join-Path $projectRoot "presentacion_publica\Recorrido_Etiquetador80mm_CirculoK.pptx"
$outputPptx = Join-Path $projectRoot "presentacion_publica\Recorrido_Etiquetador80mm_CirculoK_Narrado_v3.pptx"
$outputVideo = Join-Path $projectRoot "presentacion_publica\Recorrido_Etiquetador80mm_CirculoK_v3.mp4"
$metadataPath = Join-Path $projectRoot ".presentation_work\audio_natural\timings.json"

if (Test-Path -LiteralPath $outputPptx) {
    throw "Ya existe el PowerPoint actualizado: $outputPptx"
}
if (Test-Path -LiteralPath $outputVideo) {
    throw "Ya existe el video actualizado: $outputVideo"
}
if (-not (Test-Path -LiteralPath $metadataPath)) {
    throw "No existe la configuración de la narración: $metadataPath"
}

$metadata = Get-Content -LiteralPath $metadataPath -Raw | ConvertFrom-Json
$audioPath = [string]$metadata.audio
$durations = @($metadata.timings | ForEach-Object { [double]$_ })

if (-not (Test-Path -LiteralPath $audioPath)) {
    throw "No existe la narración natural: $audioPath"
}

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
            # La narración continua evita cortes aunque PowerPoint no exponga
            # el control de duración del efecto en todas sus versiones.
        }
    }

    $firstSlide = $presentation.Slides.Item(1)
    $left = [single]($presentation.PageSetup.SlideWidth - 2)
    $top = [single]($presentation.PageSetup.SlideHeight - 2)
    $audioShape = $firstSlide.Shapes.AddMediaObject2(
        $audioPath,
        0,
        -1,
        $left,
        $top,
        [single]1,
        [single]1
    )
    $audioShape.AnimationSettings.PlaySettings.PlayOnEntry = -1
    $audioShape.AnimationSettings.PlaySettings.HideWhileNotPlaying = -1
    $audioShape.AnimationSettings.PlaySettings.LoopUntilStopped = 0
    $audioShape.AnimationSettings.PlaySettings.StopAfterSlides = $presentation.Slides.Count

    $presentation.SaveAs($outputPptx, 24)
    $presentation.CreateVideo($outputVideo, $true, 5, 1080, 30, 85)

    $deadline = (Get-Date).AddMinutes(12)
    do {
        Start-Sleep -Seconds 3
        $status = [int]$presentation.CreateVideoStatus
        Write-Host "Estado de video actualizado: $status"
        if ((Get-Date) -gt $deadline) {
            throw "La exportación del video actualizado excedió el tiempo esperado."
        }
    } while ($status -eq 1 -or $status -eq 2)

    if ($status -ne 3) {
        throw "PowerPoint no pudo completar el video actualizado. Estado final: $status"
    }

    Write-Host "PPTX_ACTUALIZADO=$outputPptx"
    Write-Host "VIDEO_ACTUALIZADO=$outputVideo"
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
