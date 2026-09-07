$ErrorActionPreference = "Stop"

$projectRoot = "D:\CCK\EXTRA\Documentos\GitHub\Etiquetador80mm_NPV"
$inputPptx = Join-Path $projectRoot "presentacion_publica\Recorrido_Etiquetador80mm_CirculoK.pptx"
$outputPptx = Join-Path $projectRoot "presentacion_publica\Recorrido_Etiquetador80mm_CirculoK_Narrado_v2.pptx"
$outputVideo = Join-Path $projectRoot "presentacion_publica\Recorrido_Etiquetador80mm_CirculoK_v2.mp4"
$audioPath = Join-Path $projectRoot ".presentation_work\audio\narracion_completa.wav"

if (Test-Path -LiteralPath $outputPptx) {
    throw "Ya existe el PowerPoint corregido: $outputPptx"
}
if (Test-Path -LiteralPath $outputVideo) {
    throw "Ya existe el video corregido: $outputVideo"
}
if (-not (Test-Path -LiteralPath $audioPath)) {
    throw "No existe la narracion continua: $audioPath"
}

$durations = @(14.63, 16.36, 15.29, 17.08, 15.91, 16.18, 14.40, 14.57)
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
            # La pista continua evita cortes aunque la version de PowerPoint
            # no exponga el control de duracion del efecto.
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
        Write-Host "Estado de video corregido: $status"
        if ((Get-Date) -gt $deadline) {
            throw "La exportacion del video corregido excedio el tiempo esperado."
        }
    } while ($status -eq 1 -or $status -eq 2)

    if ($status -ne 3) {
        throw "PowerPoint no pudo completar el video corregido. Estado final: $status"
    }

    Write-Host "PPTX_CORREGIDO=$outputPptx"
    Write-Host "VIDEO_CORREGIDO=$outputVideo"
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
