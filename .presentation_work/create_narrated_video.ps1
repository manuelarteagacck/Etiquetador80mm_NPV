$ErrorActionPreference = "Stop"

$projectRoot = "D:\CCK\EXTRA\Documentos\GitHub\Etiquetador80mm_NPV"
$inputPptx = Join-Path $projectRoot "presentacion_publica\Recorrido_Etiquetador80mm_CirculoK.pptx"
$outputPptx = Join-Path $projectRoot "presentacion_publica\Recorrido_Etiquetador80mm_CirculoK_Narrado.pptx"
$outputVideo = Join-Path $projectRoot "presentacion_publica\Recorrido_Etiquetador80mm_CirculoK.mp4"
$audioDir = Join-Path $projectRoot ".presentation_work\audio"

if (Test-Path -LiteralPath $outputPptx) {
    throw "Ya existe el PowerPoint narrado: $outputPptx"
}
if (Test-Path -LiteralPath $outputVideo) {
    throw "Ya existe el video: $outputVideo"
}

$durations = @(12.5, 15.8, 14.7, 16.5, 15.3, 15.6, 13.8, 14.0)
$powerPoint = $null
$presentation = $null

try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $presentation = $powerPoint.Presentations.Open($inputPptx, $false, $false, $false)

    if ($presentation.Slides.Count -ne $durations.Count) {
        throw "Se esperaban $($durations.Count) diapositivas y se encontraron $($presentation.Slides.Count)."
    }

    for ($index = 1; $index -le $presentation.Slides.Count; $index++) {
        $slide = $presentation.Slides.Item($index)
        $audioPath = Join-Path $audioDir ("slide-{0:D2}.wav" -f $index)
        if (-not (Test-Path -LiteralPath $audioPath)) {
            throw "No existe el audio: $audioPath"
        }

        $left = [single]($presentation.PageSetup.SlideWidth - 2)
        $top = [single]($presentation.PageSetup.SlideHeight - 2)
        $audioShape = $slide.Shapes.AddMediaObject2(
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
        $audioShape.AnimationSettings.PlaySettings.StopAfterSlides = 1

        $transition = $slide.SlideShowTransition
        $transition.AdvanceOnClick = 0
        $transition.AdvanceOnTime = -1
        $transition.AdvanceTime = [single]$durations[$index - 1]
        try {
            $transition.EntryEffect = [Microsoft.Office.Interop.PowerPoint.PpEntryEffect]::ppEffectFadeSmoothly
        }
        catch {
            # La sincronizacion no depende del efecto visual.
        }
    }

    $presentation.SaveAs($outputPptx, 24)
    $presentation.CreateVideo($outputVideo, $true, 5, 1080, 30, 85)

    $deadline = (Get-Date).AddMinutes(12)
    do {
        Start-Sleep -Seconds 3
        $status = [int]$presentation.CreateVideoStatus
        Write-Host "Estado de video: $status"
        if ((Get-Date) -gt $deadline) {
            throw "La exportacion de video excedio el tiempo esperado."
        }
    } while ($status -eq 1 -or $status -eq 2)

    if ($status -ne 3) {
        throw "PowerPoint no pudo completar el video. Estado final: $status"
    }

    Write-Host "PPTX_NARRADO=$outputPptx"
    Write-Host "VIDEO=$outputVideo"
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
