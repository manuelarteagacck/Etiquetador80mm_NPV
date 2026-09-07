import json
import subprocess
from pathlib import Path

import imageio_ffmpeg


PROJECT_ROOT = Path(__file__).resolve().parent.parent
WORK_DIR = PROJECT_ROOT / ".presentation_work"
METADATA_PATH = WORK_DIR / "audio_v7" / "timings.json"
CONCAT_PATH = WORK_DIR / "video_v7_frames.txt"
BASE_VIDEO = WORK_DIR / "Recorrido_Etiquetador80mm_CirculoK_v7_base.mp4"
OUTPUT_VIDEO = PROJECT_ROOT / "presentacion_publica" / "Recorrido_Etiquetador80mm_CirculoK_v7_actualizado.mp4"
FRAME_DIR = WORK_DIR / "v7_frames"

IMAGES = [
    WORK_DIR / "rendered" / "slide-01.png",
    WORK_DIR / "screens" / "02_principal.png",
    WORK_DIR / "screens" / "10_precios_nuevos_version_vigente.png",
    WORK_DIR / "screens" / "12_articulo_cargado_vigente.png",
    WORK_DIR / "screens" / "13_vista_previa_vigente.png",
    WORK_DIR / "screens" / "14_configuracion_impresion_vigente.png",
    WORK_DIR / "screens" / "15_precios_especiales_vigentes.png",
    WORK_DIR / "screens" / "17_articulo_promocion_vigente.png",
    WORK_DIR / "screens" / "18_vista_previa_promocion_vigente.png",
    WORK_DIR / "rendered" / "slide-08.png",
]


def quote_concat_path(path: Path) -> str:
    return path.resolve().as_posix().replace("'", "'\\''")


def subtitle_filter(path: Path) -> str:
    relative = path.relative_to(PROJECT_ROOT).as_posix().replace("'", r"\'")
    return (
        f"subtitles='{relative}':"
        "force_style='FontName=Arial,FontSize=12,Bold=1,"
        "PrimaryColour=&H00FFFFFF,OutlineColour=&H00000000,"
        "BackColour=&H78000000,BorderStyle=3,Outline=1,Shadow=0,"
        "Alignment=2,MarginV=24,MarginL=36,MarginR=36'"
    )


def main():
    metadata = json.loads(METADATA_PATH.read_text(encoding="utf-8"))
    timings = [float(value) for value in metadata["timings"]]
    if len(timings) != len(IMAGES):
        raise RuntimeError(f"Se esperaban {len(IMAGES)} tiempos y se recibieron {len(timings)}")
    for image in IMAGES:
        if not image.exists():
            raise FileNotFoundError(image)

    ffmpeg = imageio_ffmpeg.get_ffmpeg_exe()
    base_command = [ffmpeg, "-y", "-hide_banner", "-loglevel", "warning"]
    for image, duration in zip(IMAGES, timings):
        base_command.extend([
            "-loop", "1", "-framerate", "30", "-t", f"{duration:.3f}", "-i", str(image)
        ])
    filter_parts = []
    for index, duration in enumerate(timings):
        filter_parts.append(
            f"[{index}:v]scale=1920:1080:force_original_aspect_ratio=decrease,"
            f"pad=1920:1080:(ow-iw)/2:(oh-ih)/2:color=0xF2F3F4,"
            f"setsar=1,trim=duration={duration:.3f},setpts=PTS-STARTPTS[v{index}]"
        )
    inputs = "".join(f"[v{index}]" for index in range(len(IMAGES)))
    filter_parts.append(f"{inputs}concat=n={len(IMAGES)}:v=1:a=0,format=yuv420p[vout]")
    base_command.extend([
        "-filter_complex", ";".join(filter_parts), "-map", "[vout]",
        "-c:v", "libx264", "-preset", "medium", "-crf", "18",
        "-pix_fmt", "yuv420p", "-movflags", "+faststart", str(BASE_VIDEO),
    ])
    subprocess.run(base_command, cwd=PROJECT_ROOT, check=True)

    audio_path = Path(metadata["audio"])
    subtitles_path = Path(metadata["subtitles"])
    subprocess.run([
        ffmpeg, "-y", "-hide_banner", "-loglevel", "warning",
        "-i", str(BASE_VIDEO), "-i", str(audio_path),
        "-vf", subtitle_filter(subtitles_path),
        "-map", "0:v:0", "-map", "1:a:0",
        "-c:v", "libx264", "-preset", "medium", "-crf", "18",
        "-pix_fmt", "yuv420p", "-c:a", "aac", "-b:a", "192k",
        "-shortest", "-movflags", "+faststart", str(OUTPUT_VIDEO),
    ], cwd=PROJECT_ROOT, check=True)

    FRAME_DIR.mkdir(parents=True, exist_ok=True)
    cursor = 0.0
    for index, duration in enumerate(timings, start=1):
        midpoint = cursor + duration / 2
        output = FRAME_DIR / f"frame-{index:02d}.png"
        subprocess.run([
            ffmpeg, "-y", "-hide_banner", "-loglevel", "error",
            "-ss", f"{midpoint:.3f}", "-i", str(OUTPUT_VIDEO),
            "-frames:v", "1", str(output),
        ], check=True)
        cursor += duration

    print(f"VIDEO={OUTPUT_VIDEO}")
    print(f"DURATION={metadata['total_duration']}")
    print(f"FRAMES={FRAME_DIR}")


if __name__ == "__main__":
    main()
