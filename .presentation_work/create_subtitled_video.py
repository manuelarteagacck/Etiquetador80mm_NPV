import subprocess
from pathlib import Path

import imageio_ffmpeg


PROJECT_ROOT = Path(__file__).resolve().parent.parent
INPUT_VIDEO = (
    PROJECT_ROOT
    / ".presentation_work"
    / "Recorrido_Etiquetador80mm_CirculoK_v4_base.mp4"
)
INPUT_AUDIO = (
    PROJECT_ROOT
    / ".presentation_work"
    / "audio_natural"
    / "narracion_natural_completa.wav"
)
SUBTITLES = (
    PROJECT_ROOT
    / ".presentation_work"
    / "audio_natural"
    / "subtitulos.srt"
)
OUTPUT_VIDEO = (
    PROJECT_ROOT
    / "presentacion_publica"
    / "Recorrido_Etiquetador80mm_CirculoK_v6_subtitulado.mp4"
)


def escape_subtitle_path(path: Path) -> str:
    relative = path.relative_to(PROJECT_ROOT).as_posix()
    return relative.replace("'", r"\'")


def main():
    for required in (INPUT_VIDEO, INPUT_AUDIO, SUBTITLES):
        if not required.exists():
            raise FileNotFoundError(required)
    if OUTPUT_VIDEO.exists():
        raise FileExistsError(OUTPUT_VIDEO)

    ffmpeg = imageio_ffmpeg.get_ffmpeg_exe()
    subtitle_filter = (
        f"subtitles='{escape_subtitle_path(SUBTITLES)}':"
        "force_style='FontName=Arial,FontSize=11,Bold=1,"
        "PrimaryColour=&H00FFFFFF,OutlineColour=&H00000000,"
        "BackColour=&H70000000,BorderStyle=3,Outline=1,Shadow=0,"
        "Alignment=2,MarginV=18,MarginL=30,MarginR=30'"
    )
    video_filter = subtitle_filter

    subprocess.run(
        [
            ffmpeg,
            "-y",
            "-hide_banner",
            "-i",
            str(INPUT_VIDEO),
            "-i",
            str(INPUT_AUDIO),
            "-vf",
            video_filter,
            "-map",
            "0:v:0",
            "-map",
            "1:a:0",
            "-c:v",
            "libx264",
            "-preset",
            "medium",
            "-crf",
            "18",
            "-pix_fmt",
            "yuv420p",
            "-c:a",
            "aac",
            "-b:a",
            "256k",
            "-shortest",
            "-movflags",
            "+faststart",
            str(OUTPUT_VIDEO),
        ],
        cwd=PROJECT_ROOT,
        check=True,
    )
    print(f"VIDEO_SUBTITULADO={OUTPUT_VIDEO}")


if __name__ == "__main__":
    main()
