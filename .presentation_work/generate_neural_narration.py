import asyncio
import json
import subprocess
import wave
from pathlib import Path

import edge_tts
import imageio_ffmpeg


WORK_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = WORK_DIR / "audio_neural"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

VOICE = "es-MX-DaliaNeural"
RATE = "-4%"
PITCH = "-2Hz"
PRE_SILENCE_SECONDS = 0.45
POST_SILENCE_SECONDS = 1.35

NARRATIONS = [
    "Etiquetador ochenta milímetros reúne en una sola aplicación la consulta de productos, la revisión de precios y la preparación de etiquetas para la operación de tienda de Círculo K México.",
    "La pantalla principal organiza el trabajo en tres zonas claras: búsqueda, vista previa editable y configuración de impresión. El aliado mantiene control del contenido antes de generar la etiqueta.",
    "El aliado puede localizar un producto por artículo, UPC o descripción. Una vez encontrado, la aplicación carga la información disponible para su revisión antes de imprimir.",
    "La vista previa muestra la composición completa: producto, precio, vigencia, identificadores y código de barras. Así se confirma visualmente el resultado antes de enviarlo a la impresora.",
    "Al abrir la aplicación, los precios nuevos detectados se presentan en una lista. El aliado puede seleccionar todos, limpiar la selección o imprimir únicamente los artículos elegidos.",
    "El módulo de precios especiales vigentes reúne precio regular, precio especial, disponibilidad y periodo promocional. La selección múltiple facilita preparar las etiquetas necesarias.",
    "Cuando existe un precio especial, la etiqueta promocional integra el nuevo precio, el ahorro, el precio anterior y las condiciones de vigencia en un formato de alta visibilidad.",
    "En resumen: buscar, revisar, seleccionar e imprimir. Etiquetador ochenta milímetros centraliza el recorrido para preparar las etiquetas de la operación diaria.",
]


async def synthesize_slide(index: int, text: str) -> Path:
    output = OUTPUT_DIR / f"slide-{index:02d}.mp3"
    communicator = edge_tts.Communicate(
        text,
        VOICE,
        rate=RATE,
        volume="+0%",
        pitch=PITCH,
    )
    await communicator.save(str(output))
    return output


def convert_to_wav(mp3_path: Path) -> Path:
    wav_path = mp3_path.with_suffix(".wav")
    ffmpeg = imageio_ffmpeg.get_ffmpeg_exe()
    subprocess.run(
        [
            ffmpeg,
            "-y",
            "-loglevel",
            "error",
            "-i",
            str(mp3_path),
            "-ac",
            "1",
            "-ar",
            "24000",
            "-c:a",
            "pcm_s16le",
            str(wav_path),
        ],
        check=True,
    )
    return wav_path


def wav_duration(path: Path) -> float:
    with wave.open(str(path), "rb") as wav_file:
        return wav_file.getnframes() / wav_file.getframerate()


def combine_audio(wav_paths: list[Path]) -> tuple[Path, list[float]]:
    combined_path = OUTPUT_DIR / "narracion_neural_completa.wav"
    with wave.open(str(wav_paths[0]), "rb") as first_file:
        params = first_file.getparams()

    timings = []
    with wave.open(str(combined_path), "wb") as combined:
        combined.setparams(params)
        silence_frame = b"\x00" * params.sampwidth * params.nchannels
        pre_silence = silence_frame * int(params.framerate * PRE_SILENCE_SECONDS)
        post_silence = silence_frame * int(params.framerate * POST_SILENCE_SECONDS)

        for wav_path in wav_paths:
            with wave.open(str(wav_path), "rb") as source:
                source_format = (
                    source.getnchannels(),
                    source.getsampwidth(),
                    source.getframerate(),
                    source.getcomptype(),
                )
                expected_format = (
                    params.nchannels,
                    params.sampwidth,
                    params.framerate,
                    params.comptype,
                )
                if source_format != expected_format:
                    raise RuntimeError(f"Formato de audio incompatible: {wav_path}")
                duration = source.getnframes() / source.getframerate()
                combined.writeframes(pre_silence)
                combined.writeframes(source.readframes(source.getnframes()))
                combined.writeframes(post_silence)
                timings.append(duration + PRE_SILENCE_SECONDS + POST_SILENCE_SECONDS)

    return combined_path, timings


async def main() -> None:
    voices = await edge_tts.list_voices()
    if not any(voice.get("ShortName") == VOICE for voice in voices):
        raise RuntimeError(f"La voz solicitada no está disponible: {VOICE}")

    mp3_paths = []
    for number, narration in enumerate(NARRATIONS, start=1):
        mp3_path = await synthesize_slide(number, narration)
        mp3_paths.append(mp3_path)
        print(f"MP3_{number:02d}={mp3_path}")

    wav_paths = [convert_to_wav(path) for path in mp3_paths]
    combined_path, timings = combine_audio(wav_paths)

    metadata = {
        "voice": VOICE,
        "rate": RATE,
        "pitch": PITCH,
        "audio": str(combined_path),
        "timings": [round(seconds, 3) for seconds in timings],
        "total_duration": round(wav_duration(combined_path), 3),
    }
    metadata_path = OUTPUT_DIR / "timings.json"
    metadata_path.write_text(
        json.dumps(metadata, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(f"METADATA={metadata_path}")
    print(f"AUDIO_CONTINUO={combined_path}")
    print("TIMINGS=" + ",".join(f"{seconds:.3f}" for seconds in timings))


if __name__ == "__main__":
    asyncio.run(main())
