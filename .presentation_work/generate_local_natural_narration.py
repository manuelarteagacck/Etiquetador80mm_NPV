import json
import textwrap
import wave
from pathlib import Path

import win32com.client


WORK_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = WORK_DIR / "audio_natural"
SEGMENT_DIR = OUTPUT_DIR / "segments"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
SEGMENT_DIR.mkdir(parents=True, exist_ok=True)

VOICE_NAME = "Sabina"
PRE_SILENCE_SECONDS = 0.40
POST_SILENCE_SECONDS = 1.30
FINAL_POST_SILENCE_SECONDS = 3.00

# Cada tupla contiene: texto, velocidad SAPI y pausa posterior. La variación
# moderada de ritmo y las pausas por idea evitan una lectura monótona.
SLIDES = [
    [
        ("Etiquetador ochenta milímetros reúne, en una sola aplicación, la consulta de productos.", 0, 0.30),
        ("También integra la revisión de precios.", -1, 0.20),
        ("Y prepara las etiquetas para la operación de tienda de Círculo K México.", 0, 0.00),
    ],
    [
        ("La pantalla principal organiza el trabajo en tres zonas claras:", -1, 0.20),
        ("búsqueda, vista previa editable y configuración de impresión.", 0, 0.32),
        ("Así, el aliado mantiene el control del contenido antes de generar la etiqueta.", 0, 0.00),
    ],
    [
        ("El aliado puede localizar un producto por artículo, U P C o descripción.", 0, 0.32),
        ("Una vez encontrado, la aplicación carga la información disponible.", -1, 0.20),
        ("Así puede revisarse antes de imprimir.", 0, 0.00),
    ],
    [
        ("La vista previa muestra la composición completa:", -1, 0.18),
        ("producto, precio, vigencia, identificadores y código de barras.", 0, 0.32),
        ("De esta manera, se confirma visualmente el resultado.", -1, 0.20),
        ("Después, puede enviarse a la impresora.", 0, 0.00),
    ],
    [
        ("Al abrir la aplicación, los precios nuevos detectados se presentan en una lista.", 0, 0.32),
        ("El aliado puede seleccionar todos o limpiar la selección.", -1, 0.20),
        ("También puede imprimir únicamente los artículos elegidos.", 0, 0.00),
    ],
    [
        ("El módulo de precios especiales vigentes reúne el precio regular y el precio especial.", -1, 0.20),
        ("Además, muestra la disponibilidad y el periodo promocional.", 0, 0.32),
        ("La selección múltiple facilita preparar las etiquetas necesarias.", 0, 0.00),
    ],
    [
        ("Cuando existe un precio especial, la etiqueta promocional integra el nuevo precio y el ahorro.", 0, 0.28),
        ("También presenta el precio anterior y las condiciones de vigencia.", -1, 0.20),
        ("Todo, en un formato de alta visibilidad.", 0, 0.00),
    ],
    [
        ("En resumen:", -1, 0.18),
        ("buscar, revisar, seleccionar e imprimir.", 0, 0.34),
        ("Con este flujo, las etiquetas quedan listas para la operación diaria.", -1, 0.00),
    ],
]


def get_voice():
    voice = win32com.client.Dispatch("SAPI.SpVoice")
    spanish_voice = next(
        token
        for token in voice.GetVoices()
        if VOICE_NAME in token.GetDescription()
    )
    voice.Voice = spanish_voice
    voice.Volume = 100
    return voice


def synthesize_segment(voice, slide_number, segment_number, text, rate):
    output = SEGMENT_DIR / f"slide-{slide_number:02d}-{segment_number:02d}.wav"
    voice.Rate = rate
    stream = win32com.client.Dispatch("SAPI.SpFileStream")
    stream.Open(str(output), 3, False)
    voice.AudioOutputStream = stream
    voice.Speak(text)
    stream.Close()
    return output


def audio_format(params):
    return (
        params.nchannels,
        params.sampwidth,
        params.framerate,
        params.comptype,
        params.compname,
    )


def combine_segments(paths_and_pauses, output):
    with wave.open(str(paths_and_pauses[0][0]), "rb") as first_file:
        params = first_file.getparams()

    with wave.open(str(output), "wb") as combined:
        combined.setparams(params)
        silence_frame = b"\x00" * params.sampwidth * params.nchannels
        for segment_path, pause_seconds in paths_and_pauses:
            with wave.open(str(segment_path), "rb") as source:
                if audio_format(source.getparams()) != audio_format(params):
                    raise RuntimeError(f"Formato de audio incompatible: {segment_path}")
                combined.writeframes(source.readframes(source.getnframes()))
            if pause_seconds:
                combined.writeframes(
                    silence_frame * int(params.framerate * pause_seconds)
                )


def wav_duration(path):
    with wave.open(str(path), "rb") as source:
        return source.getnframes() / source.getframerate()


def combine_slides(slide_paths):
    output = OUTPUT_DIR / "narracion_natural_completa.wav"
    with wave.open(str(slide_paths[0]), "rb") as first_file:
        params = first_file.getparams()

    timings = []
    with wave.open(str(output), "wb") as combined:
        combined.setparams(params)
        silence_frame = b"\x00" * params.sampwidth * params.nchannels
        pre_silence = silence_frame * int(params.framerate * PRE_SILENCE_SECONDS)
        for slide_index, slide_path in enumerate(slide_paths, start=1):
            duration = wav_duration(slide_path)
            post_silence_seconds = (
                FINAL_POST_SILENCE_SECONDS
                if slide_index == len(slide_paths)
                else POST_SILENCE_SECONDS
            )
            post_silence = silence_frame * int(
                params.framerate * post_silence_seconds
            )
            with wave.open(str(slide_path), "rb") as source:
                if audio_format(source.getparams()) != audio_format(params):
                    raise RuntimeError(f"Formato de audio incompatible: {slide_path}")
                combined.writeframes(pre_silence)
                combined.writeframes(source.readframes(source.getnframes()))
                combined.writeframes(post_silence)
            timings.append(
                duration + PRE_SILENCE_SECONDS + post_silence_seconds
            )

    return output, timings


def srt_timestamp(seconds):
    milliseconds = round(seconds * 1000)
    hours, remainder = divmod(milliseconds, 3_600_000)
    minutes, remainder = divmod(remainder, 60_000)
    whole_seconds, milliseconds = divmod(remainder, 1_000)
    return f"{hours:02d}:{minutes:02d}:{whole_seconds:02d},{milliseconds:03d}"


def subtitle_text(text):
    text = text.replace("U P C", "UPC")
    return "\n".join(
        textwrap.wrap(
            text,
            width=58,
            break_long_words=False,
            break_on_hyphens=False,
        )
    )


def write_subtitles():
    entries = []
    cursor = 0.0
    entry_number = 1

    for slide_number, segments in enumerate(SLIDES, start=1):
        cursor += PRE_SILENCE_SECONDS
        for segment_number, (text, _rate, pause) in enumerate(segments, start=1):
            segment_path = SEGMENT_DIR / (
                f"slide-{slide_number:02d}-{segment_number:02d}.wav"
            )
            duration = wav_duration(segment_path)
            start = cursor
            end = cursor + duration
            entries.extend(
                [
                    str(entry_number),
                    f"{srt_timestamp(start)} --> {srt_timestamp(end)}",
                    subtitle_text(text),
                    "",
                ]
            )
            entry_number += 1
            cursor = end + pause

        cursor += (
            FINAL_POST_SILENCE_SECONDS
            if slide_number == len(SLIDES)
            else POST_SILENCE_SECONDS
        )

    output = OUTPUT_DIR / "subtitulos.srt"
    output.write_text("\n".join(entries), encoding="utf-8-sig")
    return output


def main():
    voice = get_voice()
    slide_paths = []

    for slide_number, segments in enumerate(SLIDES, start=1):
        generated_segments = []
        for segment_number, (text, rate, pause) in enumerate(segments, start=1):
            path = synthesize_segment(
                voice,
                slide_number,
                segment_number,
                text,
                rate,
            )
            generated_segments.append((path, pause))

        slide_path = OUTPUT_DIR / f"slide-{slide_number:02d}.wav"
        combine_segments(generated_segments, slide_path)
        slide_paths.append(slide_path)
        print(f"SLIDE_{slide_number:02d}={slide_path}")

    combined_path, timings = combine_slides(slide_paths)
    subtitles_path = write_subtitles()
    metadata = {
        "voice": "Microsoft Sabina - español de México (local)",
        "audio": str(combined_path),
        "subtitles": str(subtitles_path),
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
    print(f"SUBTITULOS={subtitles_path}")
    print("TIMINGS=" + ",".join(f"{seconds:.3f}" for seconds in timings))


if __name__ == "__main__":
    main()
