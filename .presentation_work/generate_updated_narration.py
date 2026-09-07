import json
import textwrap
import wave
from pathlib import Path

import win32com.client


WORK_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = WORK_DIR / "audio_v7"
SEGMENT_DIR = OUTPUT_DIR / "segments"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
SEGMENT_DIR.mkdir(parents=True, exist_ok=True)

VOICE_NAME = "Sabina"
PRE_SILENCE_SECONDS = 0.40
POST_SILENCE_SECONDS = 1.15
FINAL_POST_SILENCE_SECONDS = 2.50

CHAPTERS = [
    [
        ("Esta es la versión actualizada del Etiquetador de ochenta milímetros para tienda N P V.", 0, 0.25),
        ("La aplicación reúne consulta, validación e impresión de etiquetas en un solo flujo.", -1, 0.00),
    ],
    [
        ("La pantalla principal organiza el trabajo en búsqueda, vista previa editable y configuración de impresión.", 0, 0.25),
        ("Desde aquí también se consultan precios nuevos, promociones vigentes y la impresión de prueba.", -1, 0.00),
    ],
    [
        ("Al abrir, la aplicación compara los precios por artículo y fecha de inicio.", 0, 0.25),
        ("Si encuentra cambios pendientes, muestra automáticamente la lista para seleccionar todos, limpiar la selección o imprimir solo los elegidos.", -1, 0.00),
    ],
    [
        ("La búsqueda acepta artículo, U P C o descripción.", 0, 0.20),
        ("El producto localizado carga descripción, precio, código interno y U P C para revisión.", -1, 0.20),
        ("La impresión automática al escanear puede mantenerse activa o desactivarse según el flujo de la tienda.", 0, 0.00),
    ],
    [
        ("Antes de imprimir, la vista previa confirma la etiqueta normal completa.", 0, 0.20),
        ("Se revisan producto, precio, vigencia, identificadores y código de barras, incluido el texto legible bajo las barras.", -1, 0.00),
    ],
    [
        ("La salida puede usar la impresora predeterminada o una impresora elegida de la lista.", 0, 0.22),
        ("También se define el número de copias, la impresión automática y el formato individual angosto.", -1, 0.20),
        ("El botón Probar impresión permite validar la comunicación con el equipo.", 0, 0.00),
    ],
    [
        ("Precios Especiales Vigentes muestra en una sola lista el precio regular, el precio promocional, la existencia y la vigencia.", 0, 0.25),
        ("La selección múltiple permite preparar únicamente las etiquetas necesarias.", -1, 0.00),
    ],
    [
        ("Cuando el artículo tiene promoción, la aplicación habilita el formato de precio especial.", 0, 0.22),
        ("El precio normal y el precio promocional quedan visibles antes de generar la etiqueta.", -1, 0.00),
    ],
    [
        ("La etiqueta promocional comunica el ahorro, el precio nuevo, el precio anterior y las condiciones de vigencia.", 0, 0.22),
        ("La vista previa permite revisar el resultado sin enviar todavía el trabajo a la impresora.", -1, 0.00),
    ],
    [
        ("Como soporte del flujo diario, una tarea programada conserva la instantánea de precios anteriores a las nueve de la noche.", 0, 0.25),
        ("En conjunto, el proceso es claro: buscar, revisar, seleccionar e imprimir.", -1, 0.20),
        ("Las etiquetas quedan listas para la operación diaria de tienda.", 0, 0.00),
    ],
]


def get_voice():
    voice = win32com.client.Dispatch("SAPI.SpVoice")
    spanish_voice = next(
        token for token in voice.GetVoices() if VOICE_NAME in token.GetDescription()
    )
    voice.Voice = spanish_voice
    voice.Volume = 100
    return voice


def synthesize_segment(voice, chapter_number, segment_number, text, rate):
    output = SEGMENT_DIR / f"chapter-{chapter_number:02d}-{segment_number:02d}.wav"
    voice.Rate = rate
    stream = win32com.client.Dispatch("SAPI.SpFileStream")
    stream.Open(str(output), 3, False)
    voice.AudioOutputStream = stream
    voice.Speak(text)
    stream.Close()
    return output


def comparable_format(params):
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
                if comparable_format(source.getparams()) != comparable_format(params):
                    raise RuntimeError(f"Formato de audio incompatible: {segment_path}")
                combined.writeframes(source.readframes(source.getnframes()))
            combined.writeframes(silence_frame * int(params.framerate * pause_seconds))


def wav_duration(path):
    with wave.open(str(path), "rb") as source:
        return source.getnframes() / source.getframerate()


def combine_chapters(chapter_paths):
    output = OUTPUT_DIR / "narracion_actualizada.wav"
    with wave.open(str(chapter_paths[0]), "rb") as first_file:
        params = first_file.getparams()
    timings = []
    with wave.open(str(output), "wb") as combined:
        combined.setparams(params)
        silence_frame = b"\x00" * params.sampwidth * params.nchannels
        for chapter_index, chapter_path in enumerate(chapter_paths, start=1):
            post_seconds = FINAL_POST_SILENCE_SECONDS if chapter_index == len(chapter_paths) else POST_SILENCE_SECONDS
            with wave.open(str(chapter_path), "rb") as source:
                combined.writeframes(silence_frame * int(params.framerate * PRE_SILENCE_SECONDS))
                combined.writeframes(source.readframes(source.getnframes()))
                combined.writeframes(silence_frame * int(params.framerate * post_seconds))
                timings.append(wav_duration(chapter_path) + PRE_SILENCE_SECONDS + post_seconds)
    return output, timings


def srt_timestamp(seconds):
    milliseconds = round(seconds * 1000)
    hours, remainder = divmod(milliseconds, 3_600_000)
    minutes, remainder = divmod(remainder, 60_000)
    whole_seconds, milliseconds = divmod(remainder, 1_000)
    return f"{hours:02d}:{minutes:02d}:{whole_seconds:02d},{milliseconds:03d}"


def subtitle_text(text):
    text = text.replace("U P C", "UPC").replace("N P V", "NPV")
    return "\n".join(textwrap.wrap(text, width=58, break_long_words=False, break_on_hyphens=False))


def write_subtitles():
    entries = []
    cursor = 0.0
    entry_number = 1
    for chapter_number, segments in enumerate(CHAPTERS, start=1):
        cursor += PRE_SILENCE_SECONDS
        for segment_number, (text, _rate, pause) in enumerate(segments, start=1):
            segment_path = SEGMENT_DIR / f"chapter-{chapter_number:02d}-{segment_number:02d}.wav"
            duration = wav_duration(segment_path)
            entries.extend([
                str(entry_number),
                f"{srt_timestamp(cursor)} --> {srt_timestamp(cursor + duration)}",
                subtitle_text(text),
                "",
            ])
            entry_number += 1
            cursor += duration + pause
        cursor += FINAL_POST_SILENCE_SECONDS if chapter_number == len(CHAPTERS) else POST_SILENCE_SECONDS
    output = OUTPUT_DIR / "subtitulos_actualizados.srt"
    output.write_text("\n".join(entries), encoding="utf-8-sig")
    return output


def main():
    voice = get_voice()
    chapter_paths = []
    for chapter_number, segments in enumerate(CHAPTERS, start=1):
        generated = []
        for segment_number, (text, rate, pause) in enumerate(segments, start=1):
            generated.append((synthesize_segment(voice, chapter_number, segment_number, text, rate), pause))
        chapter_path = OUTPUT_DIR / f"chapter-{chapter_number:02d}.wav"
        combine_segments(generated, chapter_path)
        chapter_paths.append(chapter_path)
        print(f"CHAPTER_{chapter_number:02d}={chapter_path}")

    combined_path, timings = combine_chapters(chapter_paths)
    subtitles_path = write_subtitles()
    metadata = {
        "voice": "Microsoft Sabina - español de México (local)",
        "audio": str(combined_path),
        "subtitles": str(subtitles_path),
        "timings": [round(seconds, 3) for seconds in timings],
        "total_duration": round(wav_duration(combined_path), 3),
    }
    metadata_path = OUTPUT_DIR / "timings.json"
    metadata_path.write_text(json.dumps(metadata, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"METADATA={metadata_path}")
    print(f"AUDIO={combined_path}")
    print(f"SUBTITLES={subtitles_path}")
    print("TIMINGS=" + ",".join(f"{seconds:.3f}" for seconds in timings))


if __name__ == "__main__":
    main()
