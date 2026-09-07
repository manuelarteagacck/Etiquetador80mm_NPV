from pathlib import Path
import wave

import win32com.client


OUTPUT_DIR = Path(__file__).resolve().parent / "audio"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

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


def synthesize(index, text):
    output = OUTPUT_DIR / f"slide-{index:02d}.wav"
    voice = win32com.client.Dispatch("SAPI.SpVoice")
    spanish_voice = next(
        token
        for token in voice.GetVoices()
        if "Sabina" in token.GetDescription()
    )
    voice.Voice = spanish_voice
    voice.Rate = -1
    voice.Volume = 100

    stream = win32com.client.Dispatch("SAPI.SpFileStream")
    stream.Open(str(output), 3, False)
    voice.AudioOutputStream = stream
    voice.Speak(text)
    stream.Close()

    with wave.open(str(output), "rb") as wav_file:
        duration = wav_file.getnframes() / wav_file.getframerate()
    return output, duration


generated = []
for number, narration in enumerate(NARRATIONS, start=1):
    path, seconds = synthesize(number, narration)
    generated.append((path, seconds))
    print(f"{number:02d}\t{seconds:.2f}\t{path}")

combined_path = OUTPUT_DIR / "narracion_completa.wav"
pre_silence_seconds = 0.4
post_silence_seconds = 1.6

with wave.open(str(generated[0][0]), "rb") as first_file:
    params = first_file.getparams()

with wave.open(str(combined_path), "wb") as combined:
    combined.setparams(params)
    silence_frame = b"\x00" * params.sampwidth * params.nchannels
    pre_silence = silence_frame * int(params.framerate * pre_silence_seconds)
    post_silence = silence_frame * int(params.framerate * post_silence_seconds)

    for audio_path, _ in generated:
        with wave.open(str(audio_path), "rb") as source:
            source_params = source.getparams()
            source_format = (
                source_params.nchannels,
                source_params.sampwidth,
                source_params.framerate,
                source_params.comptype,
                source_params.compname,
            )
            expected_format = (
                params.nchannels,
                params.sampwidth,
                params.framerate,
                params.comptype,
                params.compname,
            )
            if source_format != expected_format:
                raise RuntimeError(f"Formato de audio incompatible: {audio_path}")
            combined.writeframes(pre_silence)
            combined.writeframes(source.readframes(source.getnframes()))
            combined.writeframes(post_silence)

timings = [seconds + pre_silence_seconds + post_silence_seconds for _, seconds in generated]
print("TIMINGS=" + ",".join(f"{seconds:.2f}" for seconds in timings))
print(f"AUDIO_CONTINUO={combined_path}")
