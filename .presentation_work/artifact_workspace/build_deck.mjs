import fs from "node:fs/promises";
import path from "node:path";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const ROOT = "D:/CCK/EXTRA/Documentos/GitHub/Etiquetador80mm_NPV";
const WORK = path.join(ROOT, ".presentation_work");
const SCREENS = path.join(WORK, "screens");
const OUTPUT_DIR = path.join(ROOT, "presentacion_publica");
const OUTPUT_PPTX = path.join(OUTPUT_DIR, "Recorrido_Etiquetador80mm_CirculoK.pptx");

const W = 1280;
const H = 720;
const RED = "#E2231A";
const BLACK = "#171717";
const GRAY = "#F2F3F4";
const MID = "#5D646B";
const LIGHT = "#D8DBDE";
const WHITE = "#FFFFFF";
const FONT = "Aptos";
const FONT_DISPLAY = "Aptos Display";

async function readBytes(filePath) {
  const bytes = await fs.readFile(filePath);
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

function addText(slide, text, position, options = {}) {
  const shape = slide.shapes.add({
    geometry: "textbox",
    name: options.name,
    position,
    fill: options.fill ?? "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  shape.text = text;
  shape.text.fontSize = options.fontSize ?? 24;
  shape.text.bold = Boolean(options.bold);
  shape.text.color = options.color ?? BLACK;
  shape.text.typeface = options.typeface ?? FONT;
  shape.text.alignment = options.alignment ?? "left";
  shape.text.verticalAlignment = options.verticalAlignment ?? "top";
  shape.text.insets = options.insets ?? { left: 0, right: 0, top: 0, bottom: 0 };
  return shape;
}

function addRect(slide, position, fill, options = {}) {
  return slide.shapes.add({
    geometry: options.geometry ?? "rect",
    name: options.name,
    position,
    fill,
    line: options.line ?? { style: "solid", fill: "none", width: 0 },
    borderRadius: options.borderRadius,
    shadow: options.shadow,
  });
}

async function addImage(slide, filePath, position, alt, options = {}) {
  return slide.images.add({
    blob: await readBytes(filePath),
    contentType: "image/png",
    alt,
    fit: options.fit ?? "contain",
    crop: options.crop,
    position,
    geometry: options.geometry ?? "roundRect",
    borderRadius: options.borderRadius ?? 12,
  });
}

async function addLogo(slide, position = { left: 1015, top: 36, width: 195, height: 50 }) {
  return addImage(
    slide,
    path.join(ROOT, "LOGO CIRCLE K.png"),
    position,
    "Logotipo de Circle K",
    { fit: "contain", geometry: "rect", borderRadius: 0 },
  );
}

function addHeader(slide, title, page) {
  addText(slide, title, { left: 70, top: 42, width: 880, height: 55 }, {
    fontSize: 38,
    bold: true,
    typeface: FONT_DISPLAY,
    name: `slide-${page}-title`,
  });
  addRect(slide, { left: 70, top: 108, width: 76, height: 5 }, RED);
  addText(slide, String(page).padStart(2, "0"), { left: 1190, top: 662, width: 40, height: 24 }, {
    fontSize: 15,
    bold: true,
    color: MID,
    alignment: "right",
  });
}

function addScreenshotFrame(slide, position) {
  addRect(slide, {
    left: position.left - 8,
    top: position.top - 8,
    width: position.width + 16,
    height: position.height + 16,
  }, WHITE, {
    geometry: "roundRect",
    borderRadius: 18,
    line: { style: "solid", fill: LIGHT, width: 1 },
    shadow: "shadow-md",
  });
}

function addNotes(slide, narration, sources) {
  const sourceLines = sources.map((source) => `- ${source}`).join("\n");
  slide.speakerNotes.textFrame.setText(
    `${narration}\n\n[Sources]\n${sourceLines}\n[/Sources]`,
  );
}

async function build() {
  await fs.mkdir(OUTPUT_DIR, { recursive: true });
  await fs.mkdir(path.join(WORK, "rendered"), { recursive: true });

  const presentation = Presentation.create({ slideSize: { width: W, height: H } });

  // 1. Apertura
  {
    const slide = presentation.slides.add();
    slide.background.fill = WHITE;
    addRect(slide, { left: 0, top: 0, width: 24, height: H }, RED);
    await addLogo(slide, { left: 74, top: 60, width: 255, height: 64 });
    addText(slide, "Etiquetador\n80 mm", { left: 72, top: 174, width: 460, height: 160 }, {
      fontSize: 62,
      bold: true,
      typeface: FONT_DISPLAY,
      name: "cover-title",
    });
    addText(slide, "Recorrido funcional para operación en tienda", { left: 76, top: 350, width: 430, height: 74 }, {
      fontSize: 27,
      color: MID,
    });
    addRect(slide, { left: 76, top: 452, width: 124, height: 6 }, RED);
    addText(slide, "Consulta. Revisa. Imprime.", { left: 76, top: 478, width: 430, height: 42 }, {
      fontSize: 25,
      bold: true,
      color: BLACK,
    });
    const shot = { left: 590, top: 82, width: 585, height: 505 };
    addScreenshotFrame(slide, shot);
    await addImage(slide, path.join(SCREENS, "02_principal.png"), shot, "Pantalla principal del Etiquetador 80 mm");
    addText(slide, "CÍRCULO K MÉXICO", { left: 76, top: 642, width: 260, height: 28 }, {
      fontSize: 16,
      bold: true,
      color: RED,
    });
    addNotes(
      slide,
      "Etiquetador 80 milímetros concentra en una sola aplicación la consulta, revisión e impresión de etiquetas para la operación de tienda de Círculo K México.",
      [
        "Logo: LOGO CIRCLE K.png, recurso local proporcionado en el proyecto.",
        "Captura: 02_principal.png, aplicación local observada el 31/07/2026.",
      ],
    );
  }

  // 2. Pantalla principal
  {
    const slide = presentation.slides.add();
    slide.background.fill = GRAY;
    addHeader(slide, "Una sola pantalla concentra el flujo", 2);
    await addLogo(slide);
    const shot = { left: 70, top: 138, width: 670, height: 555 };
    addScreenshotFrame(slide, shot);
    await addImage(slide, path.join(SCREENS, "02_principal.png"), shot, "Pantalla principal con búsqueda, vista editable e impresión");
    const lines = [
      ["01", "Búsqueda", "Artículo, UPC o descripción"],
      ["02", "Revisión", "Datos editables antes de imprimir"],
      ["03", "Salida", "Impresora, formato y número de copias"],
    ];
    let top = 178;
    for (const [number, heading, body] of lines) {
      addText(slide, number, { left: 800, top, width: 62, height: 42 }, {
        fontSize: 30,
        bold: true,
        color: RED,
      });
      addText(slide, heading, { left: 880, top: top - 2, width: 280, height: 38 }, {
        fontSize: 26,
        bold: true,
      });
      addText(slide, body, { left: 880, top: top + 38, width: 300, height: 54 }, {
        fontSize: 19,
        color: MID,
      });
      top += 146;
    }
    addNotes(
      slide,
      "La pantalla principal organiza el trabajo en tres zonas claras: búsqueda, vista previa editable y configuración de impresión. El aliado mantiene control del contenido antes de generar la etiqueta.",
      ["Captura: 02_principal.png, aplicación local observada el 31/07/2026."],
    );
  }

  // 3. Búsqueda y carga
  {
    const slide = presentation.slides.add();
    slide.background.fill = WHITE;
    addHeader(slide, "La información del artículo llega lista para revisar", 3);
    await addLogo(slide);
    const shot = { left: 70, top: 138, width: 700, height: 558 };
    addScreenshotFrame(slide, shot);
    await addImage(slide, path.join(SCREENS, "03_articulo_cargado.png"), shot, "Artículo cargado con descripción, precio, código y UPC");
    addText(slide, "Una búsqueda puede iniciar por:", { left: 825, top: 172, width: 340, height: 40 }, {
      fontSize: 24,
      bold: true,
    });
    addText(slide, "Artículo\nUPC\nDescripción", { left: 825, top: 235, width: 320, height: 180 }, {
      fontSize: 35,
      bold: true,
      color: RED,
    });
    addRect(slide, { left: 825, top: 448, width: 315, height: 2 }, BLACK);
    addText(slide, "Descripción, precio, código interno y UPC quedan visibles para confirmación.", { left: 825, top: 478, width: 340, height: 120 }, {
      fontSize: 21,
      color: MID,
    });
    addNotes(
      slide,
      "El aliado puede localizar un producto por artículo, UPC o descripción. Una vez encontrado, la aplicación carga la información disponible para su revisión y, cuando corresponde, permite ajustes antes de imprimir.",
      ["Captura: 03_articulo_cargado.png, búsqueda local realizada el 31/07/2026."],
    );
  }

  // 4. Vista previa normal
  {
    const slide = presentation.slides.add();
    slide.background.fill = GRAY;
    addHeader(slide, "La vista previa confirma la etiqueta antes de imprimir", 4);
    await addLogo(slide);
    const shot = { left: 70, top: 150, width: 730, height: 480 };
    addScreenshotFrame(slide, shot);
    await addImage(slide, path.join(SCREENS, "04_vista_previa.png"), shot, "Vista previa de etiqueta normal de 80 mm");
    addText(slide, "El aliado valida", { left: 860, top: 178, width: 300, height: 38 }, {
      fontSize: 27,
      bold: true,
    });
    addText(slide, "Producto\nPrecio\nVigencia\nArtículo y UPC\nCódigo de barras", { left: 860, top: 242, width: 315, height: 245 }, {
      fontSize: 27,
      bold: true,
      color: RED,
    });
    addText(slide, "La impresión solo se envía cuando el contenido ya fue revisado.", { left: 860, top: 535, width: 300, height: 80 }, {
      fontSize: 20,
      color: MID,
    });
    addNotes(
      slide,
      "Antes de imprimir, la vista previa muestra la composición completa: producto, precio, vigencia, identificadores y código de barras. Esto permite confirmar visualmente el resultado.",
      ["Captura: 04_vista_previa.png, vista previa local generada el 31/07/2026."],
    );
  }

  // 5. Precios nuevos
  {
    const slide = presentation.slides.add();
    slide.background.fill = WHITE;
    addHeader(slide, "Los precios nuevos aparecen automáticamente al abrir", 5);
    await addLogo(slide);
    const shot = { left: 140, top: 130, width: 1000, height: 560 };
    addScreenshotFrame(slide, shot);
    await addImage(slide, path.join(SCREENS, "01_precios_nuevos.png"), shot, "Selección automática de precios nuevos disponibles");
    addRect(slide, { left: 158, top: 548, width: 964, height: 82 }, "#FFFFFFE8", {
      geometry: "roundRect",
      borderRadius: 12,
      line: { style: "solid", fill: LIGHT, width: 1 },
    });
    addText(slide, "Seleccionar  ·  revisar vigencia  ·  imprimir en lote", { left: 180, top: 570, width: 920, height: 42 }, {
      fontSize: 25,
      bold: true,
      color: BLACK,
      alignment: "center",
    });
    addNotes(
      slide,
      "Al abrir la aplicación, los precios nuevos detectados se presentan en una lista. El aliado puede seleccionar todos, limpiar la selección o enviar únicamente los artículos elegidos a impresión.",
      ["Captura: 01_precios_nuevos.png, detección automática observada el 31/07/2026."],
    );
  }

  // 6. Promociones vigentes
  {
    const slide = presentation.slides.add();
    slide.background.fill = GRAY;
    addHeader(slide, "Promociones vigentes listas para imprimir", 6);
    await addLogo(slide);
    const shot = { left: 115, top: 130, width: 1050, height: 564 };
    addScreenshotFrame(slide, shot);
    await addImage(slide, path.join(SCREENS, "05_precios_especiales.png"), shot, "Lista de precios especiales vigentes");
    addRect(slide, { left: 150, top: 536, width: 980, height: 88 }, "#171717E8", {
      geometry: "roundRect",
      borderRadius: 12,
    });
    addText(slide, "Precio regular, precio especial, existencia y vigencia en una sola revisión.", { left: 180, top: 560, width: 920, height: 46 }, {
      fontSize: 24,
      bold: true,
      color: WHITE,
      alignment: "center",
    });
    addNotes(
      slide,
      "El módulo de precios especiales vigentes reúne precio regular, precio especial, disponibilidad y periodo promocional. La selección múltiple facilita preparar las etiquetas necesarias.",
      ["Captura: 05_precios_especiales.png, aplicación local observada el 31/07/2026."],
    );
  }

  // 7. Etiqueta promocional
  {
    const slide = presentation.slides.add();
    slide.background.fill = WHITE;
    addHeader(slide, "La promoción comunica el ahorro", 7);
    await addLogo(slide);
    addText(slide, "Precio nuevo", { left: 80, top: 182, width: 390, height: 46 }, {
      fontSize: 31,
      bold: true,
      color: RED,
    });
    addText(slide, "Ahorro visible", { left: 80, top: 254, width: 390, height: 46 }, {
      fontSize: 31,
      bold: true,
      color: RED,
    });
    addText(slide, "Precio anterior", { left: 80, top: 326, width: 390, height: 46 }, {
      fontSize: 31,
      bold: true,
      color: RED,
    });
    addText(slide, "Vigencia y condiciones", { left: 80, top: 398, width: 390, height: 46 }, {
      fontSize: 31,
      bold: true,
      color: RED,
    });
    addRect(slide, { left: 80, top: 478, width: 360, height: 2 }, BLACK);
    addText(slide, "El formato destaca la promoción sin perder la información operativa.", { left: 80, top: 510, width: 390, height: 100 }, {
      fontSize: 22,
      color: MID,
    });
    const shot = { left: 535, top: 130, width: 650, height: 485 };
    addScreenshotFrame(slide, shot);
    await addImage(slide, path.join(SCREENS, "07_vista_previa_promocion.png"), shot, "Vista previa de etiqueta promocional");
    addText(slide, "Ejemplo de demostración", { left: 535, top: 646, width: 650, height: 26 }, {
      fontSize: 16,
      color: MID,
      alignment: "center",
    });
    addNotes(
      slide,
      "Cuando existe un precio especial, la vista previa promocional integra el nuevo precio, el ahorro, el precio anterior y las condiciones de vigencia en un formato de alta visibilidad.",
      ["Captura: 07_vista_previa_promocion.png, vista previa local generada el 31/07/2026."],
    );
  }

  // 8. Cierre
  {
    const slide = presentation.slides.add();
    slide.background.fill = WHITE;
    addRect(slide, { left: 0, top: 0, width: W, height: 16 }, RED);
    await addLogo(slide, { left: 78, top: 64, width: 270, height: 68 });
    addText(slide, "Un flujo claro de principio a fin", { left: 78, top: 190, width: 780, height: 72 }, {
      fontSize: 52,
      bold: true,
      typeface: FONT_DISPLAY,
    });
    addText(slide, "BUSCA", { left: 78, top: 330, width: 220, height: 60 }, {
      fontSize: 38,
      bold: true,
      color: RED,
    });
    addText(slide, "REVISA", { left: 335, top: 330, width: 230, height: 60 }, {
      fontSize: 38,
      bold: true,
      color: RED,
    });
    addText(slide, "SELECCIONA", { left: 605, top: 330, width: 300, height: 60 }, {
      fontSize: 38,
      bold: true,
      color: RED,
    });
    addText(slide, "IMPRIME", { left: 950, top: 330, width: 240, height: 60 }, {
      fontSize: 38,
      bold: true,
      color: RED,
    });
    addRect(slide, { left: 78, top: 430, width: 1110, height: 3 }, BLACK);
    addText(slide, "Etiquetador 80 mm centraliza la consulta, la validación visual y la salida de etiquetas para la operación diaria.", { left: 78, top: 476, width: 980, height: 102 }, {
      fontSize: 27,
      color: MID,
    });
    addText(slide, "CÍRCULO K MÉXICO", { left: 78, top: 642, width: 320, height: 30 }, {
      fontSize: 18,
      bold: true,
      color: RED,
    });
    addNotes(
      slide,
      "En resumen, el Etiquetador 80 milímetros ofrece un recorrido claro: buscar, revisar, seleccionar e imprimir. Una herramienta centralizada para preparar etiquetas en la operación diaria.",
      [
        "Logo: LOGO CIRCLE K.png, recurso local proporcionado en el proyecto.",
        "Síntesis funcional basada en el recorrido local observado el 31/07/2026.",
      ],
    );
  }

  for (const [index, slide] of presentation.slides.items.entries()) {
    const number = String(index + 1).padStart(2, "0");
    const png = await presentation.export({ slide, format: "png", scale: 1 });
    await fs.writeFile(path.join(WORK, "rendered", `slide-${number}.png`), new Uint8Array(await png.arrayBuffer()));
    const layout = await slide.export({ format: "layout" });
    await fs.writeFile(path.join(WORK, "rendered", `slide-${number}.layout.json`), await layout.text());
  }

  const montage = await presentation.export({ format: "webp", montage: true, scale: 1 });
  await fs.writeFile(path.join(WORK, "rendered", "montage.webp"), new Uint8Array(await montage.arrayBuffer()));

  const pptx = await PresentationFile.exportPptx(presentation);
  await pptx.save(OUTPUT_PPTX);
  console.log(OUTPUT_PPTX);
}

build().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
