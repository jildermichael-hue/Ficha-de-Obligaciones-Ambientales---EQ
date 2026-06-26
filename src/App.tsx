/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useRef } from "react";
import { GoogleGenAI, Type } from "@google/genai";
import { motion, AnimatePresence } from "motion/react";
import { 
  FileText, 
  Upload, 
  Download, 
  AlertCircle, 
  Loader2, 
  CheckCircle2,
  Trash2,
  Info,
  PlusCircle,
  Sparkles,
  Cpu,
  Copy,
  ExternalLink,
  ArrowLeft,
  Check
} from "lucide-react";
import { saveAs } from "file-saver";
import { 
  Document, 
  Packer, 
  Paragraph, 
  TextRun, 
  Table, 
  TableRow, 
  TableCell, 
  WidthType, 
  BorderStyle, 
  AlignmentType, 
  VerticalAlign,
  ShadingType,
  PageOrientation
} from "docx";
import { cn } from "./lib/utils";
import * as pdfjsLib from "pdfjs-dist";

// Set worker source for pdfjs
const PDFJS_VERSION = "4.0.379";
pdfjsLib.GlobalWorkerOptions.workerSrc = `https://cdnjs.cloudflare.com/ajax/libs/pdf.js/${PDFJS_VERSION}/pdf.worker.min.mjs`;

const SYSTEM_PROMPT = `
# PROFESIÓN / ROL
Actúa como especialista senior en derecho ambiental y fiscalización ambiental, con experiencia en:
- Normas generales de gestión ambiental.
- Normas sectoriales (minería, hidrocarburos, electricidad, industria, residuos, pesca, etc.).
- Diseño de matrices de obligaciones fiscalizables para autoridades de fiscalización ambiental.

Tu enfoque debe ser técnico, jurídico y muy estructurado, pero explicado de forma clara y operativa para que pueda usarse directamente en supervisiones y fiscalizaciones.

# ACTIVIDAD / TAREA
Tu tarea principal es identificar y sistematizar las obligaciones ambientales fiscalizables contenidas en las normas generales y sectoriales que te proporcionaré, organizándolas en una matriz clara y utilizable para fiscalización.

En concreto debes:
a) Leer y analizar el texto normativo que te entregue.
b) Detectar obligaciones ambientales (lo que el administrado debe hacer, no debe hacer o debe reportar).
c) Determinar si dichas obligaciones son fiscalizables por una autoridad de fiscalización ambiental.
d) Clasificar cada obligación según: Norma y artículo, Tipo de obligación, Sector o actividad, Sujeto obligado.
e) Presentar el resultado en una matriz o tabla lista para ser utilizada en fiscalización.

# CONTEXTO / INFORMACIÓN
- País: Perú.
- No inventes artículos ni normas.
- Explicar de manera operativa las obligaciones.
- Si la norma no es clara, indica dudas.

# MENTE / RAZONAMIENTO
1. Identificación de mandatos jurídicos: Busca verbos como debe, deberá, queda prohibido, etc.
2. Determinación de la fiscalizabilidad: ¿Puede ser verificada en inspección, revisión documental o monitoreo?
3. Clasificación técnico-jurídica:
   - Tipo: De resultado, De medio, De información.
   - Sector o actividad.
   - Sujeto obligado.
   - Instrumento asociado (EIA, plan de manejo, etc.).
4. Verificación interna: Referencia normativa exacta.

# FORMATO DE SALIDA
Debes devolver un objeto JSON que contenga los datos necesarios para construir la ficha de obligaciones.
El JSON debe tener esta estructura:
{
  "informacionGeneral": {
    "empresa": "Nombre de la empresa o [Completar]",
    "ruc": "RUC o [Completar]",
    "direccion": "Dirección o [Completar]",
    "area": "Área o [Completar]"
  },
  "fuentes": [
    {
      "nro": 1,
      "tipo": "Tipo de norma",
      "fuente": "Nombre de la norma",
      "autoridad": "Autoridad competente",
      "documento": "Documento de aprobación",
      "fecha": "Fecha de aprobación",
      "descripcion": "Breve descripción"
    }
  ],
  "obligaciones": [
    {
      "componenteGeneral": "NOMBRE GENERAL DEL COMPONENTE AMBIENTAL",
      "componente": "NOMBRE DEL COMPONENTE",
      "items": [
        {
          "referencia": "Nro. Fuente",
          "nro": "Nro. Obligación",
          "ubicacion": "Artículo/Numeral",
          "descripcion": "Descripción operativa de la obligación"
        }
      ]
    }
  ]
}
`;

const DISCLAIMER = "ESTE RESULTADO DEBE SER CORROBORADO POR EL EQUIPO DE LA EMPRESA, ES SOLO INFORMATIVO Y DEBE REVISARSE PARA EVITAR ERROR, ES RESPONSABILIDAD ABSOLUTA LA REVISIÓN FINAL, LA IA NO SE RESPONSABILIZA POR ERRORES";

const GALACTIC_PROMPT = `# PROFESIÓN / ROL
Actúa como especialista senior en derecho ambiental y fiscalización ambiental, con experiencia en:

	•	Normas generales de gestión ambiental.
	•	Normas sectoriales (minería, hidrocarburos, electricidad, industria, residuos, pesca, etc.).
	•	Diseño de matrices de obligaciones fiscalizables para autoridades de fiscalización ambiental.

Tu enfoque debe ser técnico, jurídico y muy estructurado, pero explicado de forma clara y operativa para que pueda usarse directamente en supervisiones y fiscalizaciones.

# ACTIVIDAD / TAREA
Tu tarea principal es identificar y sistematizar las obligaciones ambientales fiscalizables contenidas en las normas generales y sectoriales que te proporcionaré, organizándolas en una matriz clara y utilizable para fiscalización.

En concreto debes:
a) Leer y analizar el texto normativo que te entregue (leyes, reglamentos, decretos, resoluciones, etc.).
b) Detectar obligaciones ambientales (lo que el administrado debe hacer, no debe hacer o debe reportar).
c) Determinar si dichas obligaciones son fiscalizables por una autoridad de fiscalización ambiental.
d) Clasificar cada obligación según:
	•	Norma y artículo.
	•	Tipo de obligación (prevención, control, reporte, permisos/autorizaciones, manejo de residuos, emisiones, vertimientos, etc.).
	•	Sector o actividad (minería, hidrocarburos, electricidad, industria, otros).
	•	Sujeto obligado (titular de proyecto, operador, generador de residuos, etc.).
e) Presentar el resultado en una matriz o tabla lista para ser utilizada en fiscalización.

# CONTEXTO / INFORMACIÓN
Ten en cuenta lo siguiente:

El usuario te proporcionará uno o varios de estos elementos:

	•	Texto de normas generales de gestión ambiental.
	•	Texto de normas sectoriales aplicables a actividades específicas.
	•	Información sobre el país: Perú (trabaja sin inventar números de artículos).
	•	Información sobre el tipo de actividad o proyecto (por ejemplo: unidad minera a tajo abierto, planta industrial, lote de hidrocarburos, etc.).
	•	Tu análisis debe:
	•	Ceñirse exclusivamente al contenido normativo que se te entregue o identifique el usuario.
	•	Evitar inventar artículos, numeración normativa o normas inexistentes.
	•	Explicar de manera operativa las obligaciones, como insumos para acciones de supervisión/fiscalización.

Si la norma no es clara, indica las dudas o ambigüedades, pero no llenes los vacíos normativos con suposiciones.

# MENTE / RAZONAMIENTO
Antes de entregar la respuesta final, sigue internamente este proceso de razonamiento:

	1.	Identificación de mandatos jurídicos

	•	Recorre el texto buscando verbos como: debe, deberá, queda prohibido, se prohíbe, está obligado a, se exige, es requisito, no podrá, estará sujeto a, etc.
	•	Cada vez que identifiques un mandato, analiza si se refiere a un comportamiento ambiental (emisiones, vertimientos, residuos, ruido, permisos, reportes ambientales, monitoreos, etc.).

	2.	Determinación de la fiscalizabilidad

	•	Pregúntate: ¿esta obligación puede ser verificada en una inspección, revisión documental o monitoreo?
	•	Si la respuesta es sí, márcala como obligación fiscalizable.
	•	Si es una norma meramente declarativa o de principios generales sin mandato concreto, indícalo como no directamente fiscalizable.

	3.	Clasificación técnico-jurídica
Para cada obligación identificada, clasifícala indicando:

	•	Tipo de obligación:
	•	De resultado (lograr un límite, estándar, meta).
	•	De medio (implementar una medida, equipo, procedimiento).
	•	De información (reportar, monitorear, remitir informes).
	•	Sector o actividad al que aplica.
	•	Sujeto obligado (titular de proyecto, operador, concesionario, etc.).
	•	Instrumento asociado si corresponde (EIA, estudio ambiental, plan de manejo, etc.).

	4.	Verificación interna

	•	Verifica que cada obligación tenga referencia normativa exacta (nombre de la norma + artículo, numeral o literal).
	•	Revisa que la formulación de la obligación sea clara, verificable y operativa.

Solo después de este proceso, elabora la matriz o tablas de salida.

# AYUDA / FORMATO DE SALIDA
Entrega siempre la información organizada en una sola tabla en Word como se detalla aca:

0) Configuración de página (Word)
1.	Diseño > Tamaño: A4
2.	Diseño > Orientación: Horizontal
3.	Diseño > Márgenes > Márgenes personalizados: 1.5 cm (arriba/abajo/izq/der)
4.	Inicio > Párrafo:
o	Interlineado: Sencillo
o	Espaciado Antes: 0 pt / Después: 0 pt

Ancho útil aproximado de trabajo: 29.7 − 3.0 = 26.7 cm

1) Título superior (como en la captura)
1.	Escribe centrado: FICHA DE OBLIGACIONES AMBIENTALES FISCALIZABLES
- Fuente: Arial 16 (puede ser 18 si lo quieres más dominante)
- Negrita, MAYÚSCULA, centrado.
2.	Con el cursor en esa línea: Inicio > Bordes > Bordes y sombreado
- Borde superior: 0.75 pt (negro)
- Borde inferior: 2.25–3 pt (negro)  (esta es la línea gruesa)
Tip: si tu Word no te deja “borde superior” en el párrafo, usa un salto de línea arriba y aplica borde al párrafo completo.

2) Sección I. INFORMACIÓN GENERAL
A) Barra gris del título de sección
1.	Insertar > Tabla: 1 fila x 1 columna
2.	Escribe: I. INFORMACIÓN GENERAL
3.	Formato:
- Arial 12, negrita
- Sombreado gris claro: (Tabla > Diseño de tabla > Sombreado: Gris 15% o #D9D9D9)
- Bordes: superior e inferior negro 1.5 pt

B) Tabla de datos (Empresa / Dirección / RUC / Área)
1.	Insertar > Tabla: 2 filas x 4 columnas
2.	Anchos de columnas (muy importante):
-	Col 1: 4.2 cm
-	Col 2: 16.5 cm
-	Col 3: 2.5 cm
-	Col 4: 3.5 cm
(Total = 26.7 cm)

3.	Textos:
-	Fila 1: Empresa | (vacío) | R.U.C. | (vacío)
-	Fila 2: Dirección de la empresa | (vacío) | Área | (vacío)
4.	Bordes:
-	Marco exterior: 1.0 pt (gris oscuro o negro suave)
-	Líneas internas: 0.5 pt gris (#7F7F7F)

3) Sección II. FUENTE DE OBLIGACIONES FISCALIZABLES

A) Barra gris de sección
•	Igual que la Sección I, pero texto:
II. FUENTE DE OBLIGACIONES FISCALIZABLES

B) Tabla “DATOS DE LA FUENTE”
1.	Insertar > Tabla: 4 filas x 7 columnas
(1 fila para “DATOS DE LA FUENTE”, 1 fila encabezados, 2 filas en blanco)
2.	Fila 1: Combina las 7 celdas y escribe:
-	DATOS DE LA FUENTE (Arial 12, negrita, alineado izquierda)
3.	Fila 2 (encabezados):
1.	Nro. Fuente
2.	Tipo
3.	Fuente
4.	Autoridad Competente
5.	Documento de Aprobación
6.	Fecha de aprobación
7.	Descripción
4.	Anchos de columnas (para que calce como la captura):
-	C1: 1.8 cm
-	C2: 2.3 cm
-	C3: 4.0 cm
-	C4: 4.0 cm
-	C5: 4.2 cm
-	C6: 3.0 cm
-	C7: 7.4 cm
(Total = 26.7 cm)
5.	Bordes:
-	Exterior: 1.0 pt gris oscuro
-	Internos: 0.5 pt gris
-	Encabezados (fila 2) con borde inferior 1.0 pt gris oscuro
6.	Alturas sugeridas:
•	Fila 1 (DATOS DE LA FUENTE): 0.9 cm
•	Fila 2 (encabezados): 1.0 cm
•	Filas en blanco: 1.5–1.8 cm

4) Sección III. OBLIGACIONES FISCALIZABLES

A) Barra gris de sección
•	Igual: III. OBLIGACIONES FISCALIZABLES
B) Párrafo debajo (texto de la captura)
Pega el texto y formatea:
•	Arial 11 o 12
•	Justificado
•	Sin sangría, sin espaciados antes/después
C) Tabla principal de obligaciones
1.	Insertar > Tabla: 6 filas x 4 columnas
(1 encabezado + 1 fila “3.1” + 1 fila “3.1.1” + 3 filas en blanco)
2.	Encabezados (fila 1):
1.	Referencia (Nro. Fuente)
2.	Nro. Obligación
3.	Ubicación (artículo de la norma)
4.	Descripción de la Obligación
3.	Anchos de columnas:
-	Col 1: 3.0 cm
-	Col 2: 3.0 cm
-	Col 3: 6.0 cm
-	Col 4: 14.7 cm
(Total = 26.7 cm)
4.	Fila 2: combinar 4 celdas y poner:
- 3.1 [NOMBRE GENERAL DEL COMPONENTE AMBIENTAL]
- Negrita
5.	Fila 3: combinar 4 celdas y poner:
- 3.1.1 [NOMBRE DEL COMPONENTE]
- Negrita
6.	Bordes recomendados:
-	Tabla (exterior): 1.0 pt gris oscuro
-	Internos: 0.5 pt gris
-	Línea negra marcada entre “3.1” y “3.1.1”: aplica borde inferior 1.5–2 pt negro en la fila “3.1.1” (para el efecto de separador fuerte como la captura)

5) Ajustes finos para que se vea “pro”
•	Alineación vertical de celdas: Centrado (Tabla > Presentación > Alineación)
•	Márgenes internos de celda: 0.15 cm (Tabla > Propiedades > Opciones)
•	Encabezados centrados y en negrita.
•	Mantén el gris solo en las barras I/II/III, no en toda la tabla.

# CONDICIONES / LO QUE NO QUIERO
Respeta estrictamente estas condiciones:

	•	No inventes artículos, numerales ni nombres de normas.
	•	No quiero que me lo des por partes, dame todo en una sola tabla.
	•	No des respuestas genéricas del tipo “depende de la normativa aplicable” sin intentar trabajar con el texto que se te entregue.
	•	No te extiendas en definiciones doctrinarias o teóricas de derecho ambiental, salvo que el usuario lo pida expresamente.
	•	No mezcles obligaciones de diferentes países sin que se indique claramente.
	•	No repitas innecesariamente el texto legal completo; prioriza su traducción a lenguaje operativo de fiscalización.

En cambio, sí debes:
	•	Citar siempre norma + artículo/numeral cuando esté disponible en el texto proporcionado.
	•	Priorizar la claridad y utilidad práctica para la fiscalización ambiental.
	•	Señalar las dudas, vacíos o ambigüedades de la norma cuando existan.

# ORDEN FINAL 1: Pídele al usuario que te envié todas las normas con las que desea trabajar, con un tono amable y que sepa que estar para ayudarlo a cumplir el objetivo. 

# ORDEN FINAL 2: Todo tiene que ser devuelto en formato word 

# ORDEN FINAL 3: Al inicio y final de la respuesta o documento agrega este mensaje: “ESTE RESULTADO DEBE SER CORROBORADO POR EL EQUIPO DE LA EMPRESA, ES SOLO INFORMATIVO Y DEBE REVISARSE PARA EVITAR ERROR, ES RESPONSABILIDAD ABSOLUTA LA REVISIÓN FINAL, LA IA NO SE RESPONSABILIZA POR ERRORES” en negrita y cursiva.`;

export default function App() {
  const [currentScreen, setCurrentScreen] = useState<"home" | "engineer" | "prompt">("home");
  const [selectedAI, setSelectedAI] = useState<"chatgpt" | "gemini" | "claude" | "deepseek" | null>(null);
  const [copied, setCopied] = useState(false);

  const [inputText, setInputText] = useState("");
  const [isProcessing, setIsProcessing] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [result, setResult] = useState<any>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Manual Info State
  const [generalInfo, setGeneralInfo] = useState({
    empresa: "",
    ruc: "",
    direccion: "",
    area: "",
    sector: "Minería" // Default sector
  });

  const SECTORS = [
    "Minería",
    "Hidrocarburos",
    "Electricidad",
    "Industria",
    "Residuos Sólidos",
    "Pesca",
    "Agricultura",
    "Transportes",
    "Vivienda y Construcción",
    "Otros"
  ];

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    if (file.type === "application/pdf") {
      setIsProcessing(true);
      setError(null);
      try {
        const arrayBuffer = await file.arrayBuffer();
        const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
        let fullText = "";
        
        for (let i = 1; i <= pdf.numPages; i++) {
          const page = await pdf.getPage(i);
          const textContent = await page.getTextContent();
          const pageText = textContent.items
            .map((item: any) => {
              if ("str" in item) return item.str;
              return "";
            })
            .join(" ");
          fullText += pageText + "\n";
        }
        
        setInputText(fullText);
      } catch (err) {
        console.error("Error reading PDF:", err);
        setError("No se pudo leer el archivo PDF. Asegúrate de que no esté protegido.");
      } finally {
        setIsProcessing(false);
      }
    } else {
      const reader = new FileReader();
      reader.onload = (event) => {
        const text = event.target?.result as string;
        setInputText(text);
      };
      reader.readAsText(file);
    }
  };

  const processNorms = async (isAppending = false) => {
    if (!inputText.trim()) {
      setError("Por favor, ingresa o sube el texto de las normas.");
      return;
    }

    setIsProcessing(true);
    setError(null);

    try {
      const apiKey = process.env.GEMINI_API_KEY;
      if (!apiKey) {
        throw new Error("API Key de Gemini no encontrada. Por favor, configúrala en el panel de Secretos.");
      }

      const ai = new GoogleGenAI({ apiKey });
      
      let prompt = `Analiza las siguientes normas para una empresa del sector ${generalInfo.sector} en el área de ${generalInfo.area}. 
      Empresa: ${generalInfo.empresa} (RUC: ${generalInfo.ruc}).
      
      Extrae las obligaciones ambientales fiscalizables:\n\n${inputText}`;
      
      if (isAppending && result) {
        prompt = `Ya hemos identificado algunas obligaciones para la empresa ${generalInfo.empresa} del sector ${generalInfo.sector}. 
        Por favor, realiza un análisis EXHAUSTIVO y A FONDO del texto para identificar OTRAS obligaciones que no hayan sido listadas anteriormente. 
        Busca detalles específicos, numerales secundarios o requisitos técnicos que se hayan pasado por alto.
        
        Obligaciones ya identificadas (NO LAS REPITAS): ${JSON.stringify(result.obligaciones.map((g: any) => g.items.map((i: any) => i.descripcion)))}
        
        Texto de las normas:\n\n${inputText}`;
      }

      const response = await ai.models.generateContent({
        model: "gemini-3-flash-preview",
        contents: prompt,
        config: {
          systemInstruction: SYSTEM_PROMPT,
          responseMimeType: "application/json",
          responseSchema: {
            type: Type.OBJECT,
            properties: {
              informacionGeneral: {
                type: Type.OBJECT,
                properties: {
                  empresa: { type: Type.STRING },
                  ruc: { type: Type.STRING },
                  direccion: { type: Type.STRING },
                  area: { type: Type.STRING },
                  sector: { type: Type.STRING },
                },
                required: ["empresa", "ruc", "direccion", "area", "sector"]
              },
              fuentes: {
                type: Type.ARRAY,
                items: {
                  type: Type.OBJECT,
                  properties: {
                    nro: { type: Type.NUMBER },
                    tipo: { type: Type.STRING },
                    fuente: { type: Type.STRING },
                    autoridad: { type: Type.STRING },
                    documento: { type: Type.STRING },
                    fecha: { type: Type.STRING },
                    descripcion: { type: Type.STRING },
                  },
                  required: ["nro", "tipo", "fuente", "autoridad", "documento", "fecha", "descripcion"]
                }
              },
              obligaciones: {
                type: Type.ARRAY,
                items: {
                  type: Type.OBJECT,
                  properties: {
                    componenteGeneral: { type: Type.STRING },
                    componente: { type: Type.STRING },
                    items: {
                      type: Type.ARRAY,
                      items: {
                        type: Type.OBJECT,
                        properties: {
                          referencia: { type: Type.STRING },
                          nro: { type: Type.STRING },
                          ubicacion: { type: Type.STRING },
                          descripcion: { type: Type.STRING },
                        },
                        required: ["referencia", "nro", "ubicacion", "descripcion"]
                      }
                    }
                  },
                  required: ["componenteGeneral", "componente", "items"]
                }
              }
            },
            required: ["informacionGeneral", "fuentes", "obligaciones"]
          }
        }
      });

      const responseText = response.text || "{}";
      const cleanJson = responseText.replace(/```json\n?|\n?```/g, "").trim();
      const data = JSON.parse(cleanJson);
      
      if (isAppending && result) {
        // Merge obligations
        const mergedObligaciones = [...result.obligaciones];
        data.obligaciones.forEach((newGroup: any) => {
          const existingGroup = mergedObligaciones.find(g => g.componente === newGroup.componente);
          if (existingGroup) {
            existingGroup.items = [...existingGroup.items, ...newGroup.items];
          } else {
            mergedObligaciones.push(newGroup);
          }
        });
        
        // Merge sources
        const mergedFuentes = [...result.fuentes];
        data.fuentes.forEach((newFuente: any) => {
          if (!mergedFuentes.find(f => f.fuente === newFuente.fuente)) {
            mergedFuentes.push(newFuente);
          }
        });

        setResult({
          ...result,
          fuentes: mergedFuentes,
          obligaciones: mergedObligaciones
        });
      } else {
        // Use manual info if provided, otherwise use AI's
        const finalData = {
          ...data,
          informacionGeneral: {
            empresa: generalInfo.empresa || data.informacionGeneral.empresa,
            ruc: generalInfo.ruc || data.informacionGeneral.ruc,
            direccion: generalInfo.direccion || data.informacionGeneral.direccion,
            area: generalInfo.area || data.informacionGeneral.area,
            sector: generalInfo.sector || data.informacionGeneral.sector,
          }
        };
        setResult(finalData);
        // Update manual info state with what AI found if it was empty
        setGeneralInfo(finalData.informacionGeneral);
      }
    } catch (err: any) {
      console.error("Error processing norms:", err);
      setError(err.message || "Ocurrió un error al procesar las normas. Por favor, intenta de nuevo.");
    } finally {
      setIsProcessing(false);
    }
  };

  const generateWord = async () => {
    if (!result) return;

    const CM_TO_TWIP = 567;

    const doc = new Document({
      sections: [
        {
          properties: {
            page: {
              size: {
                orientation: PageOrientation.LANDSCAPE,
              },
              margin: {
                top: 1.5 * CM_TO_TWIP,
                bottom: 1.5 * CM_TO_TWIP,
                left: 1.5 * CM_TO_TWIP,
                right: 1.5 * CM_TO_TWIP,
              },
            },
          },
          children: [
            // Disclaimer Header
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({
                  text: DISCLAIMER,
                  bold: true,
                  italics: true,
                  size: 20,
                  font: "Arial",
                }),
              ],
            }),
            new Paragraph({ text: "" }),

            // Main Title
            new Paragraph({
              alignment: AlignmentType.CENTER,
              border: {
                top: { color: "000000", size: 6, style: BorderStyle.SINGLE },
                bottom: { color: "000000", size: 18, style: BorderStyle.SINGLE },
              },
              children: [
                new TextRun({
                  text: "FICHA DE OBLIGACIONES AMBIENTALES FISCALIZABLES",
                  bold: true,
                  size: 32, // Arial 16
                  font: "Arial",
                }),
              ],
            }),
            new Paragraph({ text: "" }),

            // Section I: Información General
            new Table({
              width: { size: 26.7 * CM_TO_TWIP, type: WidthType.DXA },
              rows: [
                new TableRow({
                  children: [
                    new TableCell({
                      shading: { fill: "D9D9D9", type: ShadingType.CLEAR },
                      borders: {
                        top: { style: BorderStyle.SINGLE, size: 12, color: "000000" },
                        bottom: { style: BorderStyle.SINGLE, size: 12, color: "000000" },
                      },
                      children: [
                        new Paragraph({
                          children: [new TextRun({ text: "I. INFORMACIÓN GENERAL", bold: true, font: "Arial", size: 24 })],
                        }),
                      ],
                    }),
                  ],
                }),
              ],
            }),
            new Table({
              width: { size: 26.7 * CM_TO_TWIP, type: WidthType.DXA },
              rows: [
                new TableRow({
                  children: [
                    new TableCell({ width: { size: 4.2 * CM_TO_TWIP, type: WidthType.DXA }, children: [new Paragraph({ children: [new TextRun({ text: "Empresa", font: "Arial", bold: true })] })] }),
                    new TableCell({ width: { size: 16.5 * CM_TO_TWIP, type: WidthType.DXA }, children: [new Paragraph({ children: [new TextRun({ text: generalInfo.empresa || result.informacionGeneral.empresa || "", font: "Arial" })] })] }),
                    new TableCell({ width: { size: 2.5 * CM_TO_TWIP, type: WidthType.DXA }, children: [new Paragraph({ children: [new TextRun({ text: "R.U.C.", font: "Arial", bold: true })] })] }),
                    new TableCell({ width: { size: 3.5 * CM_TO_TWIP, type: WidthType.DXA }, children: [new Paragraph({ children: [new TextRun({ text: generalInfo.ruc || result.informacionGeneral.ruc || "", font: "Arial" })] })] }),
                  ],
                }),
                new TableRow({
                  children: [
                    new TableCell({ width: { size: 4.2 * CM_TO_TWIP, type: WidthType.DXA }, children: [new Paragraph({ children: [new TextRun({ text: "Dirección de la empresa", font: "Arial", bold: true })] })] }),
                    new TableCell({ width: { size: 10.0 * CM_TO_TWIP, type: WidthType.DXA }, children: [new Paragraph({ children: [new TextRun({ text: generalInfo.direccion || result.informacionGeneral.direccion || "", font: "Arial" })] })] }),
                    new TableCell({ width: { size: 2.5 * CM_TO_TWIP, type: WidthType.DXA }, children: [new Paragraph({ children: [new TextRun({ text: "Sector", font: "Arial", bold: true })] })] }),
                    new TableCell({ width: { size: 4.0 * CM_TO_TWIP, type: WidthType.DXA }, children: [new Paragraph({ children: [new TextRun({ text: generalInfo.sector || result.informacionGeneral.sector || "", font: "Arial" })] })] }),
                    new TableCell({ width: { size: 2.5 * CM_TO_TWIP, type: WidthType.DXA }, children: [new Paragraph({ children: [new TextRun({ text: "Área", font: "Arial", bold: true })] })] }),
                    new TableCell({ width: { size: 3.5 * CM_TO_TWIP, type: WidthType.DXA }, children: [new Paragraph({ children: [new TextRun({ text: generalInfo.area || result.informacionGeneral.area || "", font: "Arial" })] })] }),
                  ],
                }),
              ],
            }),
            new Paragraph({ text: "" }),

            // Section II: Fuente de Obligaciones
            new Table({
              width: { size: 26.7 * CM_TO_TWIP, type: WidthType.DXA },
              rows: [
                new TableRow({
                  children: [
                    new TableCell({
                      shading: { fill: "D9D9D9", type: ShadingType.CLEAR },
                      borders: {
                        top: { style: BorderStyle.SINGLE, size: 12, color: "000000" },
                        bottom: { style: BorderStyle.SINGLE, size: 12, color: "000000" },
                      },
                      children: [
                        new Paragraph({
                          children: [new TextRun({ text: "II. FUENTE DE OBLIGACIONES FISCALIZABLES", bold: true, font: "Arial", size: 24 })],
                        }),
                      ],
                    }),
                  ],
                }),
              ],
            }),
            new Table({
              width: { size: 26.7 * CM_TO_TWIP, type: WidthType.DXA },
              rows: [
                new TableRow({
                  children: [
                    new TableCell({ columnSpan: 7, children: [new Paragraph({ children: [new TextRun({ text: "DATOS DE LA FUENTE", bold: true, font: "Arial", size: 24 })] })] }),
                  ],
                }),
                new TableRow({
                  children: [
                    new TableCell({ width: { size: 1.8 * CM_TO_TWIP, type: WidthType.DXA }, children: [new Paragraph({ children: [new TextRun({ text: "Nro. Fuente", bold: true, font: "Arial" })] })] }),
                    new TableCell({ width: { size: 2.3 * CM_TO_TWIP, type: WidthType.DXA }, children: [new Paragraph({ children: [new TextRun({ text: "Tipo", bold: true, font: "Arial" })] })] }),
                    new TableCell({ width: { size: 4.0 * CM_TO_TWIP, type: WidthType.DXA }, children: [new Paragraph({ children: [new TextRun({ text: "Fuente", bold: true, font: "Arial" })] })] }),
                    new TableCell({ width: { size: 4.0 * CM_TO_TWIP, type: WidthType.DXA }, children: [new Paragraph({ children: [new TextRun({ text: "Autoridad Competente", bold: true, font: "Arial" })] })] }),
                    new TableCell({ width: { size: 4.2 * CM_TO_TWIP, type: WidthType.DXA }, children: [new Paragraph({ children: [new TextRun({ text: "Documento de Aprobación", bold: true, font: "Arial" })] })] }),
                    new TableCell({ width: { size: 3.0 * CM_TO_TWIP, type: WidthType.DXA }, children: [new Paragraph({ children: [new TextRun({ text: "Fecha de aprobación", bold: true, font: "Arial" })] })] }),
                    new TableCell({ width: { size: 7.4 * CM_TO_TWIP, type: WidthType.DXA }, children: [new Paragraph({ children: [new TextRun({ text: "Descripción", bold: true, font: "Arial" })] })] }),
                  ],
                }),
                ...result.fuentes.map((f: any) => new TableRow({
                  children: [
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: String(f.nro), font: "Arial" })] })] }),
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: f.tipo, font: "Arial" })] })] }),
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: f.fuente, font: "Arial" })] })] }),
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: f.autoridad, font: "Arial" })] })] }),
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: f.documento, font: "Arial" })] })] }),
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: f.fecha, font: "Arial" })] })] }),
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: f.descripcion, font: "Arial" })] })] }),
                  ],
                })),
              ],
            }),
            new Paragraph({ text: "" }),

            // Section III: Obligaciones
            new Table({
              width: { size: 26.7 * CM_TO_TWIP, type: WidthType.DXA },
              rows: [
                new TableRow({
                  children: [
                    new TableCell({
                      shading: { fill: "D9D9D9", type: ShadingType.CLEAR },
                      borders: {
                        top: { style: BorderStyle.SINGLE, size: 12, color: "000000" },
                        bottom: { style: BorderStyle.SINGLE, size: 12, color: "000000" },
                      },
                      children: [
                        new Paragraph({
                          children: [new TextRun({ text: "III. OBLIGACIONES FISCALIZABLES", bold: true, font: "Arial", size: 24 })],
                        }),
                      ],
                    }),
                  ],
                }),
              ],
            }),
            new Paragraph({
              alignment: AlignmentType.JUSTIFIED,
              children: [
                new TextRun({
                  text: "A continuación se detallan las obligaciones ambientales fiscalizables identificadas en la normativa analizada.",
                  font: "Arial",
                }),
              ],
            }),
            new Table({
              width: { size: 26.7 * CM_TO_TWIP, type: WidthType.DXA },
              rows: [
                new TableRow({
                  children: [
                    new TableCell({ width: { size: 3.0 * CM_TO_TWIP, type: WidthType.DXA }, children: [new Paragraph({ children: [new TextRun({ text: "Referencia (Nro. Fuente)", bold: true, font: "Arial" })] })] }),
                    new TableCell({ width: { size: 3.0 * CM_TO_TWIP, type: WidthType.DXA }, children: [new Paragraph({ children: [new TextRun({ text: "Nro. Obligación", bold: true, font: "Arial" })] })] }),
                    new TableCell({ width: { size: 6.0 * CM_TO_TWIP, type: WidthType.DXA }, children: [new Paragraph({ children: [new TextRun({ text: "Ubicación (artículo de la norma)", bold: true, font: "Arial" })] })] }),
                    new TableCell({ width: { size: 14.7 * CM_TO_TWIP, type: WidthType.DXA }, children: [new Paragraph({ children: [new TextRun({ text: "Descripción de la Obligación", bold: true, font: "Arial" })] })] }),
                  ],
                }),
                ...result.obligaciones.flatMap((group: any, index: number) => [
                  new TableRow({
                    children: [
                      new TableCell({ columnSpan: 4, children: [new Paragraph({ children: [new TextRun({ text: `3.${index + 1} ${group.componenteGeneral}`, bold: true, font: "Arial" })] })] }),
                    ],
                  }),
                  new TableRow({
                    children: [
                      new TableCell({ 
                        columnSpan: 4, 
                        borders: { bottom: { style: BorderStyle.SINGLE, size: 12, color: "000000" } },
                        children: [new Paragraph({ children: [new TextRun({ text: `3.${index + 1}.1 ${group.componente}`, bold: true, font: "Arial" })] })] 
                      }),
                    ],
                  }),
                  ...group.items.map((item: any) => new TableRow({
                    children: [
                      new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: item.referencia, font: "Arial" })] })] }),
                      new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: item.nro, font: "Arial" })] })] }),
                      new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: item.ubicacion, font: "Arial" })] })] }),
                      new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: item.descripcion, font: "Arial" })] })] }),
                    ],
                  })),
                ]),
              ],
            }),
            new Paragraph({ text: "" }),

            // Footer Disclaimer
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({
                  text: DISCLAIMER,
                  bold: true,
                  italics: true,
                  size: 20,
                  font: "Arial",
                }),
              ],
            }),
          ],
        },
      ],
    });

    const blob = await Packer.toBlob(doc);
    saveAs(blob, "Ficha_de_Obligaciones_Ambientales.docx");
  };

  const renderHomeScreen = () => {
    return (
      <div className="max-w-4xl mx-auto space-y-8 -mt-10">
        <div className="text-center space-y-3">
          <motion.div
            initial={{ scale: 0.8, opacity: 0 }}
            animate={{ scale: 1, opacity: 1 }}
            className="w-16 h-16 mx-auto rounded-2xl bg-gradient-to-br from-brand-blue to-brand-purple flex items-center justify-center text-white shadow-xl shadow-brand-blue/20"
          >
            <Sparkles size={32} />
          </motion.div>
          <motion.h2
            initial={{ opacity: 0, y: 15 }}
            animate={{ opacity: 1, y: 0 }}
            className="text-3xl sm:text-4xl font-extrabold text-slate-900 tracking-tight"
          >
            Ficha de Obligaciones Ambientales EQ
          </motion.h2>
          <motion.p
            initial={{ opacity: 0, y: 15 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ delay: 0.1 }}
            className="text-base text-slate-600 max-w-xl mx-auto leading-relaxed"
          >
            Selecciona el método de tu preferencia para analizar y extraer tus obligaciones ambientales fiscalizables.
          </motion.p>
        </div>

        {/* Two Paths Selection Grid */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-8 max-w-4xl mx-auto pt-4">
          {/* Path 1: App del Ingeniero */}
          <motion.div
            whileHover={{ y: -4, scale: 1.005 }}
            onClick={() => setCurrentScreen("engineer")}
            className="bg-white rounded-3xl p-8 border border-slate-200 shadow-xl shadow-slate-150/40 flex flex-col justify-between h-full space-y-6 text-left cursor-pointer group hover:border-brand-blue/30 transition-all"
          >
            <div className="space-y-4">
              <div className="w-14 h-14 rounded-2xl bg-brand-blue/10 flex items-center justify-center text-brand-blue group-hover:bg-brand-blue group-hover:text-white transition-colors">
                <Cpu size={28} />
              </div>
              <h3 className="text-2xl font-bold text-slate-900">Usa la APP del Ingeniero</h3>
              <p className="text-slate-600 text-sm leading-relaxed">
                Nuestra herramienta inteligente e interactiva. Sube normas ambientales en formato PDF o texto plano, ingresa los datos de tu empresa, área y sector, y deja que la IA especializada redacte y catalogue las obligaciones en una matriz descargable en Word con formato profesional.
              </p>
            </div>
            <button className="w-full py-4 px-6 rounded-2xl bg-brand-blue text-white font-bold text-center group-hover:bg-brand-blue/90 transition-all shadow-md shadow-brand-blue/10">
              Iniciar App del Ingeniero
            </button>
          </motion.div>

          {/* Path 2: Prompt Galáctico */}
          <motion.div
            whileHover={{ y: -4, scale: 1.005 }}
            onClick={() => setCurrentScreen("prompt")}
            className="bg-white rounded-3xl p-8 border border-slate-200 shadow-xl shadow-slate-150/40 flex flex-col justify-between h-full space-y-6 text-left cursor-pointer group hover:border-brand-purple/30 transition-all"
          >
            <div className="space-y-4">
              <div className="w-14 h-14 rounded-2xl bg-brand-purple/10 flex items-center justify-center text-brand-purple group-hover:bg-brand-purple group-hover:text-white transition-colors">
                <Sparkles size={28} />
              </div>
              <h3 className="text-2xl font-bold text-slate-900">Usa el Prompt Galáctico</h3>
              <p className="text-slate-600 text-sm leading-relaxed">
                Obtén acceso directo al prompt de ingeniería de instrucciones ultra-especializado diseñado para derecho ambiental. Elige tu IA preferida (ChatGPT, Gemini, Claude, DeepSeek) y copia la estructura profesional de forma instantánea.
              </p>
            </div>
            <button className="w-full py-4 px-6 rounded-2xl bg-brand-purple text-white font-bold text-center group-hover:bg-brand-purple/90 transition-all shadow-md shadow-brand-purple/10">
              Ver Prompt Galáctico
            </button>
          </motion.div>
        </div>

        {/* Informative message box */}
        <motion.div
          initial={{ opacity: 0, y: 10 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ delay: 0.2 }}
          className="p-5 rounded-2xl bg-amber-50 border border-amber-200/80 text-amber-850 flex gap-4 items-start shadow-sm max-w-2xl mx-auto text-left"
        >
          <Info className="text-amber-550 shrink-0 mt-0.5" size={22} />
          <div>
            <h4 className="font-bold text-slate-950 text-sm">¿La App del Ingeniero está inactiva o sin límites?</h4>
            <p className="text-sm text-slate-600 mt-1">
              Si no está activa la app del ingeniero, puedes usar el <strong>Prompt Galáctico</strong> para copiar las instrucciones profesionales de análisis y utilizarlas con cualquier IA externa (ChatGPT, Gemini, Claude, DeepSeek).
            </p>
          </div>
        </motion.div>
      </div>
    );
  };

  const renderPromptScreen = () => {
    const aiList = [
      { id: "chatgpt", name: "ChatGPT", url: "https://chatgpt.com", color: "bg-[#10a37f]", textColor: "text-white", desc: "OpenAI GPT-4o / o1" },
      { id: "gemini", name: "Gemini", url: "https://gemini.google.com", color: "bg-[#1a73e8]", textColor: "text-white", desc: "Google Gemini Pro" },
      { id: "claude", name: "Claude", url: "https://claude.ai", color: "bg-[#d97706]", textColor: "text-white", desc: "Anthropic Claude 3.5" },
      { id: "deepseek", name: "DeepSeek", url: "https://chat.deepseek.com", color: "bg-[#0052cc]", textColor: "text-white", desc: "DeepSeek-V3 / Chat" }
    ] as const;

    const handleCopy = () => {
      navigator.clipboard.writeText(GALACTIC_PROMPT);
      setCopied(true);
      setTimeout(() => setCopied(false), 2500);
    };

    return (
      <div className="max-w-4xl mx-auto space-y-10 py-6 text-left">
        {/* Header */}
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 pb-6 border-b border-slate-200">
          <div>
            <button
              onClick={() => {
                setCurrentScreen("home");
                setSelectedAI(null);
              }}
              className="flex items-center gap-2 text-slate-500 hover:text-slate-900 font-semibold mb-2 text-sm transition-colors"
            >
              <ArrowLeft size={16} />
              Volver al Inicio
            </button>
            <h2 className="text-3xl font-bold text-slate-900">Usa el Prompt Galáctico</h2>
            <p className="text-slate-600 text-sm mt-1">
              Selecciona tu IA preferida para copiar el prompt de instrucciones especializado y empezar a trabajar de inmediato.
            </p>
          </div>
        </div>

        {/* AI Selector */}
        <div className="space-y-4">
          <h3 className="text-lg font-bold text-slate-800 flex items-center gap-2">
            <span className="w-2.5 h-2.5 rounded-full bg-brand-purple" />
            ¿Con qué IA quieres trabajar?
          </h3>
          <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
            {aiList.map((ai) => {
              const isSelected = selectedAI === ai.id;
              return (
                <button
                  key={ai.id}
                  onClick={() => setSelectedAI(ai.id)}
                  className={cn(
                    "p-6 rounded-2xl border text-center transition-all flex flex-col items-center justify-center gap-3 relative overflow-hidden",
                    isSelected
                      ? "border-brand-purple bg-purple-50/50 ring-2 ring-brand-purple/20 shadow-md"
                      : "border-slate-200 bg-white hover:border-slate-300 hover:shadow-md"
                  )}
                >
                  <div className={cn("w-12 h-12 rounded-xl flex items-center justify-center font-bold text-xl shadow-sm", ai.color, ai.textColor)}>
                    {ai.name[0]}
                  </div>
                  <div>
                    <h4 className="font-bold text-slate-900">{ai.name}</h4>
                    <span className="text-[10px] text-slate-500 block mt-0.5">{ai.desc}</span>
                  </div>
                  {isSelected && (
                    <span className="absolute top-2 right-2 w-5 h-5 rounded-full bg-brand-purple text-white flex items-center justify-center text-[10px] font-bold">
                      ✓
                    </span>
                  )}
                </button>
              );
            })}
          </div>
        </div>

        {/* Copy and Actions Area */}
        <AnimatePresence>
          {selectedAI && (
            <motion.div
              initial={{ opacity: 0, y: 15 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -15 }}
              className="bg-white rounded-3xl p-8 border border-slate-200 shadow-xl shadow-slate-200/50 space-y-6"
            >
              {/* Top Row with action buttons */}
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 pb-4 border-b border-slate-100">
                <div className="flex items-center gap-3">
                  <div className={cn(
                    "px-4 py-2 rounded-xl font-bold text-xs shadow-sm",
                    aiList.find(a => a.id === selectedAI)?.color,
                    aiList.find(a => a.id === selectedAI)?.textColor
                  )}>
                    {aiList.find(a => a.id === selectedAI)?.name} Seleccionado
                  </div>
                  <span className="text-slate-500 text-sm font-medium">¡Listo para usar!</span>
                </div>

                <div className="flex items-center gap-2">
                  {/* Copy Button */}
                  <button
                    onClick={handleCopy}
                    className={cn(
                      "flex items-center gap-2 px-5 py-3 rounded-xl font-bold transition-all text-sm shadow-md",
                      copied
                        ? "bg-emerald-500 text-white shadow-emerald-500/10"
                        : "bg-slate-900 text-white hover:bg-slate-800"
                    )}
                  >
                    {copied ? (
                      <>
                        <Check size={18} />
                        ¡Prompt Copiado!
                      </>
                    ) : (
                      <>
                        <Copy size={18} />
                        Copia el prompt
                      </>
                    )}
                  </button>

                  {/* Access Button */}
                  <a
                    href={aiList.find(a => a.id === selectedAI)?.url}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="flex items-center gap-2 px-5 py-3 rounded-xl border border-slate-200 bg-slate-50 hover:bg-slate-100 transition-all text-sm font-bold text-slate-700 shadow-sm"
                  >
                    <ExternalLink size={18} />
                    Acceder a {aiList.find(a => a.id === selectedAI)?.name}
                  </a>
                </div>
              </div>

              {/* Prompt box */}
              <div className="space-y-3">
                <div className="flex items-center justify-between">
                  <span className="text-xs font-bold text-slate-500 uppercase tracking-wider">Vista Previa del Prompt</span>
                  <span className="text-xs text-slate-400">Total: {GALACTIC_PROMPT.length} caracteres</span>
                </div>
                <div className="relative rounded-2xl bg-slate-50 border border-slate-200 p-6 overflow-hidden">
                  <textarea
                    readOnly
                    value={GALACTIC_PROMPT}
                    className="w-full h-80 bg-transparent text-slate-700 font-mono text-xs leading-relaxed resize-none outline-none focus:ring-0"
                  />
                  <div className="absolute bottom-0 left-0 right-0 h-12 bg-gradient-to-t from-slate-50 to-transparent pointer-events-none" />
                </div>
              </div>
            </motion.div>
          )}
        </AnimatePresence>
      </div>
    );
  };

  return (
    <div className="min-h-screen bg-slate-50 font-sans text-slate-900">
      {/* Header */}
      <header className="bg-white border-b border-slate-200 sticky top-0 z-10">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-20 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl bg-gradient-to-br from-brand-cyan to-brand-purple flex items-center justify-center text-white shadow-lg cursor-pointer" onClick={() => setCurrentScreen("home")}>
              <FileText size={24} />
            </div>
            <div>
              <h1 className="text-xl font-bold tracking-tight text-slate-900 cursor-pointer" onClick={() => setCurrentScreen("home")}>
                Ficha de Obligaciones Ambientales EQ
              </h1>
              <p className="text-xs text-slate-500 font-medium uppercase tracking-wider">
                Equilibria Consulting
              </p>
            </div>
          </div>
          <div className="flex items-center gap-4">
            {currentScreen !== "home" && (
              <button
                onClick={() => {
                  setCurrentScreen("home");
                  setSelectedAI(null);
                }}
                className="flex items-center gap-2 px-4 py-2 rounded-xl border border-slate-200 bg-white text-slate-600 hover:text-slate-900 hover:bg-slate-50 font-bold transition-all text-sm shadow-sm"
              >
                <ArrowLeft size={16} />
                <span className="hidden sm:inline">Menú Principal</span>
                <span className="sm:hidden">Menú</span>
              </button>
            )}
            <div className="hidden md:flex items-center gap-1 -space-x-2">
              <div className="w-8 h-8 rounded-full bg-brand-cyan border-2 border-white" />
              <div className="w-8 h-8 rounded-full bg-brand-light-blue border-2 border-white" />
              <div className="w-8 h-8 rounded-full bg-brand-blue border-2 border-white" />
              <div className="w-8 h-8 rounded-full bg-brand-purple border-2 border-white" />
            </div>
          </div>
        </div>
      </header>

      <main className="max-w-5xl mx-auto px-4 py-12">
        <div className="space-y-8">
          {currentScreen === "home" && renderHomeScreen()}

          {currentScreen === "prompt" && renderPromptScreen()}

          {currentScreen === "engineer" && (
            <>
              {/* Hero Section */}
              <section className="text-center space-y-4">
                <motion.h2 
                  initial={{ opacity: 0, y: 20 }}
                  animate={{ opacity: 1, y: 0 }}
                  className="text-3xl sm:text-4xl font-extrabold text-slate-900"
                >
                  Sistematización de Normas Ambientales
                </motion.h2>
                <motion.p 
                  initial={{ opacity: 0, y: 20 }}
                  animate={{ opacity: 1, y: 0 }}
                  transition={{ delay: 0.1 }}
                  className="text-lg text-slate-600 max-w-2xl mx-auto"
                >
                  Sube tus normas para extraer tus obligaciones ambientales fiscalizables.
                </motion.p>
              </section>


          {/* Input Area */}
          <section className="bg-white rounded-3xl shadow-xl shadow-slate-200/60 border border-slate-100 overflow-hidden">
            <div className="p-6 sm:p-8 space-y-6">
              {/* Manual Info Fields */}
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-4">
                <div className="space-y-2">
                  <label className="text-xs font-bold text-slate-500 uppercase tracking-wider">Nombre de Empresa</label>
                  <input 
                    type="text" 
                    value={generalInfo.empresa}
                    onChange={(e) => setGeneralInfo({...generalInfo, empresa: e.target.value})}
                    placeholder="Ej: Equilibria S.A.C."
                    className="w-full p-3 rounded-xl bg-slate-50 border border-slate-200 focus:border-brand-blue outline-none transition-all text-sm"
                  />
                </div>
                <div className="space-y-2">
                  <label className="text-xs font-bold text-slate-500 uppercase tracking-wider">RUC</label>
                  <input 
                    type="text" 
                    value={generalInfo.ruc}
                    onChange={(e) => setGeneralInfo({...generalInfo, ruc: e.target.value})}
                    placeholder="Ej: 20123456789"
                    className="w-full p-3 rounded-xl bg-slate-50 border border-slate-200 focus:border-brand-blue outline-none transition-all text-sm"
                  />
                </div>
                <div className="space-y-2">
                  <label className="text-xs font-bold text-slate-500 uppercase tracking-wider">Dirección</label>
                  <input 
                    type="text" 
                    value={generalInfo.direccion}
                    onChange={(e) => setGeneralInfo({...generalInfo, direccion: e.target.value})}
                    placeholder="Ej: Av. Principal 123"
                    className="w-full p-3 rounded-xl bg-slate-50 border border-slate-200 focus:border-brand-blue outline-none transition-all text-sm"
                  />
                </div>
                <div className="space-y-2">
                  <label className="text-xs font-bold text-slate-500 uppercase tracking-wider">Sector</label>
                  <select 
                    value={generalInfo.sector}
                    onChange={(e) => setGeneralInfo({...generalInfo, sector: e.target.value})}
                    className="w-full p-3 rounded-xl bg-slate-50 border border-slate-200 focus:border-brand-blue outline-none transition-all text-sm"
                  >
                    {SECTORS.map(s => <option key={s} value={s}>{s}</option>)}
                  </select>
                </div>
                <div className="space-y-2">
                  <label className="text-xs font-bold text-slate-500 uppercase tracking-wider">Área</label>
                  <input 
                    type="text" 
                    value={generalInfo.area}
                    onChange={(e) => setGeneralInfo({...generalInfo, area: e.target.value})}
                    placeholder="Ej: Medio Ambiente"
                    className="w-full p-3 rounded-xl bg-slate-50 border border-slate-200 focus:border-brand-blue outline-none transition-all text-sm"
                  />
                </div>
              </div>

              <div className="flex flex-col sm:flex-row gap-4 items-start sm:items-center justify-between pt-4 border-t border-slate-100">
                <div className="flex items-center gap-2">
                  <button
                    onClick={() => {
                      setInputText("");
                      setResult(null);
                      setError(null);
                    }}
                    className="p-2 text-slate-400 hover:text-red-500 transition-colors"
                    title="Limpiar todo"
                  >
                    <Trash2 size={20} />
                  </button>
                  <div className="flex items-center gap-2 text-brand-blue font-semibold">
                    <Info size={20} />
                    <span>Instrucciones</span>
                  </div>
                </div>
                <div className="flex items-center gap-2">
                  <button
                    onClick={() => fileInputRef.current?.click()}
                    className="flex items-center gap-2 px-4 py-2 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 transition-colors text-sm font-medium"
                  >
                    <Upload size={18} />
                    Subir Archivo
                  </button>
                  <input
                    type="file"
                    ref={fileInputRef}
                    onChange={handleFileUpload}
                    className="hidden"
                    accept=".txt,.md,.pdf"
                  />
                  <button
                    onClick={() => setInputText("")}
                    className="p-2 rounded-xl text-slate-400 hover:text-red-500 hover:bg-red-50 transition-all"
                    title="Limpiar"
                  >
                    <Trash2 size={20} />
                  </button>
                </div>
              </div>

              <div className="relative">
                <textarea
                  value={inputText}
                  onChange={(e) => setInputText(e.target.value)}
                  placeholder="Pega aquí el texto de las leyes, reglamentos o decretos..."
                  className="w-full h-64 p-6 rounded-2xl bg-slate-50 border-2 border-transparent focus:border-brand-blue focus:bg-white transition-all outline-none resize-none text-slate-700 leading-relaxed"
                />
                {!inputText && (
                  <div className="absolute inset-0 flex flex-col items-center justify-center pointer-events-none text-slate-400 space-y-2">
                    <FileText size={48} strokeWidth={1} />
                    <p className="text-sm">El texto aparecerá aquí</p>
                  </div>
                )}
              </div>

              <div className="flex flex-col sm:flex-row gap-4">
                <button
                  onClick={() => processNorms(false)}
                  disabled={isProcessing || !inputText.trim()}
                  className={cn(
                    "flex-1 py-4 rounded-2xl font-bold text-lg shadow-lg transition-all flex items-center justify-center gap-3",
                    isProcessing 
                      ? "bg-slate-200 text-slate-500 cursor-not-allowed" 
                      : "bg-gradient-to-r from-brand-blue to-brand-purple text-white hover:scale-[1.02] active:scale-[0.98] shadow-brand-blue/20"
                  )}
                >
                  {isProcessing ? (
                    <>
                      <Loader2 className="animate-spin" />
                      Procesando...
                    </>
                  ) : (
                    <>
                      <CheckCircle2 />
                      Generar Ficha
                    </>
                  )}
                </button>

                {result && (
                  <button
                    onClick={() => processNorms(true)}
                    disabled={isProcessing || !inputText.trim()}
                    className={cn(
                      "flex-1 py-4 rounded-2xl font-bold text-lg shadow-lg transition-all flex items-center justify-center gap-3",
                      isProcessing 
                        ? "bg-slate-200 text-slate-500 cursor-not-allowed" 
                        : "bg-white border-2 border-brand-purple text-brand-purple hover:bg-brand-purple/5 active:scale-[0.98]"
                    )}
                  >
                    {isProcessing ? (
                      <Loader2 className="animate-spin" />
                    ) : (
                      <>
                        <Upload size={20} />
                        Generar más obligaciones
                      </>
                    )}
                  </button>
                )}
              </div>

              {error && (
                <motion.div 
                  initial={{ opacity: 0, height: 0 }}
                  animate={{ opacity: 1, height: "auto" }}
                  className="p-4 rounded-xl bg-red-50 border border-red-100 flex items-center gap-3 text-red-600 text-sm"
                >
                  <AlertCircle size={20} />
                  {error}
                </motion.div>
              )}
            </div>
          </section>

          {/* Results Area */}
          <AnimatePresence>
            {result && (
              <motion.section
                initial={{ opacity: 0, y: 40 }}
                animate={{ opacity: 1, y: 0 }}
                className="space-y-6"
              >
                <div className="bg-white rounded-3xl shadow-xl shadow-slate-200/60 border border-slate-100 p-8 space-y-8">
                  {/* Word-like Header */}
                  <div className="text-center space-y-4 pb-6 border-b border-slate-100">
                    <p className="text-[10px] font-bold italic text-slate-400 uppercase tracking-widest">
                      {DISCLAIMER}
                    </p>
                    <h2 className="text-2xl font-black text-slate-900 border-y-2 border-slate-900 py-4 uppercase tracking-tighter">
                      Ficha de Obligaciones Ambientales Fiscalizables
                    </h2>
                  </div>

                  <div className="flex items-center justify-between">
                    <div className="space-y-1">
                      <h3 className="text-lg font-bold text-slate-900">Vista Previa del Documento</h3>
                      <p className="text-slate-500 text-xs">Se han identificado {result.obligaciones.reduce((acc: number, curr: any) => acc + curr.items.length, 0)} obligaciones.</p>
                    </div>
                    <button
                      onClick={generateWord}
                      className="flex items-center gap-2 px-6 py-3 rounded-xl bg-brand-cyan text-white font-bold hover:bg-brand-cyan/90 transition-all shadow-lg shadow-brand-cyan/20"
                    >
                      <Download size={20} />
                      Descargar Word
                    </button>
                  </div>

                  <div className="space-y-10">
                    {/* Info General Preview */}
                    <div className="space-y-4">
                      <div className="bg-slate-100 p-3 rounded-lg font-bold text-slate-700 border-l-4 border-brand-blue">
                        I. INFORMACIÓN GENERAL
                      </div>
                      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4 text-sm">
                        <div className="p-4 rounded-xl bg-slate-50 border border-slate-100">
                          <span className="text-slate-400 block mb-1">Empresa</span>
                          <span className="font-semibold">{result.informacionGeneral.empresa}</span>
                        </div>
                        <div className="p-4 rounded-xl bg-slate-50 border border-slate-100">
                          <span className="text-slate-400 block mb-1">R.U.C.</span>
                          <span className="font-semibold">{result.informacionGeneral.ruc}</span>
                        </div>
                        <div className="p-4 rounded-xl bg-slate-50 border border-slate-100">
                          <span className="text-slate-400 block mb-1">Sector</span>
                          <span className="font-semibold">{result.informacionGeneral.sector}</span>
                        </div>
                        <div className="p-4 rounded-xl bg-slate-50 border border-slate-100">
                          <span className="text-slate-400 block mb-1">Área</span>
                          <span className="font-semibold">{result.informacionGeneral.area}</span>
                        </div>
                        <div className="p-4 rounded-xl bg-slate-50 border border-slate-100 sm:col-span-2">
                          <span className="text-slate-400 block mb-1">Dirección</span>
                          <span className="font-semibold">{result.informacionGeneral.direccion}</span>
                        </div>
                      </div>
                    </div>

                    {/* Fuentes Preview */}
                    <div className="space-y-4">
                      <div className="bg-slate-100 p-3 rounded-lg font-bold text-slate-700 border-l-4 border-brand-light-blue">
                        II. FUENTE DE OBLIGACIONES FISCALIZABLES
                      </div>
                      <table className="w-full text-sm text-left border-collapse">
                        <thead>
                          <tr className="bg-slate-50 text-slate-500 uppercase text-xs tracking-wider">
                            <th className="p-3 border-b">Nro</th>
                            <th className="p-3 border-b">Tipo</th>
                            <th className="p-3 border-b">Fuente</th>
                            <th className="p-3 border-b">Autoridad</th>
                            <th className="p-3 border-b">Documento</th>
                            <th className="p-3 border-b">Fecha</th>
                          </tr>
                        </thead>
                        <tbody>
                          {result.fuentes.map((f: any, i: number) => (
                            <tr key={i} className="hover:bg-slate-50 transition-colors">
                              <td className="p-3 border-b font-medium">{f.nro}</td>
                              <td className="p-3 border-b">{f.tipo}</td>
                              <td className="p-3 border-b">{f.fuente}</td>
                              <td className="p-3 border-b">{f.autoridad}</td>
                              <td className="p-3 border-b">{f.documento}</td>
                              <td className="p-3 border-b">{f.fecha}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>

                    {/* Obligaciones Preview */}
                    <div className="space-y-4">
                      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 bg-slate-100 p-3 rounded-lg border-l-4 border-brand-purple">
                        <span className="font-bold text-slate-700">III. OBLIGACIONES FISCALIZABLES</span>
                        <button
                          onClick={() => processNorms(true)}
                          disabled={isProcessing}
                          className={cn(
                            "flex items-center gap-2 px-4 py-2 rounded-xl text-sm font-bold transition-all shadow-md",
                            isProcessing
                              ? "bg-slate-200 text-slate-400 cursor-not-allowed"
                              : "bg-white text-brand-purple border border-brand-purple hover:bg-brand-purple hover:text-white"
                          )}
                        >
                          {isProcessing ? (
                            <Loader2 className="animate-spin size-4" />
                          ) : (
                            <>
                              <PlusCircle size={16} />
                              Agregar más obligaciones
                            </>
                          )}
                        </button>
                      </div>
                      <div className="space-y-6">
                        {result.obligaciones.map((group: any, i: number) => (
                          <div key={i} className="space-y-3">
                            <div className="bg-slate-50 p-2 rounded font-bold text-brand-blue">
                              3.{i + 1} {group.componenteGeneral} - {group.componente}
                            </div>
                            <table className="w-full text-sm text-left border-collapse">
                              <thead>
                                <tr className="text-slate-400 text-xs uppercase">
                                  <th className="p-2 border-b w-16">Ref</th>
                                  <th className="p-2 border-b w-16">Nro</th>
                                  <th className="p-2 border-b w-48">Ubicación</th>
                                  <th className="p-2 border-b">Descripción</th>
                                </tr>
                              </thead>
                              <tbody>
                                {group.items.map((item: any, j: number) => (
                                  <tr key={j} className="hover:bg-slate-50 transition-colors">
                                    <td className="p-2 border-b">{item.referencia}</td>
                                    <td className="p-2 border-b">{item.nro}</td>
                                    <td className="p-2 border-b italic">{item.ubicacion}</td>
                                    <td className="p-2 border-b">{item.descripcion}</td>
                                  </tr>
                                ))}
                              </tbody>
                            </table>
                          </div>
                        ))}
                      </div>
                    </div>
                  </div>
                </div>

                <div className="bg-amber-50 border border-amber-100 p-6 rounded-2xl flex gap-4">
                  <AlertCircle className="text-amber-500 shrink-0" />
                  <p className="text-amber-800 text-sm leading-relaxed">
                    <strong>ESTE RESULTADO DEBE SER CORROBORADO POR EL EQUIPO DE LA EMPRESA, ES SOLO INFORMATIVO Y DEBE REVISARSE PARA EVITAR ERROR, ES RESPONSABILIDAD ABSOLUTA LA REVISIÓN FINAL, LA IA NO SE RESPONSABILIZA POR ERRORES</strong>
                  </p>
                </div>
              </motion.section>
            )}
          </AnimatePresence>
        </>
      )}
        </div>
      </main>

      {/* Footer */}
      <footer className="bg-white border-t border-slate-200 py-12 mt-12">
        <div className="max-w-7xl mx-auto px-4 text-center space-y-4">
          <div className="flex items-center justify-center gap-2 text-slate-400">
            <div className="w-8 h-px bg-slate-200" />
            <span className="text-xs font-bold uppercase tracking-widest">Equilibria Consulting</span>
            <div className="w-8 h-px bg-slate-200" />
          </div>
          <p className="text-slate-500 text-sm">
            Desarrollado By Equilibria Consulting
          </p>
        </div>
      </footer>
    </div>
  );
}
