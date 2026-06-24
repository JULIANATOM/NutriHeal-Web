# -*- coding: utf-8 -*-
"""
Actualiza el archivo de auditoria NutriHeal 360:
1. Reemplaza los prompts de imagen (columna 9) con la plantilla validada
   en nano banana / FLUX.2 Pro que preserva el producto real.
2. Agrega dos columnas nuevas con hipervínculos a las imágenes de referencia
   que el operador debe subir al generador de imágenes:
      - Col 13: Imagen Producto (Ref. 1) → URL Azure Blob
      - Col 14: Logo NutriHeal    (Ref. 2) → archivo local
3. Regenera las hojas Publer Import con los alt-texts actualizados.
"""
import json
import os
import re
import sys
import urllib.parse
import csv

sys.stdout.reconfigure(encoding='utf-8')

import openpyxl
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SRC_XLSX = os.path.join(BASE_DIR, 'NutriHeal360_DrMestizo_Auditoria_Contenido_3Meses_2026.xlsx')
CATALOGO_JSON = os.path.join(BASE_DIR, 'catalogo.json')

BLOB_BASE = 'https://nutriheal.blob.core.windows.net/productos/'
LOGO_PATH_LOCAL = os.path.join(BASE_DIR, 'Imagenes', 'logo', 'logo.png')
LOGO_URL = 'file:///' + LOGO_PATH_LOCAL.replace('\\', '/').replace('G:/', 'G:/')
WHATSAPP_LINK = 'https://wa.me/573147088080'

# ── Carga catálogo ──────────────────────────────────────────────────────────
with open(CATALOGO_JSON, encoding='utf-8') as f:
    catalog_data = json.load(f)

CATALOG_MAP = {
    p['nombre'].upper().strip(): p
    for p in catalog_data['productos']
}

def blob_url_producto(nombre):
    """Devuelve la URL pública de la imagen principal en Azure Blob Storage."""
    key = nombre.upper().strip()
    if key not in CATALOG_MAP:
        return None
    p = CATALOG_MAP[key]
    cat = urllib.parse.quote(p['categoria'])
    sub = urllib.parse.quote(p['subcategoria'])
    codigo = p['codigo']
    return f'{BLOB_BASE}{cat}/{sub}/{codigo}.png'

# ── Plantillas de prompts ────────────────────────────────────────────────────
SUPLEMENTACION_TEMPLATE = (
    'Keep the supplement bottle labeled {product} exactly as shown in the reference image — '
    'do not modify, redesign, recreate, or alter its label, text, shape, color, or branding '
    'in any way. Generate only a new background and environment around the unchanged product: '
    'dark marble surface, soft emerald green (#1B7A4A) ambient glow and lighting, '
    'gold (#C9A84C) accent details in the scene, scattered fresh herbs and botanical elements '
    'like ginger and citrus slices around the bottle, professional studio photography lighting, '
    'shallow depth of field, luxury wellness brand aesthetic. '
    'Place the NutriHeal 360 logo from the second reference image in the bottom-right corner '
    'of the composition exactly as it appears in that reference — do not recolor, redesign, '
    'modify, or alter any element of the logo in any way (not the colors, not the typography, '
    'not the icon, not the proportions). Use it as a fixed, pixel-accurate brand watermark, '
    'clearly outside the product and not overlapping or touching the bottle itself.'
)

LOGO_TAG = (
    ' Place the NutriHeal 360 logo from the reference image in the bottom-right corner '
    'exactly as it appears — do not recolor, redesign, modify, or alter any element of the '
    'logo in any way (not the colors, not the typography, not the icon, not the proportions). '
    'Use it as a fixed, pixel-accurate brand watermark.'
)

OLD_TO_NEW_PROMPT = {
    # Medicina Alternativa
    'Photorealistic professional medical marketing image, emerald green (#1B7A4A) and premium gold (#C9A84C) color palette, luxury health brand aesthetic, ultra-detailed, 8K resolution, soft cinematic lighting, homeopathic glass vials and pellets on botanical background, Colombian doctor consultation room background blurred, emerald and gold tones, trust and science emotion':
        'Photorealistic professional wellness photograph, emerald green (#1B7A4A) and premium gold (#C9A84C) color palette, luxury health brand aesthetic, ultra-detailed, 8K resolution, soft cinematic lighting, homeopathic glass vials and pellets on botanical background, blurred consultation room background, emerald and gold tones, trust and science emotion.' + LOGO_TAG,
    'Photorealistic professional medical marketing image, emerald green (#1B7A4A) and premium gold (#C9A84C) color palette, luxury health brand aesthetic, ultra-detailed, 8K resolution, soft cinematic lighting, iris of human eye close-up revealing health patterns, iridology diagnostic chart overlay, Colombian doctor consultation room background blurred, emerald and gold tones, trust and science emotion':
        'Photorealistic professional wellness photograph, emerald green (#1B7A4A) and premium gold (#C9A84C) color palette, luxury health brand aesthetic, ultra-detailed, 8K resolution, soft cinematic lighting, iris of a human eye close-up revealing health patterns with iridology diagnostic chart overlay, blurred consultation room background, emerald and gold tones.' + LOGO_TAG,
    'Photorealistic professional medical marketing image, emerald green (#1B7A4A) and premium gold (#C9A84C) color palette, luxury health brand aesthetic, ultra-detailed, 8K resolution, soft cinematic lighting, vintage botanical medical illustration with modern gold frame overlay, Colombian doctor consultation room background blurred, emerald and gold tones, trust and science emotion':
        'Photorealistic professional wellness photograph, emerald green (#1B7A4A) and premium gold (#C9A84C) color palette, luxury health brand aesthetic, ultra-detailed, 8K resolution, soft cinematic lighting, vintage botanical medical illustration with elegant modern gold frame overlay, blurred background, emerald and gold tones.' + LOGO_TAG,
    'Photorealistic professional medical marketing image, emerald green (#1B7A4A) and premium gold (#C9A84C) color palette, luxury health brand aesthetic, ultra-detailed, 8K resolution, soft cinematic lighting, human silhouette with glowing energy meridian lines, holistic healing visualization, Colombian doctor consultation room background blurred, emerald and gold tones, trust and science emotion':
        'Photorealistic professional wellness photograph, emerald green (#1B7A4A) and premium gold (#C9A84C) color palette, luxury health brand aesthetic, ultra-detailed, 8K resolution, soft cinematic lighting, human silhouette with glowing energy meridian lines, holistic healing visualization, blurred background, emerald and gold tones.' + LOGO_TAG,
    # Prevención Médica Diaria
    'Photorealistic professional medical marketing image, emerald green (#1B7A4A) and premium gold (#C9A84C) color palette, luxury health brand aesthetic, ultra-detailed, 8K resolution, soft cinematic lighting, fresh organic vegetables and medicinal plants arranged on dark slate, informative health poster style, Colombian demographic, emerald green and gold accents, hope and vitality emotion':
        'Photorealistic professional wellness photograph, emerald green (#1B7A4A) and premium gold (#C9A84C) color palette, luxury health brand aesthetic, ultra-detailed, 8K resolution, soft cinematic lighting, fresh organic vegetables and medicinal plants beautifully arranged on dark slate surface, Colombian demographic, emerald green and gold accents, hope and vitality emotion.' + LOGO_TAG,
    'Photorealistic professional medical marketing image, emerald green (#1B7A4A) and premium gold (#C9A84C) color palette, luxury health brand aesthetic, ultra-detailed, 8K resolution, soft cinematic lighting, human body anatomical illustration with glowing healthy organs, emerald green highlights, informative health poster style, Colombian demographic, emerald green and gold accents, hope and vitality emotion':
        'Photorealistic professional wellness photograph, emerald green (#1B7A4A) and premium gold (#C9A84C) color palette, luxury health brand aesthetic, ultra-detailed, 8K resolution, soft cinematic lighting, human body anatomical illustration with glowing healthy organs highlighted in emerald green, Colombian demographic, gold accents.' + LOGO_TAG,
    'Photorealistic professional medical marketing image, emerald green (#1B7A4A) and premium gold (#C9A84C) color palette, luxury health brand aesthetic, ultra-detailed, 8K resolution, soft cinematic lighting, healthy Colombian family at golden hour outdoor setting, vibrant wellness energy, informative health poster style, Colombian demographic, emerald green and gold accents, hope and vitality emotion':
        'Photorealistic professional wellness photograph, emerald green (#1B7A4A) and premium gold (#C9A84C) color palette, luxury health brand aesthetic, ultra-detailed, 8K resolution, soft cinematic lighting, healthy Colombian family at golden hour outdoor setting, vibrant wellness energy, natural environment, hope and vitality emotion.' + LOGO_TAG,
    'Photorealistic professional medical marketing image, emerald green (#1B7A4A) and premium gold (#C9A84C) color palette, luxury health brand aesthetic, ultra-detailed, 8K resolution, soft cinematic lighting, DNA helix structure with protective shield overlay, prevention concept, informative health poster style, Colombian demographic, emerald green and gold accents, hope and vitality emotion':
        'Photorealistic professional wellness photograph, emerald green (#1B7A4A) and premium gold (#C9A84C) color palette, luxury health brand aesthetic, ultra-detailed, 8K resolution, soft cinematic lighting, DNA helix structure with glowing protective shield overlay symbolizing prevention, emerald green and gold accents.' + LOGO_TAG,
    # Casos de Éxito
    'Photorealistic professional medical marketing image, emerald green (#1B7A4A) and premium gold (#C9A84C) color palette, luxury health brand aesthetic, ultra-detailed, 8K resolution, soft cinematic lighting, Colombian patient and doctor moment of success in medical consultation room, warm smile, trust and gratitude emotion, San Fernando Cali medical office aesthetic, certificate wall background, emerald green plants, gold framed diplomas, before-after transformation concept':
        'Photorealistic professional wellness photograph, emerald green (#1B7A4A) and premium gold (#C9A84C) color palette, luxury health brand aesthetic, ultra-detailed, 8K resolution, soft cinematic lighting, Colombian patient and doctor sharing a warm moment of success inside a consultation room, natural smile, trust and gratitude, San Fernando Cali office aesthetic with certificate wall, emerald green plants and gold framed diplomas in background.' + LOGO_TAG,
}

# Regex para detectar prompts de Suplementación (formato viejo)
PRODUCTO_RE_OLD = re.compile(r'of "([^"]+)" on a dark marble')
# Regex para detectar prompts de Suplementación (formato nuevo)
PRODUCTO_RE_NEW = re.compile(r'Keep the supplement bottle labeled (.+?) exactly as shown')

# Textos de logo ANTERIORES que deben reemplazarse (formato intermedio)
_OLD_LOGO_SIMPLE = (
    ' Place the NutriHeal 360 logo from the reference image as a small brand watermark '
    'in the bottom-right corner of the composition.'
)
_OLD_LOGO_SUPL = (
    'Place the NutriHeal 360 logo from the second reference image as a small separate brand '
    'watermark in the bottom-right corner of the composition, clearly outside the product and '
    'not overlapping or touching the bottle itself.'
)

def nuevo_prompt(prompt_viejo):
    """Devuelve el prompt actualizado y el nombre del producto si aplica."""
    if not prompt_viejo:
        return prompt_viejo, None

    # Suplementación — formato ORIGINAL
    m = PRODUCTO_RE_OLD.search(prompt_viejo)
    if m:
        nombre = m.group(1)
        return SUPLEMENTACION_TEMPLATE.format(product=nombre), nombre

    # No-Suplementación — formato ORIGINAL (viejo prompt de medical marketing)
    nuevo = OLD_TO_NEW_PROMPT.get(prompt_viejo)
    if nuevo:
        return nuevo, None

    # Suplementación — formato NUEVO pero con logo viejo (re-ejecución)
    m2 = PRODUCTO_RE_NEW.search(prompt_viejo)
    if m2:
        nombre = m2.group(1)
        return SUPLEMENTACION_TEMPLATE.format(product=nombre), nombre

    # No-Suplementación — formato intermedio (wellness photograph + logo viejo)
    # Solo reemplazar la instrucción del logo por la nueva versión estricta
    if _OLD_LOGO_SIMPLE in prompt_viejo:
        return prompt_viejo.replace(_OLD_LOGO_SIMPLE, LOGO_TAG), None

    return prompt_viejo, None

# ── Generación dinámica de instrucción HeyGen ────────────────────────────────

_SKIP_HEYGEN = ('📲', '📍', '🔴', '💬', '👇', '👍', '🔔', '⏱', '#', 'wa.me',
                'WhatsApp', 'www.', 'doctormestizo', 'Cra ', '+57')

def _limpiar_linea(line):
    return re.sub(r'^[^\w¿"]+', '', line).strip()

def extraer_titulo_heygen(copy, pilar):
    """Extrae el tema/título del post para la instrucción HeyGen."""
    if not copy:
        return 'Salud y bienestar con NutriHeal 360'
    lines = [l.strip() for l in copy.split('\n') if l.strip()]

    # YouTube/Facebook: título entre comillas en las primeras 3 líneas
    for line in lines[:3]:
        m = re.search(r'"([^"]{15,})"', line)
        if m:
            return m.group(1).strip()

    # Prevención: "PREVENCIÓN HOY: Tema"
    for line in lines[:2]:
        m = re.search(r'PREVENCIÓN HOY\s*:\s*(.+)', line, re.IGNORECASE)
        if m:
            return m.group(1).strip()

    # Casos de Éxito: línea tras "HISTORIA DE TRANSFORMACIÓN"
    for i, line in enumerate(lines[:3]):
        if 'HISTORIA DE' in line.upper() or 'TRAYECTORIA' in line.upper():
            for j in range(i+1, min(i+3, len(lines))):
                cand = _limpiar_linea(lines[j])
                if len(cand) > 10 and not any(s in lines[j] for s in _SKIP_HEYGEN):
                    return cand
            break

    # Suplementación: "Hoy hablo de PRODUCTO — nuestra línea SUBCATEGORIA"
    if 'Suplementación' in pilar:
        for line in lines:
            m = re.search(r'(?:hablo de|habla de)\s+([A-ZÁÉÍÓÚÑÜ][A-Z0-9\- ]{2,}[A-Z0-9])', line)
            if m:
                prod = m.group(1).strip()
                sm = re.search(r'línea\s+(\w[\w\s]+?)\.', line)
                subcat = f' — línea {sm.group(1).strip()}' if sm else ''
                return f'Beneficios de {prod}{subcat}'

    # Fallback: primera línea limpia
    return _limpiar_linea(lines[0])[:120]


def extraer_puntos_guion(copy, n=3):
    """Extrae hasta n oraciones sustantivas del copy para el guión."""
    if not copy:
        return []
    lines = [l.strip() for l in copy.split('\n') if l.strip()]
    puntos = []
    for line in lines[1:]:
        if any(s in line for s in _SKIP_HEYGEN):
            continue
        # Saltar líneas que son solo un título entre comillas (ya va en TEMA)
        if re.match(r'^"[^"]{10,}"\.?$', line.strip()):
            continue
        clean = re.sub(r'^[🌿🔬🛡️🎙️✅✨💚🧬🌱💊🧪•◦▸→\-\*✓]+\s*', '', line).strip()
        clean = re.sub(r'[^\w\sáéíóúÁÉÍÓÚñÑüÜ:.,¿?!""()\-]', '', clean).strip()
        if len(clean) > 30:
            puntos.append(clean)
        if len(puntos) >= n:
            break
    return puntos


_BROLL = {
    'Suplementación NutriHeal':    '{producto} en close-up, animaciones moleculares 3D del mecanismo de acción, testimonios en texto flotante, sello INVIMA animado',
    'Medicina Alternativa':        'ilustraciones de medicina homeopática, imágenes de iridodiagnosis, plantas medicinales en close-up, consultorio San Fernando Cali',
    'Prevención Médica Diaria':    'infografías de salud preventiva, estadísticas en texto animado, alimentos y plantas medicinales, familia colombiana saludable',
    'Casos de Éxito / Trayectoria':'tarjetas de testimonios de pacientes, antes/después en texto, consultorio con diplomas y certificados enmarcados en dorado',
}

_HEYGEN_PLAT = {
    'Facebook': {
        'encuadre':   'plano americano clásico (3/4 figura)',
        'fondo':      'consultorio NutriHeal San Fernando Cali — escritorio médico, planta monstera, vitrina con productos NutriHeal visibles',
        'iluminacion':'softbox frontal + luz esmeralda lateral',
        'formato':    '1:1 (Feed) o 16:9 (video nativo) — 1080x1080 / 1920x1080',
        'duracion':   '60-90 segundos',
    },
    'Instagram': {
        'encuadre':   'primer plano (hombros-cabeza) centrado',
        'fondo':      'consultorio blanco-verde esmeralda minimalista, planta en esquina, logo NutriHeal visible en fondo',
        'iluminacion':'luz natural lateral + relleno suave',
        'formato':    '9:16 vertical (Reels) — 1080x1920',
        'duracion':   '30-60 segundos. Entrada directa sin intro; CTA animado al final',
    },
    'YouTube': {
        'encuadre':   'plano medio-americano (cintura-cabeza)',
        'fondo':      'consultorio NutriHeal San Fernando Cali — estantería con libros médicos, plantas verdes, diploma en pared dorada, luz cálida lateral',
        'iluminacion':'softbox principal + luz de relleno dorada + backlight esmeralda',
        'formato':    '16:9 horizontal — 1920x1080 (Full HD)',
        'duracion':   '8-12 minutos',
    },
}


def generar_heygen(copy, pilar, red, producto=None):
    """Construye la instrucción completa para HeyGen basada en el copy del post."""
    titulo = extraer_titulo_heygen(copy, pilar)
    puntos = extraer_puntos_guion(copy, n=3)
    meta   = _HEYGEN_PLAT.get(red, _HEYGEN_PLAT['YouTube'])

    broll_key = pilar if pilar in _BROLL else 'Medicina Alternativa'
    broll = _BROLL[broll_key]
    if producto and '{producto}' in broll:
        broll = broll.replace('{producto}', producto)
    elif '{producto}' in broll:
        broll = broll.replace('{producto} en close-up, ', '')

    guion_puntos = ' | '.join(puntos) if puntos else titulo
    texto_pantalla = producto if ('Suplementación' in pilar and producto) else titulo[:70]

    partes = [
        'Avatar: Dr. Mestizo HD.',
        f'TEMA DEL VIDEO: {titulo}.',
        f'GUIÓN: El Dr. Mestizo habla sobre este tema. Script basado en el copy del post — puntos clave a desarrollar: {guion_puntos}. Cierra con CTA al WhatsApp +57 314 708 8080.',
        f'ENCUADRE: {meta["encuadre"]}.',
        f'FONDO: {meta["fondo"]}.',
        f'ILUMINACIÓN: {meta["iluminacion"]}.',
        f'B-ROLL: {broll}.',
        f'TEXTO EN PANTALLA: "{texto_pantalla}" en verde esmeralda (#1B7A4A).',
        f'FORMATO: {meta["formato"]}.',
        f'DURACIÓN: {meta["duracion"]}.',
    ]
    return ' '.join(partes)


def make_hyperlink_cell(ws, row, col, url, label):
    cell = ws.cell(row=row, column=col)
    cell.value = label
    cell.hyperlink = url
    cell.font = Font(color='0563C1', underline='single')
    cell.alignment = Alignment(vertical='top')

def actualizar_hoja_fuente(wb, nombre_hoja):
    ws = wb[nombre_hoja]

    # Asegurar encabezados en columnas 13 y 14
    ws.cell(row=1, column=13).value = '📦 Imagen Producto (Ref. 1 — nano banana)'
    ws.cell(row=1, column=14).value = '🏷️ Logo NutriHeal (Ref. 2 — nano banana)'
    for col in (13, 14):
        ws.cell(row=1, column=col).font = Font(bold=True)
        ws.cell(row=1, column=col).alignment = Alignment(wrap_text=True, vertical='top')
    ws.column_dimensions[get_column_letter(13)].width = 32
    ws.column_dimensions[get_column_letter(14)].width = 30

    # Red social de esta hoja (para HeyGen)
    red_nombre = nombre_hoja.replace('📘 ', '').replace('📷 ', '').replace('▶️ ', '').strip()

    filas_ok = 0
    filas_sin_map = 0
    for r in range(2, ws.max_row + 1):
        prompt_viejo = ws.cell(r, 9).value
        if not prompt_viejo:
            continue

        copy    = ws.cell(r, 8).value or ''
        pilar   = ws.cell(r, 4).value or ''

        # ── Col 9: prompt imagen ──────────────────────────────────────────────
        nuevo_p, nombre_prod = nuevo_prompt(str(prompt_viejo))
        ws.cell(r, 9).value = nuevo_p

        # ── Col 11: instrucción HeyGen dinámica ──────────────────────────────
        heygen_instr = generar_heygen(copy, pilar, red_nombre, producto=nombre_prod)
        ws.cell(r, 11).value = heygen_instr
        ws.cell(r, 11).alignment = Alignment(wrap_text=True, vertical='top')

        # ── Col 14: logo (siempre) ───────────────────────────────────────────
        make_hyperlink_cell(ws, r, 14, LOGO_URL, '🏷️ Logo NutriHeal')

        # ── Col 13: imagen del producto (solo Suplementación) ────────────────
        if nombre_prod:
            blob = blob_url_producto(nombre_prod)
            if blob:
                make_hyperlink_cell(ws, r, 13, blob, f'📦 {nombre_prod}')
            else:
                ws.cell(r, 13).value = f'⚠️ No encontrado en catálogo: {nombre_prod}'
                ws.cell(r, 13).alignment = Alignment(vertical='top')
        else:
            ws.cell(r, 13).value = 'N/A — imagen temática generativa'
            ws.cell(r, 13).alignment = Alignment(vertical='top')
            ws.cell(r, 13).font = Font(color='888888', italic=True)

        if nuevo_p == prompt_viejo and not PRODUCTO_RE_NEW.search(nuevo_p):
            filas_sin_map += 1
        filas_ok += 1

    print(f'  {nombre_hoja}: {filas_ok} filas actualizadas'
          + (f' ({filas_sin_map} sin mapeo)' if filas_sin_map else ''))

# ── Alt-text actualizado ─────────────────────────────────────────────────────
ALT_TEXT_MAP_NEW = {v: k2 for k2, v in {
    'Viales de vidrio y gránulos homeopáticos sobre fondo botánico, en tonos verde esmeralda y dorado — NutriHeal 360':
        'homeopathic glass vials and pellets on botanical background, blurred consultation room',
    'Primer plano del iris de un ojo humano con superposición de carta diagnóstica de iridología':
        'iris of a human eye close-up revealing health patterns with iridology diagnostic chart overlay',
    'Ilustración botánica médica de estilo vintage con marco dorado moderno':
        'vintage botanical medical illustration with elegant modern gold frame overlay',
    'Silueta humana con líneas de meridianos energéticos brillantes — visualización de sanación holística':
        'human silhouette with glowing energy meridian lines, holistic healing visualization',
    'Vegetales orgánicos frescos y plantas medicinales sobre pizarra oscura, estilo póster informativo de salud':
        'fresh organic vegetables and medicinal plants beautifully arranged on dark slate surface',
    'Ilustración anatómica del cuerpo humano con órganos saludables iluminados en verde esmeralda':
        'human body anatomical illustration with glowing healthy organs highlighted in emerald green',
    'Familia colombiana saludable al aire libre durante la hora dorada, energía de bienestar':
        'healthy Colombian family at golden hour outdoor setting, vibrant wellness energy',
    'Estructura de hélice de ADN con escudo protector superpuesto — concepto de prevención':
        'DNA helix structure with glowing protective shield overlay symbolizing prevention',
    'Paciente y médico colombianos en un momento de éxito dentro del consultorio, sonrisa cálida — pared de certificados, plantas verdes y diplomas enmarcados en dorado':
        'Colombian patient and doctor sharing a warm moment of success inside a consultation room',
}.items()}

ALT_TEXT_MAP_OLD = {
    'Photorealistic professional medical marketing image, emerald green (#1B7A4A) and premium gold (#C9A84C) color palette, luxury health brand aesthetic, ultra-detailed, 8K resolution, soft cinematic lighting, homeopathic glass vials and pellets on botanical background, Colombian doctor consultation room background blurred, emerald and gold tones, trust and science emotion':
        'Viales de vidrio y gránulos homeopáticos sobre fondo botánico, en tonos verde esmeralda y dorado — NutriHeal 360',
    'Photorealistic professional medical marketing image, emerald green (#1B7A4A) and premium gold (#C9A84C) color palette, luxury health brand aesthetic, ultra-detailed, 8K resolution, soft cinematic lighting, iris of human eye close-up revealing health patterns, iridology diagnostic chart overlay, Colombian doctor consultation room background blurred, emerald and gold tones, trust and science emotion':
        'Primer plano del iris de un ojo humano con superposición de carta diagnóstica de iridología',
    'Photorealistic professional medical marketing image, emerald green (#1B7A4A) and premium gold (#C9A84C) color palette, luxury health brand aesthetic, ultra-detailed, 8K resolution, soft cinematic lighting, vintage botanical medical illustration with modern gold frame overlay, Colombian doctor consultation room background blurred, emerald and gold tones, trust and science emotion':
        'Ilustración botánica médica de estilo vintage con marco dorado moderno',
    'Photorealistic professional medical marketing image, emerald green (#1B7A4A) and premium gold (#C9A84C) color palette, luxury health brand aesthetic, ultra-detailed, 8K resolution, soft cinematic lighting, human silhouette with glowing energy meridian lines, holistic healing visualization, Colombian doctor consultation room background blurred, emerald and gold tones, trust and science emotion':
        'Silueta humana con líneas de meridianos energéticos brillantes — visualización de sanación holística',
    'Photorealistic professional medical marketing image, emerald green (#1B7A4A) and premium gold (#C9A84C) color palette, luxury health brand aesthetic, ultra-detailed, 8K resolution, soft cinematic lighting, fresh organic vegetables and medicinal plants arranged on dark slate, informative health poster style, Colombian demographic, emerald green and gold accents, hope and vitality emotion':
        'Vegetales orgánicos frescos y plantas medicinales sobre pizarra oscura, estilo póster informativo de salud',
    'Photorealistic professional medical marketing image, emerald green (#1B7A4A) and premium gold (#C9A84C) color palette, luxury health brand aesthetic, ultra-detailed, 8K resolution, soft cinematic lighting, human body anatomical illustration with glowing healthy organs, emerald green highlights, informative health poster style, Colombian demographic, emerald green and gold accents, hope and vitality emotion':
        'Ilustración anatómica del cuerpo humano con órganos saludables iluminados en verde esmeralda',
    'Photorealistic professional medical marketing image, emerald green (#1B7A4A) and premium gold (#C9A84C) color palette, luxury health brand aesthetic, ultra-detailed, 8K resolution, soft cinematic lighting, healthy Colombian family at golden hour outdoor setting, vibrant wellness energy, informative health poster style, Colombian demographic, emerald green and gold accents, hope and vitality emotion':
        'Familia colombiana saludable al aire libre durante la hora dorada, energía de bienestar',
    'Photorealistic professional medical marketing image, emerald green (#1B7A4A) and premium gold (#C9A84C) color palette, luxury health brand aesthetic, ultra-detailed, 8K resolution, soft cinematic lighting, DNA helix structure with protective shield overlay, prevention concept, informative health poster style, Colombian demographic, emerald green and gold accents, hope and vitality emotion':
        'Estructura de hélice de ADN con escudo protector superpuesto — concepto de prevención',
    'Photorealistic professional medical marketing image, emerald green (#1B7A4A) and premium gold (#C9A84C) color palette, luxury health brand aesthetic, ultra-detailed, 8K resolution, soft cinematic lighting, Colombian patient and doctor moment of success in medical consultation room, warm smile, trust and gratitude emotion, San Fernando Cali medical office aesthetic, certificate wall background, emerald green plants, gold framed diplomas, before-after transformation concept':
        'Paciente y médico colombianos en un momento de éxito dentro del consultorio, sonrisa cálida — pared de certificados, plantas verdes y diplomas enmarcados en dorado',
}

def generar_alt_text(prompt):
    if not prompt:
        return ''
    m = PRODUCTO_RE_NEW.search(prompt)
    if m:
        producto = m.group(1)
        return (f'Foto de producto: frasco de {producto} sobre mármol oscuro con '
                f'brillo verde esmeralda, detalles dorados y hierbas medicinales — NutriHeal 360')
    m = PRODUCTO_RE_OLD.search(prompt)
    if m:
        producto = m.group(1)
        return (f'Foto de producto: frasco de {producto} sobre mármol oscuro con '
                f'brillo verde esmeralda, detalles dorados y hierbas medicinales — NutriHeal 360')
    return ALT_TEXT_MAP_OLD.get(prompt, ALT_TEXT_MAP_NEW.get(prompt, ''))

# ── Publer Import ────────────────────────────────────────────────────────────
PUBLER_HEADERS = [
    'Date - Intl. format or prompt', 'Text',
    'Link(s) - Separated by comma for FB carousels',
    'Media URL(s) - Separated by comma',
    'Title - For the video, pin, PDF ..',
    'Label(s) - Separated by comma',
    'Alt text(s) - Separated by ||',
    'Comment(s) - Separated by ||',
    'Pin board, FB album, or Google category',
    'Post subtype - I.e. story, reel, PDF ..',
    'CTA - For Facebook links or Google',
    'Reminder - For stories, reels, shorts, and TikToks',
]
EXTRA_HEADERS = [
    'Relación de Aspecto', 'Resolución Sugerida',
    'Duración Estimada', 'Pilar Temático', 'N° (ref. hoja origen)',
]
PLATFORM_META = {
    'Facebook':  {'sheet': '📘 Facebook',  'aspecto': '1:1 (Feed) / 16:9 (Video nativo)', 'resolucion': '1080x1080 / 1920x1080', 'duracion': '60-90 segundos (si es video nativo)', 'post_subtype': '', 'cta': 'Send Message', 'link': WHATSAPP_LINK},
    'Instagram': {'sheet': '📷 Instagram', 'aspecto': '9:16 (Vertical — Reels)',          'resolucion': '1080x1920',             'duracion': '30-60 segundos',                      'post_subtype': 'Reel',  'cta': '', 'link': ''},
    'YouTube':   {'sheet': '▶️ YouTube',   'aspecto': '16:9 (Horizontal)',                'resolucion': '1920x1080 (mínimo Full HD)', 'duracion': '8-12 minutos',                   'post_subtype': 'Video', 'cta': '', 'link': ''},
}
MEDIA_PENDIENTE = 'PENDIENTE — generar con nano banana (FLUX.2 Pro) / HeyGen y subir antes de programar'

from datetime import datetime
TITULO_RE = re.compile(r'"([^"]+)"')

def parse_fecha_hora(fecha_str, hora_str):
    fecha = datetime.strptime(str(fecha_str).strip(), '%d/%m/%Y')
    hora  = datetime.strptime(str(hora_str).strip(), '%I:%M %p')
    return fecha.replace(hour=hora.hour, minute=hora.minute).strftime('%Y-%m-%d %H:%M')

def construir_fila_publer(row, red):
    meta = PLATFORM_META[red]
    n, fecha, _dia, pilar, _red, _plat, hora, copy, prompt_mj = row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8]
    hashtags = row[11] if len(row) > 11 else ''
    fecha_pub = parse_fecha_hora(fecha, hora)
    titulo = (TITULO_RE.search(copy or '').group(1) if red == 'YouTube' else '')
    alt = generar_alt_text(prompt_mj)
    publer = [
        fecha_pub, copy, meta['link'], MEDIA_PENDIENTE, titulo,
        pilar, alt, hashtags or '', '', meta['post_subtype'], meta['cta'], '',
    ]
    extra = [meta['aspecto'], meta['resolucion'], meta['duracion'], pilar, n]
    return publer + extra

def regenerar_publer_import(wb, red):
    meta = PLATFORM_META[red]
    src = wb[meta['sheet']]
    rows = list(src.iter_rows(min_row=2, max_col=12, values_only=True))

    nombre_hoja = f'📤 Publer Import - {red}'
    if nombre_hoja in wb.sheetnames:
        del wb[nombre_hoja]
    ws = wb.create_sheet(nombre_hoja)

    headers = PUBLER_HEADERS + EXTRA_HEADERS
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(wrap_text=True, vertical='top')
    ws.freeze_panes = 'A2'

    for row in rows:
        ws.append(construir_fila_publer(row, red))

    anchos = [16, 50, 22, 50, 40, 22, 45, 30, 18, 16, 16, 18, 26, 22, 24, 22, 10]
    for i, ancho in enumerate(anchos, start=1):
        ws.column_dimensions[get_column_letter(i)].width = ancho
    for r in range(2, ws.max_row + 1):
        for c in (2, 7, 8):
            ws.cell(row=r, column=c).alignment = Alignment(wrap_text=True, vertical='top')

    return ws, len(rows)

def exportar_csv(ws, red, total_filas):
    out = os.path.join(BASE_DIR, f'publer_import_{red.lower()}.csv')
    with open(out, 'w', encoding='utf-8-sig', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(PUBLER_HEADERS)
        for r in range(2, total_filas + 2):
            fila = [ws.cell(row=r, column=c).value or '' for c in range(1, len(PUBLER_HEADERS) + 1)]
            writer.writerow(fila)
    return out

def main():
    wb = openpyxl.load_workbook(SRC_XLSX)

    print('=== Actualizando hojas fuente ===')
    for hoja in ['📘 Facebook', '📷 Instagram', '▶️ YouTube']:
        actualizar_hoja_fuente(wb, hoja)

    print('\n=== Regenerando hojas Publer Import ===')
    for red in ['Facebook', 'Instagram', 'YouTube']:
        ws, total = regenerar_publer_import(wb, red)
        csv_path = exportar_csv(ws, red, total)
        print(f'  {red}: {total} filas → hoja "📤 Publer Import - {red}" + {os.path.basename(csv_path)}')

    wb.save(SRC_XLSX)
    print(f'\n✅ Guardado: {os.path.basename(SRC_XLSX)}')
    print(f'📁 Carpeta Marketing: {os.path.join(BASE_DIR, "Marketing")}')

if __name__ == '__main__':
    main()
