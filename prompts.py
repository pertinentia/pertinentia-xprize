# -*- coding: utf-8 -*-
"""Suite de Diseño Instruccional (EC0301 / EC0217.01)."""

BASE_IDENTITY = """Eres un Arquitecto de Diseño Instruccional experto en la normativa mexicana 
(CONOCER EC0301 y EC0217.01). Operas bajo DOBLE PERFIL: (1) experto pedagógico y (2) 
experto en la materia del curso. Generas documentos profesionales y congruentes."""

INSTRUCCION_BLOOM = """REGLA DE TAXONOMÍA (FUTURO + 4 DOMINIOS):
Redacta OBJETIVOS con VERBOS DE ACCIÓN EN TIEMPO FUTURO según Bloom/Marzano (Nivel 4-6 recomendado).
Estructura: [Sujeto] + [Cuándo] + [Verbo FUTURO] + [Objeto] + [Condición] + [Finalidad].

Dominios obligatorios:
a) COGNITIVO (Saber): evaluará, fundará, contrastará...
b) PSICOMOTOR (Saber hacer): operará, ensamblará, adaptará...
c) AFECTIVO (Saber ser): se comprometerá, internalizará...
d) RELACIONAL-SOCIAL (Saber convivir): coordinará, mediará, retroalimentará..."""

CARTA_COLUMNS = "| Temas/Subtemas | Actividades | Duración | Técnicas Grupales/Instruccionales | Material y Equipo de Apoyo |"

def build_carta_prompt(tema, duracion_min, objetivo, participantes, reparto):
    comp, ap, des, cie = reparto
    return f"""{BASE_IDENTITY}
{INSTRUCCION_BLOOM}
Genera la CARTA DESCRIPTIVA (EC0301). 
Curso: {tema}. Objetivo: {objetivo}. Duración: {duracion_min} min. Participantes: {participantes}.
Incluye Ficha, Objetivo General y Particulares (4 dominios), y Tabla con columnas: 
{CARTA_COLUMNS}
Reparto: Encuadre {comp}m, Apertura {ap}m, Desarrollo {des}m, Cierre {cie}m. TOTAL: {duracion_min}m."""

def build_iec_prompt(tema, objetivo, subtipo):
    return f"""{BASE_IDENTITY}
{INSTRUCCION_BLOOM}
Genera el INSTRUMENTO DE EVALUACIÓN: {subtipo} para "{tema}".
Reactivos específicos y congruentes. Usa tabla Markdown donde aplique."""

def build_manual_instructor_prompt(tema, objetivo):
    return f"{BASE_IDENTITY}\n{INSTRUCCION_BLOOM}\nGenera MANUAL DEL INSTRUCTOR detallado para '{tema}'."

def build_manual_participante_prompt(tema, objetivo):
    return f"{BASE_IDENTITY}\n{INSTRUCCION_BLOOM}\nGenera MANUAL DEL PARTICIPANTE (contenido teórico) para '{tema}'."

DOC_TYPES = {
    "carta": "Carta Descriptiva",
    "iec": "Instrumento de Evaluación",
    "manual_instructor": "Manual del Instructor",
    "manual_participante": "Manual del Participante",
}