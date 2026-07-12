# -*- coding: utf-8 -*-
"""Reemplaza el texto sin acentos que agrego apply_pendientes_manuscrito.py por
versiones con tildes/enes correctas. Localiza cada run por un fragmento unico."""
from pathlib import Path
import docx

ROOT = Path(__file__).resolve().parents[2]
DOC = ROOT / "SEVERIDAD SINTOMAS DEPRESIVOS VS PRESION ARTEIAL ELEVADA (1).docx"

# (fragmento-ancla sin acentos) -> texto correcto completo del run
FIXES = {
    "La presion arterial se determino a partir de dos tomas":
        " La presión arterial se determinó a partir de dos tomas realizadas en una única "
        "visita domiciliaria (variables QS903 y QS905); se definió presión arterial elevada "
        "como un promedio de presión arterial sistólica ≥ 140 mmHg o de presión arterial "
        "diastólica ≥ 90 mmHg. Antes de promediar las tomas se depuraron las mediciones: se "
        "descartaron los valores centinela (999) y los fisiológicamente implausibles (sistólica "
        "< 70 o > 270 mmHg; diastólica < 30 o > 150 mmHg), y se excluyeron por inconsistencia "
        "hemodinámica los registros con inversión tensional —aquellos en que la presión sistólica "
        "no superaba a la diastólica en alguna de las tomas—, que por tanto no aportaron un "
        "desenlace válido.",
    "La justificacion causal y el respaldo bibliografico":
        " La justificación causal y el respaldo bibliográfico de cada confusor estructural "
        "—y de las aristas confusor→exposición y confusor→desenlace que motivan su ajuste— "
        "se documentan arista por arista en el grafo acíclico dirigido (Figura S2) y en la Tabla S5.",
    "Finalmente, para cuantificar la sensibilidad de los estimadores":
        " Finalmente, para cuantificar la sensibilidad de los estimadores a la confusión no "
        "medida se calcularon E-values (VanderWeele y Ding, 2017) (30) para la razón de prevalencia "
        "del PHQ-9 en los modelos principales y en la cascada de cuidado (Tabla S9).",
    "El analisis de E-values reforzo esta interpretacion":
        " El análisis de E-values reforzó esta interpretación: dado que el intervalo de "
        "confianza del Modelo 2 incluye la unidad, no se requiere confusión no medida alguna para "
        "ser compatible con la ausencia de efecto, y la pequeña asociación inversa observada sin "
        "ajuste por altitud tendría un E-value de apenas 1,08 (1,04 para el límite del intervalo "
        "de confianza), de modo que un confusor tan débil como la propia altitud basta para "
        "explicarla por completo (Tabla S9).",
    "No obstante, la magnitud de esta confusion potencial":
        " No obstante, la magnitud de esta confusión potencial es acotada: los E-values de los "
        "estimadores principales fueron muy bajos (≤ 1,11), lo que indica que asociaciones no "
        "medidas modestas bastarían para explicar las pequeñas señales observadas, en línea con "
        "la interpretación de que no existe un efecto causal independiente (Tabla S9).",
    "E-values de los estimadores principales (VanderWeele":
        "E-values de los estimadores principales (VanderWeele y Ding, 2017). Para cada razón de "
        "prevalencia (o de momios, en la sensibilidad logística) se reporta el E-value de la "
        "estimación puntual y el del límite del intervalo de confianza más cercano a la unidad. El "
        "E-value es la fuerza mínima de asociación (en escala de razón) que un confusor no medido "
        "debería tener con la exposición y con el desenlace, por encima de los confusores ya "
        "ajustados, para explicar por completo la asociación observada.",
}


def main():
    d = docx.Document(str(DOC))
    hits = {k: 0 for k in FIXES}
    for p in d.paragraphs:
        for run in p.runs:
            for anchor, correct in FIXES.items():
                if anchor in run.text:
                    run.text = correct
                    hits[anchor] += 1
    missing = [k for k, v in hits.items() if v != 1]
    if missing:
        raise AssertionError(f"Anclas no resueltas (esperaba 1 hit c/u): {[(k, hits[k]) for k in missing]}")
    d.save(str(DOC))
    print("OK - acentos corregidos:", {k[:30]: v for k, v in hits.items()})


if __name__ == "__main__":
    main()
