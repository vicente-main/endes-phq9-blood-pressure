"""Fuente UNICA de etiquetas para QS25N (nivel educativo, ENDES CSALUD01).

Auditoria 2026-07-16 (hallazgo #3): los scripts de presentacion mapeaban las
etiquetas DESPLAZADAS una posicion (asumian 1..5) y omitian "Postgrado".
Codificacion real del diccionario oficial INEI (Diccionario - CSALUD01.pdf):

    0 = Inicial, pre-escolar
    1 = Primaria
    2 = Secundaria
    3 = Superior No Universitaria
    4 = Superior Universitaria
    5 = Postgrado

Ademas, desde la correccion de la compuerta QS24 (hallazgo #2, pipeline.py),
la categoria 0 incluye tambien a quienes NUNCA asistieron a la escuela
(QS24 == 2), por lo que su etiqueta conjunta es "Sin nivel / inicial".
Postgrado (5) se reporta como fila propia; NO se fusiona con universitaria.

Todo script de presentacion debe importar de aqui; no redefinir estos dicts.
"""

# --- Espanol (tablas internas / Post_Auditoria) ---
EDU_LABELS_ES = {
    0: "Sin nivel / inicial",
    1: "Primaria",
    2: "Secundaria",
    3: "Superior no universitaria",
    4: "Superior universitaria",
    5: "Postgrado",
}
EDU_ORDER_ES = [EDU_LABELS_ES[k] for k in range(6)]

# --- Ingles (paquete PLOS ONE) ---
EDU_LABELS_EN = {
    0: "No formal education / initial",
    1: "Primary",
    2: "Secondary",
    3: "Higher non-university",
    4: "Higher university",
    5: "Postgraduate",
}
EDU_ORDER_EN = [EDU_LABELS_EN[k] for k in range(6)]

# Terminos de modelo factor(QS25N)k -> etiqueta (referencia = 0, nivel mas bajo).
EDU_MODEL_TERMS_ES = {
    "factor(QS25N)0": f"{EDU_LABELS_ES[0]} (referencia)",
    **{f"factor(QS25N){k}": EDU_LABELS_ES[k] for k in range(1, 6)},
}
EDU_MODEL_TERMS_EN = {
    "factor(QS25N)0": f"{EDU_LABELS_EN[0]} (reference)",
    **{f"factor(QS25N){k}": EDU_LABELS_EN[k] for k in range(1, 6)},
}

# Nota al pie sugerida cuando se muestre la variable en una tabla publicable.
EDU_FOOTNOTE_ES = (
    "Nivel educativo (QS25N): 0 incluye 'nunca asistio' (QS24=2) e 'inicial/pre-escolar'; "
    "'Postgrado' se reporta como categoria propia segun el diccionario ENDES."
)
EDU_FOOTNOTE_EN = (
    "Educational level (QS25N): category 0 includes 'never attended school' (QS24=2) and "
    "'initial/pre-school'; 'Postgraduate' is reported as its own category per the ENDES dictionary."
)
