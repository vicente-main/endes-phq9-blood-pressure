"""Figura S2 — DAG conceptual rediseñado para publicación.

Versión con todas las correcciones contra convenciones publicadas (Tennant 2021 IJE;
J Clin Epidemiol DAG tutorial):

1. Flujo izquierda → derecha en el eje principal (PHQ-9 → mediadores → PAE).
2. Mediadores en bandas superior e inferior del eje principal (estilo "diamante"),
   evitando que mediador → PAE vaya hacia arriba.
3. Nodo U (factores no medidos) con flechas DASHED a PHQ-9 y PAE.
4. Aristas confusor → mediador seleccionadas (las más fundamentadas).
5. Correlaciones entre confusores (RIQUEZA ↔ EDUC, EDAD ↔ ANIO) como
   arcos bidireccionales dashed.
6. Etiqueta "Efecto total a estimar" trasladada al caption del manuscrito.
7. Variables de diseño NO se grafican (entran al svydesign).
8. TIEMPO_DX_HTA_MESES NO se grafica (excluida del estudio).

La especificación formal del DAG (con isAcyclic() y adjustmentSets()) sigue
viviendo en 02_DAG_modelos.R; este script solo regenera el visual editorial.

Salidas:
  data/output/Auditoria_Integral/Figura_S2_DAG.{svg,png,pdf}
"""
from __future__ import annotations

from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyArrowPatch, FancyBboxPatch
import numpy as np


ROOT = Path(__file__).resolve().parents[2]
OUT_DIR = ROOT / "data" / "output" / "Auditoria_Integral"


# ---------------------------------------------------------------------------
# Paleta de colores por rol
# ---------------------------------------------------------------------------

COLOR = {
    "confounder":   {"fill": "#D4EDDA", "edge": "#155724", "arrow": "#2E7D32"},  # verde
    "exposure":     {"fill": "#90CAF9", "edge": "#0D47A1", "arrow": "#0D47A1"},  # azul
    "outcome":      {"fill": "#F5B7B1", "edge": "#922B21", "arrow": "#922B21"},  # rojo
    "mediator":     {"fill": "#FFE0B2", "edge": "#E65100", "arrow": "#E65100"},  # naranja
    "unmeasured":   {"fill": "#F3E5F5", "edge": "#6A1B9A", "arrow": "#6A1B9A"},  # violeta
}


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def draw_box(ax, x, y, label, role, width=1.8, height=0.7, fontsize=9,
             weight="normal", dashed_border=False):
    style = COLOR[role]
    box = FancyBboxPatch(
        (x - width / 2, y - height / 2),
        width, height,
        boxstyle="round,pad=0.02,rounding_size=0.10",
        facecolor=style["fill"],
        edgecolor=style["edge"],
        linewidth=1.0,
        linestyle="--" if dashed_border else "-",
        zorder=3,
    )
    ax.add_patch(box)
    ax.text(x, y, label,
            ha="center", va="center", fontsize=fontsize, weight=weight,
            color=style["edge"], zorder=4)
    return (x, y, width, height)


def edge_point(node, direction):
    x, y, w, h = node
    if direction == "bottom":   return (x, y - h / 2)
    if direction == "top":      return (x, y + h / 2)
    if direction == "left":     return (x - w / 2, y)
    if direction == "right":    return (x + w / 2, y)
    return (x, y)


def draw_arrow(ax, src, dst, role="confounder",
               connectionstyle="arc3,rad=0",
               linewidth=0.7, alpha=0.55, zorder=2, linestyle="-"):
    arrow = FancyArrowPatch(
        src, dst,
        arrowstyle="-|>",
        mutation_scale=14,
        linewidth=linewidth,
        color=COLOR[role]["arrow"],
        alpha=alpha,
        connectionstyle=connectionstyle,
        zorder=zorder,
        linestyle=linestyle,
    )
    ax.add_patch(arrow)


def draw_correlation(ax, src, dst, color="#155724", linewidth=0.8, alpha=0.5):
    """Dibuja una correlación bidireccional dashed entre dos puntos."""
    arrow = FancyArrowPatch(
        src, dst,
        arrowstyle="<|-|>",
        mutation_scale=8,
        linewidth=linewidth,
        color=color,
        alpha=alpha,
        connectionstyle="arc3,rad=-0.18",
        linestyle="--",
        zorder=1,
    )
    ax.add_patch(arrow)


# ---------------------------------------------------------------------------
# DAG layout
# ---------------------------------------------------------------------------

def build_dag_figure(out_dir: Path) -> None:
    fig, ax = plt.subplots(figsize=(16, 10.5))
    ax.set_xlim(-0.9, 16.4)
    ax.set_ylim(-0.5, 11.2)
    ax.axis("off")
    ax.set_aspect("equal", adjustable="datalim")

    # Título
    ax.text(7.75, 10.55, "Figura S2. Grafo dirigido acíclico (DAG) del estudio",
            ha="center", va="center", fontsize=13.5, weight="bold")
    ax.text(7.75, 10.10, "PHQ-9 → presión arterial elevada — ENDES 2019-2024",
            ha="center", va="center", fontsize=10.5, style="italic", color="#424242")

    # --- Confusores (banda superior) ---------------------------------------
    ax.text(7.75, 9.45, "Confusores estructurales — incluidos en Modelo 2 (principal)",
            ha="center", va="center", fontsize=10, weight="bold",
            color=COLOR["confounder"]["edge"])
    # Posiciones alineadas dentro del eje PHQ-9 (x=1.5) → PAE (x=14.5) para que
    # las flechas confusor→exposición/desenlace queden DENTRO del cuadro y no
    # se desborden hacia las esquinas.
    confounders = [
        ("EDAD",    "Edad",                 1.7),
        ("SEXO",    "Sexo",                 3.5),
        ("EDUC",    "Educación",            5.3),
        ("AREA",    "Área urbano/rural",    7.1),
        ("ALTITUD", "Altitud",              8.9),
        ("RIQUEZA", "Quintil de riqueza",  10.7),
        ("VPAR",    "Violencia\nde pareja", 12.5),
        ("ANIO",    "Año del estudio",     14.3),
    ]
    confounder_nodes = {}
    for key, label, x in confounders:
        node = draw_box(ax, x, 8.65, label, role="confounder",
                        width=1.7, height=0.7, fontsize=8.0)
        confounder_nodes[key] = node

    # --- Correlaciones entre confusores (dashed bidireccional) -------------
    # Ejemplos teóricos importantes:
    #   EDUC ↔ RIQUEZA (correlación socioeconómica)
    #   EDAD ↔ ANIO (estructura de cohorte)
    edu = confounder_nodes["EDUC"]
    riq = confounder_nodes["RIQUEZA"]
    edad = confounder_nodes["EDAD"]
    anio = confounder_nodes["ANIO"]
    draw_correlation(ax,
                     (edu[0], edu[1] + edu[3] / 2),
                     (riq[0], riq[1] + riq[3] / 2),
                     color="#155724", linewidth=0.9, alpha=0.45)
    draw_correlation(ax,
                     (edad[0], edad[1] + edad[3] / 2),
                     (anio[0], anio[1] + anio[3] / 2),
                     color="#155724", linewidth=0.9, alpha=0.45)
    # Etiqueta de correlación (pequeña, sobre la curva superior)
    ax.text(7.75, 9.85, "↔  correlaciones entre confusores",
            ha="center", va="center", fontsize=7.5, style="italic", color="#155724")

    # --- Eje principal: PHQ-9 → PAE (izquierda → derecha) -----------------
    phq9_node = draw_box(ax, 1.5, 5.3,
                         "Síntomas\ndepresivos\n(PHQ-9)",
                         role="exposure",
                         width=2.0, height=1.25, fontsize=10, weight="bold")
    pae_node  = draw_box(ax, 14.5, 5.3,
                         "Presión\narterial\nelevada",
                         role="outcome",
                         width=2.0, height=1.25, fontsize=10, weight="bold")

    # Flecha directa PHQ-9 → PAE (efecto total)
    src = edge_point(phq9_node, "right")
    dst = edge_point(pae_node, "left")
    arrow = FancyArrowPatch(
        src, dst,
        arrowstyle="-|>",
        mutation_scale=22,
        linewidth=2.5,
        color="#1F2937",
        alpha=0.95,
        zorder=5,
    )
    ax.add_patch(arrow)
    # La etiqueta "Efecto total a estimar" va al CAPTION (no en la figura)
    # según convención editorial.

    # Confusores → PHQ-9 y → PAE: LÍNEAS RECTAS con anclaje repartido sobre los
    # bordes de PHQ-9 y PAE. Cada confusor aterriza en un punto distinto del borde
    # (abanico), aprovechando el espacio lateral y evitando que todas las líneas
    # converjan en un único punto. Las que van al objetivo cercano usan el borde
    # superior; las que cruzan al objetivo lejano usan el borde lateral exterior
    # (izquierdo de PHQ-9 / derecho de PAE), que es espacio muerto.
    phq_x, phq_y, phq_w, phq_h = phq9_node
    pae_x, pae_y, pae_w, pae_h = pae_node
    n_conf = len(confounders)
    for i, (key, _label, x) in enumerate(confounders):
        c_node = confounder_nodes[key]
        src_c = edge_point(c_node, "bottom")
        frac = (i + 1) / (n_conf + 1)

        # → PHQ-9: si el confusor está a la derecha de PHQ-9, entra por el borde
        # superior repartido; los muy lejanos (derecha) aterrizan más a la derecha
        # del borde superior para no cruzar el centro.
        tx_phq = phq_x - phq_w / 2 + frac * phq_w
        dst_phq = (tx_phq, phq_y + phq_h / 2)
        draw_arrow(ax, src_c, dst_phq, role="confounder",
                   connectionstyle="arc3,rad=0", linewidth=1.0, alpha=0.6)

        # → PAE: anclaje repartido sobre el borde superior de PAE.
        tx_pae = pae_x - pae_w / 2 + frac * pae_w
        dst_pae = (tx_pae, pae_y + pae_h / 2)
        draw_arrow(ax, src_c, dst_pae, role="confounder",
                   connectionstyle="arc3,rad=0", linewidth=1.0, alpha=0.55)

    # --- Mediadores: 2 bandas (superior e inferior del eje principal) ------
    mediators_top = [
        ("IMC",      "IMC",                            5.0),
        ("TABACO",   "Consumo de\ntabaco (30 d)",      8.0),
        ("DIETA",    "Calidad\nde dieta",             11.0),
    ]
    mediators_bot = [
        ("CINTURA",  "Circunferencia\nabdominal",      5.0),
        ("ALC_PROB", "Consumo\nproblemático alcohol",  8.0),
        ("DIABETES", "Diagnóstico\nde diabetes",      11.0),
    ]
    mediator_nodes = {}
    for key, label, x in mediators_top:
        node = draw_box(ax, x, 6.85, label, role="mediator",
                        width=2.2, height=0.85, fontsize=8.5)
        mediator_nodes[key] = node
    for key, label, x in mediators_bot:
        node = draw_box(ax, x, 3.75, label, role="mediator",
                        width=2.2, height=0.85, fontsize=8.5)
        mediator_nodes[key] = node

    # PHQ-9 → cada mediador (izquierda → derecha, ligeras curvas)
    src_phq_right = edge_point(phq9_node, "right")
    src_phq_top = edge_point(phq9_node, "top")
    src_phq_bot = edge_point(phq9_node, "bottom")
    for key, _label, x in mediators_top:
        m_node = mediator_nodes[key]
        dst_m = edge_point(m_node, "left")
        draw_arrow(ax, src_phq_top, dst_m, role="mediator",
                   connectionstyle="arc3,rad=0",
                   linewidth=1.1, alpha=0.75)
    for key, _label, x in mediators_bot:
        m_node = mediator_nodes[key]
        dst_m = edge_point(m_node, "left")
        draw_arrow(ax, src_phq_bot, dst_m, role="mediator",
                   connectionstyle="arc3,rad=0",
                   linewidth=1.1, alpha=0.75)

    # Cada mediador → PAE
    dst_pae_top = edge_point(pae_node, "top")
    dst_pae_bot = edge_point(pae_node, "bottom")
    for key, _label, x in mediators_top:
        m_node = mediator_nodes[key]
        src_m = edge_point(m_node, "right")
        draw_arrow(ax, src_m, dst_pae_top, role="mediator",
                   connectionstyle="arc3,rad=0",
                   linewidth=1.1, alpha=0.75)
    for key, _label, x in mediators_bot:
        m_node = mediator_nodes[key]
        src_m = edge_point(m_node, "right")
        draw_arrow(ax, src_m, dst_pae_bot, role="mediator",
                   connectionstyle="arc3,rad=0",
                   linewidth=1.1, alpha=0.75)

    # --- Aristas confusor → mediador (selección teórica) -------------------
    # Las más fundamentadas en literatura:
    cm_edges = [
        ("EDAD",    "IMC"),       # envejecimiento → adiposidad
        ("SEXO",    "CINTURA"),   # dimorfismo sexual en distribución grasa
        ("RIQUEZA", "DIETA"),     # SES → patrones alimentarios
        ("EDUC",    "DIETA"),     # educación → alfabetización nutricional
        ("EDAD",    "DIABETES"),  # mayor edad → mayor prevalencia de diabetes
    ]
    for src_key, dst_key in cm_edges:
        c_node = confounder_nodes[src_key]
        m_node = mediator_nodes[dst_key]
        src_c = edge_point(c_node, "bottom")
        dst_m = edge_point(m_node, "top")
        draw_arrow(ax, src_c, dst_m, role="confounder",
                   connectionstyle="arc3,rad=0",
                   linewidth=0.9, alpha=0.55, linestyle="-")

    # Etiqueta de mediadores (debajo de la fila inferior, fuera de aristas)
    ax.text(7.75, 2.95,
            "Mediadores potenciales — añadidos en Modelo 3 (exploratorio)",
            ha="center", va="center", fontsize=10, weight="bold",
            color=COLOR["mediator"]["edge"])

    # --- Nodo U (factores no medidos) -------------------------------------
    u_node = draw_box(ax, 7.75, 1.85,
                      "U: factores no medidos\n(estrés crónico, sueño, comorbilidad,\nadherencia previa, sesgo de detección)",
                      role="unmeasured",
                      width=4.6, height=0.95, fontsize=8.5,
                      dashed_border=True)

    # U → PHQ-9 y U → PAE (dashed para indicar hipotético/no medido)
    src_u_left = edge_point(u_node, "left")
    src_u_right = edge_point(u_node, "right")
    dst_phq_bot = edge_point(phq9_node, "bottom")
    dst_pae_bot2 = edge_point(pae_node, "bottom")
    draw_arrow(ax, src_u_left, dst_phq_bot, role="unmeasured",
               connectionstyle="arc3,rad=0.15",
               linewidth=0.9, alpha=0.65, linestyle="--")
    draw_arrow(ax, src_u_right, dst_pae_bot2, role="unmeasured",
               connectionstyle="arc3,rad=-0.15",
               linewidth=0.9, alpha=0.65, linestyle="--")

    # --- Leyenda horizontal (5 items, espaciados) --------------------------
    legend_items = [
        ("Confusor (Modelo 2)",          "confounder"),
        ("Exposición (PHQ-9)",            "exposure"),
        ("Desenlace (PA elevada)",        "outcome"),
        ("Mediador (Modelo 3)",           "mediator"),
        ("U: no medidos (hipotético)",    "unmeasured"),
    ]
    legend_y = 0.40
    ax.text(0.3, legend_y, "Leyenda:", fontsize=10, weight="bold",
            va="center", ha="left")
    legend_xs = [2.3, 5.1, 7.9, 10.7, 13.3]
    for (txt, role), xc in zip(legend_items, legend_xs):
        ls = "--" if role == "unmeasured" else "-"
        rect = FancyBboxPatch(
            (xc - 0.20, legend_y - 0.15), 0.40, 0.30,
            boxstyle="round,pad=0.0,rounding_size=0.06",
            facecolor=COLOR[role]["fill"],
            edgecolor=COLOR[role]["edge"],
            linewidth=1.0,
            linestyle=ls,
        )
        ax.add_patch(rect)
        ax.text(xc + 0.32, legend_y, txt, fontsize=9, va="center", ha="left",
                color="#222")

    # Exportar
    out_dir.mkdir(parents=True, exist_ok=True)
    svg = out_dir / "Figura_S2_DAG.svg"
    png = out_dir / "Figura_S2_DAG.png"
    pdf = out_dir / "Figura_S2_DAG.pdf"
    plt.savefig(svg, format="svg", bbox_inches="tight")
    plt.savefig(png, format="png", bbox_inches="tight", dpi=200)
    plt.savefig(pdf, format="pdf", bbox_inches="tight")
    plt.close(fig)
    print(f"  -> {svg.relative_to(ROOT)}")
    print(f"  -> {png.relative_to(ROOT)}")
    print(f"  -> {pdf.relative_to(ROOT)}")


def main() -> None:
    print("Figura S2 (DAG rediseñado, A-grade):")
    build_dag_figure(OUT_DIR)


if __name__ == "__main__":
    main()
