from pathlib import Path

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt


def parse_num(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series.astype(str).str.replace(',', '.', regex=False), errors='coerce')


def categorize(series: pd.Series, min_v: float, max_v: float, edges: list[float], labels: list[str]) -> pd.Categorical:
    out = pd.Series(index=series.index, dtype='object')

    out[series.isna()] = 'Sin dato'
    out[series < min_v] = f'< {min_v} (fuera)'
    out[series > max_v] = f'> {max_v} (fuera)'

    in_range = series.ge(min_v) & series.le(max_v)
    out.loc[in_range] = pd.cut(
        series.loc[in_range],
        bins=edges,
        labels=labels,
        include_lowest=True,
        right=True,
        ordered=True,
    ).astype(str)

    order = [f'< {min_v} (fuera)', *labels, f'> {max_v} (fuera)', 'Sin dato']
    return pd.Categorical(out, categories=order, ordered=True)


def build_distribution_table(df: pd.DataFrame) -> pd.DataFrame:
    specs = {
        'PAS_PROM': {
            'title': 'PA Sistolica promedio (mmHg)',
            'series': (parse_num(df['QS903S']) + parse_num(df['QS905S'])) / 2,
            'min': 70,
            'max': 270,
            'edges': [70, 90, 110, 130, 140, 160, 180, 200, 270],
            'labels': ['70-90', '91-110', '111-130', '131-140', '141-160', '161-180', '181-200', '201-270'],
        },
        'PAD_PROM': {
            'title': 'PA Diastolica promedio (mmHg)',
            'series': (parse_num(df['QS903D']) + parse_num(df['QS905D'])) / 2,
            'min': 40,
            'max': 180,
            'edges': [40, 60, 70, 80, 90, 100, 110, 120, 180],
            'labels': ['40-60', '61-70', '71-80', '81-90', '91-100', '101-110', '111-120', '121-180'],
        },
        'PESO_KG': {
            'title': 'Peso (kg)',
            'series': parse_num(df['QS900']),
            'min': 25,
            'max': 300,
            'edges': [25, 50, 70, 90, 110, 130, 160, 300],
            'labels': ['25-50', '51-70', '71-90', '91-110', '111-130', '131-160', '161-300'],
        },
        'TALLA_CM': {
            'title': 'Talla (cm)',
            'series': parse_num(df['QS901']),
            'min': 120,
            'max': 230,
            'edges': [120, 140, 150, 160, 170, 180, 230],
            'labels': ['120-140', '141-150', '151-160', '161-170', '171-180', '181-230'],
        },
        'EDAD': {
            'title': 'Edad (anios)',
            'series': pd.to_numeric(df['EDAD'], errors='coerce'),
            'min': 18,
            'max': 97,
            'edges': [18, 30, 40, 50, 60, 70, 80, 90, 97],
            'labels': ['18-30', '31-40', '41-50', '51-60', '61-70', '71-80', '81-90', '91-97'],
        },
    }

    rows = []
    n_total = len(df)

    for key, cfg in specs.items():
        cat = categorize(cfg['series'], cfg['min'], cfg['max'], cfg['edges'], cfg['labels'])
        counts = pd.Series(cat).value_counts(sort=False, dropna=False)

        for cat_name, n in counts.items():
            rows.append(
                {
                    'variable': key,
                    'titulo': cfg['title'],
                    'categoria': str(cat_name),
                    'n': int(n),
                    'pct': round((int(n) / n_total) * 100, 4),
                }
            )

    return pd.DataFrame(rows)


def plot_distribution_bars(dist: pd.DataFrame, out_png: Path, out_pdf: Path) -> None:
    variables = ['PAS_PROM', 'PAD_PROM', 'PESO_KG', 'TALLA_CM', 'EDAD']

    fig, axes = plt.subplots(3, 2, figsize=(20, 15))
    axes = axes.flatten()

    for i, var in enumerate(variables):
        ax = axes[i]
        sub = dist[dist['variable'] == var].copy()
        title = sub['titulo'].iloc[0]

        colors = []
        for c in sub['categoria']:
            c_lower = c.lower()
            if 'fuera' in c_lower:
                colors.append('#d7301f')
            elif 'sin dato' in c_lower:
                colors.append('#969696')
            else:
                colors.append('#3182bd')

        bars = ax.bar(sub['categoria'], sub['n'], color=colors, edgecolor='black', linewidth=0.4)
        ax.set_title(title, fontsize=12, fontweight='bold')
        ax.set_ylabel('Frecuencia (n)')
        ax.tick_params(axis='x', rotation=35, labelsize=9)

        y_max = sub['n'].max()
        for bar, pct in zip(bars, sub['pct']):
            ax.text(
                bar.get_x() + bar.get_width() / 2,
                bar.get_height() + y_max * 0.01,
                f'{pct:.1f}%',
                ha='center',
                va='bottom',
                fontsize=8,
            )

    axes[-1].axis('off')
    fig.suptitle('Distribucion de variables clinicas para definir valores imposibles (ENDES 2019-2024)', fontsize=15, fontweight='bold')
    fig.tight_layout(rect=(0, 0, 1, 0.97))

    fig.savefig(out_png, dpi=300)
    fig.savefig(out_pdf)
    plt.close(fig)


def main() -> None:
    root = Path(__file__).resolve().parents[1]
    send_dir = root / 'Para enviar'
    send_dir.mkdir(parents=True, exist_ok=True)

    data_path = root / 'data' / 'output' / 'endes_hta_depresion_2019_2024.csv'
    df = pd.read_csv(data_path, low_memory=False)

    dist = build_distribution_table(df)

    dist_csv = send_dir / '04_barras_distribucion_variables_tabla.csv'
    dist.to_csv(dist_csv, index=False, encoding='utf-8-sig')

    out_png = send_dir / '04_barras_distribucion_variables.png'
    out_pdf = send_dir / '04_barras_distribucion_variables.pdf'
    plot_distribution_bars(dist, out_png, out_pdf)

    print('created', dist_csv)
    print('created', out_png)
    print('created', out_pdf)


if __name__ == '__main__':
    main()
