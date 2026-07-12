from __future__ import annotations

from pathlib import Path

import pandas as pd
from openpyxl import load_workbook


ROOT = Path(__file__).resolve().parents[1]
PUBLIC_DIR = ROOT / "data" / "output" / "Para Enviar" / "Para Publicar"


GUIDE_PAPERS: list[dict] = [
    {
        "paper_id": "fpsyt_15_1433990",
        "filename": "fpsyt-15-1433990.pdf",
        "short_citation": "Frontiers in Psychiatry 2024",
        "title": "Association of diastolic and systolic blood pressure with depression: a cross-sectional study from NHANES 2005-2018",
        "journal": "Frontiers in Psychiatry",
        "year": 2024,
        "design": "Cross-sectional NHANES analysis",
        "population": "US adults, NHANES 2005-2018, n=26,581",
        "main_exposure": "SBP and DBP",
        "main_outcome": "Prevalent depression (PHQ-9 >= 10)",
        "displays": [
            {
                "display_id": "Figure 1",
                "display_kind": "Figure",
                "display_type": "flow/STROBE",
                "location": "principal",
                "caption_or_label": "Participant selection flow from NHANES 2005-2018",
                "analytic_role": "Sample derivation",
                "objective": "Show exclusions and final analytic sample.",
                "central_variable": "Analytic population",
                "estimator": "Counts and exclusions",
                "information_load": "medium",
                "visual_resources": "Flowchart with boxes, arrows, n values",
                "strengths": "Transparent and standard for observational studies.",
                "weaknesses": "Supports context only; no substantive result.",
                "problem_solved": "Documents sample construction and exclusion logic.",
            },
            {
                "display_id": "Table 1",
                "display_kind": "Table",
                "display_type": "descriptivo basal",
                "location": "principal",
                "caption_or_label": "Weighted characteristics of the eligible 26,581 participants",
                "analytic_role": "Baseline context",
                "objective": "Describe the analytic sample and compare depressed vs non-depressed participants.",
                "central_variable": "Depression status and participant characteristics",
                "estimator": "Weighted means/proportions with p values",
                "information_load": "high",
                "visual_resources": "Dense multicolumn baseline table",
                "strengths": "Recognizable, complete, and easy to cite in Results.",
                "weaknesses": "Heavy table; lower scanability.",
                "problem_solved": "Gives clinical and sociodemographic context before modeling.",
            },
            {
                "display_id": "Figure 2",
                "display_kind": "Figure",
                "display_type": "curva dosis-respuesta/spline",
                "location": "principal",
                "caption_or_label": "Restricted cubic spline for DBP and prevalent depression",
                "analytic_role": "Functional form assessment",
                "objective": "Show whether the DBP-depression relationship is linear or nonlinear.",
                "central_variable": "DBP vs depression",
                "estimator": "Restricted cubic spline with nonlinear p value",
                "information_load": "high",
                "visual_resources": "Curve with CI, subgroup panels",
                "strengths": "Justifies the chosen functional form with a visual argument.",
                "weaknesses": "Panel complexity; requires careful captioning.",
                "problem_solved": "Shows that the relationship is close to linear and supports model specification.",
            },
            {
                "display_id": "Figure 3",
                "display_kind": "Figure",
                "display_type": "curva dosis-respuesta/spline",
                "location": "principal",
                "caption_or_label": "Spline and threshold analysis for SBP and prevalent depression",
                "analytic_role": "Nonlinear signal and inflection point",
                "objective": "Show the U-shaped pattern and the turning point for SBP.",
                "central_variable": "SBP vs depression",
                "estimator": "Restricted cubic spline and two-piecewise linear regression",
                "information_load": "high",
                "visual_resources": "Curve plus threshold panel",
                "strengths": "The figure carries a substantive finding, not just a method check.",
                "weaknesses": "Methodologically dense; depends on caption clarity.",
                "problem_solved": "Makes the threshold finding legible to readers.",
            },
            {
                "display_id": "Table 2",
                "display_kind": "Table",
                "display_type": "modelo principal",
                "location": "principal",
                "caption_or_label": "Weighted multivariate logistic regression analysis for the association of prevalent depression with SBP and DBP",
                "analytic_role": "Primary association estimates",
                "objective": "Quantify adjusted associations for the main exposure-outcome relationship.",
                "central_variable": "SBP/DBP vs depression",
                "estimator": "Weighted logistic OR with 95% CI",
                "information_load": "high",
                "visual_resources": "Compact model table with serial adjustments",
                "strengths": "Core inferential display; aligns tightly with the research question.",
                "weaknesses": "Longer once categorical and continuous analyses are combined.",
                "problem_solved": "Presents the main adjusted estimates in a citation-ready format.",
            },
            {
                "display_id": "Figure 4",
                "display_kind": "Figure",
                "display_type": "subgrupo/interaccion",
                "location": "principal",
                "caption_or_label": "Association between depression and blood pressure across subgroups",
                "analytic_role": "Effect modification",
                "objective": "Show subgroup heterogeneity and interaction results.",
                "central_variable": "SBP/DBP-depression association across strata",
                "estimator": "Subgroup OR and interaction p values",
                "information_load": "medium",
                "visual_resources": "Forest-style subgroup display",
                "strengths": "Efficient summary of heterogeneity.",
                "weaknesses": "Secondary result; can distract if underpowered.",
                "problem_solved": "Lets readers assess whether findings are stable across key strata.",
            },
            {
                "display_id": "Table 3",
                "display_kind": "Table",
                "display_type": "sensibilidad/robustez",
                "location": "principal",
                "caption_or_label": "Sensitivity analysis for the association of prevalent depression with SBP and DBP",
                "analytic_role": "Robustness confirmation",
                "objective": "Show that results hold after alternate analytic decisions.",
                "central_variable": "Main blood pressure findings under sensitivity analyses",
                "estimator": "Alternative model estimates",
                "information_load": "medium",
                "visual_resources": "Short results table",
                "strengths": "Adds credibility without needing another figure.",
                "weaknesses": "Secondary by nature.",
                "problem_solved": "Condenses robustness checks into one place.",
            },
        ],
    },
    {
        "paper_id": "fpubh_12_1461300",
        "filename": "fpubh-12-1461300.pdf",
        "short_citation": "Frontiers in Public Health 2024",
        "title": "Association between weight-adjusted waist circumference index and depression in older patients with hypertension: a study based on NHANES 2007-2016",
        "journal": "Frontiers in Public Health",
        "year": 2024,
        "design": "Cross-sectional NHANES analysis",
        "population": "Older adults with hypertension, NHANES 2007-2016, n=4,228",
        "main_exposure": "WWI",
        "main_outcome": "Depression",
        "displays": [
            {
                "display_id": "Figure 1",
                "display_kind": "Figure",
                "display_type": "flow/STROBE",
                "location": "principal",
                "caption_or_label": "Participant selection flowchart",
                "analytic_role": "Sample derivation",
                "objective": "Show exclusions leading to the final hypertensive sample.",
                "central_variable": "Analytic population",
                "estimator": "Counts and exclusion reasons",
                "information_load": "medium",
                "visual_resources": "Flowchart",
                "strengths": "Standard and easy to interpret.",
                "weaknesses": "Narrative support only.",
                "problem_solved": "Documents the denominator used in the study.",
            },
            {
                "display_id": "Table 1",
                "display_kind": "Table",
                "display_type": "descriptivo basal",
                "location": "principal",
                "caption_or_label": "Baseline characteristics of participants",
                "analytic_role": "Baseline context",
                "objective": "Describe participants with and without depression.",
                "central_variable": "Participant characteristics by depression status",
                "estimator": "Weighted means/proportions with significance tests",
                "information_load": "high",
                "visual_resources": "Dense baseline table",
                "strengths": "Standard, comprehensive, and directly linked to the target population.",
                "weaknesses": "Can become visually heavy.",
                "problem_solved": "Sets up the study population before regression.",
            },
            {
                "display_id": "Table 2",
                "display_kind": "Table",
                "display_type": "modelo principal",
                "location": "principal",
                "caption_or_label": "Logistic regression analysis between WWI index and depression",
                "analytic_role": "Primary association estimates",
                "objective": "Show the adjusted relationship between WWI and depression.",
                "central_variable": "WWI vs depression",
                "estimator": "Logistic OR with 95% CI",
                "information_load": "medium",
                "visual_resources": "Serial adjustment table",
                "strengths": "Focused and conventional.",
                "weaknesses": "Less visual than a summarized model figure.",
                "problem_solved": "Provides the main adjusted estimate.",
            },
            {
                "display_id": "Figure 2",
                "display_kind": "Figure",
                "display_type": "curva dosis-respuesta/spline",
                "location": "principal",
                "caption_or_label": "Smoothed curve fitting for WWI and depression",
                "analytic_role": "Functional form assessment",
                "objective": "Show the nonlinear positive association and threshold around WWI >= 11.6.",
                "central_variable": "WWI vs depression",
                "estimator": "Smoothed dose-response curve",
                "information_load": "medium",
                "visual_resources": "Smooth curve with CI",
                "strengths": "A figure is justified because it reveals a threshold-like pattern.",
                "weaknesses": "Needs the text to explain how the threshold is used analytically.",
                "problem_solved": "Makes the nonlinear relationship visible.",
            },
            {
                "display_id": "Figure 3",
                "display_kind": "Figure",
                "display_type": "subgrupo/interaccion",
                "location": "principal",
                "caption_or_label": "Subgroup analyses by age, sex, and BMI",
                "analytic_role": "Effect heterogeneity",
                "objective": "Show whether the main association is stronger in specific subgroups.",
                "central_variable": "WWI-depression association across strata",
                "estimator": "Subgroup OR and interaction testing",
                "information_load": "medium",
                "visual_resources": "Forest-style subgroup display",
                "strengths": "Compact way to show robustness and heterogeneity.",
                "weaknesses": "Secondary result; can be overread.",
                "problem_solved": "Shows where the association is concentrated.",
            },
        ],
    },
    {
        "paper_id": "s12888_021_03275_2",
        "filename": "s12888-021-03275-2.pdf",
        "short_citation": "BMC Psychiatry 2021",
        "title": "The association between triglyceride glucose index and depression: data from NHANES 2005-2018",
        "journal": "BMC Psychiatry",
        "year": 2021,
        "design": "Cross-sectional NHANES analysis",
        "population": "US adults, NHANES 2005-2018, n=13,350",
        "main_exposure": "TyG index",
        "main_outcome": "Depression",
        "displays": [
            {
                "display_id": "Figure 1",
                "display_kind": "Figure",
                "display_type": "flow/STROBE",
                "location": "principal",
                "caption_or_label": "Flow chart of subject selection",
                "analytic_role": "Sample derivation",
                "objective": "Show the excluded participants and final analytic sample.",
                "central_variable": "Analytic population",
                "estimator": "Counts and exclusions",
                "information_load": "medium",
                "visual_resources": "Flow diagram",
                "strengths": "Transparent and concise.",
                "weaknesses": "Purely contextual.",
                "problem_solved": "Documents who entered the analysis.",
            },
            {
                "display_id": "Table 1",
                "display_kind": "Table",
                "display_type": "descriptivo basal",
                "location": "principal",
                "caption_or_label": "Weighted baseline characteristics according to triglyceride-glucose index quartile",
                "analytic_role": "Baseline context",
                "objective": "Describe participants across exposure quartiles.",
                "central_variable": "TyG quartiles and participant characteristics",
                "estimator": "Weighted proportions/means by quartile",
                "information_load": "high",
                "visual_resources": "Exposure-stratified baseline table",
                "strengths": "Directly tied to the exposure framing used in the models.",
                "weaknesses": "Dense and long.",
                "problem_solved": "Shows how the exposure is distributed across the population.",
            },
            {
                "display_id": "Figure 2",
                "display_kind": "Figure",
                "display_type": "curva dosis-respuesta/spline",
                "location": "principal",
                "caption_or_label": "Generalized additive model curve for TyG index and depression",
                "analytic_role": "Functional form assessment",
                "objective": "Show whether the TyG-depression relationship is linear or nonlinear.",
                "central_variable": "TyG index vs depression",
                "estimator": "Generalized additive model curve",
                "information_load": "medium",
                "visual_resources": "Smooth curve with confidence band",
                "strengths": "The figure helps justify linear modeling even without a nonlinear signal.",
                "weaknesses": "Its substantive yield is lower because the relationship is reported as linear.",
                "problem_solved": "Documents functional form before presenting adjusted estimates.",
            },
            {
                "display_id": "Table 2",
                "display_kind": "Table",
                "display_type": "modelo principal",
                "location": "principal",
                "caption_or_label": "Weighted relationship between triglyceride-glucose index and depression",
                "analytic_role": "Primary association estimates",
                "objective": "Show adjusted associations across TyG parameterizations.",
                "central_variable": "TyG index vs depression",
                "estimator": "Adjusted OR with 95% CI",
                "information_load": "medium",
                "visual_resources": "Model table with multiple adjustment sets",
                "strengths": "Focused and conventional.",
                "weaknesses": "Carries both the main model and secondary covariates in one display.",
                "problem_solved": "Provides the central inferential result.",
            },
            {
                "display_id": "Figure 3",
                "display_kind": "Figure",
                "display_type": "subgrupo/interaccion",
                "location": "principal",
                "caption_or_label": "Subgroup analysis of the association between TyG index and depression",
                "analytic_role": "Robustness across strata",
                "objective": "Show that the association is similar across most subpopulations.",
                "central_variable": "TyG-depression association across strata",
                "estimator": "Stratified OR and interaction review",
                "information_load": "medium",
                "visual_resources": "Subgroup figure, likely forest-style",
                "strengths": "Compresses heterogeneity checks into a single display.",
                "weaknesses": "Secondary and potentially repetitive if the pattern is mostly null.",
                "problem_solved": "Shows where the association is stable.",
            },
            {
                "display_id": "Table S1",
                "display_kind": "Table",
                "display_type": "sensibilidad/robustez",
                "location": "soporte_referenciado_no_disponible",
                "caption_or_label": "Supplementary subgroup detail referenced in main text",
                "analytic_role": "Support detail",
                "objective": "Expand subgroup analyses mentioned in the main manuscript.",
                "central_variable": "TyG subgroup estimates",
                "estimator": "Supplementary model detail",
                "information_load": "high",
                "visual_resources": "Supplementary table",
                "strengths": "Prevents the main text from overloading.",
                "weaknesses": "Not available in the local PDF bundle.",
                "problem_solved": "Moves detailed heterogeneity results out of the main display set.",
            },
            {
                "display_id": "Table S2",
                "display_kind": "Table",
                "display_type": "otro",
                "location": "soporte_referenciado_no_disponible",
                "caption_or_label": "Supplementary univariate analysis of depression referenced in main text",
                "analytic_role": "Support detail",
                "objective": "Provide auxiliary descriptive or univariate results.",
                "central_variable": "Depression correlates",
                "estimator": "Univariate associations",
                "information_load": "high",
                "visual_resources": "Supplementary table",
                "strengths": "Keeps auxiliary detail out of the main paper.",
                "weaknesses": "Not available in the local PDF bundle.",
                "problem_solved": "Holds lower-priority supporting detail.",
            },
        ],
    },
    {
        "paper_id": "s12889_022_12942_2",
        "filename": "s12889-022-12942-2.pdf",
        "short_citation": "BMC Public Health 2022",
        "title": "Interaction between trouble sleeping and depression on hypertension in the NHANES 2005-2018",
        "journal": "BMC Public Health",
        "year": 2022,
        "design": "Cross-sectional NHANES analysis",
        "population": "US adults, NHANES 2005-2018, n=30,434",
        "main_exposure": "Trouble sleeping and depression",
        "main_outcome": "Hypertension",
        "displays": [
            {
                "display_id": "Table 1",
                "display_kind": "Table",
                "display_type": "descriptivo basal",
                "location": "principal",
                "caption_or_label": "Group differences between hypertension and non-hypertension groups",
                "analytic_role": "Baseline context",
                "objective": "Describe differences between participants with and without hypertension.",
                "central_variable": "Hypertension status and participant characteristics",
                "estimator": "Means/proportions with significance tests",
                "information_load": "high",
                "visual_resources": "Classical baseline table",
                "strengths": "Standard and clinically interpretable.",
                "weaknesses": "Dense and less visually memorable.",
                "problem_solved": "Summarizes the descriptive backbone of the study.",
            },
            {
                "display_id": "Table 2",
                "display_kind": "Table",
                "display_type": "subgrupo/interaccion",
                "location": "principal",
                "caption_or_label": "Interactive effect analysis of trouble sleeping and depression",
                "analytic_role": "Interaction result",
                "objective": "Show the joint association of trouble sleeping and depression on hypertension.",
                "central_variable": "Trouble sleeping x depression vs hypertension",
                "estimator": "OR, RERI, AP, S with 95% CI",
                "information_load": "high",
                "visual_resources": "Interaction table across models",
                "strengths": "A table is justified because additive interaction metrics need exact values.",
                "weaknesses": "Notation-heavy; less accessible to non-technical readers.",
                "problem_solved": "Presents the core interaction result with all required metrics.",
            },
            {
                "display_id": "Table 3",
                "display_kind": "Table",
                "display_type": "subgrupo/interaccion",
                "location": "principal",
                "caption_or_label": "Interactive effect analysis of trouble sleeping and depression severity",
                "analytic_role": "Interaction refinement",
                "objective": "Show whether the interaction differs by depression severity.",
                "central_variable": "Trouble sleeping x depression severity vs hypertension",
                "estimator": "OR, RERI, AP, S with 95% CI",
                "information_load": "high",
                "visual_resources": "Interaction table by severity category",
                "strengths": "Keeps the interaction story coherent and complete.",
                "weaknesses": "Secondary and notation-heavy.",
                "problem_solved": "Shows where the interaction is strongest.",
            },
        ],
    },
]
OUR_DISPLAYS: list[dict] = [
    {
        "display_id": "Tabla_1",
        "current_location": "principal",
        "display_type": "descriptivo basal",
        "current_role": "Describe the full analytic sample and stratified summaries.",
        "closest_guide_analogues": "All 4 guide papers use a baseline table.",
        "benchmark_read": "Aligned with guide-paper convention, but ours is more layered because it combines overall, outcome, sex, and severity panels.",
        "score_necesidad_editorial": 5,
        "score_claridad_mensaje": 3,
        "score_adecuacion_analitica": 5,
        "score_densidad_informativa": 3,
        "score_facilidad_lectura": 3,
        "score_consistencia_caption_contenido": 5,
        "score_utilidad_para_principal": 4,
        "score_adecuacion_rpmesp": 4,
        "veredicto": "refinar",
        "recommended_location": "principal",
        "why": "The table is necessary and methodologically strong, but the current multi-panel structure is denser than the benchmark papers and harder to scan quickly.",
        "recommended_change": "Keep as principal, but simplify the main-table version and move some stratified depth to supplement if manuscript space becomes tight.",
    },
    {
        "display_id": "Tabla_2",
        "current_location": "principal",
        "display_type": "bivariado",
        "current_role": "Summarize pooled Rao-Scott bivariate associations.",
        "closest_guide_analogues": "No guide paper relies on a technically labeled pooled bivariate table as strongly as we do.",
        "benchmark_read": "Functionally valid, but reader-facing polish is below benchmark because the term labels still expose internal placeholders.",
        "score_necesidad_editorial": 4,
        "score_claridad_mensaje": 2,
        "score_adecuacion_analitica": 4,
        "score_densidad_informativa": 4,
        "score_facilidad_lectura": 2,
        "score_consistencia_caption_contenido": 3,
        "score_utilidad_para_principal": 3,
        "score_adecuacion_rpmesp": 4,
        "veredicto": "refinar",
        "recommended_location": "principal",
        "why": "The analysis is legitimate, but the current presentation is too technical for a main-table reader because terms such as .tmp_pred leak implementation detail.",
        "recommended_change": "Keep as principal only if relabeled into clean reader-facing contrasts and stripped of internal placeholders.",
    },
    {
        "display_id": "Tabla_3",
        "current_location": "principal",
        "display_type": "modelo principal",
        "current_role": "Present the main multivariable models.",
        "closest_guide_analogues": "All guide papers include a main model table.",
        "benchmark_read": "This is the strongest and most benchmark-aligned table in the package.",
        "score_necesidad_editorial": 5,
        "score_claridad_mensaje": 4,
        "score_adecuacion_analitica": 5,
        "score_densidad_informativa": 4,
        "score_facilidad_lectura": 4,
        "score_consistencia_caption_contenido": 5,
        "score_utilidad_para_principal": 5,
        "score_adecuacion_rpmesp": 5,
        "veredicto": "mantener",
        "recommended_location": "principal",
        "why": "The table carries the main inferential result in the exact way readers expect.",
        "recommended_change": "Maintain as a principal table; only polish technical factor labels if the manuscript table is redrawn.",
    },
    {
        "display_id": "Tabla_4",
        "current_location": "principal",
        "display_type": "modelo principal",
        "current_role": "Present the hypertension care-cascade submodels.",
        "closest_guide_analogues": "Closest to robustness or interaction follow-up tables in fpsyt and s12889.",
        "benchmark_read": "Less universal than a main model table, but still justified because it expands the clinical interpretation of the main study.",
        "score_necesidad_editorial": 4,
        "score_claridad_mensaje": 4,
        "score_adecuacion_analitica": 5,
        "score_densidad_informativa": 4,
        "score_facilidad_lectura": 4,
        "score_consistencia_caption_contenido": 5,
        "score_utilidad_para_principal": 4,
        "score_adecuacion_rpmesp": 5,
        "veredicto": "mantener",
        "recommended_location": "principal",
        "why": "The table adds a clinically relevant downstream layer that is specific to our question and is not redundant with Table 3.",
        "recommended_change": "Maintain as principal; keep labels reader-facing if later redrawn for the manuscript.",
    },
    {
        "display_id": "Figura_1_STROBE",
        "current_location": "principal",
        "display_type": "flow/STROBE",
        "current_role": "Show the selection flow and exclusions.",
        "closest_guide_analogues": "Used in 3 of 4 guide papers.",
        "benchmark_read": "At benchmark level; it fits both guide-paper practice and RPMESP expectations.",
        "score_necesidad_editorial": 5,
        "score_claridad_mensaje": 5,
        "score_adecuacion_analitica": 5,
        "score_densidad_informativa": 4,
        "score_facilidad_lectura": 4,
        "score_consistencia_caption_contenido": 5,
        "score_utilidad_para_principal": 5,
        "score_adecuacion_rpmesp": 5,
        "veredicto": "mantener",
        "recommended_location": "principal",
        "why": "The flow figure is standard, necessary, and already corrected for publication use.",
        "recommended_change": "Maintain unchanged as a principal figure.",
    },
    {
        "display_id": "Figura_2_Spline",
        "current_location": "principal",
                "display_type": "curva dosis-respuesta/spline",
        "current_role": "Show the predicted probability curve for PHQ-9 and elevated blood pressure.",
        "closest_guide_analogues": "Comparable to Figure 2 in fpubh and Figure 2/3 in fpsyt.",
        "benchmark_read": "Methodologically acceptable, but weaker than benchmark nonlinear figures because ours does not carry a strong substantive nonlinear finding.",
        "score_necesidad_editorial": 3,
        "score_claridad_mensaje": 2,
        "score_adecuacion_analitica": 4,
        "score_densidad_informativa": 3,
        "score_facilidad_lectura": 4,
        "score_consistencia_caption_contenido": 4,
        "score_utilidad_para_principal": 2,
        "score_adecuacion_rpmesp": 3,
        "veredicto": "mover a suplementario",
        "recommended_location": "suplementario",
        "why": "The figure is a good specification check, but the lack of consistent nonlinearity lowers its value as a principal display.",
        "recommended_change": "Demote to supplement if a stronger robustness synthesis is available for the main text.",
    },
    {
        "display_id": "Tabla_S1_flujo_STROBE",
        "current_location": "suplementario",
        "display_type": "flow/STROBE",
        "current_role": "Tabular backup for the STROBE figure.",
        "closest_guide_analogues": "Some guide papers keep the flow only as a figure.",
        "benchmark_read": "Good supporting file; it avoids overloading the main package.",
        "score_necesidad_editorial": 2,
        "score_claridad_mensaje": 4,
        "score_adecuacion_analitica": 5,
        "score_densidad_informativa": 4,
        "score_facilidad_lectura": 4,
        "score_consistencia_caption_contenido": 5,
        "score_utilidad_para_principal": 2,
        "score_adecuacion_rpmesp": 5,
        "veredicto": "mantener",
        "recommended_location": "suplementario",
        "why": "Useful as support, but unnecessary in the main narrative because Figura_1 already covers the same function visually.",
        "recommended_change": "Keep in supplement only.",
    },
    {
        "display_id": "Tabla_S3_sensibilidad_interacciones",
        "current_location": "suplementario",
        "display_type": "sensibilidad/robustez",
        "current_role": "Collect interaction and sensitivity models outside the main package.",
        "closest_guide_analogues": "Closest to Table 3 in fpsyt and subgroup figures in fpubh and s12888.",
        "benchmark_read": "Good supplement, but still too technical to be reader-optimal as a main-table replacement.",
        "score_necesidad_editorial": 4,
        "score_claridad_mensaje": 3,
        "score_adecuacion_analitica": 5,
        "score_densidad_informativa": 3,
        "score_facilidad_lectura": 3,
        "score_consistencia_caption_contenido": 5,
        "score_utilidad_para_principal": 3,
        "score_adecuacion_rpmesp": 5,
        "veredicto": "mantener",
        "recommended_location": "suplementario",
        "why": "A useful repository of supporting models, but not as concise as a purpose-built robustness table.",
        "recommended_change": "Keep as supplement; do not use as the main robustness summary.",
    },
    {
        "display_id": "Tabla_S4_datos_figuras",
        "current_location": "suplementario",
        "display_type": "otro",
        "current_role": "Provide the source data behind the figures.",
        "closest_guide_analogues": "Analogous to supplementary detail referenced but not displayed in some guide papers.",
        "benchmark_read": "Appropriate as support material only.",
        "score_necesidad_editorial": 1,
        "score_claridad_mensaje": 2,
        "score_adecuacion_analitica": 2,
        "score_densidad_informativa": 5,
        "score_facilidad_lectura": 2,
        "score_consistencia_caption_contenido": 5,
        "score_utilidad_para_principal": 1,
        "score_adecuacion_rpmesp": 4,
        "veredicto": "mantener",
        "recommended_location": "suplementario",
        "why": "Valuable for transparency, but not designed as a narrative display.",
        "recommended_change": "Keep as supporting data only.",
    },
    {
        "display_id": "Tabla_S5_robustez_hallazgo_principal",
        "current_location": "suplementario",
        "display_type": "sensibilidad/robustez",
        "current_role": "Synthesize the principal result, second BP measure, and spline readout.",
        "closest_guide_analogues": "Most comparable to Table 3 in fpsyt and to subgroup/robustness displays in fpubh and s12888.",
        "benchmark_read": "This is the cleanest cross-check display in our package and compares well with the strongest support displays in the guide papers.",
        "score_necesidad_editorial": 5,
        "score_claridad_mensaje": 5,
        "score_adecuacion_analitica": 5,
        "score_densidad_informativa": 5,
        "score_facilidad_lectura": 5,
        "score_consistencia_caption_contenido": 5,
        "score_utilidad_para_principal": 4,
        "score_adecuacion_rpmesp": 5,
        "veredicto": "promover a principal",
        "recommended_location": "principal",
        "why": "It answers the robustness question directly and more efficiently than the current main spline figure.",
        "recommended_change": "Promote to principal if one current principal item is demoted.",
    },
]
PACKAGE_LEVEL_DIAGNOSIS: list[dict] = [
    {
        "question": "Estado general de la seleccion principal",
        "answer": "Bien balanceada, pero con una pieza principal debil.",
        "evidence": "Tabla_3, Tabla_4 y Figura_1_STROBE estan a nivel benchmark; Figura_2 es la principal mas debil porque no aporta una no linealidad sustantiva.",
    },
    {
        "question": "Elemento mas fuerte del paquete",
        "answer": "Tabla_3 y Figura_1_STROBE.",
        "evidence": "Ambas coinciden con lo que todos o casi todos los papers guia hacen bien: claridad, utilidad narrativa y adecuacion editorial.",
    },
    {
        "question": "Elemento principal por debajo del benchmark",
        "answer": "Figura_2_Spline.",
        "evidence": "Los guide papers usan curvas principales cuando revelan umbral o forma no lineal fuerte; en nuestro caso la figura opera mas como chequeo metodologico.",
    },
    {
        "question": "Elemento suplementario con mayor valor para ascender",
        "answer": "Tabla_S5_robustez_hallazgo_principal.",
        "evidence": "Resume estabilidad del hallazgo principal de forma compacta y comparable con las mejores piezas de robustez de los papers guia.",
    },
    {
        "question": "Display faltante mas util",
        "answer": "Una version visual tipo forest/robustez seria el analogo faltante mas potente a mediano plazo.",
        "evidence": "Los guide papers muestran heterogeneidad o robustez de forma sintetica; hoy nosotros lo resolvemos mejor con Tabla_S5 que con una figura dedicada.",
    },
    {
        "question": "Redundancia relevante",
        "answer": "No hay gran redundancia entre Tabla_3 y Tabla_4; la redundancia potencial esta entre Figura_2 y Tabla_S5.",
        "evidence": "Ambas hablan de robustez de la forma funcional, pero Tabla_S5 da una lectura mas util para decision editorial.",
    },
]
RUBRIC_ROWS: list[dict] = [
    {"criterio": "1", "significado": "Muy debil o poco defendible"},
    {"criterio": "2", "significado": "Debil; requiere rediseño importante para paper"},
    {"criterio": "3", "significado": "Aceptable, pero no al nivel de los mejores analogos"},
    {"criterio": "4", "significado": "Fuerte y util para manuscrito"},
    {"criterio": "5", "significado": "Muy fuerte; comparable con el mejor benchmark del set"},
]


DISPLAY_SCORE_COLUMNS = [
    "score_necesidad_editorial",
    "score_claridad_mensaje",
    "score_adecuacion_analitica",
    "score_densidad_informativa",
    "score_facilidad_lectura",
    "score_consistencia_caption_contenido",
    "score_utilidad_para_principal",
    "score_adecuacion_rpmesp",
]


def _ensure_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def _autofit_workbook(path: Path) -> None:
    wb = load_workbook(path)
    for ws in wb.worksheets:
        for column_cells in ws.columns:
            values = [str(cell.value) for cell in column_cells if cell.value is not None]
            if not values:
                continue
            width = min(max(len(value) for value in values) + 2, 54)
            ws.column_dimensions[column_cells[0].column_letter].width = width
    wb.save(path)


def _write_workbook(path: Path, sheets: dict[str, pd.DataFrame]) -> None:
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for sheet_name, frame in sheets.items():
            frame.to_excel(writer, sheet_name=sheet_name[:31], index=False)
    _autofit_workbook(path)


def _guide_summary_df() -> pd.DataFrame:
    rows = []
    for paper in GUIDE_PAPERS:
        rows.append(
            {
                "paper_id": paper["paper_id"],
                "filename": paper["filename"],
                "short_citation": paper["short_citation"],
                "title": paper["title"],
                "journal": paper["journal"],
                "year": paper["year"],
                "design": paper["design"],
                "population": paper["population"],
                "main_exposure": paper["main_exposure"],
                "main_outcome": paper["main_outcome"],
                "n_displays_total": len(paper["displays"]),
                "n_tables": sum(1 for d in paper["displays"] if d["display_kind"] == "Table"),
                "n_figures": sum(1 for d in paper["displays"] if d["display_kind"] == "Figure"),
            }
        )
    return pd.DataFrame(rows)


def _guide_inventory_df() -> pd.DataFrame:
    rows = []
    for paper in GUIDE_PAPERS:
        for display in paper["displays"]:
            row = {
                "paper_id": paper["paper_id"],
                "short_citation": paper["short_citation"],
                "filename": paper["filename"],
                "title": paper["title"],
            }
            row.update(display)
            rows.append(row)
    return pd.DataFrame(rows)


def _guide_patterns_df(inventory: pd.DataFrame) -> pd.DataFrame:
    main_only = inventory.loc[inventory["location"].isin(["principal", "support"])]
    counts = (
        main_only.groupby("display_type")
        .agg(
            n_displays=("display_id", "count"),
            n_papers=("paper_id", "nunique"),
        )
        .reset_index()
    )
    counts["papers_with_type_pct"] = (counts["n_papers"] / len(GUIDE_PAPERS) * 100).round(1)
    return counts.sort_values(["n_papers", "n_displays"], ascending=[False, False])


def _inspect_public_workbook(path: Path) -> str:
    if not path.exists():
        return "Archivo no disponible en paquete actual"
    xls = pd.ExcelFile(path)
    parts = []
    for sheet in xls.sheet_names:
        frame = pd.read_excel(path, sheet_name=sheet)
        parts.append(f"{sheet}: {len(frame)} filas")
    return "; ".join(parts)


def _our_display_matrix_df(public_dir: Path) -> pd.DataFrame:
    rows = []
    workbook_map = {
        "Tabla_1": public_dir / "Principal" / "Tabla_1.xlsx",
        "Tabla_2": public_dir / "Principal" / "Tabla_2.xlsx",
        "Tabla_3": public_dir / "Principal" / "Tabla_3.xlsx",
        "Tabla_4": public_dir / "Principal" / "Tabla_4.xlsx",
        "Tabla_S1_flujo_STROBE": public_dir / "Suplementario" / "Tabla_S1_flujo_STROBE.xlsx",
        "Tabla_S3_sensibilidad_interacciones": public_dir / "Suplementario" / "Tabla_S3_sensibilidad_interacciones.xlsx",
        "Tabla_S4_datos_figuras": public_dir / "Suplementario" / "Tabla_S4_datos_figuras.xlsx",
        "Tabla_S5_robustez_hallazgo_principal": public_dir / "Suplementario" / "Tabla_S5_robustez_hallazgo_principal.xlsx",
    }
    for display in OUR_DISPLAYS:
        row = dict(display)
        if display["display_id"] in workbook_map:
            row["source_snapshot"] = _inspect_public_workbook(workbook_map[display["display_id"]])
        else:
            row["source_snapshot"] = "Principal figure asset"
        row["total_score"] = sum(int(display[col]) for col in DISPLAY_SCORE_COLUMNS)
        rows.append(row)
    return pd.DataFrame(rows)


def _diagnosis_df(matrix: pd.DataFrame) -> pd.DataFrame:
    return matrix[
        [
            "display_id",
            "current_location",
            "display_type",
            "total_score",
            "veredicto",
            "recommended_location",
            "why",
            "recommended_change",
            "closest_guide_analogues",
        ]
    ].rename(
        columns={
            "display_id": "display",
            "current_location": "ubicacion_actual",
            "display_type": "tipo_display",
            "total_score": "puntaje_total",
            "recommended_location": "ubicacion_recomendada",
            "why": "diagnostico",
            "recommended_change": "accion_recomendada",
            "closest_guide_analogues": "analogo_guia",
        }
    )


def _memo_text(guide_summary: pd.DataFrame, patterns: pd.DataFrame, matrix: pd.DataFrame) -> str:
    total_guides = len(guide_summary)
    principal_displays = matrix.loc[matrix["current_location"] == "principal"]
    weakest = principal_displays.sort_values("total_score", ascending=True).iloc[0]["display_id"]
    promote = matrix.loc[matrix["veredicto"] == "promover a principal"].iloc[0]["display_id"]
    pattern_lines = "\n".join(
        f"- {row.display_type}: presente en {int(row.n_papers)} de {total_guides} papers guia"
        for row in patterns.itertuples()
    )
    display_lines = "\n".join(
        f"- {row.display_id}: {row.veredicto} ({row.recommended_location})"
        for row in matrix.itertuples()
    )
    return (
        "# Memo de auditoria comparativa de tablas y figuras\n\n"
        "## Conclusion ejecutiva\n"
        "Nuestra seleccion actual de displays es fuerte en su columna vertebral editorial: "
        "Tabla_3, Tabla_4 y Figura_1_STROBE estan a nivel comparable con los mejores analogos de los papers guia. "
        f"El punto mas debil del paquete principal es {weakest}, porque aporta mas como verificacion metodologica que como hallazgo central.\n\n"
        f"El cambio con mayor rendimiento editorial seria promover {promote} y considerar el traslado de Figura_2_Spline a material suplementario. "
        "Eso alinearia mejor el paquete con RPMESP y con el patron de los papers guia: las curvas principales se reservan para relaciones no lineales o umbrales que cambian la lectura del hallazgo.\n\n"
        "## Lo que hacen los papers guia y que nos sirve de benchmark\n"
        f"{pattern_lines}\n\n"
        "## Diagnostico de nuestra eleccion de displays\n"
        "- La seleccion principal no esta incompleta, pero si esta ligeramente sesgada hacia la completitud metodologica en lugar de la maxima utilidad narrativa.\n"
        "- Tabla_1 esta bien elegida, aunque su forma actual es mas densa que la de los benchmarks y convendria depurarla si se maqueta para manuscrito.\n"
        "- Tabla_2 esta bien elegida por el plan, pero su ejecucion actual esta por debajo del benchmark porque deja ver terminos internos de implementacion.\n"
        "- Tabla_3 y Tabla_4 estan bien elegidas y bien ubicadas.\n"
        "- Figura_1_STROBE esta bien elegida y bien resuelta.\n"
        "- Figura_2_Spline es la pieza mas discutible en principal porque no cambia la interpretacion del resultado central.\n"
        "- Tabla_S5_robustez_hallazgo_principal es la mejor candidata a ascenso editorial.\n\n"
        "## Decisiones concretas sugeridas\n"
        "- Mantener: Tabla_3, Tabla_4, Figura_1_STROBE.\n"
        "- Refinar: Tabla_1, Tabla_2.\n"
        "- Mantener en suplementario: Tabla_S1_flujo_STROBE, Tabla_S3_sensibilidad_interacciones, Tabla_S4_datos_figuras.\n"
        "- Promover a principal: Tabla_S5_robustez_hallazgo_principal.\n"
        "- Mover a suplementario: Figura_2_Spline.\n\n"
        "## Que estilo conviene adoptar y que no\n"
        "Si conviene adoptar:\n"
        "- Tablas principales con mensaje unico y rapido de leer.\n"
        "- Displays de robustez sinteticamente orientados a decision.\n"
        "- Curvas solo cuando muestran forma funcional relevante.\n\n"
        "No conviene adoptar:\n"
        "- Curvas principales que solo prueban ausencia de no linealidad sin cambiar la interpretacion.\n"
        "- Tablas tecnicas con nomenclatura de implementacion expuesta al lector.\n"
        "- Basales demasiado fragmentadas en la version final del manuscrito.\n\n"
        "## Snapshot por display\n"
        f"{display_lines}\n\n"
        "## Respuesta corta a la pregunta central\n"
        "Si comparamos nuestro paquete con estos cuatro trabajos, elegimos bien la mayoria de nuestras tablas y figuras, pero no del todo la jerarquia final. "
        "La arquitectura seria mas fuerte si la robustez sintetica sube al cuerpo principal y la spline baja a suplementario.\n"
    )


def build_audit(public_dir: Path | None = None) -> Path:
    target_public_dir = public_dir or PUBLIC_DIR
    audit_dir = target_public_dir / "Auditoria_Guias"
    _ensure_dir(audit_dir)

    guide_summary = _guide_summary_df()
    guide_inventory = _guide_inventory_df()
    guide_patterns = _guide_patterns_df(guide_inventory)
    matrix = _our_display_matrix_df(target_public_dir)
    diagnosis = _diagnosis_df(matrix)
    package_level = pd.DataFrame(PACKAGE_LEVEL_DIAGNOSIS)
    rubric = pd.DataFrame(RUBRIC_ROWS)

    per_paper_sheets = {
        paper["paper_id"][:31]: guide_inventory.loc[guide_inventory["paper_id"] == paper["paper_id"]].drop(
            columns=["title", "filename"]
        )
        for paper in GUIDE_PAPERS
    }

    _write_workbook(
        audit_dir / "01_fichas_resumen_papers_guia.xlsx",
        {
            "Resumen_papers": guide_summary,
            "Inventario_displays": guide_inventory,
            "Patrones_cruzados": guide_patterns,
            **per_paper_sheets,
        },
    )
    _write_workbook(
        audit_dir / "02_matriz_comparativa_displays.xlsx",
        {
            "Matriz": matrix,
            "Anclajes_1a5": rubric,
            "Patrones_guias": guide_patterns,
        },
    )
    _write_workbook(
        audit_dir / "03_diagnostico_displays_propios.xlsx",
        {
            "Diagnostico": diagnosis,
            "Arquitectura_narrativa": package_level,
        },
    )
    (audit_dir / "04_memo_recomendaciones_editoriales.md").write_text(
        _memo_text(guide_summary, guide_patterns, matrix),
        encoding="utf-8-sig",
    )
    print(f"Auditoria de guias lista en: {audit_dir}")
    return audit_dir


def main() -> None:
    build_audit()


if __name__ == "__main__":
    main()
