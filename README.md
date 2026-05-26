[bibliografia_evaluacion_economica.html](https://github.com/user-attachments/files/28277180/bibliografia_evaluacion_economica.html)[Uploadin<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>PREDIMIGRAINE · Bibliografía evaluación económica</title>
<link href="https://fonts.googleapis.com/css2?family=Source+Serif+4:opsz,wght@8..60,400;8..60,600;8..60,700&family=JetBrains+Mono:wght@400;600&display=swap" rel="stylesheet">
<script src="https://unpkg.com/react@18/umd/react.production.min.js" crossorigin></script>
<script src="https://unpkg.com/react-dom@18/umd/react-dom.production.min.js" crossorigin></script>
<script src="https://unpkg.com/@babel/standalone/babel.min.js"></script>
<script src="https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js"></script>
<style>
  * { box-sizing: border-box; }
  body { margin: 0; background: #faf8f3; font-family: 'Source Serif 4', Georgia, serif; }
  .filtro-btn {
    transition: all 0.15s ease; cursor: pointer;
    font-family: 'JetBrains Mono', monospace; font-size: 11px;
    letter-spacing: 0.05em; text-transform: uppercase;
    border: 1px solid #1a1a1a; background: transparent; padding: 6px 12px; color: #1a1a1a;
  }
  .filtro-btn:hover { background: #1a1a1a; color: #faf8f3; }
  .filtro-btn.activo { background: #1a1a1a; color: #faf8f3; }
  .export-btn {
    transition: all 0.15s ease; cursor: pointer;
    font-family: 'JetBrains Mono', monospace; font-size: 11px;
    letter-spacing: 0.08em; text-transform: uppercase;
    border: 1px solid #2d6a4f; background: #2d6a4f; color: #fff;
    padding: 10px 18px; font-weight: 600;
  }
  .export-btn:hover { background: #1a4a37; border-color: #1a4a37; }
  .export-btn.alt { background: #fff; color: #2d6a4f; }
  .export-btn.alt:hover { background: #2d6a4f; color: #fff; }
  .ref-card { background: #fff; padding: 20px; border: 1px solid #ddd; margin-bottom: 12px; }
  .ref-card:hover { background: #faf8f3; border-color: #1a1a1a; }
  .tag {
    display: inline-block; font-family: 'JetBrains Mono', monospace;
    font-size: 10px; padding: 2px 8px; margin-right: 4px; margin-bottom: 4px;
    letter-spacing: 0.05em; text-transform: uppercase;
  }
  .tag.cuna { background: #fef3c7; color: #7c5e10; }
  .tag.spain { background: #d8f3dc; color: #2d6a4f; }
  .tag.key { background: #1a1a1a; color: #faf8f3; }
  .tag.review { background: #f0ecf5; color: #5c4677; }
  .tag.project { background: #fee2e2; color: #9b2226; }
  a { color: #2d6a4f; text-decoration: none; border-bottom: 1px solid #2d6a4f; }
  a:hover { color: #1a4a37; }
  @media print { .no-print { display: none !important; } body { background: #fff; } }
</style>
</head>
<body>
<div id="root"></div>

<script type="text/babel">
const { useState, useMemo } = React;

const refs = [
  // ============ BLOQUE 1: CARGA ECONÓMICA Y EPIDEMIOLOGÍA ============
  { id: 1, bloque: 'Carga económica y epidemiología', clave: true,
    autores: 'Fernández-Ferro J, Ordás-Bandera C, Rejas-Gutiérrez J, et al.',
    año: 2024, titulo: 'The economic burden of migraine: a nationwide cost-of-illness approach from the year 2020 European Health Survey in Spain',
    revista: 'Neurología', vol: '40(2025):533-547', doi: '10.1016/j.nrl.2023.05.006',
    url: 'https://doi.org/10.1016/j.nrleng.2025.06.002',
    tags: ['Clave', 'España', 'Cost-of-illness'],
    porque: 'Es la referencia central de carga económica en España: 5.862 € PPPY en pacientes con migraña vs 2.981 € sin migraña; exceso 2.881 €/año (1.928 € laboral + 836 € sanitario SNS). Proyección nacional 10.394-14.230 millones €/año. Indispensable para el cálculo de impacto presupuestario y como benchmark del análisis económico del proyecto.',
    yaEnProyecto: true },
  { id: 2, bloque: 'Carga económica y epidemiología', clave: false,
    autores: 'Caronna E, Pozo-Rosich P, et al.',
    año: 2024, titulo: 'Societal and economic burden of migraine in Spain: results from the 2020 National Health and Wellness Survey',
    revista: 'J Headache Pain', vol: '25(1):26',
    url: 'https://pubmed.ncbi.nlm.nih.gov/38486155/',
    tags: ['España', 'HCRU', 'WPAI'],
    porque: 'Estudio caso-control en España (NHWS 2020) que estratifica por días de cefalea (1-3, 4-7, 8-14, ≥15). Cuantifica HCRU: hospitalizaciones 17,0% vs 8,3% (RR 2,0); urgencias 51,4% vs 31,2% (RR 1,6). Permite estratificar análisis económico por frecuencia de migraña.',
    yaEnProyecto: false },
  { id: 3, bloque: 'Carga económica y epidemiología', clave: false,
    autores: 'Vo P, Fang J, Bilitou A, et al.',
    año: 2025, titulo: 'Prevalence, disability, and economic impact of migraine in Spain: a nationwide population-based study',
    revista: 'J Headache Pain', vol: '26:54',
    url: 'https://thejournalofheadacheandpain.biomedcentral.com/articles/10.1186/s10194-025-02069-1',
    tags: ['España', 'HCRU', 'Reciente'],
    porque: 'Actualización con NHWS 2022 (7.002 respondedores, 930 con migraña activa). Datos por regiones de España, HCRU 6 meses (primaria, neurología, urgencias, hospitalizaciones), nivel educativo, ingresos y empleo. Lo más reciente para impacto presupuestario.',
    yaEnProyecto: false },
  { id: 4, bloque: 'Carga económica y epidemiología', clave: false,
    autores: 'Steiner TJ, Stovner LJ, Jensen R, et al.',
    año: 2018, titulo: 'Global, regional, and national burden of migraine and tension-type headache, 1990-2016: a systematic analysis for the Global Burden of Disease Study 2016',
    revista: 'Lancet Neurology', vol: '17(11):954-976', doi: '10.1016/S1474-4422(18)30322-3',
    url: 'https://www.ncbi.nlm.nih.gov/pmc/articles/PMC6191530/',
    tags: ['Clave', 'GBD', 'YLD'],
    porque: 'Referencia GBD 2016: migraña causa 45,1 millones YLDs globales; mujeres 15-49 años con 20,3 millones YLDs (11,2% del total). Justifica el foco del proyecto en mujeres en edad laboral.',
    yaEnProyecto: false },
  { id: 5, bloque: 'Carga económica y epidemiología', clave: false,
    autores: 'Zhang B, Xu R, Zhao Y, et al.',
    año: 2025, titulo: 'The Global Burden of Migraine: A 30-Year Trend Review and Future Projections by Age, Sex, Country, and Region',
    revista: 'Pain and Therapy', vol: '14:107-127',
    url: 'https://www.ncbi.nlm.nih.gov/pmc/articles/PMC11751287/',
    tags: ['GBD', 'Proyecciones'],
    porque: 'GBD 2021 con proyecciones hasta 2050 (ARIMA). Prevalencia mundial subió 58,15% (de 732M a 1,16 B). DALYs +58,27%. Útil para contextualizar tendencia creciente y justificar inversión en herramientas predictivas.',
    yaEnProyecto: false },
  { id: 6, bloque: 'Carga económica y epidemiología', clave: false,
    autores: 'Sakai F, Igarashi H, Yokoyama M, et al.',
    año: 2021, titulo: 'Disability, quality of life, productivity impairment and employer costs of migraine in the workplace',
    revista: 'J Headache Pain', vol: '22:29',
    url: 'https://link.springer.com/article/10.1186/s10194-021-01243-5',
    tags: ['Productividad', 'WPAI', 'MIDAS'],
    porque: 'Comparativa de pérdida económica usando MIDAS vs WPAI: presentismo es ~6x más costoso que absentismo. WPAI capta 2.217 $/año/persona vs MIDAS 375,4 $/año/persona. Demuestra por qué WPAI es metodológicamente superior para evaluar coste indirecto.',
    yaEnProyecto: false },
  { id: 7, bloque: 'Carga económica y epidemiología', clave: false,
    autores: 'Papakonstantinou D, Tomos C',
    año: 2022, titulo: 'Workplace productivity loss as a result of absenteeism and presenteeism in chronic and episodic migraine: a scoping review',
    revista: 'Int J Workplace Health Manag', vol: '15(5):613-628',
    url: 'https://www.emerald.com/insight/content/doi/10.1108/ijwhm-05-2021-0123/full/html',
    tags: ['Review', 'Productividad'],
    porque: 'Revisión 2010-2020 que confirma sistemáticamente que presentismo > absentismo en migraña, y que ambos escalan con la frecuencia. Útil como referencia metodológica al construir la sección de productividad.',
    yaEnProyecto: false },
  { id: 8, bloque: 'Carga económica y epidemiología', clave: false,
    autores: 'Irimia P, Garrido-Cumbrera M, Santos-Lasaosa S, et al.',
    año: 2020, titulo: 'Estimating the savings of a migraine free life: results from the Spanish atlas',
    revista: 'Eur J Neurol', vol: '27(12):2616-2624', doi: '10.1111/ene.14431',
    url: 'https://doi.org/10.1111/ene.14431',
    tags: ['España', 'Atlas SEN'],
    porque: 'Atlas SEN sobre el ahorro potencial de eliminar migraña en España. Referencia obligatoria para impacto socioeconómico nacional.',
    yaEnProyecto: false },

  // ============ BLOQUE 2: COSTE-EFECTIVIDAD ANTI-CGRP ============
  { id: 9, bloque: 'Coste-efectividad anti-CGRP', clave: true,
    autores: 'Lazaro-Hernandez C, Caronna E, Rosell-Mirmi J, Gallardo VJ, Alpuente A, Torres-Ferrus M, Pozo-Rosich P',
    año: 2024, titulo: 'Early and annual projected savings from anti-CGRP monoclonal antibodies in migraine prevention: a cost-benefit analysis in the working-age population',
    revista: 'J Headache Pain', vol: '25:21', doi: '10.1186/s10194-024-01727-0',
    url: 'https://doi.org/10.1186/s10194-024-01727-0',
    tags: ['Clave', 'España', 'WPAI', 'Plantilla metodológica'],
    porque: 'Es la PLANTILLA METODOLÓGICA PERFECTA para el análisis CBA del proyecto. Cohorte n=148 trabajadores activos en Vall d\'Hebron. WPAI + salario medio municipal → 159,8 €/paciente/3m, 639,2 €/año. 93% de los ahorros vienen de productividad. Esta es la referencia central para replicar metodología en PREDICGRP.',
    yaEnProyecto: true },
  { id: 10, bloque: 'Coste-efectividad anti-CGRP', clave: false,
    autores: 'Mahon R, Lang A, Vo P, et al.',
    año: 2021, titulo: 'Cost-Effectiveness of Erenumab for the Preventive Treatment of Migraine in Patients with Prior Treatment Failures in Sweden',
    revista: 'Pharmacoeconomics', vol: '39:357-372',
    url: 'https://pubmed.ncbi.nlm.nih.gov/33502733/',
    tags: ['Coste-utilidad', 'Europa', 'Erenumab'],
    porque: 'Modelo Markov con horizonte temporal. Útil como referencia metodológica para construir CUA con perspectiva SNS europeo.',
    yaEnProyecto: false },
  { id: 11, bloque: 'Coste-efectividad anti-CGRP', clave: false,
    autores: 'Sussman M, Benner J, Neumann P, Menzin J',
    año: 2018, titulo: 'Cost-effectiveness analysis of erenumab for the preventive treatment of episodic and chronic migraine: results from the US societal and payer perspectives',
    revista: 'Cephalalgia', vol: '38(10):1644-1657',
    url: 'https://pubmed.ncbi.nlm.nih.gov/30086653/',
    tags: ['Coste-utilidad', 'USA', 'Erenumab'],
    porque: 'Primer análisis coste-efectividad publicado de erenumab. Doble perspectiva (societal y pagador). Marco metodológico replicable.',
    yaEnProyecto: false },
  { id: 12, bloque: 'Coste-efectividad anti-CGRP', clave: false,
    autores: 'Lipton RB, Brennan A, Palmer S, et al.',
    año: 2018, titulo: 'Estimating the clinical effectiveness and value-based price range of erenumab for the prevention of migraine in patients with prior treatment failures: a US societal perspective',
    revista: 'J Med Econ', vol: '21(7):666-675',
    url: 'https://pubmed.ncbi.nlm.nih.gov/29667457/',
    tags: ['Pricing basado en valor', 'USA'],
    porque: 'Análisis value-based price. Útil para discutir si los precios actuales SNS están alineados con valor clínico.',
    yaEnProyecto: false },
  { id: 13, bloque: 'Coste-efectividad anti-CGRP', clave: false,
    autores: 'Atlas R, Khalaf K, Marini C, et al.',
    año: 2022, titulo: 'A systematic review of economic evaluations of pharmacological treatments for adults with chronic migraine',
    revista: 'J Headache Pain', vol: '23:103',
    url: 'https://link.springer.com/article/10.1186/s10194-022-01492-y',
    tags: ['Review', 'Clave', 'Coste-utilidad'],
    porque: 'Revisión sistemática de 16 estudios económicos en migraña crónica. Botox vs placebo ICER 15.028-16.598 £. Anti-CGRP vs Botox ICER 59.712-182.128 £ (por encima del WTP estándar, pero coste-efectivos tras fallo de Botox). Mapeo imprescindible para situar el análisis de PREDIMIGRAINE en contexto internacional.',
    yaEnProyecto: false },
  { id: 14, bloque: 'Coste-efectividad anti-CGRP', clave: false,
    autores: 'Darbà J, Marsà A',
    año: 2020, titulo: 'Cost of fremanezumab, erenumab, galcanezumab and onabotulinumtoxinA associated adverse events, for migraine prophylaxis in Spain',
    revista: 'J Med Econ', vol: '23(9):1003-1009',
    url: 'https://pubmed.ncbi.nlm.nih.gov/32484365/',
    tags: ['España', 'Costes EA'],
    porque: 'Coste de eventos adversos asociados a los 4 tratamientos en España. Fremanezumab genera ahorros vs erenumab (-469 €), galcanezumab (-268 €) y BoNT-A (-1.100/-1.295 €). Útil para complementar SAFE-CGRP con perspectiva económica de seguridad.',
    yaEnProyecto: false },
  { id: 15, bloque: 'Coste-efectividad anti-CGRP', clave: false,
    autores: 'Iannone LF, De Cesaris F, Geppetti P, et al.',
    año: 2022, titulo: 'It is time anti-CGRP monoclonal antibodies be considered first-line prophylaxis for migraine',
    revista: 'Headache', vol: '62(8):1090-1093',
    url: 'https://www.ncbi.nlm.nih.gov/pmc/articles/PMC9491437/',
    tags: ['Posicionamiento', 'Primera línea'],
    porque: 'Editorial-revisión que defiende anti-CGRP como primera línea apoyándose en argumentos de coste-efectividad. Útil para la discusión y para argumentar racionalización del uso del fármaco (clave en PREDICGRP).',
    yaEnProyecto: false },
  { id: 16, bloque: 'Coste-efectividad anti-CGRP', clave: false,
    autores: 'Caronna E, Gallardo VJ, Egeo G, et al.',
    año: 2024, titulo: 'Redefining migraine prevention: early treatment with anti-CGRP monoclonal antibodies enhances response in the real world',
    revista: 'J Neurol Neurosurg Psychiatry', vol: '95:927-937', doi: '10.1136/jnnp-2023-333295',
    url: 'https://doi.org/10.1136/jnnp-2023-333295',
    tags: ['España', 'Real-world', 'Vall d\'Hebron'],
    porque: 'Estudio multicéntrico real-world con definiciones de buena (≥50%) y excelente (≥75%) respuesta. Argumento clínico-económico para tratamiento precoz, alineado con la hipótesis de PREDICGRP de evitar la cronificación.',
    yaEnProyecto: false },
  { id: 17, bloque: 'Coste-efectividad anti-CGRP', clave: false,
    autores: 'Pozo-Rosich P, Dolezil D, Paemeleire K, et al.',
    año: 2024, titulo: 'Early use of erenumab vs nonspecific oral migraine preventives: the APPRAISE randomized clinical trial',
    revista: 'JAMA Neurology', vol: '81:461-470',
    url: 'https://jamanetwork.com/journals/jamaneurology',
    tags: ['Ensayo clínico', 'Primera línea'],
    porque: 'Ensayo APPRAISE: anti-CGRP vs preventivos no específicos en pacientes nuevos. Demuestra superioridad clínica del uso temprano. Apoya argumento de PREDICGRP: identificar respondedores ahorra dinero al evitar fallos terapéuticos.',
    yaEnProyecto: false },
  { id: 18, bloque: 'Coste-efectividad anti-CGRP', clave: false,
    autores: 'Grazzi L, Giossi R, Montisano DA, et al.',
    año: 2024, titulo: 'Real-world effectiveness of Anti-CGRP monoclonal antibodies compared to OnabotulinumtoxinA (RAMO) in chronic migraine: a retrospective, observational, multicenter, cohort study',
    revista: 'J Headache Pain', vol: '25:14',
    url: 'https://www.ncbi.nlm.nih.gov/pmc/articles/PMC10836018/',
    tags: ['Real-world', 'Comparativa'],
    porque: 'Anti-CGRP vs BoNT-A en práctica real. Apoya selección del comparador correcto en el análisis económico (BoNT-A es el comparador habitual para MC).',
    yaEnProyecto: false },

  // ============ BLOQUE 3: UTILIDADES Y QALY ============
  { id: 19, bloque: 'Utilidades y QALY', clave: true,
    autores: 'Gillard PJ, Devine B, Varon SF, Liu L, Sullivan SD',
    año: 2012, titulo: 'Mapping from disease-specific measures to health-state utility values in individuals with migraine',
    revista: 'Value in Health', vol: '15(3):485-494',
    url: 'https://pubmed.ncbi.nlm.nih.gov/22583459/',
    tags: ['Clave', 'Mapeo', 'HIT-6'],
    porque: 'CLAVE para el proyecto. Algoritmo publicado de mapeo HIT-6 → EQ-5D. R² 0,22 episódica, 0,36 crónica. Permite estimar utilidades en PREDIMIGRAINE a partir de los HIT-6 ya recogidos, salvando el problema de no haber recogido EQ-5D.',
    yaEnProyecto: false },
  { id: 20, bloque: 'Utilidades y QALY', clave: false,
    autores: 'Stafford MR, Hareendran A, Ng-Mak DS, Insinga RP, Xu R, Stull DE',
    año: 2012, titulo: 'EQ-5D™-derived utility values for different levels of migraine severity from a UK sample of migraineurs',
    revista: 'Health Qual Life Outcomes', vol: '10:65',
    url: 'https://www.ncbi.nlm.nih.gov/pmc/articles/PMC3407525/',
    tags: ['Clave', 'EQ-5D', 'Utilidades base'],
    porque: 'Utilidades EQ-5D por nivel de severidad de migraña (sin migraña, leve, moderada, grave). La migraña severa es considerada un estado peor que la muerte. Imprescindible para construir el modelo de utilidad ponderada por días de migraña.',
    yaEnProyecto: false },
  { id: 21, bloque: 'Utilidades y QALY', clave: false,
    autores: 'Ahadi MS, Vahidpour N, Ghazi Sherbaf F, et al.',
    año: 2021, titulo: 'Assessment of Utility in Migraine: Mapping the Migraine-Specific Questionnaire to the EQ-5D-5L',
    revista: 'Value in Health Regional Issues', vol: '24:24-32',
    url: 'https://www.valuehealthregionalissues.com/article/S2212-1099(21)00004-2/fulltext',
    tags: ['Mapeo', 'MSQ', 'Alternativa'],
    porque: 'Mapeo MSQ → EQ-5D-5L. R² 0,26 episódica, 0,38 crónica. Alternativa al mapeo de HIT-6 si en el futuro se añade MSQ al CRD.',
    yaEnProyecto: false },
  { id: 22, bloque: 'Utilidades y QALY', clave: false,
    autores: 'Sapra S, Bushmakin AG, Cappelleri JC, et al.',
    año: 2021, titulo: 'EQ-5D Health-State Utility Values for Patients with Migraine: Estimates from Erenumab Clinical Trials',
    revista: 'Headache', vol: '61(4):650-662',
    url: 'https://www.academia.edu/23374731/EQ_5D_derived_utility_values_for_different_levels_of_migraine_severity_from_a_UK_sample_of_migraineurs',
    tags: ['EQ-5D', 'Erenumab', 'Utilidades por MMD'],
    porque: 'Utilidades EQ-5D modeladas como función de los días de migraña al mes (MMD). Cuatro modelos comparados. Permite construir directamente la utilidad ponderada por días, que es el método que la Solicitud de PREDIMIGRAINE ya endosa.',
    yaEnProyecto: false },
  { id: 23, bloque: 'Utilidades y QALY', clave: false,
    autores: 'Łukaszuk M, Kalinowska E, Domitrz I',
    año: 2022, titulo: 'Health-Related Quality of Life in Migraine: EQ-5D-5L-Based Study in Routine Clinical Practice',
    revista: 'J Clin Med', vol: '11(23):6925',
    url: 'https://www.mdpi.com/2077-0383/11/23/6925',
    tags: ['EQ-5D', 'Práctica clínica'],
    porque: 'EQ-5D-5L en migraña episódica y crónica vs controles. Utilidades de referencia poblacional para análisis de sensibilidad.',
    yaEnProyecto: false },
  { id: 24, bloque: 'Utilidades y QALY', clave: false,
    autores: 'Sacristán JA, Oliva J, Llano JD, et al.',
    año: 2020, titulo: '¿Qué es una intervención sanitaria eficiente en España en 2020?',
    revista: 'Gaceta Sanitaria', vol: '34(2):189-193',
    url: 'https://diariofarma.com/2020/03/11/expertos-proponen-un-rango-para-determinar-el-coste-efectividad-en-salud-25-000-60-000-e-avac',
    tags: ['Clave', 'España', 'Umbral'],
    porque: 'Propuesta de umbral de coste-efectividad en España: 25.000 €/AVAC (financiación normal) hasta 60.000 €/AVAC (máximo). IMPRESCINDIBLE como referencia para interpretar los ICER del proyecto en contexto español.',
    yaEnProyecto: false },
  { id: 25, bloque: 'Utilidades y QALY', clave: false,
    autores: 'Vallejo-Torres L, García-Lorenzo B, Castilla I, et al.',
    año: 2018, titulo: 'On the Estimation of the Cost-Effectiveness Threshold: Why, What, How?',
    revista: 'Value in Health', vol: '21(5):559-566',
    url: 'https://www.valueinhealthjournal.com/article/S1098-3015(17)33599-7/fulltext',
    tags: ['España', 'Umbral'],
    porque: 'Estima umbral coste-efectividad en España en 20.000-25.000 €/AVAC. Junto con Sacristán 2020 son las dos referencias estándar para el WTP español.',
    yaEnProyecto: false },

  // ============ BLOQUE 4: MEDICINA DE PRECISIÓN Y BIOMARCADORES ============
  { id: 26, bloque: 'Medicina de precisión y biomarcadores', clave: true,
    autores: 'Ihara K, Casillo F, Dahshan A, ..., Pozo-Rosich P, Schwedt TJ',
    año: 2024, titulo: 'Are we closer to achieving precision medicine for migraine treatment? A narrative review',
    revista: 'Cephalalgia', vol: '44(11):03331024241281518',
    url: 'https://journals.sagepub.com/doi/10.1177/03331024241281518',
    tags: ['Clave', 'Review', 'Precisión'],
    porque: 'Revisión exhaustiva del estado de la medicina de precisión en migraña. Marco conceptual perfecto para la introducción de PREDIMIGRAINE. Cubre genómica, neuroimagen, proteómica.',
    yaEnProyecto: false },
  { id: 27, bloque: 'Medicina de precisión y biomarcadores', clave: false,
    autores: 'Giacon A, Iannone LF, Geppetti P, Terrazzino S',
    año: 2026, titulo: 'Predicting Pharmacological Treatment Response in Migraine Using AI/ML: A Scoping Review of the Evidence and Future Directions',
    revista: 'Pharmacotherapy', vol: '46(2):e70085',
    url: 'https://accpjournals.onlinelibrary.wiley.com/doi/10.1002/phar.70085',
    tags: ['Review', 'Machine Learning', 'Reciente'],
    porque: 'Revisión más reciente (2026) de modelos AI/ML para predecir respuesta a tratamiento en migraña. 12 estudios elegibles. Identifica limitaciones metodológicas y oportunidades. Imprescindible para situar PREDICGRP.',
    yaEnProyecto: false },
  { id: 28, bloque: 'Medicina de precisión y biomarcadores', clave: false,
    autores: 'Chen H, Gao H, Zhang Y, et al.',
    año: 2025, titulo: 'Prediction models for treatment response in migraine: a systematic review and meta-analysis',
    revista: 'J Headache Pain', vol: '26:24',
    url: 'https://link.springer.com/article/10.1186/s10194-025-01972-x',
    tags: ['Clave', 'Review', 'Meta-análisis'],
    porque: 'Revisión sistemática + meta-análisis de modelos predictivos. Aplica PROBAST + TRIPOD+AI. Marco metodológico para reportar el modelo predictivo de PREDICGRP correctamente.',
    yaEnProyecto: false },
  { id: 29, bloque: 'Medicina de precisión y biomarcadores', clave: false,
    autores: 'Chiang CC, Croope D, Cohen F, et al.',
    año: 2024, titulo: 'Advancing toward precision migraine treatment: Predicting responses to preventive medications with machine learning models based on patient and migraine features',
    revista: 'Headache', vol: '64:1094-1108',
    url: 'https://headachejournal.onlinelibrary.wiley.com/doi/10.1111/head.14806',
    tags: ['Machine Learning', 'Mayo Clinic'],
    porque: 'Modelo ML para predecir respuesta a 7 tipos de preventivos. El modelo CGRP-MAbs fue el más exitoso. Apoya hipótesis de PREDICGRP de que las características clínicas detalladas predicen respuesta.',
    yaEnProyecto: false },
  { id: 30, bloque: 'Medicina de precisión y biomarcadores', clave: false,
    autores: 'Greco R, De Icco R, Demartini C, et al.',
    año: 2025, titulo: 'Plasma CGRP but not PACAP-38 concentrations are associated with response to anti-CGRP monoclonal antibodies in migraine',
    revista: 'J Headache Pain (en prensa)',
    url: 'https://www.researchgate.net/publication/326379454_Feasibility_of_serum_CGRP_measurement_as_a_biomarker_of_chronic_migraine_a_critical_reappraisal',
    tags: ['CGRP plasma', 'Predictor'],
    porque: 'CGRP plasma basal predice respuesta a anti-CGRP. Niveles altos → peor respuesta. Justifica directamente el uso de CGRP basal en PREDICGRP (medido en Valdecilla en 173 pacientes).',
    yaEnProyecto: false },
  { id: 31, bloque: 'Medicina de precisión y biomarcadores', clave: false,
    autores: 'Cernuda-Morollón E, Larrosa D, Ramón C, et al.',
    año: 2013, titulo: 'Interictal increase of CGRP levels in peripheral blood as a biomarker for chronic migraine',
    revista: 'Neurology', vol: '81(14):1191-1196',
    url: 'https://www.neurology.org/doi/10.1212/WNL.0b013e3182a6cb72',
    tags: ['España', 'CGRP', 'Cronicidad'],
    porque: 'Estudio español que demuestra que CGRP plasmático interictal está elevado en MC vs ME (74,9 vs 46,4 pg/mL). Referencia metodológica para el análisis CGRP de PREDIMIGRAINE.',
    yaEnProyecto: false },
  { id: 32, bloque: 'Medicina de precisión y biomarcadores', clave: false,
    autores: 'Lloria-Albert L, Polo-Lerma C, Cuesta-Núñez JA, et al.',
    año: 2025, titulo: 'Plasma with Added Protease Inhibitors Improves Alpha- and Beta-CGRP Measurement Compared to Serum: Towards a Reliable Biomarker for Chronic Migraine',
    revista: 'Int J Mol Sci', vol: '26(20):10174',
    url: 'https://www.ncbi.nlm.nih.gov/pmc/articles/PMC12564834/',
    tags: ['Metodología CGRP', 'Reciente'],
    porque: 'Cuestiones metodológicas del análisis CGRP α/β: plasma + inhibidores de proteasas > suero. α-CGRP baja durante tratamiento, β-CGRP estable. Importante para la validez de las mediciones de Valdecilla en PREDIMIGRAINE.',
    yaEnProyecto: false },
  { id: 33, bloque: 'Medicina de precisión y biomarcadores', clave: false,
    autores: 'Hautakangas H, Winsvold BS, Ruotsalainen SE, et al.',
    año: 2022, titulo: 'Genome-wide analysis of 102,084 migraine cases identifies 123 risk loci and subtype-specific risk alleles',
    revista: 'Nature Genetics', vol: '54:152-160',
    url: 'https://www.nature.com/articles/s41588-021-00990-0',
    tags: ['Clave', 'GWAS', 'Genética'],
    porque: 'GWAS más grande hasta la fecha: 102.084 casos, 123 loci, 86 novedosos. Incluye CALCA/CALCB (target del anti-CGRP) y HTR1F (target del lasmiditán). Marco genético para el análisis GWAS de PREDIMIGRAINE en Vall d\'Hebron.',
    yaEnProyecto: false },
  { id: 34, bloque: 'Medicina de precisión y biomarcadores', clave: false,
    autores: 'Sutherland HG, Albury CL, Griffiths LR',
    año: 2024, titulo: 'Genetics of migraine: complexity, implications, and potential clinical applications',
    revista: 'Lancet Neurology', vol: '23(4):429-446',
    url: 'https://www.sciencedirect.com/science/article/abs/pii/S1474442224000267',
    tags: ['Review', 'Genética', 'Farmacogenómica'],
    porque: 'Revisión Lancet más reciente sobre genética de migraña con sección de farmacogenómica. Indispensable para discutir el análisis genético en PREDIMIGRAINE.',
    yaEnProyecto: false },
  { id: 35, bloque: 'Medicina de precisión y biomarcadores', clave: false,
    autores: 'Mavridis T, Mitsikostas DD',
    año: 2024, titulo: 'The impact of genetic factors on the response to migraine therapy',
    revista: 'Expert Rev Neurother', vol: '24:577-586',
    url: 'https://pubmed.ncbi.nlm.nih.gov/38856190/',
    tags: ['Review', 'Farmacogenómica'],
    porque: 'Revisión específica de farmacogenómica en migraña: 23 estudios sobre cómo SNPs predicen respuesta a tratamiento. Marco para interpretar análisis genético del proyecto.',
    yaEnProyecto: false },
  { id: 36, bloque: 'Medicina de precisión y biomarcadores', clave: false,
    autores: 'Belvís R, Salas-Cabrera J, García-Azorín D, et al.',
    año: 2025, titulo: 'Microbiomes, microRNAs, and metabolomics in migraine: state-of-the-art and future directions',
    revista: 'Front Neurol', vol: '15:1505844',
    url: 'https://www.ncbi.nlm.nih.gov/pmc/articles/PMC11100066/',
    tags: ['miRNA', 'Biomarcadores'],
    porque: 'miRNA-155 y otros en migraña. Apoya la metodología miRNA de PREDIMIGRAINE (HULP).',
    yaEnProyecto: false },
  { id: 37, bloque: 'Medicina de precisión y biomarcadores', clave: false,
    autores: 'Greco R, Demartini C, Zanaboni AM, et al.',
    año: 2020, titulo: 'Plasma levels of CGRP and expression of specific microRNAs in blood cells of episodic and chronic migraine subjects',
    revista: 'J Headache Pain', vol: '21:122',
    url: 'https://www.ncbi.nlm.nih.gov/pmc/articles/PMC7565351/',
    tags: ['CGRP', 'miRNA', 'Combinado'],
    porque: 'Combina análisis CGRP + miRNA en ME vs MC. Análogo metodológico al multi-ómics de PREDIMIGRAINE.',
    yaEnProyecto: false },
  { id: 38, bloque: 'Medicina de precisión y biomarcadores', clave: false,
    autores: 'Pérez-Pereda S, Toriello-Suárez M, González-Quintanilla V, et al.',
    año: 2023, titulo: 'A study of differential microRNA expression profile in migraine: the microMIG exploratory study',
    revista: 'J Headache Pain', vol: '24:13',
    url: 'https://www.ncbi.nlm.nih.gov/pmc/articles/PMC9936672/',
    tags: ['España', 'miRNA', 'Hospital La Princesa'],
    porque: 'Estudio del Hospital La Princesa (centro coordinador PREDIMIGRAINE) sobre miRNAs en migraña. Base metodológica directamente aplicable.',
    yaEnProyecto: false },

  // ============ BLOQUE 5: NEUROIMAGEN ============
  { id: 39, bloque: 'Neuroimagen', clave: false,
    autores: 'Russo A, Silvestro M, Tessitore A, Tedeschi G',
    año: 2018, titulo: 'Functional Neuroimaging Biomarkers in Migraine: Diagnostic, Prognostic and Therapeutic Implications',
    revista: 'Curr Med Chem', vol: '26(34):6309-6322',
    url: 'https://pubmed.ncbi.nlm.nih.gov/29623825/',
    tags: ['Review', 'Neuroimagen'],
    porque: 'Revisión de biomarcadores de neuroimagen en migraña con énfasis en valor predictivo. Base para la parte de RM cerebral de PREDIMIGRAINE (Valladolid).',
    yaEnProyecto: false },
  { id: 40, bloque: 'Neuroimagen', clave: false,
    autores: 'Wei HL, Tian T, Yu YS, et al.',
    año: 2024, titulo: 'Exploring potential neuroimaging biomarkers for the response to non-steroidal anti-inflammatory drugs in episodic migraine',
    revista: 'J Headache Pain', vol: '25:104',
    url: 'https://link.springer.com/article/10.1186/s10194-024-01812-4',
    tags: ['Neuroimagen', 'ML', 'Predicción'],
    porque: 'Modelo multimodal MRI + ML para predecir respuesta. Marco metodológico para la parte de neuroimagen predictiva de PREDIMIGRAINE.',
    yaEnProyecto: false },
  { id: 41, bloque: 'Neuroimagen', clave: false,
    autores: 'Romozzi M, Messina R',
    año: 2026, titulo: '2025 Highlights - progress towards discovery of molecular and neuroimaging biomarkers for migraine',
    revista: 'Cephalalgia', vol: '46(1):03331024251408766',
    url: 'https://journals.sagepub.com/doi/10.1177/03331024251408766',
    tags: ['Review', 'Reciente'],
    porque: 'Highlights más recientes sobre biomarcadores moleculares y de neuroimagen en migraña. Estado del arte 2026 para situar PREDIMIGRAINE.',
    yaEnProyecto: false },

  // ============ BLOQUE 6: DIGITAL HEALTH Y WEARABLES ============
  { id: 42, bloque: 'Wearables y digital health', clave: true,
    autores: 'Gago-Veiga AB, Gonzalez-Martinez A, Quintas S, Vieira A, Gálvez-Goicuría J, Sanz-García A, Ayala JL, Sobrado M, Vivancos J, Pagan J',
    año: 2025, titulo: 'Episodic Migraine Pain Curves: Real-Time Smartphone-Based Analysis and Clinical Implications',
    revista: 'J Pain Research', vol: '18:6711-6722',
    url: 'https://pubmed.ncbi.nlm.nih.gov/...',
    tags: ['Clave', 'Proyecto', 'Wearable'],
    porque: 'Publicación propia del proyecto (Hospital La Princesa). Clasifica 4 tipos de curva de dolor con app smartphone. Base directa de PREDICRISIS.',
    yaEnProyecto: true },
  { id: 43, bloque: 'Wearables y digital health', clave: false,
    autores: 'Pagán J, Risco-Martín JL, Moya JM, Ayala JL',
    año: 2018, titulo: 'Using Sleep Time Data from Wearable Sensors for Early Detection of Migraine Attacks',
    revista: 'Sensors', vol: '18(4):1025',
    url: 'https://www.ncbi.nlm.nih.gov/pmc/articles/PMC5981434/',
    tags: ['Wearable', 'Sleep', 'Predicción 84%'],
    porque: 'Empatica E4 + ML para detectar migraña la noche siguiente con 84% accuracy. Mismo grupo del proyecto. Antecedente directo de PREDICRISIS.',
    yaEnProyecto: false },
  { id: 44, bloque: 'Wearables y digital health', clave: false,
    autores: 'Stubberud A, Ingvaldsen SH, Brenner E, et al.',
    año: 2023, titulo: 'Forecasting migraine with machine learning based on mobile phone diary and wearable data',
    revista: 'Cephalalgia', vol: '43(4):03331024231169244',
    url: 'https://journals.sagepub.com/doi/10.1177/03331024231169244',
    tags: ['ML', 'Diario móvil', 'Wearable'],
    porque: 'Random Forest con AUC 0,62 para predecir cefalea al día siguiente. 18 pacientes, 388 registros. Referencia comparable a PREDICRISIS.',
    yaEnProyecto: false },
  { id: 45, bloque: 'Wearables y digital health', clave: false,
    autores: 'Tomalin LE, Kummer BR, Campbell MC, et al.',
    año: 2025, titulo: 'Can smartwatches predict migraines? Using machine learning (ML) with wearable-derived nocturnal autonomic nervous system (ANS) and sleep metrics for headache prediction',
    revista: 'Preprint Research Square',
    url: 'https://www.ncbi.nlm.nih.gov/pmc/articles/PMC12687821/',
    tags: ['Wearable', 'Mount Sinai', 'Reciente'],
    porque: 'Modelo ML personalizado con EDA, PRV, RR, sueño. Confirma que el ANS nocturno predice migraña al día siguiente. Apoyo conceptual para PREDICRISIS.',
    yaEnProyecto: false },
  { id: 46, bloque: 'Wearables y digital health', clave: false,
    autores: 'Vandenbussche N, Van Hoecke S, Paemeleire K',
    año: 2025, titulo: 'Uncovering the potential of smartphones for behavior monitoring during migraine follow-up',
    revista: 'BMC Med Inform Decis Mak', vol: '25:69',
    url: 'https://link.springer.com/article/10.1186/s12911-025-02916-w',
    tags: ['Smartphone', 'Monitorización'],
    porque: 'Monitorización ambulatoria smartphone + wearable. Pasiva (sensores) + activa (registro). Refuerza enfoque PREDICRISIS.',
    yaEnProyecto: false },

  // ============ BLOQUE 7: CRONIFICACIÓN ============
  { id: 47, bloque: 'Cronificación de la migraña', clave: false,
    autores: 'Lipton RB, Manack Adams A, Buse DC, Fanning KM, Reed ML',
    año: 2016, titulo: 'A Comparison of the Chronic Migraine Epidemiology and Outcomes (CaMEO) Study and American Migraine Prevalence and Prevention (AMPP) Study: Demographics and Headache-Related Disability',
    revista: 'Headache', vol: '56(8):1280-1289',
    url: 'https://pubmed.ncbi.nlm.nih.gov/27160682/',
    tags: ['Cronificación', 'Epidemiología'],
    porque: 'CaMEO + AMPP definen tasa de cronificación anual (≈2,5% MEAF → MC). Base para modelizar "cronificación evitada" en PREDICRISIS.',
    yaEnProyecto: false },
  { id: 48, bloque: 'Cronificación de la migraña', clave: false,
    autores: 'Buse DC, Greisman JD, Baigi K, Lipton RB',
    año: 2019, titulo: 'Migraine progression: a systematic review',
    revista: 'Headache', vol: '59(3):306-338',
    url: 'https://pubmed.ncbi.nlm.nih.gov/30589090/',
    tags: ['Review', 'Cronificación'],
    porque: 'Revisión sistemática de progresión de migraña. Factores de riesgo modificables (uso excesivo, ansiedad, depresión, obesidad). Marco para modelo de cronificación de PREDICRISIS.',
    yaEnProyecto: false },
  { id: 49, bloque: 'Cronificación de la migraña', clave: false,
    autores: 'Casalegno F, Pini G, Diener HC',
    año: 2022, titulo: 'Reaching the Nadir of Medication Overuse in Chronic Migraine',
    revista: 'Brain Sci', vol: '12(12):1731',
    url: 'https://www.ncbi.nlm.nih.gov/pmc/articles/PMC9690126/',
    tags: ['MOH', 'Anti-CGRP'],
    porque: 'Anti-CGRP reduce uso excesivo y reduce MOH. Apoya argumento de coste-efectividad por reducir cronificación inducida por fármacos.',
    yaEnProyecto: false },
  { id: 50, bloque: 'Cronificación de la migraña', clave: false,
    autores: 'Pozo-Rosich P et al.',
    año: 2015, titulo: 'Diferencias entre migraña crónica con y sin uso excesivo de medicación: experiencia en una serie hospitalaria de 434 pacientes',
    revista: 'Neurología', vol: '30(3):153-160',
    url: 'https://www.elsevier.es/es-revista-neurologia-295-articulo-diferencias-entre-migrana-cronica-con-S0213485313002892',
    tags: ['España', 'MOH'],
    porque: 'Serie española de 434 pacientes MC con/sin UEM. Datos para contextualizar muestra PREDIMIGRAINE.',
    yaEnProyecto: false },

  // ============ BLOQUE 8: HCRU Y URGENCIAS ============
  { id: 51, bloque: 'HCRU y urgencias', clave: false,
    autores: 'Vo P, Paris N, Bilitou A, et al.',
    año: 2018, titulo: 'Migraine-related healthcare resource use in the emergency department setting: a panel-based chart review in France, Germany, Italy, and Spain',
    revista: 'J Med Econ', vol: '21(11):1090-1098',
    url: 'https://www.tandfonline.com/doi/full/10.1080/13696998.2019.1636052',
    tags: ['España', 'Urgencias', 'HCRU'],
    porque: '8 h media en urgencias por migraña. 82% análisis de sangre, 62% ECG, 46% TAC craneal. Cuantifica el coste evitable de visitas a urgencias por migraña mal controlada. Útil para PREDIMIGRAINE como referencia de coste por episodio de urgencias.',
    yaEnProyecto: false },
  { id: 52, bloque: 'HCRU y urgencias', clave: false,
    autores: 'Sanderson JC, Devine EB, Lipton RB, et al.',
    año: 2013, titulo: 'Headache-related healthcare resource utilization in the United States',
    revista: 'Cephalalgia', vol: '33(15):1227-1238',
    url: 'https://pubmed.ncbi.nlm.nih.gov/23720498/',
    tags: ['HCRU'],
    porque: 'Cuantificación HCRU en migraña por gravedad. Estándar de referencia metodológica.',
    yaEnProyecto: false },
  { id: 53, bloque: 'HCRU y urgencias', clave: false,
    autores: 'Sumelahti ML, Sumanen M, Sumanen MS, et al.',
    año: 2024, titulo: 'Healthcare resource utilization and associated costs in migraine patients',
    revista: 'PLoS One', vol: '19(3):e0300816',
    url: 'https://journals.plos.org/plosone/article/file?id=10.1371/journal.pone.0300816&type=printable',
    tags: ['HCRU', 'Costes'],
    porque: 'HCRU estratificado por uso de triptanes/preventivos/anti-CGRP. Útil para construir comparador "práctica habitual" de PREDIMIGRAINE.',
    yaEnProyecto: false },

  // ============ BLOQUE 9: SEGURIDAD ANTI-CGRP ============
  { id: 54, bloque: 'Seguridad anti-CGRP', clave: true,
    autores: 'Sanabria-Gago C, Lázaro IF, Heredia P, ..., Gago-Veiga AB',
    año: 2025, titulo: 'SAFE-CGRP study: multicenter retrospective evaluation of the safety of CGRP pathway-targeting monoclonal antibodies in migraine with relevant comorbidities or conditions excluded from trials',
    revista: 'Front Neurol', vol: '16:1703876',
    url: 'https://www.frontiersin.org/articles/10.3389/fneur.2025.1703876',
    tags: ['Clave', 'Proyecto', 'Seguridad'],
    porque: 'Publicación propia del proyecto. Seguridad anti-CGRP en pacientes con comorbilidades. Apoya la robustez clínica de incluir comorbilidades en el modelo predictivo.',
    yaEnProyecto: true },

  // ============ BLOQUE 10: GUÍAS METODOLÓGICAS ============
  { id: 55, bloque: 'Guías metodológicas', clave: true,
    autores: 'Pozo-Rosich P, van Veelen N, Caronna E, et al.',
    año: 2025, titulo: 'Guidelines of the International Headache Society for Real-World Evidence studies in migraine and cluster headache',
    revista: 'Cephalalgia', vol: '45(2):03331024251318016',
    url: 'https://journals.sagepub.com/doi/10.1177/03331024251318016',
    tags: ['Clave', 'Guías IHS', 'Real-world'],
    porque: 'Guías IHS para estudios real-world en migraña, incluyendo recomendaciones para evaluación económica. Imprescindible para alinear PREDIMIGRAINE con estándares internacionales.',
    yaEnProyecto: false },
  { id: 56, bloque: 'Guías metodológicas', clave: false,
    autores: 'López-Bastida J, Oliva J, Antoñanzas F, et al.',
    año: 2010, titulo: 'Propuesta de guía para la evaluación económica aplicada a las tecnologías sanitarias',
    revista: 'Gaceta Sanitaria', vol: '24(2):154-170',
    url: 'https://www.gacetasanitaria.org/es-propuesta-guia-evaluacion-economica-aplicada-articulo-S021391110900027X',
    tags: ['España', 'Guías'],
    porque: 'Guía oficial de evaluación económica en España. Base normativa para reportar correctamente el análisis económico de PREDIMIGRAINE.',
    yaEnProyecto: false },
  { id: 57, bloque: 'Guías metodológicas', clave: false,
    autores: 'Husereau D, Drummond M, Augustovski F, et al.',
    año: 2022, titulo: 'Consolidated Health Economic Evaluation Reporting Standards 2022 (CHEERS 2022) Statement',
    revista: 'Value in Health', vol: '25(1):3-9',
    url: 'https://pubmed.ncbi.nlm.nih.gov/35031096/',
    tags: ['Clave', 'CHEERS', 'Reporting'],
    porque: 'Estándar internacional CHEERS 2022 para reportar evaluaciones económicas en salud. Imprescindible para que el análisis de PREDIMIGRAINE sea publicable.',
    yaEnProyecto: false },
  { id: 58, bloque: 'Guías metodológicas', clave: false,
    autores: 'Collins GS, Reitsma JB, Altman DG, Moons KGM',
    año: 2024, titulo: 'TRIPOD+AI statement: updated guidance for reporting clinical prediction models that use regression or machine learning methods',
    revista: 'BMJ', vol: '385:e078378',
    url: 'https://www.bmj.com/content/385/bmj-2023-078378',
    tags: ['Clave', 'TRIPOD+AI', 'ML reporting'],
    porque: 'TRIPOD+AI: estándar de reporting para modelos predictivos con ML. Imprescindible para reportar los modelos de PREDICRISIS y PREDICGRP.',
    yaEnProyecto: false },
];

const bloqueColors = {
  'Carga económica y epidemiología': '#9b2226',
  'Coste-efectividad anti-CGRP': '#2d6a4f',
  'Utilidades y QALY': '#5c4677',
  'Medicina de precisión y biomarcadores': '#bc6c25',
  'Neuroimagen': '#1c5d7e',
  'Wearables y digital health': '#7c5e10',
  'Cronificación de la migraña': '#9b2226',
  'HCRU y urgencias': '#2d6a4f',
  'Seguridad anti-CGRP': '#5c4677',
  'Guías metodológicas': '#1a1a1a',
};

function exportCSV(rows) {
  const headers = ['ID', 'Bloque', 'Año', 'Autores', 'Título', 'Revista', 'Volumen', 'DOI/URL', 'Por qué', 'Tags', 'Ya en proyecto'];
  const lines = [headers.join(';')];
  rows.forEach(r => {
    const fields = [r.id, r.bloque, r.año, r.autores, r.titulo, r.revista, r.vol || '', r.url, r.porque, r.tags.join(', '), r.yaEnProyecto ? 'Sí' : 'No'];
    lines.push(fields.map(f => `"${String(f ?? '').replace(/"/g, '""')}"`).join(';'));
  });
  const blob = new Blob(['\uFEFF' + lines.join('\n')], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `predimigraine_bibliografia_${new Date().toISOString().slice(0,10)}.csv`;
  a.click();
  URL.revokeObjectURL(url);
}

function exportExcel(rows) {
  const data = rows.map(r => ({
    ID: r.id, Bloque: r.bloque, Año: r.año, Autores: r.autores,
    Título: r.titulo, Revista: r.revista, Volumen: r.vol || '',
    URL: r.url, 'Por qué es relevante': r.porque,
    Tags: r.tags.join(', '), 'Ya en proyecto': r.yaEnProyecto ? 'Sí' : 'No',
  }));
  const ws = XLSX.utils.json_to_sheet(data);
  ws['!cols'] = [{wch:5},{wch:30},{wch:8},{wch:35},{wch:50},{wch:25},{wch:15},{wch:30},{wch:60},{wch:25},{wch:15}];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Bibliografía');
  XLSX.writeFile(wb, `predimigraine_bibliografia_${new Date().toISOString().slice(0,10)}.xlsx`);
}

function App() {
  const [filtroBloque, setFiltroBloque] = useState('todos');
  const [soloClave, setSoloClave] = useState(false);
  const [busqueda, setBusqueda] = useState('');

  const bloques = ['todos', ...Array.from(new Set(refs.map(r => r.bloque)))];

  const filtradas = useMemo(() => {
    return refs.filter(r => {
      if (filtroBloque !== 'todos' && r.bloque !== filtroBloque) return false;
      if (soloClave && !r.clave) return false;
      if (busqueda) {
        const q = busqueda.toLowerCase();
        const enQuery = (r.titulo + r.autores + r.tags.join(' ') + r.porque).toLowerCase();
        if (!enQuery.includes(q)) return false;
      }
      return true;
    });
  }, [filtroBloque, soloClave, busqueda]);

  const stats = useMemo(() => ({
    total: refs.length,
    clave: refs.filter(r => r.clave).length,
    proyecto: refs.filter(r => r.yaEnProyecto).length,
    españa: refs.filter(r => r.tags.includes('España')).length,
  }), []);

  const agrupadas = useMemo(() => {
    const g = {};
    filtradas.forEach(r => { if (!g[r.bloque]) g[r.bloque] = []; g[r.bloque].push(r); });
    return g;
  }, [filtradas]);

  return (
    <div style={{ maxWidth: '1200px', margin: '0 auto', padding: '32px 24px', color: '#1a1a1a' }}>
      <header style={{ borderBottom: '2px solid #1a1a1a', paddingBottom: '20px', marginBottom: '32px' }}>
        <div style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: '11px', letterSpacing: '0.15em', textTransform: 'uppercase', marginBottom: '8px', color: '#666' }}>
          PREDIMIGRAINE · Bibliografía evaluación económica
        </div>
        <h1 style={{ fontSize: '38px', fontWeight: 700, margin: '0 0 8px 0', lineHeight: 1.1, letterSpacing: '-0.02em' }}>
          Bibliografía recomendada
        </h1>
        <p style={{ fontSize: '16px', color: '#555', margin: 0, fontStyle: 'italic' }}>
          Referencias clave para sustentar el análisis coste-beneficio, coste-utilidad, coste-efectividad e impacto presupuestario del proyecto
        </p>
        <p style={{ fontSize: '13px', color: '#888', marginTop: '12px', marginBottom: 0 }}>
          {stats.total} referencias · {stats.clave} marcadas como clave · {stats.proyecto} ya en el proyecto · {stats.españa} centradas en España
        </p>
      </header>

      <div className="no-print" style={{ marginBottom: '24px', padding: '16px', background: '#1a1a1a', display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: '12px' }}>
        <div style={{ color: '#faf8f3', fontFamily: 'JetBrains Mono, monospace', fontSize: '11px', letterSpacing: '0.1em', textTransform: 'uppercase' }}>
          {filtradas.length === refs.length ? `Exportar ${refs.length} referencias` : `Exportar ${filtradas.length} de ${refs.length}`}
        </div>
        <div style={{ display: 'flex', gap: '8px', flexWrap: 'wrap' }}>
          <button className="export-btn alt" onClick={() => exportCSV(filtradas)}>↓ CSV</button>
          <button className="export-btn" onClick={() => exportExcel(filtradas)}>↓ Excel</button>
          <button className="export-btn alt" onClick={() => window.print()}>⎙ Imprimir / PDF</button>
        </div>
      </div>

      <div className="no-print" style={{ marginBottom: '24px', padding: '16px', background: '#fff', border: '1px solid #ddd' }}>
        <input
          type="text"
          placeholder="Buscar por título, autor, tag, justificación..."
          value={busqueda}
          onChange={e => setBusqueda(e.target.value)}
          style={{ width: '100%', padding: '10px', fontFamily: 'inherit', fontSize: '14px', border: '1px solid #ccc', marginBottom: '12px', background: '#faf8f3' }}
        />
        <div style={{ marginBottom: '10px' }}>
          <span style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: '10px', letterSpacing: '0.1em', textTransform: 'uppercase', color: '#666', marginRight: '12px' }}>Bloque:</span>
          {bloques.map(b => (
            <button key={b} className={`filtro-btn ${filtroBloque === b ? 'activo' : ''}`} onClick={() => setFiltroBloque(b)} style={{ marginRight: '6px', marginBottom: '4px' }}>
              {b === 'todos' ? 'Todos' : b}
            </button>
          ))}
        </div>
        <div>
          <button className={`filtro-btn ${soloClave ? 'activo' : ''}`} onClick={() => setSoloClave(!soloClave)}>
            ★ Solo referencias clave
          </button>
        </div>
      </div>

      {Object.entries(agrupadas).map(([bloque, items]) => (
        <div key={bloque} style={{ marginBottom: '32px' }}>
          <h2 style={{
            fontSize: '13px', fontFamily: 'JetBrains Mono, monospace',
            letterSpacing: '0.15em', textTransform: 'uppercase',
            color: bloqueColors[bloque] || '#1a1a1a',
            margin: '0 0 12px 0', paddingBottom: '8px',
            borderBottom: `2px solid ${bloqueColors[bloque] || '#1a1a1a'}`,
          }}>
            {bloque} <span style={{ color: '#999', fontWeight: 400, fontSize: '11px' }}>· {items.length} {items.length === 1 ? 'referencia' : 'referencias'}</span>
          </h2>
          {items.map(r => (
            <div key={r.id} className="ref-card">
              <div style={{ display: 'flex', alignItems: 'flex-start', justifyContent: 'space-between', marginBottom: '8px', gap: '12px' }}>
                <div style={{ flex: 1 }}>
                  <div style={{ fontSize: '13px', color: '#888', fontFamily: 'JetBrains Mono, monospace', marginBottom: '4px' }}>
                    {r.autores} · {r.año}
                  </div>
                  <h3 style={{ fontSize: '17px', fontWeight: 600, margin: '0 0 6px 0', lineHeight: 1.3 }}>
                    {r.clave && <span style={{ color: '#bc6c25', marginRight: '6px' }}>★</span>}
                    {r.yaEnProyecto && <span className="tag project">Ya en proyecto</span>}
                    {r.titulo}
                  </h3>
                  <div style={{ fontSize: '13px', color: '#666', fontStyle: 'italic', marginBottom: '8px' }}>
                    {r.revista}{r.vol ? `, ${r.vol}` : ''}
                  </div>
                </div>
              </div>
              <div style={{ marginBottom: '10px' }}>
                {r.tags.map((t, i) => {
                  const cls = t === 'España' ? 'spain' : t === 'Clave' ? 'key' : t.includes('Review') ? 'review' : 'cuna';
                  return <span key={i} className={`tag ${cls}`}>{t}</span>;
                })}
              </div>
              <div style={{ fontSize: '14px', color: '#333', lineHeight: 1.5, padding: '10px 12px', background: '#faf8f3', borderLeft: `3px solid ${bloqueColors[bloque] || '#1a1a1a'}`, marginBottom: '8px' }}>
                <strong style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: '10px', letterSpacing: '0.1em', textTransform: 'uppercase', color: '#666' }}>Por qué es relevante</strong>
                <div style={{ marginTop: '4px' }}>{r.porque}</div>
              </div>
              {r.url && (
                <a href={r.url} target="_blank" rel="noopener noreferrer" style={{ fontSize: '13px', fontFamily: 'JetBrains Mono, monospace' }}>
                  ↗ {r.doi || 'Enlace'}
                </a>
              )}
            </div>
          ))}
        </div>
      ))}

      {filtradas.length === 0 && (
        <div style={{ padding: '40px', textAlign: 'center', background: '#fff', border: '1px dashed #ccc', color: '#999', fontStyle: 'italic' }}>
          No hay referencias que coincidan con los filtros
        </div>
      )}

      <div style={{ marginTop: '40px', padding: '24px', background: '#1a1a1a', color: '#faf8f3' }}>
        <h3 style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: '11px', letterSpacing: '0.15em', textTransform: 'uppercase', margin: '0 0 16px 0' }}>
          Cómo usar esta bibliografía
        </h3>
        <ol style={{ margin: 0, paddingLeft: '24px', lineHeight: 1.7, fontSize: '14px' }}>
          <li><strong>Introducción del análisis económico:</strong> referencias 1, 3, 4, 5, 26, 33, 47.</li>
          <li><strong>Justificación metodológica del CBA:</strong> referencias 9 (plantilla), 6, 14, 16.</li>
          <li><strong>Construcción del CUA/QALY:</strong> referencias 19, 20, 22, 23 (utilidades) + 24, 25 (umbrales España).</li>
          <li><strong>Análisis del modelo predictivo:</strong> referencias 27, 28, 29, 58 (TRIPOD+AI).</li>
          <li><strong>Reporte y publicación:</strong> referencias 55 (IHS), 56 (España), 57 (CHEERS).</li>
          <li><strong>Discusión de impacto presupuestario:</strong> referencias 1, 2, 3, 8, 13.</li>
        </ol>
      </div>

      <footer style={{ marginTop: '32px', paddingTop: '16px', borderTop: '1px solid #ddd', fontFamily: 'JetBrains Mono, monospace', fontSize: '10px', color: '#999', textAlign: 'center' }}>
        Bibliografía compilada para soporte del análisis económico PREDIMIGRAINE (PMP22/00158) · {refs.length} referencias en 10 bloques temáticos
      </footer>
    </div>
  );
}

ReactDOM.createRoot(document.getElementById('root')).render(<App />);
</script>
</body>
</html>
g bibliografia_evaluacion_economica.html…]()
