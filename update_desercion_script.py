import re
import os

evid_path = r'C:\Users\GECA\.gemini\antigravity\scratch\sim-cbc-para-github\pages\evidencias.js'
sigeca_evid_path = r'C:\Users\GECA\.gemini\antigravity\scratch\sigeca\pages\evidencias.js'

with open(evid_path, 'r', encoding='utf-8') as f:
    content = f.read()

new_section_code = '''
  // ==========================================================================
  // PUNTO 2: DESERCIÓN ESTUDIANTIL (DATOS OFICIALES GOOGLE SHEETS UPT)
  // ==========================================================================
  const DESERCION_SHEET_DATA = {
    institucional: {
      semestres: [
        { sem: '2025-I', val: 3.46, color: '#3b82f6' },
        { sem: '2025-II', val: 4.26, color: '#f59e0b' },
        { sem: '2026-I', val: 4.17, color: '#10b981' }
      ],
      promedio: 3.97,
      maxAlcanzado: 14.30
    },
    rankingTop4: [
      { rank: 1, carrera: 'PSICOLOGÍA', facultad: 'FAEDCOH', promedio: 11.90, max: 14.30, sem2026: 8.60, riesgo: 'CRÍTICO', color: '#f43f5e' },
      { rank: 2, carrera: 'CIENCIAS DE LA COMUNICACIÓN', facultad: 'FAEDCOH', promedio: 8.68, max: 9.09, sem2026: 8.18, riesgo: 'MEDIO', color: '#f59e0b' },
      { rank: 3, carrera: 'INGENIERÍA AGROINDUSTRIAL', facultad: 'FAING', promedio: 8.44, max: 9.09, sem2026: 9.09, riesgo: 'MEDIO', color: '#f59e0b' },
      { rank: 4, carrera: 'ODONTOLOGÍA', facultad: 'FACSA', promedio: 7.73, max: 9.80, sem2026: 6.59, riesgo: 'MEDIO', color: '#f59e0b' }
    ],
    carreras: [
      { num: 1, facultad: 'FAEDCOH', carrera: 'EDUCACIÓN INICIAL', sem2025_1: 0.00, sem2025_2: 0.00, sem2026_1: 0.00, promedio: 0.00, max: 0.00, riesgo: 'BAJO' },
      { num: 2, facultad: 'FAEDCOH', carrera: 'EDUCACIÓN PRIMARIA', sem2025_1: 0.00, sem2025_2: 0.00, sem2026_1: 3.00, promedio: 1.00, max: 3.00, riesgo: 'BAJO' },
      { num: 3, facultad: 'FAEDCOH', carrera: 'EDUCACIÓN FISICA Y DEPORTES', sem2025_1: 0.00, sem2025_2: 0.00, sem2026_1: 0.00, promedio: 0.00, max: 0.00, riesgo: 'BAJO' },
      { num: 4, facultad: 'FAEDCOH', carrera: 'CIENCIAS DE LA COMUNICACIÓN', sem2025_1: 9.09, sem2025_2: 8.78, sem2026_1: 8.18, promedio: 8.68, max: 9.09, riesgo: 'MEDIO' },
      { num: 5, facultad: 'FAEDCOH', carrera: 'PSICOLOGÍA', sem2025_1: 14.30, sem2025_2: 12.80, sem2026_1: 8.60, promedio: 11.90, max: 14.30, riesgo: 'CRÍTICO' },
      { num: 6, facultad: 'FAING', carrera: 'INGENIERÍA DE SISTEMAS', sem2025_1: 4.85, sem2025_2: 7.25, sem2026_1: 5.00, promedio: 5.70, max: 7.25, riesgo: 'MEDIO' },
      { num: 7, facultad: 'FAING', carrera: 'INGENIERÍA ELECTRÓNICA', sem2025_1: 7.46, sem2025_2: 4.55, sem2026_1: 6.41, promedio: 6.14, max: 7.46, riesgo: 'MEDIO' },
      { num: 8, facultad: 'FAING', carrera: 'INGENIERÍA CIVIL', sem2025_1: 2.43, sem2025_2: 6.91, sem2026_1: 10.81, promedio: 6.72, max: 10.81, riesgo: 'MEDIO' },
      { num: 9, facultad: 'FAING', carrera: 'INGENIERÍA AGROINDUSTRIAL', sem2025_1: 7.14, sem2025_2: 9.08, sem2026_1: 9.09, promedio: 8.44, max: 9.09, riesgo: 'MEDIO' },
      { num: 10, facultad: 'FAING', carrera: 'INGENIERÍA AMBIENTAL', sem2025_1: 3.68, sem2025_2: 3.84, sem2026_1: 4.02, promedio: 3.85, max: 4.02, riesgo: 'BAJO' },
      { num: 11, facultad: 'FAING', carrera: 'INGENIERÍA INDUSTRIAL', sem2025_1: 0.00, sem2025_2: 7.32, sem2026_1: 3.17, promedio: 3.50, max: 7.32, riesgo: 'BAJO' },
      { num: 12, facultad: 'FADE', carrera: 'DERECHO', sem2025_1: 5.76, sem2025_2: 7.13, sem2026_1: 6.13, promedio: 6.34, max: 7.13, riesgo: 'MEDIO' },
      { num: 13, facultad: 'FACSA', carrera: 'MEDICINA', sem2025_1: 0.00, sem2025_2: 0.00, sem2026_1: 0.28, promedio: 0.09, max: 0.28, riesgo: 'BAJO' },
      { num: 14, facultad: 'FACSA', carrera: 'ODONTOLOGÍA', sem2025_1: 9.80, sem2025_2: 6.80, sem2026_1: 6.59, promedio: 7.73, max: 9.80, riesgo: 'MEDIO' },
      { num: 15, facultad: 'FACSA', carrera: 'TECNOLOGÍA MÉDICA TERAPIA FÍSICA', sem2025_1: 3.00, sem2025_2: 2.00, sem2026_1: 2.00, promedio: 2.33, max: 3.00, riesgo: 'BAJO' },
      { num: 16, facultad: 'FACSA', carrera: 'TECNOLOGÍA MÉDICA LABORATORIO CLÍNICO', sem2025_1: 3.00, sem2025_2: 9.00, sem2026_1: 3.00, promedio: 5.00, max: 9.00, riesgo: 'MEDIO' },
      { num: 17, facultad: 'FAU', carrera: 'ARQUITECTURA', sem2025_1: 3.40, sem2025_2: 3.24, sem2026_1: 4.48, promedio: 3.71, max: 4.48, riesgo: 'BAJO' },
      { num: 18, facultad: 'FAU', carrera: 'URBANISMO', sem2025_1: 0.00, sem2025_2: 0.00, sem2026_1: 0.00, promedio: 0.00, max: 0.00, riesgo: 'BAJO' },
      { num: 19, facultad: 'FACEM', carrera: 'INGENIERÍA COMERCIAL', sem2025_1: 3.00, sem2025_2: 2.00, sem2026_1: 7.00, promedio: 4.00, max: 7.00, riesgo: 'BAJO' },
      { num: 20, facultad: 'FACEM', carrera: 'CIENCIAS CONTABLES Y FINANCIERAS', sem2025_1: 4.50, sem2025_2: 6.70, sem2026_1: 6.60, promedio: 5.93, max: 6.70, riesgo: 'MEDIO' },
      { num: 21, facultad: 'FACEM', carrera: 'ADMINISTRACIÓN DE EMPRESAS', sem2025_1: 0.00, sem2025_2: 0.00, sem2026_1: 0.00, promedio: 0.00, max: 0.00, riesgo: 'BAJO' },
      { num: 22, facultad: 'FACEM', carrera: 'ADMINISTRACIÓN DE NEGOCIOS INTERNACIONALES', sem2025_1: 0.70, sem2025_2: 1.32, sem2026_1: 1.00, promedio: 1.01, max: 1.32, riesgo: 'BAJO' },
      { num: 23, facultad: 'FACEM', carrera: 'ADMINISTRACIÓN TURÍSTICA HOTELERA', sem2025_1: 0.93, sem2025_2: 3.57, sem2026_1: 4.81, promedio: 3.10, max: 4.81, riesgo: 'BAJO' },
      { num: 24, facultad: 'FACEM', carrera: 'ECONOMÍA', sem2025_1: 0.00, sem2025_2: 0.00, sem2026_1: 0.00, promedio: 0.00, max: 0.00, riesgo: 'BAJO' }
    ]
  };

  const renderDesercionYRezagoSectionHTML = (activeSem, selFac, selCarrera) => {
    if (!window._desercionSubView) window._desercionSubView = "historico_general";
    if (!window._desercionSelectedFac) window._desercionSelectedFac = "FAEDCOH";

    const subView = window._desercionSubView;
    const selectedFac = window._desercionSelectedFac;
    const allFaculties = ["FAEDCOH", "FAING", "FADE", "FACSA", "FAU", "FACEM"];

    let dynamicViewHTML = "";

    if (subView === "historico_general") {
      const chartBars = DESERCION_SHEET_DATA.institucional.semestres;
      const maxVal = 6.0;

      dynamicViewHTML = `
        <div style="display: flex; flex-direction: column; gap: 1.25rem;">
          <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 0.5rem;">
            <h4 style="font-size: 1.05rem; font-weight: 700; color: #fff; margin: 0; display: flex; align-items: center; gap: 8px;">
              <i data-lucide="bar-chart-2" style="color: var(--accent-light); width: 18px; height: 18px;"></i>
              <span>Histórico General de Deserción Institucional UPT (2025 - 2026)</span>
            </h4>
            <span style="font-size: 0.8rem; color: var(--text-300); font-weight: 600; background: rgba(59, 130, 246, 0.15); border: 1px solid rgba(59, 130, 246, 0.3); padding: 0.2rem 0.6rem; border-radius: 6px;">
              Promedio Histórico UPT: <strong style="color: #60a5fa;">3.97%</strong>
            </span>
          </div>

          <!-- Gráfico de Barras Histórico General -->
          <div style="background: var(--bg-dark-900); padding: 1.25rem; border-radius: 8px; border: 1px solid var(--border-color);">
            <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 1.5rem; margin-bottom: 1.5rem;">
              ${chartBars.map(b => `
                <div style="display: flex; flex-direction: column; align-items: center; background: var(--bg-dark-800); padding: 1rem; border-radius: 8px; border: 1px solid var(--border-color);">
                  <span style="font-size: 0.8rem; color: var(--text-400); font-weight: 600; margin-bottom: 0.4rem;">Semestre ${b.sem}</span>
                  <div style="height: 120px; width: 45px; background: var(--bg-dark-900); border-radius: 6px; display: flex; align-items: flex-end; padding: 3px; position: relative;">
                    <div style="width: 100%; height: ${((b.val / maxVal) * 100).toFixed(1)}%; background: ${b.color}; border-radius: 4px; transition: height 0.4s ease;"></div>
                  </div>
                  <span style="font-size: 1.1rem; font-weight: 800; color: #fff; margin-top: 0.5rem;">${b.val.toFixed(2)}%</span>
                </div>
              `).join('')}
            </div>

            <!-- Tabla Resumen Institucional -->
            <table class="summary-table" style="width: 100%; font-size: 0.82rem; margin: 0; background: var(--bg-dark-800);">
              <thead>
                <tr>
                  <th>Semestre Académico</th>
                  <th class="text-center">Tasa Deserción UPT (%)</th>
                  <th class="text-center">Variación vs Periodo Previo</th>
                  <th class="text-center">Nivel de Riesgo Global</th>
                </tr>
              </thead>
              <tbody>
                <tr>
                  <td><strong>2025-I</strong></td>
                  <td class="text-center" style="font-weight: 700; color: #60a5fa;">3.46%</td>
                  <td class="text-center" style="color: var(--text-400);">Base Inicial</td>
                  <td class="text-center"><span class="badge-status cumple">BAJO</span></td>
                </tr>
                <tr>
                  <td><strong>2025-II</strong></td>
                  <td class="text-center" style="font-weight: 700; color: #fbbf24;">4.26%</td>
                  <td class="text-center" style="color: #fbbf24;">+0.80%</td>
                  <td class="text-center"><span class="badge-status alerta">MEDIO</span></td>
                </tr>
                <tr>
                  <td><strong>2026-I</strong></td>
                  <td class="text-center" style="font-weight: 700; color: #34d399;">4.17%</td>
                  <td class="text-center" style="color: #34d399;">-0.09%</td>
                  <td class="text-center"><span class="badge-status alerta">MEDIO</span></td>
                </tr>
                <tr style="background: rgba(255,255,255,0.03); font-weight: bold; border-top: 1.5px solid var(--border-color);">
                  <td>PROMEDIO HISTÓRICO ACUMULADO</td>
                  <td class="text-center" style="font-weight: 800; color: var(--accent-light);">3.97%</td>
                  <td class="text-center" style="color: var(--text-300);">Media UPT</td>
                  <td class="text-center"><span class="badge-status cumple">BAJO / ESTABLE</span></td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>
      `;
    } else if (subView === "historico_facultad") {
      const facCarreras = DESERCION_SHEET_DATA.carreras.filter(c => c.facultad === selectedFac);
      
      const sem2025_1_avg = facCarreras.length > 0 ? (facCarreras.reduce((s, c) => s + c.sem2025_1, 0) / facCarreras.length) : 0;
      const sem2025_2_avg = facCarreras.length > 0 ? (facCarreras.reduce((s, c) => s + c.sem2025_2, 0) / facCarreras.length) : 0;
      const sem2026_1_avg = facCarreras.length > 0 ? (facCarreras.reduce((s, c) => s + c.sem2026_1, 0) / facCarreras.length) : 0;

      dynamicViewHTML = `
        <div style="display: flex; flex-direction: column; gap: 1.25rem;">
          <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 0.75rem;">
            <h4 style="font-size: 1.05rem; font-weight: 700; color: #fff; margin: 0; display: flex; align-items: center; gap: 8px;">
              <i data-lucide="building-2" style="color: #60a5fa; width: 18px; height: 18px;"></i>
              <span>Evolución Histórica por Facultad: <strong style="color: var(--accent-light);">${selectedFac}</strong></span>
            </h4>

            <div style="display: flex; gap: 6px; flex-wrap: wrap;">
              ${allFaculties.map(f => `
                <button class="btn ${selectedFac === f ? 'btn-primary' : 'btn-secondary'} btn-desercion-fac-select" data-fac="${f}" style="padding: 0.3rem 0.65rem; font-size: 0.75rem;">
                  ${f}
                </button>
              `).join('')}
            </div>
          </div>

          <!-- Histórico por Facultad -->
          <div class="card" style="padding: 1.25rem; background: var(--bg-dark-900);">
            <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 1rem; margin-bottom: 1.25rem;">
              <div style="background: var(--bg-dark-800); padding: 0.75rem; border-radius: 6px; text-align: center; border: 1px solid var(--border-color);">
                <span style="font-size: 0.75rem; color: var(--text-400);">2025-I</span>
                <div style="font-size: 1.1rem; font-weight: 800; color: #60a5fa; margin-top: 2px;">${sem2025_1_avg.toFixed(2)}%</div>
              </div>
              <div style="background: var(--bg-dark-800); padding: 0.75rem; border-radius: 6px; text-align: center; border: 1px solid var(--border-color);">
                <span style="font-size: 0.75rem; color: var(--text-400);">2025-II</span>
                <div style="font-size: 1.1rem; font-weight: 800; color: #fbbf24; margin-top: 2px;">${sem2025_2_avg.toFixed(2)}%</div>
              </div>
              <div style="background: var(--bg-dark-800); padding: 0.75rem; border-radius: 6px; text-align: center; border: 1px solid var(--border-color);">
                <span style="font-size: 0.75rem; color: var(--text-400);">2026-I</span>
                <div style="font-size: 1.1rem; font-weight: 800; color: #34d399; margin-top: 2px;">${sem2026_1_avg.toFixed(2)}%</div>
              </div>
            </div>

            <!-- Tabla por Carreras de la Facultad -->
            <table class="summary-table" style="width: 100%; font-size: 0.8rem; margin: 0;">
              <thead>
                <tr>
                  <th>N°</th>
                  <th>Carrera Profesional</th>
                  <th class="text-center">2025-I</th>
                  <th class="text-center">2025-II</th>
                  <th class="text-center">2026-I</th>
                  <th class="text-center">Promedio Hist.</th>
                  <th class="text-center">Máx. Alcanzado</th>
                  <th class="text-center">Estado de Riesgo</th>
                </tr>
              </thead>
              <tbody>
                ${facCarreras.map(c => `
                  <tr>
                    <td>${c.num}</td>
                    <td><strong>${c.carrera}</strong></td>
                    <td class="text-center">${c.sem2025_1.toFixed(2)}%</td>
                    <td class="text-center">${c.sem2025_2.toFixed(2)}%</td>
                    <td class="text-center" style="font-weight: 700;">${c.sem2026_1.toFixed(2)}%</td>
                    <td class="text-center" style="font-weight: 800; color: var(--accent-light);">${c.promedio.toFixed(2)}%</td>
                    <td class="text-center" style="color: #f87171;">${c.max.toFixed(2)}%</td>
                    <td class="text-center">
                      <span class="badge-status ${c.riesgo === 'CRÍTICO' ? 'nocumple' : (c.riesgo === 'MEDIO' ? 'alerta' : 'cumple')}">
                        ${c.riesgo}
                      </span>
                    </td>
                  </tr>
                `).join('')}
              </tbody>
            </table>
          </div>
        </div>
      `;
    } else if (subView === "comparativa") {
      const facAverages = allFaculties.map(fac => {
        const facC = DESERCION_SHEET_DATA.carreras.filter(c => c.facultad === fac);
        const avg = facC.length > 0 ? (facC.reduce((s, c) => s + c.promedio, 0) / facC.length) : 0;
        return { facultad: fac, promedio: avg, count: facC.length };
      }).sort((a, b) => b.promedio - a.promedio);

      const maxFacAvg = Math.max(...facAverages.map(f => f.promedio), 1);

      dynamicViewHTML = `
        <div style="display: flex; flex-direction: column; gap: 1.25rem;">
          <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 0.5rem;">
            <h4 style="font-size: 1.05rem; font-weight: 700; color: #fff; margin: 0; display: flex; align-items: center; gap: 8px;">
              <i data-lucide="git-compare" style="color: #a78bfa; width: 18px; height: 18px;"></i>
              <span>Pestaña Comparativa Inter-Facultades y Carreras Profesionales</span>
            </h4>
          </div>

          <!-- Gráfico Comparativo entre Facultades -->
          <div class="card" style="padding: 1.25rem; background: var(--bg-dark-900);">
            <h5 style="font-size: 0.88rem; color: var(--text-300); margin-bottom: 1rem;">Promedio de Deserción por Facultad (Cruce Multisemestral)</h5>

            <div style="display: flex; flex-direction: column; gap: 0.75rem; margin-bottom: 1.5rem;">
              ${facAverages.map(f => `
                <div style="display: flex; align-items: center; gap: 1rem; font-size: 0.8rem;">
                  <span style="width: 90px; font-weight: 700; color: #fff;">${f.facultad}</span>
                  <div style="flex-grow: 1; height: 22px; background: var(--bg-dark-800); border-radius: 4px; overflow: hidden; position: relative;">
                    <div style="width: ${((f.promedio / maxFacAvg) * 100).toFixed(1)}%; height: 100%; background: ${f.promedio > 6 ? '#f43f5e' : (f.promedio > 3.5 ? '#f59e0b' : '#10b981')}; border-radius: 4px; transition: width 0.4s ease;"></div>
                    <span style="position: absolute; right: 10px; top: 2px; font-weight: 800; font-size: 0.75rem; color: #fff;">${f.promedio.toFixed(2)}%</span>
                  </div>
                  <span style="width: 80px; font-size: 0.72rem; color: var(--text-400); text-align: right;">${f.count} carreras</span>
                </div>
              `).join('')}
            </div>

            <!-- Tabla Matriz Comparativa General -->
            <table class="summary-table" style="width: 100%; font-size: 0.78rem; margin: 0;">
              <thead>
                <tr>
                  <th>N°</th>
                  <th>Facultad</th>
                  <th>Carrera Profesional</th>
                  <th class="text-center">2025-I</th>
                  <th class="text-center">2025-II</th>
                  <th class="text-center">2026-I</th>
                  <th class="text-center">Promedio</th>
                  <th class="text-center">Riesgo</th>
                </tr>
              </thead>
              <tbody>
                ${DESERCION_SHEET_DATA.carreras.map(c => `
                  <tr>
                    <td>${c.num}</td>
                    <td><strong>${c.facultad}</strong></td>
                    <td>${c.carrera}</td>
                    <td class="text-center">${c.sem2025_1.toFixed(2)}%</td>
                    <td class="text-center">${c.sem2025_2.toFixed(2)}%</td>
                    <td class="text-center" style="font-weight: 700;">${c.sem2026_1.toFixed(2)}%</td>
                    <td class="text-center" style="font-weight: 800; color: var(--accent-light);">${c.promedio.toFixed(2)}%</td>
                    <td class="text-center">
                      <span class="badge-status ${c.riesgo === 'CRÍTICO' ? 'nocumple' : (c.riesgo === 'MEDIO' ? 'alerta' : 'cumple')}">
                        ${c.riesgo}
                      </span>
                    </td>
                  </tr>
                `).join('')}
              </tbody>
            </table>
          </div>
        </div>
      `;
    } else if (subView === "ranking_top4") {
      const top4 = DESERCION_SHEET_DATA.rankingTop4;

      dynamicViewHTML = `
        <div style="display: flex; flex-direction: column; gap: 1.25rem;">
          <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 0.5rem;">
            <h4 style="font-size: 1.05rem; font-weight: 700; color: #fff; margin: 0; display: flex; align-items: center; gap: 8px;">
              <i data-lucide="trophy" style="color: #f43f5e; width: 18px; height: 18px;"></i>
              <span>Ranking Top 4: Carreras con Mayor Porcentaje de Deserción Estudiantil</span>
            </h4>
            <span style="font-size: 0.78rem; color: #f43f5e; font-weight: 700; background: rgba(244, 63, 94, 0.12); border: 1px solid rgba(244, 63, 94, 0.3); padding: 0.2rem 0.6rem; border-radius: 6px;">
              Foco Crítico Institucional
            </span>
          </div>

          <!-- Cards de Ranking Top 4 -->
          <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(230px, 1fr)); gap: 1rem;">
            ${top4.map(item => `
              <div class="card" style="padding: 1.25rem; background: var(--bg-dark-900); border: 1.5px solid ${item.color}; position: relative; overflow: hidden; display: flex; flex-direction: column; justify-content: space-between;">
                <div style="position: absolute; right: 10px; top: 10px; font-size: 2.2rem; font-weight: 900; opacity: 0.15; color: ${item.color};">
                  #${item.rank}
                </div>
                <div>
                  <div style="display: flex; align-items: center; gap: 6px; margin-bottom: 0.5rem;">
                    <span class="badge-status ${item.riesgo === 'CRÍTICO' ? 'nocumple' : 'alerta'}" style="font-size: 0.68rem; font-weight: 800; padding: 0.15rem 0.45rem;">
                      RANK #${item.rank} • ${item.riesgo}
                    </span>
                    <span style="font-size: 0.72rem; color: var(--text-400); font-weight: 700;">${item.facultad}</span>
                  </div>
                  <h4 style="font-size: 0.95rem; font-weight: 800; color: #fff; margin: 0 0 0.75rem 0; line-height: 1.3;">
                    ${item.carrera}
                  </h4>
                </div>
                
                <div style="background: var(--bg-dark-800); padding: 0.6rem 0.75rem; border-radius: 6px; border: 1px solid var(--border-color); display: flex; justify-content: space-between; align-items: center;">
                  <span style="font-size: 0.75rem; color: var(--text-300);">Promedio Histórico:</span>
                  <span style="font-size: 1.1rem; font-weight: 900; color: ${item.color};">${item.promedio.toFixed(2)}%</span>
                </div>
              </div>
            `).join('')}
          </div>

          <!-- Tabla Detallada Top 4 -->
          <div class="card" style="padding: 1.25rem; background: var(--bg-dark-900);">
            <h5 style="font-size: 0.88rem; color: var(--text-200); margin-bottom: 1rem;">Desglose Semestral de las 4 Carreras en Mayor Riesgo</h5>
            <table class="summary-table" style="width: 100%; font-size: 0.82rem; margin: 0;">
              <thead>
                <tr>
                  <th>Posición</th>
                  <th>Carrera Profesional</th>
                  <th>Facultad</th>
                  <th class="text-center">2025-I</th>
                  <th class="text-center">2025-II</th>
                  <th class="text-center">2026-I</th>
                  <th class="text-center">Promedio Hist.</th>
                  <th class="text-center">Máx. Alcanzado</th>
                  <th class="text-center">Nivel de Riesgo</th>
                </tr>
              </thead>
              <tbody>
                ${top4.map(t => `
                  <tr style="border-left: 3px solid ${t.color};">
                    <td><strong style="color: ${t.color};">#${t.rank}</strong></td>
                    <td><strong>${t.carrera}</strong></td>
                    <td>${t.facultad}</td>
                    <td class="text-center">${t.promedio === 11.9 ? '14.30%' : (t.promedio === 8.68 ? '9.09%' : (t.promedio === 8.44 ? '7.14%' : '9.80%'))}</td>
                    <td class="text-center">${t.promedio === 11.9 ? '12.80%' : (t.promedio === 8.68 ? '8.78%' : (t.promedio === 8.44 ? '9.08%' : '6.80%'))}</td>
                    <td class="text-center" style="font-weight: 700;">${t.sem2026.toFixed(2)}%</td>
                    <td class="text-center" style="font-weight: 900; color: ${t.color}; font-size: 0.9rem;">${t.promedio.toFixed(2)}%</td>
                    <td class="text-center" style="color: #f87171;">${t.max.toFixed(2)}%</td>
                    <td class="text-center">
                      <span class="badge-status ${t.riesgo === 'CRÍTICO' ? 'nocumple' : 'alerta'}">
                        ${t.riesgo}
                      </span>
                    </td>
                  </tr>
                `).join('')}
              </tbody>
            </table>
          </div>
        </div>
      `;
    }

    return `
      <div style="margin-top: 2.5rem; border-top: 2px solid var(--border-color); padding-top: 2rem;">
        
        <!-- ENCABEZADO PUNTO 2: DESERCIÓN ESTUDIANTIL -->
        <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 1rem; margin-bottom: 1.5rem;">
          <div>
            <h3 style="font-size: 1.35rem; font-weight: 800; color: #fff; margin: 0; display: flex; align-items: center; gap: 10px;">
              <i data-lucide="trending-down" style="color: #f43f5e; width: 24px; height: 24px;"></i>
              <span>2. Deserción Estudiantil</span>
            </h3>
            <p style="font-size: 0.85rem; color: var(--text-300); margin-top: 4px;">
              Monitoreo multisemestral por facultad y carrera profesional con indicadores automáticos de tendencia y nivel de riesgo (Fuente oficial UPT).
            </p>
          </div>

          <!-- Pestañas internas de navegación para Deserción -->
          <div class="tab-container" style="border: none; margin: 0; background: var(--bg-dark-900); padding: 4px; border-radius: 8px;">
            <button class="tab-btn ${subView === 'historico_general' ? 'active' : ''} btn-desercion-view" data-view="historico_general">
              <i data-lucide="bar-chart-2" style="width: 14px; height: 14px;"></i>
              <span>Histórico General UPT</span>
            </button>
            <button class="tab-btn ${subView === 'historico_facultad' ? 'active' : ''} btn-desercion-view" data-view="historico_facultad">
              <i data-lucide="building-2" style="width: 14px; height: 14px;"></i>
              <span>Histórico por Facultad</span>
            </button>
            <button class="tab-btn ${subView === 'comparativa' ? 'active' : ''} btn-desercion-view" data-view="comparativa">
              <i data-lucide="git-compare" style="width: 14px; height: 14px;"></i>
              <span>Pestaña Comparativa</span>
            </button>
            <button class="tab-btn ${subView === 'ranking_top4' ? 'active' : ''} btn-desercion-view" data-view="ranking_top4" style="color: #f43f5e;">
              <i data-lucide="trophy" style="width: 14px; height: 14px;"></i>
              <span>Ranking Top 4</span>
            </button>
          </div>
        </div>

        <!-- KPI Summary Cards -->
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 1rem; margin-bottom: 1.5rem;">
          <div class="card" style="padding: 1rem; background: var(--bg-dark-900);">
            <span style="font-size: 0.75rem; color: var(--text-400); font-weight: 600;">Promedio Histórico UPT</span>
            <div style="font-size: 1.4rem; font-weight: 900; color: #60a5fa; margin-top: 4px;">3.97%</div>
            <span style="font-size: 0.7rem; color: var(--text-300);">Semestres 2025 - 2026</span>
          </div>

          <div class="card" style="padding: 1rem; background: var(--bg-dark-900);">
            <span style="font-size: 0.75rem; color: var(--text-400); font-weight: 600;">Corte Actual (2026-I)</span>
            <div style="font-size: 1.4rem; font-weight: 900; color: #34d399; margin-top: 4px;">4.17%</div>
            <span style="font-size: 0.7rem; color: var(--text-300);">Institucional global</span>
          </div>

          <div class="card" style="padding: 1rem; background: var(--bg-dark-900);">
            <span style="font-size: 0.75rem; color: var(--text-400); font-weight: 600;">Máxima Deserción Registrada</span>
            <div style="font-size: 1.4rem; font-weight: 900; color: #f43f5e; margin-top: 4px;">14.30%</div>
            <span style="font-size: 0.7rem; color: var(--text-300);">Psicología (2025-I)</span>
          </div>

          <div class="card" style="padding: 1rem; background: var(--bg-dark-900);">
            <span style="font-size: 0.75rem; color: var(--text-400); font-weight: 600;">Top 1 Riesgo Crítico</span>
            <div style="font-size: 1.1rem; font-weight: 900; color: #f43f5e; margin-top: 4px;">Psicología (11.90%)</div>
            <span style="font-size: 0.7rem; color: var(--text-300);">FAEDCOH</span>
          </div>
        </div>

        <!-- Vista Dinámica Seleccionada -->
        ${dynamicViewHTML}

      </div>
    `;
  };
'''

pattern = r'  // HELPER PARA RENDERIZAR ANÁLISIS DE DESERCIÓN Y REZAGO.*?return `.*?<\/div>\s*`;\s*};'
match = re.search(pattern, content, re.DOTALL)
if match:
    updated = content[:match.start()] + new_section_code.strip() + content[match.end():]
    with open(evid_path, 'w', encoding='utf-8') as f:
        f.write(updated)
    with open(sigeca_evid_path, 'w', encoding='utf-8') as f:
        f.write(updated)
    print('SUCCESS: Updated renderDesercionYRezagoSectionHTML in both files!')
else:
    print('ERROR: Pattern match failed!')
