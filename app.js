// Main Application Logic
let globalData = [];
let currentFilteredData = [];
let charts = {};

document.addEventListener('DOMContentLoaded', () => {
    setupDragAndDrop();
});

function setupDragAndDrop() {
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('csv-file');

    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, preventDefaults, false);
    });

    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }

    ['dragenter', 'dragover'].forEach(eventName => {
        dropZone.addEventListener(eventName, () => {
            dropZone.classList.add('dragover');
        }, false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, () => {
            dropZone.classList.remove('dragover');
        }, false);
    });

    dropZone.addEventListener('drop', (e) => {
        const dt = e.dataTransfer;
        handleFiles(dt.files);
    });

    fileInput.addEventListener('change', function() {
        handleFiles(this.files);
    });
}

function handleFiles(files) {
    const uploadError = document.getElementById('upload-error');
    if (!files.length) return;
    
    const file = files[0];
    
    if (file.type !== 'text/csv' && !file.name.endsWith('.csv')) {
        uploadError.innerText = "Por favor, selecciona un archivo .csv válido.";
        return;
    }
    
    uploadError.innerText = "Procesando archivo...";
    
    Papa.parse(file, {
        header: true,
        skipEmptyLines: true,
        encoding: "ISO-8859-1",
        complete: function(results) {
            processData(results.data);
        },
        error: function(err) {
            uploadError.innerText = "Error leyendo el archivo: " + err.message;
        }
    });
}

function normalizeState(stateStr) {
    if (!stateStr) return null;
    const s = stateStr.toString().toUpperCase().trim();
    if (s.includes('INICIO')) return 1;
    if (s.includes('PROCESO')) return 2;
    if (s.includes('LOGRADO')) return 3;
    const num = parseInt(s);
    if (!isNaN(num) && num >= 1 && num <= 3) return num;
    return null;
}

function stateToString(num) {
    if (num === 1) return "INICIO";
    if (num === 2) return "PROCESO";
    if (num === 3) return "LOGRADO";
    return "N/A";
}

function processData(rows) {
    const uploadError = document.getElementById('upload-error');
    if (rows.length === 0) {
        uploadError.innerText = "El archivo CSV está vacío.";
        return;
    }
    
    const sample = rows[0];
    if (!sample.AREA || !sample.DOCENTE || !sample.NOMBRES) {
        uploadError.innerText = "Faltan columnas requeridas (AREA, DOCENTE, NOMBRES).";
        return;
    }

    const processedRows = [];
    let errorFound = false;
    const compRegex = /^COMP(\d+)_C(\d+)$/;

    rows.forEach((row, index) => {
        if (errorFound) return;
        if (!row.NOMBRES || row.NOMBRES.trim() === '') return;

        let comps = {};
        let hasAnyNote = false;

        Object.keys(row).forEach(key => {
            const match = key.match(compRegex);
            if (match) {
                const compId = parseInt(match[1]);
                const note = normalizeState(row[key]);
                if (note !== null) {
                    hasAnyNote = true;
                    if (!comps[compId]) comps[compId] = [];
                    comps[compId].push(note);
                }
            }
        });

        if (!hasAnyNote) {
            uploadError.innerText = `Error fila ${index + 2}: El alumno ${row.NOMBRES} no tiene validaciones evaluadas.`;
            errorFound = true;
            return;
        }

        let compAverages = [];
        let rawComps = {};
        Object.keys(comps).forEach(compId => {
            const notes = comps[compId];
            if (notes.length > 0) {
                const sum = notes.reduce((a, b) => a + b, 0);
                const avg = Math.round(sum / notes.length);
                compAverages.push(avg);
                rawComps[compId] = avg;
            }
        });

        let finalStatus = null;
        if (compAverages.length > 0) {
            const sumFinal = compAverages.reduce((a, b) => a + b, 0);
            finalStatus = Math.round(sumFinal / compAverages.length);
        }

        processedRows.push({
            student: row.NOMBRES,
            course: row.AREA,
            teacher: row.DOCENTE,
            grade: row.GRADO,
            section: row.SECCION,
            finalStatus: finalStatus,
            compAverages: compAverages,
            rawComps: rawComps
        });
    });

    if (errorFound) return;

    globalData = processedRows;
    initDashboard();
}

function initDashboard() {
    document.getElementById('upload-section').classList.add('hidden');
    document.getElementById('dashboard').classList.remove('hidden');
    document.getElementById('export-controls').classList.remove('hidden');
    
    document.getElementById('filter-course').addEventListener('change', () => { updateFilters('course'); updateDashboard(); });
    document.getElementById('filter-teacher').addEventListener('change', () => { updateFilters('teacher'); updateDashboard(); });
    document.getElementById('filter-grade').addEventListener('change', () => { updateFilters('grade'); updateDashboard(); });
    document.getElementById('filter-section').addEventListener('change', () => { updateFilters('section'); updateDashboard(); });
    
    updateFilters('all');
    updateDashboard();
}

function updateFilters(changedSelect) {
    const selects = {
        course: document.getElementById('filter-course'),
        teacher: document.getElementById('filter-teacher'),
        grade: document.getElementById('filter-grade'),
        section: document.getElementById('filter-section')
    };

    const vals = {
        course: selects.course.value,
        teacher: selects.teacher.value,
        grade: selects.grade.value,
        section: selects.section.value
    };

    const getAvailable = (field) => {
        return [...new Set(globalData.filter(d => {
            if (field !== 'course' && vals.course !== 'all' && d.course !== vals.course) return false;
            if (field !== 'teacher' && vals.teacher !== 'all' && d.teacher !== vals.teacher) return false;
            if (field !== 'grade' && vals.grade !== 'all' && d.grade !== vals.grade) return false;
            if (field !== 'section' && vals.section !== 'all' && d.section !== vals.section) return false;
            return true;
        }).map(d => d[field]))].filter(Boolean).sort();
    };

    const repopulate = (field, defaultLabel) => {
        if (changedSelect === field) return;
        const available = getAvailable(field);
        const el = selects[field];
        const currentVal = el.value;
        
        el.innerHTML = `<option value="all">${defaultLabel}</option>`;
        available.forEach(optVal => {
            const opt = document.createElement('option');
            opt.value = optVal;
            opt.innerText = optVal;
            el.appendChild(opt);
        });
        
        if (available.includes(currentVal)) {
            el.value = currentVal;
        } else {
            el.value = 'all'; 
        }
    };

    repopulate('course', 'Todos los cursos');
    repopulate('teacher', 'Todos los docentes');
    repopulate('grade', 'Todos los grados');
    repopulate('section', 'Todas las secciones');
}

function updateDashboard() {
    const courseFilter = document.getElementById('filter-course').value;
    const teacherFilter = document.getElementById('filter-teacher').value;
    const gradeFilter = document.getElementById('filter-grade').value;
    const sectionFilter = document.getElementById('filter-section').value;

    let filtered = globalData;
    if (courseFilter !== 'all') filtered = filtered.filter(d => d.course === courseFilter);
    if (teacherFilter !== 'all') filtered = filtered.filter(d => d.teacher === teacherFilter);
    if (gradeFilter !== 'all') filtered = filtered.filter(d => d.grade === gradeFilter);
    if (sectionFilter !== 'all') filtered = filtered.filter(d => d.section === sectionFilter);

    currentFilteredData = filtered;

    updateKPIs(filtered, courseFilter);
    updateCharts(filtered);
    generateInsights(filtered, courseFilter);
}

function updateKPIs(data, courseFilter) {
    const students = new Set(data.map(d => d.student)).size;
    const courses = new Set(data.map(d => d.course)).size;
    
    let logs = { inicio: 0, proceso: 0, logrado: 0 };
    data.forEach(d => {
        if (d.finalStatus === 1) logs.inicio++;
        else if (d.finalStatus === 2) logs.proceso++;
        else if (d.finalStatus === 3) logs.logrado++;
    });

    const totalEvals = logs.inicio + logs.proceso + logs.logrado;
    const pctLogrado = totalEvals > 0 ? Math.round((logs.logrado / totalEvals) * 100) : 0;

    document.getElementById('kpi-students').innerText = students;
    document.getElementById('kpi-courses').innerText = courses;
    document.getElementById('kpi-achievement').innerText = pctLogrado + '%';
    
    let coursesInAlert = 0;
    const groupedByCourse = groupBy(data, 'course');
    Object.keys(groupedByCourse).forEach(cName => {
        if (courseFilter !== 'all' && cName !== courseFilter) return;
        const cData = groupedByCourse[cName];
        let bad = 0;
        cData.forEach(d => {
            if (d.finalStatus === 1 || d.finalStatus === 2) bad++;
        });
        const badPct = (bad / cData.length) * 100;
        if (badPct >= 30) coursesInAlert++;
    });
    document.getElementById('kpi-alerts').innerText = coursesInAlert;
}

function updateCharts(data) {
    let logs = { inicio: 0, proceso: 0, logrado: 0 };
    data.forEach(d => {
        if (d.finalStatus === 1) logs.inicio++;
        else if (d.finalStatus === 2) logs.proceso++;
        else if (d.finalStatus === 3) logs.logrado++;
    });

    const ColorInicio = '#ef4444';
    const ColorProceso = '#f59e0b';
    const ColorLogrado = '#10b981';

    const ctxGlobal = document.getElementById('global-achievement-chart').getContext('2d');
    if (charts.global) charts.global.destroy();
    charts.global = new Chart(ctxGlobal, {
        type: 'doughnut',
        data: {
            labels: ['Inicio', 'Proceso', 'Logrado'],
            datasets: [{
                data: [logs.inicio, logs.proceso, logs.logrado],
                backgroundColor: [ColorInicio, ColorProceso, ColorLogrado]
            }]
        },
        options: { responsive: true, maintainAspectRatio: false }
    });

    const groupedByCourse = groupBy(data, 'course');
    const courseLabels = Object.keys(groupedByCourse);
    const logradoArr = [], procesoArr = [], inicioArr = [];

    courseLabels.forEach(c => {
        let i=0, p=0, l=0;
        groupedByCourse[c].forEach(d => {
            if (d.finalStatus === 1) i++;
            else if (d.finalStatus === 2) p++;
            else if (d.finalStatus === 3) l++;
        });
        inicioArr.push(i); procesoArr.push(p); logradoArr.push(l);
    });

    const ctxCourse = document.getElementById('course-performance-chart').getContext('2d');
    if (charts.course) charts.course.destroy();
    charts.course = new Chart(ctxCourse, {
        type: 'bar',
        data: {
            labels: courseLabels,
            datasets: [
                { label: 'Logrado', data: logradoArr, backgroundColor: ColorLogrado },
                { label: 'Proceso', data: procesoArr, backgroundColor: ColorProceso },
                { label: 'Inicio', data: inicioArr, backgroundColor: ColorInicio }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                x: { stacked: true },
                y: { stacked: true }
            }
        }
    });

    const groupedByTeacher = groupBy(data, 'teacher');
    const teacherLabels = Object.keys(groupedByTeacher).slice(0, 30);
    const tLogrado = [], tProceso = [], tInicio = [];
    
    teacherLabels.forEach(t => {
        let i=0, p=0, l=0;
        groupedByTeacher[t].forEach(d => {
            if (d.finalStatus === 1) i++;
            else if (d.finalStatus === 2) p++;
            else if (d.finalStatus === 3) l++;
        });
        tInicio.push(i); tProceso.push(p); tLogrado.push(l);
    });

    const ctxTeacher = document.getElementById('teacher-comparison-chart').getContext('2d');
    if (charts.teacher) charts.teacher.destroy();
    charts.teacher = new Chart(ctxTeacher, {
        type: 'bar',
        data: {
            labels: teacherLabels,
            datasets: [
                { label: 'Logrado', data: tLogrado, backgroundColor: ColorLogrado },
                { label: 'Proceso', data: tProceso, backgroundColor: ColorProceso },
                { label: 'Inicio', data: tInicio, backgroundColor: ColorInicio }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { title: { display: false } }
        }
    });

    let radarData = {};
    data.forEach(d => {
        if (!d.rawComps) return;
        Object.keys(d.rawComps).forEach(cId => {
            if (!radarData[cId]) radarData[cId] = { sum: 0, count: 0 };
            radarData[cId].sum += d.rawComps[cId];
            radarData[cId].count += 1;
        });
    });

    const compLabels = Object.keys(radarData).sort((a,b) => parseInt(a)-parseInt(b)).map(k => 'Comp ' + k);
    const compValues = Object.keys(radarData).sort((a,b) => parseInt(a)-parseInt(b)).map(k => (radarData[k].sum / radarData[k].count).toFixed(2));

    const ctxRadar = document.getElementById('competency-radar-chart').getContext('2d');
    if (charts.radar) charts.radar.destroy();
    charts.radar = new Chart(ctxRadar, {
        type: 'radar',
        data: {
            labels: compLabels,
            datasets: [{
                label: 'Promedio de Logro',
                data: compValues,
                fill: true,
                backgroundColor: 'rgba(14, 165, 233, 0.2)',
                borderColor: 'rgb(14, 165, 233)',
                pointBackgroundColor: 'rgb(14, 165, 233)',
                pointBorderColor: '#fff',
                pointHoverBackgroundColor: '#fff',
                pointHoverBorderColor: 'rgb(14, 165, 233)'
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                r: {
                    angleLines: { display: true },
                    min: 0,
                    max: 3,
                    ticks: {
                        stepSize: 1,
                        callback: function(value) {
                            if (value === 1) return 'Inicio';
                            if (value === 2) return 'Proceso';
                            if (value === 3) return 'Logrado';
                            return value;
                        }
                    }
                }
            }
        }
    });
}

function generateInsights(data, courseFilter) {
    const sList = document.getElementById('strengths-list');
    const wList = document.getElementById('weaknesses-list');
    sList.innerHTML = ''; wList.innerHTML = '';

    const coursesMap = groupBy(data, 'course');
    
    let strengthsCount = 0;
    let weaknessesCount = 0;

    Object.keys(coursesMap).forEach(cName => {
        const cData = coursesMap[cName];
        let bad = 0; let good = 0;
        cData.forEach(d => {
            if (d.finalStatus === 1 || d.finalStatus === 2) bad++;
            if (d.finalStatus === 3) good++;
        });
        const badPct = Math.round((bad / cData.length) * 100);
        const goodPct = Math.round((good / cData.length) * 100);

        if (badPct >= 30) {
            const li = document.createElement('li');
            li.innerHTML = `<strong>${cName}</strong> presenta una <strong>debilidad</strong>: ${badPct}% de evaluaciones están en Inicio o Proceso. ¡Alerta!`;
            wList.appendChild(li);
            weaknessesCount++;
        }
        
        if (goodPct >= 80) {
            const li = document.createElement('li');
            li.innerHTML = `<strong>${cName}</strong> es una <strong>fortaleza</strong>: ${goodPct}% de evaluaciones lograron el máximo nivel.`;
            sList.appendChild(li);
            strengthsCount++;
        }
    });

    if (strengthsCount === 0) sList.innerHTML = `<li>No se detectaron fortalezas sobresalientes en este cruce de filtros.</li>`;
    if (weaknessesCount === 0) wList.innerHTML = `<li>No hay alertas o debilidades relevantes (Umbral < 30%).</li>`;
}

function groupBy(xs, key) {
    return xs.reduce(function(rv, x) {
      if(x) {
        (rv[x[key]] = rv[x[key]] || []).push(x);
      }
      return rv;
    }, {});
}

// Export Module
function extractInsightsText(listId) {
    const list = document.getElementById(listId);
    if (!list) return [];
    return Array.from(list.querySelectorAll('li')).map(li => li.innerText);
}

function generateReport(type) {
    if (!currentFilteredData || currentFilteredData.length === 0) {
        alert('No hay datos para exportar.');
        return;
    }

    const stats = {
        students: document.getElementById('kpi-students').innerText,
        courses: document.getElementById('kpi-courses').innerText,
        pctLogrado: document.getElementById('kpi-achievement').innerText,
        alerts: document.getElementById('kpi-alerts').innerText,
        strengths: extractInsightsText('strengths-list'),
        weaknesses: extractInsightsText('weaknesses-list')
    };

    const reportName = `Reporte_Academico_${new Date().toISOString().split('T')[0]}`;

    try {
        if (type === 'ppt') generatePPT(reportName, stats);
        else if (type === 'word') generateWord(reportName, stats);
    } catch (e) {
        console.error("Error generando reporte:", e);
        alert("Ocurrió un error inicializando el exportador: " + e.message);
    }
}

function generatePPT(name, stats) {
    try {
        let pptx = new PptxGenJS();
        
        let slide1 = pptx.addSlide();
        slide1.addText("Reporte de Rendimiento Escolar", { x: 1, y: 1.5, w: "80%", fontSize: 32, bold: true, color: "0369a1", align: "center" });
        slide1.addText(`Emitido el: ${new Date().toLocaleDateString()}`, { x: 1, y: 2.5, w: "80%", fontSize: 16, color: "64748b", align: "center" });

        let slide2 = pptx.addSlide();
        slide2.addText("Métricas Generales", { x: 0.5, y: 0.3, fontSize: 24, bold: true, color: "0ea5e9" });
        slide2.addText(`Total Alumnos: ${stats.students}\nCursos Evaluados: ${stats.courses}\n% Logrado Global: ${stats.pctLogrado}\nAlertas Activas: ${stats.alerts}`, 
            { x: 0.5, y: 1.0, w: "90%", h: 2, fontSize: 18, bullet: true });

        // Fortalezas y Debilidades en la misma o separadas
        let slide3 = pptx.addSlide();
        slide3.addText("Diagnóstico Cualitativo", { x: 0.5, y: 0.3, fontSize: 24, bold: true, color: "334155" });
        slide3.addText("Fortalezas:", { x: 0.5, y: 1.0, fontSize: 16, bold: true, color: "10b981" });
        slide3.addText(stats.strengths.join('\n') || "No hay fortalezas registradas.", { x: 0.5, y: 1.4, w: "90%", h: 1.5, fontSize: 12, bullet: true });
        
        slide3.addText("Oportunidades de Mejora:", { x: 0.5, y: 3.0, fontSize: 16, bold: true, color: "ef4444" });
        slide3.addText(stats.weaknesses.join('\n') || "No hay riesgos críticos bajo estos filtros.", { x: 0.5, y: 3.4, w: "90%", h: 1.5, fontSize: 12, bullet: true });

        // Acciones recomendadas
        let recomendaciones = [];
        if (stats.weaknesses.length > 0) {
            recomendaciones = [
                "Implementar tutorías extracurriculares inmediatas en los cursos en estado de Alerta.",
                "Revisar y adaptar las estrategias pedagógicas para alumnos en nivel 'Inicio'.",
                "Establecer reuniones con los docentes de cursos afectados para planificar la recuperación académica."
            ];
        } else {
            recomendaciones = [
                "Fomentar sesiones de enriquecimiento para alumnos que están en nivel 'Logrado'.",
                "Documentar y compartir las buenas prácticas aplicadas por los docentes con alto rendimiento general.",
                "Mantener el monitoreo continuo para prevenir descensos en trimestres futuros."
            ];
        }

        let slide4 = pptx.addSlide();
        slide4.addText("Acciones Recomendadas", { x: 0.5, y: 0.3, fontSize: 24, bold: true, color: "f59e0b" });
        slide4.addText(recomendaciones.join('\n\n'), { x: 0.5, y: 1.2, w: "90%", h: 3, fontSize: 16, bullet: true });

        // Desglose por Grados y Secciones
        const gradeGroups = groupBy(currentFilteredData, 'grade');
        const grades = Object.keys(gradeGroups).sort();
        
        grades.forEach(g => {
            let gradeSlide = pptx.addSlide();
            gradeSlide.addText(`Desglose: Grado ${g || 'No Definido'}`, { x: 0.5, y: 0.3, fontSize: 24, bold: true, color: "0369a1" });
            
            const sectionsGroup = groupBy(gradeGroups[g], 'section');
            const sections = Object.keys(sectionsGroup).sort();
            
            let tableData = [
                [
                    { text:"Sección", options:{ bold:true, fill:"f1f5f9" } },
                    { text:"Alumnos", options:{ bold:true, fill:"f1f5f9" } },
                    { text:"Logrado (Verde)", options:{ bold:true, fill:"f1f5f9", color:"10b981" } },
                    { text:"Proceso (Amar.)", options:{ bold:true, fill:"f1f5f9", color:"f59e0b" } },
                    { text:"Inicio (Rojo)", options:{ bold:true, fill:"f1f5f9", color:"ef4444" } },
                    { text:"% Logro", options:{ bold:true, fill:"f1f5f9" } }
                ]
            ];

            sections.forEach(s => {
                const studentsData = sectionsGroup[s];
                const uniqueStudents = new Set(studentsData.map(d => d.student)).size;
                let logrados = 0, procesos = 0, inicios = 0;
                
                studentsData.forEach(d => {
                    if (d.finalStatus === 1) inicios++;
                    else if (d.finalStatus === 2) procesos++;
                    else if (d.finalStatus === 3) logrados++;
                });
                
                const totalEvals = logrados + procesos + inicios;
                const pctLog = totalEvals > 0 ? Math.round((logrados / totalEvals) * 100) : 0;
                
                tableData.push([
                    s || 'N/D',
                    uniqueStudents.toString(),
                    logrados.toString(),
                    procesos.toString(),
                    inicios.toString(),
                    pctLog.toString() + '%'
                ]);
            });

            gradeSlide.addTable(tableData, { 
                x: 0.5, y: 1.0, w: "90%", 
                colW: [1.2, 1.2, 1.5, 1.5, 1.5, 1.2],
                border: {pt: 1, color: "cbd5e1"},
                align: "center", fontSize: 12 
            });
        });

        pptx.writeFile({ fileName: `${name}.pptx` }).then(() => {
            console.log("PPTX Exportado exitosamente");
        }).catch(err => {
            alert("Error al guardar el PPTX: " + err);
        });
    } catch(err) {
        alert("Error construyendo diapositivas PPTX: " + err.message);
    }
}

function generateWord(name, stats) {
    let recomendaciones = [];
    if (stats.weaknesses.length > 0) {
        recomendaciones = [
            "Implementar tutorías extracurriculares inmediatas en los cursos en estado de Alerta.",
            "Revisar y adaptar las estrategias pedagógicas para alumnos en nivel 'Inicio'.",
            "Establecer reuniones con los docentes de cursos afectados para planificar la recuperación académica."
        ];
    } else {
        recomendaciones = [
            "Fomentar sesiones de enriquecimiento para alumnos que están en nivel 'Logrado'.",
            "Documentar y compartir las buenas prácticas aplicadas por los docentes con alto rendimiento general.",
            "Mantener el monitoreo continuo para prevenir descensos en trimestres futuros."
        ];
    }

    const recosHtml = recomendaciones.map(r => `<li>${r}</li>`).join('');

    let content = `\uFEFF
    <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
    <head><meta charset='utf-8'><title>Reporte Académico</title></head>
    <body style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;">
        <h1 style="color: #0369a1; text-align: center;">Informe Académico Detallado</h1>
        <p style="text-align: center;"><strong>Fecha de emisión:</strong> ${new Date().toLocaleDateString()}</p>
        <hr>
        <h2>1. Resumen General</h2>
        <ul>
            <li><strong>Total Alumnos (Muestra):</strong> ${stats.students}</li>
            <li><strong>Cursos Evaluados:</strong> ${stats.courses}</li>
            <li><strong>Porcentaje Logrado Global:</strong> ${stats.pctLogrado}</li>
            <li><strong>Cursos en Alerta:</strong> ${stats.alerts}</li>
        </ul>
        
        <h2>2. Diagnóstico Cualitativo</h2>
        <h3 style="color: #166534;">Fortalezas</h3>
        <ul>${stats.strengths.map(s => `<li>${s}</li>`).join('') || '<li>No hay fortalezas sobresalientes registradas.</li>'}</ul>
        <h3 style="color: #991b1b;">Oportunidades de Mejora (Riesgos)</h3>
        <ul>${stats.weaknesses.map(w => `<li>${w}</li>`).join('') || '<li>No hay riesgos críticos bajo estos filtros.</li>'}</ul>
        
        <h2>3. Acciones Formativas Recomendadas</h2>
        <ul>${recosHtml}</ul>

        <h2>4. Desglose Detallado por Grado y Sección</h2>
    `;
    
    const gradeGroups = groupBy(currentFilteredData, 'grade');
    const grades = Object.keys(gradeGroups).sort();
    
    if (grades.length === 0) {
        content += `<p>No hay datos disponibles para mostrar el desglose.</p>`;
    }

    grades.forEach(g => {
        const sectionsGroup = groupBy(gradeGroups[g], 'section');
        const sections = Object.keys(sectionsGroup).sort();
        
        content += `<h3 style="color: #475569; margin-top: 30px;">▶ GRADO: ${g || 'No Definido'}</h3>`;
        
        sections.forEach(s => {
            const studentsData = sectionsGroup[s];
            const uniqueStudents = new Set(studentsData.map(d => d.student)).size;
            let logrados = 0, procesos = 0, inicios = 0;
            
            studentsData.forEach(d => {
                if (d.finalStatus === 1) inicios++;
                else if (d.finalStatus === 2) procesos++;
                else if (d.finalStatus === 3) logrados++;
            });
            
            const totalEvals = logrados + procesos + inicios;
            const pctLog = totalEvals > 0 ? Math.round((logrados / totalEvals) * 100) : 0;
            
            content += `
                <div style="margin-left: 20px; margin-bottom: 25px;">
                    <h4 style="color: #334155; margin-bottom: 5px;">Sección: ${s || 'No Definida'}</h4>
                    <table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse; width: 100%; max-width: 600px; text-align: center;">
                        <tr style="background-color: #f1f5f9;">
                            <th>Alumnos únicos</th>
                            <th>Total Notas</th>
                            <th>Logrado (Verde)</th>
                            <th>Proceso (Amarillo)</th>
                            <th>Inicio (Rojo)</th>
                            <th>% Logrado</th>
                        </tr>
                        <tr>
                            <td>${uniqueStudents}</td>
                            <td>${totalEvals}</td>
                            <td style="color: #10b981; font-weight: bold;">${logrados}</td>
                            <td style="color: #f59e0b; font-weight: bold;">${procesos}</td>
                            <td style="color: #ef4444; font-weight: bold;">${inicios}</td>
                            <td><strong>${pctLog}%</strong></td>
                        </tr>
                    </table>
                </div>
            `;
        });
    });

    content += `
    </body>
    </html>
    `;

    const blob = new Blob([content], { type: 'application/msword;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `${name}.doc`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}
