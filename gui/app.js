/* ============================================================================
   gui/app.js — Controlador de Interfaz de Usuario para APP_REVISA_EXCEL_V2
   ============================================================================ */

document.addEventListener("DOMContentLoaded", () => {
    // Referencias a elementos
    const htmlEl = document.documentElement;
    const btnThemeToggle = document.getElementById("btn-theme-toggle");
    const themeIcon = document.getElementById("theme-icon");
    const btnExit = document.getElementById("btn-exit");

    // Selector de modo
    const btnModoA = document.getElementById("btn-modo-a");
    const btnModoB = document.getElementById("btn-modo-b");

    // Inputs y botones de selección
    const inputPlantilla = document.getElementById("plantilla-path");
    const inputTrabajos = document.getElementById("trabajos-path");
    const plantillaHint = document.getElementById("plantilla-hint");
    const trabajosHint = document.getElementById("trabajos-hint");
    const btnReloadRubric = document.getElementById("btn-reload-rubric");
    const btnBrowseFile = document.getElementById("btn-browse-file");
    const btnBrowseFolder = document.getElementById("btn-browse-folder");
    const btnGenTemplate = document.getElementById("btn-gen-template");

    const inputSeccion = document.getElementById("input-seccion");
    const inputFecha = document.getElementById("input-fecha");

    const shutdownModal = document.getElementById("shutdown-modal");
    const btnCloseWindow = document.getElementById("btn-close-window");

    // Pre-chequeo de rúbrica
    const rubricContainer = document.getElementById("rubric-preview-container");
    const rubricSourceBadge = document.getElementById("rubric-source-badge");
    const rubricTotalBadge = document.getElementById("rubric-total-badge");
    const rubricTableBody = document.getElementById("rubric-table-body");
    const rubricAlert = document.getElementById("rubric-status-alert");

    // Consola y Ejecución
    const btnRun = document.getElementById("btn-run");
    const btnStop = document.getElementById("btn-stop");
    const consoleSection = document.getElementById("console-section");
    const consoleStatusText = document.getElementById("console-status-text");
    const btnClearConsole = document.getElementById("btn-clear-console");
    const terminalBody = document.getElementById("terminal-body");

    // Progreso
    const progressFill = document.getElementById("progress-fill");
    const progressPctText = document.getElementById("progress-pct-text");
    const progressStudentText = document.getElementById("progress-student-text");

    // Resumen Final
    const resultsSummary = document.getElementById("results-summary");
    const summaryStatsGrid = document.getElementById("summary-stats-grid");
    const resultsTableBody = document.getElementById("results-table-body");
    const toastContainer = document.getElementById("toast-container");

    let selectedMode = "modo_b_rubrica";
    let isRunning = false;
    let eventSource = null;
    let collectedResults = [];

    // -------------------------------------------------------------------------
    // 1. Control de Tema (Dark / Light)
    // -------------------------------------------------------------------------
    const savedTheme = localStorage.getItem("theme") || "dark";
    htmlEl.setAttribute("data-theme", savedTheme);
    themeIcon.textContent = savedTheme === "dark" ? "🌙" : "☀️";

    btnThemeToggle.addEventListener("click", () => {
        const currentTheme = htmlEl.getAttribute("data-theme");
        const nextTheme = currentTheme === "dark" ? "light" : "dark";
        htmlEl.setAttribute("data-theme", nextTheme);
        localStorage.setItem("theme", nextTheme);
        themeIcon.textContent = nextTheme === "dark" ? "🌙" : "☀️";
    });

    // -------------------------------------------------------------------------
    // 2. Control de Selector de Modo A / Modo B
    // -------------------------------------------------------------------------
    btnModoA.addEventListener("click", () => {
        if (isRunning) return;
        selectedMode = "modo_a_espejo";
        btnModoA.classList.add("active");
        btnModoA.setAttribute("aria-checked", "true");
        btnModoB.classList.remove("active");
        btnModoB.setAttribute("aria-checked", "false");
        rubricContainer.style.display = "none";
        appendTerminalLine("Modo seleccionado: Modo A (Plantilla Espejo - Strict Diff).", "system");
    });

    btnModoB.addEventListener("click", () => {
        if (isRunning) return;
        selectedMode = "modo_b_rubrica";
        btnModoB.classList.add("active");
        btnModoB.setAttribute("aria-checked", "true");
        btnModoA.classList.remove("active");
        btnModoA.setAttribute("aria-checked", "false");
        appendTerminalLine("Modo seleccionado: Modo B (Rúbrica Inteligente por Criterios).", "system");

        if (inputPlantilla.value.trim()) {
            inspectRubric(inputPlantilla.value.trim());
        }
    });

    // -------------------------------------------------------------------------
    // 3. Helper para TextBoxes de Rutas (Bajo tamaño, auto-scroll y visualización de archivo)
    // -------------------------------------------------------------------------
    function setPathInput(inputEl, path, hintEl, defaultHintHtml) {
        if (!inputEl) return;
        inputEl.value = path || "";
        inputEl.title = path || "";
        if (path) {
            // Desplazar al final para que el nombre del archivo o carpeta sea visible
            setTimeout(() => {
                inputEl.scrollLeft = inputEl.scrollWidth;
            }, 60);

            if (hintEl) {
                const clean = path.replace(/\\/g, "/");
                const parts = clean.split("/").filter(Boolean);
                const baseName = parts.pop() || "";
                if (baseName) {
                    hintEl.innerHTML = `<span class="path-badge">📌 Seleccionado:</span> <strong>${escapeHtml(baseName)}</strong>`;
                }
            }
        } else if (hintEl && defaultHintHtml) {
            hintEl.innerHTML = defaultHintHtml;
        }
    }

    inputPlantilla.addEventListener("input", () => {
        setPathInput(inputPlantilla, inputPlantilla.value, plantillaHint);
        if (selectedMode === "modo_b_rubrica" && inputPlantilla.value.trim()) {
            inspectRubric(inputPlantilla.value.trim());
        }
    });

    inputTrabajos.addEventListener("input", () => {
        setPathInput(inputTrabajos, inputTrabajos.value, trabajosHint);
    });

    // -------------------------------------------------------------------------
    // 4. Cargar Valores por Defecto
    // -------------------------------------------------------------------------
    fetch("/api/defaults")
        .then(res => res.json())
        .then(data => {
            if (data.plantilla) setPathInput(inputPlantilla, data.plantilla, plantillaHint, 'Busca hoja <code>_RUBRICA</code> o archivo <code>.rubrica.json</code>');
            if (data.trabajos) setPathInput(inputTrabajos, data.trabajos, trabajosHint, 'Carpeta con archivos <code>.xlsx</code> a evaluar');

            // Fecha actual por defecto dd/mm/aaaa
            const now = new Date();
            const dd = String(now.getDate()).padStart(2, "0");
            const mm = String(now.getMonth() + 1).padStart(2, "0");
            const yyyy = now.getFullYear();
            inputFecha.value = `${dd}/${mm}/${yyyy}`;

            if (data.plantilla && selectedMode === "modo_b_rubrica") {
                inspectRubric(data.plantilla);
            }
        })
        .catch(err => {
            console.warn("No se pudieron cargar rutas por defecto:", err);
        });

    // -------------------------------------------------------------------------
    // 5. Pre-chequeo e Inspección de Rúbrica
    // -------------------------------------------------------------------------
    function inspectRubric(plantillaPath) {
        if (!plantillaPath) return;

        fetch(`/api/inspect-rubric?plantilla=${encodeURIComponent(plantillaPath)}`)
            .then(res => res.json())
            .then(data => {
                if (data.status === "ok" && data.criterios && data.criterios.length > 0) {
                    renderRubricPreview(data);
                } else {
                    rubricContainer.style.display = "none";
                    if (selectedMode === "modo_b_rubrica") {
                        showToast(data.message || "No se detectó rúbrica en la plantilla.", "info");
                    }
                }
            })
            .catch(err => {
                console.error("Error inspeccionando rúbrica:", err);
            });
    }

    function renderRubricPreview(data) {
        rubricTableBody.innerHTML = "";
        rubricSourceBadge.textContent = data.origen || "Rúbrica";

        const ptsTotales = data.puntos_totales || 0;
        rubricTotalBadge.textContent = `${ptsTotales.toFixed(1)} / 100 pts`;

        const esExacto100 = Math.abs(ptsTotales - 100.0) < 0.01;
        if (esExacto100) {
            rubricTotalBadge.className = "points-badge valid";
            rubricAlert.style.display = "none";
        } else {
            rubricTotalBadge.className = "points-badge invalid";
            rubricAlert.textContent = `⚠️ Atención: Los criterios suman ${ptsTotales.toFixed(1)} pts en lugar de 100.0 pts.`;
            rubricAlert.style.display = "block";
        }

        data.criterios.forEach(c => {
            const tr = document.createElement("tr");
            tr.innerHTML = `
                <td><strong>${escapeHtml(c.id)}</strong></td>
                <td>${escapeHtml(c.criterio)}</td>
                <td><span class="field-hint"><code>${escapeHtml(c.hoja || "(todas)")}</code></span></td>
                <td><code>${escapeHtml(c.rango || "-")}</code></td>
                <td>${escapeHtml(c.tipo_validacion)}</td>
                <td><strong>${c.puntos}</strong></td>
                <td>${c.obligatorio ? "🔴 SI" : "NO"}</td>
            `;
            rubricTableBody.appendChild(tr);
        });

        rubricContainer.style.display = "block";
    }

    // -------------------------------------------------------------------------
    // 6. Diálogos de Selección de Archivos y Generación de Plantilla
    // -------------------------------------------------------------------------
    const fileBtnOriginalSvg = btnBrowseFile.innerHTML;
    const folderBtnOriginalSvg = btnBrowseFolder.innerHTML;

    btnBrowseFile.addEventListener("click", () => {
        if (isRunning) return;
        btnBrowseFile.disabled = true;
        btnBrowseFile.innerHTML = `<span style="font-size:0.9rem;">⏳</span>`;

        fetch("/api/select-file")
            .then(res => res.json())
            .then(data => {
                if (data.path) {
                    setPathInput(inputPlantilla, data.path, plantillaHint);
                    showToast("Plantilla seleccionada.", "info");
                    if (selectedMode === "modo_b_rubrica") {
                        inspectRubric(data.path);
                    }
                }
            })
            .finally(() => {
                btnBrowseFile.disabled = false;
                btnBrowseFile.innerHTML = fileBtnOriginalSvg;
            });
    });

    btnBrowseFolder.addEventListener("click", () => {
        if (isRunning) return;
        btnBrowseFolder.disabled = true;
        btnBrowseFolder.innerHTML = `<span style="font-size:0.9rem;">⏳</span>`;

        fetch("/api/select-folder")
            .then(res => res.json())
            .then(data => {
                if (data.path) {
                    setPathInput(inputTrabajos, data.path, trabajosHint);
                    showToast("Carpeta de trabajos seleccionada.", "info");

                    // Sugerir sección
                    if (!inputSeccion.value.trim()) {
                        const parts = data.path.split(/[\\/]/);
                        const folderName = parts.pop() || parts.pop();
                        if (folderName && folderName !== "TRABAJOS_ESTUDIANTES") {
                            inputSeccion.value = folderName;
                        }
                    }
                }
            })
            .finally(() => {
                btnBrowseFolder.disabled = false;
                btnBrowseFolder.innerHTML = folderBtnOriginalSvg;
            });
    });

    // Botón Recargar Rúbrica tras editar en Excel (Ctrl+S)
    if (btnReloadRubric) {
        btnReloadRubric.addEventListener("click", () => {
            const path = inputPlantilla.value.trim();
            if (!path) {
                showToast("Selecciona primero un archivo de plantilla.", "warning");
                return;
            }
            btnReloadRubric.classList.add("spinning");
            inspectRubric(path);
            showToast("Actualizando criterios desde Excel...", "info");
            setTimeout(() => {
                btnReloadRubric.classList.remove("spinning");
            }, 700);
        });
    }

    btnGenTemplate.addEventListener("click", () => {
        if (isRunning) return;

        const confirmar = confirm(
            "¿Deseas crear una nueva plantilla de demostración en la carpeta PLANTILLAS/?\n\n" +
            "• Esto creará un nuevo archivo con hojas 'Ventas', 'Resumen' y '_RUBRICA'.\n" +
            "• Tus archivos existentes NO serán sobreescritos."
        );
        if (!confirmar) return;

        btnGenTemplate.disabled = true;
        btnGenTemplate.innerHTML = `<span>⏳</span> Creando...`;

        fetch("/api/generate-rubric-template")
            .then(res => res.json())
            .then(data => {
                if (data.status === "ok") {
                    // Si estamos en Modo A, cambiar automáticamente a Modo B para ver la rúbrica
                    if (selectedMode !== "modo_b_rubrica") {
                        btnModoB.click();
                    }
                    setPathInput(inputPlantilla, data.path, plantillaHint);
                    showToast("¡Plantilla de demostración creada exitosamente!", "success");
                    appendTerminalLine(`Plantilla generada con éxito: ${data.path}. Contiene hoja _RUBRICA con 8 criterios de ejemplo.`, "success");
                    inspectRubric(data.path);
                } else {
                    showToast(data.message || "Error al crear plantilla.", "error");
                }
            })
            .catch(err => {
                showToast("Error de conexión al generar plantilla.", "error");
            })
            .finally(() => {
                btnGenTemplate.disabled = false;
                btnGenTemplate.innerHTML = `<span>📄</span> Crear Plantilla de Demostración (.xlsx)`;
            });
    });

    // -------------------------------------------------------------------------
    // 6. Ejecución de la Auditoría en Tiempo Real (SSE)
    // -------------------------------------------------------------------------
    btnRun.addEventListener("click", () => {
        if (isRunning) return;

        const plantilla = inputPlantilla.value.trim();
        const trabajos = inputTrabajos.value.trim();
        const seccion = inputSeccion.value.trim() || "SEC_DEFAULT";
        const fecha = inputFecha.value.trim();

        if (!plantilla) {
            showToast("Por favor selecciona una plantilla de Excel.", "error");
            return;
        }
        if (!trabajos) {
            showToast("Por favor selecciona la carpeta de trabajos.", "error");
            return;
        }

        // Configurar estado
        isRunning = true;
        collectedResults = [];
        btnRun.disabled = true;
        btnRun.style.display = "none";
        btnStop.style.display = "inline-flex";

        consoleSection.style.display = "block";
        terminalBody.innerHTML = "";
        resultsSummary.style.display = "none";
        resultsTableBody.innerHTML = "";

        updateProgress(0, "Iniciando proceso de auditoría...");
        consoleStatusText.textContent = "Ejecutando auditoría...";

        const params = new URLSearchParams({
            plantilla: plantilla,
            trabajos: trabajos,
            modo: selectedMode,
            seccion: seccion,
            fecha: fecha
        });

        eventSource = new EventSource(`/api/run-audit?${params.toString()}`);

        eventSource.onmessage = (event) => {
            try {
                const data = JSON.parse(event.data);

                if (data.evento === "inicio_archivo") {
                    updateProgress(data.porcentaje, `Evaluando [${data.indice}/${data.total}]: ${data.archivo}`);
                } else if (data.evento === "fin_archivo") {
                    updateProgress(data.porcentaje, `Finalizado [${data.indice}/${data.total}]: ${data.archivo}`);
                    if (data.resultado) {
                        collectedResults.push(data.resultado);
                    }
                } else if (data.status === "done") {
                    finishExecution(data.code === 0);
                    eventSource.close();
                } else if (data.status === "error") {
                    appendTerminalLine(`Error: ${data.message}`, "error");
                    finishExecution(false);
                    eventSource.close();
                } else if (data.text) {
                    // Determinar tipo de línea
                    const t = data.text;
                    let type = "info";
                    if (t.includes("❌") || t.toLowerCase().includes("error")) type = "error";
                    else if (t.includes("✅") || t.includes("ÉXITO") || t.includes("COMPLETADA")) type = "success";
                    else if (t.includes("⚠️") || t.toLowerCase().includes("advertencia")) type = "warning";
                    else if (t.startsWith("===") || t.includes("INICIANDO")) type = "system";

                    appendTerminalLine(t, type);
                }
            } catch (e) {
                console.error("Error parseando SSE:", e);
            }
        };

        eventSource.onerror = (err) => {
            console.error("Error SSE EventSource:", err);
            appendTerminalLine("Conexión con el servidor interrumpida o finalizada.", "warning");
            finishExecution(true);
            if (eventSource) eventSource.close();
        };
    });

    btnStop.addEventListener("click", () => {
        if (!isRunning) return;
        fetch("/api/abort")
            .then(() => {
                showToast("Auditoría cancelada.", "info");
                appendTerminalLine("Proceso cancelado por el usuario.", "warning");
                finishExecution(false);
                if (eventSource) eventSource.close();
            });
    });

    btnClearConsole.addEventListener("click", () => {
        terminalBody.innerHTML = "";
    });

    if (btnCloseWindow) {
        btnCloseWindow.addEventListener("click", () => {
            window.close();
        });
    }

    btnExit.addEventListener("click", () => {
        if (confirm("¿Estás seguro de que deseas apagar el servidor local?")) {
            // Mostrar inmediatamente la pantalla de apagado con contraste total
            if (shutdownModal) {
                shutdownModal.style.display = "flex";
            }
            fetch("/api/shutdown")
                .catch(() => {
                    // El servidor finalizó, esto es el comportamiento esperado
                });
        }
    });

    // -------------------------------------------------------------------------
    // 7. Funciones Auxiliares
    // -------------------------------------------------------------------------
    function updateProgress(pct, text) {
        const p = Math.max(0, Math.min(100, pct));
        progressFill.style.width = `${p}%`;
        progressPctText.textContent = `${p.toFixed(0)}%`;
        if (text) progressStudentText.textContent = text;
    }

    function appendTerminalLine(text, type = "info") {
        const line = document.createElement("div");
        line.className = `terminal-line ${type}`;
        line.textContent = text;
        terminalBody.appendChild(line);
        terminalBody.scrollTop = terminalBody.scrollHeight;
    }

    function finishExecution(isSuccess) {
        isRunning = false;
        btnRun.disabled = false;
        btnRun.style.display = "inline-flex";
        btnStop.style.display = "none";
        consoleStatusText.textContent = isSuccess ? "Auditoría completada" : "Auditoría finalizada con observaciones";

        updateProgress(100, "Auditoría concluida.");

        if (collectedResults.length > 0) {
            renderFinalResults(collectedResults);
        }
    }

    function renderFinalResults(results) {
        const total = results.length;
        const aprobados = results.filter(r => r.aprobado).length;
        const reprobados = total - aprobados;
        const promedio = total > 0 ? (results.reduce((acc, r) => acc + (r.nota_final_100 || 0), 0) / total) : 0;

        summaryStatsGrid.innerHTML = `
            <div class="stat-box">
                <div class="stat-num">${total}</div>
                <div class="stat-label">Total Evaluados</div>
            </div>
            <div class="stat-box">
                <div class="stat-num" style="color:var(--success-color);">${aprobados}</div>
                <div class="stat-label">Aprobados</div>
            </div>
            <div class="stat-box">
                <div class="stat-num" style="color:var(--danger-color);">${reprobados}</div>
                <div class="stat-label">Reprobados</div>
            </div>
            <div class="stat-box">
                <div class="stat-num" style="color:var(--accent-color);">${promedio.toFixed(1)}</div>
                <div class="stat-label">Promedio General</div>
            </div>
        `;

        resultsTableBody.innerHTML = "";
        results.forEach(r => {
            const tr = document.createElement("tr");
            const estadoClase = r.aprobado ? "aprobado" : "reprobado";
            const estadoTexto = r.aprobado ? "Aprobado" : "Reprobado";
            const fallosCount = (r.criterios_fallados || []).length;
            const detTxt = fallosCount === 0 ? "Sin errores" : `${fallosCount} criterio(s) con fallos`;

            tr.innerHTML = `
                <td><strong>${escapeHtml(r.nombre_estudiante)}</strong></td>
                <td><strong>${(r.nota_final_100 || 0).toFixed(1)}</strong> / 100</td>
                <td><span class="status-pill ${estadoClase}">${estadoTexto}</span></td>
                <td><span class="field-hint">${detTxt}</span></td>
            `;
            resultsTableBody.appendChild(tr);
        });

        resultsSummary.style.display = "block";
    }

    function showToast(message, type = "info") {
        const toast = document.createElement("div");
        toast.className = `toast ${type}`;
        let icon = "ℹ️";
        if (type === "success") icon = "✅";
        else if (type === "error") icon = "❌";
        else if (type === "warning") icon = "⚠️";

        toast.innerHTML = `<span class="toast-icon">${icon}</span> <span>${escapeHtml(message)}</span>`;
        toastContainer.appendChild(toast);
        setTimeout(() => {
            toast.style.opacity = "0";
            toast.style.transform = "translateY(10px)";
            toast.style.transition = "all 0.3s ease";
            setTimeout(() => toast.remove(), 300);
        }, 4500);
    }

    function escapeHtml(str) {
        if (!str) return "";
        return String(str)
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&#039;");
    }
});
