/**
 * CobranzApp — Frontend Controller & Executive Dashboard Logic
 */

let appMetadata = null;
let chartInstance = null;

// ==============================================================
// SISTEMA ELEGANTE DE NOTIFICACIONES TOAST (CobranzApp Theme)
// ==============================================================
function showToast(type, title, message, duration = 4500) {
  const container = document.getElementById('toast-container');
  if (!container) return;

  const toastId = 'toast-' + Date.now() + '-' + Math.floor(Math.random() * 1000);

  const configs = {
    success: {
      border: 'border-emerald-200/90',
      iconBg: 'bg-emerald-50 text-emerald-600 border border-emerald-200',
      icon: 'fa-solid fa-circle-check',
      progressBg: 'bg-emerald-500'
    },
    error: {
      border: 'border-rose-200/90',
      iconBg: 'bg-rose-50 text-rose-600 border border-rose-200',
      icon: 'fa-solid fa-circle-xmark',
      progressBg: 'bg-rose-500'
    },
    warning: {
      border: 'border-amber-200/90',
      iconBg: 'bg-amber-50 text-amber-600 border border-amber-200',
      icon: 'fa-solid fa-triangle-exclamation',
      progressBg: 'bg-amber-500'
    },
    info: {
      border: 'border-blue-200/90',
      iconBg: 'bg-blue-50 text-blue-600 border border-blue-200',
      icon: 'fa-solid fa-circle-info',
      progressBg: 'bg-blue-500'
    }
  };

  const cfg = configs[type] || configs.info;

  const toast = document.createElement('div');
  toast.id = toastId;
  toast.className = `pointer-events-auto bg-white/95 backdrop-blur-sm border ${cfg.border} rounded-2xl shadow-xl shadow-slate-900/10 p-4 transition-all duration-300 ease-out transform translate-x-8 opacity-0 relative overflow-hidden flex items-start gap-3.5`;

  toast.innerHTML = `
    <div class="w-9 h-9 rounded-xl ${cfg.iconBg} flex items-center justify-center text-base shrink-0 shadow-2xs">
      <i class="${cfg.icon}"></i>
    </div>
    <div class="flex-1 min-w-0 pr-1">
      <div class="text-xs font-bold text-slate-900 leading-snug">${title}</div>
      <div class="text-xs text-slate-600 mt-0.5 leading-relaxed break-words">${message}</div>
    </div>
    <button onclick="dismissToast('${toastId}')" class="text-slate-400 hover:text-slate-600 transition-colors p-1 cursor-pointer shrink-0">
      <i class="fa-solid fa-xmark text-sm"></i>
    </button>
    <div class="absolute bottom-0 left-0 h-1 bg-slate-100 rounded-full w-full overflow-hidden">
      <div id="${toastId}-bar" class="h-full ${cfg.progressBg} rounded-full transition-all linear" style="width: 100%;"></div>
    </div>
  `;

  container.appendChild(toast);

  requestAnimationFrame(() => {
    toast.classList.remove('translate-x-8', 'opacity-0');
    toast.classList.add('translate-x-0', 'opacity-100');
  });

  const startTime = Date.now();
  const bar = document.getElementById(`${toastId}-bar`);
  const interval = setInterval(() => {
    const elapsed = Date.now() - startTime;
    const remainingPct = Math.max(0, 100 - (elapsed / duration) * 100);
    if (bar) bar.style.width = `${remainingPct}%`;
    if (elapsed >= duration) {
      clearInterval(interval);
      dismissToast(toastId);
    }
  }, 40);

  toast._interval = interval;
}

function dismissToast(id) {
  const toast = document.getElementById(id);
  if (!toast) return;
  if (toast._interval) clearInterval(toast._interval);
  toast.classList.add('translate-x-8', 'opacity-0');
  setTimeout(() => {
    if (toast.parentNode) toast.remove();
  }, 300);
}

// Interceptor global para que ningún alert() nativo vuelva a interrumpir la interfaz
window.alert = function(msg) {
  showToast('info', 'Aviso del Sistema', msg);
};

// ==============================================================
// CONTROLADOR DE BARRA DE PROGRESO DE CARGA Y PROCESAMIENTO
// ==============================================================
const ProgressController = {
  interval: null,
  start(title, initialStatus) {
    const box = document.getElementById('upload-progress-box');
    const bar = document.getElementById('progress-bar-fill');
    const percLabel = document.getElementById('progress-percentage-label');
    const titleElem = document.getElementById('progress-title');
    const descElem = document.getElementById('progress-status-desc');
    const icon = document.getElementById('progress-spinner-icon');

    if (!box) return;
    box.classList.remove('hidden');
    titleElem.textContent = title || 'Procesando archivo...';
    descElem.textContent = initialStatus || 'Iniciando carga de datos...';
    icon.className = 'fa-solid fa-spinner fa-spin text-blue-400';
    bar.className = 'bg-gradient-to-r from-blue-500 via-indigo-500 to-emerald-400 h-2 rounded-full transition-all duration-300 ease-out';
    percLabel.className = 'text-xs font-mono font-bold px-3 py-1 rounded-full bg-blue-500/20 text-blue-400 border border-blue-500/30';
    bar.style.width = '12%';
    percLabel.textContent = '12%';

    let current = 12;
    clearInterval(this.interval);
    this.interval = setInterval(() => {
      if (current < 35) {
        current += 6;
        descElem.textContent = 'Transfiriendo bytes al servidor...';
      } else if (current < 70) {
        current += 4;
        descElem.textContent = 'Procesando hojas de cálculo con motor Calamine...';
      } else if (current < 88) {
        current += 2;
        descElem.textContent = 'Estandarizando cédulas y unificando cartera con contactos...';
      } else if (current < 95) {
        current += 0.5;
        descElem.textContent = 'Generando métricas ejecutivas y matriz de campañas...';
      }
      bar.style.width = `${Math.min(current, 95)}%`;
      percLabel.textContent = `${Math.round(Math.min(current, 95))}%`;
    }, 220);
  },
  complete(successMessage, callback) {
    clearInterval(this.interval);
    const box = document.getElementById('upload-progress-box');
    const bar = document.getElementById('progress-bar-fill');
    const percLabel = document.getElementById('progress-percentage-label');
    const descElem = document.getElementById('progress-status-desc');
    const icon = document.getElementById('progress-spinner-icon');

    if (!box) {
      if (callback) callback();
      return;
    }

    bar.style.width = '100%';
    percLabel.textContent = '100%';
    descElem.textContent = successMessage || '¡Información cargada con éxito!';
    icon.className = 'fa-solid fa-circle-check text-emerald-400';

    setTimeout(() => {
      box.classList.add('hidden');
      bar.style.width = '0%';
      if (callback) callback();
    }, 650);
  },
  error(errorMessage) {
    clearInterval(this.interval);
    const box = document.getElementById('upload-progress-box');
    const bar = document.getElementById('progress-bar-fill');
    const percLabel = document.getElementById('progress-percentage-label');
    const descElem = document.getElementById('progress-status-desc');
    const icon = document.getElementById('progress-spinner-icon');

    if (!box) return;
    bar.className = 'bg-rose-500 h-2 rounded-full transition-all duration-300';
    percLabel.className = 'text-xs font-mono font-bold px-3 py-1 rounded-full bg-rose-500/20 text-rose-400 border border-rose-500/30';
    descElem.textContent = errorMessage || 'Ocurrió un error durante la carga.';
    icon.className = 'fa-solid fa-circle-xmark text-rose-400';

    setTimeout(() => {
      box.classList.add('hidden');
      bar.className = 'bg-gradient-to-r from-blue-500 via-indigo-500 to-emerald-400 h-2 rounded-full transition-all duration-300 ease-out';
      percLabel.className = 'text-xs font-mono font-bold px-3 py-1 rounded-full bg-blue-500/20 text-blue-400 border border-blue-500/30';
      bar.style.width = '0%';
    }, 4500);
  }
};

// ==============================================================
// ESTADO GLOBAL DE ARCHIVOS Y NAVEGACIÓN
// ==============================================================
let currentFilesState = {
  cartera_loaded: false,
  contacts_loaded: false,
  cartera_filename: '',
  contacts_filename: ''
};
let hasProcessedOnce = false;
let currentActiveModule = 'general';
let enableMoraAvila = true;
let currentModulesSummary = {};
let hasUserSelectedModule = false;

function renderSuiteModulesState(data) {
  if (data.active_module) currentActiveModule = data.active_module;
  if (typeof data.enable_mora_avila === 'boolean') enableMoraAvila = data.enable_mora_avila;
  if (data.modules_summary) currentModulesSummary = data.modules_summary;
  if (data.loaded) hasUserSelectedModule = true;

  const headerSwitcher = document.getElementById('header-module-switcher');
  const homeSelector = document.getElementById('home-suite-selector');
  const uploadWorkspace = document.getElementById('module-upload-workspace');
  const uploadBadge = document.getElementById('upload-active-module-badge');
  const uploadIcon = document.getElementById('upload-box-icon');
  const sampleBtnLabel = document.getElementById('btn-load-samples-label');

  if (!enableMoraAvila) {
    if (headerSwitcher) headerSwitcher.classList.add('hidden');
    if (homeSelector) homeSelector.classList.add('hidden');
    if (uploadBadge) uploadBadge.classList.add('hidden');
    if (uploadWorkspace) uploadWorkspace.classList.remove('hidden');
    return;
  }

  if (headerSwitcher) headerSwitcher.classList.remove('hidden');
  if (homeSelector) homeSelector.classList.remove('hidden');
  if (uploadBadge) uploadBadge.classList.remove('hidden');

  if (uploadWorkspace) {
    if (hasUserSelectedModule) {
      uploadWorkspace.classList.remove('hidden');
    } else {
      uploadWorkspace.classList.add('hidden');
    }
  }

  const isAvila = currentActiveModule === 'avila';

  // 1. Header Switcher Buttons
  const hBtnGen = document.getElementById('header-btn-mod-general');
  const hBtnAvi = document.getElementById('header-btn-mod-avila');
  if (hBtnGen && hBtnAvi) {
    if (!hasUserSelectedModule) {
      hBtnGen.className = 'inline-flex items-center gap-1.5 px-3 py-1 rounded-lg text-xs font-semibold transition-all cursor-pointer text-slate-600 hover:text-slate-900';
      hBtnAvi.className = 'inline-flex items-center gap-1.5 px-3 py-1 rounded-lg text-xs font-semibold transition-all cursor-pointer text-slate-600 hover:text-slate-900';
    } else if (isAvila) {
      hBtnGen.className = 'inline-flex items-center gap-1.5 px-3 py-1 rounded-lg text-xs font-semibold transition-all cursor-pointer text-slate-600 hover:text-slate-900';
      hBtnAvi.className = 'inline-flex items-center gap-1.5 px-3 py-1 rounded-lg text-xs font-bold transition-all cursor-pointer bg-emerald-600 text-white shadow-xs';
    } else {
      hBtnGen.className = 'inline-flex items-center gap-1.5 px-3 py-1 rounded-lg text-xs font-bold transition-all cursor-pointer bg-blue-600 text-white shadow-xs';
      hBtnAvi.className = 'inline-flex items-center gap-1.5 px-3 py-1 rounded-lg text-xs font-semibold transition-all cursor-pointer text-slate-600 hover:text-slate-900';
    }
  }

  // 2. Upload Box Active Module Indicator
  if (uploadBadge) {
    if (isAvila) {
      uploadBadge.className = 'inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-emerald-50 text-emerald-700 border border-emerald-200';
      uploadBadge.innerHTML = `<i class="fa-solid fa-mountain-sun text-[10px]"></i> Módulo Activo: Mora Ávila`;
    } else {
      uploadBadge.className = 'inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-blue-50 text-blue-700 border border-blue-200';
      uploadBadge.innerHTML = `<i class="fa-solid fa-building-columns text-[10px]"></i> Módulo Activo: Mora Ubi`;
    }
  }
  if (uploadIcon) {
    uploadIcon.className = isAvila
      ? 'w-12 h-12 rounded-2xl bg-emerald-600 text-white flex items-center justify-center text-xl shadow-xs shrink-0 transition-colors'
      : 'w-12 h-12 rounded-2xl bg-blue-600 text-white flex items-center justify-center text-xl shadow-xs shrink-0 transition-colors';
  }
  if (sampleBtnLabel) {
    sampleBtnLabel.textContent = isAvila ? 'Cargar Muestra Mora Ávila' : 'Cargar Muestra Mora Ubi';
  }

  // 3. Home Suite Cards (General & Avila)
  const cardGen = document.getElementById('card-module-general');
  const cardAvi = document.getElementById('card-module-avila');
  const badgeActGen = document.getElementById('badge-active-general');
  const badgeActAvi = document.getElementById('badge-active-avila');
  const btnLblGen = document.getElementById('btn-label-open-general');
  const btnLblAvi = document.getElementById('btn-label-open-avila');

  if (cardGen && cardAvi) {
    if (!hasUserSelectedModule) {
      cardGen.className = 'relative rounded-2xl border-2 border-slate-200 bg-white hover:border-blue-500 p-5 transition-all cursor-pointer hover:shadow-md flex flex-col justify-between gap-4';
      cardAvi.className = 'relative rounded-2xl border-2 border-slate-200 bg-white hover:border-emerald-500 p-5 transition-all cursor-pointer hover:shadow-md flex flex-col justify-between gap-4';
      if (badgeActGen) badgeActGen.className = 'hidden items-center gap-1 text-[10px] font-bold px-2 py-0.5 rounded-full bg-blue-600 text-white';
      if (badgeActAvi) badgeActAvi.className = 'hidden items-center gap-1 text-[10px] font-bold px-2 py-0.5 rounded-full bg-emerald-600 text-white';
      if (btnLblGen) btnLblGen.textContent = 'Seleccionar Módulo Ubi';
      if (btnLblAvi) btnLblAvi.textContent = 'Seleccionar Módulo Ávila';
    } else if (isAvila) {
      cardGen.className = 'relative rounded-2xl border-2 border-slate-200 bg-white hover:border-blue-400 p-5 transition-all cursor-pointer hover:shadow-md flex flex-col justify-between gap-4';
      cardAvi.className = 'relative rounded-2xl border-2 border-emerald-600 bg-emerald-50/20 p-5 transition-all cursor-pointer hover:shadow-md flex flex-col justify-between gap-4';
      if (badgeActGen) badgeActGen.className = 'hidden items-center gap-1 text-[10px] font-bold px-2 py-0.5 rounded-full bg-blue-600 text-white';
      if (badgeActAvi) badgeActAvi.className = 'inline-flex items-center gap-1 text-[10px] font-bold px-2 py-0.5 rounded-full bg-emerald-600 text-white';
      if (btnLblGen) btnLblGen.textContent = 'Cambiar a Mora Ubi';
      if (btnLblAvi) btnLblAvi.textContent = 'Módulo Activo';
    } else {
      cardGen.className = 'relative rounded-2xl border-2 border-blue-600 bg-blue-50/20 p-5 transition-all cursor-pointer hover:shadow-md flex flex-col justify-between gap-4';
      cardAvi.className = 'relative rounded-2xl border-2 border-slate-200 bg-white hover:border-emerald-400 p-5 transition-all cursor-pointer hover:shadow-md flex flex-col justify-between gap-4';
      if (badgeActGen) badgeActGen.className = 'inline-flex items-center gap-1 text-[10px] font-bold px-2 py-0.5 rounded-full bg-blue-600 text-white';
      if (badgeActAvi) badgeActAvi.className = 'hidden items-center gap-1 text-[10px] font-bold px-2 py-0.5 rounded-full bg-emerald-600 text-white';
      if (btnLblGen) btnLblGen.textContent = 'Módulo Activo';
      if (btnLblAvi) btnLblAvi.textContent = 'Cambiar a Mora Ávila';
    }
  }

  // 4. Stats & Status Dots per Module
  ['general', 'avila'].forEach(modKey => {
    const modInfo = currentModulesSummary[modKey] || { loaded: false, total_cuotas: 0, total_clientes: 0, monto_total: 0 };
    const pill = document.getElementById(`status-pill-${modKey}`);
    const statsLine = document.getElementById(`stats-line-${modKey}`);
    const dot = document.getElementById(`header-dot-${modKey}`);

    if (modInfo.loaded) {
      if (pill) {
        pill.className = 'text-[11px] font-bold px-2.5 py-1 rounded-full bg-emerald-100 text-emerald-800 border border-emerald-300 shrink-0';
        pill.innerHTML = `<i class="fa-solid fa-check mr-1"></i> Cargado`;
      }
      if (statsLine) {
        const montoFmt = '$' + Number(modInfo.monto_total || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
        statsLine.innerHTML = `<span class="text-slate-900 font-bold">${Number(modInfo.total_cuotas || 0).toLocaleString()} cuotas</span> · <span class="text-slate-700">${Number(modInfo.total_clientes || 0).toLocaleString()} clientes</span> · <span class="text-emerald-700 font-extrabold">${montoFmt} USD</span>`;
      }
      if (dot) dot.className = 'w-2 h-2 rounded-full bg-emerald-400 ml-0.5 shadow-2xs';
    } else {
      if (pill) {
        pill.className = 'text-[11px] font-semibold px-2.5 py-1 rounded-full bg-slate-100 text-slate-600 border border-slate-200 shrink-0';
        pill.textContent = 'Sin cargar';
      }
      if (statsLine) {
        statsLine.textContent = 'Ningún archivo cargado en este módulo';
      }
      if (dot) dot.className = 'w-1.5 h-1.5 rounded-full bg-slate-300 ml-0.5';
    }
  });
}

async function selectActiveModule(moduleId, scrollToUpload = false) {
  if (moduleId !== 'general' && moduleId !== 'avila') return;
  hasUserSelectedModule = true;
  currentActiveModule = moduleId;

  // 1. Feedback visual instantáneo para que la UI reaccione de inmediato al clic
  const ws = document.getElementById('module-upload-workspace');
  if (ws) ws.classList.remove('hidden');
  renderSuiteModulesState({
    active_module: moduleId,
    enable_mora_avila: enableMoraAvila,
    modules_summary: currentModulesSummary
  });

  if (scrollToUpload && ws) {
    setTimeout(() => {
      ws.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }, 40);
  }

  // 2. Sincronizar estado con el servidor
  try {
    const res = await fetch('/api/module/select', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ module_id: moduleId })
    });
    const data = await res.json();
    if (!res.ok) {
      showToast('error', 'Error al cambiar módulo', data.detail || 'No se pudo cambiar el módulo.');
      return;
    }

    updateUIState(data);
    const modName = moduleId === 'avila' ? 'Módulo Mora Ávila' : 'Módulo Mora Ubi';

    // Si estamos en una pestaña de resultados o call center y el módulo tiene datos, refrescar vista
    const resultsTab = document.getElementById('tab-results');
    const contactsTab = document.getElementById('tab-contacts');
    const segmentTab = document.getElementById('tab-segment');

    if (data.loaded) {
      if (resultsTab && !resultsTab.classList.contains('hidden')) {
        await processFilters();
      } else if (contactsTab && !contactsTab.classList.contains('hidden')) {
        await loadCallCenterData();
      }
      showToast('info', modName, `Visualizando ${Number(data.metadata?.total_rows || 0).toLocaleString()} cuotas activas de ${modName}.`);
    } else {
      if ((resultsTab && !resultsTab.classList.contains('hidden')) ||
          (contactsTab && !contactsTab.classList.contains('hidden')) ||
          (segmentTab && !segmentTab.classList.contains('hidden'))) {
        switchTab('tab-upload', true);
      }
      showToast('info', modName, `${modName} habilitado. Sube tu archivo o haz clic en el botón de muestra.`);
    }
  } catch (err) {
    showToast('error', 'Error de conexión', 'No se pudo cambiar de módulo: ' + err.message);
  }
}

function highlightMissingCard(cardType) {
  const cardId = cardType === 'cartera' ? 'card-status-cartera' : 'card-status-contacts';
  const card = document.getElementById(cardId);
  if (!card) return;

  card.classList.remove('border-slate-200/90');
  card.classList.add('ring-2', 'ring-rose-500', 'bg-rose-50/70', 'border-rose-300', 'scale-[1.02]');

  setTimeout(() => {
    card.classList.remove('ring-2', 'ring-rose-500', 'bg-rose-50/70', 'border-rose-300', 'scale-[1.02]');
    card.classList.add('border-slate-200/90');
  }, 2500);
}

function handleProcessClick() {
  const hasCartera = Boolean(currentFilesState.cartera_loaded);

  if (!hasCartera) {
    showToast('error', 'Falta el Archivo de Cartera', 'Debes subir el archivo de Cartera / Mora (mora.csv o .xlsx) para poder procesar y cruzar con la Base de Datos.');
    highlightMissingCard('cartera');
    return false;
  }

  showToast('success', '¡Cartera Lista!', 'Cartera cruzada con la Base de Datos de Contactos y Campañas. Abriendo segmentación...');
  switchTab('tab-segment', true);
  return true;
}

function switchTab(tabId, bypassCheck = false) {
  if (tabId !== 'tab-upload' && !bypassCheck) {
    if (!currentFilesState.cartera_loaded) {
      handleProcessClick();
      return;
    }
  }

  if (tabId === 'tab-results' && !hasProcessedOnce) {
    processFilters();
    return;
  }

  ['tab-upload', 'tab-segment', 'tab-results', 'tab-contacts'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.classList.add('hidden');
  });
  ['nav-upload', 'nav-segment', 'nav-results', 'nav-contacts'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.classList.remove('tab-active');
  });

  const targetTab = document.getElementById(tabId);
  if (targetTab) targetTab.classList.remove('hidden');
  const targetNav = document.getElementById('nav-' + tabId.replace('tab-', ''));
  if (targetNav) targetNav.classList.add('tab-active');

  if (tabId === 'tab-contacts') {
    loadCallCenterData();
  }
}

window.addEventListener('DOMContentLoaded', async () => {
  try {
    const res = await fetch('/api/status');
    const data = await res.json();
    updateUIState(data);
  } catch (err) {
    console.error("Error al obtener status:", err);
  }
});

function updateUIState(data) {
  const meta = data.metadata || {};
  const totals = data.totals || {};

  renderSuiteModulesState(data);

  currentFilesState.cartera_loaded = Boolean(data.loaded);
  currentFilesState.contacts_loaded = Boolean(data.contacts_loaded);
  currentFilesState.cartera_filename = data.cartera_filename || meta.cartera_filename || '';
  if (data.contacts_filename || meta.contacts_filename) {
    currentFilesState.contacts_filename = data.contacts_filename || meta.contacts_filename;
  }

  const carteraCard = document.getElementById('card-status-cartera');
  const carteraBadge = document.getElementById('badge-cartera');
  const carteraIcon = document.getElementById('icon-status-cartera');
  const carteraName = document.getElementById('cartera-filename-label');
  const carteraCount = document.getElementById('cartera-count-label');
  const carteraDetail = document.getElementById('cartera-status-detail');

  if (currentFilesState.cartera_loaded) {
    if (carteraBadge) {
      carteraBadge.className = 'text-[10px] font-bold px-2.5 py-0.5 rounded-full bg-emerald-50 text-emerald-700 border border-emerald-200';
      carteraBadge.textContent = '✓ Listo';
    }
    if (carteraIcon) {
      carteraIcon.className = 'w-9 h-9 rounded-xl bg-emerald-50 text-emerald-600 border border-emerald-200 flex items-center justify-center text-sm shrink-0 transition-colors';
    }
    if (carteraName) carteraName.textContent = currentFilesState.cartera_filename || 'Cartera de Mora activa';
    if (carteraCount) carteraCount.textContent = `${Number(meta.total_rows || 0).toLocaleString()} cuotas`;
    if (carteraDetail) carteraDetail.textContent = 'Verificado y cruzado con Base de Datos';
    if (carteraCard) carteraCard.className = 'flex items-center justify-between bg-emerald-50/20 p-3.5 rounded-xl border border-emerald-200/80 shadow-2xs transition-all';
  } else {
    if (carteraBadge) {
      carteraBadge.className = 'text-[10px] font-bold px-2 py-0.5 rounded-full bg-rose-50 text-rose-700 border border-rose-200';
      carteraBadge.textContent = 'Obligatorio';
    }
    if (carteraIcon) {
      carteraIcon.className = 'w-9 h-9 rounded-xl bg-slate-100 text-slate-400 flex items-center justify-center text-sm shrink-0 transition-colors';
    }
    if (carteraName) carteraName.textContent = '1. Archivo de Cartera / Mora';
    if (carteraCount) carteraCount.textContent = '0 cuotas';
    if (carteraDetail) carteraDetail.textContent = 'Obligatorio: sube mora.csv o .xlsx para procesar';
    if (carteraCard) carteraCard.className = 'flex items-center justify-between bg-white p-3.5 rounded-xl border border-slate-200/90 shadow-2xs transition-all';
  }

  const contactsCard = document.getElementById('card-status-contacts');
  const contactsBadge = document.getElementById('badge-contacts');
  const contactsIcon = document.getElementById('icon-status-contacts');
  const contactsName = document.getElementById('contacts-filename-label');
  const contactsCount = document.getElementById('contacts-count-label');
  const contactsDetail = document.getElementById('contacts-status-detail');

  if (currentFilesState.contacts_loaded) {
    if (contactsBadge) {
      contactsBadge.className = 'text-[10px] font-bold px-2.5 py-0.5 rounded-full bg-emerald-50 text-emerald-700 border border-emerald-200';
      contactsBadge.textContent = '✓ En Base de Datos';
    }
    if (contactsIcon) {
      contactsIcon.className = 'w-9 h-9 rounded-xl bg-emerald-50 text-emerald-600 border border-emerald-200 flex items-center justify-center text-sm shrink-0 transition-colors';
    }
    if (contactsName) contactsName.textContent = currentFilesState.contacts_filename || 'Base de Contactos y Campañas (BD)';
    const totalContactsIdx = Number(data.total_contacts || meta.total_contacts || 0);
    const totalWithPhone = Number(totals.total_con_telefono || 0);
    if (contactsCount) {
      if (currentFilesState.cartera_loaded && totalWithPhone > 0) {
        contactsCount.textContent = `${totalWithPhone.toLocaleString()} con WhatsApp`;
      } else {
        contactsCount.textContent = `${(totalContactsIdx || totalWithPhone).toLocaleString()} en BD`;
      }
    }
    if (contactsDetail) contactsDetail.textContent = 'Activa en BD local (Opcional: sube archivo para sumar más)';
    if (contactsCard) contactsCard.className = 'flex items-center justify-between bg-emerald-50/20 p-3.5 rounded-xl border border-emerald-200/80 shadow-2xs transition-all';
  } else {
    if (contactsBadge) {
      contactsBadge.className = 'text-[10px] font-bold px-2 py-0.5 rounded-full bg-amber-50 text-amber-700 border border-amber-200';
      contactsBadge.textContent = 'Opcional';
    }
    if (contactsIcon) {
      contactsIcon.className = 'w-9 h-9 rounded-xl bg-slate-100 text-slate-400 flex items-center justify-center text-sm shrink-0 transition-colors';
    }
    if (contactsName) contactsName.textContent = '2. Base de Contactos y Campañas';
    if (contactsCount) contactsCount.textContent = '0 en BD';
    if (contactsDetail) contactsDetail.textContent = 'Opcional: sube archivo para cargar a la BD';
    if (contactsCard) contactsCard.className = 'flex items-center justify-between bg-white p-3.5 rounded-xl border border-slate-200/90 shadow-2xs transition-all';
  }

  const summaryBadge = document.getElementById('files-status-summary');
  const btnProcess = document.getElementById('btn-process-files');

  if (summaryBadge) {
    if (currentFilesState.cartera_loaded) {
      summaryBadge.className = 'text-xs font-bold px-3 py-1 rounded-full bg-emerald-100 text-emerald-800 border border-emerald-300 flex items-center gap-1.5';
      summaryBadge.innerHTML = `<i class="fa-solid fa-check"></i> Cartera + BD Listas para Procesar`;
    } else if (currentFilesState.contacts_loaded) {
      summaryBadge.className = 'text-xs font-bold px-3 py-1 rounded-full bg-blue-100 text-blue-800 border border-blue-300 flex items-center gap-1.5';
      summaryBadge.innerHTML = `<i class="fa-solid fa-database"></i> BD Activa (Sube Cartera de Mora)`;
    } else {
      summaryBadge.className = 'text-xs font-bold px-3 py-1 rounded-full bg-slate-200 text-slate-700';
      summaryBadge.textContent = 'Esperando Cartera de Mora';
    }
  }

  if (btnProcess) {
    if (currentFilesState.cartera_loaded) {
      btnProcess.className = 'w-full sm:w-auto inline-flex items-center justify-center gap-2 px-6 py-3 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold rounded-xl shadow-md transition-all cursor-pointer ring-2 ring-emerald-400/50';
      btnProcess.innerHTML = `<span>Procesar Cartera</span> <i class="fa-solid fa-arrow-right"></i>`;
    } else {
      btnProcess.className = 'w-full sm:w-auto inline-flex items-center justify-center gap-2 px-6 py-3 bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold rounded-xl shadow-sm transition-all cursor-pointer';
      btnProcess.innerHTML = `<span>Procesar Cartera</span> <i class="fa-solid fa-arrow-right"></i>`;
    }
  }

  const topBadge = document.getElementById('file-status-badge');
  if (topBadge) {
    if (currentFilesState.cartera_loaded) {
      topBadge.className = 'flex items-center gap-2 px-3 py-1.5 rounded-full text-xs font-medium bg-emerald-50 text-emerald-700 border border-emerald-200';
      topBadge.innerHTML = `<span class="w-2 h-2 rounded-full bg-emerald-500"></span><span>Cartera + BD vinculadas (${Number(meta.total_rows || 0).toLocaleString()} cuotas)</span>`;
    } else if (currentFilesState.contacts_loaded) {
      const totalContactsIdx = Number(data.total_contacts || meta.total_contacts || 0);
      topBadge.className = 'flex items-center gap-2 px-3 py-1.5 rounded-full text-xs font-medium bg-blue-50 text-blue-700 border border-blue-200';
      topBadge.innerHTML = `<span class="w-2 h-2 rounded-full bg-blue-500"></span><span>BD Activa (${totalContactsIdx.toLocaleString()} clientes) · Sube Cartera</span>`;
    } else {
      topBadge.className = 'flex items-center gap-2 px-3 py-1.5 rounded-full text-xs font-medium bg-slate-100 text-slate-500 border border-slate-200';
      topBadge.innerHTML = `<span class="w-2 h-2 rounded-full bg-slate-400"></span><span>Sin cartera cargada</span>`;
    }
  }

  hasProcessedOnce = false;

  if (currentFilesState.cartera_loaded) {
    populateMetadata(meta, totals);
  } else {
    const summaryCard = document.getElementById('file-summary-card');
    if (summaryCard) summaryCard.classList.add('hidden');
  }

  const globalSummary = data.resumen_global_tramos || (data.metadata && data.metadata.resumen_global_tramos);
  const fibexSummary = data.resumen_fibex || (data.metadata && data.metadata.resumen_fibex);
  if (globalSummary) {
    currentGlobalSummary = globalSummary;
    updateQuickFilterBadges(globalSummary);
  }
  if (fibexSummary) currentFibexSummary = fibexSummary;

  if (currentExecutiveView === 'fibex' && currentFibexSummary) {
    renderFibexScorecard(currentFibexSummary);
  } else if (currentGlobalSummary) {
    renderGlobalTramosScorecard(currentGlobalSummary);
  }

  const matrixData = data.matriz_campanas_tramos || (data.metadata && data.metadata.matriz_campanas_tramos);
  if (matrixData) {
    renderCampanasMatrix(matrixData);
  }
}

async function loadSampleFiles(moduleId = null) {
  const targetMod = moduleId || currentActiveModule || 'general';
  const btn = document.getElementById('btn-load-samples');
  const originalText = btn ? btn.innerHTML : '';
  if (btn) {
    btn.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> Cargando y cruzando...`;
    btn.disabled = true;
  }

  const sampleLabel = targetMod === 'avila' ? 'reporte_mora_2026-10-28.csv (Mora Ávila)' : 'mora.csv (Mora Ubi)';
  ProgressController.start('Cargando cartera de prueba...', `Leyendo ${sampleLabel} y cruzando con la Base de Datos SQLite...`);

  try {
    const url = moduleId ? `/api/load-sample?module_id=${encodeURIComponent(moduleId)}` : '/api/load-sample';
    const res = await fetch(url, { method: 'POST' });
    const data = await res.json();
    if (res.ok) {
      updateUIState({
        loaded: true,
        contacts_loaded: true,
        active_module: data.active_module || targetMod,
        enable_mora_avila: data.enable_mora_avila,
        modules_summary: data.modules_summary,
        cartera_filename: data.metadata?.cartera_filename || sampleLabel,
        contacts_filename: data.metadata?.contacts_filename || 'Base de Datos SQLite (Contactos + Campañas)',
        metadata: data.metadata,
        totals: data.totals,
        resumen_global_tramos: data.resumen_global_tramos,
        resumen_fibex: data.resumen_fibex,
        matriz_campanas_tramos: data.matriz_campanas_tramos
      });
      ProgressController.complete('¡Cartera y Base de Datos sincronizadas!', () => {
        showToast('success', '¡Carga Completada con Éxito!', `${Number(data.metadata?.total_rows || 0).toLocaleString()} cuotas de cartera cruzadas con ${Number(data.metadata?.total_contacts || 0).toLocaleString()} clientes en Base de Datos.`);
        switchTab('tab-segment', true);
      });
    } else {
      ProgressController.error(data.detail || 'Error al procesar archivos');
      showToast('error', 'Error en la Carga', data.detail || 'No se pudieron procesar los archivos de prueba.');
    }
  } catch (err) {
    ProgressController.error(err.message);
    showToast('error', 'Error de Conexión', 'No se pudo conectar con el servidor: ' + err.message);
  } finally {
    if (btn) {
      btn.innerHTML = originalText;
      btn.disabled = false;
    }
  }
}

async function handleUnifiedUpload(e) {
  const files = (e.target && e.target.files) ? Array.from(e.target.files) : [];
  if (!files.length) return;
  for (const file of files) {
    await processSingleFile(file);
  }
  if (e.target) e.target.value = '';
}

async function handleCarteraUpload(e) {
  const file = (e.target && e.target.files) ? e.target.files[0] : null;
  if (file) await processSingleFile(file);
  if (e.target) e.target.value = '';
}

async function handleContactsUpload(e) {
  const file = (e.target && e.target.files) ? e.target.files[0] : null;
  if (file) await processSingleFile(file);
  if (e.target) e.target.value = '';
}

async function processSingleFile(file) {
  if (!file) return;

  const promptDiv = document.getElementById('unified-prompt');
  const loadDiv = document.getElementById('unified-loading');
  const loadText = document.getElementById('unified-loading-text');

  if (promptDiv) promptDiv.classList.add('hidden');
  if (loadDiv) loadDiv.classList.remove('hidden');
  if (loadText) loadText.textContent = `Procesando ${file.name}...`;

  ProgressController.start(`Subiendo: ${file.name}`, 'Analizando estructura del archivo y sincronizando con Base de Datos...');

  const formData = new FormData();
  formData.append('file', file);

  try {
    const res = await fetch('/api/upload', { method: 'POST', body: formData });
    const data = await res.json();
    if (res.ok) {
      if (data.auto_routed === 'contacts') {
        updateUIState({
          loaded: currentFilesState.cartera_loaded,
          contacts_loaded: true,
          active_module: data.active_module,
          enable_mora_avila: data.enable_mora_avila,
          modules_summary: data.modules_summary,
          cartera_filename: currentFilesState.cartera_filename,
          contacts_filename: `BD Actualizada (+ ${file.name})`,
          total_contacts: data.total_contacts,
          metadata: data.metadata,
          totals: data.totals,
          resumen_global_tramos: data.resumen_global_tramos,
          resumen_fibex: data.resumen_fibex,
          matriz_campanas_tramos: data.matriz_campanas_tramos
        });
        ProgressController.complete('¡Base de Datos actualizada!', () => {
          const nuevos = Number(data.added_contacts || data.nuevos_agregados || 0);
          const extraMsg = nuevos > 0 ? ` (${nuevos.toLocaleString()} nuevos agregados)` : '';
          if (currentFilesState.cartera_loaded) {
            showToast('success', '¡Base de Datos Actualizada!', `Total en BD: ${Number(data.total_contacts).toLocaleString()} clientes${extraMsg}. Ya puedes hacer clic en 'Procesar Cartera'.`);
          } else {
            showToast('success', '¡Base de Datos Actualizada!', `Total en BD: ${Number(data.total_contacts).toLocaleString()} clientes${extraMsg}. Sube tu archivo de Cartera (mora.csv) cuando quieras procesar.`);
          }
        });
      } else {
        updateUIState({
          loaded: true,
          contacts_loaded: (data.metadata && data.metadata.has_contacts) || currentFilesState.contacts_loaded,
          active_module: data.active_module,
          enable_mora_avila: data.enable_mora_avila,
          modules_summary: data.modules_summary,
          cartera_filename: file.name,
          contacts_filename: currentFilesState.contacts_filename,
          metadata: data.metadata,
          totals: data.totals,
          resumen_global_tramos: data.resumen_global_tramos,
          resumen_fibex: data.resumen_fibex,
          matriz_campanas_tramos: data.matriz_campanas_tramos
        });
        ProgressController.complete('¡Cartera procesada y cruzada con BD!', () => {
          showToast('success', '¡Cartera Lista para Procesar!', `Cartera cargada (${Number(data.metadata?.total_rows || 0).toLocaleString()} cuotas) y cruzada automáticamente con la Base de Datos de Contactos y Campañas.`);
        });
      }
    } else {
      ProgressController.error(data.detail || 'Error en el archivo');
      showToast('error', 'Error al procesar archivo', data.detail || 'No se pudo interpretar el archivo.');
    }
  } catch (err) {
    ProgressController.error(err.message);
    showToast('error', 'Error de conexión', 'No se pudo subir el archivo: ' + err.message);
  } finally {
    if (promptDiv) promptDiv.classList.remove('hidden');
    if (loadDiv) loadDiv.classList.add('hidden');
  }
}

function populateMetadata(meta, totals) {
  appMetadata = meta;

  const summaryCard = document.getElementById('file-summary-card');
  if (summaryCard) {
    if (currentFilesState.cartera_loaded && currentFilesState.contacts_loaded) {
      summaryCard.classList.remove('hidden');
    } else {
      summaryCard.classList.add('hidden');
    }
  }

  document.getElementById('meta-rows').textContent = Number(meta.total_rows || 0).toLocaleString();
  document.getElementById('meta-clients').textContent = Number(meta.total_clientes || 0).toLocaleString();
  document.getElementById('meta-mora-range').textContent = `${meta.min_dias || 0} a ${meta.max_dias || 0} días`;
  document.getElementById('meta-total-amount').textContent = `$${Number(totals.monto_total || 0).toLocaleString('en-US', {minimumFractionDigits: 2})} USD`;
  document.getElementById('meta-tel-clients').textContent = Number(totals.total_con_telefono || 0).toLocaleString();
  document.getElementById('meta-tel-pct').textContent = `${totals.pct_con_telefono || 0}% de cobertura`;
  document.getElementById('meta-contactable-amount').textContent = `$${Number(totals.monto_contactable || 0).toLocaleString('en-US', {minimumFractionDigits: 2})} USD`;
  document.getElementById('meta-contactable-pct').textContent = `${totals.pct_monto_contactable || 0}% de cartera`;

  document.getElementById('kpi-clientes').textContent = Number(totals.total_clientes || meta.total_clientes || 0).toLocaleString();
  document.getElementById('kpi-cuotas').textContent = `${Number(totals.total_cuotas || meta.total_rows || 0).toLocaleString()} cuotas en mora`;
  document.getElementById('kpi-monto-total').textContent = `$${Number(totals.monto_total || 0).toLocaleString('en-US', {minimumFractionDigits: 2})} USD`;
  document.getElementById('kpi-promedio').textContent = `Promedio: $${Number(totals.promedio_cliente || 0).toLocaleString('en-US', {minimumFractionDigits: 2})} / cliente`;
  document.getElementById('kpi-con-telefono').textContent = Number(totals.total_con_telefono || 0).toLocaleString();
  document.getElementById('kpi-pct-telefono').textContent = `${totals.pct_con_telefono || 0}% de cobertura`;
  document.getElementById('kpi-monto-contactable').textContent = `$${Number(totals.monto_contactable || 0).toLocaleString('en-US', {minimumFractionDigits: 2})} USD`;
  document.getElementById('kpi-pct-monto').textContent = `${totals.pct_monto_contactable || 0}% de la deuda`;

  const campanaContainer = document.getElementById('campanas-list');
  campanaContainer.innerHTML = '';
  const campanas = meta.tipos_campana || [];
  if (campanas.length === 0) {
    campanaContainer.innerHTML = '<span class="text-slate-400">Sin campañas en los contactos.</span>';
  } else {
    campanas.forEach(c => {
      const div = document.createElement('div');
      div.className = 'flex items-center gap-2 campana-item';
      div.dataset.name = c.toLowerCase();
      div.innerHTML = `
        <input type="checkbox" id="camp-${c}" value="${c}" checked class="campana-checkbox rounded border-slate-300 text-blue-600 focus:ring-blue-500">
        <label for="camp-${c}" class="text-slate-700 cursor-pointer select-none truncate" title="${c}">${c}</label>
      `;
      campanaContainer.appendChild(div);
    });
  }

  const progContainer = document.getElementById('programas-list');
  progContainer.innerHTML = '';
  (meta.programas || []).forEach(p => {
    const div = document.createElement('div');
    div.className = 'flex items-center gap-2';
    div.innerHTML = `
      <input type="checkbox" id="prog-${p}" value="${p}" checked class="program-checkbox rounded border-slate-300 text-blue-600 focus:ring-blue-500">
      <label for="prog-${p}" class="text-slate-700 cursor-pointer select-none truncate" title="${p}">${p}</label>
    `;
    progContainer.appendChild(div);
  });

  const tiposContainer = document.getElementById('tipos-list');
  tiposContainer.innerHTML = '';
  (meta.tipos || []).forEach(t => {
    const div = document.createElement('div');
    div.className = 'flex items-center gap-2';
    div.innerHTML = `
      <input type="checkbox" id="tipo-${t}" value="${t}" checked class="tipo-checkbox rounded border-slate-300 text-blue-600 focus:ring-blue-500">
      <label for="tipo-${t}" class="text-slate-700 cursor-pointer select-none">${t}</label>
    `;
    tiposContainer.appendChild(div);
  });
}

function setMoraRange(min, max, customLabel = null) {
  document.getElementById('input-min-dias').value = min;
  document.getElementById('input-max-dias').value = max;
  document.getElementById('range-badge').textContent = customLabel || `${min} - ${max} días`;
}

document.getElementById('input-min-dias').addEventListener('input', () => {
  const min = document.getElementById('input-min-dias').value;
  const max = document.getElementById('input-max-dias').value;
  document.getElementById('range-badge').textContent = `${min} - ${max} días`;
});

document.getElementById('input-max-dias').addEventListener('input', () => {
  const min = document.getElementById('input-min-dias').value;
  const max = document.getElementById('input-max-dias').value;
  document.getElementById('range-badge').textContent = `${min} - ${max} días`;
});

function toggleAllCheckboxes(className, state) {
  document.querySelectorAll('.' + className).forEach(cb => cb.checked = state);
}

function filterCheckboxList(className, query) {
  const q = query.toLowerCase().trim();
  document.querySelectorAll('.campana-item').forEach(item => {
    if (!q || item.dataset.name.includes(q)) {
      item.style.display = 'flex';
    } else {
      item.style.display = 'none';
    }
  });
}

const dropUnified = document.getElementById('drop-zone-unified');
if (dropUnified) {
  dropUnified.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropUnified.classList.add('border-blue-500', 'bg-blue-50/50', 'ring-2', 'ring-blue-400/50');
  });
  dropUnified.addEventListener('dragleave', () => {
    dropUnified.classList.remove('border-blue-500', 'bg-blue-50/50', 'ring-2', 'ring-blue-400/50');
  });
  dropUnified.addEventListener('drop', async (e) => {
    e.preventDefault();
    dropUnified.classList.remove('border-blue-500', 'bg-blue-50/50', 'ring-2', 'ring-blue-400/50');
    if (e.dataTransfer.files && e.dataTransfer.files.length) {
      for (const file of Array.from(e.dataTransfer.files)) {
        await processSingleFile(file);
      }
    }
  });
}

async function processFilters() {
  const minDias = parseInt(document.getElementById('input-min-dias').value) || 0;
  const maxDias = parseInt(document.getElementById('input-max-dias').value) || 999;
  const onlyPhone = document.getElementById('check-only-phone').checked;

  const campanas = Array.from(document.querySelectorAll('.campana-checkbox:checked')).map(cb => cb.value);
  const programas = Array.from(document.querySelectorAll('.program-checkbox:checked')).map(cb => cb.value);
  const tipos = Array.from(document.querySelectorAll('.tipo-checkbox:checked')).map(cb => cb.value);

  const btn = document.getElementById('btn-process');
  const originalText = btn.innerHTML;
  btn.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> Segmentando cartera...`;
  btn.disabled = true;

  try {
    const res = await fetch('/api/process', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        min_dias: minDias,
        max_dias: maxDias,
        tipos_campana: campanas,
        programas: programas,
        tipos: tipos,
        only_with_phone: onlyPhone,
        group_by: "Tramo_Mora",
        cartera_activa: currentCarteraActivaManual
      })
    });

    const data = await res.json();
    if (res.ok) {
      hasProcessedOnce = true;
      renderResults(data, minDias, maxDias);
      showToast('success', 'Filtros Aplicados', `${Number(data.total_filtered).toLocaleString()} cuotas coincidentes encontradas.`);
      switchTab('tab-results', true);
    } else {
      showToast('error', 'Error al filtrar', data.detail || 'Ocurrió un error al aplicar los filtros.');
    }
  } catch (err) {
    showToast('error', 'Error al procesar', 'No se pudieron aplicar los filtros: ' + err.message);
  } finally {
    btn.innerHTML = originalText;
    btn.disabled = false;
  }
}

function renderResults(data, minDias, maxDias) {
  const k = data.kpis;

  document.getElementById('results-subtitle').textContent = `Filtro: ${minDias} a ${maxDias} días de mora | ${Number(data.total_filtered).toLocaleString()} cuotas coincidentes`;

  document.getElementById('kpi-clientes').textContent = Number(k.total_clientes).toLocaleString();
  document.getElementById('kpi-cuotas').textContent = `${Number(k.total_cuotas).toLocaleString()} cuotas en mora`;
  document.getElementById('kpi-monto-total').textContent = `$${Number(k.monto_total).toLocaleString('en-US', {minimumFractionDigits: 2})} USD`;
  document.getElementById('kpi-promedio').textContent = `Promedio: $${Number(k.promedio_cliente).toLocaleString('en-US', {minimumFractionDigits: 2})} / cliente`;
  document.getElementById('kpi-con-telefono').textContent = Number(k.total_con_telefono).toLocaleString();
  document.getElementById('kpi-pct-telefono').textContent = `${k.pct_con_telefono}% de cobertura`;
  document.getElementById('kpi-monto-contactable').textContent = `$${Number(k.monto_contactable).toLocaleString('en-US', {minimumFractionDigits: 2})} USD`;
  document.getElementById('kpi-pct-monto').textContent = `${k.pct_monto_contactable}% de la deuda`;

  const globalSummary = data.resumen_global_tramos || (data.metadata && data.metadata.resumen_global_tramos);
  const fibexSummary = data.resumen_fibex || (data.metadata && data.metadata.resumen_fibex);
  if (globalSummary) {
    currentGlobalSummary = globalSummary;
    updateQuickFilterBadges(globalSummary);
  }
  if (fibexSummary) currentFibexSummary = fibexSummary;

  if (currentExecutiveView === 'fibex' && currentFibexSummary) {
    renderFibexScorecard(currentFibexSummary);
  } else if (currentGlobalSummary) {
    renderGlobalTramosScorecard(currentGlobalSummary);
  }

  const matrixData = data.matriz_campanas_tramos || (data.metadata && data.metadata.matriz_campanas_tramos);
  if (matrixData) {
    renderCampanasMatrix(matrixData);
  }

  renderPreviewTable(data.preview, data.total_filtered);
}

// ==============================================================
// INDICADOR EJECUTIVO UNIFICADO (TRAMOS DE MORA / FIBEX VS. NO FIBEX)
// ==============================================================
let currentCarteraActivaManual = null;
let currentGlobalSummary = null;
let currentFibexSummary = null;
let currentExecutiveView = 'tramos';

function switchExecutiveView(viewType) {
  currentExecutiveView = viewType;
  const btnTramos = document.getElementById('tab-btn-tramos');
  const btnFibex = document.getElementById('tab-btn-fibex');

  if (viewType === 'tramos') {
    if (btnTramos) btnTramos.className = 'inline-flex items-center gap-1.5 px-3 py-1 rounded-lg text-xs font-bold transition-all shadow-xs bg-white text-slate-900 cursor-pointer';
    if (btnFibex) btnFibex.className = 'inline-flex items-center gap-1.5 px-3 py-1 rounded-lg text-xs font-bold transition-all text-slate-500 hover:text-slate-800 cursor-pointer';
    if (currentGlobalSummary) renderGlobalTramosScorecard(currentGlobalSummary);
  } else {
    if (btnTramos) btnTramos.className = 'inline-flex items-center gap-1.5 px-3 py-1 rounded-lg text-xs font-bold transition-all text-slate-500 hover:text-slate-800 cursor-pointer';
    if (btnFibex) btnFibex.className = 'inline-flex items-center gap-1.5 px-3 py-1 rounded-lg text-xs font-bold transition-all shadow-xs bg-white text-slate-900 cursor-pointer';
    if (currentFibexSummary) renderFibexScorecard(currentFibexSummary);
  }
}

function parseManualNumber(val) {
  if (!val) return null;
  let cleaned = String(val).replace(/[\$\s]/g, '');
  if ((cleaned.match(/\./g) || []).length > 1) {
    cleaned = cleaned.replace(/\./g, '');
  } else if ((cleaned.match(/,/g) || []).length > 1) {
    cleaned = cleaned.replace(/,/g, '');
  } else if (cleaned.includes('.') && cleaned.includes(',')) {
    if (cleaned.lastIndexOf('.') > cleaned.lastIndexOf(',')) {
      cleaned = cleaned.replace(/,/g, '');
    } else {
      cleaned = cleaned.replace(/\./g, '').replace(',', '.');
    }
  } else if (cleaned.includes(',')) {
    cleaned = cleaned.replace(',', '.');
  }
  const n = parseFloat(cleaned);
  return isNaN(n) || n <= 0 ? null : n;
}

function handleCarteraActivaManualInput(val) {
  currentCarteraActivaManual = parseManualNumber(val);

  const qParam = currentCarteraActivaManual ? `?cartera_activa=${currentCarteraActivaManual}` : '';
  const excelBtn = document.getElementById('btn-download-tramos-excel');
  const pdfBtn = document.getElementById('btn-download-tramos-pdf');
  if (currentExecutiveView === 'fibex') {
    if (excelBtn) excelBtn.href = `/api/fibex/download/excel${qParam}`;
    if (pdfBtn) pdfBtn.href = `/api/fibex/download/pdf${qParam}`;
  } else {
    if (excelBtn) excelBtn.href = `/api/tramos/download/excel${qParam}`;
    if (pdfBtn) pdfBtn.href = `/api/tramos/download/pdf${qParam}`;
  }

  if (window.carteraActivaDebounce) clearTimeout(window.carteraActivaDebounce);
  window.carteraActivaDebounce = setTimeout(() => {
    fetch('/api/tramos/set-cartera-activa', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ cartera_activa: currentCarteraActivaManual })
    }).catch(() => {});
  }, 400);

  if (currentExecutiveView === 'fibex' && currentFibexSummary) {
    renderFibexScorecard(currentFibexSummary);
  } else if (currentGlobalSummary) {
    renderGlobalTramosScorecard(currentGlobalSummary);
  }
}

function updateQuickFilterBadges(globalSummary) {
  try {
    if (!globalSummary || !globalSummary.tramos) return;
    const total = Number(globalSummary.total_cartera_usd || 0);
    const map = {
      '1-30': 'btn-pct-tramo-1-30',
      '31-60': 'btn-pct-tramo-31-60',
      '61-90': 'btn-pct-tramo-61-90',
      '91-180': 'btn-pct-tramo-91-180',
      '181+': 'btn-pct-tramo-181'
    };
    globalSummary.tramos.forEach(t => {
      const badgeId = map[t.id];
      if (badgeId) {
        const el = document.getElementById(badgeId);
        if (el) {
          const pctVal = total > 0 ? ((Number(t.total_monto || 0) / total) * 100).toFixed(1) + '%' : '0.0%';
          el.textContent = pctVal;
        }
      }
    });
    const allEl = document.getElementById('btn-pct-tramo-all');
    if (allEl) allEl.textContent = '100%';
  } catch (err) {
    console.warn('Error en updateQuickFilterBadges:', err);
  }
}

/**
 * Motor compartido de renderizado para el Indicador Ejecutivo (Tramos Oficiales y Fibex vs. NO Fibex).
 */
function renderExecutiveScorecard(cfg) {
  const {
    summary,
    rows,
    title,
    subtitle,
    col1Header,
    col2Header,
    chartTitleText,
    chartTagClass,
    chartTagText,
    excelBaseUrl,
    excelLabel,
    excelTitle,
    pdfBaseUrl,
    pdfLabel,
    pdfTitle,
    footerLabel,
    footerSublabel,
    buildRowMeta
  } = cfg;

  const titleElem = document.getElementById('executive-indicator-title');
  const subtitleElem = document.getElementById('executive-indicator-subtitle');
  if (titleElem) titleElem.textContent = title;
  if (subtitleElem) subtitleElem.textContent = subtitle;

  const th1 = document.getElementById('th-exec-col1');
  const th2 = document.getElementById('th-exec-col2');
  if (th1) th1.textContent = col1Header;
  if (th2) th2.textContent = col2Header;

  const thPct = document.getElementById('th-exec-col-pct');
  if (thPct) {
    thPct.textContent = (currentCarteraActivaManual && currentCarteraActivaManual > 0) ? '% Cartera Activa' : '% de la Mora';
  }

  const chartTitle = document.getElementById('chart-exec-title');
  const chartTag = document.getElementById('chart-exec-tag');
  if (chartTitle) chartTitle.textContent = chartTitleText;
  if (chartTag) {
    chartTag.className = chartTagClass;
    chartTag.textContent = chartTagText;
  }

  const inputElem = document.getElementById('input-cartera-activa-manual');
  if (inputElem) {
    if (currentCarteraActivaManual === null && summary.cartera_activa_usd) {
      currentCarteraActivaManual = Number(summary.cartera_activa_usd);
      inputElem.value = currentCarteraActivaManual.toLocaleString('es-CL');
    } else if (inputElem.value && currentCarteraActivaManual === null) {
      currentCarteraActivaManual = parseManualNumber(inputElem.value);
    }
  }

  const excelBtn = document.getElementById('btn-download-tramos-excel');
  const pdfBtn = document.getElementById('btn-download-tramos-pdf');
  const qParam = currentCarteraActivaManual ? `?cartera_activa=${currentCarteraActivaManual}` : '';
  if (excelBtn) {
    excelBtn.href = `${excelBaseUrl}${qParam}`;
    excelBtn.innerHTML = `<i class="fa-solid fa-file-excel text-emerald-600"></i> ${excelLabel}`;
    excelBtn.title = excelTitle;
  }
  if (pdfBtn) {
    pdfBtn.href = `${pdfBaseUrl}${qParam}`;
    pdfBtn.innerHTML = `<i class="fa-solid fa-file-pdf text-rose-600"></i> ${pdfLabel}`;
    pdfBtn.title = pdfTitle;
  }

  const bannerElem = document.getElementById('banner-cartera-activa');
  if (bannerElem) {
    const formattedTotal = Number(summary.total_cartera_usd).toLocaleString('es-CL', {
      maximumFractionDigits: 0
    });
    bannerElem.textContent = `$${formattedTotal} USD`;
  }

  const tbody = document.getElementById('global-tramos-tbody');
  const tfoot = document.getElementById('global-tramos-tfoot');
  if (!tbody) return;

  tbody.innerHTML = '';

  const chartLabels = [];
  const chartValues = [];
  const chartColors = [];

  rows.forEach(item => {
    const meta = buildRowMeta(item);
    chartLabels.push(meta.chartLabel);
    chartValues.push(item.total_monto);
    chartColors.push(item.color || '#3b82f6');

    let pctBadge = '';
    if (currentCarteraActivaManual && currentCarteraActivaManual > 0) {
      const pct = (item.total_monto / currentCarteraActivaManual) * 100;
      const isHigh = pct >= 8.0;
      pctBadge = isHigh
        ? `<span class="inline-block px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-amber-100 text-amber-800 border border-amber-200 shadow-2xs">${pct.toFixed(2)}%</span>`
        : `<span class="inline-block px-2.5 py-0.5 rounded-full text-[11px] font-semibold bg-slate-100 text-slate-700 border border-slate-200">${pct.toFixed(2)}%</span>`;
    } else {
      const pctMora = summary.total_cartera_usd > 0 ? ((item.total_monto / summary.total_cartera_usd) * 100) : 0;
      pctBadge = `<span class="inline-block px-2.5 py-0.5 rounded-full text-[11px] font-semibold bg-slate-100 text-slate-700 border border-slate-200" title="% sobre la mora cargada">${pctMora.toFixed(1)}%</span>`;
    }

    const tr = document.createElement('tr');
    tr.className = 'hover:bg-slate-50/80 transition-colors border-b border-slate-100 cursor-pointer';
    tr.title = meta.rowTitle;
    tr.onclick = meta.onClick;

    tr.innerHTML = `
      <td class="py-3 px-3.5 text-slate-800 font-medium text-xs">
        <div class="flex items-center gap-2">
          <span class="w-2.5 h-2.5 rounded-full shrink-0 shadow-2xs" style="background-color: ${item.color}"></span>
          ${meta.col1InnerHtml}
        </div>
      </td>
      <td class="py-3 px-3.5">
        <span class="${item.badge_class}">${meta.col2Text}</span>
      </td>
      <td class="py-3 px-3.5 text-right font-mono text-xs text-slate-700">${Number(item.cuotas).toLocaleString()}</td>
      <td class="py-3 px-3.5 text-right font-mono text-xs text-slate-700">${Number(item.clientes).toLocaleString()}</td>
      <td class="py-3 px-3.5 text-right font-mono font-bold text-xs text-slate-900">$${Number(item.total_monto).toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>
      <td class="py-3 px-3.5 text-right">
        ${pctBadge}
      </td>
    `;
    tbody.appendChild(tr);
  });

  if (tfoot) {
    let footerPctDisplay = '';
    if (currentCarteraActivaManual && currentCarteraActivaManual > 0) {
      const totPct = (summary.total_cartera_usd / currentCarteraActivaManual) * 100;
      footerPctDisplay = `<span class="text-blue-700 font-black">${totPct.toFixed(2)}%</span>`;
    } else {
      footerPctDisplay = `<span class="text-slate-700 font-bold font-mono text-xs" title="% sobre la mora cargada">100.0%</span>`;
    }

    tfoot.innerHTML = `
      <tr class="bg-slate-100 font-bold border-t-2 border-slate-300 text-slate-900 text-xs">
        <td class="py-3.5 px-3.5 uppercase tracking-wider text-[11px] text-slate-800">
          <i class="fa-solid fa-layer-group mr-1.5 text-blue-600"></i> ${footerLabel}
        </td>
        <td class="py-3.5 px-3.5 text-slate-600 text-xs font-semibold">${footerSublabel}</td>
        <td class="py-3.5 px-3.5 text-right font-mono text-xs text-slate-800">${Number(summary.total_cuotas).toLocaleString()}</td>
        <td class="py-3.5 px-3.5 text-right font-mono text-xs text-slate-800">${Number(summary.total_clientes).toLocaleString()}</td>
        <td class="py-3.5 px-3.5 text-right font-mono text-sm text-emerald-700 font-black">$${Number(summary.total_cartera_usd).toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>
        <td class="py-3.5 px-3.5 text-right text-xs">
          ${footerPctDisplay}
        </td>
      </tr>
    `;
  }

  renderChart(chartLabels, chartValues, chartColors);
}

function renderGlobalTramosScorecard(globalSummary) {
  if (!globalSummary || !globalSummary.tramos || globalSummary.tramos.length === 0) return;
  currentGlobalSummary = globalSummary;
  updateQuickFilterBadges(globalSummary);

  renderExecutiveScorecard({
    summary: globalSummary,
    rows: globalSummary.tramos,
    title: 'Resumen Agrupado por Tramos de Mora',
    subtitle: 'Distribución consolidada de los 5 tramos sobre la Cartera Activa Total (fijo e independiente de los filtros).',
    col1Header: 'Fase / Categoría',
    col2Header: 'Rango / Etiqueta',
    chartTitleText: 'Distribución de Tarifa',
    chartTagClass: 'text-[10px] bg-blue-50 text-blue-700 px-2 py-0.5 rounded font-semibold border border-blue-200',
    chartTagText: 'Tramos Oficiales',
    excelBaseUrl: '/api/tramos/download/excel',
    excelLabel: 'Excel Tramos',
    excelTitle: 'Descargar Resumen de Tramos en formato Excel',
    pdfBaseUrl: '/api/tramos/download/pdf',
    pdfLabel: 'PDF Tramos',
    pdfTitle: 'Descargar Reporte PDF de Tramos',
    footerLabel: 'Total Tarifa',
    footerSublabel: 'Gestión total',
    buildRowMeta: (t) => ({
      chartLabel: t.etiqueta,
      rowTitle: `Clic para filtrar rápido de ${t.min_dias} a ${t.max_dias} días`,
      col1InnerHtml: `<span class="font-semibold text-slate-800">${t.fase}</span>`,
      col2Text: t.etiqueta,
      onClick: () => {
        setMoraRange(t.min_dias, Math.min(t.max_dias, 999), t.etiqueta);
        processFilters();
      }
    })
  });
}

function renderFibexScorecard(fibexSummary) {
  if (!fibexSummary || !fibexSummary.segments || fibexSummary.segments.length === 0) return;
  currentFibexSummary = fibexSummary;

  renderExecutiveScorecard({
    summary: fibexSummary,
    rows: fibexSummary.segments,
    title: 'Resumen Comparativo: Clientes Fibex vs. NO Fibex',
    subtitle: 'Distribución de cartera entre clientes sin servicio Fibex (NO Fibex) y clientes con suscripción Fibex / Resto.',
    col1Header: 'Segmento de Cartera',
    col2Header: 'Definición / Origen',
    chartTitleText: 'Proporción de Cartera',
    chartTagClass: 'text-[10px] bg-amber-50 text-amber-700 px-2 py-0.5 rounded font-semibold border border-amber-200',
    chartTagText: 'Fibex vs. NO Fibex',
    excelBaseUrl: '/api/fibex/download/excel',
    excelLabel: 'Excel Fibex',
    excelTitle: 'Descargar Resumen Fibex vs. NO Fibex en formato Excel',
    pdfBaseUrl: '/api/fibex/download/pdf',
    pdfLabel: 'PDF Fibex',
    pdfTitle: 'Descargar Reporte PDF Ejecutivo Fibex vs. NO Fibex',
    footerLabel: 'Total Cartera',
    footerSublabel: 'Gestión total evaluada',
    buildRowMeta: (s) => ({
      chartLabel: s.etiqueta,
      rowTitle: `Clic para ver ${s.etiqueta} (${s.porcentaje_mora}% de la morosidad)`,
      col1InnerHtml: `
        <div>
          <span class="font-bold text-slate-900">${s.etiqueta}</span>
          <span class="ml-1.5 text-[10px] font-mono text-slate-500 font-semibold">(${s.porcentaje_mora}% mora)</span>
        </div>
      `,
      col2Text: s.subtitulo,
      onClick: () => {
        const matrixSearch = document.getElementById('matrix-search-input');
        if (matrixSearch) {
          matrixSearch.value = s.id === 'no_fibex' ? 'NO Fibex' : '';
          filterMatrixTable(matrixSearch.value);
          matrixSearch.scrollIntoView({ behavior: 'smooth', block: 'center' });
        }
      }
    })
  });
}

function filterMatrixTable(query) {
  const q = (query || '').toLowerCase().trim();
  const rows = document.querySelectorAll('#matrix-campaigns-tbody tr');
  rows.forEach(tr => {
    const name = (tr.dataset.campana || '').toLowerCase();
    if (!q || name.includes(q)) {
      tr.style.display = '';
    } else {
      tr.style.display = 'none';
    }
  });
}

function renderCampanasMatrix(matrixData) {
  if (!matrixData || !matrixData.matrix) return;

  const countElem = document.getElementById('matrix-campaigns-count');
  if (countElem) {
    countElem.textContent = `${matrixData.matrix.length} campañas analizadas`;
  }

  const tbody = document.getElementById('matrix-campaigns-tbody');
  const tfoot = document.getElementById('matrix-campaigns-tfoot');
  if (!tbody) return;

  tbody.innerHTML = '';

  const tramoIds = ['1-30', '31-60', '61-90', '91-180', '181+'];

  function fmtMonto(val) {
    const n = Number(val || 0);
    if (n === 0) return '<span class="text-slate-300 font-normal">-</span>';
    return '$' + n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  function fmtCant(val) {
    const n = Number(val || 0);
    if (n === 0) return '<span class="text-slate-300 font-normal">-</span>';
    return n.toLocaleString();
  }

  matrixData.matrix.forEach(r => {
    const tr = document.createElement('tr');
    tr.className = 'hover:bg-slate-50/80 transition-colors border-b border-slate-100 group';
    tr.dataset.campana = r.campana;

    let tramosHtml = '';
    tramoIds.forEach(tId => {
      const t = r.tramos[tId] || { monto: 0, clientes: 0 };
      tramosHtml += `
        <td class="py-2 px-2.5 text-right font-mono border-r border-slate-100/70 text-slate-800">${fmtMonto(t.monto)}</td>
        <td class="py-2 px-2 text-right font-mono border-r border-slate-200 text-slate-600">${fmtCant(t.clientes)}</td>
      `;
    });

    tr.innerHTML = `
      <td class="py-2 px-3.5 font-bold text-slate-800 sticky left-0 bg-white group-hover:bg-slate-50 border-r border-slate-200 shadow-2xs">
        <span class="truncate max-w-[210px] inline-block align-middle" title="${r.campana}">${r.campana}</span>
      </td>
      ${tramosHtml}
      <td class="py-2 px-2.5 text-right font-mono font-bold text-slate-900 bg-slate-50/60 border-r border-slate-100">$${Number(r.total_monto).toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>
      <td class="py-2 px-2 text-right font-mono font-bold text-slate-800 bg-slate-50/60 border-r border-slate-100">${Number(r.total_clientes).toLocaleString()}</td>
      <td class="py-2 px-2.5 text-right font-mono font-semibold text-blue-700 bg-slate-50/60">${Number(r.porcentaje).toFixed(2)}%</td>
    `;
    tbody.appendChild(tr);
  });

  if (tfoot) {
    const totTramos = matrixData.totals_by_tramo || {};
    let footTramosHtml = '';
    tramoIds.forEach(tId => {
      const t = totTramos[tId] || { monto: 0, clientes: 0 };
      footTramosHtml += `
        <td class="py-2.5 px-2.5 text-right font-mono text-xs border-r border-slate-200 text-slate-900">$${Number(t.monto).toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>
        <td class="py-2.5 px-2 text-right font-mono text-xs border-r border-slate-300 text-slate-800">${Number(t.clientes).toLocaleString()}</td>
      `;
    });

    tfoot.innerHTML = `
      <tr class="bg-slate-100 border-t-2 border-slate-300 text-slate-900 font-bold text-xs">
        <td class="py-2.5 px-3.5 sticky left-0 bg-slate-100 border-r border-slate-300 text-slate-900 uppercase tracking-wider text-[11px]">
          <i class="fa-solid fa-layer-group mr-1.5 text-blue-600"></i> TOTAL TARIFA CONSOLIDADA
        </td>
        ${footTramosHtml}
        <td class="py-2.5 px-2.5 text-right font-mono text-xs text-emerald-700 font-black border-r border-slate-200">$${Number(matrixData.total_cartera_usd).toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>
        <td class="py-2.5 px-2 text-right font-mono text-xs text-slate-900 font-bold border-r border-slate-200">${Number(matrixData.total_clientes).toLocaleString()}</td>
        <td class="py-2.5 px-2.5 text-right font-mono text-xs text-blue-700 font-black">100.00%</td>
      </tr>
    `;
  }
}

function renderChart(labels, values, colorsList) {
  const ctx = document.getElementById('chart-agrupacion').getContext('2d');
  if (chartInstance) {
    chartInstance.destroy();
  }

  chartInstance = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [{
        label: 'Tarifa Vencida (USD)',
        data: values,
        backgroundColor: colorsList && colorsList.length ? colorsList : ['#10b981', '#3b82f6', '#f59e0b', '#f97316', '#ef4444'],
        borderRadius: 6
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false }
      },
      scales: {
        x: {
          ticks: {
            font: { size: 10 }
          }
        },
        y: {
          ticks: {
            callback: value => '$' + Number(value).toLocaleString()
          }
        }
      }
    }
  });
}

function renderPreviewTable(rows, totalFiltered) {
  document.getElementById('preview-tag').textContent = `Mostrando ${rows.length} de ${Number(totalFiltered).toLocaleString()} cuotas`;

  const thead = document.getElementById('preview-thead');
  const tbody = document.getElementById('preview-tbody');
  thead.innerHTML = '';
  tbody.innerHTML = '';

  if (!rows || rows.length === 0) {
    tbody.innerHTML = `<tr><td colspan="10" class="p-4 text-center text-slate-400">No hay registros con los filtros seleccionados</td></tr>`;
    return;
  }

  const displayCols = [
    { key: 'Cedula Cliente', label: 'Cédula' },
    { key: 'Nombre Cliente', label: 'Nombre Cliente' },
    { key: 'Dias de Mora', label: 'Días' },
    { key: 'Tramo_Mora', label: 'Tramo Mora' },
    { key: 'Pendiente Tarifa', label: 'Tarifa ($)' },
    { key: 'Pendiente Cuota', label: 'Cuota ($)' },
    { key: 'Pendiente Mora', label: 'Mora ($)' },
    { key: 'Total', label: 'Total ($)' },
    { key: 'Tipo_Campana', label: 'Tipo de Campaña' },
    { key: 'Programa', label: 'Programa' },
    { key: 'Telefono_1', label: 'Teléfono 1' },
    { key: 'Correo', label: 'Correo' }
  ];

  const trHead = document.createElement('tr');
  displayCols.forEach(col => {
    const th = document.createElement('th');
    th.className = 'py-2 px-3';
    th.textContent = col.label;
    trHead.appendChild(th);
  });
  thead.appendChild(trHead);

  rows.forEach(r => {
    const tr = document.createElement('tr');
    tr.className = 'hover:bg-blue-50/40 transition-colors';

    displayCols.forEach(col => {
      const td = document.createElement('td');
      td.className = 'py-1.5 px-3 border-b border-slate-100 text-[10px]';

      let val = r[col.key] !== undefined ? r[col.key] : '';

      if (col.key === 'Telefono_1' && r.wa_link) {
        td.innerHTML = `<a href="${r.wa_link}" target="_blank" class="inline-flex items-center gap-1 font-mono text-emerald-700 bg-emerald-50 px-1.5 py-0.5 rounded hover:bg-emerald-100 font-semibold"><i class="fa-brands fa-whatsapp text-emerald-600"></i> ${val}</a>`;
      } else if (['Pendiente Tarifa', 'Pendiente Cuota', 'Pendiente Mora', 'Total'].includes(col.key)) {
        td.className += ' font-mono text-right';
        td.textContent = '$' + Number(val || 0).toFixed(2);
      } else if (col.key === 'Dias de Mora') {
        td.className += ' font-mono text-center font-bold text-slate-800';
        td.textContent = val;
      } else {
        td.textContent = val;
      }
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
}

// ==============================================================
// PESTAÑA 4: CALL CENTER & WHATSAPP
// ==============================================================
async function loadCallCenterData() {
  const tbody = document.getElementById('callcenter-tbody');
  tbody.innerHTML = `<tr><td colspan="10" class="p-6 text-center text-slate-400"><i class="fa-solid fa-spinner fa-spin mr-2"></i> Cargando deudores y números telefónicos...</td></tr>`;

  try {
    const res = await fetch('/api/contacts/cross-match', { method: 'POST' });
    const data = await res.json();
    if (res.ok) {
      renderCallCenterView(data);
    } else {
      tbody.innerHTML = `<tr><td colspan="10" class="p-6 text-center text-rose-500">${data.detail}</td></tr>`;
    }
  } catch (err) {
    tbody.innerHTML = `<tr><td colspan="10" class="p-6 text-center text-rose-500">Error al cargar: ${err.message}</td></tr>`;
  }
}

function renderCallCenterView(data) {
  const k = data.kpis || {};
  document.getElementById('cc-kpi-total').textContent = Number(k.total_deudores || 0).toLocaleString();
  document.getElementById('cc-kpi-tel').textContent = Number(k.total_con_telefono || 0).toLocaleString();
  document.getElementById('cc-kpi-tel-pct').textContent = `${k.porcentaje_cobertura || 0}% de cobertura`;
  document.getElementById('cc-kpi-mail').textContent = Number(k.total_con_correo || 0).toLocaleString();
  document.getElementById('cc-kpi-monto').textContent = `$${Number(k.monto_contactable || 0).toLocaleString('en-US', {minimumFractionDigits: 2})} USD`;
  document.getElementById('cc-kpi-monto-pct').textContent = `${k.porcentaje_monto_contactable || 0}% del monto adeudado`;

  const tbody = document.getElementById('callcenter-tbody');
  tbody.innerHTML = '';

  if (!data.preview || data.preview.length === 0) {
    tbody.innerHTML = `<tr><td colspan="10" class="p-6 text-center text-slate-400">No hay deudores en el tramo actual.</td></tr>`;
    return;
  }

  data.preview.forEach(r => {
    const tr = document.createElement('tr');
    tr.className = 'hover:bg-blue-50/40 transition-colors';

    const tel1Html = r.wa_link
      ? `<a href="${r.wa_link}" target="_blank" class="inline-flex items-center gap-1 font-mono text-emerald-700 bg-emerald-50 px-2 py-0.5 rounded hover:bg-emerald-100 font-semibold" title="Abrir chat de WhatsApp">
           <i class="fa-brands fa-whatsapp text-emerald-600"></i> ${r.telefono_1}
         </a>`
      : (r.telefono_1 ? `<span class="font-mono text-slate-700">${r.telefono_1}</span>` : `<span class="text-slate-300">-</span>`);

    const tel2Html = r.telefono_2 ? `<span class="font-mono text-slate-600">${r.telefono_2}</span>` : `<span class="text-slate-300">-</span>`;
    const mailHtml = r.correo ? `<span class="font-mono text-[11px] text-blue-600 truncate max-w-[150px] inline-block" title="${r.correo}">${r.correo}</span>` : `<span class="text-slate-300">-</span>`;

    tr.innerHTML = `
      <td class="py-2.5 px-3 font-mono font-bold text-slate-800">${r.cedula}</td>
      <td class="py-2.5 px-3 font-medium text-slate-900">${r.nombre}</td>
      <td class="py-2.5 px-3 text-right font-bold text-emerald-600 font-mono">$${Number(r.monto_total).toLocaleString('en-US', {minimumFractionDigits: 2})}</td>
      <td class="py-2.5 px-3 text-center font-mono">${r.cuotas}</td>
      <td class="py-2.5 px-3 text-center font-mono font-bold text-slate-700">${r.dias_mora}d</td>
      <td class="py-2.5 px-3"><span class="px-2 py-0.5 rounded text-[10px] font-semibold bg-blue-50 text-blue-700 border border-blue-200">${r.tramo_mora}</span></td>
      <td class="py-2.5 px-3 text-slate-600 font-medium">${r.tipo_campana}</td>
      <td class="py-2.5 px-3">${tel1Html}</td>
      <td class="py-2.5 px-3">${tel2Html}</td>
      <td class="py-2.5 px-3">${mailHtml}</td>
    `;
    tbody.appendChild(tr);
  });
}

function toggleGuideModal(show) {
  const modal = document.getElementById('guide-modal');
  if (show) {
    modal.classList.remove('hidden');
    document.body.style.overflow = 'hidden';
  } else {
    modal.classList.add('hidden');
    document.body.style.overflow = 'auto';
  }
}
