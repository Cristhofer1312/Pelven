import React, { useState, useEffect, useMemo, useRef } from 'react';

const PELVEN_FACTORES = [
  0.9524, 0.9048, 0.8571, 0.8095, 0.7619, 0.7143, 0.6667, 0.6190, 0.5714, 0.5238,
  0.4762, 0.4286, 0.3810, 0.3333, 0.2857, 0.2381, 0.1905, 0.1429, 0.0952, 0.0476
];

const emptyMaestroForm = () => ({
  nro: '',
  circuito: '',
  administradora: '',
  complejo: '',
  estado: '',
  ciudad: '',
  sala: '',
  aforo: '',
});

export default function App() {
  const [maestro, setMaestro] = useState([]);
  const [selectedFiles, setSelectedFiles] = useState([]);
  const [isDragging, setIsDragging] = useState(false);
  const [loading, setLoading] = useState(false);
  const [exporting, setExporting] = useState(false);
  const [result, setResult] = useState(null);
  const [toast, setToast] = useState(null);
  const [exportarFiltrado, setExportarFiltrado] = useState(false);

  // Maestro CNAC (CRUD persistente)
  const [maestroQuery, setMaestroQuery] = useState('');
  const [maestroForm, setMaestroForm] = useState(emptyMaestroForm);
  const [maestroEditIdx, setMaestroEditIdx] = useState(null);
  const [savingMaestro, setSavingMaestro] = useState(false);
  const [runs, setRuns] = useState([]);
  const [loadingRuns, setLoadingRuns] = useState(false);
  
  // Parámetros de cálculo
  const [anio, setAnio] = useState(2026);
  const [factorDiario, setFactorDiario] = useState(4.20);
  
  // Filtros de navegación
  const [activeTab, setActiveTab] = useState('oficial'); // 'oficial' | 'analitico' | 'maestro' | 'auditoria'
  const [searchTerm, setSearchTerm] = useState('');
  const [filterCircuito, setFilterCircuito] = useState('TODOS');
  const [filterComplejo, setFilterComplejo] = useState('TODOS');
  const [incluirTodosCircuitos, setIncluirTodosCircuitos] = useState(false);
  
  const fileInputRef = useRef(null);

  const cargarMaestro = async () => {
    try {
      const res = await fetch('/api/maestro');
      const data = await res.json();
      setMaestro(Array.isArray(data) ? data : []);
    } catch {
      setMaestro([]);
    }
  };

  const cargarRuns = async () => {
    setLoadingRuns(true);
    try {
      const res = await fetch('/api/runs');
      const data = await res.json();
      setRuns(Array.isArray(data) ? data : []);
    } catch {
      setRuns([]);
    } finally {
      setLoadingRuns(false);
    }
  };

  useEffect(() => {
    cargarMaestro();
  }, []);

  useEffect(() => {
    if (activeTab === 'auditoria') cargarRuns();
  }, [activeTab, result?.manifiesto?.id]);

  const showToast = (text, type = 'info') => {
    setToast({ text, type });
    setTimeout(() => setToast(null), 4500);
  };

  // Manejador de selección de archivos
  const handleFileChange = (e) => {
    if (e.target.files?.length) {
      setSelectedFiles(Array.from(e.target.files));
    }
  };

  const handleDragOver = (e) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = () => {
    setIsDragging(false);
  };

  const handleDrop = (e) => {
    e.preventDefault();
    setIsDragging(false);
    if (e.dataTransfer.files?.length) {
      setSelectedFiles(Array.from(e.dataTransfer.files));
    }
  };

  // Enviar archivo al backend para cálculo
  const procesarReporte = async (e) => {
    if (e) e.preventDefault();
    if (!selectedFiles.length) {
      showToast('Por favor selecciona al menos un archivo de taquilla', 'error');
      return;
    }

    setLoading(true);
    showToast('Procesando datos de taquilla...', 'info');

    try {
      const fd = new FormData();
      selectedFiles.forEach((f) => fd.append('files', f));
      fd.append('anio', String(anio));
      fd.append('factorDiario', String(factorDiario));
      fd.append('incluirTodosCircuitos', String(incluirTodosCircuitos));

      const res = await fetch('/api/upload', {
        method: 'POST',
        body: fd
      });

      const data = await res.json();
      if (!res.ok || data.error) {
        throw new Error(data.error || 'Error en el servidor al procesar el reporte');
      }

      setResult(data);
      setExportarFiltrado(false);
      showToast(`¡Procesado exitoso! ${data.totalFilas.toLocaleString()} filas y ${data.salasDetectadas} salas`, 'success');
    } catch (err) {
      showToast(err.message, 'error');
    } finally {
      setLoading(false);
    }
  };

  // Descargar consolidado en Excel oficial. Si "exportarFiltrado" está activo,
  // se exportan solo las filas visibles con el filtro/búsqueda actual.
  const descargarExcel = async () => {
    const filas = exportarFiltrado ? filasFiltradas : result?._full;
    if (!filas?.length) {
      showToast('No hay filas para exportar con el filtro actual', 'error');
      return;
    }
    if (result.bloqueado) {
      showToast('Exportación bloqueada: corrige los errores de validación antes de descargar', 'error');
      return;
    }
    setExporting(true);
    try {
      const sufijo = exportarFiltrado ? 'filtrado' : '';
      const res = await fetch('/api/export', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ filas, anio, bloqueado: result.bloqueado, sufijo })
      });
      if (res.status === 409) {
        const data = await res.json().catch(() => ({}));
        throw new Error(data.error || 'Exportación bloqueada por errores de validación');
      }
      if (!res.ok) throw new Error('Error al exportar archivo');
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const nombre = `Pelven_aplica_${anio}${sufijo ? '_' + sufijo : ''}.xlsx`;
      const a = document.createElement('a');
      a.href = url;
      a.download = nombre;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
      showToast(`Descarga iniciada: ${nombre}${exportarFiltrado ? ' (solo filas filtradas)' : ''}`, 'success');
    } catch (err) {
      showToast('Error al descargar el archivo: ' + err.message, 'error');
    } finally {
      setExporting(false);
    }
  };

  // Descarga el manifiesto completo de una corrida (trazabilidad).
  const descargarManifiesto = async (id = result?.manifiesto?.id) => {
    if (!id) return;
    try {
      const res = await fetch(`/api/runs/${encodeURIComponent(id)}`);
      if (!res.ok) throw new Error('No se encontró el manifiesto');
      const data = await res.json();
      const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `manifiesto_${id}.json`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
      showToast('Manifiesto descargado', 'success');
    } catch (err) {
      showToast('Error al descargar el manifiesto: ' + err.message, 'error');
    }
  };

  const resetMaestroForm = () => {
    setMaestroForm(emptyMaestroForm());
    setMaestroEditIdx(null);
  };

  const editarMaestro = (item, idx) => {
    setMaestroForm({
      nro: item.nro ?? '',
      circuito: item.circuito ?? '',
      administradora: item.administradora ?? '',
      complejo: item.complejo ?? '',
      estado: item.estado ?? '',
      ciudad: item.ciudad ?? '',
      sala: item.sala ?? '',
      aforo: item.aforo ?? '',
    });
    setMaestroEditIdx(idx);
  };

  const guardarMaestro = async (e) => {
    e?.preventDefault();
    if (!String(maestroForm.complejo || '').trim() || !String(maestroForm.sala || '').trim()) {
      showToast('Complejo y sala son obligatorios', 'error');
      return;
    }
    setSavingMaestro(true);
    try {
      const esEdicion = maestroEditIdx != null;
      const res = await fetch(esEdicion ? `/api/maestro/${maestroEditIdx}` : '/api/maestro', {
        method: esEdicion ? 'PUT' : 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(maestroForm),
      });
      const data = await res.json().catch(() => ({}));
      if (!res.ok) throw new Error(data.error || 'No se pudo guardar la sala');
      await cargarMaestro();
      resetMaestroForm();
      showToast(esEdicion ? 'Sala actualizada en el maestro' : 'Sala agregada al maestro', 'success');
    } catch (err) {
      showToast(err.message, 'error');
    } finally {
      setSavingMaestro(false);
    }
  };

  const eliminarMaestro = async (item, idx) => {
    if (!window.confirm(`¿Eliminar "${item.complejo}" sala ${item.sala} del maestro?`)) return;
    try {
      const res = await fetch(`/api/maestro/${idx}`, { method: 'DELETE' });
      const data = await res.json().catch(() => ({}));
      if (!res.ok) throw new Error(data.error || 'No se pudo eliminar la sala');
      await cargarMaestro();
      resetMaestroForm();
      showToast('Sala eliminada del maestro', 'success');
    } catch (err) {
      showToast(err.message, 'error');
    }
  };

  // Listados únicos de circuitos y complejos para filtros rápidos
  const circuitosDisponibles = useMemo(() => {
    if (!result?._full) return [];
    const setCirc = new Set(result._full.map((r) => r.circuito).filter(Boolean));
    return ['TODOS', ...Array.from(setCirc).sort()];
  }, [result]);

  const complejosDisponibles = useMemo(() => {
    if (!result?._full) return [];
    const filteredByCirc = filterCircuito === 'TODOS'
      ? result._full
      : result._full.filter((r) => r.circuito === filterCircuito);
    const setComp = new Set(filteredByCirc.map((r) => r.complejo).filter(Boolean));
    return ['TODOS', ...Array.from(setComp).sort()];
  }, [result, filterCircuito]);

  // Filtrado de filas en la tabla
  const filasFiltradas = useMemo(() => {
    if (!result?._full) return [];
    return result._full.filter((r) => {
      const matchesSearch =
        !searchTerm ||
        r.circuito?.toLowerCase().includes(searchTerm.toLowerCase()) ||
        r.administradora?.toLowerCase().includes(searchTerm.toLowerCase()) ||
        r.complejo?.toLowerCase().includes(searchTerm.toLowerCase()) ||
        String(r.sala).toLowerCase().includes(searchTerm.toLowerCase()) ||
        r.ciudad?.toLowerCase().includes(searchTerm.toLowerCase()) ||
        r.estado?.toLowerCase().includes(searchTerm.toLowerCase());

      const matchesCircuito =
        filterCircuito === 'TODOS' || r.circuito === filterCircuito;

      const matchesComplejo =
        filterComplejo === 'TODOS' || r.complejo === filterComplejo;

      return matchesSearch && matchesCircuito && matchesComplejo;
    });
  }, [result, searchTerm, filterCircuito, filterComplejo]);

  const maestroFiltrado = useMemo(() => {
    const q = maestroQuery.trim().toLowerCase();
    if (!q) return maestro;
    return maestro.filter((m) =>
      [m.nro, m.circuito, m.administradora, m.complejo, m.estado, m.ciudad, m.sala].some((v) =>
        String(v ?? '').toLowerCase().includes(q)
      )
    );
  }, [maestro, maestroQuery]);

  const duplicados = result?.duplicados || { grupos: 0, filasExtra: 0, detalle: [] };
  const cobertura = result?.cobertura || null;
  const coberturaOk = cobertura ? cobertura.completa !== false : true;
  const manifiesto = result?.manifiesto || null;

  const maestroCampos = [
    { key: 'nro', label: 'Nro', type: 'number' },
    { key: 'circuito', label: 'Circuito' },
    { key: 'administradora', label: 'Administradora' },
    { key: 'complejo', label: 'Complejo' },
    { key: 'estado', label: 'Estado' },
    { key: 'ciudad', label: 'Ciudad' },
    { key: 'sala', label: 'Sala' },
    { key: 'aforo', label: 'Aforo', type: 'number' },
  ];

  // Cálculos de KPI ejecutivos
  const kpis = useMemo(() => {
    if (!result?._full) return null;
    const totalEspectadores = result._full.reduce((acc, r) => acc + (r.totalEspectadores || 0), 0);
    const promBase = result._full.length
      ? result._full.reduce((acc, r) => acc + (r.base || 0), 0) / result._full.length
      : 0;
    return {
      totalFilas: result.totalFilas || 0,
      totalSalas: result._full.length,
      totalComplejos: complejosDisponibles.length - 1,
      totalEspectadores,
      promBase: promBase.toFixed(1),
      errores: result.totalErrores || 0
    };
  }, [result, complejosDisponibles]);

  const sinMaestro = useMemo(
    () => (result?._full ? result._full.filter((r) => !r.enMaestro).length : 0),
    [result]
  );

  const salasSinData = useMemo(
    () => (result?._full ? result._full.filter((r) => r.sinData).length : 0),
    [result]
  );

  const auditOk = result ? result.totalErrores === 0 : true;

  const rangoGlobal = useMemo(() => {
    if (!result?.archivos?.length) return '';
    const fecha = (f) => {
      const m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(f || '');
      return m ? new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1])) : null;
    };
    const fechas = result.archivos.flatMap((a) => (a.rango || []).map(fecha)).filter(Boolean).sort((a, b) => a - b);
    if (!fechas.length) return '';
    const fmt = (d) => `${String(d.getDate()).padStart(2, '0')}/${String(d.getMonth() + 1).padStart(2, '0')}/${d.getFullYear()}`;
    return `${fmt(fechas[0])} al ${fmt(fechas[fechas.length - 1])}`;
  }, [result]);

  return (
    <div className="app-container">
      {/* Header Bar */}
      <header className="header-bar">
        <div className="brand-wrapper">
          <div className="brand-logo">
            <svg width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="#ffffff" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <rect x="2" y="2" width="20" height="20" rx="2.18" ry="2.18"></rect>
              <line x1="7" y1="2" x2="7" y2="22"></line>
              <line x1="17" y1="2" x2="17" y2="22"></line>
              <line x1="2" y1="12" x2="22" y2="12"></line>
              <line x1="2" y1="7" x2="7" y2="7"></line>
              <line x1="2" y1="17" x2="7" y2="17"></line>
              <line x1="17" y1="17" x2="22" y2="17"></line>
              <line x1="17" y1="7" x2="22" y2="7"></line>
            </svg>
          </div>
          <div>
            <div className="brand-title">
              PELVEN <span className="brand-badge">CNAC Oficial</span>
            </div>
            <div className="brand-sub">
              Sistema de Determinación de Promedio de Salas Cinematográficas · Providencia N° 002
            </div>
          </div>
        </div>

        <div className="header-status-pills">
          <div className="status-pill active">
            <span className="dot-indicator"></span> Motor Activo
          </div>
          <div className="status-pill">
            Aplica Año: <strong>{anio}</strong>
          </div>
          <div className="status-pill">
            Factor: <strong>{Number(factorDiario).toFixed(2)}</strong> (Vesp+Inter+Noche)
          </div>
          <div className="status-pill">
            Redondeo: <strong>≥ 0.70</strong> (al entero mayor)
          </div>
          <div className="status-pill">
            Circuito(s): <strong>{result?.circuitosDetectados?.join(', ') || 'Multicircuito soportado'}</strong>
          </div>
        </div>
      </header>

      {/* Main Upload & Configuration Card */}
      <section className="glass-panel upload-card">
        <div
          className={`drop-zone ${isDragging ? 'dragover' : ''}`}
          onDragOver={handleDragOver}
          onDragLeave={handleDragLeave}
          onDrop={handleDrop}
          onClick={() => fileInputRef.current?.click()}
        >
          <input
            ref={fileInputRef}
            type="file"
            multiple
            accept=".xls,.xlsx,.csv,.ods,.tsv,.txt"
            onChange={handleFileChange}
            style={{ display: 'none' }}
          />
          <div className="drop-icon-circle">
            <svg width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
              <polyline points="17 8 12 3 7 8"></polyline>
              <line x1="12" y1="3" x2="12" y2="15"></line>
            </svg>
          </div>
          <div className="drop-title">
            Arrastra el reporte de taquilla semanal o anual aquí
          </div>
          <div className="drop-sub">
            Formatos soportados: <code>.csv</code> (comas o punto y coma), <code>.ods</code> (LibreOffice/OpenDocument), <code>.xls</code> o <code>.xlsx</code>
          </div>

          {selectedFiles.length > 0 && (
            <div className="file-chip-container" onClick={(e) => e.stopPropagation()}>
              {selectedFiles.map((f, i) => (
                <div key={i} className="file-chip">
                  <span>📄 {f.name}</span>
                  <span style={{ opacity: 0.6 }}>({(f.size / (1024 * 1024)).toFixed(1)} MB)</span>
                </div>
              ))}
            </div>
          )}
        </div>

        {/* Configuration Toolbar */}
        <div className="controls-bar">
          <div className="inputs-group">
            <div className="input-field-wrapper">
              <span className="input-label">Año a Publicar</span>
              <input
                type="number"
                className="custom-input"
                style={{ width: 110 }}
                value={anio}
                onChange={(e) => setAnio(Number(e.target.value))}
              />
            </div>

            <div className="input-field-wrapper">
              <span className="input-label">Factor Diario Calibrado</span>
              <input
                type="number"
                step="0.05"
                className="custom-input"
                style={{ width: 120 }}
                value={factorDiario}
                onChange={(e) => setFactorDiario(Number(e.target.value))}
              />
            </div>

            <label style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: 12.5, cursor: 'pointer', color: 'var(--text-muted)', marginTop: 18 }}>
              <input
                type="checkbox"
                checked={incluirTodosCircuitos}
                onChange={(e) => setIncluirTodosCircuitos(e.target.checked)}
                style={{ width: 16, height: 16, accentColor: 'var(--wine-primary)' }}
              />
              <span>Incluir plantilla nacional completa (323 salas de los 4 circuitos)</span>
            </label>
          </div>

          <div className="actions-group">
            <button
              type="button"
              className="btn-primary"
              disabled={loading || !selectedFiles.length}
              onClick={procesarReporte}
            >
              {loading ? (
                <>
                  <span className="spinner"></span> Procesando reporte...
                </>
              ) : (
                <>
                  <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                    <polygon points="5 3 19 12 5 21 5 3"></polygon>
                  </svg>
                  Procesar y Generar Reporte
                </>
              )}
            </button>

            <button
              type="button"
              className="btn-secondary"
              onClick={() => setActiveTab('maestro')}
            >
              🏛️ Maestro CNAC
            </button>

            {result && (
              <button
                type="button"
                className="btn-download"
                disabled={exporting || result?.bloqueado}
                onClick={descargarExcel}
              >
                {exporting ? (
                  <>
                    <span className="spinner"></span> Generando XLSX...
                  </>
                ) : (
                  <>
                    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
                      <polyline points="7 10 12 15 17 10"></polyline>
                      <line x1="12" y1="15" x2="12" y2="3"></line>
                    </svg>
                    Descargar Excel Oficial (.xlsx)
                  </>
                )}
              </button>
            )}
          </div>
        </div>

        {/* Regulatory Info Callout */}
        <div className="info-callout">
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" style={{ flexShrink: 0, marginTop: 2 }}>
            <circle cx="12" cy="12" r="10"></circle>
            <line x1="12" y1="16" x2="12" y2="12"></line>
            <line x1="12" y1="8" x2="12.01" y2="8"></line>
          </svg>
          <div>
            <strong>Regla de Cálculo Pelven:</strong> Promedio diario de espectadores en funciones <em>Vespertina</em>, <em>Intermedia</em> y <em>Noche</em> multiplicado por el factor calibrado <code>4.20</code> (=21/5). <strong>Criterio de Redondeo:</strong> los cálculos se redondean al entero mayor a partir de <code>.70</code> (ej. 10.70 → 11; 10.69 → 10). Las columnas 20 a 1 aplican los factores porcentuales reglamentarios oficiales.
          </div>
        </div>
      </section>

      {result?.bloqueado && (
        <div className="info-callout" style={{ marginTop: 16, borderColor: 'rgba(248, 113, 113, 0.5)', background: 'rgba(248, 113, 113, 0.08)' }}>
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#f87171" strokeWidth="2" style={{ flexShrink: 0, marginTop: 2 }}>
            <path d="M10.29 3.86 1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"></path>
            <line x1="12" y1="9" x2="12" y2="13"></line>
            <line x1="12" y1="17" x2="12.01" y2="17"></line>
          </svg>
          <div>
            <strong>Exportación bloqueada:</strong> {result.totalErrores} inconsistencias en {result.filasInvalidas} filas.
            Las filas inválidas no se incluyeron en el cálculo. Corrige el archivo fuente y vuelve a procesar.
          </div>
        </div>
      )}

      {/* KPI Dashboard Metrics (when result available) */}
      {kpis && (
        <section className="stats-grid">
          <div className="glass-panel stat-card">
            <div className="stat-header">
              <span className="stat-label">Filas de Taquilla</span>
              <div className="stat-icon">📊</div>
            </div>
            <div className="stat-value">{kpis.totalFilas.toLocaleString()}</div>
            <div className="stat-caption">
              {result.filasValidas?.toLocaleString()} válidas · {result.filasInvalidas || 0} excluidas por error
            </div>
          </div>

          <div className="glass-panel stat-card">
            <div className="stat-header">
              <span className="stat-label">Salas Calculadas</span>
              <div className="stat-icon">🎬</div>
            </div>
            <div className="stat-value">{kpis.totalSalas}</div>
            <div className="stat-caption">En {kpis.totalComplejos} complejos cinematográficos</div>
          </div>

          <div className="glass-panel stat-card">
            <div className="stat-header">
              <span className="stat-label">Espectadores Totales</span>
              <div className="stat-icon">🎟️</div>
            </div>
            <div className="stat-value">{kpis.totalEspectadores.toLocaleString()}</div>
            <div className="stat-caption">Asistencia anual acumulada</div>
          </div>

          <div className="glass-panel stat-card">
            <div className="stat-header">
              <span className="stat-label">Promedio Base (Col 21)</span>
              <div className="stat-icon">📈</div>
            </div>
            <div className="stat-value">{kpis.promBase}</div>
            <div className="stat-caption">Espectadores promedio sala/día</div>
          </div>
        </section>
      )}

      {/* Results Workspace with Tabs */}
      {(result || activeTab === 'maestro') && (
        <section className="glass-panel" style={{ padding: '24px 0 0 0' }}>
          {/* Navigation Tabs */}
          <div style={{ padding: '0 24px' }}>
            <div className="tabs-header">
              {result && (
                <>
                  <button
                    className={`tab-btn ${activeTab === 'oficial' ? 'active' : ''}`}
                    onClick={() => setActiveTab('oficial')}
                  >
                    <span>📑 Formato Oficial Excel</span>
                    <span className="tab-badge">{filasFiltradas.length}</span>
                  </button>

                  <button
                    className={`tab-btn ${activeTab === 'analitico' ? 'active' : ''}`}
                    onClick={() => setActiveTab('analitico')}
                  >
                    <span>🔍 Desglose de Asistencia</span>
                  </button>
                </>
              )}

              <button
                className={`tab-btn ${activeTab === 'maestro' ? 'active' : ''}`}
                onClick={() => setActiveTab('maestro')}
              >
                <span>🏛️ Maestro CNAC</span>
                <span className="tab-badge">{maestro.length}</span>
              </button>

              {result && (
                <button
                  className={`tab-btn ${activeTab === 'auditoria' ? 'active' : ''}`}
                  onClick={() => setActiveTab('auditoria')}
                >
                  <span>🛡️ Auditoría ({result.totalErrores})</span>
                </button>
              )}
            </div>
          </div>

          {/* Table Search & Filter Toolbar */}
          {result && (activeTab === 'oficial' || activeTab === 'analitico') && (
            <div className="table-toolbar">
              <div className="search-box">
                <span className="search-icon">🔍</span>
                <input
                  type="text"
                  placeholder="Buscar complejo, sala, ciudad o estado..."
                  className="search-input"
                  value={searchTerm}
                  onChange={(e) => setSearchTerm(e.target.value)}
                />
              </div>

              <div style={{ display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap' }}>
                <span className="input-label" style={{ margin: 0 }}>Circuito:</span>
                <select
                  className="custom-input"
                  style={{ maxWidth: 170 }}
                  value={filterCircuito}
                  onChange={(e) => {
                    setFilterCircuito(e.target.value);
                    setFilterComplejo('TODOS');
                  }}
                >
                  {circuitosDisponibles.map((c, i) => (
                    <option key={i} value={c}>{c}</option>
                  ))}
                </select>

                <span className="input-label" style={{ margin: 0 }}>Complejo:</span>
                <select
                  className="custom-input"
                  style={{ maxWidth: 220 }}
                  value={filterComplejo}
                  onChange={(e) => setFilterComplejo(e.target.value)}
                >
                  {complejosDisponibles.map((c, i) => (
                    <option key={i} value={c}>
                      {c}
                    </option>
                  ))}
                </select>

                <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 12, color: 'var(--text-muted)', cursor: 'pointer' }}>
                  <input
                    type="checkbox"
                    checked={exportarFiltrado}
                    onChange={(e) => setExportarFiltrado(e.target.checked)}
                    style={{ width: 15, height: 15, accentColor: 'var(--wine-primary)' }}
                  />
                  Solo filas filtradas ({filasFiltradas.length})
                </label>

                <button
                  type="button"
                  className="btn-download"
                  style={{ padding: '8px 16px', fontSize: 13 }}
                  onClick={descargarExcel}
                  disabled={exporting || result?.bloqueado}
                >
                  📥 Exportar (.xlsx)
                </button>
              </div>
            </div>
          )}

          {/* Tab 1: Official Excel View */}
          {result && activeTab === 'oficial' && (
            <div className="table-scroll-wrapper">
              <table className="pelven-table">
                <thead>
                  <tr>
                    <th rowSpan={3} style={{ width: 50 }}>Nro</th>
                    <th rowSpan={3} style={{ width: 130 }}>Circuito</th>
                    <th rowSpan={3} style={{ width: 220 }}>Administradora</th>
                    <th rowSpan={3} style={{ width: 180 }}>Complejos</th>
                    <th rowSpan={3} style={{ width: 100 }}>Estado</th>
                    <th rowSpan={3} style={{ width: 100 }}>Ciudad</th>
                    <th rowSpan={3} style={{ width: 50 }}>Salas</th>
                    <th rowSpan={3} style={{ width: 60 }}>Aforo</th>
                    <th colSpan={21}>FUNCIONES</th>
                  </tr>
                  <tr className="secondary-header">
                    <th style={{ background: '#501322' }}>21</th>
                    {Array.from({ length: 20 }, (_, i) => (
                      <th key={i}>{20 - i}</th>
                    ))}
                  </tr>
                  <tr className="factor-header">
                    <th>100 % PELVEN</th>
                    {PELVEN_FACTORES.map((f, i) => (
                      <th key={i}>{(f * 100).toFixed(2)}%</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {filasFiltradas.map((r, i) => (
                    <tr key={i}>
                      <td className="col-center" style={{ color: 'var(--text-dim)' }}>{i + 1}</td>
                      <td>{r.circuito}</td>
                      <td style={{ maxWidth: 220, overflow: 'hidden', textOverflow: 'ellipsis' }}>{r.administradora}</td>
                      <td style={{ fontWeight: 600, color: '#fff' }}>{r.complejo}</td>
                      <td className="col-center">{r.estado}</td>
                      <td className="col-center">{r.ciudad}</td>
                      <td className="col-center" style={{ fontWeight: 600 }}>{r.sala}</td>
                      <td className="col-num">{r.aforo || 0}</td>
                      {/* Columna Base FUNCIONES destacada */}
                      <td className="col-highlight">{r.base.toLocaleString()}</td>
                      {Array.from({ length: 20 }, (_, k) => {
                        const val = r['f' + (20 - k)] || 0;
                        return (
                          <td key={k} className="col-num">
                            {val.toLocaleString()}
                          </td>
                        );
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}

          {/* Tab 2: Detailed Analytics Breakdown */}
          {result && activeTab === 'analitico' && (
            <div className="table-scroll-wrapper">
              <table className="pelven-table">
                <thead>
                  <tr>
                    <th>#</th>
                    <th>Complejo</th>
                    <th>Sala</th>
                    <th>Aforo</th>
                    <th>Espectadores Totales</th>
                    <th>Espectadores Base (V+I+N)</th>
                    <th>Días con Función</th>
                    <th>Días Base</th>
                    <th>Semanas</th>
                    <th>Factor</th>
                    <th>Base Oficial (21)</th>
                    <th>f20 (95.24%)</th>
                    <th>f10 (47.62%)</th>
                    <th>f1 (4.76%)</th>
                    <th>Maestro</th>
                  </tr>
                </thead>
                <tbody>
                  {filasFiltradas.map((r, i) => (
                    <tr key={i}>
                      <td className="col-center">{i + 1}</td>
                      <td style={{ fontWeight: 600 }}>{r.complejo}</td>
                      <td className="col-center">{r.sala}</td>
                      <td className="col-num">{r.aforo}</td>
                      <td className="col-num">{r.totalEspectadores.toLocaleString()}</td>
                      <td className="col-num" style={{ color: '#38bdf8' }}>{Math.round(r.totalBase || 0).toLocaleString()}</td>
                      <td className="col-center">{r.diasActivos}</td>
                      <td className="col-center">{r.diasBase}</td>
                      <td className="col-center">{r.semanasConData}</td>
                      <td className="col-center">{r.factorUsado || 4.2}</td>
                      <td className="col-highlight">
                        {r.base} <span style={{ fontSize: 10, opacity: 0.6 }}>({r.rawBase?.toFixed(2)})</span>
                      </td>
                      <td className="col-num">{r.f20}</td>
                      <td className="col-num">{r.f10}</td>
                      <td className="col-num">{r.f1}</td>
                      <td className="col-center">
                        <span style={{ color: r.enMaestro ? '#34d399' : '#f87171', fontWeight: 600 }}>
                          {r.enMaestro ? '✓ Mapeado' : '✗ Sin maestro'}
                        </span>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}

          {/* Tab 3: Maestro CNAC Directory (CRUD persistente) */}
          {activeTab === 'maestro' && (
            <div style={{ padding: 24 }}>
              <div style={{ marginBottom: 16, display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 16, flexWrap: 'wrap' }}>
                <div>
                  <h3 style={{ fontSize: 16, fontWeight: 700 }}>Directorio Nacional de Salas CNAC</h3>
                  <p style={{ fontSize: 13, color: 'var(--text-muted)' }}>
                    {maestro.length} salas registradas. Los cambios se guardan en <code>maestro.json</code> y aplican a la próxima corrida.
                  </p>
                </div>
                <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                  {!result && (
                    <button type="button" className="btn-secondary" style={{ padding: '9px 14px' }} onClick={() => setActiveTab('oficial')}>
                      Volver al inicio
                    </button>
                  )}
                  <input
                    type="text"
                    className="custom-input"
                    placeholder="Buscar sala, complejo, circuito..."
                    value={maestroQuery}
                    onChange={(e) => setMaestroQuery(e.target.value)}
                    style={{ minWidth: 240 }}
                  />
                  <button type="button" className="btn-secondary" onClick={resetMaestroForm} style={{ padding: '9px 14px' }}>
                    + Nueva sala
                  </button>
                </div>
              </div>

              <form onSubmit={guardarMaestro} style={{ marginBottom: 18, padding: 16, background: 'rgba(255,255,255,0.03)', borderRadius: 12, border: '1px solid var(--border-subtle)' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 10 }}>
                  <strong style={{ fontSize: 13 }}>
                    {maestroEditIdx != null ? `Editando sala #${maestroEditIdx}` : 'Agregar sala al maestro'}
                  </strong>
                  {maestroEditIdx != null && (
                    <button type="button" className="btn-secondary" style={{ padding: '5px 10px', fontSize: 12 }} onClick={resetMaestroForm}>
                      Cancelar edición
                    </button>
                  )}
                </div>
                <div className="maestro-form-grid">
                  {maestroCampos.map((campo) => (
                    <div key={campo.key} className="input-field-wrapper">
                      <span className="input-label">{campo.label}</span>
                      <input
                        type={campo.type || 'text'}
                        className="custom-input"
                        value={maestroForm[campo.key] ?? ''}
                        onChange={(e) => setMaestroForm((prev) => ({ ...prev, [campo.key]: e.target.value }))}
                      />
                    </div>
                  ))}
                  <div style={{ display: 'flex', alignItems: 'flex-end' }}>
                    <button type="submit" className="btn-primary" style={{ padding: '10px 18px' }} disabled={savingMaestro}>
                      {savingMaestro ? 'Guardando…' : maestroEditIdx != null ? 'Guardar cambios' : 'Agregar sala'}
                    </button>
                  </div>
                </div>
              </form>

              <div className="table-scroll-wrapper" style={{ maxHeight: 480 }}>
                <table className="pelven-table">
                  <thead>
                    <tr>
                      <th>Nro</th>
                      <th>Circuito</th>
                      <th>Administradora</th>
                      <th>Complejo</th>
                      <th>Estado</th>
                      <th>Ciudad</th>
                      <th>Sala</th>
                      <th>Aforo</th>
                      <th>Acciones</th>
                    </tr>
                  </thead>
                  <tbody>
                    {maestroFiltrado.map((m) => {
                      const idx = maestro.indexOf(m);
                      return (
                        <tr key={`${m.nro}-${idx}`}>
                          <td className="col-center">{m.nro || idx + 1}</td>
                          <td style={{ fontWeight: 600 }}>{m.circuito}</td>
                          <td>{m.administradora}</td>
                          <td style={{ color: '#fff' }}>{m.complejo}</td>
                          <td className="col-center">{m.estado}</td>
                          <td className="col-center">{m.ciudad}</td>
                          <td className="col-center">{m.sala}</td>
                          <td className="col-num">{m.aforo}</td>
                          <td className="col-center">
                            <div className="row-actions">
                              <button type="button" className="mini-btn" onClick={() => editarMaestro(m, idx)}>
                                Editar
                              </button>
                              <button type="button" className="mini-btn danger" onClick={() => eliminarMaestro(m, idx)}>
                                Eliminar
                              </button>
                            </div>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
              {maestroFiltrado.length === 0 && (
                <p style={{ marginTop: 12, fontSize: 12, color: 'var(--text-dim)', textAlign: 'center' }}>
                  No hay salas que coincidan con la búsqueda.
                </p>
              )}
            </div>
          )}

          {/* Tab 4: Integrity & Validation Audit */}
          {result && activeTab === 'auditoria' && (
            <div style={{ padding: 24 }}>
              <h3 style={{ fontSize: 16, fontWeight: 700, marginBottom: 8 }}>
                Auditoría de Integridad y Validación
              </h3>
              <p style={{ fontSize: 13, color: 'var(--text-muted)', marginBottom: 20 }}>
                Verificación estricta de consistencia en boletos, horas de función, tarifas e impuestos según especificaciones técnicas CNAC.
              </p>

              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(280px, 1fr))', gap: 16, marginBottom: 24 }}>
                <div style={{ padding: 16, background: auditOk ? 'rgba(16, 185, 129, 0.08)' : 'rgba(248, 113, 113, 0.08)', borderRadius: 12, border: `1px solid ${auditOk ? 'rgba(16, 185, 129, 0.2)' : 'rgba(248, 113, 113, 0.35)'}` }}>
                  <div style={{ color: auditOk ? '#34d399' : '#f87171', fontWeight: 700, marginBottom: 4 }}>
                    {auditOk ? '✓' : '✗'} Recaudación vs Boletos
                  </div>
                  <div style={{ fontSize: 12, color: 'var(--text-muted)' }}>
                    {auditOk
                      ? `Sin inconsistencias en ${result.filasValidas} filas válidas.`
                      : `${result.totalErrores} inconsistencias en ${result.filasInvalidas} filas excluidas del cálculo.`}
                  </div>
                </div>

                <div style={{ padding: 16, background: coberturaOk ? 'rgba(56, 189, 248, 0.08)' : 'rgba(251, 191, 36, 0.08)', borderRadius: 12, border: `1px solid ${coberturaOk ? 'rgba(56, 189, 248, 0.2)' : 'rgba(251, 191, 36, 0.35)'}` }}>
                  <div style={{ color: coberturaOk ? '#38bdf8' : '#fbbf24', fontWeight: 700, marginBottom: 4 }}>
                    {coberturaOk ? '✓' : '⚠'} Cobertura / Rango de fechas
                  </div>
                  <div style={{ fontSize: 12, color: 'var(--text-muted)' }}>
                    {cobertura
                      ? `${cobertura.detectado?.desde || '?'} al ${cobertura.detectado?.hasta || '?'}${cobertura.esperado ? ` · declarado: ${cobertura.esperado.desde} al ${cobertura.esperado.hasta}` : ''}${cobertura.totalFaltantes > 0 ? ` · faltan ${cobertura.totalFaltantes} día(s)` : ''}`
                      : rangoGlobal || 'Sin rango disponible'}
                  </div>
                </div>

                <div style={{ padding: 16, background: sinMaestro ? 'rgba(251, 191, 36, 0.08)' : 'rgba(16, 185, 129, 0.08)', borderRadius: 12, border: `1px solid ${sinMaestro ? 'rgba(251, 191, 36, 0.35)' : 'rgba(16, 185, 129, 0.2)'}` }}>
                  <div style={{ color: sinMaestro ? '#fbbf24' : '#34d399', fontWeight: 700, marginBottom: 4 }}>
                    {sinMaestro ? '⚠' : '✓'} Salas Identificadas
                  </div>
                  <div style={{ fontSize: 12, color: 'var(--text-muted)' }}>
                    {result.salasDetectadas} salas procesadas; {sinMaestro} sin maestro; {salasSinData} sin datos.
                  </div>
                </div>

                <div style={{ padding: 16, background: duplicados.grupos ? 'rgba(251, 191, 36, 0.08)' : 'rgba(16, 185, 129, 0.08)', borderRadius: 12, border: `1px solid ${duplicados.grupos ? 'rgba(251, 191, 36, 0.35)' : 'rgba(16, 185, 129, 0.2)'}` }}>
                  <div style={{ color: duplicados.grupos ? '#fbbf24' : '#34d399', fontWeight: 700, marginBottom: 4 }}>
                    {duplicados.grupos ? '⚠' : '✓'} Posibles duplicados exactos
                  </div>
                  <div style={{ fontSize: 12, color: 'var(--text-muted)' }}>
                    {duplicados.grupos
                      ? `${duplicados.grupos.toLocaleString()} grupo(s) · ${duplicados.filasExtra.toLocaleString()} fila(s) repetidas. Informativo; no se excluyen automáticamente.`
                      : 'No se detectaron filas idénticas.'}
                  </div>
                </div>

                <div style={{ padding: 16, background: 'rgba(56, 189, 248, 0.08)', borderRadius: 12, border: '1px solid rgba(56, 189, 248, 0.2)' }}>
                  <div style={{ color: '#38bdf8', fontWeight: 700, marginBottom: 4 }}>🧾 Manifiesto de la corrida</div>
                  <div style={{ fontSize: 12, color: 'var(--text-muted)', marginBottom: 8 }}>
                    {manifiesto
                      ? `${manifiesto.id} · ${manifiesto.fecha ? new Date(manifiesto.fecha).toLocaleString() : ''}${manifiesto.guardado === false ? ' · no persistido' : ''}`
                      : 'Sin manifiesto disponible'}
                  </div>
                  {manifiesto?.id && (
                    <button type="button" className="btn-secondary" style={{ padding: '5px 10px', fontSize: 12 }} disabled={manifiesto.guardado === false} onClick={() => descargarManifiesto()}>
                      Descargar JSON
                    </button>
                  )}
                </div>
              </div>

              {result.archivos?.length > 0 && (
                <div style={{ marginBottom: 24 }}>
                  <h4 style={{ marginBottom: 8 }}>Archivos procesados</h4>
                  <div className="table-scroll-wrapper">
                    <table className="pelven-table">
                      <thead>
                        <tr>
                          <th>Archivo</th>
                          <th>SHA-256</th>
                          <th>Filas</th>
                          <th>Válidas</th>
                          <th>Inválidas</th>
                          <th>Errores</th>
                          <th>Avisos</th>
                          <th>Semanas</th>
                          <th>Rango</th>
                        </tr>
                      </thead>
                      <tbody>
                        {result.archivos.map((a, i) => (
                          <tr key={i}>
                            <td style={{ fontWeight: 600 }}>{a.file}</td>
                            <td className="col-center" style={{ fontFamily: 'monospace', fontSize: 11 }} title={a.sha256 || ''}>
                              {a.sha256 ? `${a.sha256.slice(0, 12)}…` : '—'}
                            </td>
                            <td className="col-num">{a.filas?.toLocaleString()}</td>
                            <td className="col-num">{a.filasValidas?.toLocaleString()}</td>
                            <td className="col-num" style={{ color: a.filasInvalidas ? '#f87171' : undefined }}>{a.filasInvalidas}</td>
                            <td className="col-num" style={{ color: a.errores ? '#f87171' : undefined }}>{a.errores}</td>
                            <td className="col-num">{a.avisos}</td>
                            <td className="col-num">{a.semanas}</td>
                            <td className="col-center">{a.rango?.join(' al ')}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}

              {cobertura && (cobertura.totalFaltantes > 0 || cobertura.esperado) && (
                <div style={{ marginBottom: 24 }}>
                  <h4 style={{ color: cobertura.totalFaltantes ? '#fbbf24' : '#38bdf8', marginBottom: 8 }}>
                    Cobertura declarada vs detectada
                  </h4>
                  <p style={{ fontSize: 13, color: 'var(--text-muted)', marginBottom: 8 }}>
                    Detectado: {cobertura.detectado?.desde || '—'} al {cobertura.detectado?.hasta || '—'}
                    {cobertura.esperado ? ` · Declarado en el nombre: ${cobertura.esperado.desde} al ${cobertura.esperado.hasta}` : ''}
                  </p>
                  {cobertura.totalFaltantes > 0 && (
                    <div style={{ padding: 12, borderRadius: 8, background: 'rgba(251, 191, 36, 0.08)', border: '1px solid rgba(251, 191, 36, 0.3)', fontSize: 12, color: '#fbbf24' }}>
                      Faltan {cobertura.totalFaltantes} día(s): {cobertura.faltantes.join(', ')}
                      {cobertura.totalFaltantes > cobertura.faltantes.length ? ' …' : ''}
                    </div>
                  )}
                </div>
              )}

              {duplicados.detalle?.length > 0 && (
                <div style={{ marginBottom: 24 }}>
                  <h4 style={{ color: '#fbbf24', marginBottom: 8 }}>
                    Posibles duplicados exactos ({duplicados.grupos.toLocaleString()} grupos · {duplicados.filasExtra.toLocaleString()} filas extra)
                  </h4>
                  <p style={{ fontSize: 12, color: 'var(--text-muted)', marginBottom: 8 }}>
                    Informativo: una misma combinación puede provenir de lotes de venta distintos. No se excluye del cálculo ni bloquea la exportación.
                  </p>
                  <div className="table-scroll-wrapper" style={{ maxHeight: 320 }}>
                    <table className="pelven-table">
                      <thead>
                        <tr>
                          <th>Fecha</th>
                          <th>Complejo</th>
                          <th>Sala</th>
                          <th>Hora</th>
                          <th>Función</th>
                          <th>Tipo</th>
                          <th>Cantidad</th>
                          <th>Veces</th>
                          <th>Origen</th>
                        </tr>
                      </thead>
                      <tbody>
                        {duplicados.detalle.map((d, i) => (
                          <tr key={i}>
                            <td className="col-center">{d.fecha}</td>
                            <td>{d.complejo}</td>
                            <td className="col-center">{d.sala}</td>
                            <td className="col-center">{d.hora}</td>
                            <td className="col-center">{d.funcion}</td>
                            <td className="col-center">{d.tipoBoleto}</td>
                            <td className="col-num">{d.cantidad}</td>
                            <td className="col-center" style={{ fontWeight: 700, color: '#fbbf24' }}>{d.ocurrencias}</td>
                            <td style={{ fontSize: 11, color: 'var(--text-muted)' }}>
                              {d.filas.map((f, j) => `${f.archivo || '—'}:${f.linea}`).join(' · ')}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                  {duplicados.detalleTruncado && (
                    <p style={{ marginTop: 8, fontSize: 12, color: 'var(--text-dim)' }}>
                      Mostrando los primeros {duplicados.detalle.length} grupos. El manifiesto conserva el resumen completo.
                    </p>
                  )}
                </div>
              )}

              {manifiesto && (
                <div style={{ marginBottom: 24 }}>
                  <h4 style={{ marginBottom: 8 }}>Manifiesto de la corrida</h4>
                  <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))', gap: 10, fontSize: 12, color: 'var(--text-muted)' }}>
                    <div><strong>ID:</strong> {manifiesto.id}</div>
                    <div><strong>Versión:</strong> {manifiesto.version}</div>
                    <div><strong>Fecha:</strong> {manifiesto.fecha ? new Date(manifiesto.fecha).toLocaleString() : '—'}</div>
                    <div><strong>Bloqueado:</strong> {manifiesto.salidas?.bloqueado ? 'Sí' : 'No'}</div>
                    <div><strong>Filas:</strong> {manifiesto.totales?.filas?.toLocaleString()}</div>
                    <div><strong>Errores:</strong> {manifiesto.totales?.errores}</div>
                    <div><strong>Avisos:</strong> {manifiesto.totales?.avisos}</div>
                    <div><strong>Salas:</strong> {manifiesto.salidas?.salas}</div>
                  </div>
                  <div style={{ marginTop: 10 }}>
                    <button type="button" className="btn-secondary" style={{ padding: '6px 12px', fontSize: 12 }} disabled={manifiesto.guardado === false} onClick={() => descargarManifiesto()}>
                      Descargar manifiesto JSON
                    </button>
                  </div>
                </div>
              )}

              <div style={{ marginBottom: 24 }}>
                <h4 style={{ marginBottom: 8 }}>Corridas recientes</h4>
                {loadingRuns ? (
                  <p style={{ fontSize: 12, color: 'var(--text-muted)' }}>Cargando historial…</p>
                ) : runs.length === 0 ? (
                  <p style={{ fontSize: 12, color: 'var(--text-muted)' }}>Sin corridas persistidas todavía.</p>
                ) : (
                  <div className="table-scroll-wrapper" style={{ maxHeight: 260 }}>
                    <table className="pelven-table">
                      <thead>
                        <tr>
                          <th>Fecha</th>
                          <th>ID</th>
                          <th>Archivos</th>
                          <th>Filas</th>
                          <th>Errores</th>
                          <th>Salas</th>
                          <th>Estado</th>
                          <th>Manifiesto</th>
                        </tr>
                      </thead>
                      <tbody>
                        {runs.map((r) => (
                          <tr key={r.id}>
                            <td className="col-center">{r.fecha ? new Date(r.fecha).toLocaleString() : '—'}</td>
                            <td style={{ fontSize: 11, fontFamily: 'monospace' }}>{r.id}</td>
                            <td className="col-center">{r.archivos?.length || 0}</td>
                            <td className="col-num">{r.totales?.filas?.toLocaleString?.() ?? r.totales?.filas}</td>
                            <td className="col-num" style={{ color: r.totales?.errores ? '#f87171' : undefined }}>{r.totales?.errores}</td>
                            <td className="col-num">{r.salidas?.salas}</td>
                            <td className="col-center">{r.bloqueado ? 'Bloqueada' : 'OK'}</td>
                            <td className="col-center">
                              <button type="button" className="mini-btn" onClick={() => descargarManifiesto(r.id)}>
                                JSON
                              </button>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>

              {result.totalAvisos > 0 && (
                <div style={{ marginBottom: 24 }}>
                  <h4 style={{ color: '#fbbf24', marginBottom: 8 }}>Avisos ({result.totalAvisos})</h4>
                  <pre style={{ background: '#111827', padding: 16, borderRadius: 8, fontSize: 12, overflow: 'auto', maxHeight: 280 }}>
                    {JSON.stringify(result.avisos, null, 2)}
                  </pre>
                </div>
              )}

              {result.totalErrores > 0 && (
                <div>
                  <h4 style={{ color: '#f87171', marginBottom: 8 }}>
                    Inconsistencias Detectadas ({result.totalErrores}) · exportación bloqueada
                  </h4>
                  <pre style={{ background: '#111827', padding: 16, borderRadius: 8, fontSize: 12, overflow: 'auto', maxHeight: 280 }}>
                    {JSON.stringify(result.errores, null, 2)}
                  </pre>
                </div>
              )}
            </div>
          )}
        </section>
      )}

      {/* Floating Toast Notification */}
      {toast && (
        <div className={`toast-msg ${toast.type}`}>
          <span>{toast.type === 'error' ? '⚠️' : toast.type === 'success' ? '✅' : 'ℹ️'}</span>
          <span>{toast.text}</span>
        </div>
      )}
    </div>
  );
}
