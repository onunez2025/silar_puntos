import React, { useState, useCallback } from 'react';
import * as XLSX from 'xlsx';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import { Upload, FileText, Download, CheckCircle2, AlertCircle, Loader2, Calendar, User, Hash } from 'lucide-react';
import JSZip from 'jszip';

interface ServiceRow {
  id: string;
  customer: string;
  technician: string;
  equipment: string;
  date: string;
  workDone: string;
  address: string;
  city: string;
  phone: string;
  remarks: string;
  points: number;
  rawData: any;
}

const calculatePoints = (equipment: string, serviceType: string): number => {
  const eq = equipment.toUpperCase();
  const st = serviceType.toUpperCase();
  
  if (st.includes('INSTALACIÓN') || st.includes('INSTALACION') || st.includes('INST_')) {
    // Rule: TERMAS MAYORES DE 80LT -> 2 PUNTOS
    if (eq.includes('TERMA')) {
      const match = eq.match(/(\d+)\s*LT/);
      if (match && parseInt(match[1]) >= 80) return 2;
    }
    
    // Rule: CAMPANAS DECORATIVAS -> 2 PUNTOS
    if (eq.includes('CAMPANA') && (eq.includes('DECORATIVA') || eq.includes('VENUS') || eq.includes('INKA') || eq.includes('ISLA'))) {
      return 2;
    }
  }
  
  // Rule: TODO LO DEMAS -> 1 PUNTO
  return 1;
};

function App() {
  const [data, setData] = useState<ServiceRow[]>([]);
  const [isDragging, setIsDragging] = useState(false);
  const [loading, setLoading] = useState(false);
  const [status, setStatus] = useState<{ type: 'success' | 'error', msg: string } | null>(null);
  
  // Filters
  const [selectedTechnician, setSelectedTechnician] = useState<string>('');
  const [selectedDate, setSelectedDate] = useState<string>('');
  const [atsNumber, setAtsNumber] = useState<string>('');
  const [seriesAts, setSeriesAts] = useState<string>('');

  const handleFileUpload = (file: File) => {
    setLoading(true);
    setStatus(null);
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const bstr = e.target?.result;
        const workbook = XLSX.read(bstr, { type: 'binary' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(worksheet);
        const mappedData: ServiceRow[] = json.map((row: any) => {
          const equipment = row['Nombre del equipo'] || 'N/A';
          const serviceType = row['Tipo de servicio (ACT UDF)'] || row['Tipo de servicio actualizado FSM (SC UDF)'] || 'N/A';
          return {
            id: row['ID externo de llamada de servicio'] || 'N/A',
            customer: row['Nombre del cliente'] || 'N/A',
            technician: row['Técnico'] || 'N/A',
            equipment,
            date: row['Fecha de servicio (ACT UDF)'] || row['Fecha de inicio planificada']?.split(' ')[0] || 'N/A',
            workDone: serviceType,
            address: row['Calle'] || 'N/A',
            city: row['Ciudad'] || 'N/A',
            phone: row['Celular de comunicación del cliente C4 (SC UDF)'] || 'N/A',
            remarks: row['Referencia (ACT UDF)'] || row['Service Call Remarks'] || 'N/A',
            points: calculatePoints(equipment, serviceType),
            rawData: row
          };
        });

        setData(mappedData);
        setStatus({ type: 'success', msg: `Se cargaron ${mappedData.length} servicios correctamente.` });
      } catch (err) {
        console.error(err);
        setStatus({ type: 'error', msg: 'Error al procesar el archivo Excel.' });
      } finally {
        setLoading(false);
      }
    };

    reader.readAsBinaryString(file);
  };

  const onDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    const file = e.dataTransfer.files[0];
    if (file && (file.name.endsWith('.xlsx') || file.name.endsWith('.xls'))) {
      handleFileUpload(file);
    } else {
      setStatus({ type: 'error', msg: 'Por favor sube un archivo Excel (.xlsx o .xls).' });
    }
  }, []);

  const generateAtsPdf = async (technician: string, date: string, services: ServiceRow[]) => {
    const doc = new jsPDF({
      orientation: 'portrait',
      unit: 'mm',
      format: 'a4'
    });

    const margin = 10;
    const pageWidth = 210;
    let y = 10;

    // Header Table
    doc.setDrawColor(0);
    doc.setLineWidth(0.2);
    
    // Outer box for header
    doc.rect(margin, y, pageWidth - (margin * 2), 25);
    doc.line(margin + 45, y, margin + 45, y + 25); // Vertical after logo
    doc.line(margin + 130, y, margin + 130, y + 25); // Vertical before right info
    doc.line(margin + 45, y + 8, margin + 130, y + 8); // Horizontal in center
    doc.line(margin + 130, y + 8, margin + 190, y + 8); // Horizontal in right
    doc.line(margin + 130, y + 16, margin + 190, y + 16); // Horizontal 2 in right
    doc.line(margin + 160, y + 8, margin + 160, y + 25); // Vertical split in right

    // Logo area
    doc.setFontSize(10);
    doc.setFont('helvetica', 'bold');
    doc.text('SILAR', margin + 22.5, y + 10, { align: 'center' });
    doc.text('PERÚ', margin + 22.5, y + 15, { align: 'center' });
    doc.setFontSize(6);
    doc.text('Oil & Gas', margin + 22.5, y + 18, { align: 'center' });

    // Center header
    doc.setFontSize(12);
    doc.text('REGISTRO', margin + 87.5, y + 6, { align: 'center' });
    doc.setFontSize(10);
    doc.text('EXTENSIÓN DE ATS: SP-SIG-', margin + 87.5, y + 14, { align: 'center' });
    doc.text('REG- 030', margin + 87.5, y + 19, { align: 'center' });

    // Right header
    doc.setFontSize(8);
    doc.text('S/C', margin + 160, y + 5, { align: 'center' });
    doc.text('Fecha aprobada:', margin + 132, y + 12);
    doc.text('Versión', margin + 132, y + 21);
    doc.text('01', margin + 175, y + 21, { align: 'center' });
    doc.text('Página', margin + 132, y + 24.5); // Fixed version
    doc.text('01', margin + 175, y + 24.5, { align: 'center' });

    y += 25;

    // Subheader section
    doc.rect(margin, y, pageWidth - (margin * 2), 10);
    doc.line(margin + 30, y, margin + 30, y + 10);
    doc.line(margin + 130, y, margin + 130, y + 10);
    doc.line(margin + 155, y, margin + 155, y + 10);

    doc.setTextColor(255, 0, 0);
    doc.text(`N° ${atsNumber || '000000'}`, margin + 5, y + 6);
    doc.setTextColor(0);
    doc.setFontSize(7);
    doc.text('N° serie de ATS:', margin + 32, y + 4);
    doc.setFontSize(10);
    doc.text(seriesAts || '', margin + 55, y + 7);
    
    doc.setFontSize(8);
    doc.text('FECHA:', margin + 132, y + 6);
    doc.setFontSize(10);
    doc.text(date, margin + 172.5, y + 6, { align: 'center' });

    y += 10;

    // Main Table
    const tableBody = [];
    
    // Header for "EFECTUADOS"
    tableBody.push([{ content: 'N°', styles: { halign: 'center', fontStyle: 'bold' } }, { content: 'EFECTUADOS', styles: { halign: 'center', fontStyle: 'bold' } }, { content: 'PUNTOS', styles: { halign: 'center', fontStyle: 'bold' } }]);
    
    let totalPoints = 0;
    services.forEach((s, i) => {
      tableBody.push([i + 1, s.customer, s.points]);
      totalPoints += s.points;
    });

    // Fill remaining rows to reach at least 13 for "EFECTUADOS"
    for (let i = services.length; i < 13; i++) {
      tableBody.push([i + 1, '', '']);
    }

    tableBody.push([{ content: 'TOTAL', colSpan: 2, styles: { halign: 'right', fontStyle: 'bold' } }, { content: totalPoints, styles: { halign: 'center', fontStyle: 'bold' } }]);
    
    // "NO EFECTUADOS" section
    tableBody.push([{ content: 'NO EFECTUADOS', colSpan: 3, styles: { halign: 'center', fontStyle: 'bold' } }]);
    for (let i = 0; i < 5; i++) {
      tableBody.push([14 + i, '', '']);
    }

    autoTable(doc, {
      startY: y,
      body: tableBody,
      theme: 'grid',
      styles: { fontSize: 8, cellPadding: 1, textColor: [0, 0, 0], lineColor: [0, 0, 0], lineWidth: 0.2 },
      columnStyles: {
        0: { cellWidth: 10, halign: 'center' },
        1: { cellWidth: 155 },
        2: { cellWidth: 25, halign: 'center' }
      },
      margin: { left: margin },
      didDrawPage: (data) => {
        y = data.cursor?.y || y;
      }
    });

    y = (doc as any).lastAutoTable.finalY + 2;

    // Footer
    doc.rect(margin, y, pageWidth - (margin * 2), 20);
    doc.line(margin + 100, y, margin + 100, y + 20);
    doc.line(margin, y + 6, margin + pageWidth - (margin * 2), y + 6);

    doc.setFontSize(9);
    doc.text('Nombre del Responsable:', margin + 2, y + 4);
    doc.text(technician, margin + 45, y + 4);
    doc.text('Firma:', margin + 102, y + 4);

    y += 22;
    doc.setFontSize(7);
    doc.text('Nota: esta extensión es válido siempre y cuando sean las mismas actividades que se realizan durante el día, y no cambie', margin, y);
    doc.text('los integrantes de la cuadrilla de actividades.', margin, y + 3);

    // Save logic with File System Access API
    const safeTech = technician ? technician.replace(/[^a-zA-Z0-9]/g, '_') : 'Tecnico';
    const safeDate = date ? date.replace(/[^a-zA-Z0-9]/g, '_') : 'Fecha';
    const filename = `ATS_${safeTech}_${safeDate}.pdf`;
    
    try {
      if ('showSaveFilePicker' in window) {
        // Modern approach: Forces native "Save As" dialog and respects filename
        const handle = await (window as any).showSaveFilePicker({
          suggestedName: filename,
          types: [{
            description: 'Archivo PDF',
            accept: { 'application/pdf': ['.pdf'] },
          }],
        });
        const writable = await handle.createWritable();
        const pdfBlob = doc.output('blob');
        await writable.write(pdfBlob);
        await writable.close();
        setStatus({ type: 'success', msg: 'Archivo PDF guardado correctamente.' });
      } else {
        // Fallback
        doc.save(filename);
      }
    } catch (err: any) {
      if (err.name !== 'AbortError') {
        console.error(err);
        // Fallback in case of unexpected errors
        doc.save(filename);
      }
    }
  };

  const technicians = Array.from(new Set(data.map(d => d.technician))).sort();
  const dates = Array.from(new Set(data.map(d => d.date))).sort().reverse();
  
  const filteredData = data.filter(d => 
    (selectedTechnician === '' || d.technician === selectedTechnician) &&
    (selectedDate === '' || d.date === selectedDate)
  );

  return (
    <div className="container">
      <header style={{ marginBottom: '3rem', textAlign: 'center' }} className="animate-in">
        <h1 style={{ fontSize: '2.5rem', fontWeight: 800, marginBottom: '0.5rem', background: 'linear-gradient(135deg, #2563eb, #10b981)', WebkitBackgroundClip: 'text', WebkitTextFillColor: 'transparent' }}>
          Silar PDF Generator
        </h1>
        <p style={{ color: '#64748b', fontSize: '1.1rem' }}>
          Calculadora de puntos y generador de reportes FSM.
        </p>
      </header>

      <main className="glass-card animate-in" style={{ animationDelay: '0.1s' }}>
        {!data.length ? (
          <div 
            className={`upload-zone ${isDragging ? 'dragging' : ''}`}
            onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
            onDragLeave={() => setIsDragging(false)}
            onDrop={onDrop}
            onClick={() => document.getElementById('file-upload')?.click()}
          >
            <input 
              id="file-upload" 
              type="file" 
              accept=".xlsx,.xls" 
              onChange={(e) => e.target.files?.[0] && handleFileUpload(e.target.files[0])}
              style={{ display: 'none' }}
            />
            <Upload className="upload-icon" />
            <h2 style={{ marginBottom: '0.5rem' }}>Suelta tu archivo Excel aquí</h2>
            <p style={{ color: '#64748b' }}>o haz clic para seleccionar desde tu equipo</p>
          </div>
        ) : (
          <div>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: '2rem', flexWrap: 'wrap', gap: '1rem' }}>
              <div style={{ display: 'flex', gap: '1.5rem', alignItems: 'center', flexWrap: 'wrap' }}>
                <div>
                  <h2 style={{ fontSize: '1.5rem', marginBottom: '0.25rem' }}>Generar Extensión ATS</h2>
                  <span className="badge badge-success">
                    {data.length} registros cargados
                  </span>
                </div>
                
                {/* Custom Filters */}
                <div style={{ display: 'flex', gap: '1rem', background: 'rgba(255,255,255,0.5)', padding: '1rem', borderRadius: '16px', border: '1px solid var(--border)' }}>
                  <div className="filter-group">
                    <label style={{ display: 'block', fontSize: '0.75rem', fontWeight: 600, color: 'var(--text-muted)', marginBottom: '0.25rem' }}>TÉCNICO</label>
                    <div style={{ position: 'relative' }}>
                      <User size={14} style={{ position: 'absolute', left: '10px', top: '50%', transform: 'translateY(-50%)', color: 'var(--primary)' }} />
                      <select 
                        value={selectedTechnician} 
                        onChange={(e) => setSelectedTechnician(e.target.value)}
                        className="select-custom"
                        style={{ paddingLeft: '2.2rem' }}
                      >
                        <option value="">Seleccionar Técnico</option>
                        {technicians.map(t => <option key={t} value={t}>{t}</option>)}
                      </select>
                    </div>
                  </div>

                  <div className="filter-group">
                    <label style={{ display: 'block', fontSize: '0.75rem', fontWeight: 600, color: 'var(--text-muted)', marginBottom: '0.25rem' }}>FECHA</label>
                    <div style={{ position: 'relative' }}>
                      <Calendar size={14} style={{ position: 'absolute', left: '10px', top: '50%', transform: 'translateY(-50%)', color: 'var(--primary)' }} />
                      <select 
                        value={selectedDate} 
                        onChange={(e) => setSelectedDate(e.target.value)}
                        className="select-custom"
                        style={{ paddingLeft: '2.2rem' }}
                      >
                        <option value="">Seleccionar Fecha</option>
                        {dates.map(d => <option key={d} value={d}>{d}</option>)}
                      </select>
                    </div>
                  </div>

                  <div className="filter-group">
                    <label style={{ display: 'block', fontSize: '0.75rem', fontWeight: 600, color: 'var(--text-muted)', marginBottom: '0.25rem' }}>N° CORRELATIVO</label>
                    <div style={{ position: 'relative' }}>
                      <Hash size={14} style={{ position: 'absolute', left: '10px', top: '50%', transform: 'translateY(-50%)', color: 'var(--primary)' }} />
                      <input 
                        type="text"
                        placeholder="Ej. 003700"
                        value={atsNumber}
                        onChange={(e) => setAtsNumber(e.target.value)}
                        className="select-custom"
                        style={{ paddingLeft: '2.2rem', width: '120px' }}
                      />
                    </div>
                  </div>
                </div>
              </div>

              <div style={{ display: 'flex', gap: '1rem' }}>
                <button className="btn btn-outline" onClick={() => setData([])}>
                  Cambiar archivo
                </button>
                <button 
                  className="btn btn-primary" 
                  disabled={!selectedTechnician || !selectedDate || loading}
                  onClick={() => generateAtsPdf(selectedTechnician, selectedDate, filteredData)}
                >
                  {loading ? <Loader2 className="animate-spin" /> : <FileText size={20} />}
                  Descargar PDF
                </button>
              </div>
            </div>

            {selectedTechnician && selectedDate ? (
              <div className="animate-in">
                <div style={{ marginBottom: '1.5rem', display: 'flex', gap: '1rem' }}>
                  <div className="stats-card">
                    <p>Servicios Filtrados</p>
                    <h3>{filteredData.length}</h3>
                  </div>
                  <div className="stats-card">
                    <p>Total Puntos Día</p>
                    <h3>{filteredData.reduce((acc, s) => acc + s.points, 0)}</h3>
                  </div>
                </div>
                
                <div style={{ overflowX: 'auto' }}>
                  <table className="data-table">
                    <thead>
                      <tr>
                        <th>N°</th>
                        <th>Cliente (EFECTUADOS)</th>
                        <th>ID Servicio</th>
                        <th>Puntos</th>
                      </tr>
                    </thead>
                    <tbody>
                      {filteredData.map((row, idx) => (
                        <tr key={idx}>
                          <td>{idx + 1}</td>
                          <td style={{ fontWeight: 600 }}>{row.customer}</td>
                          <td style={{ color: 'var(--text-muted)' }}>{row.id}</td>
                          <td>
                            <span className={`points-badge ${row.points > 1 ? 'high' : 'normal'}`}>
                              {row.points}
                            </span>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            ) : (
              <div style={{ textAlign: 'center', padding: '4rem', background: 'rgba(0,0,0,0.02)', borderRadius: '24px', border: '2px dashed var(--border)' }}>
                <User size={48} style={{ color: 'var(--border)', marginBottom: '1rem' }} />
                <h3>Selecciona un técnico y una fecha</h3>
                <p style={{ color: 'var(--text-muted)' }}>Para visualizar los servicios y generar el reporte ATS</p>
              </div>
            )}
          </div>
        )}

        {status && (
          <div style={{ 
            marginTop: '2rem', 
            padding: '1rem', 
            borderRadius: '12px', 
            display: 'flex', 
            alignItems: 'center', 
            gap: '0.75rem',
            background: status.type === 'success' ? 'rgba(16, 185, 129, 0.1)' : 'rgba(239, 68, 68, 0.1)',
            color: status.type === 'success' ? '#059669' : '#dc2626',
            border: `1px solid ${status.type === 'success' ? 'rgba(16, 185, 129, 0.2)' : 'rgba(239, 68, 68, 0.2)'}`
          }}>
            {status.type === 'success' ? <CheckCircle2 size={20} /> : <AlertCircle size={20} />}
            {status.msg}
          </div>
        )}
      </main>

      <footer style={{ marginTop: 'auto', padding: '2rem 0', textAlign: 'center', color: '#64748b', fontSize: '0.875rem' }}>
        &copy; {new Date().getFullYear()} Silar Perú S.A.C. Todos los derechos reservados.
      </footer>
    </div>
  );
}

export default App;
