
import { useState, useRef, useEffect } from 'react';
import Tesseract from 'tesseract.js';
import * as XLSX from 'xlsx';
import { getDocument, GlobalWorkerOptions } from 'pdfjs-dist';
import pdfjsWorker from 'pdfjs-dist/build/pdf.worker?url';
import './App.css';
// Set pdfjs-dist worker
GlobalWorkerOptions.workerSrc = pdfjsWorker;



function App() {
  const [ocrText, setOcrText] = useState('');
  const [progress, setProgress] = useState(0);
  const [processing, setProcessing] = useState(false);
  const [fileName, setFileName] = useState('');
  const fileInput = useRef();

  const [extractionMethod, setExtractionMethod] = useState('');
  const resultRef = useRef();
  const topScrollRef = useRef(null);
  const contentScrollRef = useRef(null);
  // State for table extraction (must be above useEffect)
  const [extractedTables, setExtractedTables] = useState([]);
  const [isTable, setIsTable] = useState(false);

  // Synchronize top and bottom scroll bars
  useEffect(() => {
    const top = topScrollRef.current;
    const content = contentScrollRef.current;
    if (!top || !content) return;
    const handleTopScroll = () => {
      content.scrollLeft = top.scrollLeft;
    };
    const handleContentScroll = () => {
      top.scrollLeft = content.scrollLeft;
    };
    top.addEventListener('scroll', handleTopScroll);
    content.addEventListener('scroll', handleContentScroll);
    return () => {
      top.removeEventListener('scroll', handleTopScroll);
      content.removeEventListener('scroll', handleContentScroll);
    };
  }, [ocrText, extractedTables, isTable]);

  useEffect(() => {
    if (ocrText && !processing && resultRef.current) {
      resultRef.current.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }
    // eslint-disable-next-line
  }, [ocrText, processing]);

  const handleFileChange = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setFileName(file.name);
    setOcrText('');
    setProgress(0);
    setProcessing(true);
    setExtractionMethod('');
    let method = '';
    if (file.type === 'application/pdf') {
      const { usedOcr } = await handlePdf(file, true);
      method = usedOcr ? 'OCR' : 'Text Extraction';
    } else if (file.type === 'text/html' || file.name.endsWith('.html')) {
      await handleHtml(file);
      method = 'Text Extraction';
    } else {
      setOcrText('Unsupported file type.');
      setProcessing(false);
      setExtractionMethod('');
      return;
    }
    setExtractionMethod(method);
    setProcessing(false);
  };

  const handleHtml = async (file) => {
    const html = await file.text();
    const doc = new DOMParser().parseFromString(html, 'text/html');
    // Try to extract tables
    const tables = Array.from(doc.querySelectorAll('table'));
    if (tables.length > 0) {
      // Extract all tables as arrays of arrays
      const allTables = tables.map(table => {
        return Array.from(table.rows).map(row =>
          Array.from(row.cells).map(cell => cell.textContent.trim())
        );
      });
      // Flatten to a single array if only one table, else keep as multiple sheets
      setExtractedTables(allTables);
      setIsTable(true);
      setProgress(100);
      return 'Table(s) extracted from HTML.';
    } else {
      // Fallback: extract all text
      const text = doc.body.textContent || '';
      setIsTable(false);
      setProgress(100);
      return text;
    }
  };

  // handlePdf now sets extracted data directly and returns { usedOcr }
  const handlePdf = async (file, setDataDirectly = false) => {
    try {
      const arrayBuffer = await file.arrayBuffer();
      const pdf = await getDocument({ data: arrayBuffer }).promise;
      let allRows = [];
      let usedOcr = false;
      for (let i = 1; i <= pdf.numPages; i++) {
        const page = await pdf.getPage(i);
        const content = await page.getTextContent();
        if (content.items && content.items.length > 0 && content.items.some(item => item.str && item.str.trim() !== '')) {
          const lines = {};
          content.items.forEach(item => {
            const y = Math.round(item.transform[5]);
            if (!lines[y]) lines[y] = [];
            lines[y].push(item);
          });
          const sortedLines = Object.values(lines)
            .sort((a, b) => b[0].transform[5] - a[0].transform[5])
            .map(line => line.sort((a, b) => a.transform[4] - b.transform[4]))
            .map(line => line.map(item => item.str));
          allRows = allRows.concat(sortedLines);
        } else {
          usedOcr = true;
          const viewport = page.getViewport({ scale: 2 });
          const canvas = document.createElement('canvas');
          canvas.width = viewport.width;
          canvas.height = viewport.height;
          const context = canvas.getContext('2d');
          await page.render({ canvasContext: context, viewport }).promise;
          const dataUrl = canvas.toDataURL('image/png');
          const ocrText = await runOcrOnImage(dataUrl, i, pdf.numPages);
          const ocrRows = ocrText.split('\n').map(line => [line]);
          allRows = allRows.concat(ocrRows);
        }
        setProgress(Math.round((i / pdf.numPages) * 100));
      }
      setExtractedTables([allRows]);
      setIsTable(true);
      // Set ocrText for textarea fallback if no table
      setOcrText(allRows.map(row => row.join(' ')).join('\n'));
      return { usedOcr };
    } catch (err) {
      setOcrText('PDF parsing error: ' + err.message);
      setProcessing(false);
      setIsTable(false);
      return { usedOcr: false };
    }
  };

  const runOcrOnImage = async (image, page, totalPages) => {
    return new Promise((resolve) => {
      Tesseract.recognize(image, 'eng', {
        logger: (m) => {
          if (m.status === 'recognizing text') {
            setProgress(Math.round((m.progress + (page - 1)) / totalPages * 100));
          }
        },
      }).then(({ data: { text } }) => {
        resolve(text);
      });
    });
  };

  const runOcrOnText = async (text) => {
    // For HTML, just return the text (no OCR needed)
    setProgress(100);
    return text;
  };


  const handleExport = () => {
    const wb = XLSX.utils.book_new();
    if (isTable && extractedTables.length > 0) {
      extractedTables.forEach((table, idx) => {
        const ws = XLSX.utils.aoa_to_sheet(table);
        XLSX.utils.book_append_sheet(wb, ws, `Table${idx + 1}`);
      });
    } else {
      const ws = XLSX.utils.aoa_to_sheet([[ocrText]]);
      XLSX.utils.book_append_sheet(wb, ws, 'OCR Text');
    }
    XLSX.writeFile(wb, 'ocr_result.xlsx');
  };

  return (
    <div style={{ minHeight: '100vh', width: '100%', background: '#f4f6fa', display: 'flex', flexDirection: 'column', alignItems: 'center' }}>
      <div style={{ width: '100%', maxWidth: 600, display: 'flex', flexDirection: 'column', alignItems: 'center', marginTop: '2rem' }}>
        <h1>Document OCR to Excel</h1>
        <input
          type="file"
          accept=".pdf,.html,application/pdf,text/html"
          ref={fileInput}
          onChange={handleFileChange}
          disabled={processing}
        />
        {fileName && <p>Selected file: {fileName}</p>}
        {processing && (
          <div>
            <p>Processing... {progress}%</p>
            <progress value={progress} max="100" />
          </div>
        )}
      </div>
      {ocrText && !processing && (
        <div style={{ width: '100%', display: 'flex', justifyContent: 'center', margin: '2rem 0' }} ref={resultRef}>
          <div style={{ width: '100%', maxWidth: 1200, background: '#fff', borderRadius: 18, boxShadow: '0 4px 24px rgba(0,0,0,0.10)', padding: '2.5rem 2rem', display: 'flex', flexDirection: 'column', alignItems: 'center' }}>
            {/* Top scroll bar */}
            <div
              ref={topScrollRef}
              style={{
                width: '100%',
                overflowX: 'auto',
                overflowY: 'hidden',
                height: 16,
                marginBottom: 0,
                position: 'relative',
                background: 'transparent',
                pointerEvents: 'auto',
              }}
            >
              {/* Fake content to create scroll bar */}
              <div style={{ width: isTable && extractedTables.length > 0
                ? Math.max(900, ...extractedTables.map(table => (table[0]?.length || 1) * 180))
                : 900, height: 1 }} />
            </div>
            <h2 style={{marginTop:0, fontWeight:600, fontSize:'2rem', color:'#222'}}>Result</h2>
            {extractionMethod && (
              <p style={{margin:'0 0 1.5rem 0', color:'#666'}}><strong>Extraction Method:</strong> {extractionMethod}</p>
            )}
            <button onClick={handleExport} style={{marginBottom:'2rem',padding:'0.75rem 2rem',fontSize:'1.1rem',fontWeight:600,background:'#2563eb',color:'#fff',border:'none',borderRadius:8,boxShadow:'0 2px 8px rgba(37,99,235,0.08)',cursor:'pointer',transition:'background 0.2s'}} onMouseOver={e=>e.target.style.background='#1741a6'} onMouseOut={e=>e.target.style.background='#2563eb'}>
              Export to Excel
            </button>
            <div
              ref={contentScrollRef}
              style={{ width: '100%', overflowX: 'auto', overflowY: 'visible' }}
            >
              {isTable && extractedTables.length > 0 ? (
                <div style={{width:'100%', maxWidth: '100%' }}>
                  {extractedTables.map((table, idx) => (
                    <table key={idx} style={{marginBottom:'2rem',width:'100%',borderCollapse:'collapse',background:'#fafbfc',borderRadius:10,overflow:'hidden',boxShadow:'0 1px 4px rgba(0,0,0,0.04)'}}>
                      <tbody>
                        {table.map((row, rIdx) => (
                          <tr key={rIdx}>
                            {row.map((cell, cIdx) => <td key={cIdx} style={{whiteSpace:'pre-wrap',padding:'0.5rem 1rem',border:'1px solid #e0e3ea',fontSize:'1rem',color:'#222'}}>{cell}</td>)}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  ))}
                </div>
              ) : (
                <textarea value={ocrText} readOnly rows={12} style={{ width: '100%', minHeight: 200, borderRadius: 10, border: '1px solid #e0e3ea', padding: '1rem', fontSize: '1.1rem', background:'#fafbfc', color:'#222', boxShadow:'0 1px 4px rgba(0,0,0,0.04)', resize:'vertical' }} />
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

export default App;
