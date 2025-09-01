
import { useState, useRef } from 'react';
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

  const handleFileChange = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setFileName(file.name);
    setOcrText('');
    setProgress(0);
    setProcessing(true);
    let text = '';
    if (file.type === 'application/pdf') {
      text = await handlePdf(file);
    } else if (file.type === 'text/html' || file.name.endsWith('.html')) {
      text = await handleHtml(file);
    } else {
      setOcrText('Unsupported file type.');
      setProcessing(false);
      return;
    }
    setOcrText(text);
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

  const handlePdf = async (file) => {
    try {
      // Read PDF as ArrayBuffer
      const arrayBuffer = await file.arrayBuffer();
      const pdf = await getDocument({ data: arrayBuffer }).promise;
      let allRows = [];
      for (let i = 1; i <= pdf.numPages; i++) {
        const page = await pdf.getPage(i);
        const content = await page.getTextContent();
        // Group items by y position (line), then sort by x (column)
        const lines = {};
        content.items.forEach(item => {
          const y = Math.round(item.transform[5]);
          if (!lines[y]) lines[y] = [];
          lines[y].push(item);
        });
        const sortedLines = Object.values(lines)
          .sort((a, b) => b[0].transform[5] - a[0].transform[5]) // top to bottom
          .map(line => line.sort((a, b) => a.transform[4] - b.transform[4])) // left to right
          .map(line => line.map(item => item.str));
        allRows = allRows.concat(sortedLines);
        setProgress(Math.round((i / pdf.numPages) * 100));
      }
      setExtractedTables([allRows]);
      setIsTable(true);
      return 'Table-like structure extracted from PDF.';
    } catch (err) {
      setOcrText('PDF parsing error: ' + err.message);
      setProcessing(false);
      setIsTable(false);
      return '';
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

  // State for table extraction
  const [extractedTables, setExtractedTables] = useState([]);
  const [isTable, setIsTable] = useState(false);

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
    <div className="container">
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
      {ocrText && !processing && (
        <div>
          <h2>Result</h2>
          {isTable && extractedTables.length > 0 ? (
            <div style={{overflowX:'auto'}}>
              {extractedTables.map((table, idx) => (
                <table key={idx} border="1" cellPadding="4" style={{marginBottom:'1rem',width:'100%'}}>
                  <tbody>
                    {table.map((row, rIdx) => (
                      <tr key={rIdx}>
                        {row.map((cell, cIdx) => <td key={cIdx}>{cell}</td>)}
                      </tr>
                    ))}
                  </tbody>
                </table>
              ))}
            </div>
          ) : (
            <textarea value={ocrText} readOnly rows={10} style={{ width: '100%' }} />
          )}
          <button onClick={handleExport}>Export to Excel</button>
        </div>
      )}
    </div>
  );
}

export default App;
