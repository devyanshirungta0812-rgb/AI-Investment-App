import React, { useState, useEffect } from 'react';
import { 
  LayoutDashboard, 
  FileText, 
  MessageSquare, 
  TrendingUp, 
  ShieldCheck, 
  Search,
  ChevronRight,
  Loader2,
  ExternalLink,
  ArrowUpRight,
  ArrowDownRight,
  Send,
  User,
  Bot,
  Filter,
  RefreshCw,
  FileDown,
  FileText as FileIcon
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { geminiService, NCD, InvestmentMemo } from './services/geminiService';
import ReactMarkdown from 'react-markdown';
import { jsPDF } from 'jspdf';
import { Document, Packer, Paragraph, TextRun, ImageRun, HeadingLevel, AlignmentType } from 'docx';
import { saveAs } from 'file-saver';
import * as htmlToImage from 'html-to-image';

// --- Helpers ---

const safeString = (val: any): string => {
  if (typeof val === 'string') return val;
  if (val === null || val === undefined) return '';
  if (typeof val === 'object') {
    return Object.entries(val)
      .map(([k, v]) => `${k}: ${v}`)
      .join('\n');
  }
  return String(val);
};

// --- Components ---

const Sidebar = ({ activeTab, setActiveTab }: { activeTab: string, setActiveTab: (t: string) => void }) => {
  const menuItems = [
    { id: 'explorer', label: 'Market Explorer', icon: LayoutDashboard },
    { id: 'analyser', label: 'Deep Dive Analyser', icon: FileText },
    { id: 'assistant', label: 'Chat Assistant', icon: MessageSquare },
  ];

  return (
    <div className="w-64 border-r border-border h-screen flex flex-col bg-bg sticky top-0">
      <div className="p-6 border-b border-border">
        <div className="flex items-center gap-3">
          <div className="w-8 h-8 bg-accent rounded flex items-center justify-center">
            <TrendingUp className="text-white w-5 h-5" />
          </div>
          <div>
            <h1 className="font-bold text-sm tracking-tight text-accent">NBHI</h1>
            <p className="text-[10px] text-text-dim uppercase tracking-widest">Investment Pro</p>
          </div>
        </div>
      </div>
      <nav className="flex-1 p-4 space-y-2">
        {menuItems.map((item) => (
          <button
            key={item.id}
            onClick={() => setActiveTab(item.id)}
            className={`w-full flex items-center gap-3 px-4 py-3 rounded-lg transition-all text-sm font-medium ${
              activeTab === item.id 
                ? 'bg-accent/10 text-accent border border-accent/20' 
                : 'text-slate-600 hover:bg-slate-100 hover:text-accent'
            }`}
          >
            <item.icon size={18} />
            {item.label}
          </button>
        ))}
      </nav>
      <div className="p-6 border-t border-border">
        <div className="flex items-center gap-3 p-3 rounded-lg bg-slate-50">
          <div className="w-8 h-8 rounded-full bg-slate-200 flex items-center justify-center text-xs text-slate-700">
            ALM
          </div>
          <div className="flex-1 min-w-0">
            <p className="text-xs font-semibold truncate text-slate-900">Investment Team</p>
            <p className="text-[10px] text-text-dim">Niva Bupa</p>
          </div>
        </div>
      </div>
    </div>
  );
};

const MarketExplorer = ({ onSelectBond }: { onSelectBond: (bond: NCD) => void }) => {
  const [bonds, setBonds] = useState<NCD[]>([]);
  const [loading, setLoading] = useState(true);

  const fetchBonds = async () => {
    setLoading(true);
    const data = await geminiService.getMarketScan();
    setBonds(data);
    setLoading(false);
  };

  useEffect(() => {
    fetchBonds();
  }, []);

  if (loading) {
    return (
      <div className="flex-1 flex items-center justify-center">
        <div className="flex flex-col items-center gap-4">
          <Loader2 className="w-8 h-8 animate-spin text-accent" />
          <p className="text-sm text-text-dim animate-pulse">Scanning Debt Markets...</p>
        </div>
      </div>
    );
  }

  return (
    <div className="flex-1 p-8 overflow-auto">
      <div className="flex justify-between items-end mb-8">
        <div>
          <div className="flex items-center gap-3 mb-2">
            <h2 className="text-2xl font-bold tracking-tight text-accent">Live Market Scan</h2>
            <span className="px-2 py-0.5 rounded bg-accent/10 text-accent text-[10px] font-bold uppercase tracking-wider border border-accent/20">
              Live Data Connection Active
            </span>
          </div>
          <p className="text-sm text-text-dim">
            Sources: <span className="text-slate-600 font-medium">IndiaBonds | NSE Debt | GoldenPi | Jiraaf</span>
          </p>
        </div>
        <div className="flex gap-3">
          <button className="flex items-center gap-2 px-4 py-2 rounded-lg border border-border text-sm hover:bg-slate-50 text-slate-700 font-medium">
            <Filter size={16} />
            Filters
          </button>
          <button 
            onClick={fetchBonds}
            className="flex items-center gap-2 px-4 py-2 rounded-lg bg-accent text-white font-bold text-sm hover:bg-accent/90 shadow-sm transition-all active:scale-95"
          >
            <RefreshCw size={16} />
            Refresh
          </button>
        </div>
      </div>

      <div className="border border-border rounded-xl overflow-hidden bg-white shadow-sm">
        <table className="w-full text-left border-collapse">
          <thead>
            <tr className="border-b border-border bg-slate-50">
              <th className="p-4 text-[10px] font-bold text-slate-500 uppercase tracking-widest">Issuer Name</th>
              <th className="p-4 text-[10px] font-bold text-slate-500 uppercase tracking-widest">Coupon Rate %</th>
              <th className="p-4 text-[10px] font-bold text-slate-500 uppercase tracking-widest">Rating</th>
              <th className="p-4 text-[10px] font-bold text-slate-500 uppercase tracking-widest">Issuance Date</th>
              <th className="p-4 text-[10px] font-bold text-slate-500 uppercase tracking-widest">Maturity Date</th>
              <th className="p-4 text-[10px] font-bold text-slate-500 uppercase tracking-widest">ALM Fit</th>
              <th className="p-4 text-[10px] font-bold text-slate-500 uppercase tracking-widest">Verdict</th>
              <th className="p-4 text-[10px] font-bold text-slate-500 uppercase tracking-widest text-right">Action</th>
            </tr>
          </thead>
          <tbody>
            {bonds.map((bond, idx) => (
              <tr key={idx} className="border-b border-border/50 hover:bg-slate-50 transition-colors group">
                <td className="p-4">
                  <div className="font-bold text-sm text-slate-900">{bond.issuerName}</div>
                  <div className="text-[10px] text-text-dim font-mono mt-1">{bond.isin}</div>
                  <div className="mt-2">
                    <span className={`text-[9px] font-bold px-1.5 py-0.5 rounded ${
                      bond.marketLabel === 'PRIMARY MARKET' ? 'bg-blue-100 text-blue-700 border border-blue-200' :
                      bond.marketLabel === 'UPCOMING' ? 'bg-amber-100 text-amber-700 border border-amber-200' :
                      'bg-emerald-100 text-emerald-700 border border-emerald-200'
                    }`}>
                      {bond.marketLabel}
                    </span>
                  </div>
                </td>
                <td className="p-4 font-mono text-sm text-accent font-semibold">{bond.couponRate}%</td>
                <td className="p-4">
                  <div className="flex items-center gap-2">
                    <ShieldCheck size={14} className="text-accent" />
                    <span className="text-sm font-semibold text-slate-700">{bond.rating}</span>
                  </div>
                </td>
                <td className="p-4 text-sm text-slate-600">{bond.issuanceDate}</td>
                <td className="p-4 text-sm text-slate-600">{bond.maturityDate}</td>
                <td className="p-4">
                  <div className={`text-xs font-bold ${
                    bond.almFit === 'High' ? 'text-emerald-600' :
                    bond.almFit === 'Medium' ? 'text-amber-600' : 'text-rose-600'
                  }`}>
                    {bond.almFit}
                  </div>
                  <div className="text-[10px] text-slate-500 mt-1 max-w-[120px] leading-tight font-medium">
                    {bond.almExplanation}
                  </div>
                </td>
                <td className="p-4">
                  <span className={`px-2 py-1 rounded text-[10px] font-bold ${
                    bond.verdict === 'BUY' ? 'bg-emerald-600 text-white' :
                    bond.verdict === 'HOLD' ? 'bg-amber-500 text-white' : 'bg-rose-600 text-white'
                  }`}>
                    {bond.verdict}
                  </span>
                </td>
                <td className="p-4 text-right">
                  <button 
                    onClick={() => onSelectBond(bond)}
                    className="inline-flex items-center gap-1 text-xs font-bold text-accent hover:underline"
                  >
                    Deep Dive
                    <ChevronRight size={14} />
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
};

const DeepDiveAnalyser = ({ selectedBond, onBack }: { selectedBond: NCD | null, onBack: () => void }) => {
  const [memo, setMemo] = useState<InvestmentMemo | null>(null);
  const [loading, setLoading] = useState(false);
  const [exporting, setExporting] = useState<'pdf' | 'word' | null>(null);
  const memoRef = React.useRef<HTMLDivElement>(null);

  const validateMemo = (m: InvestmentMemo | null) => {
    if (!m) return false;
    // Core sections required for a professional memo
    const requiredSections = [
      'executiveSummary',
      'financialAnnexure',
      'recommendation'
    ];
    return requiredSections.every(section => {
      const val = (m as any)[section];
      if (Array.isArray(val)) return val.length > 0;
      return !!val;
    });
  };

  const handleExportPDF = async () => {
    if (!memo || !memoRef.current) return;
    
    if (!validateMemo(memo)) {
      alert('The Investment Memo is still being generated or is incomplete. Please wait for the financial data to load before exporting.');
      return;
    }

    const element = memoRef.current;
    
    setExporting('pdf');
    element.classList.add('pdf-export-mode');
    
    try {
      // Small delay to ensure styles are applied
      await new Promise(resolve => setTimeout(resolve, 500));
      
      const pdf = new jsPDF({
        orientation: 'portrait',
        unit: 'px',
        format: 'a4',
      });

      const pdfWidth = pdf.internal.pageSize.getWidth();
      const pageHeight = pdf.internal.pageSize.getHeight();
      const margin = 40; // Margin in px
      const contentWidth = pdfWidth - (margin * 2);
      
      let currentY = margin;

      // Get all sections and the header
      const sections = Array.from(element.querySelectorAll('.pdf-section'));
      
      for (let i = 0; i < sections.length; i++) {
        const section = sections[i] as HTMLElement;
        
        // Force page break if class is present
        if (section.classList.contains('pdf-page-break') && currentY > margin) {
          pdf.addPage();
          currentY = margin;
        }

        const canvas = await htmlToImage.toCanvas(section, {
          quality: 1,
          backgroundColor: '#ffffff',
          pixelRatio: 2,
        });

        const imgWidth = contentWidth;
        const imgHeight = (canvas.height * imgWidth) / canvas.width;

        // Check if section fits on current page
        if (currentY + imgHeight > pageHeight - margin) {
          // If it's a single section that's larger than a whole page, we'll have to split it or scale it
          // But usually sections are smaller. For now, we add a new page.
          if (currentY > margin) {
            pdf.addPage();
            currentY = margin;
          }
        }

        const imgData = canvas.toDataURL('image/png');
        pdf.addImage(imgData, 'PNG', margin, currentY, imgWidth, imgHeight);
        
        currentY += imgHeight + 20; // Add some spacing between sections
      }

      pdf.save(`${memo.issuerName}_Investment_Memo.pdf`);
    } catch (error) {
      console.error('PDF Export failed:', error);
      alert('Failed to export PDF. Please try again.');
    } finally {
      element.classList.remove('pdf-export-mode');
      setExporting(null);
    }
  };

  const handleExportWord = async () => {
    if (!memo) return;

    if (!validateMemo(memo)) {
      alert('The Investment Memo is still being generated or is incomplete. Please wait for the financial data to load before exporting.');
      return;
    }

    setExporting('word');
    const captureAsImage = async (id: string) => {
      const el = document.getElementById(id);
      if (!el) return null;
      try {
        // Ensure element is fully visible and rendered
        await new Promise(resolve => setTimeout(resolve, 200));
        
        const rect = el.getBoundingClientRect();
        const aspectRatio = rect.height / rect.width;
        
        const dataUrl = await htmlToImage.toPng(el, { 
          quality: 1, 
          backgroundColor: '#ffffff', 
          pixelRatio: 2,
          style: {
            borderRadius: '0',
            boxShadow: 'none',
            margin: '0',
            padding: '20px' // Add some padding for the image
          }
        });
        const blob = await (await fetch(dataUrl)).blob();
        const arrayBuffer = await blob.arrayBuffer();
        return { data: arrayBuffer, aspectRatio };
      } catch (e) {
        console.error(`Failed to capture ${id}`, e);
        return null;
      }
    };

    const sections = [
      { id: 'instrument-details', title: 'Instrument Details' },
      { id: 'management-table', title: 'Management' },
      { id: 'yield-table', title: 'Yield Analysis' },
      { id: 'financial-table', title: 'Financial Performance' },
      { id: 'trend-graph', title: 'Financial Trend Graph' },
      { id: 'swot-table', title: 'SWOT Analysis' },
      { id: 'annexure-table', title: 'Financial Annexure' }
    ];

    try {
      const capturedImages: Record<string, { data: ArrayBuffer, aspectRatio: number } | null> = {};
      for (const section of sections) {
        capturedImages[section.id] = await captureAsImage(section.id);
      }

      const doc = new Document({
        sections: [{
          properties: {
            page: {
              margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 }, // 1 inch = 1440 twips
            }
          },
          children: [
            new Paragraph({
              text: "Investment Memo",
              heading: HeadingLevel.HEADING_1,
              alignment: AlignmentType.CENTER,
            }),
            new Paragraph({
              text: memo.issuerName,
              heading: HeadingLevel.HEADING_2,
              alignment: AlignmentType.CENTER,
            }),
            new Paragraph({
              text: `Generated Date: ${memo.generatedDate}`,
              alignment: AlignmentType.CENTER,
              spacing: { after: 400 },
            }),

            new Paragraph({ text: "01. Verdict Header", heading: HeadingLevel.HEADING_1, spacing: { before: 400, after: 200 } }),
            new Paragraph({ 
              children: [new TextRun({ text: `Final Recommendation: ${memo.verdict}`, bold: true })] 
            }),
            ...(capturedImages['instrument-details'] ? [
              new Paragraph({
                children: [new ImageRun({ 
                  data: capturedImages['instrument-details']!.data, 
                  transformation: { width: 600, height: 600 * capturedImages['instrument-details']!.aspectRatio },
                  type: 'png'
                })],
                alignment: AlignmentType.CENTER,
                spacing: { before: 200, after: 200 }
              }),
              new Paragraph({ text: "Table: Instrument Details", alignment: AlignmentType.CENTER, spacing: { after: 200 } })
            ] : []),

            new Paragraph({ text: "02. Executive Summary", heading: HeadingLevel.HEADING_1, spacing: { before: 400, after: 200 } }),
            new Paragraph({ text: memo.executiveSummary, spacing: { after: 200 } }),

            new Paragraph({ text: "03. Company Overview", heading: HeadingLevel.HEADING_1, spacing: { before: 400, after: 200 } }),
            new Paragraph({ text: memo.companyOverview.description, spacing: { after: 200 } }),

            new Paragraph({ text: "04. Management", heading: HeadingLevel.HEADING_1, spacing: { before: 400, after: 200 } }),
            ...(capturedImages['management-table'] ? [
              new Paragraph({
                children: [new ImageRun({ 
                  data: capturedImages['management-table']!.data, 
                  transformation: { width: 600, height: 600 * capturedImages['management-table']!.aspectRatio },
                  type: 'png'
                })],
                alignment: AlignmentType.CENTER
              }),
              new Paragraph({ text: "Table: Management Stakeholders", alignment: AlignmentType.CENTER, spacing: { after: 200 } })
            ] : []),

            new Paragraph({ text: "05. ALM Suitability", heading: HeadingLevel.HEADING_1, spacing: { before: 400, after: 200 } }),
            new Paragraph({ 
              children: [new TextRun({ text: `Duration: ${memo.almSuitability.duration}`, bold: true })] 
            }),
            new Paragraph({ text: memo.almSuitability.explanation }),
            new Paragraph({ text: `Solvency Impact: ${memo.almSuitability.solvencyImpact}`, spacing: { after: 200 } }),

            new Paragraph({ text: "06. Instrument Overview", heading: HeadingLevel.HEADING_1, spacing: { before: 400, after: 200 } }),
            new Paragraph({ text: memo.instrumentOverview, spacing: { after: 200 } }),

            new Paragraph({ text: "07. Yield Analysis", heading: HeadingLevel.HEADING_1, spacing: { before: 400, after: 200 } }),
            ...(capturedImages['yield-table'] ? [
              new Paragraph({
                children: [new ImageRun({ 
                  data: capturedImages['yield-table']!.data, 
                  transformation: { width: 600, height: 600 * capturedImages['yield-table']!.aspectRatio },
                  type: 'png'
                })],
                alignment: AlignmentType.CENTER
              }),
              new Paragraph({ text: "Table: Yield & Spread Analysis", alignment: AlignmentType.CENTER, spacing: { after: 200 } })
            ] : []),

            new Paragraph({ text: "08. Business & Industry Overview", heading: HeadingLevel.HEADING_1, spacing: { before: 400, after: 200 } }),
            new Paragraph({ text: memo.businessIndustryOverview, spacing: { after: 200 } }),

            new Paragraph({ text: "09. Financial Performance", heading: HeadingLevel.HEADING_1, spacing: { before: 400, after: 200 } }),
            ...(capturedImages['financial-table'] ? [
              new Paragraph({
                children: [new ImageRun({ 
                  data: capturedImages['financial-table']!.data, 
                  transformation: { width: 600, height: 600 * capturedImages['financial-table']!.aspectRatio },
                  type: 'png'
                })],
                alignment: AlignmentType.CENTER
              }),
              new Paragraph({ text: "Table: Financial Performance Summary", alignment: AlignmentType.CENTER, spacing: { after: 200 } })
            ] : []),

            new Paragraph({ text: "10. Financial Trend Graph", heading: HeadingLevel.HEADING_1, spacing: { before: 400, after: 200 } }),
            ...(capturedImages['trend-graph'] ? [
              new Paragraph({
                children: [new ImageRun({ 
                  data: capturedImages['trend-graph']!.data, 
                  transformation: { width: 600, height: 600 * capturedImages['trend-graph']!.aspectRatio },
                  type: 'png'
                })],
                alignment: AlignmentType.CENTER
              }),
              new Paragraph({ text: "Graph: Financial Trends", alignment: AlignmentType.CENTER, spacing: { after: 200 } })
            ] : []),

            new Paragraph({ text: "11. Credit Profile", heading: HeadingLevel.HEADING_1, spacing: { before: 400, after: 200 } }),
            new Paragraph({ text: memo.creditProfile.discussion, spacing: { after: 200 } }),

            new Paragraph({ text: "12. SWOT Analysis", heading: HeadingLevel.HEADING_1, spacing: { before: 400, after: 200 } }),
            ...(capturedImages['swot-table'] ? [
              new Paragraph({
                children: [new ImageRun({ 
                  data: capturedImages['swot-table']!.data, 
                  transformation: { width: 600, height: 600 * capturedImages['swot-table']!.aspectRatio },
                  type: 'png'
                })],
                alignment: AlignmentType.CENTER
              }),
              new Paragraph({ text: "Table: SWOT Analysis", alignment: AlignmentType.CENTER, spacing: { after: 200 } })
            ] : []),

            new Paragraph({ text: "13. Recommendation", heading: HeadingLevel.HEADING_1, spacing: { before: 400, after: 200 } }),
            new Paragraph({ text: memo.recommendation, spacing: { after: 200 } }),

            new Paragraph({ text: "14. Financial Annexure", heading: HeadingLevel.HEADING_1, spacing: { before: 400, after: 200 } }),
            ...(capturedImages['annexure-table'] ? [
              new Paragraph({
                children: [new ImageRun({ 
                  data: capturedImages['annexure-table']!.data, 
                  transformation: { width: 600, height: 600 * capturedImages['annexure-table']!.aspectRatio },
                  type: 'png'
                })],
                alignment: AlignmentType.CENTER
              }),
              new Paragraph({ text: "Table: Detailed Financial Annexure", alignment: AlignmentType.CENTER, spacing: { after: 200 } })
            ] : []),
          ]
        }]
      });

      const blob = await Packer.toBlob(doc);
      saveAs(blob, `${memo.issuerName}_Investment_Memo.docx`);
    } catch (error) {
      console.error('Word Export failed:', error);
      alert('Failed to export Word document. Please try again.');
    } finally {
      setExporting(null);
    }
  };

  useEffect(() => {
    if (selectedBond) {
      const fetchMemo = async () => {
        setLoading(true);
        const data = await geminiService.generateMemo(selectedBond);
        setMemo(data);
        setLoading(false);
      };
      fetchMemo();
    }
  }, [selectedBond]);

  if (!selectedBond) {
    return (
      <div className="flex-1 flex flex-col items-center justify-center p-8 text-center">
        <div className="w-16 h-16 bg-white/5 rounded-full flex items-center justify-center mb-6">
          <Search className="text-text-dim" size={32} />
        </div>
        <h3 className="text-xl font-bold mb-2">No Issuer Selected</h3>
        <p className="text-sm text-text-dim max-w-xs">
          Select a bond from the Market Explorer to generate a full institutional investment memo.
        </p>
      </div>
    );
  }

  if (loading) {
    return (
      <div className="flex-1 flex items-center justify-center">
        <div className="flex flex-col items-center gap-4">
          <Loader2 className="w-8 h-8 animate-spin text-accent" />
          <p className="text-sm text-text-dim animate-pulse">Generating Institutional Memo for {selectedBond.issuerName}...</p>
        </div>
      </div>
    );
  }

  if (!memo) return null;

  return (
    <div className="flex-1 p-8 overflow-auto bg-slate-50">
      <div className="max-w-4xl mx-auto">
        <div className="flex justify-between items-center mb-8">
          <button 
            onClick={onBack}
            className="flex items-center gap-2 text-sm text-slate-500 hover:text-accent transition-colors font-medium"
          >
            <ChevronRight className="rotate-180" size={16} />
            Back to Explorer
          </button>

          <div className="flex gap-4">
            <button 
              onClick={handleExportPDF}
              disabled={exporting !== null}
              className="flex items-center gap-2 px-4 py-2 rounded-lg bg-white border border-border text-xs font-bold text-slate-700 hover:bg-slate-50 transition-all shadow-sm disabled:opacity-50"
            >
              {exporting === 'pdf' ? <Loader2 size={14} className="animate-spin" /> : <FileDown size={14} className="text-rose-600" />}
              {exporting === 'pdf' ? 'Preparing PDF...' : 'Download as PDF'}
            </button>
            <button 
              onClick={handleExportWord}
              disabled={exporting !== null}
              className="flex items-center gap-2 px-4 py-2 rounded-lg bg-white border border-border text-xs font-bold text-slate-700 hover:bg-slate-50 transition-all shadow-sm disabled:opacity-50"
            >
              {exporting === 'word' ? <Loader2 size={14} className="animate-spin" /> : <FileIcon size={14} className="text-blue-600" />}
              {exporting === 'word' ? 'Preparing Word...' : 'Download as Word Document'}
            </button>
          </div>
        </div>

        <div ref={memoRef} className="bg-white p-12 rounded-2xl border border-border shadow-sm">
          <header className="mb-12 border-b border-border pb-8 pdf-section">
            <div className="flex justify-between items-start mb-6">
              <div>
                <h1 className="text-4xl font-bold tracking-tight mb-2 text-slate-900">Investment Memo</h1>
                <h2 className="text-2xl font-bold text-slate-700 mb-1">{memo.issuerName}</h2>
                <p className="text-sm text-text-dim uppercase tracking-widest font-semibold">{memo.generatedDate}</p>
              </div>
              <div className="text-right">
                <div className="text-[10px] text-text-dim uppercase tracking-widest mb-1 font-bold">Final Recommendation</div>
                <div className={`text-2xl font-black ${
                  memo.verdict === 'BUY' ? 'text-emerald-600' :
                  memo.verdict === 'HOLD' ? 'text-amber-600' : 'text-rose-600'
                }`}>
                  {memo.verdict}
                </div>
              </div>
            </div>

            <div id="instrument-details" className="grid grid-cols-6 gap-4 p-4 bg-slate-50 rounded-xl border border-border">
              {Object.entries(memo.instrumentDetails || {}).map(([key, val]) => (
                <div key={key}>
                  <div className="text-[9px] text-slate-400 uppercase tracking-widest mb-1 font-bold">{key.replace(/([A-Z])/g, ' $1')}</div>
                  <div className="text-xs font-bold text-slate-800">{val}</div>
                </div>
              ))}
            </div>
          </header>

          <div className="space-y-12">
            {/* 02. EXECUTIVE SUMMARY */}
            <section className="pdf-avoid-break pdf-section">
              <h3 className="text-xs font-bold text-accent uppercase tracking-[0.2em] mb-4">02. Executive Summary</h3>
              <div className="text-sm leading-relaxed text-slate-800">
                <ReactMarkdown>{safeString(memo.executiveSummary)}</ReactMarkdown>
              </div>
            </section>

            {/* 03. COMPANY OVERVIEW */}
            <section className="pdf-avoid-break pdf-section">
              <h3 className="text-xs font-bold text-accent uppercase tracking-[0.2em] mb-4">03. Company Overview</h3>
              <div className="mb-6">
                <p className="text-sm text-slate-800 leading-relaxed mb-6">{memo.companyOverview?.description}</p>
                <div className="grid grid-cols-4 gap-4">
                  {(memo.companyOverview?.metrics || []).map((m, i) => (
                    <div key={i} className="p-3 rounded-lg bg-slate-50 border border-slate-100">
                      <div className="text-[9px] text-slate-500 uppercase tracking-widest mb-1 font-bold">{m.label}</div>
                      <div className="text-xs font-bold text-slate-900">{m.value}</div>
                    </div>
                  ))}
                </div>
              </div>
            </section>

            {/* 04. LIST OF STAKEHOLDERS / MANAGEMENT */}
            <section id="management-table" className="pdf-avoid-break pdf-section">
              <h3 className="text-xs font-bold text-accent uppercase tracking-[0.2em] mb-4">04. List of Stakeholders / Management</h3>
              <div className="border border-border rounded-xl overflow-hidden">
                <table className="w-full text-left border-collapse text-xs">
                  <thead>
                    <tr className="border-b border-border bg-slate-100">
                      <th className="p-3 font-bold text-slate-600 uppercase tracking-widest">Name</th>
                      <th className="p-3 font-bold text-slate-600 uppercase tracking-widest">Designation</th>
                    </tr>
                  </thead>
                  <tbody>
                    {(memo.management || []).map((m, i) => (
                      <tr key={i} className="border-b border-border/50">
                        <td className="p-3 font-bold text-slate-900">{m.name}</td>
                        <td className="p-3 text-slate-700">{m.designation}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </section>

            {/* 05. ALM SUITABILITY ANALYSIS */}
            <section className="pdf-avoid-break pdf-section">
              <h3 className="text-xs font-bold text-accent uppercase tracking-[0.2em] mb-4">05. ALM Suitability Analysis</h3>
              <div className="p-6 rounded-xl border border-border bg-slate-50">
                <div className="flex items-center gap-3 mb-6">
                  <span className={`px-3 py-1 rounded-full text-[10px] font-bold uppercase tracking-wider border ${
                    memo.almSuitability?.duration === 'Long' ? 'bg-emerald-50 text-emerald-700 border-emerald-200' :
                    memo.almSuitability?.duration === 'Medium' ? 'bg-blue-50 text-blue-700 border-blue-200' :
                    'bg-slate-50 text-slate-700 border-slate-200'
                  }`}>
                    {memo.almSuitability?.duration} Duration Bucket
                  </span>
                </div>
                <div className="grid grid-cols-2 gap-8">
                  <div>
                    <h4 className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-2">Duration & Liability Matching</h4>
                    <p className="text-sm text-slate-800 leading-relaxed">{memo.almSuitability?.explanation}</p>
                  </div>
                  <div>
                    <h4 className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-2">Solvency Impact</h4>
                    <p className="text-sm text-slate-800 leading-relaxed">{memo.almSuitability?.solvencyImpact}</p>
                  </div>
                </div>
              </div>
            </section>

            {/* 06. INSTRUMENT OVERVIEW */}
            <section className="pdf-avoid-break pdf-section">
              <h3 className="text-xs font-bold text-accent uppercase tracking-[0.2em] mb-4">06. Instrument Overview</h3>
              <p className="text-sm text-slate-800 leading-relaxed">{safeString(memo.instrumentOverview)}</p>
            </section>

            {/* 07. YIELD & SPREAD ANALYSIS */}
            <section id="yield-table" className="pdf-avoid-break pdf-section">
              <h3 className="text-xs font-bold text-accent uppercase tracking-[0.2em] mb-4">07. Yield & Spread Analysis</h3>
              <div className="grid grid-cols-3 gap-6 mb-6">
                <div className="p-4 rounded-xl border border-border bg-slate-50">
                  <div className="text-[9px] text-slate-500 uppercase tracking-widest mb-2">Selected NCD Yield</div>
                  <div className="text-xl font-mono font-bold text-accent">{memo.yieldAnalysis?.ncdYield || 'N/A'}</div>
                </div>
                <div className="p-4 rounded-xl border border-border bg-slate-50">
                  <div className="text-[9px] text-slate-500 uppercase tracking-widest mb-2">10Y G-Sec Yield</div>
                  <div className="text-xl font-mono font-bold text-blue-600">{memo.yieldAnalysis?.gSecYield || 'N/A'}</div>
                </div>
                <div className="p-4 rounded-xl border border-border bg-slate-50">
                  <div className="text-[9px] text-slate-500 uppercase tracking-widest mb-2">Large Bank FD Yield</div>
                  <div className="text-xl font-mono font-bold text-amber-600">{memo.yieldAnalysis?.fdYield || 'N/A'}</div>
                </div>
              </div>
              <div className="grid grid-cols-2 gap-4">
                <div className="p-4 rounded-lg bg-slate-50 border border-slate-100">
                  <div className="text-[9px] text-slate-500 uppercase tracking-widest mb-1 font-bold">Credit Spread vs G-Sec</div>
                  <div className="text-sm font-bold text-slate-900">{memo.yieldAnalysis?.creditSpread}</div>
                </div>
                <div className="p-4 rounded-lg bg-slate-50 border border-slate-100">
                  <div className="text-[9px] text-slate-500 uppercase tracking-widest mb-1 font-bold">Yield Premium vs FD</div>
                  <div className="text-sm font-bold text-slate-900">{memo.yieldAnalysis?.yieldPremium}</div>
                </div>
              </div>
            </section>

            {/* 08. BUSINESS & INDUSTRY OVERVIEW */}
            <section className="pdf-avoid-break pdf-section">
              <h3 className="text-xs font-bold text-accent uppercase tracking-[0.2em] mb-4">08. Business & Industry Overview</h3>
              <p className="text-sm text-slate-800 leading-relaxed">{safeString(memo.businessIndustryOverview)}</p>
            </section>

            {/* 09. FINANCIAL PERFORMANCE */}
            <section id="financial-table" className="pdf-avoid-break pdf-section">
              <div className="flex justify-between items-end mb-4">
                <h3 className="text-xs font-bold text-accent uppercase tracking-[0.2em]">09. Financial Performance</h3>
                <span className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Values in ₹ Crore</span>
              </div>
              <div className="border border-border rounded-xl overflow-hidden">
                <table className="w-full text-left border-collapse text-xs">
                  <thead>
                    <tr className="border-b border-border bg-slate-100">
                      <th className="p-3 font-bold text-slate-600 uppercase tracking-widest">Year</th>
                      <th className="p-3 font-bold text-slate-600 uppercase tracking-widest">Revenue / Operating Income</th>
                      <th className="p-3 font-bold text-slate-600 uppercase tracking-widest">Net Profit</th>
                    </tr>
                  </thead>
                  <tbody>
                    {(memo.financialPerformanceTable || []).map((row, i) => (
                      <tr key={i} className="border-b border-border/50">
                        <td className="p-3 font-bold text-slate-900">{row.year}</td>
                        <td className="p-3 font-mono text-slate-700">{row.revenue}</td>
                        <td className="p-3 font-mono text-slate-700">{row.netProfit}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </section>

            {/* 10. FINANCIAL TREND GRAPH */}
            <section id="trend-graph" className="pdf-avoid-break pdf-section">
              <h3 className="text-xs font-bold text-accent uppercase tracking-[0.2em] mb-4">10. Financial Trend Graph</h3>
              <div className="p-8 rounded-xl border border-border bg-slate-50">
                <div className="space-y-8">
                  <div>
                    <h4 className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-4">Revenue Trend</h4>
                    <div className="space-y-3">
                      {(memo.financialTrendData || []).map((d, i) => (
                        <div key={i} className="flex items-center gap-4">
                          <div className="w-12 text-[10px] font-bold text-slate-500">{d.year}</div>
                          <div className="flex-1 h-3 bg-slate-200 rounded-full overflow-hidden">
                            <div 
                              className="h-full bg-accent transition-all duration-1000" 
                              style={{ width: `${Math.min(100, (d.revenueValue / Math.max(...memo.financialTrendData.map(x => x.revenueValue))) * 100)}%` }}
                            />
                          </div>
                          <div className="w-24 text-[10px] font-mono text-right text-slate-600">{d.revenueLabel}</div>
                        </div>
                      ))}
                    </div>
                  </div>
                  <div>
                    <h4 className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-4">Profit Trend</h4>
                    <div className="space-y-3">
                      {(memo.financialTrendData || []).map((d, i) => (
                        <div key={i} className="flex items-center gap-4">
                          <div className="w-12 text-[10px] font-bold text-slate-500">{d.year}</div>
                          <div className="flex-1 h-3 bg-slate-200 rounded-full overflow-hidden">
                            <div 
                              className="h-full bg-emerald-500 transition-all duration-1000" 
                              style={{ width: `${Math.min(100, (d.netProfitValue / Math.max(...memo.financialTrendData.map(x => x.netProfitValue))) * 100)}%` }}
                            />
                          </div>
                          <div className="w-24 text-[10px] font-mono text-right text-slate-600">{d.netProfitLabel}</div>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>
              </div>
            </section>

            {/* 11. CREDIT PROFILE */}
            <section className="pdf-avoid-break pdf-section">
              <h3 className="text-xs font-bold text-accent uppercase tracking-[0.2em] mb-4">11. Credit Profile</h3>
              <div className="mb-6">
                <p className="text-sm text-slate-800 leading-relaxed mb-6">{memo.creditProfile?.discussion}</p>
                <div className="grid grid-cols-3 gap-4">
                  {(memo.creditProfile?.metrics || []).map((m, i) => (
                    <div key={i} className="p-4 rounded-lg bg-slate-50 border border-slate-100">
                      <div className="text-[9px] text-slate-500 uppercase tracking-widest mb-1 font-bold">{m.label}</div>
                      <div className="text-sm font-bold text-slate-900">{m.value}</div>
                    </div>
                  ))}
                </div>
              </div>
            </section>

            {/* 12. SWOT ANALYSIS */}
            <section id="swot-table" className="pdf-avoid-break pdf-section">
              <h3 className="text-xs font-bold text-accent uppercase tracking-[0.2em] mb-4">12. SWOT Analysis</h3>
              <div className="grid grid-cols-2 gap-4">
                <div className="p-6 rounded-xl border border-emerald-500/20 bg-emerald-50">
                  <div className="text-[10px] font-bold text-emerald-700 uppercase tracking-widest mb-4">Strengths</div>
                  <ul className="space-y-2">
                    {(memo.swot?.strengths || []).map((s, i) => (
                      <li key={i} className="text-xs text-slate-800 flex gap-2">
                        <span className="text-emerald-600">•</span> {s}
                      </li>
                    ))}
                  </ul>
                </div>
                <div className="p-6 rounded-xl border border-rose-500/20 bg-rose-50">
                  <div className="text-[10px] font-bold text-rose-700 uppercase tracking-widest mb-4">Weaknesses</div>
                  <ul className="space-y-2">
                    {(memo.swot?.weaknesses || []).map((s, i) => (
                      <li key={i} className="text-xs text-slate-800 flex gap-2">
                        <span className="text-rose-600">•</span> {s}
                      </li>
                    ))}
                  </ul>
                </div>
                <div className="p-6 rounded-xl border border-blue-500/20 bg-blue-50">
                  <div className="text-[10px] font-bold text-blue-700 uppercase tracking-widest mb-4">Opportunities</div>
                  <ul className="space-y-2">
                    {(memo.swot?.opportunities || []).map((s, i) => (
                      <li key={i} className="text-xs text-slate-800 flex gap-2">
                        <span className="text-blue-600">•</span> {s}
                      </li>
                    ))}
                  </ul>
                </div>
                <div className="p-6 rounded-xl border border-amber-500/20 bg-amber-50">
                  <div className="text-[10px] font-bold text-amber-700 uppercase tracking-widest mb-4">Threats</div>
                  <ul className="space-y-2">
                    {(memo.swot?.threats || []).map((s, i) => (
                      <li key={i} className="text-xs text-slate-800 flex gap-2">
                        <span className="text-amber-600">•</span> {s}
                      </li>
                    ))}
                  </ul>
                </div>
              </div>
            </section>

            {/* 13. INVESTMENT TEAM RECOMMENDATION */}
            <section className="pdf-page-break pdf-avoid-break pdf-section">
              <h3 className="text-xs font-bold text-accent uppercase tracking-[0.2em] mb-4">13. Investment Team Recommendation</h3>
              <div className="p-8 rounded-2xl bg-accent text-white shadow-lg">
                <h4 className="text-lg font-black uppercase tracking-tight mb-4">NBHI Portfolio Recommendation</h4>
                <p className="text-sm font-medium leading-relaxed mb-6">{safeString(memo.recommendation)}</p>
                <div className="w-24 h-8 border-b-2 border-white/30"></div>
              </div>
            </section>

            {/* 14. FINANCIAL ANNEXURE */}
            <section id="annexure-table" className="pdf-page-break pdf-avoid-break pdf-section">
              <div className="flex justify-between items-end mb-6">
                <h3 className="text-xs font-bold text-accent uppercase tracking-[0.2em]">14. Financial Annexure</h3>
                <span className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Values in ₹ Cr (EPS in ₹)</span>
              </div>
              <div className="border border-border rounded-xl overflow-hidden bg-white">
                <table className="w-full text-left border-collapse text-xs">
                  <thead>
                    <tr className="border-b border-border bg-slate-100">
                      <th className="p-3 font-bold text-slate-600 uppercase tracking-widest">Metric</th>
                      <th className="p-3 font-bold text-slate-600 uppercase tracking-widest">FY21</th>
                      <th className="p-3 font-bold text-slate-600 uppercase tracking-widest">FY22</th>
                      <th className="p-3 font-bold text-slate-600 uppercase tracking-widest">FY23</th>
                      <th className="p-3 font-bold text-slate-600 uppercase tracking-widest">FY24</th>
                      <th className="p-3 font-bold text-slate-600 uppercase tracking-widest">FY25</th>
                    </tr>
                  </thead>
                  <tbody>
                    {memo.financialAnnexure && memo.financialAnnexure.length > 0 ? (
                      memo.financialAnnexure.map((row, i) => (
                        <tr key={i} className="border-b border-border/50">
                          <td className="p-3 font-medium text-slate-900">{row.metric}</td>
                          <td className="p-3 font-mono text-slate-700">{row.fy21 || 'Not Available'}</td>
                          <td className="p-3 font-mono text-slate-700">{row.fy22 || 'Not Available'}</td>
                          <td className="p-3 font-mono text-slate-700">{row.fy23 || 'Not Available'}</td>
                          <td className="p-3 font-mono text-slate-700">{row.fy24 || 'Not Available'}</td>
                          <td className="p-3 font-mono text-slate-700">{row.fy25 || 'Not Available'}</td>
                        </tr>
                      ))
                    ) : (
                      <tr>
                        <td colSpan={6} className="p-8 text-center text-slate-400 italic">
                          Financial data currently being retrieved...
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </section>

            {/* 15. DATA SOURCES */}
            <section className="pdf-avoid-break pdf-section">
              <h3 className="text-xs font-bold text-accent uppercase tracking-[0.2em] mb-4">15. Data Sources</h3>
              <div className="p-6 rounded-xl border border-border bg-slate-50">
                <ul className="grid grid-cols-2 gap-4">
                  {(memo.sources || []).map((source, i) => (
                    <li key={i} className="flex items-center gap-2 text-[10px] text-slate-600 font-medium">
                      <div className="w-1 h-1 rounded-full bg-accent" />
                      {source}
                    </li>
                  ))}
                </ul>
                <p className="mt-6 text-[9px] text-slate-400 italic">
                  * All data is sourced from publicly available financial disclosures, exchange filings, and reliable market terminals.
                </p>
              </div>
            </section>
          </div>
        </div>
      </div>
    </div>
  );
};

const ChatAssistant = () => {
  const [messages, setMessages] = useState<{ role: 'user' | 'model'; text: string }[]>([
    { role: 'model', text: 'Welcome to NBHI Investment Assistant. How can I help you with NCD research today?' }
  ]);
  const [input, setInput] = useState('');
  const [loading, setLoading] = useState(false);

  const handleSend = async () => {
    if (!input.trim() || loading) return;

    const userMsg = input;
    setInput('');
    setMessages(prev => [...prev, { role: 'user', text: userMsg }]);
    setLoading(true);

    try {
      const history = messages.map(m => ({
        role: m.role,
        parts: [{ text: m.text }]
      }));
      
      const response = await geminiService.chat(userMsg, history);
      setMessages(prev => [...prev, { role: 'model', text: response.text || 'I encountered an error processing your request.' }]);
    } catch (error) {
      setMessages(prev => [...prev, { role: 'model', text: 'Error: Failed to connect to research engine.' }]);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="flex-1 flex flex-col h-full bg-white">
      <div className="p-6 border-b border-border flex justify-between items-center bg-slate-50">
        <div>
          <h2 className="text-lg font-bold tracking-tight text-accent">Research Assistant</h2>
          <p className="text-xs text-text-dim font-medium">Powered by Gemini 3 Flash • Google Search Grounding</p>
        </div>
        <div className="flex items-center gap-2">
          <div className="w-2 h-2 rounded-full bg-accent animate-pulse"></div>
          <span className="text-[10px] font-bold uppercase tracking-widest text-accent">Active</span>
        </div>
      </div>

      <div className="flex-1 overflow-auto p-6 space-y-6 bg-white">
        {messages.map((msg, i) => (
          <motion.div 
            initial={{ opacity: 0, y: 10 }}
            animate={{ opacity: 1, y: 0 }}
            key={i} 
            className={`flex gap-4 ${msg.role === 'user' ? 'flex-row-reverse' : ''}`}
          >
            <div className={`w-8 h-8 rounded-lg flex items-center justify-center shrink-0 ${
              msg.role === 'user' ? 'bg-accent text-white shadow-md' : 'bg-slate-100 text-accent border border-border'
            }`}>
              {msg.role === 'user' ? <User size={16} /> : <Bot size={16} />}
            </div>
            <div className={`max-w-[80%] p-4 rounded-2xl text-sm leading-relaxed shadow-sm ${
              msg.role === 'user' ? 'bg-accent/5 text-slate-800 border border-accent/20' : 'bg-slate-50 text-slate-700 border border-border'
            }`}>
              <ReactMarkdown>{msg.text}</ReactMarkdown>
            </div>
          </motion.div>
        ))}
        {loading && (
          <div className="flex gap-4">
            <div className="w-8 h-8 rounded-lg bg-slate-100 text-accent border border-border flex items-center justify-center">
              <Bot size={16} />
            </div>
            <div className="bg-slate-50 p-4 rounded-2xl border border-border shadow-sm">
              <Loader2 className="w-4 h-4 animate-spin text-accent" />
            </div>
          </div>
        )}
      </div>

      <div className="p-6 border-t border-border bg-slate-50">
        <div className="max-w-4xl mx-auto relative">
          <input 
            type="text"
            value={input}
            onChange={(e) => setInput(e.target.value)}
            onKeyDown={(e) => e.key === 'Enter' && handleSend()}
            placeholder="Ask about specific NCDs, yields, or ALM suitability..."
            className="w-full bg-white border border-border rounded-xl py-4 pl-6 pr-14 text-sm focus:outline-none focus:border-accent transition-colors shadow-sm"
          />
          <button 
            onClick={handleSend}
            disabled={loading}
            className="absolute right-2 top-2 bottom-2 w-10 bg-accent text-white rounded-lg flex items-center justify-center hover:bg-accent/90 transition-colors disabled:opacity-50 shadow-sm"
          >
            <Send size={18} />
          </button>
        </div>
        <p className="text-center text-[10px] text-text-dim mt-4 uppercase tracking-widest font-bold">
          Institutional Research Mode • Data sourced from Public Disclosures
        </p>
      </div>
    </div>
  );
};

// --- Main App ---

export default function App() {
  const [activeTab, setActiveTab] = useState('explorer');
  const [selectedBond, setSelectedBond] = useState<NCD | null>(null);

  const handleSelectBond = (bond: NCD) => {
    setSelectedBond(bond);
    setActiveTab('analyser');
  };

  return (
    <div className="flex h-screen bg-bg text-white overflow-hidden">
      <Sidebar activeTab={activeTab} setActiveTab={setActiveTab} />
      
      <main className="flex-1 flex flex-col relative overflow-hidden">
        <AnimatePresence mode="wait">
          {activeTab === 'explorer' && (
            <motion.div 
              key="explorer"
              initial={{ opacity: 0, x: 20 }}
              animate={{ opacity: 1, x: 0 }}
              exit={{ opacity: 0, x: -20 }}
              className="flex-1 flex flex-col overflow-hidden"
            >
              <MarketExplorer onSelectBond={handleSelectBond} />
            </motion.div>
          )}
          {activeTab === 'analyser' && (
            <motion.div 
              key="analyser"
              initial={{ opacity: 0, x: 20 }}
              animate={{ opacity: 1, x: 0 }}
              exit={{ opacity: 0, x: -20 }}
              className="flex-1 flex flex-col overflow-hidden"
            >
              <DeepDiveAnalyser 
                selectedBond={selectedBond} 
                onBack={() => setActiveTab('explorer')} 
              />
            </motion.div>
          )}
          {activeTab === 'assistant' && (
            <motion.div 
              key="assistant"
              initial={{ opacity: 0, x: 20 }}
              animate={{ opacity: 1, x: 0 }}
              exit={{ opacity: 0, x: -20 }}
              className="flex-1 flex flex-col overflow-hidden"
            >
              <ChatAssistant />
            </motion.div>
          )}
        </AnimatePresence>
      </main>
    </div>
  );
}
