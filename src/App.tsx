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
import * as XLSX from 'xlsx';

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

const validateMemo = (m: InvestmentMemo | null) => {
  if (!m) return false;
  // Core sections required for a professional memo
  const requiredSections = [
    'instrumentOverview',
    'issuerOverview',
    'industryOverview',
    'creditRatingAnalysis',
    'financialPerformanceAnalysis',
    'balanceSheetStrength',
    'assetQualityCreditRisk',
    'yieldRelativeValueAnalysis',
    'liquidityMarketability',
    'riskAnalysis',
    'swotAnalysis',
    'almSuitability',
    'financialAnnexure',
    'finalInvestmentRecommendation'
  ];
  
  const hasBasicSections = requiredSections.every(section => {
    const val = (m as any)[section];
    if (Array.isArray(val)) return val.length > 0;
    if (typeof val === 'object' && val !== null) {
      // Specific deep checks
      if (section === 'issuerOverview' && (!val.management || val.management.length === 0)) return false;
      if (section === 'financialPerformanceAnalysis' && (!val.trends || val.trends.length === 0)) return false;
      return Object.keys(val).length > 0;
    }
    return !!val;
  });

  if (!hasBasicSections) return false;

  // Deep check for Financial Annexure metrics
  const requiredMetrics = ["Total Income", "Interest Income", "Operating Expenses", "PBT", "Net Profit", "EPS"];
  const hasAllMetrics = requiredMetrics.every(metric => {
    const row = m.financialAnnexure.find(r => r.metric.toLowerCase().includes(metric.toLowerCase()));
    if (!row) return false;
    // Check if at least 3 years have data (allowing for some missing historical data if absolutely necessary, but prompt is strict)
    const years = ['fy21', 'fy22', 'fy23', 'fy24', 'fy25'];
    const dataPoints = years.filter(y => {
      const val = (row as any)[y];
      return val && val !== 'N/A' && val !== '-' && val !== 'Not Available';
    });
    return dataPoints.length >= 3;
  });

  return hasAllMetrics;
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
    
    // Sort bonds: UPCOMING and PRIMARY MARKET first, then SECONDARY MARKET
    const sortedBonds = [...data].sort((a, b) => {
      const order: Record<string, number> = {
        'UPCOMING': 0,
        'PRIMARY MARKET': 1,
        'SECONDARY MARKET': 2
      };
      const orderA = order[a.marketLabel] ?? 3;
      const orderB = order[b.marketLabel] ?? 3;
      return orderA - orderB;
    });
    
    setBonds(sortedBonds);
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
              <th className="p-4 text-[10px] font-bold text-slate-500 uppercase tracking-widest">Bond Type</th>
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
                <td className="p-4 text-sm text-slate-600">{bond.bondType}</td>
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

const DeepDiveAnalyser = ({ selectedBond, onBack, onSelectBond }: { selectedBond: NCD | null, onBack: () => void, onSelectBond: (b: NCD | null) => void }) => {
  const [memo, setMemo] = useState<InvestmentMemo | null>(null);
  const [loading, setLoading] = useState(false);
  const [searching, setSearching] = useState(false);
  const [searchQuery, setSearchQuery] = useState('');
  const [exporting, setExporting] = useState<'pdf' | 'word' | null>(null);
  const memoRef = React.useRef<HTMLDivElement>(null);

  const getTrendFromAnnexure = (metricName: string) => {
    if (!memo?.financialAnnexure) return null;
    const row = memo.financialAnnexure.find(r => 
      r.metric.toLowerCase().replace(/[^a-z]/g, '').includes(metricName.toLowerCase().replace(/[^a-z]/g, ''))
    );
    if (!row) return null;
    
    return {
      metric: row.metric,
      values: [
        { year: 'FY21', value: row.fy21 },
        { year: 'FY22', value: row.fy22 },
        { year: 'FY23', value: row.fy23 },
        { year: 'FY24', value: row.fy24 },
        { year: 'FY25', value: row.fy25 }
      ].filter(v => v.value && v.value !== 'N/A' && v.value !== '-' && v.value !== 'Not Available')
    };
  };

  const totalIncomeTrend = getTrendFromAnnexure('Total Income') || 
    memo?.financialPerformanceAnalysis?.trends?.find(t => t.metric.toLowerCase().includes('total income'));
  
  const netProfitTrend = getTrendFromAnnexure('Net Profit') || 
    memo?.financialPerformanceAnalysis?.trends?.find(t => t.metric.toLowerCase().includes('net profit'));

  const otherTrends = memo?.financialPerformanceAnalysis?.trends?.filter(t => 
    !t.metric.toLowerCase().includes('total income') && 
    !t.metric.toLowerCase().includes('net profit')
  ) || [];

  const allTrends = [
    ...(totalIncomeTrend ? [totalIncomeTrend] : []),
    ...(netProfitTrend ? [netProfitTrend] : []),
    ...otherTrends
  ];

  const handleSearch = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!searchQuery.trim()) return;

    setSearching(true);
    try {
      const bond = await geminiService.searchBond(searchQuery);
      if (bond) {
        onSelectBond(bond);
      } else {
        alert('Could not find bond details for the given query. Please try a valid ISIN or full company name.');
      }
    } catch (error) {
      console.error('Search error:', error);
      alert('An error occurred while searching for the bond.');
    } finally {
      setSearching(false);
    }
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
      { id: 'instrument-overview', title: '01. Instrument Overview' },
      { id: 'issuer-overview', title: '02. Issuer Overview' },
      { id: 'industry-overview', title: '03. Industry Overview' },
      { id: 'credit-rating-analysis', title: '04. Credit Rating Analysis' },
      { id: 'financial-performance-analysis', title: '05. Financial Performance Analysis' },
      { id: 'balance-sheet-strength', title: '06. Balance Sheet Strength' },
      { id: 'asset-quality-credit-risk', title: '07. Asset Quality and Credit Risk' },
      { id: 'yield-relative-value-analysis', title: '08. Yield and Relative Value Analysis' },
      { id: 'liquidity-marketability', title: '09. Liquidity and Marketability' },
      { id: 'risk-analysis', title: '10. Risk Analysis' },
      { id: 'swot-analysis', title: '11. SWOT Analysis' },
      { id: 'alm-suitability', title: '12. ALM Suitability' },
      { id: 'annexure-table', title: '13. Financial Annexure' },
      { id: 'final-recommendation', title: '14. Final Investment Recommendation' }
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

            new Paragraph({ text: "01. Instrument Overview", heading: HeadingLevel.HEADING_1, spacing: { before: 400, after: 200 } }),
            ...(capturedImages['instrument-overview'] ? [
              new Paragraph({
                children: [new ImageRun({ 
                  data: capturedImages['instrument-overview']!.data, 
                  transformation: { width: 600, height: 600 * capturedImages['instrument-overview']!.aspectRatio },
                  type: 'png'
                })],
                alignment: AlignmentType.CENTER,
                spacing: { before: 200, after: 200 }
              })
            ] : []),

            new Paragraph({ text: "02. Issuer Overview", heading: HeadingLevel.HEADING_1, spacing: { before: 400, after: 200 } }),
            new Paragraph({ text: memo.issuerOverview.description }),
            new Paragraph({ text: `Ownership Structure: ${memo.issuerOverview.ownershipStructure}` }),
            new Paragraph({ text: `Status: ${memo.issuerOverview.status}` }),
            new Paragraph({ text: `Sector Positioning: ${memo.issuerOverview.sectorPositioning}` }),
            new Paragraph({ text: `Major Business Lines: ${memo.issuerOverview.majorBusinessLines.join(', ')}` }),
            ...(memo.issuerOverview.loanBookComposition ? [new Paragraph({ text: `Loan Book Composition: ${memo.issuerOverview.loanBookComposition}` })] : []),
            new Paragraph({ text: `Strategic Role: ${memo.issuerOverview.strategicRole}` }),
            ...(memo.issuerOverview.management && memo.issuerOverview.management.length > 0 ? [
              new Paragraph({ 
                children: [new TextRun({ text: "Key Management:", bold: true })], 
                spacing: { before: 200 } 
              }),
              ...memo.issuerOverview.management.map(p => new Paragraph({ text: `• ${p.name} - ${p.designation}`, bullet: { level: 0 } }))
            ] : []),
            new Paragraph({ text: "", spacing: { after: 200 } }),

            new Paragraph({ text: "03. Industry Overview", heading: HeadingLevel.HEADING_1, spacing: { before: 400, after: 200 } }),
            new Paragraph({ text: memo.industryOverview.content }),
            new Paragraph({ text: `Growth Drivers: ${memo.industryOverview.growthDrivers.join(', ')}` }),
            new Paragraph({ text: `Regulatory Framework: ${memo.industryOverview.regulatoryFramework}` }),
            new Paragraph({ text: `Risks: ${memo.industryOverview.risks.join(', ')}`, spacing: { after: 200 } }),

            new Paragraph({ text: "04. Credit Rating Analysis", heading: HeadingLevel.HEADING_1, spacing: { before: 400, after: 200 } }),
            new Paragraph({ children: [new TextRun({ text: `Rating: ${memo.creditRatingAnalysis.rating}`, bold: true })] }),
            new Paragraph({ text: `Rationale: ${memo.creditRatingAnalysis.rationale}` }),
            new Paragraph({ text: `Key Drivers: ${memo.creditRatingAnalysis.keyDrivers.join(', ')}` }),
            new Paragraph({ text: `Outlook: ${memo.creditRatingAnalysis.outlook}` }),
            new Paragraph({ text: `Downgrade Triggers: ${memo.creditRatingAnalysis.downgradeTriggers.join(', ')}`, spacing: { after: 200 } }),

            new Paragraph({ text: "05. Financial Performance Analysis", heading: HeadingLevel.HEADING_1, spacing: { before: 400, after: 200 } }),
            new Paragraph({ text: memo.financialPerformanceAnalysis.discussion }),
            ...(capturedImages['financial-performance-analysis'] ? [
              new Paragraph({
                children: [new ImageRun({ 
                  data: capturedImages['financial-performance-analysis']!.data, 
                  transformation: { width: 600, height: 600 * capturedImages['financial-performance-analysis']!.aspectRatio },
                  type: 'png'
                })],
                alignment: AlignmentType.CENTER
              })
            ] : []),

            new Paragraph({ text: "06. Balance Sheet Strength", heading: HeadingLevel.HEADING_1, spacing: { before: 400, after: 200 } }),
            new Paragraph({ text: memo.balanceSheetStrength.discussion }),
            ...(capturedImages['balance-sheet-strength'] ? [
              new Paragraph({
                children: [new ImageRun({ 
                  data: capturedImages['balance-sheet-strength']!.data, 
                  transformation: { width: 600, height: 600 * capturedImages['balance-sheet-strength']!.aspectRatio },
                  type: 'png'
                })],
                alignment: AlignmentType.CENTER
              })
            ] : []),

            new Paragraph({ text: "07. Asset Quality and Credit Risk", heading: HeadingLevel.HEADING_1, spacing: { before: 400, after: 200 } }),
            new Paragraph({ text: memo.assetQualityCreditRisk.discussion }),
            ...(capturedImages['asset-quality-credit-risk'] ? [
              new Paragraph({
                children: [new ImageRun({ 
                  data: capturedImages['asset-quality-credit-risk']!.data, 
                  transformation: { width: 600, height: 600 * capturedImages['asset-quality-credit-risk']!.aspectRatio },
                  type: 'png'
                })],
                alignment: AlignmentType.CENTER
              })
            ] : []),

            new Paragraph({ text: "08. Yield and Relative Value Analysis", heading: HeadingLevel.HEADING_1, spacing: { before: 400, after: 200 } }),
            new Paragraph({ text: memo.yieldRelativeValueAnalysis.spreadCalculation }),
            ...(capturedImages['yield-relative-value-analysis'] ? [
              new Paragraph({
                children: [new ImageRun({ 
                  data: capturedImages['yield-relative-value-analysis']!.data, 
                  transformation: { width: 600, height: 600 * capturedImages['yield-relative-value-analysis']!.aspectRatio },
                  type: 'png'
                })],
                alignment: AlignmentType.CENTER
              })
            ] : []),

            new Paragraph({ text: "09. Liquidity and Marketability", heading: HeadingLevel.HEADING_1, spacing: { before: 400, after: 200 } }),
            new Paragraph({ text: memo.liquidityMarketability.discussion }),
            new Paragraph({ text: `Listing Exchange: ${memo.liquidityMarketability.listingExchange}` }),
            new Paragraph({ text: `Typical Liquidity: ${memo.liquidityMarketability.typicalLiquidity}` }),
            new Paragraph({ text: `Bid-Ask Spreads: ${memo.liquidityMarketability.bidAskSpreads}` }),
            new Paragraph({ text: `Institutional Participation: ${memo.liquidityMarketability.institutionalParticipation}`, spacing: { after: 200 } }),

            new Paragraph({ text: "10. Risk Analysis", heading: HeadingLevel.HEADING_1, spacing: { before: 400, after: 200 } }),
            ...(capturedImages['risk-analysis'] ? [
              new Paragraph({
                children: [new ImageRun({ 
                  data: capturedImages['risk-analysis']!.data, 
                  transformation: { width: 600, height: 600 * capturedImages['risk-analysis']!.aspectRatio },
                  type: 'png'
                })],
                alignment: AlignmentType.CENTER
              })
            ] : []),

            new Paragraph({ text: "11. SWOT Analysis", heading: HeadingLevel.HEADING_1, spacing: { before: 400, after: 200 } }),
            ...(capturedImages['swot-analysis'] ? [
              new Paragraph({
                children: [new ImageRun({ 
                  data: capturedImages['swot-analysis']!.data, 
                  transformation: { width: 600, height: 600 * capturedImages['swot-analysis']!.aspectRatio },
                  type: 'png'
                })],
                alignment: AlignmentType.CENTER
              })
            ] : []),

            new Paragraph({ text: "12. ALM Suitability", heading: HeadingLevel.HEADING_1, spacing: { before: 400, after: 200 } }),
            new Paragraph({ children: [new TextRun({ text: `Duration: ${memo.almSuitability.duration}`, bold: true })] }),
            new Paragraph({ text: memo.almSuitability.explanation }),
            new Paragraph({ text: `Liability Matching: ${memo.almSuitability.liabilityMatching}`, spacing: { after: 200 } }),

            new Paragraph({ text: "13. Financial Annexure", heading: HeadingLevel.HEADING_1, spacing: { before: 400, after: 200 } }),
            ...(capturedImages['annexure-table'] ? [
              new Paragraph({
                children: [new ImageRun({ 
                  data: capturedImages['annexure-table']!.data, 
                  transformation: { width: 600, height: 600 * capturedImages['annexure-table']!.aspectRatio },
                  type: 'png'
                })],
                alignment: AlignmentType.CENTER
              })
            ] : []),

            new Paragraph({ text: "14. Final Investment Recommendation", heading: HeadingLevel.HEADING_1, spacing: { before: 400, after: 200 } }),
            new Paragraph({ children: [new TextRun({ text: `Verdict: ${memo.finalInvestmentRecommendation.verdict}`, bold: true })] }),
            new Paragraph({ text: memo.finalInvestmentRecommendation.justification }),
            new Paragraph({ text: `Credit Strength: ${memo.finalInvestmentRecommendation.creditStrength}` }),
            new Paragraph({ text: `Yield Attractiveness: ${memo.finalInvestmentRecommendation.yieldAttractiveness}` }),
            new Paragraph({ text: `Relative Spread: ${memo.finalInvestmentRecommendation.relativeSpread}` }),
            new Paragraph({ text: `Investor Suitability: ${memo.finalInvestmentRecommendation.investorSuitability}`, spacing: { after: 200 } }),

            new Paragraph({ text: "Data Sources", heading: HeadingLevel.HEADING_1, spacing: { before: 400, after: 200 } }),
            new Paragraph({ text: memo.sources.join(', '), spacing: { after: 200 } }),
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
        setMemo(null);
        setLoading(true);
        try {
          const data = await geminiService.generateMemo(selectedBond);
          if (data && validateMemo(data)) {
            setMemo(data);
          } else {
            alert('Failed to generate a complete investment memo. The AI service might be experiencing high load or returned incomplete data. Please try again or select another bond.');
          }
        } catch (error) {
          console.error('Memo generation error:', error);
          alert('An error occurred while generating the memo. Please check your connection and try again.');
        } finally {
          setLoading(false);
        }
      };
      fetchMemo();
    }
  }, [selectedBond]);

  if (!selectedBond) {
    return (
      <div className="flex-1 flex flex-col items-center justify-center p-8 text-center bg-slate-50">
        <div className="max-w-md w-full">
          <div className="w-16 h-16 bg-accent/10 rounded-full flex items-center justify-center mb-6 mx-auto">
            <Search className="text-accent" size={32} />
          </div>
          <h3 className="text-xl font-bold mb-2 text-slate-900">Deep Dive Analyser</h3>
          <p className="text-sm text-slate-500 mb-8">
            Enter an ISIN or Bond Name to generate a full institutional investment memo.
          </p>
          
          <form onSubmit={handleSearch} className="relative group">
            <input 
              type="text" 
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              placeholder="Enter ISIN (e.g. INE...) or Issuer Name"
              className="w-full px-6 py-4 rounded-xl border border-border bg-white text-slate-900 focus:outline-none focus:ring-2 focus:ring-accent/50 shadow-sm transition-all pr-16"
              disabled={searching}
            />
            <button 
              type="submit"
              disabled={searching || !searchQuery.trim()}
              className="absolute right-2 top-2 bottom-2 px-4 bg-accent text-white rounded-lg flex items-center justify-center hover:bg-accent/90 transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
            >
              {searching ? <Loader2 size={20} className="animate-spin" /> : <ChevronRight size={20} />}
            </button>
          </form>
          
          <div className="mt-8 pt-8 border-t border-border">
            <p className="text-[10px] text-slate-400 uppercase tracking-widest font-bold mb-4">Or Browse Market</p>
            <button 
              onClick={onBack}
              className="text-sm font-bold text-accent hover:underline flex items-center gap-2 mx-auto"
            >
              Go to Market Explorer
              <ArrowUpRight size={14} />
            </button>
          </div>
        </div>
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
              onClick={() => {
                onSelectBond(null);
                setMemo(null);
              }}
              className="flex items-center gap-2 px-4 py-2 rounded-lg bg-white border border-border text-xs font-bold text-slate-700 hover:bg-slate-50 transition-all shadow-sm"
            >
              <Search size={14} className="text-accent" />
              New Search
            </button>
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

            <div id="instrument-overview" className="grid grid-cols-5 gap-4 p-4 bg-slate-50 rounded-xl border border-border">
              {Object.entries(memo.instrumentOverview || {}).map(([key, val]) => (
                <div key={key}>
                  <div className="text-[9px] text-slate-400 uppercase tracking-widest mb-1 font-bold">{key.replace(/([A-Z])/g, ' $1')}</div>
                  <div className="text-xs font-bold text-slate-800">{val}</div>
                </div>
              ))}
            </div>
          </header>

          <div className="space-y-12">
            {/* 02. ISSUER OVERVIEW */}
            <section id="issuer-overview" className="pdf-avoid-break pdf-section">
              <h3 className="text-xs font-bold text-accent uppercase tracking-[0.2em] mb-4">02. Issuer Overview</h3>
              <div className="mb-6">
                <p className="text-sm text-slate-800 leading-relaxed mb-6">{memo.issuerOverview?.description}</p>
                <div className="grid grid-cols-2 gap-4 mb-6">
                  <div className="p-3 rounded-lg bg-slate-50 border border-slate-100">
                    <div className="text-[9px] text-slate-500 uppercase tracking-widest mb-1 font-bold">Ownership Structure</div>
                    <div className="text-xs font-bold text-slate-900">{memo.issuerOverview?.ownershipStructure}</div>
                  </div>
                  <div className="p-3 rounded-lg bg-slate-50 border border-slate-100">
                    <div className="text-[9px] text-slate-500 uppercase tracking-widest mb-1 font-bold">Status</div>
                    <div className="text-xs font-bold text-slate-900">{memo.issuerOverview?.status}</div>
                  </div>
                  <div className="p-3 rounded-lg bg-slate-50 border border-slate-100">
                    <div className="text-[9px] text-slate-500 uppercase tracking-widest mb-1 font-bold">Sector Positioning</div>
                    <div className="text-xs font-bold text-slate-900">{memo.issuerOverview?.sectorPositioning}</div>
                  </div>
                  <div className="p-3 rounded-lg bg-slate-50 border border-slate-100">
                    <div className="text-[9px] text-slate-500 uppercase tracking-widest mb-1 font-bold">Strategic Role</div>
                    <div className="text-xs font-bold text-slate-900">{memo.issuerOverview?.strategicRole}</div>
                  </div>
                </div>
                <div className="space-y-4">
                  <div>
                    <h4 className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-2">Major Business Lines</h4>
                    <div className="flex flex-wrap gap-2">
                      {memo.issuerOverview?.majorBusinessLines.map((line, i) => (
                        <span key={i} className="px-2 py-1 bg-slate-100 text-slate-700 rounded text-[10px] font-medium border border-slate-200">{line}</span>
                      ))}
                    </div>
                  </div>
                  {memo.issuerOverview?.loanBookComposition && (
                    <div>
                      <h4 className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-2">Loan Book Composition</h4>
                      <p className="text-xs text-slate-700">{memo.issuerOverview.loanBookComposition}</p>
                    </div>
                  )}
                  {memo.issuerOverview?.management && memo.issuerOverview.management.length > 0 && (
                    <div>
                      <h4 className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-3">Key Management & Stakeholders</h4>
                      <div className="border border-border rounded-lg overflow-hidden">
                        <table className="w-full text-left border-collapse text-[11px]">
                          <thead>
                            <tr className="bg-slate-50 border-b border-border">
                              <th className="p-2 font-bold text-slate-500 uppercase tracking-wider">Name</th>
                              <th className="p-2 font-bold text-slate-500 uppercase tracking-wider">Designation</th>
                            </tr>
                          </thead>
                          <tbody>
                            {memo.issuerOverview.management.map((person, i) => (
                              <tr key={i} className="border-b border-border/50 last:border-0">
                                <td className="p-2 font-bold text-slate-900">{person.name}</td>
                                <td className="p-2 text-slate-600">{person.designation}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  )}
                </div>
              </div>
            </section>

            {/* 03. INDUSTRY OVERVIEW */}
            <section id="industry-overview" className="pdf-avoid-break pdf-section">
              <h3 className="text-xs font-bold text-accent uppercase tracking-[0.2em] mb-4">03. Industry Overview</h3>
              <div className="space-y-6">
                <p className="text-sm text-slate-800 leading-relaxed">{memo.industryOverview?.content}</p>
                <div className="grid grid-cols-2 gap-6">
                  <div className="p-4 rounded-xl bg-slate-50 border border-slate-100">
                    <h4 className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-3">Growth Drivers</h4>
                    <ul className="space-y-2">
                      {memo.industryOverview?.growthDrivers.map((d, i) => (
                        <li key={i} className="text-xs text-slate-700 flex gap-2"><span className="text-accent">•</span> {d}</li>
                      ))}
                    </ul>
                  </div>
                  <div className="p-4 rounded-xl bg-slate-50 border border-slate-100">
                    <h4 className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-3">Regulatory Framework</h4>
                    <p className="text-xs text-slate-700">{memo.industryOverview?.regulatoryFramework}</p>
                  </div>
                </div>
              </div>
            </section>

            {/* 04. CREDIT RATING ANALYSIS */}
            <section id="credit-rating-analysis" className="pdf-avoid-break pdf-section">
              <h3 className="text-xs font-bold text-accent uppercase tracking-[0.2em] mb-4">04. Credit Rating Analysis</h3>
              <div className="p-6 rounded-xl border border-border bg-slate-50">
                <div className="flex items-center gap-3 mb-4">
                  <ShieldCheck className="text-accent" size={20} />
                  <span className="text-lg font-bold text-slate-900">{memo.creditRatingAnalysis?.rating}</span>
                  <span className="px-2 py-0.5 bg-blue-100 text-blue-700 text-[10px] font-bold rounded border border-blue-200 uppercase">{memo.creditRatingAnalysis?.outlook} Outlook</span>
                </div>
                <p className="text-sm text-slate-800 mb-6 leading-relaxed">{memo.creditRatingAnalysis?.rationale}</p>
                <div className="grid grid-cols-2 gap-8">
                  <div>
                    <h4 className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-3">Key Rating Drivers</h4>
                    <ul className="space-y-2">
                      {memo.creditRatingAnalysis?.keyDrivers.map((d, i) => (
                        <li key={i} className="text-xs text-slate-700 flex gap-2"><span className="text-emerald-600">↑</span> {d}</li>
                      ))}
                    </ul>
                  </div>
                  <div>
                    <h4 className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-3">Potential Downgrade Triggers</h4>
                    <ul className="space-y-2">
                      {memo.creditRatingAnalysis?.downgradeTriggers.map((d, i) => (
                        <li key={i} className="text-xs text-slate-700 flex gap-2"><span className="text-rose-600">↓</span> {d}</li>
                      ))}
                    </ul>
                  </div>
                </div>
              </div>
            </section>

            {/* 05. FINANCIAL PERFORMANCE ANALYSIS */}
            <section id="financial-performance-analysis" className="pdf-avoid-break pdf-section">
              <h3 className="text-xs font-bold text-accent uppercase tracking-[0.2em] mb-4">05. Financial Performance Analysis</h3>
              <p className="text-sm text-slate-800 leading-relaxed mb-6">{memo.financialPerformanceAnalysis?.discussion}</p>
              <div className="space-y-8">
                {allTrends.length > 0 ? (
                  allTrends.map((trend, idx) => (
                    <div key={idx} className="p-6 rounded-xl border border-border bg-white shadow-sm">
                      <div className="flex justify-between items-center mb-6">
                        <h4 className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">{trend.metric} Trend</h4>
                        <div className="px-2 py-1 bg-accent/5 rounded text-[9px] font-bold text-accent uppercase tracking-wider border border-accent/10">
                          Sourced from Financial Annexure
                        </div>
                      </div>
                      <div className="flex items-end gap-3 h-40 pt-8">
                        {trend.values.map((v, i) => {
                          const numericValue = parseFloat(v.value.replace(/[^0-9.-]/g, '')) || 0;
                          const maxVal = Math.max(...trend.values.map(x => Math.abs(parseFloat(x.value.replace(/[^0-9.-]/g, '')) || 0)));
                          const height = maxVal > 0 ? (Math.abs(numericValue) / maxVal) * 100 : 0;
                          const isNegative = numericValue < 0;
                          
                          return (
                            <div key={i} className="flex-1 flex flex-col items-center gap-3 group h-full justify-end">
                              <div className="relative w-full flex flex-col items-center justify-end h-full">
                                <div 
                                  className={`w-full max-w-[48px] transition-all rounded-t-sm relative ${isNegative ? 'bg-rose-500/30 group-hover:bg-rose-500/50' : 'bg-accent/20 group-hover:bg-accent/40'}`}
                                  style={{ height: `${height}%` }}
                                >
                                  <div className="absolute -top-7 left-1/2 -translate-x-1/2 text-[10px] font-bold text-accent whitespace-nowrap bg-white/80 px-1.5 py-0.5 rounded shadow-sm border border-border/50 z-10">
                                    {v.value}
                                  </div>
                                </div>
                              </div>
                              <div className="text-[9px] font-bold text-slate-500 uppercase tracking-tighter">{v.year}</div>
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  ))
                ) : (
                  <div className="p-12 text-center border border-dashed border-border rounded-xl bg-slate-50">
                    <p className="text-sm text-slate-400 italic">Financial trend data is being analyzed and verified...</p>
                  </div>
                )}
              </div>
            </section>

            {/* 06. BALANCE SHEET STRENGTH */}
            <section id="balance-sheet-strength" className="pdf-avoid-break pdf-section">
              <h3 className="text-xs font-bold text-accent uppercase tracking-[0.2em] mb-4">06. Balance Sheet Strength</h3>
              <div className="p-6 rounded-xl border border-border bg-slate-50">
                <p className="text-sm text-slate-800 leading-relaxed mb-6">{memo.balanceSheetStrength?.discussion}</p>
                <div className="grid grid-cols-3 gap-4">
                  {memo.balanceSheetStrength?.metrics.map((m, i) => (
                    <div key={i} className="p-4 rounded-lg bg-white border border-border shadow-sm">
                      <div className="text-[9px] text-slate-500 uppercase tracking-widest mb-1 font-bold">{m.label}</div>
                      <div className="text-sm font-bold text-slate-900">{m.value}</div>
                    </div>
                  ))}
                </div>
              </div>
            </section>

            {/* 07. ASSET QUALITY AND CREDIT RISK */}
            <section id="asset-quality-credit-risk" className="pdf-avoid-break pdf-section">
              <h3 className="text-xs font-bold text-accent uppercase tracking-[0.2em] mb-4">07. Asset Quality and Credit Risk</h3>
              <div className="p-6 rounded-xl border border-border bg-slate-50">
                <p className="text-sm text-slate-800 leading-relaxed mb-6">{memo.assetQualityCreditRisk?.discussion}</p>
                <div className="grid grid-cols-3 gap-4">
                  {memo.assetQualityCreditRisk?.metrics.map((m, i) => (
                    <div key={i} className="p-4 rounded-lg bg-white border border-border shadow-sm">
                      <div className="text-[9px] text-slate-500 uppercase tracking-widest mb-1 font-bold">{m.label}</div>
                      <div className="text-sm font-bold text-slate-900">{m.value}</div>
                    </div>
                  ))}
                </div>
              </div>
            </section>

            {/* 08. YIELD AND RELATIVE VALUE ANALYSIS */}
            <section id="yield-relative-value-analysis" className="pdf-avoid-break pdf-section">
              <h3 className="text-xs font-bold text-accent uppercase tracking-[0.2em] mb-4">08. Yield and Relative Value Analysis</h3>
              <div className="space-y-6">
                <div className="p-4 rounded-xl bg-accent/5 border border-accent/10">
                  <h4 className="text-[10px] font-bold text-accent uppercase tracking-widest mb-2">Spread Calculation</h4>
                  <p className="text-sm font-bold text-slate-900">{memo.yieldRelativeValueAnalysis?.spreadCalculation}</p>
                </div>
                <div className="border border-border rounded-xl overflow-hidden">
                  <table className="w-full text-left border-collapse text-xs">
                    <thead>
                      <tr className="border-b border-border bg-slate-100">
                        <th className="p-3 font-bold text-slate-600 uppercase tracking-widest">Instrument</th>
                        <th className="p-3 font-bold text-slate-600 uppercase tracking-widest">Yield %</th>
                        <th className="p-3 font-bold text-slate-600 uppercase tracking-widest">Spread (bps)</th>
                      </tr>
                    </thead>
                    <tbody>
                      {memo.yieldRelativeValueAnalysis?.comparisonTable.map((row, i) => (
                        <tr key={i} className="border-b border-border/50">
                          <td className="p-3 font-medium text-slate-900">{row.instrument}</td>
                          <td className="p-3 font-mono text-accent font-bold">{row.yield}</td>
                          <td className="p-3 font-mono text-slate-700">{row.spread}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </section>

            {/* 09. LIQUIDITY AND MARKETABILITY */}
            <section id="liquidity-marketability" className="pdf-avoid-break pdf-section">
              <h3 className="text-xs font-bold text-accent uppercase tracking-[0.2em] mb-4">09. Liquidity and Marketability</h3>
              <div className="p-6 rounded-xl border border-border bg-slate-50">
                <p className="text-sm text-slate-800 leading-relaxed mb-6">{memo.liquidityMarketability?.discussion}</p>
                <div className="grid grid-cols-2 gap-4">
                  <div className="p-3 rounded-lg bg-white border border-border">
                    <div className="text-[9px] text-slate-500 uppercase tracking-widest mb-1 font-bold">Listing Exchange</div>
                    <div className="text-xs font-bold text-slate-900">{memo.liquidityMarketability?.listingExchange}</div>
                  </div>
                  <div className="p-3 rounded-lg bg-white border border-border">
                    <div className="text-[9px] text-slate-500 uppercase tracking-widest mb-1 font-bold">Typical Liquidity</div>
                    <div className="text-xs font-bold text-slate-900">{memo.liquidityMarketability?.typicalLiquidity}</div>
                  </div>
                  <div className="p-3 rounded-lg bg-white border border-border">
                    <div className="text-[9px] text-slate-500 uppercase tracking-widest mb-1 font-bold">Bid-Ask Spreads</div>
                    <div className="text-xs font-bold text-slate-900">{memo.liquidityMarketability?.bidAskSpreads}</div>
                  </div>
                  <div className="p-3 rounded-lg bg-white border border-border">
                    <div className="text-[9px] text-slate-500 uppercase tracking-widest mb-1 font-bold">Institutional Participation</div>
                    <div className="text-xs font-bold text-slate-900">{memo.liquidityMarketability?.institutionalParticipation}</div>
                  </div>
                </div>
              </div>
            </section>

            {/* 10. RISK ANALYSIS */}
            <section id="risk-analysis" className="pdf-avoid-break pdf-section">
              <h3 className="text-xs font-bold text-accent uppercase tracking-[0.2em] mb-4">10. Risk Analysis</h3>
              <div className="space-y-4">
                {memo.riskAnalysis?.map((risk, i) => (
                  <div key={i} className="p-5 rounded-xl border border-border bg-white shadow-sm">
                    <div className="flex items-center gap-2 mb-2">
                      <div className="w-1.5 h-1.5 rounded-full bg-rose-500" />
                      <h4 className="text-xs font-bold text-slate-900 uppercase tracking-wider">{risk.type}</h4>
                    </div>
                    <p className="text-sm text-slate-700 mb-3">{risk.description}</p>
                    <div className="p-3 rounded-lg bg-slate-50 border border-slate-100 text-[11px] text-slate-600 italic">
                      <span className="font-bold text-slate-400 uppercase mr-2">Evidence:</span>
                      {risk.evidence}
                    </div>
                  </div>
                ))}
              </div>
            </section>

            {/* 11. SWOT ANALYSIS */}
            <section id="swot-analysis" className="pdf-avoid-break pdf-section">
              <h3 className="text-xs font-bold text-accent uppercase tracking-[0.2em] mb-4">11. SWOT Analysis</h3>
              <div className="grid grid-cols-2 gap-4">
                <div className="p-6 rounded-xl border border-emerald-500/20 bg-emerald-50">
                  <div className="text-[10px] font-bold text-emerald-700 uppercase tracking-widest mb-4">Strengths</div>
                  <ul className="space-y-2">
                    {(memo.swotAnalysis?.strengths || []).map((s, i) => (
                      <li key={i} className="text-xs text-slate-800 flex gap-2">
                        <span className="text-emerald-600 font-bold">•</span> {s}
                      </li>
                    ))}
                  </ul>
                </div>
                <div className="p-6 rounded-xl border border-rose-500/20 bg-rose-50">
                  <div className="text-[10px] font-bold text-rose-700 uppercase tracking-widest mb-4">Weaknesses</div>
                  <ul className="space-y-2">
                    {(memo.swotAnalysis?.weaknesses || []).map((s, i) => (
                      <li key={i} className="text-xs text-slate-800 flex gap-2">
                        <span className="text-rose-600 font-bold">•</span> {s}
                      </li>
                    ))}
                  </ul>
                </div>
                <div className="p-6 rounded-xl border border-blue-500/20 bg-blue-50">
                  <div className="text-[10px] font-bold text-blue-700 uppercase tracking-widest mb-4">Opportunities</div>
                  <ul className="space-y-2">
                    {(memo.swotAnalysis?.opportunities || []).map((s, i) => (
                      <li key={i} className="text-xs text-slate-800 flex gap-2">
                        <span className="text-blue-600 font-bold">•</span> {s}
                      </li>
                    ))}
                  </ul>
                </div>
                <div className="p-6 rounded-xl border border-amber-500/20 bg-amber-50">
                  <div className="text-[10px] font-bold text-amber-700 uppercase tracking-widest mb-4">Threats</div>
                  <ul className="space-y-2">
                    {(memo.swotAnalysis?.threats || []).map((s, i) => (
                      <li key={i} className="text-xs text-slate-800 flex gap-2">
                        <span className="text-amber-600 font-bold">•</span> {s}
                      </li>
                    ))}
                  </ul>
                </div>
              </div>
            </section>

            {/* 12. ALM SUITABILITY */}
            <section id="alm-suitability" className="pdf-avoid-break pdf-section">
              <h3 className="text-xs font-bold text-accent uppercase tracking-[0.2em] mb-4">12. ALM Suitability</h3>
              <div className="p-6 rounded-xl border border-border bg-slate-50">
                <div className="flex items-center gap-3 mb-6">
                  <span className="px-3 py-1 rounded-full bg-accent text-white text-[10px] font-bold uppercase tracking-wider">
                    {memo.almSuitability?.duration} Duration Bucket
                  </span>
                </div>
                <div className="grid grid-cols-2 gap-8">
                  <div>
                    <h4 className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-2">Liability Matching Analysis</h4>
                    <p className="text-sm text-slate-800 leading-relaxed">{memo.almSuitability?.explanation}</p>
                  </div>
                  <div>
                    <h4 className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-2">ALM Fit Justification</h4>
                    <p className="text-sm text-slate-800 leading-relaxed">{memo.almSuitability?.liabilityMatching}</p>
                  </div>
                </div>
              </div>
            </section>

            {/* 13. FINANCIAL ANNEXURE */}
            <section id="annexure-table" className="pdf-page-break pdf-avoid-break pdf-section">
              <div className="flex justify-between items-end mb-6">
                <h3 className="text-xs font-bold text-accent uppercase tracking-[0.2em]">13. Financial Annexure</h3>
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
                          <td className="p-3 font-mono text-slate-700">{row.fy21 || 'N/A'}</td>
                          <td className="p-3 font-mono text-slate-700">{row.fy22 || 'N/A'}</td>
                          <td className="p-3 font-mono text-slate-700">{row.fy23 || 'N/A'}</td>
                          <td className="p-3 font-mono text-slate-700">{row.fy24 || 'N/A'}</td>
                          <td className="p-3 font-mono text-slate-700">{row.fy25 || 'N/A'}</td>
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

            {/* 14. FINAL INVESTMENT RECOMMENDATION */}
            <section id="final-recommendation" className="pdf-avoid-break pdf-section">
              <h3 className="text-xs font-bold text-accent uppercase tracking-[0.2em] mb-4">14. Final Investment Recommendation</h3>
              <div className="p-8 rounded-2xl bg-accent text-white shadow-lg">
                <div className="flex justify-between items-start mb-6">
                  <div>
                    <h4 className="text-lg font-black uppercase tracking-tight mb-1">NBHI Portfolio Recommendation</h4>
                    <p className="text-xs text-white/70 font-bold uppercase tracking-widest">Institutional Credit View</p>
                  </div>
                  <div className="px-4 py-2 bg-white text-accent rounded-xl font-black text-xl shadow-inner">
                    {memo.finalInvestmentRecommendation?.verdict}
                  </div>
                </div>
                
                <div className="grid grid-cols-2 gap-8 mb-8">
                  <div className="space-y-4">
                    <div>
                      <h5 className="text-[10px] font-bold text-white/60 uppercase tracking-widest mb-1">Credit Strength</h5>
                      <p className="text-sm font-medium">{memo.finalInvestmentRecommendation?.creditStrength}</p>
                    </div>
                    <div>
                      <h5 className="text-[10px] font-bold text-white/60 uppercase tracking-widest mb-1">Yield Attractiveness</h5>
                      <p className="text-sm font-medium">{memo.finalInvestmentRecommendation?.yieldAttractiveness}</p>
                    </div>
                  </div>
                  <div className="space-y-4">
                    <div>
                      <h5 className="text-[10px] font-bold text-white/60 uppercase tracking-widest mb-1">Relative Spread</h5>
                      <p className="text-sm font-medium">{memo.finalInvestmentRecommendation?.relativeSpread}</p>
                    </div>
                    <div>
                      <h5 className="text-[10px] font-bold text-white/60 uppercase tracking-widest mb-1">Investor Suitability</h5>
                      <p className="text-sm font-medium">{memo.finalInvestmentRecommendation?.investorSuitability}</p>
                    </div>
                  </div>
                </div>

                <div className="pt-6 border-t border-white/20">
                  <h5 className="text-[10px] font-bold text-white/60 uppercase tracking-widest mb-2">Final Justification</h5>
                  <p className="text-sm font-medium leading-relaxed italic">"{memo.finalInvestmentRecommendation?.justification}"</p>
                </div>
              </div>
            </section>

            {/* DATA SOURCES */}
            <section className="pdf-avoid-break pdf-section">
              <h3 className="text-xs font-bold text-accent uppercase tracking-[0.2em] mb-4">Data Sources</h3>
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


const BondTable = ({ bonds, onSelectBond }: { bonds: NCD[], onSelectBond: (b: NCD) => void }) => {
  return (
    <div className="mt-4 border border-border rounded-lg overflow-hidden bg-white shadow-sm overflow-x-auto">
      <table className="w-full text-left border-collapse min-w-[600px]">
        <thead>
          <tr className="border-b border-border bg-slate-50">
            <th className="p-3 text-[9px] font-bold text-slate-500 uppercase tracking-widest">Issuer</th>
            <th className="p-3 text-[9px] font-bold text-slate-500 uppercase tracking-widest">Coupon</th>
            <th className="p-3 text-[9px] font-bold text-slate-500 uppercase tracking-widest">Rating</th>
            <th className="p-3 text-[9px] font-bold text-slate-500 uppercase tracking-widest">Maturity</th>
            <th className="p-3 text-[9px] font-bold text-slate-500 uppercase tracking-widest">Verdict</th>
            <th className="p-3 text-[9px] font-bold text-slate-500 uppercase tracking-widest text-right">Action</th>
          </tr>
        </thead>
        <tbody>
          {bonds.map((bond, idx) => (
            <tr key={idx} className="border-b border-border/50 hover:bg-slate-50 transition-colors">
              <td className="p-3">
                <div className="font-bold text-xs text-slate-900">{bond.issuerName}</div>
                <div className="text-[9px] text-text-dim font-mono">{bond.isin}</div>
              </td>
              <td className="p-3 font-mono text-xs text-accent font-semibold">{bond.couponRate}%</td>
              <td className="p-3 text-xs font-semibold text-slate-700">{bond.rating}</td>
              <td className="p-3 text-xs text-slate-600">{bond.maturityDate}</td>
              <td className="p-3">
                <span className={`px-1.5 py-0.5 rounded text-[9px] font-bold ${
                  bond.verdict === 'BUY' ? 'bg-emerald-600 text-white' :
                  bond.verdict === 'HOLD' ? 'bg-amber-500 text-white' : 'bg-rose-600 text-white'
                }`}>
                  {bond.verdict}
                </span>
              </td>
              <td className="p-3 text-right">
                <button 
                  onClick={() => onSelectBond(bond)}
                  className="text-[10px] font-bold text-accent hover:underline"
                >
                  Deep Dive
                </button>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
};

const ChatAssistant = ({ 
  onSelectBond, 
  setActiveTab, 
  messages, 
  setMessages 
}: { 
  onSelectBond: (b: NCD) => void, 
  setActiveTab: (t: string) => void,
  messages: { role: 'user' | 'model'; text: string; bonds?: NCD[] }[],
  setMessages: React.Dispatch<React.SetStateAction<{ role: 'user' | 'model'; text: string; bonds?: NCD[] }[]>>
}) => {
  const [input, setInput] = useState('');
  const [loading, setLoading] = useState(false);
  const fileInputRef = React.useRef<HTMLInputElement>(null);

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setLoading(true);
    setMessages(prev => [...prev, { role: 'user', text: `Uploaded file: ${file.name}` }]);

    try {
      const reader = new FileReader();
      reader.onload = async (evt) => {
        const fileData = evt.target?.result;
        const wb = XLSX.read(fileData, { type: 'array' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const sheetData = XLSX.utils.sheet_to_json(ws, { header: 1 }) as any[][];
        
        // Extract potential bond names or ISINs (assuming they are in the first column or contain keywords)
        const bondQueries: string[] = [];
        sheetData.forEach(row => {
          row.forEach(cell => {
            if (typeof cell === 'string' && cell.length > 3) {
              // Simple heuristic: if it looks like a company name or ISIN
              if (/^[A-Z]{2}[0-9A-Z]{10}$/.test(cell.toUpperCase()) || cell.split(' ').length > 1) {
                bondQueries.push(cell);
              }
            }
          });
        });

        const uniqueQueries = Array.from(new Set(bondQueries)).slice(0, 5); // Limit to 5 for now
        
        if (uniqueQueries.length === 0) {
          setMessages(prev => [...prev, { role: 'model', text: "I couldn't find any clear bond names or ISINs in that file. Please ensure the file contains bond identifiers." }]);
          setLoading(false);
          return;
        }

        setMessages(prev => [...prev, { role: 'model', text: `Found ${uniqueQueries.length} potential bonds. Fetching market data...` }]);
        
        const fetchedBonds: NCD[] = [];
        for (const query of uniqueQueries) {
          const bond = await geminiService.searchBond(query);
          if (bond) fetchedBonds.push(bond);
        }

        if (fetchedBonds.length > 0) {
          setMessages(prev => [...prev, { 
            role: 'model', 
            text: `Here is the market data for the bonds found in your file:`,
            bonds: fetchedBonds
          }]);
        } else {
          setMessages(prev => [...prev, { role: 'model', text: "I found bond names but couldn't retrieve specific market data for them. Please try searching manually." }]);
        }
        setLoading(false);
      };
      reader.readAsArrayBuffer(file);
    } catch (error) {
      console.error("File upload error:", error);
      setMessages(prev => [...prev, { role: 'model', text: "Error processing the file. Please ensure it's a valid Excel or CSV file." }]);
      setLoading(false);
    }
  };

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
      const responseText = response.text || 'I encountered an error processing your request.';
      
      // Intent detection: GENERATE_MEMO
      const memoMatch = responseText.match(/\[GENERATE_MEMO:\s*(.*?)\]/);
      if (memoMatch) {
        const bondQuery = memoMatch[1].trim();
        setMessages(prev => [...prev, { role: 'model', text: `Initiating deep dive for: **${bondQuery}**. Redirecting you to the Analyser...` }]);
        
        const bond = await geminiService.searchBond(bondQuery);
        if (bond) {
          onSelectBond(bond);
          setActiveTab('analyser');
        } else {
          setMessages(prev => [...prev, { role: 'model', text: `Sorry, I couldn't find official details for "${bondQuery}" to start a deep dive. Please try again with a valid ISIN.` }]);
        }
        setLoading(false);
        return;
      }

      // Intent detection: SCAN_MARKET
      const scanMatch = responseText.match(/\[SCAN_MARKET:\s*(.*?)\]/);
      if (scanMatch) {
        setMessages(prev => [...prev, { role: 'model', text: `Scanning market for: **${scanMatch[1]}**...` }]);
        
        // Fetch data directly for the chat
        const data = await geminiService.getMarketScan();
        if (data && data.length > 0) {
          setMessages(prev => [...prev, { 
            role: 'model', 
            text: `Here are the top bonds matching your request:`,
            bonds: data.slice(0, 5)
          }]);
        } else {
          setMessages(prev => [...prev, { role: 'model', text: "I couldn't find any bonds matching those criteria at the moment." }]);
        }
        
        setLoading(false);
        return;
      }

      setMessages(prev => [...prev, { role: 'model', text: responseText }]);
    } catch (error) {
      setMessages(prev => [...prev, { role: 'model', text: 'Error: Failed to connect to research engine.' }]);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="flex-1 flex flex-col h-full bg-white text-slate-900">
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
            <div className={`max-w-[85%] p-4 rounded-2xl text-sm leading-relaxed shadow-sm ${
              msg.role === 'user' ? 'bg-accent/5 text-slate-800 border border-accent/20' : 'bg-slate-50 text-slate-700 border border-border'
            }`}>
              <ReactMarkdown>{msg.text}</ReactMarkdown>
              {msg.bonds && <BondTable bonds={msg.bonds} onSelectBond={onSelectBond} />}
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
        <div className="max-w-4xl mx-auto flex gap-3 items-center">
          <button 
            onClick={() => fileInputRef.current?.click()}
            className="p-3 rounded-xl border border-border bg-white text-slate-600 hover:text-accent hover:border-accent transition-all shadow-sm"
            title="Upload Excel/CSV"
          >
            <FileIcon size={20} />
          </button>
          <input 
            type="file" 
            ref={fileInputRef} 
            onChange={handleFileUpload} 
            accept=".xlsx,.xls,.csv,.xlsb" 
            className="hidden" 
          />
          <div className="flex-1 relative">
            <input 
              type="text"
              value={input}
              onChange={(e) => setInput(e.target.value)}
              onKeyDown={(e) => e.key === 'Enter' && handleSend()}
              placeholder="Ask about specific NCDs, yields, or ALM suitability..."
              className="w-full bg-white border border-border rounded-xl py-4 pl-6 pr-14 text-sm text-slate-900 focus:outline-none focus:border-accent transition-colors shadow-sm"
            />
            <button 
              onClick={handleSend}
              disabled={loading}
              className="absolute right-2 top-2 bottom-2 w-10 bg-accent text-white rounded-lg flex items-center justify-center hover:bg-accent/90 transition-colors disabled:opacity-50 shadow-sm"
            >
              <Send size={18} />
            </button>
          </div>
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
  const [chatHistory, setChatHistory] = useState<{ role: 'user' | 'model'; text: string; bonds?: NCD[] }[]>([
    { role: 'model', text: 'Welcome to NBHI Investment Assistant. How can I help you with NCD research today?' }
  ]);

  const handleSelectBond = (bond: NCD) => {
    setSelectedBond(bond);
    setActiveTab('analyser');
  };

  return (
    <div className="flex h-screen bg-bg text-slate-900 overflow-hidden">
      <Sidebar activeTab={activeTab} setActiveTab={setActiveTab} />
      
      <main className="flex-1 flex flex-col relative overflow-hidden bg-white">
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
                onSelectBond={setSelectedBond}
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
              <ChatAssistant 
                onSelectBond={handleSelectBond} 
                setActiveTab={setActiveTab} 
                messages={chatHistory}
                setMessages={setChatHistory}
              />
            </motion.div>
          )}
        </AnimatePresence>
      </main>
    </div>
  );
}
