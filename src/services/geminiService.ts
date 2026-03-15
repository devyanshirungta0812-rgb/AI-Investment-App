import { GoogleGenAI } from "@google/genai";

export interface NCD {
  issuerName: string;
  isin: string;
  marketLabel: 'UPCOMING' | 'PRIMARY MARKET' | 'SECONDARY MARKET';
  couponRate: number;
  rating: string;
  issuanceDate: string;
  maturityDate: string;
  almFit: 'High' | 'Medium' | 'Low';
  almExplanation: string;
  verdict: 'BUY' | 'HOLD' | 'SELL';
}

export interface InvestmentMemo {
  issuerName: string;
  generatedDate: string;
  verdict: 'BUY' | 'HOLD' | 'SELL';
  instrumentDetails: {
    isin: string;
    couponRate: string;
    rating: string;
    issueDate: string;
    maturityDate: string;
    marketType: string;
  };
  executiveSummary: string;
  companyOverview: {
    description: string;
    metrics: { label: string; value: string }[];
  };
  management: { name: string; designation: string }[];
  almSuitability: {
    duration: 'Short' | 'Medium' | 'Long';
    explanation: string;
    solvencyImpact: string;
  };
  instrumentOverview: string;
  yieldAnalysis: {
    ncdYield: string;
    gSecYield: string;
    fdYield: string;
    creditSpread: string;
    yieldPremium: string;
    significance: string;
  };
  businessIndustryOverview: string;
  financialPerformanceTable: {
    year: string;
    revenue: string;
    netProfit: string;
  }[];
  financialTrendData: {
    year: string;
    revenueValue: number;
    netProfitValue: number;
    revenueLabel: string;
    netProfitLabel: string;
  }[];
  creditProfile: {
    discussion: string;
    metrics: { label: string; value: string }[];
  };
  swot: {
    strengths: string[];
    weaknesses: string[];
    opportunities: string[];
    threats: string[];
  };
  recommendation: string;
  financialAnnexure: {
    metric: string;
    fy21: string;
    fy22: string;
    fy23: string;
    fy24: string;
    fy25: string;
  }[];
  sources: string[];
}

const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY || "" });

export const geminiService = {
  async getMarketScan(): Promise<NCD[]> {
    const response = await ai.models.generateContent({
      model: "gemini-3-flash-preview",
      contents: `Perform a LIVE market scan (as of March 2026) for Indian Corporate Non-Convertible Debentures (NCDs). 
      Search reliable sources: IndiaBonds, NSE Debt Market, GoldenPi, and Jiraaf.
      
      CRITICAL: You MUST provide the MOST RECENT data available.
      MANDATORY: Return at least 10 unique NCD instruments in the list.
      
      Filter for:
      - Rating: AA+ or AAA (from CRISIL, ICRA, or India Ratings)
      - Type: Corporate NCD
      - Availability: Must be active in Primary or Secondary markets.
      
      Return a JSON array of objects with exactly these fields:
      - issuerName: Full legal name of the issuer
      - isin: Valid 12-character ISIN
      - marketLabel: "UPCOMING" | "PRIMARY MARKET" | "SECONDARY MARKET"
      - couponRate: Number (e.g. 8.25)
      - rating: Agency + Rating (e.g. "CRISIL AAA")
      - issuanceDate: Date string (e.g. "2024-01-15")
      - maturityDate: Date string (e.g. "2030-01-15")
      - almFit: "High" | "Medium" | "Low" (based on insurance ALM suitability)
      - almExplanation: Short professional justification
      - verdict: "BUY" | "HOLD" | "SELL"`,
      config: {
        tools: [{ googleSearch: {} }],
        responseMimeType: "application/json",
      },
    });

    try {
      const text = response.text || "[]";
      
      // Robust JSON extraction for arrays or objects
      let jsonStr = "";
      const startBracket = text.indexOf('[');
      const startBrace = text.indexOf('{');
      
      // Determine if we are looking for an array or an object
      const start = (startBracket !== -1 && (startBrace === -1 || startBracket < startBrace)) ? startBracket : startBrace;
      const openChar = text[start];
      const closeChar = openChar === '[' ? ']' : '}';

      if (start !== -1) {
        let depth = 0;
        for (let i = start; i < text.length; i++) {
          if (text[i] === openChar) depth++;
          else if (text[i] === closeChar) depth--;
          
          if (depth === 0) {
            jsonStr = text.substring(start, i + 1);
            break;
          }
        }
      }

      if (!jsonStr) {
        throw new Error("No valid JSON found in response");
      }
      
      const result = JSON.parse(jsonStr);
      return Array.isArray(result) ? result : [result];
    } catch (e) {
      console.error("Failed to parse market scan", e);
      return [];
    }
  },

  async generateMemo(bond: NCD): Promise<InvestmentMemo | null> {
    const prompt = `Generate a complete institutional investment memo for ${bond.issuerName} (ISIN: ${bond.isin}) for an insurance ALM team. 
    
    CONSISTENCY RULE:
    You MUST use the following data already identified for this instrument:
    - Issuer: ${bond.issuerName}
    - ISIN: ${bond.isin}
    - Coupon Rate: ${bond.couponRate}%
    - Rating: ${bond.rating}
    - Issuance Date: ${bond.issuanceDate}
    - Maturity Date: ${bond.maturityDate}
    - Verdict: ${bond.verdict}
    
    CRITICAL FINANCIAL DATA EXTRACTION RULES:
    1. PRIMARY SOURCES: Screener.in > Moneycontrol Financials > Company Annual Report > NSE/BSE disclosures > Investor Presentation.
    2. FALLBACK SOURCES: Tickertape, Trendlyne, Investing.com, ET Markets, Business Standard, Economic Times, Yahoo Finance, Credit rating agency reports (CRISIL / ICRA / CARE / India Ratings).
    3. EXTRACTION METHOD:
       - Step 1: Search for "${bond.issuerName} Screener financials". Locate P&L table.
       - Step 2: Cross-verify with "${bond.issuerName} Moneycontrol financials".
       - Step 3: Prioritize: Annual Report > NSE/BSE filings > Screener.in > Moneycontrol > Investor Presentation.
       - Step 4: For non-listed banks/NBFCs/PSUs, use Annual reports, RBI disclosures, and Rating reports.
    4. MANDATORY: The "financialAnnexure" section MUST contain actual reported numbers for the last 5 financial years (FY21, FY22, FY23, FY24, FY25).
    5. CURRENCY: All values must be in ₹ Crore, except EPS which should be in ₹.
    6. VALIDATION:
       - Revenue MUST be greater than Net Profit for all years.
       - Numbers MUST follow a logical growth trend.
       - No financial year should be skipped.
       - Search fallback sources before returning "Not Available".
    7. LAST RESORT: Only return "Not Available" if the company is private and does not disclose statements, or data is absolutely not found in any listed source.
    
    Structure the JSON response exactly as follows:
    - issuerName: ${bond.issuerName}
    - generatedDate: Current date
    - verdict: ${bond.verdict}
    - instrumentDetails: { isin: "${bond.isin}", couponRate: "${bond.couponRate}%", rating: "${bond.rating}", issueDate: "${bond.issuanceDate}", maturityDate: "${bond.maturityDate}", marketType: "${bond.marketLabel}" }
    - executiveSummary: 3-4 concise institutional points on credit strength, yield, positioning, and ALM fit.
    - companyOverview: { 
        description: Core business model, segments, and scale.
        metrics: Array of { label, value } including Total assets, Loan book size, Branch network, Customer base.
      }
    - management: Array of { name, designation } including Chairman, MD, CEO, Executive Directors.
    - almSuitability: { 
        duration: "Short" | "Medium" | "Long", 
        explanation: Duration matching and reinvestment risk.
        solvencyImpact: Impact on insurance solvency.
      }
    - instrumentOverview: Bond structure, seniority, and secondary market liquidity.
    - yieldAnalysis: { ncdYield, gSecYield, fdYield, creditSpread, yieldPremium, significance }
    - businessIndustryOverview: Industry structure, market share, and competitive landscape.
    - financialPerformanceTable: Array of { year, revenue, netProfit } for FY21-FY25. MANDATORY: Always include denominations (e.g., "₹ Cr") in the values.
    - financialTrendData: Array of { year, revenueValue, netProfitValue, revenueLabel, netProfitLabel } for FY21-FY25 (use numeric values for graphing).
    - creditProfile: {
        discussion: Analysis of credit metrics and rating outlook.
        metrics: Array of { label, value } including GNPA, NNPA, Capital adequacy.
      }
    - swot: { strengths: string[], weaknesses: string[], opportunities: string[], threats: string[] }. MANDATORY: All 4 sections must be present. Content must be FACT and NUMBER HEAVY (e.g., mention specific growth percentages, market share numbers, or debt ratios), not just general text.
    - recommendation: Final institutional recommendation explaining yield, stability, ALM fit, and diversification.
    - financialAnnexure: Detailed table for FY21-FY25. MANDATORY: Include rows for "Total Revenue / Operating Income", "Operating Profit", "Net Profit", and "Earnings Per Share (EPS)". STRICT RULE: Populate actual numbers from the mentioned sources. Do NOT use "N/A" or "Not Available" if data exists in public domain.
    - sources: List of specific public sources used, cited accurately to reflect the retrieval process (e.g., "Screener.in Financial Statements for Revenue/PAT", "Moneycontrol for EPS verification", "Company Annual Report 2024 for Management details", etc.)`;

    const response = await ai.models.generateContent({
      model: "gemini-3-flash-preview",
      contents: prompt,
      config: {
        tools: [{ googleSearch: {} }],
        responseMimeType: "application/json",
      },
    });

    try {
      const text = response.text || "null";
      
      // Robust JSON extraction: find the first '{' and its matching '}'
      let jsonStr = "";
      const start = text.indexOf('{');
      if (start !== -1) {
        let depth = 0;
        for (let i = start; i < text.length; i++) {
          if (text[i] === '{') depth++;
          else if (text[i] === '}') depth--;
          
          if (depth === 0) {
            jsonStr = text.substring(start, i + 1);
            break;
          }
        }
      }

      if (!jsonStr) {
        throw new Error("No valid JSON object found in response");
      }
      
      const data = JSON.parse(jsonStr);

      // Normalization to prevent runtime errors in UI
      if (data.financialAnnexure && !Array.isArray(data.financialAnnexure)) {
        data.financialAnnexure = Object.entries(data.financialAnnexure).map(([metric, values]: [string, any]) => {
          const row: any = { metric };
          if (typeof values === 'object' && values !== null) {
            // Map FY21, fy21, 2021, etc. to fy21
            Object.entries(values).forEach(([k, v]) => {
              const key = k.toLowerCase().replace('fy', '');
              if (key.includes('21')) row.fy21 = v;
              else if (key.includes('22')) row.fy22 = v;
              else if (key.includes('23')) row.fy23 = v;
              else if (key.includes('24')) row.fy24 = v;
              else if (key.includes('25')) row.fy25 = v;
            });
          } else {
            row.fy25 = values;
          }
          return row;
        });
      }

      if (data.sources && !Array.isArray(data.sources)) {
        data.sources = typeof data.sources === 'string' ? [data.sources] : [];
      }

      return data;
    } catch (e) {
      console.error("Failed to generate memo", e);
      return null;
    }
  },

  async chat(message: string, history: { role: string; parts: { text: string }[] }[]) {
    const chat = ai.chats.create({
      model: "gemini-3-flash-preview",
      config: {
        systemInstruction: "You are the NBHI Investment Pro Suite AI. You provide institutional fixed-income research. Use Google Search for real-time data. Maintain a professional tone.",
        tools: [{ googleSearch: {} }],
      },
      history: history,
    });

    return await chat.sendMessage({ message });
  }
};
