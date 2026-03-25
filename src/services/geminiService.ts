import { GoogleGenAI } from "@google/genai";

export interface NCD {
  issuerName: string;
  isin: string;
  marketLabel: 'UPCOMING' | 'PRIMARY MARKET' | 'SECONDARY MARKET';
  bondType: string;
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
  verdict: 'BUY' | 'SELECTIVE BUY' | 'HOLD' | 'AVOID';
  
  // 1. Instrument Overview
  instrumentOverview: {
    issuerName: string;
    ncdSeries: string;
    couponRate: string;
    maturity: string;
    creditRating: string;
    issueSize: string;
    listingExchange: string;
    isin: string;
    yieldSpreadVsGSec: string;
  };
  
  // 2. Issuer Overview
  issuerOverview: {
    description: string;
    ownershipStructure: string;
    status: string; // PSU / Private
    sectorPositioning: string;
    majorBusinessLines: string[];
    loanBookComposition?: string; // if NBFC
    strategicRole: string;
    management: { name: string; designation: string }[];
  };
  
  // 3. Industry Overview
  industryOverview: {
    content: string;
    growthDrivers: string[];
    regulatoryFramework: string;
    risks: string[];
  };
  
  // 4. Credit Rating Analysis
  creditRatingAnalysis: {
    rating: string;
    rationale: string;
    keyDrivers: string[];
    outlook: string;
    downgradeTriggers: string[];
  };
  
  // 5. Financial Performance Analysis
  financialPerformanceAnalysis: {
    discussion: string;
    trends: {
      metric: string;
      values: { year: string; value: string }[];
    }[];
  };
  
  // 6. Balance Sheet Strength
  balanceSheetStrength: {
    discussion: string;
    metrics: { label: string; value: string }[];
  };
  
  // 7. Asset Quality and Credit Risk
  assetQualityCreditRisk: {
    discussion: string;
    metrics: { label: string; value: string }[];
  };
  
  // 8. Yield and Relative Value Analysis
  yieldRelativeValueAnalysis: {
    comparisonTable: { instrument: string; yield: string; spread: string }[];
    spreadCalculation: string;
  };
  
  // 9. Liquidity and Marketability
  liquidityMarketability: {
    listingExchange: string;
    typicalLiquidity: string;
    bidAskSpreads: string;
    institutionalParticipation: string;
    discussion: string;
  };
  
  // 10. Risk Analysis
  riskAnalysis: {
    type: string;
    description: string;
    evidence: string;
  }[];
  
  // 11. SWOT Analysis
  swotAnalysis: {
    strengths: string[];
    weaknesses: string[];
    opportunities: string[];
    threats: string[];
  };
  
  // 12. ALM Suitability
  almSuitability: {
    duration: string;
    explanation: string;
    liabilityMatching: string;
  };
  
  // 13. Financial Annexure
  financialAnnexure: {
    metric: string;
    fy21: string;
    fy22: string;
    fy23: string;
    fy24: string;
    fy25: string;
  }[];
  
  // 14. Final Investment Recommendation
  finalInvestmentRecommendation: {
    verdict: 'BUY' | 'SELECTIVE BUY' | 'HOLD' | 'AVOID';
    justification: string;
    creditStrength: string;
    yieldAttractiveness: string;
    relativeSpread: string;
    investorSuitability: string;
  };
  
  sources: string[];
}

const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY || "" });

export const geminiService = {
  async callAIWithRetry(fn: () => Promise<any>, retries = 3): Promise<any> {
    for (let i = 0; i <= retries; i++) {
      try {
        return await fn();
      } catch (e: any) {
        const errorMsg = JSON.stringify(e);
        const isRPCError = e?.message?.includes('Rpc failed') || e?.message?.includes('xhr error');
        const isQuotaError = e?.status === 429 || errorMsg.includes('429') || errorMsg.includes('RESOURCE_EXHAUSTED');
        
        if ((isRPCError || isQuotaError) && i < retries) {
          const delay = Math.pow(2, i) * 2000; // 2s, 4s, 8s, 16s...
          console.warn(`AI call failed (${isQuotaError ? 'Quota' : 'RPC'} error), retrying in ${delay}ms... (${i + 1}/${retries})`);
          await new Promise(r => setTimeout(r, delay));
          continue;
        }
        throw e;
      }
    }
  },

  async getMarketScan(): Promise<NCD[]> {
    try {
      const response = await this.callAIWithRetry(() => ai.models.generateContent({
        model: "gemini-3-flash-preview",
        contents: `Perform a LIVE market scan (as of March 2026) for Indian Corporate Bonds and Non-Convertible Debentures (NCDs). 
        Search reliable sources: IndiaBonds, NSE Debt Market, GoldenPi, and Jiraaf.
        
        CRITICAL: You MUST provide the MOST RECENT data available.
        MANDATORY: Return at least 10 unique Corporate Bond/NCD instruments in the list.
        
        Filter for:
        - Rating: AA+ or AAA (from CRISIL, ICRA, or India Ratings)
        - Type: Corporate Bond / NCD
        - Availability: Must be active in Primary or Secondary markets.
        
        Return a JSON array of objects with exactly these fields:
        - issuerName: Full legal name of the issuer
        - isin: Valid 12-character ISIN
        - marketLabel: "UPCOMING" | "PRIMARY MARKET" | "SECONDARY MARKET"
        - bondType: Type of bond (e.g., "Corporate Bond", "NCD")
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
      }));

      const text = response.text || "[]";
      const result = this.extractJSON(text);
      return Array.isArray(result) ? result : [result];
    } catch (e) {
      console.error("Market scan failed:", e);
      // If it's a 500/RPC error, it might be transient or tool-related
      return [];
    }
  },

  extractJSON(text: string): any {
    try {
      // Robust JSON extraction for arrays or objects
      let jsonStr = "";
      const startBracket = text.indexOf('[');
      const startBrace = text.indexOf('{');
      
      const start = (startBracket !== -1 && (startBrace === -1 || startBracket < startBrace)) ? startBracket : startBrace;
      if (start === -1) return null;

      const openChar = text[start];
      const closeChar = openChar === '[' ? ']' : '}';

      let depth = 0;
      for (let i = start; i < text.length; i++) {
        if (text[i] === openChar) depth++;
        else if (text[i] === closeChar) depth--;
        
        if (depth === 0) {
          jsonStr = text.substring(start, i + 1);
          break;
        }
      }

      return jsonStr ? JSON.parse(jsonStr) : null;
    } catch (e) {
      console.error("JSON extraction failed", e);
      return null;
    }
  },

  async generateMemo(bond: NCD): Promise<InvestmentMemo | null> {
    const prompt = `Generate a complete institutional investment memo for ${bond.issuerName} (ISIN: ${bond.isin}) for an insurance ALM team. 
    
    CRITICAL: The final recommendation verdict MUST be "${bond.verdict}". Do not change this.

    Apply the following refinement rules within each section:

    1. Instrument Overview:
       Ensure all details are accurate and verified using issuer filings, exchange listing documents, and rating reports.
       Include: Issuer name, NCD series, coupon rate, maturity, credit rating, issue size, listing exchange, ISIN, yield spread vs G-sec.
       Verify coupon and maturity from official filings. Calculate spread vs current government bond yield using latest data.

    2. Issuer Overview:
       Factually correct and concise description using company website, annual reports, investor presentations, and financial databases.
       Include: ownership structure, PSU / private status, sector positioning, major business lines, loan book composition (if NBFC), strategic role in sector.
       MANDATORY: Include a list of key management personnel (Board of Directors, CEO, CFO, etc.) with their exact designations. Use only official company website or reliable financial portals (Moneycontrol, Bloomberg).
       No marketing language.

    3. Industry Overview:
       Data-backed overview using recent sector data (within last 2 years).
       Mention growth drivers, regulatory framework, and risks. Use government policy references.

    4. Credit Rating Analysis:
       Explain rationale, key drivers, outlook, and potential downgrade triggers using reports from CRISIL, ICRA, CARE, or India Ratings.
       Summarize analyst reasoning.

    5. Financial Performance Analysis:
       MANDATORY: Extract all numbers using live search (Moneycontrol, Screener, annual reports, exchange filings).
       Provide 3-5 year trends for: "Total Income" and "Net Profit". These MUST NOT be empty.
       These trends MUST be 100% consistent with the data provided in the Financial Annexure (Section 13).
       Also include: loan book (if NBFC), NIM, ROE, ROA, capital adequacy, leverage, asset quality (GNPA/NNPA).

    6. Balance Sheet Strength:
       Analyze using verified metrics: net worth, capital adequacy ratio, debt/equity, loan book growth, provisioning coverage, liquidity buffers.

    7. Asset Quality and Credit Risk:
       Factual analysis of GNPA/NNPA trends, sector exposure, borrower concentration, stressed assets, resolution trends.

    8. Yield and Relative Value Analysis:
       Bond spread analysis comparing NCD yield with G-secs, SDLs, AAA PSU bonds, and similar maturity corporate bonds.
       Calculate spread = NCD yield - government bond yield. Provide comparison tables.

    9. Liquidity and Marketability:
       Assess trading liquidity realistically. Discuss listing exchange, typical liquidity, expected bid-ask spreads, institutional participation.

    10. Risk Analysis:
        Structured section including credit, sector, regulatory, interest rate, refinancing, and concentration risks with supporting evidence.

    11. SWOT Analysis:
        Specific to the issuer and instrument.
        Strengths (balance sheet, govt support, rating), Weaknesses (concentration, dependency), Opportunities (sector growth, policy), Threats (regulatory, borrower distress).

    12. ALM Suitability:
        Analyze suitability for insurance, pension funds, debt funds, banks using duration and liability matching concepts.

    13. Financial Annexure:
        MANDATORY: Never display "Not Available". Retrieve data via live Google Search (Moneycontrol, Screener, Annual Reports).
        Present P&L statement for last 5 years (FY21 to FY25).
        You MUST include rows for exactly: "Total Income", "Interest Income", "Operating Expenses", "PBT", "Net Profit", and "EPS".
        These metrics are CRITICAL for the investment memo graphs. Ensure the values are accurate and consistent across all sections.

    14. Final Investment Recommendation:
        Balanced credit view. Verdict MUST be: ${bond.verdict}.
        Justify using credit risk, spread vs alternatives, and duration risk.

    FINAL QUALITY CHECK:
    - Verify all numbers using at least two financial sources.
    - No placeholder text or missing data.
    - Remove marketing language.
    - Analysis must resemble institutional credit research quality.

    Return the response as a JSON object matching this structure:
    {
      "issuerName": "Full Legal Name",
      "generatedDate": "Month Day, Year (e.g. March 16, 2026)",
      "verdict": "${bond.verdict}",
      "instrumentOverview": { 
        "issuerName": "...", 
        "ncdSeries": "...", 
        "couponRate": "e.g. 8.25%", 
        "maturity": "e.g. Jan 15, 2030", 
        "creditRating": "e.g. CRISIL AAA", 
        "issueSize": "e.g. ₹500 Cr", 
        "listingExchange": "NSE/BSE", 
        "isin": "...", 
        "yieldSpreadVsGSec": "e.g. 150 bps" 
      },
      "issuerOverview": { 
        "description": "...", 
        "ownershipStructure": "...", 
        "status": "...", 
        "sectorPositioning": "...", 
        "majorBusinessLines": ["..."], 
        "loanBookComposition": "...", 
        "strategicRole": "...",
        "management": [ { "name": "...", "designation": "..." } ]
      },
      "industryOverview": { "content": "...", "growthDrivers": ["..."], "regulatoryFramework": "...", "risks": ["..."] },
      "creditRatingAnalysis": { "rating": "...", "rationale": "...", "keyDrivers": ["..."], "outlook": "...", "downgradeTriggers": ["..."] },
      "financialPerformanceAnalysis": { "discussion": "...", "trends": [ { "metric": "...", "values": [ { "year": "...", "value": "..." } ] } ] },
      "balanceSheetStrength": { "discussion": "...", "metrics": [ { "label": "...", "value": "..." } ] },
      "assetQualityCreditRisk": { "discussion": "...", "metrics": [ { "label": "...", "value": "..." } ] },
      "yieldRelativeValueAnalysis": { "comparisonTable": [ { "instrument": "...", "yield": "...", "spread": "..." } ], "spreadCalculation": "..." },
      "liquidityMarketability": { "listingExchange": "...", "typicalLiquidity": "...", "bidAskSpreads": "...", "institutionalParticipation": "...", "discussion": "..." },
      "riskAnalysis": [ { "type": "...", "description": "...", "evidence": "..." } ],
      "swotAnalysis": { "strengths": ["..."], "weaknesses": ["..."], "opportunities": ["..."], "threats": ["..."] },
      "almSuitability": { "duration": "...", "explanation": "...", "liabilityMatching": "..." },
      "financialAnnexure": [ { "metric": "...", "fy21": "...", "fy22": "...", "fy23": "...", "fy24": "...", "fy25": "..." } ],
      "finalInvestmentRecommendation": { "verdict": "${bond.verdict}", "justification": "...", "creditStrength": "...", "yieldAttractiveness": "...", "relativeSpread": "...", "investorSuitability": "..." },
      "sources": ["..."]
    }`;

    try {
      const response = await this.callAIWithRetry(() => ai.models.generateContent({
        model: "gemini-3-flash-preview",
        contents: prompt,
        config: {
          tools: [{ googleSearch: {} }],
          responseMimeType: "application/json",
        },
      }));

      const text = response.text || "null";
      const data = this.extractJSON(text);
      if (!data) return null;

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

  async searchBond(query: string): Promise<NCD | null> {
    try {
      const response = await this.callAIWithRetry(() => ai.models.generateContent({
        model: "gemini-3-flash-preview",
        contents: `Search for the Indian Corporate Bond / NCD (Non-Convertible Debenture) matching this query: "${query}".
        Find the latest official details for this specific bond.
        
        Return a JSON object with exactly these fields:
        - issuerName: Full legal name of the issuer
        - isin: Valid 12-character ISIN
        - marketLabel: "UPCOMING" | "PRIMARY MARKET" | "SECONDARY MARKET" (Determine based on current status)
        - bondType: Type of bond (e.g., "Corporate Bond", "NCD")
        - couponRate: Number (e.g. 8.25)
        - rating: Agency + Rating (e.g. "CRISIL AAA")
        - issuanceDate: Date string (e.g. "2024-01-15")
        - maturityDate: Date string (e.g. "2030-01-15")
        - almFit: "High" | "Medium" | "Low" (based on insurance ALM suitability)
        - almExplanation: Short professional justification
        - verdict: "BUY" | "HOLD" | "SELL" (Provide a preliminary research-based verdict)`,
        config: {
          tools: [{ googleSearch: {} }],
          responseMimeType: "application/json",
        },
      }));

      const text = response.text || "null";
      return this.extractJSON(text);
    } catch (e) {
      console.error("Bond search failed:", e);
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

    return await this.callAIWithRetry(() => chat.sendMessage({ message }));
  }
};
