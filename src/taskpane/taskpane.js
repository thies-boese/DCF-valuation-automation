const PROXY_CONFIG = {
  baseUrl: 'http://localhost:3001/api',
  timeout: 10000
};

let searchTimeout;
let currentMatches = [];
let selectedIndex = -1;
let selectedSymbol = "";

//DATA VALIDATION & SANITIZATION

function sanitizeForExcel(value, type = 'auto') {
  if (value === null || value === undefined) {
    return "";
  }
  
  if (type === 'number' || typeof value === 'number') {
    if (isNaN(value) || !isFinite(value)) {
      return 0;
    }
    return Number(value);
  }
  
  if (type === 'string' || typeof value === 'string') {
    return String(value).trim();
  }
  
  if (typeof value === 'number') {
    return isNaN(value) || !isFinite(value) ? 0 : value;
  }
  
  if (typeof value === 'string') {
    return value.trim();
  }
  
  return String(value);
}

function validateDataArray(dataArray, arrayName) {
  console.log(`🔍 Validating ${arrayName}:`, dataArray);
  
  if (!Array.isArray(dataArray)) {
    console.error(`❌ ${arrayName} is not an array:`, typeof dataArray);
    return [];
  }
  
  const sanitizedArray = dataArray.map((row, rowIndex) => {
    if (!Array.isArray(row)) {
      console.warn(`⚠️ ${arrayName} row ${rowIndex} is not an array:`, row);
      return [String(row), ""];
    }
    
    return row.map((cell, cellIndex) => {
      const sanitized = sanitizeForExcel(cell);
      if (sanitized !== cell) {
        console.log(`🔧 Sanitized ${arrayName}[${rowIndex}][${cellIndex}]: ${cell} -> ${sanitized}`);
      }
      return sanitized;
    });
  });
  
  console.log(`✅ ${arrayName} validation complete. Rows: ${sanitizedArray.length}`);
  return sanitizedArray;
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("btnGetData").addEventListener("click", onGetDataClick);
    
    const queryInput = document.getElementById("queryInput");
    const dropdown = document.getElementById("dropdown");
    
    queryInput.addEventListener("input", onInputChange);
    queryInput.addEventListener("keydown", onKeyDown);
    queryInput.addEventListener("focus", onInputFocus);
    
    document.addEventListener("click", (e) => {
      if (!e.target.closest(".search-container")) {
        hideDropdown();
      }
    });

    testServerConnection();
  }
});

async function testServerConnection() {
  try {
    console.log('🔌 Testing server connection...');
    const response = await fetch('http://localhost:3001/health');
    if (response.ok) {
      const data = await response.json();
      console.log('✅ Proxy server connection successful:', data);
    } else {
      console.warn('⚠️ Proxy server responded with error:', response.status);
      showServerError();
    }
  } catch (error) {
    console.error('❌ Proxy server not available:', error);
    showServerError();
  }
}

function showServerError() {
  const statusElement = document.getElementById("status");
  statusElement.textContent = "⚠️ Proxy server not running. Please start the server: cd server && npm run dev";
  statusElement.style.color = "#d73a49";
}

//PROXY SERVER API HELPER

async function proxyFetch(endpoint) {
  try {
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), PROXY_CONFIG.timeout);
    
    const response = await fetch(`${PROXY_CONFIG.baseUrl}${endpoint}`, {
      signal: controller.signal
    });
    
    clearTimeout(timeoutId);
    
    if (!response.ok) {
      throw new Error(`HTTP ${response.status}: ${response.statusText}`);
    }
    
    return await response.json();
  } catch (error) {
    if (error.name === 'AbortError') {
      throw new Error('Request timeout - server may be overloaded');
    }
    console.error('Proxy fetch error:', error);
    throw error;
  }
}

async function searchCompanies(query) {
  const endpoint = `/search/${encodeURIComponent(query)}`;
  const result = await proxyFetch(endpoint);
  
  if (!result.success) {
    throw new Error(result.error || 'Search failed');
  }
  
  return result.data || [];
}

async function getCompanyData(symbol) {
  const endpoint = `/company/${encodeURIComponent(symbol)}`;
  const result = await proxyFetch(endpoint);
  
  if (!result.success) {
    throw new Error(result.error || 'Failed to get company data');
  }
  
  return result.data;
}

async function getHistoricalData(symbol) {
	const endpoint = `/historical/${encodeURIComponent(symbol)}`;
	try {
		const result = await proxyFetch(endpoint);
		
		if (!result.success) {
			console.warn('Historical data not available:', result.error);
			return null;
		}
		
		return result.data;
	} catch (error) {
		console.error('Historical fetch error:', error);
		return null;
	}
}

//TIMESERIES VARIANT → STATEMENT MAPPING

const VARIANT_TO_STATEMENT = {
  // INCOME STATEMENT (ANNUAL) - Revenue
  "annualTotalRevenue": "Income Statement",
  "annualOperatingRevenue": "Income Statement",
  "annualForeignSales": "Income Statement",
  "annualDomesticSales": "Income Statement",
  "annualAdjustedGeographySegmentData": "Income Statement",

  // Cost of Revenue
  "annualCostOfRevenue": "Income Statement",
  "annualOtherCostofRevenue": "Income Statement",
  "annualReconciledCostOfRevenue": "Income Statement",

  // Gross Profit
  "annualGrossProfit": "Income Statement",

  // Operating Expenses
  "annualSellingGeneralAndAdministrative": "Income Statement",
  "annualSellingAndMarketingExpense": "Income Statement",
  "annualGeneralAndAdministrativeExpense": "Income Statement",
  "annualOtherGandA": "Income Statement",
  "annualResearchAndDevelopment": "Income Statement",
  "annualProvisionForDoubtfulAccounts": "Income Statement",
  "annualInsuranceAndClaims": "Income Statement",
  "annualSalariesAndWages": "Income Statement",
  "annualProfessionalExpenseAndContractServicesExpense": "Income Statement",
  "annualOccupancyAndEquipment": "Income Statement",
  "annualEquipment": "Income Statement",
  "annualNetOccupancyExpense": "Income Statement",
  "annualOperationAndMaintenance": "Income Statement",
  "annualExplorationDevelopmentAndMineralPropertyLeaseExpenses": "Income Statement",
  "annualOtherOperatingExpenses": "Income Statement",
  "annualOperatingExpense": "Income Statement",

  // Depreciation, Amortization & Depletion
  "annualDepreciationAmortizationDepletionIncomeStatement": "Income Statement",
  "annualDepreciationIncomeStatement": "Income Statement",
  "annualDepreciationAndAmortizationInIncomeStatement": "Income Statement",
  "annualAmortization": "Income Statement",
  "annualAmortizationOfIntangiblesIncomeStatement": "Income Statement",
  "annualDepletionIncomeStatement": "Income Statement",
  "annualReconciledDepreciation": "Income Statement",

  // Operating Income
  "annualOperatingIncome": "Income Statement",
  "annualTotalOperatingIncomeAsReported": "Income Statement",

  // Non-Operating Income & Expenses
  "annualInterestExpense": "Income Statement",
  "annualInterestExpenseNonOperating": "Income Statement",
  "annualOtherInterestExpense": "Income Statement",
  "annualInterestIncome": "Income Statement",
  "annualInterestIncomeNonOperating": "Income Statement",
  "annualOtherInterestIncome": "Income Statement",
  "annualNetInterestIncome": "Income Statement",
  "annualNetNonOperatingInterestIncomeExpense": "Income Statement",
  "annualTotalOtherFinanceCost": "Income Statement",
  "annualLossonExtinguishmentofDebt": "Income Statement",

  // Other Income / Special Items
  "annualOtherIncomeExpense": "Income Statement",
  "annualOtherNonOperatingIncomeExpenses": "Income Statement",
  "annualWriteOff": "Income Statement",
  "annualSpecialIncomeCharges": "Income Statement",
  "annualOtherSpecialCharges": "Income Statement",
  "annualGainOnSaleOfPPE": "Income Statement",
  "annualGainOnSaleOfBusiness": "Income Statement",
  "annualGainOnSaleOfSecurity": "Income Statement",
  "annualGainLossonSaleofAssets": "Income Statement",
  "annualGainonSaleofInvestmentProperty": "Income Statement",
  "annualGainonSaleofLoans": "Income Statement",
  "annualForeignExchangeTradingGains": "Income Statement",
  "annualTradingGainLoss": "Income Statement",
  "annualInvestmentBankingProfit": "Income Statement",

  // Equity & Associate Income
  "annualEarningsFromEquityInterest": "Income Statement",
  "annualEarningsFromEquityInterestNetOfTax": "Income Statement",
  "annualIncomefromAssociatesandOtherParticipatingInterests": "Income Statement",

  // Taxes
  "annualTaxProvision": "Income Statement",
  "annualOtherTaxes": "Income Statement",
  "annualExciseTaxes": "Income Statement",
  "annualTaxRateForCalcs": "Income Statement",
  "annualTaxEffectOfUnusualItems": "Income Statement",

  // Net Income
  "annualPretaxIncome": "Income Statement",
  "annualNetIncome": "Income Statement",
  "annualMinorityInterests": "Income Statement",

  // EPS & Shares
  "annualBasicAverageShares": "Income Statement",
  "annualDilutedAverageShares": "Income Statement",
  "annualBasicEPS": "Income Statement",
  "annualNormalizedBasicEPS": "Income Statement",
  "annualReportedNormalizedBasicEPS": "Income Statement",
  "annualContinuingAndDiscontinuedBasicEPS": "Income Statement",
  "annualBasicEPSOtherGainsLosses": "Income Statement",
  "annualTaxLossCarryforwardBasicEPS": "Income Statement",
  "annualBasicContinuousOperations": "Income Statement",
  "annualBasicDiscontinuousOperations": "Income Statement",
  "annualBasicExtraordinary": "Income Statement",
  "annualBasicAccountingChange": "Income Statement",
  "annualDilutedEPS": "Income Statement",
  "annualNormalizedDilutedEPS": "Income Statement",
  "annualReportedNormalizedDilutedEPS": "Income Statement",
  "annualContinuingAndDiscontinuedDilutedEPS": "Income Statement",
  "annualDilutedEPSOtherGainsLosses": "Income Statement",
  "annualTaxLossCarryforwardDilutedEPS": "Income Statement",
  "annualDilutedContinuousOperations": "Income Statement",
  "annualDilutedDiscontinuousOperations": "Income Statement",
  "annualDilutedExtraordinary": "Income Statement",
  "annualDilutedAccountingChange": "Income Statement",
  "annualDilutedNIAvailtoComStockholders": "Income Statement",

  // Dividends
  "annualDividendPerShare": "Income Statement",
  "annualPreferredStockDividends": "Income Statement",
  "annualOtherunderPreferredStockDividend": "Income Statement",
  "annualOtherThanPreferredStockDividend": "Income Statement",

  // Performance Metrics
  "annualEBIT": "Income Statement",
  "annualEBITDA": "Income Statement",
  "annualNormalizedEBITDA": "Income Statement",
  "annualTaxRateForCalcs": "Income Statement",

  // BALANCE SHEET (ANNUAL) - Assets - Current
  "annualCashFinancial": "Balance Sheet",
  "annualCashEquivalents": "Balance Sheet",
  "annualCashAndCashEquivalents": "Balance Sheet",
  "annualOtherShortTermInvestments": "Balance Sheet",
  "annualCashCashEquivalentsAndShortTermInvestments": "Balance Sheet",
  "annualRestrictedCash": "Balance Sheet",
  "annualRestrictedCashAndCashEquivalents": "Balance Sheet",
  "annualReceivables": "Balance Sheet",
  "annualAccountsReceivable": "Balance Sheet",
  "annualGrossAccountsReceivable": "Balance Sheet",
  "annualAllowanceForDoubtfulAccountsReceivable": "Balance Sheet",
  "annualOtherReceivables": "Balance Sheet",
  "annualNotesReceivable": "Balance Sheet",
  "annualLoansReceivable": "Balance Sheet",
  "annualAccruedInterestReceivable": "Balance Sheet",
  "annualTaxesReceivable": "Balance Sheet",
  "annualInventory": "Balance Sheet",
  "annualOtherInventories": "Balance Sheet",
  "annualFinishedGoods": "Balance Sheet",
  "annualWorkInProcess": "Balance Sheet",
  "annualRawMaterials": "Balance Sheet",
  "annualInventoriesAdjustmentsAllowances": "Balance Sheet",
  "annualPrepaidAssets": "Balance Sheet",
  "annualCurrentDeferredAssets": "Balance Sheet",
  "annualCurrentDeferredTaxesAssets": "Balance Sheet",
  "annualOtherCurrentAssets": "Balance Sheet",
  "annualAssetsHeldForSaleCurrent": "Balance Sheet",
  "annualCurrentAssets": "Balance Sheet",

  // Assets - Non-Current
  "annualInvestmentsAndAdvances": "Balance Sheet",
  "annualInvestmentinFinancialAssets": "Balance Sheet",
  "annualOtherInvestments": "Balance Sheet",
  "annualHeldToMaturitySecurities": "Balance Sheet",
  "annualAvailableForSaleSecurities": "Balance Sheet",
  "annualTradingSecurities": "Balance Sheet",
  "annualLongTermEquityInvestment": "Balance Sheet",
  "annualGoodwill": "Balance Sheet",
  "annualOtherIntangibleAssets": "Balance Sheet",
  "annualGoodwillAndOtherIntangibleAssets": "Balance Sheet",
  "annualGrossPPE": "Balance Sheet",
  "annualAccumulatedDepreciation": "Balance Sheet",
  "annualNetPPE": "Balance Sheet",
  "annualMachineryFurnitureEquipment": "Balance Sheet",
  "annualBuildingsAndImprovements": "Balance Sheet",
  "annualLandAndImprovements": "Balance Sheet",
  "annualProperties": "Balance Sheet",
  "annualOtherProperties": "Balance Sheet",
  "annualConstructionInProgress": "Balance Sheet",
  "annualLeases": "Balance Sheet",
  "annualInvestmentProperties": "Balance Sheet",
  "annualMineralProperties": "Balance Sheet",
  "annualDeferredAssets": "Balance Sheet",
  "annualDeferredTaxAssets": "Balance Sheet",
  "annualOtherNonCurrentAssets": "Balance Sheet",
  "annualTotalNonCurrentAssets": "Balance Sheet",
  "annualTotalAssets": "Balance Sheet",

  // Liabilities – Current
  "annualAccountsPayable": "Balance Sheet",
  "annualPayables": "Balance Sheet",
  "annualOtherPayable": "Balance Sheet",
  "annualPayablesAndAccruedExpenses": "Balance Sheet",
  "annualCurrentAccruedExpenses": "Balance Sheet",
  "annualCurrentDebt": "Balance Sheet",
  "annualCurrentCapitalLeaseObligation": "Balance Sheet",
  "annualCurrentDebtAndCapitalLeaseObligation": "Balance Sheet",
  "annualOtherCurrentBorrowings": "Balance Sheet",
  "annualCommercialPaper": "Balance Sheet",
  "annualLineOfCredit": "Balance Sheet",
  "annualIncomeTaxPayable": "Balance Sheet",
  "annualTotalTaxPayable": "Balance Sheet",
  "annualInterestPayable": "Balance Sheet",
  "annualDividendsPayable": "Balance Sheet",
  "annualCurrentDeferredRevenue": "Balance Sheet",
  "annualCurrentDeferredLiabilities": "Balance Sheet",
  "annualCurrentDeferredTaxesLiabilities": "Balance Sheet",
  "annualOtherCurrentLiabilities": "Balance Sheet",
  "annualCurrentLiabilities": "Balance Sheet",

  // Liabilities – Non-Current
  "annualLongTermDebt": "Balance Sheet",
  "annualLongTermCapitalLeaseObligation": "Balance Sheet",
  "annualLongTermDebtAndCapitalLeaseObligation": "Balance Sheet",
  "annualNonCurrentDeferredRevenue": "Balance Sheet",
  "annualNonCurrentDeferredLiabilities": "Balance Sheet",
  "annualNonCurrentDeferredTaxesLiabilities": "Balance Sheet",
  "annualTradeandOtherPayablesNonCurrent": "Balance Sheet",
  "annualOtherNonCurrentLiabilities": "Balance Sheet",
  "annualTotalNonCurrentLiabilitiesNetMinorityInterest": "Balance Sheet",
  "annualTotalLiabilitiesNetMinorityInterest": "Balance Sheet",

  // Equity
  "annualCommonStock": "Balance Sheet",
  "annualPreferredStock": "Balance Sheet",
  "annualCapitalStock": "Balance Sheet",
  "annualAdditionalPaidInCapital": "Balance Sheet",
  "annualRetainedEarnings": "Balance Sheet",
  "annualGainsLossesNotAffectingRetainedEarnings": "Balance Sheet",
  "annualOtherEquityAdjustments": "Balance Sheet",
  "annualForeignCurrencyTranslationAdjustments": "Balance Sheet",
  "annualUnrealizedGainLoss": "Balance Sheet",
  "annualTreasuryStock": "Balance Sheet",
  "annualTreasurySharesNumber": "Balance Sheet",
  "annualCommonStockEquity": "Balance Sheet",
  "annualPreferredStockEquity": "Balance Sheet",
  "annualStockholdersEquity": "Balance Sheet",
  "annualTotalEquityGrossMinorityInterest": "Balance Sheet",
  "annualTotalCapitalization": "Balance Sheet",

  // Balance-sheet derived metrics
  "annualNetDebt": "Balance Sheet",
  "annualTotalDebt": "Balance Sheet",
  "annualInvestedCapital": "Balance Sheet",
  "annualWorkingCapital": "Balance Sheet",
  "annualNetTangibleAssets": "Balance Sheet",
  "annualTangibleBookValue": "Balance Sheet",

  // CASH FLOW STATEMENT (ANNUAL) - Operating Activities
  "annualOperatingCashFlow": "Cash Flow",
  "annualProvisionForLoanLeaseAndOtherLosses": "Cash Flow",
  "annualDepreciationAmortizationDepletion": "Cash Flow",
  "annualAmortizationOfFinancingCostsAndDiscounts": "Cash Flow",
  "annualChangeInLoans": "Cash Flow",
  "annualChangeInDeferredCharges": "Cash Flow",
  "annualIncreaseDecreaseInDeposit": "Cash Flow",
  "annualCashReceiptsfromCustomers": "Cash Flow",
  "annualPaymentstoSuppliersforGoodsandServices": "Cash Flow",
  "annualPaymentsonBehalfofEmployees": "Cash Flow",
  "annualOtherCashPaymentsfromOperatingActivities": "Cash Flow",
  "annualTaxesRefundPaidDirect": "Cash Flow",
  "annualInterestPaidDirect": "Cash Flow",
  "annualInterestReceivedDirect": "Cash Flow",
  "annualDividendsReceivedDirect": "Cash Flow",
  "annualDividendsPaidDirect": "Cash Flow",

  // Investing Activities
  "annualInvestingCashFlow": "Cash Flow",
  "annualCapitalExpenditure": "Cash Flow",
  "annualCapitalExpenditureReported": "Cash Flow",
  "annualPurchaseOfPPE": "Cash Flow",
  "annualSaleOfPPE": "Cash Flow",
  "annualNetPPEPurchaseAndSale": "Cash Flow",
  "annualPurchaseOfInvestment": "Cash Flow",
  "annualSaleOfInvestment": "Cash Flow",
  "annualNetInvestmentPurchaseAndSale": "Cash Flow",
  "annualPurchaseOfBusiness": "Cash Flow",
  "annualSaleOfBusiness": "Cash Flow",
  "annualNetBusinessPurchaseAndSale": "Cash Flow",
  "annualPurchaseOfIntangibles": "Cash Flow",
  "annualSaleOfIntangibles": "Cash Flow",
  "annualNetIntangiblesPurchaseAndSale": "Cash Flow",

  // Financing Activities
  "annualFinancingCashFlow": "Cash Flow",
  "annualIssuanceOfDebt": "Cash Flow",
  "annualRepaymentOfDebt": "Cash Flow",
  "annualNetIssuancePaymentsOfDebt": "Cash Flow",
  "annualIssuanceOfCapitalStock": "Cash Flow",
  "annualRepurchaseOfCapitalStock": "Cash Flow",
  "annualCommonStockIssuance": "Cash Flow",
  "annualCommonStockPayments": "Cash Flow",
  "annualNetCommonStockIssuance": "Cash Flow",
  "annualPreferredStockIssuance": "Cash Flow",
  "annualPreferredStockPayments": "Cash Flow",
  "annualNetPreferredStockIssuance": "Cash Flow",
  "annualCashDividendsPaid": "Cash Flow",
  "annualCommonStockDividendPaid": "Cash Flow",
  "annualPreferredStockDividendPaid": "Cash Flow",

  // Cash Reconciliation
  "annualBeginningCashPosition": "Cash Flow",
  "annualEffectOfExchangeRateChanges": "Cash Flow",
  "annualChangesInCash": "Cash Flow",
  "annualOtherCashAdjustmentInsideChangeinCash": "Cash Flow",
  "annualOtherCashAdjustmentOutsideChangeinCash": "Cash Flow",
  "annualEndCashPosition": "Cash Flow",

  // Free Cash Flow
  "annualFreeCashFlow": "Cash Flow"
};

const STATEMENT_SECTIONS = {
  "Income Statement": [
    { title: "Revenue", variants: ["annualTotalRevenue", "annualOperatingRevenue", "annualForeignSales", "annualDomesticSales", "annualAdjustedGeographySegmentData"] },
    { title: "Cost of Revenue", variants: ["annualCostOfRevenue", "annualOtherCostofRevenue", "annualReconciledCostOfRevenue"] },
    { title: "Gross Profit", variants: ["annualGrossProfit"] },
    { title: "Operating Expenses", variants: ["annualSellingGeneralAndAdministrative", "annualSellingAndMarketingExpense", "annualGeneralAndAdministrativeExpense", "annualOtherGandA", "annualResearchAndDevelopment", "annualProvisionForDoubtfulAccounts", "annualInsuranceAndClaims", "annualSalariesAndWages", "annualProfessionalExpenseAndContractServicesExpense", "annualOccupancyAndEquipment", "annualEquipment", "annualNetOccupancyExpense", "annualOperationAndMaintenance", "annualExplorationDevelopmentAndMineralPropertyLeaseExpenses", "annualOtherOperatingExpenses", "annualOperatingExpense"] },
    { title: "Depreciation, Amortization & Depletion", variants: ["annualDepreciationAmortizationDepletionIncomeStatement", "annualDepreciationIncomeStatement", "annualDepreciationAndAmortizationInIncomeStatement", "annualAmortization", "annualAmortizationOfIntangiblesIncomeStatement", "annualDepletionIncomeStatement", "annualReconciledDepreciation"] },
    { title: "Operating Income", variants: ["annualOperatingIncome", "annualTotalOperatingIncomeAsReported"] },
    { title: "Non-Operating Income & Expenses", variants: ["annualInterestExpense", "annualInterestExpenseNonOperating", "annualOtherInterestExpense", "annualInterestIncome", "annualInterestIncomeNonOperating", "annualOtherInterestIncome", "annualNetInterestIncome", "annualNetNonOperatingInterestIncomeExpense", "annualTotalOtherFinanceCost", "annualLossonExtinguishmentofDebt"] },
    { title: "Other Income / Special Items", variants: ["annualOtherIncomeExpense", "annualOtherNonOperatingIncomeExpenses", "annualWriteOff", "annualSpecialIncomeCharges", "annualOtherSpecialCharges", "annualGainOnSaleOfPPE", "annualGainOnSaleOfBusiness", "annualGainOnSaleOfSecurity", "annualGainLossonSaleofAssets", "annualGainonSaleofInvestmentProperty", "annualGainonSaleofLoans", "annualForeignExchangeTradingGains", "annualTradingGainLoss", "annualInvestmentBankingProfit"] },
    { title: "Equity & Associate Income", variants: ["annualEarningsFromEquityInterest", "annualEarningsFromEquityInterestNetOfTax", "annualIncomefromAssociatesandOtherParticipatingInterests"] },
    { title: "Taxes", variants: ["annualTaxProvision", "annualOtherTaxes", "annualExciseTaxes", "annualTaxRateForCalcs", "annualTaxEffectOfUnusualItems"] },
    { title: "Net Income", variants: ["annualPretaxIncome", "annualNetIncome", "annualNetIncomeContinuousOperations", "annualNetIncomeDiscontinuousOperations", "annualNetIncomeIncludingNoncontrollingInterests", "annualNetIncomeFromContinuingOperationNetMinorityInterest", "annualNetIncomeFromContinuingAndDiscontinuedOperation", "annualNetIncomeExtraordinary", "annualNetIncomeFromTaxLossCarryforward", "annualMinorityInterests"] },
    { title: "Earnings Per Share", variants: ["annualBasicAverageShares", "annualDilutedAverageShares", "annualBasicEPS", "annualNormalizedBasicEPS", "annualReportedNormalizedBasicEPS", "annualContinuingAndDiscontinuedBasicEPS", "annualBasicEPSOtherGainsLosses", "annualTaxLossCarryforwardBasicEPS", "annualBasicContinuousOperations", "annualBasicDiscontinuousOperations", "annualBasicExtraordinary", "annualBasicAccountingChange", "annualDilutedEPS", "annualNormalizedDilutedEPS", "annualReportedNormalizedDilutedEPS", "annualContinuingAndDiscontinuedDilutedEPS", "annualDilutedEPSOtherGainsLosses", "annualTaxLossCarryforwardDilutedEPS", "annualDilutedContinuousOperations", "annualDilutedDiscontinuousOperations", "annualDilutedExtraordinary", "annualDilutedAccountingChange", "annualDilutedNIAvailtoComStockholders"] },
    { title: "Dividends", variants: ["annualDividendPerShare", "annualPreferredStockDividends", "annualOtherunderPreferredStockDividend", "annualOtherThanPreferredStockDividend"] },
    { title: "Performance Metrics", variants: ["annualEBIT", "annualEBITDA", "annualNormalizedEBITDA"] }
  ],
  "Balance Sheet": [
    { title: "Assets - Current", variants: ["annualCashFinancial", "annualCashEquivalents", "annualCashAndCashEquivalents", "annualOtherShortTermInvestments", "annualCashCashEquivalentsAndShortTermInvestments", "annualRestrictedCash", "annualRestrictedCashAndCashEquivalents", "annualReceivables", "annualAccountsReceivable", "annualGrossAccountsReceivable", "annualAllowanceForDoubtfulAccountsReceivable", "annualOtherReceivables", "annualNotesReceivable", "annualLoansReceivable", "annualAccruedInterestReceivable", "annualTaxesReceivable", "annualInventory", "annualOtherInventories", "annualFinishedGoods", "annualWorkInProcess", "annualRawMaterials", "annualInventoriesAdjustmentsAllowances", "annualPrepaidAssets", "annualCurrentDeferredAssets", "annualCurrentDeferredTaxesAssets", "annualOtherCurrentAssets", "annualAssetsHeldForSaleCurrent", "annualCurrentAssets"] },
    { title: "Assets - Non-Current", variants: ["annualInvestmentsAndAdvances", "annualInvestmentinFinancialAssets", "annualOtherInvestments", "annualHeldToMaturitySecurities", "annualAvailableForSaleSecurities", "annualTradingSecurities", "annualLongTermEquityInvestment", "annualGoodwill", "annualOtherIntangibleAssets", "annualGoodwillAndOtherIntangibleAssets", "annualGrossPPE", "annualAccumulatedDepreciation", "annualNetPPE", "annualMachineryFurnitureEquipment", "annualBuildingsAndImprovements", "annualLandAndImprovements", "annualProperties", "annualOtherProperties", "annualConstructionInProgress", "annualLeases", "annualInvestmentProperties", "annualMineralProperties", "annualDeferredAssets", "annualDeferredTaxAssets", "annualOtherNonCurrentAssets", "annualTotalNonCurrentAssets", "annualTotalAssets"] },
    { title: "Liabilities - Current", variants: ["annualAccountsPayable", "annualPayables", "annualOtherPayable", "annualPayablesAndAccruedExpenses", "annualCurrentAccruedExpenses", "annualCurrentDebt", "annualCurrentCapitalLeaseObligation", "annualCurrentDebtAndCapitalLeaseObligation", "annualOtherCurrentBorrowings", "annualCommercialPaper", "annualLineOfCredit", "annualIncomeTaxPayable", "annualTotalTaxPayable", "annualInterestPayable", "annualDividendsPayable", "annualCurrentDeferredRevenue", "annualCurrentDeferredLiabilities", "annualCurrentDeferredTaxesLiabilities", "annualOtherCurrentLiabilities", "annualCurrentLiabilities"] },
    { title: "Liabilities - Non-Current", variants: ["annualLongTermDebt", "annualLongTermCapitalLeaseObligation", "annualLongTermDebtAndCapitalLeaseObligation", "annualNonCurrentDeferredRevenue", "annualNonCurrentDeferredLiabilities", "annualNonCurrentDeferredTaxesLiabilities", "annualTradeandOtherPayablesNonCurrent", "annualOtherNonCurrentLiabilities", "annualTotalNonCurrentLiabilitiesNetMinorityInterest", "annualTotalLiabilitiesNetMinorityInterest"] },
    { title: "Equity", variants: ["annualCommonStock", "annualPreferredStock", "annualCapitalStock", "annualAdditionalPaidInCapital", "annualRetainedEarnings", "annualGainsLossesNotAffectingRetainedEarnings", "annualOtherEquityAdjustments", "annualForeignCurrencyTranslationAdjustments", "annualUnrealizedGainLoss", "annualTreasuryStock", "annualTreasurySharesNumber", "annualCommonStockEquity", "annualPreferredStockEquity", "annualStockholdersEquity", "annualTotalEquityGrossMinorityInterest", "annualTotalCapitalization"] },
    { title: "Derived Metrics", variants: ["annualNetDebt", "annualTotalDebt", "annualInvestedCapital", "annualWorkingCapital", "annualNetTangibleAssets", "annualTangibleBookValue"] }
  ],
  "Cash Flow": [
    { title: "Operating Activities", variants: ["annualOperatingCashFlow", "annualProvisionForLoanLeaseAndOtherLosses", "annualDepreciationAmortizationDepletion", "annualAmortizationOfFinancingCostsAndDiscounts", "annualChangeInLoans", "annualChangeInDeferredCharges", "annualIncreaseDecreaseInDeposit", "annualCashReceiptsfromCustomers", "annualPaymentstoSuppliersforGoodsandServices", "annualPaymentsonBehalfofEmployees", "annualOtherCashPaymentsfromOperatingActivities", "annualTaxesRefundPaidDirect", "annualInterestPaidDirect", "annualInterestReceivedDirect", "annualDividendsReceivedDirect", "annualDividendsPaidDirect"] },
    { title: "Investing Activities", variants: ["annualInvestingCashFlow", "annualCapitalExpenditure", "annualCapitalExpenditureReported", "annualPurchaseOfPPE", "annualSaleOfPPE", "annualNetPPEPurchaseAndSale", "annualPurchaseOfInvestment", "annualSaleOfInvestment", "annualNetInvestmentPurchaseAndSale", "annualPurchaseOfBusiness", "annualSaleOfBusiness", "annualNetBusinessPurchaseAndSale", "annualPurchaseOfIntangibles", "annualSaleOfIntangibles", "annualNetIntangiblesPurchaseAndSale"] },
    { title: "Financing Activities", variants: ["annualFinancingCashFlow", "annualIssuanceOfDebt", "annualRepaymentOfDebt", "annualNetIssuancePaymentsOfDebt", "annualIssuanceOfCapitalStock", "annualRepurchaseOfCapitalStock", "annualCommonStockIssuance", "annualCommonStockPayments", "annualNetCommonStockIssuance", "annualPreferredStockIssuance", "annualPreferredStockPayments", "annualNetPreferredStockIssuance", "annualCashDividendsPaid", "annualCommonStockDividendPaid", "annualPreferredStockDividendPaid"] },
    { title: "Cash Reconciliation", variants: ["annualBeginningCashPosition", "annualEffectOfExchangeRateChanges", "annualChangesInCash", "annualOtherCashAdjustmentInsideChangeinCash", "annualOtherCashAdjustmentOutsideChangeinCash", "annualEndCashPosition"] },
    { title: "Free Cash Flow", variants: ["annualFreeCashFlow"] }
  ]
};

//DYNAMIC SEARCH

function onInputChange(e) {
  const query = e.target.value.trim();
  selectedIndex = -1;
  selectedSymbol = "";
  
  if (searchTimeout) {
    clearTimeout(searchTimeout);
  }
  
  if (query.length < 1) {
    hideDropdown();
    document.getElementById("status").textContent = "Type a company name or ticker symbol to search...";
    return;
  }
  
  searchTimeout = setTimeout(() => {
    performSearch(query);
  }, 400);
  
  showLoadingDropdown();
}

function onKeyDown(e) {
  const dropdown = document.getElementById("dropdown");
  
  if (dropdown.style.display === "none") return;
  
  switch (e.key) {
    case "ArrowDown":
      e.preventDefault();
      selectedIndex = Math.min(selectedIndex + 1, currentMatches.length - 1);
      updateSelection();
      break;
      
    case "ArrowUp":
      e.preventDefault();
      selectedIndex = Math.max(selectedIndex - 1, -1);
      updateSelection();
      break;
      
    case "Enter":
      e.preventDefault();
      if (selectedIndex >= 0 && currentMatches[selectedIndex]) {
        selectCompany(currentMatches[selectedIndex]);
      } else if (currentMatches.length > 0) {
        selectCompany(currentMatches[0]);
      }
      break;
      
    case "Escape":
      e.preventDefault();
      hideDropdown();
      break;
  }
}

function onInputFocus(e) {
  const query = e.target.value.trim();
  if (query.length >= 1 && currentMatches.length > 0) {
    showDropdown();
  }
}

async function performSearch(query) {
  const statusElement = document.getElementById("status");
  
  statusElement.textContent = `Searching for "${query}"...`;
  statusElement.style.color = "#555";
  
  try {
    const matches = await searchCompanies(query);
    
    currentMatches = matches;
    selectedIndex = -1;
    
    if (matches.length === 0) {
      statusElement.textContent = "No matches found. Try a different search term.";
      showNoResultsDropdown();
    } else {
      statusElement.textContent = `Found ${matches.length} match(es). Select one to get financial data.`;
      showMatchesDropdown(matches);
    }
    
  } catch (err) {
    console.error('Search error:', err);
    statusElement.textContent = `Search error: ${err.message}`;
    statusElement.style.color = "#d73a49";
    showErrorDropdown(err.message);
  }
}

function showLoadingDropdown() {
  const dropdown = document.getElementById("dropdown");
  dropdown.innerHTML = '<div class="loading">Searching...</div>';
  dropdown.style.display = "block";
}

function showMatchesDropdown(matches) {
  const dropdown = document.getElementById("dropdown");
  
  dropdown.innerHTML = matches.map((match, index) => {
    const symbol = match.symbol || "";
    const name = match.shortname || match.longname || "";
        
    return `
      <div class="dropdown-item" data-index="${index}">
        <div>
          <span class="company-symbol">${symbol}</span>
          <span class="company-name">${name}</span>
        </div>
      </div>
    `;
  }).join("");
  
  dropdown.querySelectorAll(".dropdown-item").forEach((item, index) => {
    item.addEventListener("click", () => {
      selectCompany(matches[index]);
    });
  });
  
  dropdown.style.display = "block";
}

function showNoResultsDropdown() {
  const dropdown = document.getElementById("dropdown");
  dropdown.innerHTML = '<div class="no-results">No companies found</div>';
  dropdown.style.display = "block";
}

function showErrorDropdown(error) {
  const dropdown = document.getElementById("dropdown");
  dropdown.innerHTML = `<div class="no-results">Error: ${error}</div>`;
  dropdown.style.display = "block";
}

function updateSelection() {
  const dropdown = document.getElementById("dropdown");
  const items = dropdown.querySelectorAll(".dropdown-item");
  
  items.forEach((item, index) => {
    item.classList.toggle("highlighted", index === selectedIndex);
  });
}

function selectCompany(match) {
  const symbol = match.symbol || "";
  const name = match.shortname || match.longname || "";
  
  selectedSymbol = symbol;
  document.getElementById("queryInput").value = `${symbol} - ${name}`;
  hideDropdown();
  
  document.getElementById("status").textContent = `Selected: ${symbol} - ${name}. Click "Get Financial Data" to proceed.`;
}

function showDropdown() {
  document.getElementById("dropdown").style.display = "block";
} 

function hideDropdown() {
  document.getElementById("dropdown").style.display = "none";
  selectedIndex = -1;
}

//GET DATA

async function onGetDataClick() {
  const statusElement = document.getElementById("status");
  const queryInput = document.getElementById("queryInput");
  
  let symbol = selectedSymbol || queryInput.value.trim().toUpperCase();
  
  if (!selectedSymbol && queryInput.value.includes(" - ")) {
    symbol = queryInput.value.split(" - ")[0].trim();
  }
  
  if (!symbol) {
    statusElement.textContent = "Please search and select a company first.";
    statusElement.style.color = "#d73a49";
    queryInput.focus();
    return;
  }
  
  statusElement.textContent = `Fetching data for ${symbol}...`;
  statusElement.style.color = "#555";

  try {
    console.log(`🚀 Starting data fetch for symbol: ${symbol}`);
    
    const [companyData, historicalData] = await Promise.all([
      getCompanyData(symbol),
      getHistoricalData(symbol)
    ]);

    if (!companyData) {
      throw new Error(`No financial data available for ${symbol}`);
    }

    console.log(`📊 Raw company data received:`, companyData);

    const assetProfile = companyData.assetProfile || {};
    const summaryDetail = companyData.summaryDetail || {};
    const financialData = companyData.financialData || {};
    const keyStats = companyData.defaultKeyStatistics || {};
    const priceData = companyData.price || {};
    const calendarEvents = companyData.calendarEvents || {};
    const holdersBreakdown = companyData.majorHoldersBreakdown || {};
    const upgradeDowngradeHistory = companyData.upgradeDowngradeHistory || {};

    console.log(`🏢 Asset Profile:`, assetProfile);
    console.log(`📈 Price Data:`, priceData);
    console.log(`📊 Financial Data:`, financialData);

    const currentYear = (keyStats.lastFiscalYearEnd?.fmt).split("-")[0]

    const price = sanitizeForExcel(priceData.regularMarketPrice?.raw || summaryDetail.previousClose?.raw || 0, 'number');
    const volume = sanitizeForExcel(summaryDetail.regularMarketVolume?.raw || summaryDetail.volume?.raw || 0, 'number');
    const latestTradeDay = sanitizeForExcel(priceData.regularMarketTime ? new Date(priceData.regularMarketTime * 1000).toLocaleDateString() : "", 'string');
    const change = sanitizeForExcel(priceData.regularMarketChange?.raw || 0, 'number');
    const changePercent = sanitizeForExcel(priceData.regularMarketChangePercent?.raw || 0, 'number');

    const companyName = sanitizeForExcel(assetProfile.longName || priceData.longName || priceData.shortName || symbol, 'string');
    const companySector = sanitizeForExcel(assetProfile.sector || "", 'string');
    const companyIndustry = sanitizeForExcel(assetProfile.industry || "", 'string');
    const currency = sanitizeForExcel(priceData.currency || "USD", 'string');
    const companyAddress = sanitizeForExcel(assetProfile.address1 ? `${assetProfile.address1}, ${assetProfile.city}, ${assetProfile.state}, ${assetProfile.country}, ${assetProfile.zip}` : "", 'string');
    const fiscalYearEnd = sanitizeForExcel(keyStats.lastFiscalYearEnd ? new Date(keyStats.lastFiscalYearEnd.raw * 1000).toLocaleDateString() : "", 'string');
    const companyMarketCap = sanitizeForExcel(priceData.marketCap?.raw || summaryDetail.marketCap?.raw || 0, 'number');
    const companyDividendYield = sanitizeForExcel(summaryDetail.dividendYield?.raw || 0, 'number');
    const companyDividendPerShare = sanitizeForExcel(summaryDetail.dividendRate?.raw || 0, 'number');
    const companySharesOutstanding = sanitizeForExcel(keyStats.sharesOutstanding?.raw || 0, 'number');
    const companyBeta = sanitizeForExcel(summaryDetail.beta?.raw || keyStats.beta?.raw || 1, 'number');
    const companyLatestQuarter = sanitizeForExcel(keyStats.mostRecentQuarter?.fmt || "", 'string');
    const companyPayoutRatio = sanitizeForExcel(summaryDetail.payoutRatio?.raw || 0, 'number');

    const analystForecasts = processAnalystForecasts(upgradeDowngradeHistory, price);

    Excel.run(async (context) => {
      console.log(`📝 Starting Excel operations for ${symbol}`);
      
      try {
        const workbook = context.workbook;
        
        //STEP 1: CREATE FINANCIALS SHEET
        if (historicalData) {
          console.log(`📅 Historical data found:`, historicalData);
          try {
            console.log(`📈 Creating comprehensive financials sheets (Income Statement, Balance Sheet, Cash Flow)`);
            await createFinancialsSheet(context, symbol, companyName || symbol, historicalData, currentYear);
            console.log(`✅ Financials sheets created`);
          } catch (e) {
            console.error(`❌ Financials sheets failed:`, e);
          }
        }

        //STEP 2: CREATE DCF VALUATION SHEET
        try {
          console.log(`📊 Creating DCF Valuation sheet`);
          await createDCFSheet(context, symbol, companyName || symbol, currentYear);
          console.log(`✅ DCF Valuation sheet created`);
        } catch (e) {
          console.error(`❌ DCF Valuation sheet failed:`, e);
        }

        //STEP 3: CREATE ANALYST FORECASTS SHEET
        try {
          console.log(`📊 Creating Analyst Forecasts sheet`);
          await createAnalystForecastsSheet(context, symbol, companyName || symbol, analystForecasts, price);
          console.log(`✅ Analyst Forecasts sheet created`);
        } catch (e) {
          console.error(`❌ Analyst Forecasts sheet failed:`, e);
        }

        //STEP 4: CREATE COMPANY OVERVIEW SHEET
        let overviewSheet = workbook.worksheets.getItemOrNullObject("Company Overview");
        overviewSheet.load("name");
        await context.sync();

        if (overviewSheet.isNullObject) {
          overviewSheet = workbook.worksheets.getItem("Sheet1");
          overviewSheet.name = "Company Overview";
        } else {
          const used = overviewSheet.getUsedRange(false);
          used.clear();
        }

        try {
          console.log(`🧹 Clearing existing content`);
          const usedRange = overviewSheet.getUsedRange(false);
          if (usedRange) {
            usedRange.clear();
          }
        } catch (e) {
          console.log(`ℹ️ No existing content to clear`);
        }

        try {
          console.log(`📋 Creating overview header ${companyName} (${symbol})`);
          const headerRange = overviewSheet.getRange("A1:E1");
          headerRange.merge();
          headerRange.values = [[sanitizeForExcel(`${companyName} (${symbol}) - Company Overview`, 'string'), "", "", "", ""]];
          headerRange.format.font.bold = true;
          headerRange.format.font.size = 14;
          headerRange.format.verticalAlignment = "Center";
          headerRange.format.horizontalAlignment = "Center";
          headerRange.format.fill.color = "#1F4E78";
          headerRange.format.font.color = "white";
          await context.sync();
        } catch (e) {
          console.error(`❌ Overview header failed:`, e);
          throw e;
        }

        try {
          const basicInfoStartRow = 3;
          overviewSheet.getRange(`A${basicInfoStartRow}:B${basicInfoStartRow}`).values = [["COMPANY INFORMATION", ""]];
          overviewSheet.getRange(`A${basicInfoStartRow}:B${basicInfoStartRow}`).format.font.bold = true;
          overviewSheet.getRange(`A${basicInfoStartRow}:B${basicInfoStartRow}`).format.fill.color = "#D9E2F3";

          const basicInfoData = validateDataArray([
            ["Symbol", symbol],
            ["Company Name", companyName],
            ["Sector", companySector],
            ["Industry", companyIndustry],
            ["Country", sanitizeForExcel(assetProfile.country || "", 'string')],
            ["Currency", currency],
            ["Fiscal Year End", fiscalYearEnd],
            ["Latest Quarter", companyLatestQuarter],
            ["Employees", sanitizeForExcel(assetProfile.fullTimeEmployees || "", 'string')],
            ["Website", sanitizeForExcel(assetProfile.website || "", 'string')]
          ], 'basicInfoData');

          const basicInfoRange = overviewSheet.getRangeByIndexes(basicInfoStartRow + 1, 0, basicInfoData.length, 2);
          basicInfoRange.values = basicInfoData;
          basicInfoRange.getColumn(0).format.font.bold = true;
          await context.sync();
        } catch (e) {
          console.error(`❌ Basic info failed:`, e);
          throw e;
        }

        try {
          const tradingStartRow = 16;
          overviewSheet.getRange(`A${tradingStartRow}:B${tradingStartRow}`).values = [["STOCK PRICE & TRADING", ""]];
          overviewSheet.getRange(`A${tradingStartRow}:B${tradingStartRow}`).format.font.bold = true;
          overviewSheet.getRange(`A${tradingStartRow}:B${tradingStartRow}`).format.fill.color = "#D9E2F3";

          const tradingData = validateDataArray([
            ["Current Price", price],
            ["Previous Close", sanitizeForExcel(summaryDetail.previousClose?.raw || 0, 'number')],
            ["Change", change],
            ["Change %", changePercent],
            ["Volume", volume],
            ["52 Week High", sanitizeForExcel(summaryDetail.fiftyTwoWeekHigh?.raw || 0, 'number')],
            ["52 Week Low", sanitizeForExcel(summaryDetail.fiftyTwoWeekLow?.raw || 0, 'number')],
            ["50 Day Average", sanitizeForExcel(summaryDetail.fiftyDayAverage?.raw || 0, 'number')],
            ["200 Day Average", sanitizeForExcel(summaryDetail.twoHundredDayAverage?.raw || 0, 'number')]
          ], 'tradingData');

          const tradingRange = overviewSheet.getRangeByIndexes(tradingStartRow + 1, 0, tradingData.length, 2);
          tradingRange.values = tradingData;
          tradingRange.getColumn(0).format.font.bold = true;
          await context.sync();
        } catch (e) {
          console.error(`❌ Trading section failed:`, e);
          throw e;
        }

        try {
          overviewSheet.getRange("D3:E3").values = [["ASSUMPTIONS", ""]];
          overviewSheet.getRange("D3:E3").format.font.bold = true;
          overviewSheet.getRange("D3:E3").format.fill.color = "#D9E2F3";

          overviewSheet.getRange("D4:E4").values = [["General Information", ""]];
          overviewSheet.getRange("D4:E4").format.font.bold = true;
          overviewSheet.getRange("D4:E4").format.fill.color = "#E7E6E6";

          overviewSheet.getRange("D5:E8").values = [
            ["Market Capitalization", companyMarketCap],
            ["Shares Outstanding", companySharesOutstanding],
            ["Dividend Paid Annually", companyDividendPerShare],
            ["Payout Ratio", sanitizeForExcel(companyPayoutRatio, 'number')]
          ];
          overviewSheet.getRange("D5:D8").format.font.bold = true;
          overviewSheet.getRange("E5").numberFormat = "#,###"; // market cap
          overviewSheet.getRange("E6").numberFormat = "#,###"; // Shares Outstanding
          overviewSheet.getRange("E7").numberFormat = "0.000"; // Dividend Paid
          overviewSheet.getRange("E8").numberFormat = "0.00%"; // Payout Ratio

          overviewSheet.getRange("D10:E10").values = [["Cost of Equity", ""]];
          overviewSheet.getRange("D10:E10").format.font.bold = true;
          overviewSheet.getRange("D10:E10").format.fill.color = "#E7E6E6";

          const riskFreeRate = 0.04;
          const marketReturn = 0.12;

          overviewSheet.getRange("D11:E15").values = [
            ["Risk-Free Rate", riskFreeRate],
            ["Market Return", marketReturn],
            ["Market Risk Premium", "=E12-E11"],
            ["Beta", companyBeta],
            ["Cost of Equity", "=E11 + (E14 * E13)"]
          ];
          overviewSheet.getRange("D11:D15").format.font.bold = true;
          overviewSheet.getRange("E11").numberFormat = "0.00%"; // Risk-Free Rate
          overviewSheet.getRange("E12").numberFormat = "0.00%"; // Market Return
          overviewSheet.getRange("E13").numberFormat = "0.00%"; // Market Risk Premium
          overviewSheet.getRange("E14").numberFormat = "0.000"; // Beta
          overviewSheet.getRange("E15").numberFormat = "0.00%"; // Cost of Equity

          overviewSheet.getRange("D17:E17").values = [["Cost of Debt", ""]];
          overviewSheet.getRange("D17:E17").format.font.bold = true;
          overviewSheet.getRange("D17:E17").format.fill.color = "#E7E6E6";

          overviewSheet.getRange("D18:D20").values = [["Net Debt/EBITDA"],["Implied Credit Spread"],["Cost of Debt"]];
          overviewSheet.getRange("D18:D20").format.font.bold = true;

          overviewSheet.getRange("E18").formulasLocal = [[`=(INDEX(All_Financials_${symbol}!B:B;MATCH("Total Debt";All_Financials_${symbol}!A:A;0))-INDEX(All_Financials_${symbol}!B:B;MATCH("Cash And Cash Equivalents";All_Financials_${symbol}!A:A;0)))/INDEX(All_Financials_${symbol}!B:B;MATCH("E B I T D A";All_Financials_${symbol}!A:A;0))`]];

          overviewSheet.getRange("E19").formulasLocal = [[`=IF(E18<=0;0,003;IF(E18<=0,5;0,005;IF(E18<=1;0,008;IF(E18<=1,5;0,011;IF(E18<=2;0,015;IF(E18<=2,5;0,020;IF(E18<=3;0,028;IF(E18<=4;0,038;IF(E18<=5;0,050;0,070)))))))))`]];

          overviewSheet.getRange("E20").formulasLocal = [["=E11+E19"]];

          overviewSheet.getRange("E18").numberFormat = [["0.00"]];
          overviewSheet.getRange("E19").numberFormat = [["0.00%"]];
          overviewSheet.getRange("E20").numberFormat = [["0.00%"]];

          overviewSheet.getRange("D22:E22").values = [["Weighted Average Cost of Capital (WACC)", ""]];
          overviewSheet.getRange("D22:E22").format.font.bold = true;
          overviewSheet.getRange("D22:E22").format.fill.color = "#E7E6E6";

          overviewSheet.getRange("D23:D26").values = [["Debt Portion"], ["Equity Portion"], ["Tax rate"], ["WACC"]];
          overviewSheet.getRange("D23:D26").format.font.bold = true;

          overviewSheet.getRange("E23").formulasLocal = [[`=INDEX(All_Financials_${symbol}!B:B;MATCH("Total Debt";All_Financials_${symbol}!A:A;0))/(INDEX(All_Financials_${symbol}!B:B;MATCH("Total Debt";All_Financials_${symbol}!A:A;0))+INDEX(All_Financials_${symbol}!B:B;MATCH("Stockholders Equity";All_Financials_${symbol}!A:A;0)))`]];
          overviewSheet.getRange("E24").formulasLocal = [["=1-E23"]];
          overviewSheet.getRange("E25").formulasLocal = [[`=AVERAGE(OFFSET(All_Financials_${symbol}!$A$1;MATCH("Tax Rate For Calcs";All_Financials_${symbol}!$A:$A;0)-1;1;1;4))`]];
          overviewSheet.getRange("E26").formulasLocal = [["=(E23*E20*(1-E25))+(E24*E15)"]];

          overviewSheet.getRange("E23:E26").numberFormat = [["0.00%"], ["0.00%"], ["0.00%"], ["0.00%"]];

          await context.sync();

          await context.sync();
        } catch (e) {
          console.error(`❌ Assumptions section failed:`, e);
          throw e;
        }

        try {
          console.log(`📊 Creating Analyst Forecasts sheet`);
          await createAnalystForecastsSheet(context, symbol, companyName || symbol, analystForecasts, price);
          console.log(`✅ Analyst Forecasts sheet created`);
        } catch (e) {
          console.error(`❌ Analyst Forecasts sheet failed:`, e);
        }

        //STEP 4: CREATE DCF VALUATION SHEET
        try {
          console.log(`📊 Creating DCF Valuation sheet`);
          await createDCFSheet(context, symbol, companyName || symbol, currentYear);
          console.log(`✅ DCF Valuation sheet created`);
        } catch (e) {
          console.error(`❌ DCF Valuation sheet failed:`, e);
        }

        await context.sync();
        console.log(`🎉 Excel operations completed successfully for ${symbol}`);

      } catch (excelError) {
        console.error(`💥 Excel operation failed:`, excelError);
        throw excelError;
      }
    });

    statusElement.textContent = `✅ Complete analysis created for "${symbol}". ${historicalData ? 'Historical data included.' : 'Current data only.'}`;
    statusElement.style.color = "#28a745";


  } catch (err) {
    console.error('🚨 Data fetch error:', err);
    statusElement.textContent = `❌ Error: ${err.message}`;
    statusElement.style.color = "#d73a49";
  }
}

//ANALYST FORECASTS

function processAnalystForecasts(upgradeDowngradeHistory, currentPrice) {
  console.log('📊 Processing analyst forecasts data');
  
  if (!upgradeDowngradeHistory || !upgradeDowngradeHistory.history || !Array.isArray(upgradeDowngradeHistory.history)) {
    console.log('⚠️ No analyst forecasts data available');
    return { validForecasts: [], summaryStats: {} };
  }

  const history = upgradeDowngradeHistory.history;
  const oneYearAgo = Date.now() - (180 * 24 * 60 * 60 * 1000); //180 days
  
  console.log(`📅 Filtering ${history.length} analyst records from last 180 days`);

  const firmForecasts = {};
  
  for (const entry of history) {
    const { epochGradeDate, firm, currentPriceTarget, priorPriceTarget } = entry;
    
    if (!epochGradeDate || !firm || !currentPriceTarget) continue;
    
    if (currentPriceTarget === null || currentPriceTarget === priorPriceTarget) continue;
    
    const entryDate = epochGradeDate * 1000;
    
    if (entryDate < oneYearAgo) continue;
    
    if (!firmForecasts[firm] || entryDate > firmForecasts[firm].epochGradeDate * 1000) {
      firmForecasts[firm] = entry;
    }
  }
  
  console.log(`🏢 Found ${Object.keys(firmForecasts).length} unique firms with valid forecasts`);
  
  const validForecasts = [];
  const impliedReturns = [];
  
  for (const [firm, forecast] of Object.entries(firmForecasts)) {
    const impliedReturn = (forecast.currentPriceTarget / currentPrice) - 1;
    
    const cappedReturn = Math.max(-0.5, Math.min(1.0, impliedReturn));
    
    validForecasts.push({
      firm,
      date: new Date(forecast.epochGradeDate * 1000).toLocaleDateString(),
      fromGrade: forecast.fromGrade || '',
      toGrade: forecast.toGrade || '',
      currentPriceTarget: forecast.currentPriceTarget,
      impliedReturn: cappedReturn,
      wasCapped: impliedReturn !== cappedReturn
    });
    
    impliedReturns.push(cappedReturn);
  }
  
  let summaryStats = {};
  
  if (impliedReturns.length > 0) {
    const sortedReturns = [...impliedReturns].sort((a, b) => a - b);
    const sum = impliedReturns.reduce((acc, val) => acc + val, 0);
    const mean = sum / impliedReturns.length;
    
    const squaredDiffs = impliedReturns.map(val => Math.pow(val - mean, 2));
    const avgSquaredDiff = squaredDiffs.reduce((acc, val) => acc + val, 0) / impliedReturns.length;
    const standardDev = Math.sqrt(avgSquaredDiff);
    
    summaryStats = {
      average: mean,
      median: sortedReturns[Math.floor(sortedReturns.length / 2)],
      count: impliedReturns.length,
      min: sortedReturns[0],
      max: sortedReturns[sortedReturns.length - 1],
      standardDeviation: standardDev
    };
  }
  
  console.log(`📈 Processed ${validForecasts.length} valid forecasts with average implied return: ${(summaryStats.average * 100).toFixed(1)}%`);
  
  return { validForecasts, summaryStats };
}

async function createAnalystForecastsSheet(context, symbol, companyName, processedForecasts, currentPrice) {
  console.log(`📊 Creating analyst forecasts sheet for: ${symbol}`);

  const sheetName = `Analyst_Forecasts_${symbol}`;
  let sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
  sheet.load("name");
  await context.sync();

  if (sheet.isNullObject) {
    sheet = context.workbook.worksheets.add(sheetName);
  } else {
    try {
      const used = sheet.getUsedRange(false);
      used.load("address");
      await context.sync();
      used.clear();
    } catch (e) {
      console.log(`ℹ️ No existing content to clear in ${sheetName}`);
    }
  }

  const { validForecasts, summaryStats } = processedForecasts;

  const headerRange = sheet.getRange("A1:E1");
  headerRange.merge();
  headerRange.values = [[`${companyName} (${symbol}) - Analyst Forecasts`, "", "", "", ""]];
  headerRange.format.font.bold = true;
  headerRange.format.font.size = 14;
  headerRange.format.horizontalAlignment = "Center";
  headerRange.format.fill.color = "#1F4E78";
  headerRange.format.font.color = "white";
  await context.sync();

  if (validForecasts.length === 0) {
    sheet.getRange("A3").values = [["No recent analyst price targets available"]];
    sheet.getRange("A3").format.font.bold = true;
    sheet.getRange("A3").format.font.size = 12;
    return;
  }

  sheet.getRange("A3:E3").values = [["Firm", "Date", "From→To Grade", "Price Target", "Implied Return"]];
  sheet.getRange("A3:E3").format.font.bold = true;
  sheet.getRange("A3:E3").format.fill.color = "#D9E2F3";
  sheet.getRange("A3:E3").format.horizontalAlignment = "Center";

  let currentRow = 4;
  for (const forecast of validForecasts) {
    const gradeChange = forecast.fromGrade === forecast.toGrade 
      ? forecast.toGrade 
      : `${forecast.fromGrade}→${forecast.toGrade}`;
    
    const dataRange = sheet.getRange(`A${currentRow}:E${currentRow}`);
    dataRange.values = [[
      forecast.firm,
      forecast.date,
      gradeChange,
      forecast.currentPriceTarget,
      forecast.impliedReturn
    ]];
    
    sheet.getRange(`D${currentRow}`).numberFormat = "#,##0.00";
    sheet.getRange(`E${currentRow}`).numberFormat = "0.00%";
    
    if (forecast.wasCapped) {
      sheet.getRange(`E${currentRow}`).format.fill.color = "#FFE6E6";
    }
    
    currentRow++;
  }

  currentRow += 2;
  sheet.getRange(`A${currentRow}`).values = [["SUMMARY STATISTICS"]];
  sheet.getRange(`A${currentRow}:B${currentRow}`).format.font.bold = true;
  sheet.getRange(`A${currentRow}:B${currentRow}`).format.fill.color = "#D9E2F3";
  currentRow++;

  const summaryData = [
    ["Average Implied Return", summaryStats.average || 0],
    ["Median Implied Return", summaryStats.median || 0],
    ["Number of Forecasts", summaryStats.count || 0],
    ["Range (Min to Max)", `${((summaryStats.min || 0) * 100).toFixed(1)}% to ${((summaryStats.max || 0) * 100).toFixed(1)}%`],
    ["Standard Deviation", summaryStats.standardDeviation || 0],
    ["", ""],
    ["Short-term growth rate", (summaryStats.average - 0.5*summaryStats.standardDeviation) || 0],
    ["Current Stock Price", currentPrice]
  ];

  for (let i = 0; i < summaryData.length; i++) {
    const summaryRange = sheet.getRange(`A${currentRow + i}:B${currentRow + i}`);
    summaryRange.values = [summaryData[i]];
    
    if (summaryData[i][0] !== "") {
      sheet.getRange(`A${currentRow + i}`).format.font.bold = true;
    }
    
    if (i === 0 || i === 1) { // Average and Median
      sheet.getRange(`B${currentRow + i}`).numberFormat = "0.00%";
    } else if (i === 4) { // Standard Deviation
      sheet.getRange(`B${currentRow + i}`).numberFormat = "0.00%";
    } else if (i === 6) { // Short-term growth rate
      sheet.getRange(`B${currentRow + i}`).numberFormat = "0.00%";
    } else if (i === 7) { // Current Stock Price
      sheet.getRange(`B${currentRow + i}`).numberFormat = "#,##0.00";
    }
  }

  try {
    sheet.getUsedRange().format.autofitColumns();
    await context.sync();
  } catch (e) {
    console.error(`Formatting failed for analyst forecasts sheet:`, e);
  }

  console.log(`🎉 Analyst forecasts sheet completed with ${validForecasts.length} forecasts`);
}

//ALL FINANCIALS SHEET

async function createFinancialsSheet(context, symbol, companyName, timeseriesData, currentYear) {
  if (!timeseriesData || typeof timeseriesData !== 'object') {
    console.log('No timeseries data available for financials sheet');
    return;
  }

  const result = timeseriesData.result || timeseriesData;
  if (!result || !Array.isArray(result)) {
    console.log('Timeseries data structure unexpected; no array found');
    return;
  }
  const timeseriesArray = result;

  const incomeVariants = Object.keys(VARIANT_TO_STATEMENT).filter(k => VARIANT_TO_STATEMENT[k] === 'Income Statement');
  const balanceVariants = Object.keys(VARIANT_TO_STATEMENT).filter(k => VARIANT_TO_STATEMENT[k] === 'Balance Sheet');
  const cashflowVariants = Object.keys(VARIANT_TO_STATEMENT).filter(k => VARIANT_TO_STATEMENT[k] === 'Cash Flow');

  //await createStatementSheet(context, symbol, companyName, timeseriesArray, 'Income Statement', `Income_Statement_${symbol}`, incomeVariants);
  //await createStatementSheet(context, symbol, companyName, timeseriesArray, 'Balance Sheet', `Balance_Sheet_${symbol}`, balanceVariants);
  //await createStatementSheet(context, symbol, companyName, timeseriesArray, 'Cash Flow', `Cash_Flow_${symbol}`, cashflowVariants);
  
  await createConsolidatedFinancialsSheet(context, symbol, companyName, timeseriesArray, incomeVariants, balanceVariants, cashflowVariants, currentYear);

}

async function createStatementSheet(context, symbol, companyName, timeseriesArray, statementType, sheetName, variantsList) {
  console.log(`📊 Creating ${statementType} sheet: ${sheetName}`);

  let sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
  sheet.load("name");
  await context.sync();

  if (sheet.isNullObject) {
    sheet = context.workbook.worksheets.add(sheetName);
  } else {
    try {
      const used = sheet.getUsedRange(false);
      used.load("address");
      await context.sync();
      used.clear();
    } catch (e) {
      // ignore
    }
  }

  const headerRange = sheet.getRange("A1:I1");
  headerRange.merge();
  headerRange.values = [[`${companyName} (${symbol}) - ${statementType} (Annual)`, "", "", "", "", "", "", "", ""]];
  headerRange.format.font.bold = true;
  headerRange.format.font.size = 14;
  headerRange.format.horizontalAlignment = "Center";
  headerRange.format.fill.color = "#1F4E78";
  headerRange.format.font.color = "white";
  await context.sync();

  const sortedYears = [2025, 2024, 2023, 2022];

  const processedData = [];

  for (const variant of variantsList) {
    const matchObj = timeseriesArray.find(m => {
      if (Array.isArray(m.meta?.type) && m.meta.type[0] === variant) return true;
      return !!m[variant];
    });

    const variable = {
      variant,
      name: formatVariableName(variant),
      values: {}
    };

    if (matchObj) {
      const dataPoints = matchObj[variant] || (() => {
        const t = Array.isArray(matchObj.meta?.type) ? matchObj.meta.type[0] : null;
        return t && matchObj[t] ? matchObj[t] : null;
      })();

      if (Array.isArray(dataPoints)) {
        for (const pt of dataPoints) {
          if (!pt) continue;
          
          const asOf = pt.asOfDate || pt.endDate || pt.timestamp || pt.date;
          if (!asOf) continue;
          
          let year = null;
          if (typeof asOf === 'string') {
            year = parseInt(asOf.substring(0, 4), 10);
          } else if (typeof asOf === 'number') {
            year = new Date(asOf * 1000).getFullYear();
          }
          
          if (!year) continue;
          
          let raw = null;
          if (pt.reportedValue && pt.reportedValue.raw !== undefined && pt.reportedValue.raw !== null) {
            raw = pt.reportedValue.raw;
          } else if (pt.value && pt.value.raw !== undefined && pt.value.raw !== null) {
            raw = pt.value.raw;
          } else if (typeof pt.raw === 'number') {
            raw = pt.raw;
          }
          
          if (raw !== null && raw !== undefined) {
            variable.values[year] = sanitizeForExcel(raw, 'number');
            console.log(`✅ ${variant} [${year}]: ${raw}`);
          }
        }
      }
    }

    processedData.push(variable);
  }

  const filteredData = processedData.filter(variable => {
    const hasAnyValue = sortedYears.some(year => variable.values[year] !== undefined);
    if (!hasAnyValue) {
      console.log(`⏭️ Skipping ${variable.variant} - no data for any year`);
    }
    return hasAnyValue;
  });

  console.log(`📅 ${statementType} - Using fixed years: ${sortedYears.join(', ')}`);
  console.log(`📊 ${statementType} - ${filteredData.length} of ${processedData.length} variants have data`);

  const headerRow = ["Financial Metric", ...sortedYears];
  const headerColEnd = String.fromCharCode(65 + sortedYears.length);
  sheet.getRange(`A3:${headerColEnd}3`).values = [headerRow];
  sheet.getRange(`A3:${headerColEnd}3`).format.font.bold = true;
  sheet.getRange(`A3:${headerColEnd}3`).format.fill.color = "#D9E2F3";
  sheet.getRange(`A3:${headerColEnd}3`).format.horizontalAlignment = "Center";
  await context.sync();

  const variantToData = {};
  for (const variable of filteredData) {
    variantToData[variable.variant] = variable;
  }

  const sections = STATEMENT_SECTIONS[statementType] || [];

  let currentRow = 4;
  for (const section of sections) {
    const sectionVariants = section.variants.filter(v => variantToData[v]);
    
    if (sectionVariants.length === 0) {
      console.log(`⏭️ Skipping section "${section.title}" - no data`);
      continue;
    }

    const sectionHeaderRange = sheet.getRange(`A${currentRow}:${headerColEnd}${currentRow}`);
    sectionHeaderRange.values = [[section.title, "", "", "", ""]];
    sectionHeaderRange.format.font.bold = true;
    sectionHeaderRange.format.fill.color = "#B4C7E7";
    sectionHeaderRange.format.font.color = "#1F4E78";
    currentRow++;

    for (const variant of sectionVariants) {
      const variable = variantToData[variant];
      const row = [variable.name];
      for (const y of sortedYears) {
        const cellValue = variable.values[y];
        row.push(cellValue !== undefined ? cellValue : "");
      }
      const dataRange = sheet.getRange(`A${currentRow}:${headerColEnd}${currentRow}`);
      const validated = validateDataArray([row], `${statementType}-${variable.name}`);
      dataRange.values = validated;
      currentRow++;
    }
    
    currentRow++; 
  }
  await context.sync();

  try {
    const dataStartCol = String.fromCharCode(66);
    const dataEndCol = headerColEnd;
    const dataRange = sheet.getRange(`${dataStartCol}4:${dataEndCol}${currentRow - 1}`);
    dataRange.numberFormat = [["#,##0.00"]];
    sheet.getUsedRange().format.autofitColumns();
    await context.sync();
  } catch (e) {
    console.error(`Formatting failed for ${statementType}:`, e);
  }

  console.log(`🎉 ${statementType} sheet completed with section headers and ${filteredData.length} metrics across 4 fixed years`);
}

async function createConsolidatedFinancialsSheet(context, symbol, companyName, timeseriesArray, incomeVariants, balanceVariants, cashflowVariants, currentYear) {
  console.log(`📊 Creating consolidated financials sheet for: ${symbol}`);

  const sheetName = `All_Financials_${symbol}`;
  let sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
  sheet.load("name");
  await context.sync();

  if (sheet.isNullObject) {
    sheet = context.workbook.worksheets.add(sheetName);
  } else {
    try {
      const used = sheet.getUsedRange(false);
      used.load("address");
      await context.sync();
      used.clear();
    } catch (e) {
      // ignore
    }
  }

  const headerRange = sheet.getRange("A1:E1");
  headerRange.merge();
  headerRange.values = [[`${companyName} (${symbol}) - All Financial Statements (Annual)`, "", "", "", ""]];
  headerRange.format.font.bold = true;
  headerRange.format.font.size = 14;
  headerRange.format.horizontalAlignment = "Center";
  headerRange.format.fill.color = "#1F4E78";
  headerRange.format.font.color = "white";
  await context.sync();

  const sortedYears = [currentYear, currentYear - 1, currentYear - 2, currentYear - 3];
  console.log(`📅 Using dynamic years for ${symbol}: ${sortedYears.join(', ')}`);
  const headerColEnd = String.fromCharCode(65 + sortedYears.length);

  const statementConfigs = [
    { type: 'Income Statement', variants: incomeVariants },
    { type: 'Balance Sheet', variants: balanceVariants },
    { type: 'Cash Flow', variants: cashflowVariants }
  ];

  let currentRow = 3;

  for (const config of statementConfigs) {
    const statementType = config.type;
    const variantsList = config.variants;

    console.log(`📝 Processing ${statementType} for consolidated sheet`);

    const processedData = [];

    for (const variant of variantsList) {
      const matchObj = timeseriesArray.find(m => {
        if (Array.isArray(m.meta?.type) && m.meta.type[0] === variant) return true;
        return !!m[variant];
      });

      const variable = {
        variant,
        name: formatVariableName(variant),
        values: {}
      };

      if (matchObj) {
        const dataPoints = matchObj[variant] || (() => {
          const t = Array.isArray(matchObj.meta?.type) ? matchObj.meta.type[0] : null;
          return t && matchObj[t] ? matchObj[t] : null;
        })();

        if (Array.isArray(dataPoints)) {
          for (const pt of dataPoints) {
            if (!pt) continue;
            
            const asOf = pt.asOfDate || pt.endDate || pt.timestamp || pt.date;
            if (!asOf) continue;
            
            let year = null;
            if (typeof asOf === 'string') {
              year = parseInt(asOf.substring(0, 4), 10);
            } else if (typeof asOf === 'number') {
              year = new Date(asOf * 1000).getFullYear();
            }
            
            if (!year) continue;
            
            let raw = null;
            if (pt.reportedValue && pt.reportedValue.raw !== undefined && pt.reportedValue.raw !== null) {
              raw = pt.reportedValue.raw;
            } else if (pt.value && pt.value.raw !== undefined && pt.value.raw !== null) {
              raw = pt.value.raw;
            } else if (typeof pt.raw === 'number') {
              raw = pt.raw;
            }
            
            if (raw !== null && raw !== undefined) {
              variable.values[year] = sanitizeForExcel(raw, 'number');
            }
          }
        }
      }

      processedData.push(variable);
    }

    const filteredData = processedData.filter(variable => {
      return sortedYears.some(year => variable.values[year] !== undefined);
    });

    console.log(`📊 ${statementType} - ${filteredData.length} variants with data`);

    const statementHeaderRange = sheet.getRange(`A${currentRow}:${headerColEnd}${currentRow}`);
    statementHeaderRange.values = [[statementType, "", "", "", ""]];
    statementHeaderRange.format.font.bold = true;
    statementHeaderRange.format.font.size = 12;
    statementHeaderRange.format.fill.color = "#4472C4";
    statementHeaderRange.format.font.color = "white";
    currentRow++;

    const headerRow = ["Financial Metric", ...sortedYears.map(y => `'${y}`)];
    const colHeaderRange = sheet.getRange(`A${currentRow}:${headerColEnd}${currentRow}`);
    colHeaderRange.values = [headerRow];
    colHeaderRange.format.font.bold = true;
    colHeaderRange.format.fill.color = "#D9E2F3";
    colHeaderRange.format.horizontalAlignment = "Center";
    currentRow++;

    const variantToData = {};
    for (const variable of filteredData) {
      variantToData[variable.variant] = variable;
    }

    const sections = STATEMENT_SECTIONS[statementType] || [];

    for (const section of sections) {
      const sectionVariants = section.variants.filter(v => variantToData[v]);
      
      if (sectionVariants.length === 0) {
        continue;
      }

      const sectionHeaderRange = sheet.getRange(`A${currentRow}:${headerColEnd}${currentRow}`);
      sectionHeaderRange.values = [[section.title, "", "", "", ""]];
      sectionHeaderRange.format.font.bold = true;
      sectionHeaderRange.format.fill.color = "#B4C7E7";
      sectionHeaderRange.format.font.color = "#1F4E78";
      currentRow++;

      for (const variant of sectionVariants) {
        const variable = variantToData[variant];
        const row = [variable.name];
        for (const y of sortedYears) {
          const cellValue = variable.values[y];
          row.push(cellValue !== undefined ? cellValue : "");
        }
        const dataRange = sheet.getRange(`A${currentRow}:${headerColEnd}${currentRow}`);
        const validated = validateDataArray([row], `${statementType}-${variable.name}`);
        dataRange.values = validated;
        currentRow++;
      }

    }

    currentRow++;
  }

  await context.sync();

  try {
    const dataStartCol = String.fromCharCode(66);
    const dataEndCol = headerColEnd;
    const dataRange = sheet.getRange(`${dataStartCol}4:${dataEndCol}${currentRow - 1}`);
    dataRange.numberFormat = [["#,##0.00"]];

    sheet.getRange("A:A").columnWidth = 37;
    sheet.getUsedRange().format.autofitColumns();
    await context.sync();
  } catch (e) {
    console.error(`Formatting failed for consolidated sheet:`, e);
  }

  console.log(`🎉 Consolidated financials sheet completed with all statements`);
}

//DCF VALUATION SHEET

async function createDCFSheet(context, symbol, companyName, currentYear) {
  console.log(`📊 Creating DCF Valuation sheet for: ${symbol}`);

  const sheetName = `DCF_Valuation_${symbol}`;
  let sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
  sheet.load("name");
  await context.sync();

  if (sheet.isNullObject) {
    sheet = context.workbook.worksheets.add(sheetName);
  } else {
    try {
      const used = sheet.getUsedRange(false);
      used.load("address");
      await context.sync();
      used.clear();
    } catch (e) {
      // ignore
    }
  }

  try {
    sheet.getRange("A1:B1").values = [["ASSUMPTIONS", ""]];
    sheet.getRange("A1:B1").format.font.bold = true;
    sheet.getRange("A1:B1").format.fill.color = "#D9E2F3";

    sheet.getRange("A3:B3").values = [["General Information", ""]];
    sheet.getRange("A3:B3").format.font.bold = true;
    sheet.getRange("A3:B3").format.fill.color = "#E7E6E6";

    sheet.getRange("A4:B7").values = [
      ["Market Capitalization", `='Company Overview'!E5`],
      ["Shares Outstanding", `='Company Overview'!E6`],
      ["Dividend Paid Annually", `='Company Overview'!E7`],
      ["Payout Ratio", `='Company Overview'!E8`]
    ];
    sheet.getRange("A4:A7").format.font.bold = true;
    sheet.getRange("B4").numberFormat = "#,###";
    sheet.getRange("B5").numberFormat = "#,###";
    sheet.getRange("B6").numberFormat = "0.00";
    sheet.getRange("B7").numberFormat = "0.000%";

    sheet.getRange("A10:B10").values = [["Cost of Equity", ""]];
    sheet.getRange("A10:B10").format.font.bold = true;
    sheet.getRange("A10:B10").format.fill.color = "#E7E6E6";

    sheet.getRange("A11:B15").values = [
      ["Risk-Free Rate", `='Company Overview'!E11`],
      ["Market Return", `='Company Overview'!E12`],
      ["Market Risk Premium", `='Company Overview'!E13`],
      ["Beta", `='Company Overview'!E14`],
      ["Cost of Equity", `='Company Overview'!E15`]
    ];
    sheet.getRange("A11:A15").format.font.bold = true;
    sheet.getRange("B11").numberFormat = "0.00%";
    sheet.getRange("B12").numberFormat = "0.00%";
    sheet.getRange("B13").numberFormat = "0.00%";
    sheet.getRange("B14").numberFormat = "0.000";
    sheet.getRange("B15").numberFormat = "0.00%";


    sheet.getRange("A17:B17").values = [["Cost of Debt & WACC", ""]];
    sheet.getRange("A17:B17").format.font.bold = true;
    sheet.getRange("A17:B17").format.fill.color = "#E7E6E6";

    sheet.getRange("A18:B22").values = [
      ["Cost of Debt", `='Company Overview'!E20`],
      ["Debt Portion", `='Company Overview'!E23`],
      ["Equity Portion", `='Company Overview'!E24`],
      ["Tax Rate", `='Company Overview'!E25`],
      ["WACC", `='Company Overview'!E26`]
    ];
    sheet.getRange("A18:A22").format.font.bold = true;
    sheet.getRange("B18:B22").numberFormat = "0.00%";

    await context.sync();
  } catch (e) {
    console.error(`❌ Assumptions section failed:`, e);
  }

  try {
    sheet.getRange("D1:E1").values = [["ADDITIONAL ASSUMPTIONS", ""]];
    sheet.getRange("D1:E1").format.font.bold = true;
    sheet.getRange("D1:E1").format.fill.color = "#D9E2F3";

    sheet.getRange("D3:E3").values = [["Short-term Growth Rate", 0.08]]; 
    //sheet.getRange("E3:E3").formulasLocal = [["=INDEX(Analyst_Forecasts_${symbol}!B:B,MATCH(\"Short-term growth rate\",Analyst_Forecasts_${symbol}!A:A,0))"]]; //=INDEX(Analyst_Forecasts_${symbol}!B:B,MATCH(\"Short-term growth rate\",Analyst_Forecasts_${symbol}!A:A,0))
    sheet.getRange("D3:D3").format.font.bold = true;
    sheet.getRange("E3").numberFormat = "0.00%";
    sheet.getRange("E3").format.fill.color = "#E7F0F7";

    sheet.getRange("D4:E4").values = [["Long-term Growth Rate", 0.04]];
    sheet.getRange("D4:D4").format.font.bold = true;
    sheet.getRange("E4").numberFormat = "0.00%";
    sheet.getRange("E4").format.fill.color = "#E7F0F7";

    sheet.getRange("D5:E5").values = [["Stock Price", `='Company Overview'!B18`]];
    sheet.getRange("D5:D5").format.font.bold = true;
    sheet.getRange("E5").numberFormat = "#,##0.00";

    sheet.getRange("D6:E6").values = [["Stock Price Date", `=TODAY()`]];
    sheet.getRange("D6:D6").format.font.bold = true;

    await context.sync();
  } catch (e) {
    console.error(`❌ Additional Assumptions section failed:`, e);
  }

  try {
    sheet.getRange("D10:E10").values = [["HISTORICAL METRICS", ""]];
    sheet.getRange("D10:E10").format.font.bold = true;
    sheet.getRange("D10:E10").format.fill.color = "#D9E2F3";

    sheet.getRange("D11").values = [["EBIT / Revenue"]];
    sheet.getRange("D11").format.font.bold = true;
    sheet.getRange("E11").formulasLocal = [[
      `=AVERAGE(OFFSET(All_Financials_${symbol}!$A$1;MATCH("E B I T";All_Financials_${symbol}!$A:$A;0)-1;1;1;4)/OFFSET(All_Financials_${symbol}!$A$1;MATCH("Total Revenue";All_Financials_${symbol}!$A:$A;0)-1;1;1;4))`
    ]];
    sheet.getRange("E11").numberFormat = [["0.00%"]];

    sheet.getRange("D12").values = [["Gross PPE / Revenue"]];
    sheet.getRange("D12").format.font.bold = true;
    sheet.getRange("E12").formulasLocal = [[
      `=AVERAGE(OFFSET(All_Financials_${symbol}!$A$1;MATCH("Gross P P E";All_Financials_${symbol}!$A:$A;0)-1;1;1;4)/OFFSET(All_Financials_${symbol}!$A$1;MATCH("Total Revenue";All_Financials_${symbol}!$A:$A;0)-1;1;1;4))`
    ]];
    sheet.getRange("E12").numberFormat = [["0.00%"]];

    sheet.getRange("D13").values = [["Depreciation / Gross PPE"]];
    sheet.getRange("D13").format.font.bold = true;
    sheet.getRange("E13").formulasLocal = [[
      `=AVERAGE(-(OFFSET(All_Financials_${symbol}!$A$1;MATCH("Accumulated Depreciation";All_Financials_${symbol}!$A:$A;0)-1;1;1;3)-OFFSET(All_Financials_${symbol}!$A$1;MATCH("Accumulated Depreciation";All_Financials_${symbol}!$A:$A;0)-1;2;1;3))/OFFSET(All_Financials_${symbol}!$A$1;MATCH("Gross P P E";All_Financials_${symbol}!$A:$A;0)-1;1;1;3))`
    ]];
    sheet.getRange("E13").numberFormat = [["0.00%"]];

    sheet.getRange("D14").values = [["NWC / Revenue"]];
    sheet.getRange("D14").format.font.bold = true;
    sheet.getRange("E14").formulasLocal = [[
      `=AVERAGE((OFFSET(All_Financials_${symbol}!$A$1;MATCH("Current Assets";All_Financials_${symbol}!$A:$A;0)-1;1;1;4)-OFFSET(All_Financials_${symbol}!$A$1;MATCH("Current Liabilities";All_Financials_${symbol}!$A:$A;0)-1;1;1;4))/OFFSET(All_Financials_${symbol}!$A$1;MATCH("Total Revenue";All_Financials_${symbol}!$A:$A;0)-1;1;1;4))`
    ]];
    sheet.getRange("E14").numberFormat = [["0.00%"]];

    await context.sync();
  } catch (e) {
    console.error(`❌ Historical Metrics section failed:`, e);
  }

  try {
    currentYear = Number(currentYear);
    const historicalYears = [currentYear - 3, currentYear - 2, currentYear - 1, currentYear, currentYear + 1, currentYear + 2, currentYear + 3, currentYear + 4, currentYear + 5, currentYear + 6, currentYear + 7, currentYear + 8, currentYear + 9];
    const allYearLabels = historicalYears.map(y => String(y));

    
    console.log(`📅 DCF using dynamic years: ${allYearLabels.join(', ')}`);
    console.log(typeof currentYear)
    
    sheet.getRange("B25:E25").values = [allYearLabels.slice(0, 4)];
    sheet.getRange("B25:E25").format.font.bold = true;
    sheet.getRange("B25:D25").format.fill.color = "#E8E8E8"; 
    sheet.getRange("E25").format.fill.color = "#C6E0B4"; 
    sheet.getRange("E25").format.font.bold = true;

    sheet.getRange("F25:K25").values = [allYearLabels.slice(4, 10)];
    sheet.getRange("F25:K25").format.font.bold = true;
    sheet.getRange("F25:K25").format.fill.color = "#D9E2F3"; 

    sheet.getRange("L25:N25").values = [allYearLabels.slice(10, 13)];
    sheet.getRange("L25:N25").format.font.bold = true;
    sheet.getRange("L25:N25").format.fill.color = "#D9E2F3"; 

    sheet.getRange("E26").values = [["[Current Year]"]];
    sheet.getRange("E26").format.italic = true;
    sheet.getRange("E26").format.font.size = 9;

    const metrics = ["Revenue", "PPE", "NWC", "Depreciation", "", "EBIT", "NOPAT", "CapEx", "Change in NWC", "", "Free Cash Flow", "", "Terminal Value", "", "PV of FCF"];
    for (let i = 0; i < metrics.length; i++) {
      if (metrics[i] !== "") {
        sheet.getRange(`A${27 + i}:A${27 + i}`).values = [[metrics[i]]];
        sheet.getRange(`A${27 + i}:A${27 + i}`).format.font.bold = true;
      }
    }

    sheet.getRange("B27").formulas = [[`=INDEX(All_Financials_${symbol}!E:E,MATCH("Total Revenue",All_Financials_${symbol}!A:A,0))`]];
    sheet.getRange("C27").formulas = [[`=INDEX(All_Financials_${symbol}!D:D,MATCH("Total Revenue",All_Financials_${symbol}!A:A,0))`]];
    sheet.getRange("D27").formulas = [[`=INDEX(All_Financials_${symbol}!C:C,MATCH("Total Revenue",All_Financials_${symbol}!A:A,0))`]];
    sheet.getRange("E27").formulas = [[`=INDEX(All_Financials_${symbol}!B:B,MATCH("Total Revenue",All_Financials_${symbol}!A:A,0))`]];
    sheet.getRange("B27:E27").numberFormat = "#,##0";

    sheet.getRange("B28").formulas = [[`=INDEX(All_Financials_${symbol}!E:E,MATCH("Gross P P E",All_Financials_${symbol}!A:A,0))`]];
    sheet.getRange("C28").formulas = [[`=INDEX(All_Financials_${symbol}!D:D,MATCH("Gross P P E",All_Financials_${symbol}!A:A,0))`]];
    sheet.getRange("D28").formulas = [[`=INDEX(All_Financials_${symbol}!C:C,MATCH("Gross P P E",All_Financials_${symbol}!A:A,0))`]];
    sheet.getRange("E28").formulas = [[`=INDEX(All_Financials_${symbol}!B:B,MATCH("Gross P P E",All_Financials_${symbol}!A:A,0))`]];
    sheet.getRange("B28:E28").numberFormat = "#,##0";

    sheet.getRange("B29").formulas = [[`=INDEX(All_Financials_${symbol}!E:E,MATCH("Current Assets",All_Financials_${symbol}!A:A,0))-INDEX(All_Financials_${symbol}!E:E,MATCH("Current Liabilities",All_Financials_${symbol}!A:A,0))`]];
    sheet.getRange("C29").formulas = [[`=INDEX(All_Financials_${symbol}!D:D,MATCH("Current Assets",All_Financials_${symbol}!A:A,0))-INDEX(All_Financials_${symbol}!D:D,MATCH("Current Liabilities",All_Financials_${symbol}!A:A,0))`]];
    sheet.getRange("D29").formulas = [[`=INDEX(All_Financials_${symbol}!C:C,MATCH("Current Assets",All_Financials_${symbol}!A:A,0))-INDEX(All_Financials_${symbol}!C:C,MATCH("Current Liabilities",All_Financials_${symbol}!A:A,0))`]];
    sheet.getRange("E29").formulas = [[`=INDEX(All_Financials_${symbol}!B:B,MATCH("Current Assets",All_Financials_${symbol}!A:A,0))-INDEX(All_Financials_${symbol}!B:B,MATCH("Current Liabilities",All_Financials_${symbol}!A:A,0))`]];
    sheet.getRange("B29:E29").numberFormat = "#,##0";

    sheet.getRange("B30").formulas = [[``]];
    sheet.getRange("C30").formulas = [[`=-(INDEX(All_Financials_${symbol}!D:D,MATCH("Accumulated Depreciation",All_Financials_${symbol}!A:A,0))-INDEX(All_Financials_${symbol}!E:E,MATCH("Accumulated Depreciation",All_Financials_${symbol}!A:A,0)))`]];
    sheet.getRange("D30").formulas = [[`=-(INDEX(All_Financials_${symbol}!C:C,MATCH("Accumulated Depreciation",All_Financials_${symbol}!A:A,0))-INDEX(All_Financials_${symbol}!D:D,MATCH("Accumulated Depreciation",All_Financials_${symbol}!A:A,0)))`]];
    sheet.getRange("E30").formulas = [[`=-(INDEX(All_Financials_${symbol}!B:B,MATCH("Accumulated Depreciation",All_Financials_${symbol}!A:A,0))-INDEX(All_Financials_${symbol}!C:C,MATCH("Accumulated Depreciation",All_Financials_${symbol}!A:A,0)))`]];
    sheet.getRange("B30:E30").numberFormat = "#,##0";

    sheet.getRange("B32").formulas = [[`=INDEX(All_Financials_${symbol}!E:E,MATCH("E B I T",All_Financials_${symbol}!A:A,0))`]];
    sheet.getRange("C32").formulas = [[`=INDEX(All_Financials_${symbol}!D:D,MATCH("E B I T",All_Financials_${symbol}!A:A,0))`]];
    sheet.getRange("D32").formulas = [[`=INDEX(All_Financials_${symbol}!C:C,MATCH("E B I T",All_Financials_${symbol}!A:A,0))`]];
    sheet.getRange("E32").formulas = [[`=INDEX(All_Financials_${symbol}!B:B,MATCH("E B I T",All_Financials_${symbol}!A:A,0))`]];
    sheet.getRange("B32:E32").numberFormat = "#,##0";

    sheet.getRange("B33").formulas = [["=B32*(1-$B$21)"]];
    sheet.getRange("C33").formulas = [["=C32*(1-$B$21)"]];
    sheet.getRange("D33").formulas = [["=D32*(1-$B$21)"]];
    sheet.getRange("E33").formulas = [["=E32*(1-$B$21)"]];
    sheet.getRange("B33:E33").numberFormat = "#,##0";

    sheet.getRange("B34").formulas = [[""]];
    sheet.getRange("C34").formulas = [["=C28-B28"]];
    sheet.getRange("D34").formulas = [["=D28-C28"]];
    sheet.getRange("E34").formulas = [["=E28-D28"]];
    sheet.getRange("B34:E34").numberFormat = "#,##0";

    sheet.getRange("B35").formulas = [[""]];
    sheet.getRange("C35").formulas = [["=C29-B29"]];
    sheet.getRange("D35").formulas = [["=D29-C29"]];
    sheet.getRange("E35").formulas = [["=E29-D29"]];
    sheet.getRange("B35:E35").numberFormat = "#,##0";

    sheet.getRange("B37").formulas = [[""]];
    sheet.getRange("C37").formulas = [["=C33+C30-C34-C35"]];
    sheet.getRange("D37").formulas = [["=D33+D30-D34-D35"]];
    sheet.getRange("E37").formulas = [["=E33+E30-E34-E35"]];
    sheet.getRange("B37:E37").numberFormat = "#,##0";

    sheet.getRange("B41").formulas = [[""]];
    sheet.getRange("C41").formulas = [["=C37"]];
    sheet.getRange("D41").formulas = [["=D37"]];
    sheet.getRange("E41").formulas = [["=E37"]];
    sheet.getRange("B41:E41").numberFormat = "#,##0";



    sheet.getRange("F27").formulas = [["=E27*(1+$E$3)"]];
    sheet.getRange("G27").formulas = [["=F27*(1+$E$3)"]];
    sheet.getRange("H27").formulas = [["=G27*(1+$E$3)"]];
    sheet.getRange("I27").formulas = [["=H27*(1+$E$3)"]];
    sheet.getRange("J27").formulas = [["=I27*(1+$E$3)"]];
    sheet.getRange("K27").formulas = [["=J27*(1+$E$3)"]];
    sheet.getRange("L27").formulas = [["=K27*(1+$E$4)"]];
    sheet.getRange("M27").formulas = [["=L27*(1+$E$4)"]];
    sheet.getRange("N27").formulas = [["=M27*(1+$E$4)"]];
    sheet.getRange("F27:N27").numberFormat = "#,##0";

    sheet.getRange("F28").formulas = [["=F27*$E$12"]];
    sheet.getRange("G28").formulas = [["=G27*$E$12"]];
    sheet.getRange("H28").formulas = [["=H27*$E$12"]];
    sheet.getRange("I28").formulas = [["=I27*$E$12"]];
    sheet.getRange("J28").formulas = [["=J27*$E$12"]];
    sheet.getRange("K28").formulas = [["=K27*$E$12"]];
    sheet.getRange("L28").formulas = [["=L27*$E$12"]];
    sheet.getRange("M28").formulas = [["=M27*$E$12"]];
    sheet.getRange("N28").formulas = [["=N27*$E$12"]];
    sheet.getRange("F28:N28").numberFormat = "#,##0";

    sheet.getRange("F29").formulas = [["=F27*$E$14"]];
    sheet.getRange("G29").formulas = [["=G27*$E$14"]];
    sheet.getRange("H29").formulas = [["=H27*$E$14"]];
    sheet.getRange("I29").formulas = [["=I27*$E$14"]];
    sheet.getRange("J29").formulas = [["=J27*$E$14"]];
    sheet.getRange("K29").formulas = [["=K27*$E$14"]];
    sheet.getRange("L29").formulas = [["=L27*$E$14"]];
    sheet.getRange("M29").formulas = [["=M27*$E$14"]];
    sheet.getRange("N29").formulas = [["=N27*$E$14"]];
    sheet.getRange("F29:N29").numberFormat = "#,##0";

    sheet.getRange("F30").formulas = [["=F28*$E$13"]];
    sheet.getRange("G30").formulas = [["=G28*$E$13"]];
    sheet.getRange("H30").formulas = [["=H28*$E$13"]];
    sheet.getRange("I30").formulas = [["=I28*$E$13"]];
    sheet.getRange("J30").formulas = [["=J28*$E$13"]];
    sheet.getRange("K30").formulas = [["=K28*$E$13"]];
    sheet.getRange("L30").formulas = [["=L28*$E$13"]];
    sheet.getRange("M30").formulas = [["=M28*$E$13"]];
    sheet.getRange("N30").formulas = [["=N28*$E$13"]];
    sheet.getRange("F30:N30").numberFormat = "#,##0";

    sheet.getRange("F32").formulas = [["=F27*$E$11"]];
    sheet.getRange("G32").formulas = [["=G27*$E$11"]];
    sheet.getRange("H32").formulas = [["=H27*$E$11"]];
    sheet.getRange("I32").formulas = [["=I27*$E$11"]];
    sheet.getRange("J32").formulas = [["=J27*$E$11"]];
    sheet.getRange("K32").formulas = [["=K27*$E$11"]];
    sheet.getRange("L32").formulas = [["=L27*$E$11"]];
    sheet.getRange("M32").formulas = [["=M27*$E$11"]];
    sheet.getRange("N32").formulas = [["=N27*$E$11"]];
    sheet.getRange("F32:N32").numberFormat = "#,##0";

    sheet.getRange("F33").formulas = [["=F32*(1-$B$21)"]];
    sheet.getRange("G33").formulas = [["=G32*(1-$B$21)"]];
    sheet.getRange("H33").formulas = [["=H32*(1-$B$21)"]];
    sheet.getRange("I33").formulas = [["=I32*(1-$B$21)"]];
    sheet.getRange("J33").formulas = [["=J32*(1-$B$21)"]];
    sheet.getRange("K33").formulas = [["=K32*(1-$B$21)"]];
    sheet.getRange("L33").formulas = [["=L32*(1-$B$21)"]];
    sheet.getRange("M33").formulas = [["=M32*(1-$B$21)"]];
    sheet.getRange("N33").formulas = [["=N32*(1-$B$21)"]];
    sheet.getRange("F33:N33").numberFormat = "#,##0";

    sheet.getRange("F34").formulas = [["=F28-E28"]];
    sheet.getRange("G34").formulas = [["=G28-F28"]];
    sheet.getRange("H34").formulas = [["=H28-G28"]];
    sheet.getRange("I34").formulas = [["=I28-H28"]];
    sheet.getRange("J34").formulas = [["=J28-I28"]];
    sheet.getRange("K34").formulas = [["=K28-J28"]];
    sheet.getRange("L34").formulas = [["=L28-K28"]];
    sheet.getRange("M34").formulas = [["=M28-L28"]];
    sheet.getRange("N34").formulas = [["=N28-M28"]];
    sheet.getRange("F34:N34").numberFormat = "#,##0";

    sheet.getRange("F35").formulas = [["=F29-E29"]];
    sheet.getRange("G35").formulas = [["=G29-F29"]];
    sheet.getRange("H35").formulas = [["=H29-G29"]];
    sheet.getRange("I35").formulas = [["=I29-H29"]];
    sheet.getRange("J35").formulas = [["=J29-I29"]];
    sheet.getRange("K35").formulas = [["=K29-J29"]];
    sheet.getRange("L35").formulas = [["=L29-K29"]];
    sheet.getRange("M35").formulas = [["=M29-L29"]];
    sheet.getRange("N35").formulas = [["=N29-M29"]];
    sheet.getRange("F35:N35").numberFormat = "#,##0";

    sheet.getRange("F37").formulas = [["=F33+F30-F34-F35"]];
    sheet.getRange("G37").formulas = [["=G33+G30-G34-G35"]];
    sheet.getRange("H37").formulas = [["=H33+H30-H34-H35"]];
    sheet.getRange("I37").formulas = [["=I33+I30-I34-I35"]];
    sheet.getRange("J37").formulas = [["=J33+J30-J34-J35"]];
    sheet.getRange("K37").formulas = [["=K33+K30-K34-K35"]];
    sheet.getRange("L37").formulas = [["=L33+L30-L34-L35"]];
    sheet.getRange("M37").formulas = [["=M33+M30-M34-M35"]];
    sheet.getRange("N37").formulas = [["=N33+N30-N34-N35"]];
    sheet.getRange("F37:N37").numberFormat = "#,##0";

    sheet.getRange("N39").formulas = [["=N37*(1+$E$4)/( $B$22 - $E$4)"]]; 
    sheet.getRange("N39").numberFormat = "#,##0";

    sheet.getRange("F41").formulas = [["=F37/(1+$B$22)^1"]];
    sheet.getRange("G41").formulas = [["=G37/(1+$B$22)^2"]];
    sheet.getRange("H41").formulas = [["=H37/(1+$B$22)^3"]];
    sheet.getRange("I41").formulas = [["=I37/(1+$B$22)^4"]];
    sheet.getRange("J41").formulas = [["=J37/(1+$B$22)^5"]];
    sheet.getRange("K41").formulas = [["=K37/(1+$B$22)^6"]];
    sheet.getRange("L41").formulas = [["=L37/(1+$B$22)^7"]];
    sheet.getRange("M41").formulas = [["=M37/(1+$B$22)^8"]];
    sheet.getRange("N41").formulas = [["=(N37+N39)/(1+$B$22)^9"]];
    sheet.getRange("F41:N41").numberFormat = "#,##0";



    sheet.getRange("D43").values = [["DCF SUMMARY METRICS"]];
    sheet.getRange("D43").format.font.bold = true;
    sheet.getRange("D43").format.fill.color = "#D9E2F3";

    sheet.getRange("D44").values = [["Enterprise Value"]];
    sheet.getRange("E44").formulas = [["=SUM(F41:N41)"]];
    
    sheet.getRange("D46").values = [["Debt"]];
    sheet.getRange("E46").formulas = [[`=INDEX(All_Financials_${symbol}!B:B,MATCH("Total Debt",All_Financials_${symbol}!A:A,0))`]];
    
    sheet.getRange("D47").values = [["Cash"]];
    sheet.getRange("E47").formulas = [[`=INDEX(All_Financials_${symbol}!B:B,MATCH("Cash Cash Equivalents And Short Term Investments",All_Financials_${symbol}!A:A,0))`]];
    
    sheet.getRange("D49").values = [["Equity Value"]];
    sheet.getRange("E49").formulas = [["=E44-E46+E47"]];
    
    sheet.getRange("D51").values = [["Stock Price"]];
    sheet.getRange("E51").formulas = [["=E49/B5"]];

    sheet.getRange("D44").format.font.bold = true;
    sheet.getRange("D46").format.font.bold = true;
    sheet.getRange("D47").format.font.bold = true;
    sheet.getRange("D49").format.font.bold = true;
    sheet.getRange("D51").format.font.bold = true;
    
    sheet.getRange("E44").numberFormat = "#,##0";
    sheet.getRange("E46").numberFormat = "#,##0";
    sheet.getRange("E47").numberFormat = "#,##0";
    sheet.getRange("E49").numberFormat = "#,##0";
    sheet.getRange("E51").numberFormat = "#,##0.00";

    await context.sync();
  } catch (e) {
    console.error(`❌ DCF Projection Table failed:`, e);
    throw e;
  }

  try {
    sheet.getUsedRange().format.autofitColumns();
    await context.sync();
    console.log(`🎉 DCF Valuation sheet completed successfully`);
  } catch (e) {
    console.error(`❌ DCF formatting failed:`, e);
  }
}

function formatVariableName(variableName) {
  return variableName
    .replace(/^annual/, '')
    .replace(/([A-Z])/g, ' $1')
    .trim()
    .split(' ')
    .map(word => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase())
    .join(' ');
}

