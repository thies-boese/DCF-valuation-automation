const express = require('express');
const cors = require('cors');
const fetch = require('node-fetch');

const app = express();
const PORT = 3001; // Different port from your Excel add-in

// Yahoo Finance configuration (same as your client config)
const YAHOO_CONFIG = {
  crumb: 'jf7gEo2ICRR',
  cookie: 'A3=d=AQABBBrRNmkCEMKOkK7xMWe8TU15_ajSd38FEgABCAEgOGllaRSUJm0A9qMCAAcIGtE2aajSd38&S=AQAAAsYiwAyOfOuePX5isAikjiw',
  userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/142.0.0.0 Safari/537.36'
};

// Enable CORS for Excel add-in origin
app.use(cors({
  origin: ['https://localhost:3000', 'http://localhost:3000'],
  credentials: true,
  optionsSuccessStatus: 200
}));

app.use(express.json());

// Utility function to make Yahoo Finance requests
async function yahooFetch(url, needsAuth = true) {
  try {
    console.log(`📡 Fetching: ${url}`);
    
    const headers = {
      'User-Agent': YAHOO_CONFIG.userAgent
    };
    
    if (needsAuth) {
      headers['Cookie'] = YAHOO_CONFIG.cookie;
    }
    
    const response = await fetch(url, { headers });
    
    if (!response.ok) {
      throw new Error(`HTTP ${response.status}: ${response.statusText}`);
    }
    
    const data = await response.json();
    console.log(`✅ Success for: ${url}`);
    return data;
    
  } catch (error) {
    console.error(`❌ Error fetching ${url}:`, error.message);
    throw error;
  }
}

// Health check endpoint
app.get('/health', (req, res) => {
  res.json({ 
    status: 'Server running', 
    timestamp: new Date().toISOString(),
    port: PORT
  });
});

// Search companies endpoint
app.get('/api/search/:query', async (req, res) => {
  try {
    const { query } = req.params;
    console.log(`🔍 Searching for: ${query}`);
    
    const searchUrl = `https://query1.finance.yahoo.com/v1/finance/search?q=${encodeURIComponent(query)}&lang=en-US&region=US&quotesCount=10&newsCount=0`;
    const data = await yahooFetch(searchUrl, false); // Search doesn't need auth
    
    // Filter for stocks and ETFs
    const quotes = data.quotes || [];
    const stockQuotes = quotes.filter(quote => 
      quote.quoteType === 'EQUITY' || quote.quoteType === 'ETF'
    );
    
    res.json({
      success: true,
      data: stockQuotes,
      count: stockQuotes.length,
      query: query
    });
    
  } catch (error) {
    console.error(`Search error for "${req.params.query}":`, error);
    res.status(500).json({
      success: false,
      error: error.message,
      query: req.params.query,
      type: 'SEARCH_ERROR'
    });
  }
});

// Get company data endpoint
app.get('/api/company/:symbol', async (req, res) => {
  try {
    const { symbol } = req.params;
    console.log(`📊 Getting company data for: ${symbol}`);
    
    const modules = 'price,defaultKeyStatistics,financialData,calendarEvents,assetProfile,summaryDetail,majorHoldersBreakdown,upgradeDowngradeHistory';
    const companyUrl = `https://query1.finance.yahoo.com/v10/finance/quoteSummary/${encodeURIComponent(symbol)}?modules=${modules}&crumb=${YAHOO_CONFIG.crumb}&t=${Date.now()}`;
    
    const data = await yahooFetch(companyUrl, true);
    const companyData = data.quoteSummary?.result?.[0] || null;
    
    if (!companyData) {
      throw new Error(`No company data found for symbol: ${symbol}`);
    }
    
    res.json({
      success: true,
      data: companyData,
      symbol: symbol,
      timestamp: new Date().toISOString()
    });
    
  } catch (error) {
    console.error(`Company data error for ${req.params.symbol}:`, error);
    res.status(500).json({
      success: false,
      error: error.message,
      symbol: req.params.symbol,
      type: 'COMPANY_DATA_ERROR'
    });
  }
});

// Get historical financial data endpoint - UPDATED
app.get('/api/historical/:symbol', async (req, res) => {
  try {
    const { symbol } = req.params;
    console.log(`📈 Getting fundamentals timeseries data for: ${symbol}`);
    
    // Use the new fundamentals-timeseries endpoint
    const timeseriesUrl = `https://query1.finance.yahoo.com/ws/fundamentals-timeseries/v1/finance/timeseries/${encodeURIComponent(symbol)}?merge=false&padTimeSeries=true&period1=493590046&period2=${(Date.now()/1000).toFixed(0)}&type=annualTotalRevenue%2CannualOperatingRevenue%2CannualCostOfRevenue%2CannualGrossProfit%2CannualSellingGeneralAndAdministrative%2CannualSellingAndMarketingExpense%2CannualGeneralAndAdministrativeExpense%2CannualOtherGandA%2CannualResearchAndDevelopment%2CannualDepreciationAmortizationDepletionIncomeStatement%2CannualDepletionIncomeStatement%2CannualDepreciationAndAmortizationInIncomeStatement%2CannualAmortization%2CannualAmortizationOfIntangiblesIncomeStatement%2CannualDepreciationIncomeStatement%2CannualOtherOperatingExpenses%2CannualOperatingExpense%2CannualOperatingIncome%2CannualInterestExpenseNonOperating%2CannualInterestIncomeNonOperating%2CannualTotalOtherFinanceCost%2CannualNetNonOperatingInterestIncomeExpense%2CannualWriteOff%2CannualSpecialIncomeCharges%2CannualGainOnSaleOfPPE%2CannualGainOnSaleOfBusiness%2CannualGainOnSaleOfSecurity%2CannualOtherSpecialCharges%2CannualOtherIncomeExpense%2CannualOtherNonOperatingIncomeExpenses%2CannualTotalExpenses%2CannualPretaxIncome%2CannualTaxProvision%2CannualNetIncomeContinuousOperations%2CannualNetIncomeIncludingNoncontrollingInterests%2CannualMinorityInterests%2CannualNetIncomeFromTaxLossCarryforward%2CannualNetIncomeExtraordinary%2CannualNetIncomeDiscontinuousOperations%2CannualPreferredStockDividends%2CannualOtherunderPreferredStockDividend%2CannualNetIncomeCommonStockholders%2CannualNetIncome%2CannualBasicAverageShares%2CannualDilutedAverageShares%2CannualDividendPerShare%2CannualReportedNormalizedBasicEPS%2CannualContinuingAndDiscontinuedBasicEPS%2CannualBasicEPSOtherGainsLosses%2CannualTaxLossCarryforwardBasicEPS%2CannualNormalizedBasicEPS%2CannualBasicEPS%2CannualBasicAccountingChange%2CannualBasicExtraordinary%2CannualBasicDiscontinuousOperations%2CannualBasicContinuousOperations%2CannualReportedNormalizedDilutedEPS%2CannualContinuingAndDiscontinuedDilutedEPS%2CannualTaxLossCarryforwardDilutedEPS%2CannualAverageDilutionEarnings%2CannualNormalizedDilutedEPS%2CannualDilutedEPS%2CannualDilutedAccountingChange%2CannualDilutedExtraordinary%2CannualDilutedContinuousOperations%2CannualDilutedDiscontinuousOperations%2CannualDilutedNIAvailtoComStockholders%2CannualDilutedEPSOtherGainsLosses%2CannualTotalOperatingIncomeAsReported%2CannualNetIncomeFromContinuingAndDiscontinuedOperation%2CannualNormalizedIncome%2CannualNetInterestIncome%2CannualEBIT%2CannualEBITDA%2CannualReconciledCostOfRevenue%2CannualReconciledDepreciation%2CannualNetIncomeFromContinuingOperationNetMinorityInterest%2CannualTotalUnusualItemsExcludingGoodwill%2CannualTotalUnusualItems%2CannualNormalizedEBITDA%2CannualTaxRateForCalcs%2CannualTaxEffectOfUnusualItems%2CannualRentExpenseSupplemental%2CannualEarningsFromEquityInterestNetOfTax%2CannualImpairmentOfCapitalAssets%2CannualRestructuringAndMergernAcquisition%2CannualSecuritiesAmortization%2CannualEarningsFromEquityInterest%2CannualOtherTaxes%2CannualProvisionForDoubtfulAccounts%2CannualInsuranceAndClaims%2CannualRentAndLandingFees%2CannualSalariesAndWages%2CannualExciseTaxes%2CannualInterestExpense%2CannualInterestIncome%2CannualTotalMoneyMarketInvestments%2CannualInterestIncomeAfterProvisionForLoanLoss%2CannualOtherThanPreferredStockDividend%2CannualLossonExtinguishmentofDebt%2CannualIncomefromAssociatesandOtherParticipatingInterests%2CannualNonInterestExpense%2CannualOtherNonInterestExpense%2CannualProfessionalExpenseAndContractServicesExpense%2CannualOccupancyAndEquipment%2CannualEquipment%2CannualNetOccupancyExpense%2CannualCreditLossesProvision%2CannualNonInterestIncome%2CannualOtherNonInterestIncome%2CannualGainLossonSaleofAssets%2CannualGainonSaleofInvestmentProperty%2CannualGainonSaleofLoans%2CannualForeignExchangeTradingGains%2CannualTradingGainLoss%2CannualInvestmentBankingProfit%2CannualDividendIncome%2CannualFeesAndCommissions%2CannualFeesandCommissionExpense%2CannualFeesandCommissionIncome%2CannualOtherCustomerServices%2CannualCreditCard%2CannualSecuritiesActivities%2CannualTrustFeesbyCommissions%2CannualServiceChargeOnDepositorAccounts%2CannualTotalPremiumsEarned%2CannualOtherInterestExpense%2CannualInterestExpenseForFederalFundsSoldAndSecuritiesPurchaseUnderAgreementsToResell%2CannualInterestExpenseForLongTermDebtAndCapitalSecurities%2CannualInterestExpenseForShortTermDebt%2CannualInterestExpenseForDeposit%2CannualOtherInterestIncome%2CannualInterestIncomeFromFederalFundsSoldAndSecuritiesPurchaseUnderAgreementsToResell%2CannualInterestIncomeFromDeposits%2CannualInterestIncomeFromSecurities%2CannualInterestIncomeFromLoansAndLease%2CannualInterestIncomeFromLeases%2CannualInterestIncomeFromLoans%2CannualDepreciationDepreciationIncomeStatement%2CannualOperationAndMaintenance%2CannualOtherCostofRevenue%2CannualExplorationDevelopmentAndMineralPropertyLeaseExpenses%2CannualNetDebt%2CannualTreasurySharesNumber%2CannualPreferredSharesNumber%2CannualOrdinarySharesNumber%2CannualShareIssued%2CannualTotalDebt%2CannualTangibleBookValue%2CannualInvestedCapital%2CannualWorkingCapital%2CannualNetTangibleAssets%2CannualCapitalLeaseObligations%2CannualCommonStockEquity%2CannualPreferredStockEquity%2CannualTotalCapitalization%2CannualTotalEquityGrossMinorityInterest%2CannualMinorityInterest%2CannualStockholdersEquity%2CannualOtherEquityInterest%2CannualGainsLossesNotAffectingRetainedEarnings%2CannualOtherEquityAdjustments%2CannualFixedAssetsRevaluationReserve%2CannualForeignCurrencyTranslationAdjustments%2CannualMinimumPensionLiabilities%2CannualUnrealizedGainLoss%2CannualTreasuryStock%2CannualRetainedEarnings%2CannualAdditionalPaidInCapital%2CannualCapitalStock%2CannualOtherCapitalStock%2CannualCommonStock%2CannualPreferredStock%2CannualTotalPartnershipCapital%2CannualGeneralPartnershipCapital%2CannualLimitedPartnershipCapital%2CannualTotalLiabilitiesNetMinorityInterest%2CannualTotalNonCurrentLiabilitiesNetMinorityInterest%2CannualOtherNonCurrentLiabilities%2CannualLiabilitiesHeldforSaleNonCurrent%2CannualRestrictedCommonStock%2CannualPreferredSecuritiesOutsideStockEquity%2CannualDerivativeProductLiabilities%2CannualEmployeeBenefits%2CannualNonCurrentPensionAndOtherPostretirementBenefitPlans%2CannualNonCurrentAccruedExpenses%2CannualDuetoRelatedPartiesNonCurrent%2CannualTradeandOtherPayablesNonCurrent%2CannualNonCurrentDeferredLiabilities%2CannualNonCurrentDeferredRevenue%2CannualNonCurrentDeferredTaxesLiabilities%2CannualLongTermDebtAndCapitalLeaseObligation%2CannualLongTermCapitalLeaseObligation%2CannualLongTermDebt%2CannualLongTermProvisions%2CannualCurrentLiabilities%2CannualOtherCurrentLiabilities%2CannualCurrentDeferredLiabilities%2CannualCurrentDeferredRevenue%2CannualCurrentDeferredTaxesLiabilities%2CannualCurrentDebtAndCapitalLeaseObligation%2CannualCurrentCapitalLeaseObligation%2CannualCurrentDebt%2CannualOtherCurrentBorrowings%2CannualLineOfCredit%2CannualCommercialPaper%2CannualCurrentNotesPayable%2CannualPensionandOtherPostRetirementBenefitPlansCurrent%2CannualCurrentProvisions%2CannualPayablesAndAccruedExpenses%2CannualCurrentAccruedExpenses%2CannualInterestPayable%2CannualPayables%2CannualOtherPayable%2CannualDuetoRelatedPartiesCurrent%2CannualDividendsPayable%2CannualTotalTaxPayable%2CannualIncomeTaxPayable%2CannualAccountsPayable%2CannualTotalAssets%2CannualTotalNonCurrentAssets%2CannualOtherNonCurrentAssets%2CannualDefinedPensionBenefit%2CannualNonCurrentPrepaidAssets%2CannualNonCurrentDeferredAssets%2CannualNonCurrentDeferredTaxesAssets%2CannualDuefromRelatedPartiesNonCurrent%2CannualNonCurrentNoteReceivables%2CannualNonCurrentAccountsReceivable%2CannualFinancialAssets%2CannualInvestmentsAndAdvances%2CannualOtherInvestments%2CannualInvestmentinFinancialAssets%2CannualHeldToMaturitySecurities%2CannualAvailableForSaleSecurities%2CannualFinancialAssetsDesignatedasFairValueThroughProfitorLossTotal%2CannualTradingSecurities%2CannualLongTermEquityInvestment%2CannualInvestmentsinJointVenturesatCost%2CannualInvestmentsInOtherVenturesUnderEquityMethod%2CannualInvestmentsinAssociatesatCost%2CannualInvestmentsinSubsidiariesatCost%2CannualInvestmentProperties%2CannualGoodwillAndOtherIntangibleAssets%2CannualOtherIntangibleAssets%2CannualGoodwill%2CannualNetPPE%2CannualAccumulatedDepreciation%2CannualGrossPPE%2CannualLeases%2CannualConstructionInProgress%2CannualOtherProperties%2CannualMachineryFurnitureEquipment%2CannualBuildingsAndImprovements%2CannualLandAndImprovements%2CannualProperties%2CannualCurrentAssets%2CannualOtherCurrentAssets%2CannualHedgingAssetsCurrent%2CannualAssetsHeldForSaleCurrent%2CannualCurrentDeferredAssets%2CannualCurrentDeferredTaxesAssets%2CannualRestrictedCash%2CannualPrepaidAssets%2CannualInventory%2CannualInventoriesAdjustmentsAllowances%2CannualOtherInventories%2CannualFinishedGoods%2CannualWorkInProcess%2CannualRawMaterials%2CannualReceivables%2CannualReceivablesAdjustmentsAllowances%2CannualOtherReceivables%2CannualDuefromRelatedPartiesCurrent%2CannualTaxesReceivable%2CannualAccruedInterestReceivable%2CannualNotesReceivable%2CannualLoansReceivable%2CannualAccountsReceivable%2CannualAllowanceForDoubtfulAccountsReceivable%2CannualGrossAccountsReceivable%2CannualCashCashEquivalentsAndShortTermInvestments%2CannualOtherShortTermInvestments%2CannualCashAndCashEquivalents%2CannualCashEquivalents%2CannualCashFinancial%2CannualOtherLiabilities%2CannualLiabilitiesOfDiscontinuedOperations%2CannualSubordinatedLiabilities%2CannualAdvanceFromFederalHomeLoanBanks%2CannualTradingLiabilities%2CannualDuetoRelatedParties%2CannualSecuritiesLoaned%2CannualFederalFundsPurchasedAndSecuritiesSoldUnderAgreementToRepurchase%2CannualFinancialInstrumentsSoldUnderAgreementsToRepurchase%2CannualFederalFundsPurchased%2CannualTotalDeposits%2CannualNonInterestBearingDeposits%2CannualInterestBearingDepositsLiabilities%2CannualCustomerAccounts%2CannualDepositsbyBank%2CannualOtherAssets%2CannualAssetsHeldForSale%2CannualDeferredAssets%2CannualDeferredTaxAssets%2CannualDueFromRelatedParties%2CannualAllowanceForNotesReceivable%2CannualGrossNotesReceivable%2CannualNetLoan%2CannualUnearnedIncome%2CannualAllowanceForLoansAndLeaseLosses%2CannualGrossLoan%2CannualOtherLoanAssets%2CannualMortgageLoan%2CannualConsumerLoan%2CannualCommercialLoan%2CannualLoansHeldForSale%2CannualDerivativeAssets%2CannualSecuritiesAndInvestments%2CannualBankOwnedLifeInsurance%2CannualOtherRealEstateOwned%2CannualForeclosedAssets%2CannualCustomerAcceptances%2CannualFederalHomeLoanBankStock%2CannualSecurityBorrowed%2CannualCashCashEquivalentsAndFederalFundsSold%2CannualMoneyMarketInvestments%2CannualFederalFundsSoldAndSecuritiesPurchaseUnderAgreementsToResell%2CannualSecurityAgreeToBeResell%2CannualFederalFundsSold%2CannualRestrictedCashAndInvestments%2CannualRestrictedInvestments%2CannualRestrictedCashAndCashEquivalents%2CannualInterestBearingDepositsAssets%2CannualCashAndDueFromBanks%2CannualBankIndebtedness%2CannualMineralProperties%2CannualFreeCashFlow%2CannualForeignSales%2CannualDomesticSales%2CannualAdjustedGeographySegmentData%2CannualRepurchaseOfCapitalStock%2CannualRepaymentOfDebt%2CannualIssuanceOfDebt%2CannualIssuanceOfCapitalStock%2CannualCapitalExpenditure%2CannualInterestPaidSupplementalData%2CannualIncomeTaxPaidSupplementalData%2CannualEndCashPosition%2CannualOtherCashAdjustmentOutsideChangeinCash%2CannualBeginningCashPosition%2CannualEffectOfExchangeRateChanges%2CannualChangesInCash%2CannualOtherCashAdjustmentInsideChangeinCash%2CannualCashFlowFromDiscontinuedOperation%2CannualFinancingCashFlow%2CannualCashFromDiscontinuedFinancingActivities%2CannualCashFlowFromContinuingFinancingActivities%2CannualNetOtherFinancingCharges%2CannualInterestPaidCFF%2CannualProceedsFromStockOptionExercised%2CannualCashDividendsPaid%2CannualPreferredStockDividendPaid%2CannualCommonStockDividendPaid%2CannualNetPreferredStockIssuance%2CannualPreferredStockPayments%2CannualPreferredStockIssuance%2CannualNetCommonStockIssuance%2CannualCommonStockPayments%2CannualCommonStockIssuance%2CannualNetIssuancePaymentsOfDebt%2CannualNetShortTermDebtIssuance%2CannualShortTermDebtPayments%2CannualShortTermDebtIssuance%2CannualNetLongTermDebtIssuance%2CannualLongTermDebtPayments%2CannualLongTermDebtIssuance%2CannualInvestingCashFlow%2CannualCashFromDiscontinuedInvestingActivities%2CannualCashFlowFromContinuingInvestingActivities%2CannualNetOtherInvestingChanges%2CannualInterestReceivedCFI%2CannualDividendsReceivedCFI%2CannualNetInvestmentPurchaseAndSale%2CannualSaleOfInvestment%2CannualPurchaseOfInvestment%2CannualNetInvestmentPropertiesPurchaseAndSale%2CannualSaleOfInvestmentProperties%2CannualPurchaseOfInvestmentProperties%2CannualNetBusinessPurchaseAndSale%2CannualSaleOfBusiness%2CannualPurchaseOfBusiness%2CannualNetIntangiblesPurchaseAndSale%2CannualSaleOfIntangibles%2CannualPurchaseOfIntangibles%2CannualNetPPEPurchaseAndSale%2CannualSaleOfPPE%2CannualPurchaseOfPPE%2CannualCapitalExpenditureReported%2CannualOperatingCashFlow%2CannualCashFromDiscontinuedOperatingActivities%2CannualCashFlowFromContinuingOperatingActivities%2CannualTaxesRefundPaidDirect%2CannualInterestReceivedDirect%2CannualInterestPaidDirect%2CannualDividendsReceivedDirect%2CannualDividendsPaidDirect%2CannualClassesofCashPayments%2CannualOtherCashPaymentsfromOperatingActivities%2CannualPaymentsonBehalfofEmployees%2CannualPaymentstoSuppliersforGoodsandServices%2CannualClassesofCashReceiptsfromOperatingActivities%2CannualOtherCashReceiptsfromOperatingActivities%2CannualReceiptsfromGovernmentGrants%2CannualReceiptsfromCustomers%2CannualIncreaseDecreaseInDeposit%2CannualChangeInFederalFundsAndSecuritiesSoldForRepurchase%2CannualNetProceedsPaymentForLoan%2CannualPaymentForLoans%2CannualProceedsFromLoans%2CannualProceedsPaymentInInterestBearingDepositsinBank%2CannualIncreaseinInterestBearingDepositsinBank%2CannualDecreaseinInterestBearingDepositsinBank%2CannualProceedsPaymentFederalFundsSoldAndSecuritiesPurchasedUnderAgreementToResell%2CannualChangeInLoans%2CannualChangeInDeferredCharges%2CannualProvisionForLoanLeaseAndOtherLosses%2CannualAmortizationOfFinancingCostsAndDiscounts%2CannualDepreciationAmortizationDepletion%2CannualRealizedGainLossOnSaleOfLoansAndLease%2CannualAllTaxesPaid%2CannualInterestandCommissionPaid%2CannualCashPaymentsforLoans%2CannualCashPaymentsforDepositsbyBanksandCustomers%2CannualCashReceiptsfromFeesandCommissions%2CannualCashReceiptsfromSecuritiesRelatedActivities%2CannualCashReceiptsfromLoans%2CannualCashReceiptsfromDepositsbyBanksandCustomers%2CannualCashReceiptsfromTaxRefunds%2CannualAmortizationAmortizationCashFlow&lang=en-US&region=US`;
    
    const data = await yahooFetch(timeseriesUrl, true);
    const timeseriesData = data.timeseries?.result || [];
    
    res.json({
      success: true,
      data: timeseriesData,
      symbol: symbol,
      timestamp: new Date().toISOString()
    });
    
  } catch (error) {
    console.warn(`Fundamentals data warning for ${req.params.symbol}:`, error.message);
    // Historical data is optional, so return success with null data
    res.json({
      success: true,
      data: null,
      symbol: req.params.symbol,
      warning: error.message,
      type: 'FUNDAMENTALS_DATA_WARNING'
    });
  }
});

// Error handling middleware
app.use((error, req, res, next) => {
  console.error('🚨 Unhandled server error:', error);
  res.status(500).json({
    success: false,
    error: 'Internal server error',
    type: 'SERVER_ERROR',
    timestamp: new Date().toISOString()
  });
});

// 404 handler
app.use('*', (req, res) => {
  res.status(404).json({
    success: false,
    error: `Endpoint not found: ${req.method} ${req.originalUrl}`,
    type: 'NOT_FOUND',
    availableEndpoints: [
      'GET /health',
      'GET /api/search/:query',
      'GET /api/company/:symbol',
      'GET /api/historical/:symbol'
    ]
  });
});

// Start server
app.listen(PORT, () => {
  console.log('\n🚀 Yahoo Finance Proxy Server Started!');
  console.log(`📍 Server running on: http://localhost:${PORT}`);
  console.log('\n📋 Available endpoints:');
  console.log(`   GET /health - Health check`);
  console.log(`   GET /api/search/:query - Search companies`);
  console.log(`   GET /api/company/:symbol - Get company data`);
  console.log(`   GET /api/historical/:symbol - Get historical data`);
  console.log('\n💡 Usage:');
  console.log(`   Test: http://localhost:${PORT}/health`);
  console.log(`   Search: http://localhost:${PORT}/api/search/AAPL`);
  console.log('\n🔄 Ready for Excel add-in requests...\n');
});

// Graceful shutdown
process.on('SIGTERM', () => {
  console.log('🛑 SIGTERM received, shutting down gracefully...');
  process.exit(0);
});

process.on('SIGINT', () => {
  console.log('🛑 SIGINT received, shutting down gracefully...');
  process.exit(0);
});
