function _createForOfIteratorHelper(r, e) { var t = "undefined" != typeof Symbol && r[Symbol.iterator] || r["@@iterator"]; if (!t) { if (Array.isArray(r) || (t = _unsupportedIterableToArray(r)) || e && r && "number" == typeof r.length) { t && (r = t); var _n = 0, F = function F() {}; return { s: F, n: function n() { return _n >= r.length ? { done: !0 } : { done: !1, value: r[_n++] }; }, e: function e(r) { throw r; }, f: F }; } throw new TypeError("Invalid attempt to iterate non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."); } var o, a = !0, u = !1; return { s: function s() { t = t.call(r); }, n: function n() { var r = t.next(); return a = r.done, r; }, e: function e(r) { u = !0, o = r; }, f: function f() { try { a || null == t.return || t.return(); } finally { if (u) throw o; } } }; }
function _toConsumableArray(r) { return _arrayWithoutHoles(r) || _iterableToArray(r) || _unsupportedIterableToArray(r) || _nonIterableSpread(); }
function _nonIterableSpread() { throw new TypeError("Invalid attempt to spread non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."); }
function _iterableToArray(r) { if ("undefined" != typeof Symbol && null != r[Symbol.iterator] || null != r["@@iterator"]) return Array.from(r); }
function _arrayWithoutHoles(r) { if (Array.isArray(r)) return _arrayLikeToArray(r); }
function _slicedToArray(r, e) { return _arrayWithHoles(r) || _iterableToArrayLimit(r, e) || _unsupportedIterableToArray(r, e) || _nonIterableRest(); }
function _nonIterableRest() { throw new TypeError("Invalid attempt to destructure non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."); }
function _unsupportedIterableToArray(r, a) { if (r) { if ("string" == typeof r) return _arrayLikeToArray(r, a); var t = {}.toString.call(r).slice(8, -1); return "Object" === t && r.constructor && (t = r.constructor.name), "Map" === t || "Set" === t ? Array.from(r) : "Arguments" === t || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(t) ? _arrayLikeToArray(r, a) : void 0; } }
function _arrayLikeToArray(r, a) { (null == a || a > r.length) && (a = r.length); for (var e = 0, n = Array(a); e < a; e++) n[e] = r[e]; return n; }
function _iterableToArrayLimit(r, l) { var t = null == r ? null : "undefined" != typeof Symbol && r[Symbol.iterator] || r["@@iterator"]; if (null != t) { var e, n, i, u, a = [], f = !0, o = !1; try { if (i = (t = t.call(r)).next, 0 === l) { if (Object(t) !== t) return; f = !1; } else for (; !(f = (e = i.call(t)).done) && (a.push(e.value), a.length !== l); f = !0); } catch (r) { o = !0, n = r; } finally { try { if (!f && null != t.return && (u = t.return(), Object(u) !== u)) return; } finally { if (o) throw n; } } return a; } }
function _arrayWithHoles(r) { if (Array.isArray(r)) return r; }
function _regenerator() { /*! regenerator-runtime -- Copyright (c) 2014-present, Facebook, Inc. -- license (MIT): https://github.com/babel/babel/blob/main/packages/babel-helpers/LICENSE */ var e, t, r = "function" == typeof Symbol ? Symbol : {}, n = r.iterator || "@@iterator", o = r.toStringTag || "@@toStringTag"; function i(r, n, o, i) { var c = n && n.prototype instanceof Generator ? n : Generator, u = Object.create(c.prototype); return _regeneratorDefine2(u, "_invoke", function (r, n, o) { var i, c, u, f = 0, p = o || [], y = !1, G = { p: 0, n: 0, v: e, a: d, f: d.bind(e, 4), d: function d(t, r) { return i = t, c = 0, u = e, G.n = r, a; } }; function d(r, n) { for (c = r, u = n, t = 0; !y && f && !o && t < p.length; t++) { var o, i = p[t], d = G.p, l = i[2]; r > 3 ? (o = l === n) && (u = i[(c = i[4]) ? 5 : (c = 3, 3)], i[4] = i[5] = e) : i[0] <= d && ((o = r < 2 && d < i[1]) ? (c = 0, G.v = n, G.n = i[1]) : d < l && (o = r < 3 || i[0] > n || n > l) && (i[4] = r, i[5] = n, G.n = l, c = 0)); } if (o || r > 1) return a; throw y = !0, n; } return function (o, p, l) { if (f > 1) throw TypeError("Generator is already running"); for (y && 1 === p && d(p, l), c = p, u = l; (t = c < 2 ? e : u) || !y;) { i || (c ? c < 3 ? (c > 1 && (G.n = -1), d(c, u)) : G.n = u : G.v = u); try { if (f = 2, i) { if (c || (o = "next"), t = i[o]) { if (!(t = t.call(i, u))) throw TypeError("iterator result is not an object"); if (!t.done) return t; u = t.value, c < 2 && (c = 0); } else 1 === c && (t = i.return) && t.call(i), c < 2 && (u = TypeError("The iterator does not provide a '" + o + "' method"), c = 1); i = e; } else if ((t = (y = G.n < 0) ? u : r.call(n, G)) !== a) break; } catch (t) { i = e, c = 1, u = t; } finally { f = 1; } } return { value: t, done: y }; }; }(r, o, i), !0), u; } var a = {}; function Generator() {} function GeneratorFunction() {} function GeneratorFunctionPrototype() {} t = Object.getPrototypeOf; var c = [][n] ? t(t([][n]())) : (_regeneratorDefine2(t = {}, n, function () { return this; }), t), u = GeneratorFunctionPrototype.prototype = Generator.prototype = Object.create(c); function f(e) { return Object.setPrototypeOf ? Object.setPrototypeOf(e, GeneratorFunctionPrototype) : (e.__proto__ = GeneratorFunctionPrototype, _regeneratorDefine2(e, o, "GeneratorFunction")), e.prototype = Object.create(u), e; } return GeneratorFunction.prototype = GeneratorFunctionPrototype, _regeneratorDefine2(u, "constructor", GeneratorFunctionPrototype), _regeneratorDefine2(GeneratorFunctionPrototype, "constructor", GeneratorFunction), GeneratorFunction.displayName = "GeneratorFunction", _regeneratorDefine2(GeneratorFunctionPrototype, o, "GeneratorFunction"), _regeneratorDefine2(u), _regeneratorDefine2(u, o, "Generator"), _regeneratorDefine2(u, n, function () { return this; }), _regeneratorDefine2(u, "toString", function () { return "[object Generator]"; }), (_regenerator = function _regenerator() { return { w: i, m: f }; })(); }
function _regeneratorDefine2(e, r, n, t) { var i = Object.defineProperty; try { i({}, "", {}); } catch (e) { i = 0; } _regeneratorDefine2 = function _regeneratorDefine(e, r, n, t) { function o(r, n) { _regeneratorDefine2(e, r, function (e) { return this._invoke(r, n, e); }); } r ? i ? i(e, r, { value: n, enumerable: !t, configurable: !t, writable: !t }) : e[r] = n : (o("next", 0), o("throw", 1), o("return", 2)); }, _regeneratorDefine2(e, r, n, t); }
function asyncGeneratorStep(n, t, e, r, o, a, c) { try { var i = n[a](c), u = i.value; } catch (n) { return void e(n); } i.done ? t(u) : Promise.resolve(u).then(r, o); }
function _asyncToGenerator(n) { return function () { var t = this, e = arguments; return new Promise(function (r, o) { var a = n.apply(t, e); function _next(n) { asyncGeneratorStep(a, r, o, _next, _throw, "next", n); } function _throw(n) { asyncGeneratorStep(a, r, o, _next, _throw, "throw", n); } _next(void 0); }); }; }
import YahooFinance from 'yahoo-finance2';
var yahooFinance = new YahooFinance({
  suppressNotices: ['yahooSurvey']
});
var searchTimeout;
var currentMatches = [];
var selectedIndex = -1;
var selectedSymbol = "";
Office.onReady(function (info) {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("btnGetData").addEventListener("click", onGetDataClick);

    // Set up dynamic search
    var queryInput = document.getElementById("queryInput");
    var dropdown = document.getElementById("dropdown");
    queryInput.addEventListener("input", onInputChange);
    queryInput.addEventListener("keydown", onKeyDown);
    queryInput.addEventListener("focus", onInputFocus);

    // Click outside to close dropdown
    document.addEventListener("click", function (e) {
      if (!e.target.closest(".search-container")) {
        hideDropdown();
      }
    });
  }
});

// ================== DYNAMIC SEARCH FUNCTIONALITY ==================

function onInputChange(e) {
  var query = e.target.value.trim();
  selectedIndex = -1;
  selectedSymbol = "";

  // Clear existing timeout
  if (searchTimeout) {
    clearTimeout(searchTimeout);
  }
  if (query.length < 1) {
    hideDropdown();
    document.getElementById("status").textContent = "Type a company name or ticker symbol to search...";
    return;
  }

  // Debounce search - wait 400ms after user stops typing
  searchTimeout = setTimeout(function () {
    performSearch(query);
  }, 400);

  // Show loading immediately
  showLoadingDropdown();
}
function onKeyDown(e) {
  var dropdown = document.getElementById("dropdown");
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
  var query = e.target.value.trim();
  if (query.length >= 1 && currentMatches.length > 0) {
    showDropdown();
  }
}
function performSearch(_x) {
  return _performSearch.apply(this, arguments);
}
function _performSearch() {
  _performSearch = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee(query) {
    var statusElement, searchResults, matches, _t;
    return _regenerator().w(function (_context) {
      while (1) switch (_context.p = _context.n) {
        case 0:
          statusElement = document.getElementById("status");
          statusElement.textContent = "Searching for \"".concat(query, "\"...");
          _context.p = 1;
          _context.n = 2;
          return yahooFinance.search(query);
        case 2:
          searchResults = _context.v;
          matches = searchResults.quotes || [];
          currentMatches = matches;
          selectedIndex = -1;
          if (matches.length === 0) {
            statusElement.textContent = "No matches found. Try a different search term.";
            showNoResultsDropdown();
          } else {
            statusElement.textContent = "Found ".concat(matches.length, " match(es). Select one to get financial data.");
            showMatchesDropdown(matches);
          }
          _context.n = 4;
          break;
        case 3:
          _context.p = 3;
          _t = _context.v;
          console.error(_t);
          statusElement.textContent = "Search error: ".concat(_t.message || _t);
          showErrorDropdown(_t.message || _t);
        case 4:
          return _context.a(2);
      }
    }, _callee, null, [[1, 3]]);
  }));
  return _performSearch.apply(this, arguments);
}
function showLoadingDropdown() {
  var dropdown = document.getElementById("dropdown");
  dropdown.innerHTML = '<div class="loading">Searching...</div>';
  dropdown.style.display = "block";
}
function showMatchesDropdown(matches) {
  var dropdown = document.getElementById("dropdown");
  dropdown.innerHTML = matches.map(function (match, index) {
    var symbol = match.symbol || "";
    var name = match.shortname || "";
    var details = "";
    return "\n      <div class=\"dropdown-item\" data-index=\"".concat(index, "\">\n        <div>\n          <span class=\"company-symbol\">").concat(symbol, "</span>\n          <span class=\"company-name\">").concat(name, "</span>\n        </div>\n        ").concat(details ? "<div class=\"company-details\">".concat(details, "</div>") : "", "\n      </div>\n    ");
  }).join("");

  // Add click handlers
  dropdown.querySelectorAll(".dropdown-item").forEach(function (item, index) {
    item.addEventListener("click", function () {
      selectCompany(matches[index]);
    });
  });
  dropdown.style.display = "block";
}
function showNoResultsDropdown() {
  var dropdown = document.getElementById("dropdown");
  dropdown.innerHTML = '<div class="no-results">No companies found</div>';
  dropdown.style.display = "block";
}
function showErrorDropdown(error) {
  var dropdown = document.getElementById("dropdown");
  dropdown.innerHTML = "<div class=\"no-results\">Error: ".concat(error, "</div>");
  dropdown.style.display = "block";
}
function updateSelection() {
  var dropdown = document.getElementById("dropdown");
  var items = dropdown.querySelectorAll(".dropdown-item");
  items.forEach(function (item, index) {
    item.classList.toggle("highlighted", index === selectedIndex);
  });
}
function selectCompany(match) {
  var symbol = match.symbol || "";
  var name = match.shortname || "";
  selectedSymbol = symbol;
  document.getElementById("queryInput").value = "".concat(symbol, " - ").concat(name);
  hideDropdown();
  document.getElementById("status").textContent = "Selected: ".concat(symbol, " - ").concat(name, ". Click \"Get Financial Data\" to proceed.");
}
function showDropdown() {
  document.getElementById("dropdown").style.display = "block";
}
function hideDropdown() {
  document.getElementById("dropdown").style.display = "none";
  selectedIndex = -1;
}

// ================== GET DATA FUNCTION (YAHOO FINANCE) ==================
function onGetDataClick() {
  return _onGetDataClick.apply(this, arguments);
} // ================== FINANCIAL STATEMENT SHEETS (UPDATED FOR YAHOO FINANCE) ==================
function _onGetDataClick() {
  _onGetDataClick = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee3() {
    var statusElement, queryInput, symbol, _priceData$regularMar, _summaryDetail$previo, _summaryDetail$volume, _priceData$regularMar2, _priceData$regularMar3, _priceData$regularMar4, _summaryDetail$market, _keyStats$enterpriseV, _summaryDetail$traili, _keyStats$trailingEps, _summaryDetail$divide, _keyStats$bookValue, _summaryDetail$divide2, _financialData_module, _financialData_module2, _financialData_module3, _financialData_module4, _financialData_module5, _keyStats$sharesOutst, _keyStats$beta, _financialData_module6, _yield$Promise$all, _yield$Promise$all2, financialData, quoteSummary, sortedData, latestData, overview, summaryDetail, financialData_module, keyStats, priceData, price, volume, latestTradeDay, change, changePercent, companyName, assetType, companySector, companyIndustry, currency, companyAddress, fiscalYearEnd, companyMarketCap, companyEBITDA, companyPERatio, companyEPS, companyDividendYield, companyBookValue, companyDividendPerShare, companyProfitMargin, companyOperatingMarginTTM, companyROATTM, companyROETTM, companyRevenueTTM, companyGrossProfitTTM, companySharesOutstanding, companyBeta, companyLatestQuarter, companyTaxRate, companyTotalAssets, companyTotalLiabilities, companyTotalShareholderEquity, companyTotalCurrentAssets, companyTotalCurrentLiabilities, companyCashAndCashEquivalentsAtCarryingValue, companyShortTermInvestments, companyNetReceivables, companyInventory, companyOtherCurrentAssets, companyLongTermInvestments, companyPropertyPlantEquipment, companyGoodwill, companyIntangibleAssets, companyShortTermDebt, companyLongTermDebt, companyOtherCurrentLiabilities, companyOtherLiabilities, companyCommonStock, companyRetainedEarnings, companyOtherShareholderEquity, companyTotalDebt, companyTotalRevenue, companyCostOfRevenue, companyGrossProfit, companyOperatingIncome, companyNetIncome, companyOperatingExpenses, companySellingGeneralAndAdministrative, companyResearchAndDevelopment, companyInterestExpense, companyIncomeTaxExpense, companyDepreciationAndAmortization, companyIncomeBeforeTax, companyCostOfGoodsAndServicesSold, companyOperatingCashFlow, companyCapitalExpenditures, companyChangeInOperatingAssets, companyChangeInOperatingLiabilities, companyDividendPayouts, companyNetDebt, companyEBITMargin, companyEBITDAMargin, companyGrossMargin, companyNetWorkingCapital, companyFreeCashFlow, MarketPremium, RiskFreeRate, companyCostOfEquity, companyCostOfDebt, companyWeightDebt, companyWeightEquity, companyWACC, growthRates, i, currentRevenue, previousRevenue, growth, companyRevenueGrowth, companyDAPercentage, companyCapExPercentage, companyNWCPercentage, _t2;
    return _regenerator().w(function (_context3) {
      while (1) switch (_context3.p = _context3.n) {
        case 0:
          statusElement = document.getElementById("status");
          queryInput = document.getElementById("queryInput");
          symbol = selectedSymbol || queryInput.value.trim().toUpperCase(); // If user typed something but didn't select from dropdown, try to extract symbol
          if (!selectedSymbol && queryInput.value.includes(" - ")) {
            symbol = queryInput.value.split(" - ")[0].trim();
          }
          if (symbol) {
            _context3.n = 1;
            break;
          }
          statusElement.textContent = "Please search and select a company first.";
          queryInput.focus();
          return _context3.a(2);
        case 1:
          statusElement.textContent = "Fetching data for ".concat(symbol, "...");
          _context3.p = 2;
          _context3.n = 3;
          return Promise.all([yahooFinance.fundamentalsTimeSeries(symbol, {
            period1: "2018-01-01",
            type: "annual",
            module: "all"
          }), yahooFinance.quoteSummary(symbol, {
            modules: ['summaryDetail', 'financialData', 'defaultKeyStatistics', 'assetProfile', 'price']
          })]);
        case 3:
          _yield$Promise$all = _context3.v;
          _yield$Promise$all2 = _slicedToArray(_yield$Promise$all, 2);
          financialData = _yield$Promise$all2[0];
          quoteSummary = _yield$Promise$all2[1];
          if (!(!financialData || financialData.length === 0)) {
            _context3.n = 4;
            break;
          }
          throw new Error("No financial data available for ".concat(symbol));
        case 4:
          // Sort by date to get most recent first
          sortedData = financialData.sort(function (a, b) {
            return new Date(b.date) - new Date(a.date);
          });
          latestData = sortedData[0]; // Extract data from Yahoo Finance structure
          overview = quoteSummary.assetProfile || {};
          summaryDetail = quoteSummary.summaryDetail || {};
          financialData_module = quoteSummary.financialData || {};
          keyStats = quoteSummary.defaultKeyStatistics || {};
          priceData = quoteSummary.price || {}; // Map Yahoo Finance data to our existing variable structure
          //QUOTE DATA
          price = parseFloat(((_priceData$regularMar = priceData.regularMarketPrice) === null || _priceData$regularMar === void 0 ? void 0 : _priceData$regularMar.raw) || ((_summaryDetail$previo = summaryDetail.previousClose) === null || _summaryDetail$previo === void 0 ? void 0 : _summaryDetail$previo.raw) || "0");
          volume = parseFloat(((_summaryDetail$volume = summaryDetail.volume) === null || _summaryDetail$volume === void 0 ? void 0 : _summaryDetail$volume.raw) || "0");
          latestTradeDay = ((_priceData$regularMar2 = priceData.regularMarketTime) === null || _priceData$regularMar2 === void 0 ? void 0 : _priceData$regularMar2.fmt) || "";
          change = parseFloat(((_priceData$regularMar3 = priceData.regularMarketChange) === null || _priceData$regularMar3 === void 0 ? void 0 : _priceData$regularMar3.raw) || "0");
          changePercent = ((_priceData$regularMar4 = priceData.regularMarketChangePercent) === null || _priceData$regularMar4 === void 0 ? void 0 : _priceData$regularMar4.fmt) || ""; //OVERVIEW DATA
          companyName = overview.longName || priceData.longName || "";
          assetType = "Common Stock";
          companySector = overview.sector || "";
          companyIndustry = overview.industry || "";
          currency = priceData.currency || "USD";
          companyAddress = overview.address1 ? "".concat(overview.address1, ", ").concat(overview.city, ", ").concat(overview.state, ", ").concat(overview.country, ", ").concat(overview.zip) : "";
          fiscalYearEnd = overview.fiscalYearEnd || "";
          companyMarketCap = parseFloat(((_summaryDetail$market = summaryDetail.marketCap) === null || _summaryDetail$market === void 0 ? void 0 : _summaryDetail$market.raw) || ((_keyStats$enterpriseV = keyStats.enterpriseValue) === null || _keyStats$enterpriseV === void 0 ? void 0 : _keyStats$enterpriseV.raw) || "0");
          companyEBITDA = parseFloat(latestData.EBITDA || "0");
          companyPERatio = parseFloat(((_summaryDetail$traili = summaryDetail.trailingPE) === null || _summaryDetail$traili === void 0 ? void 0 : _summaryDetail$traili.raw) || "0");
          companyEPS = parseFloat(((_keyStats$trailingEps = keyStats.trailingEps) === null || _keyStats$trailingEps === void 0 ? void 0 : _keyStats$trailingEps.raw) || latestData.basicEPS || "0");
          companyDividendYield = parseFloat(((_summaryDetail$divide = summaryDetail.dividendYield) === null || _summaryDetail$divide === void 0 ? void 0 : _summaryDetail$divide.raw) || "0");
          companyBookValue = parseFloat(((_keyStats$bookValue = keyStats.bookValue) === null || _keyStats$bookValue === void 0 ? void 0 : _keyStats$bookValue.raw) || "0");
          companyDividendPerShare = parseFloat(((_summaryDetail$divide2 = summaryDetail.dividendRate) === null || _summaryDetail$divide2 === void 0 ? void 0 : _summaryDetail$divide2.raw) || "0");
          companyProfitMargin = parseFloat(((_financialData_module = financialData_module.profitMargins) === null || _financialData_module === void 0 ? void 0 : _financialData_module.raw) || "0");
          companyOperatingMarginTTM = parseFloat(((_financialData_module2 = financialData_module.operatingMargins) === null || _financialData_module2 === void 0 ? void 0 : _financialData_module2.raw) || "0");
          companyROATTM = parseFloat(((_financialData_module3 = financialData_module.returnOnAssets) === null || _financialData_module3 === void 0 ? void 0 : _financialData_module3.raw) || "0");
          companyROETTM = parseFloat(((_financialData_module4 = financialData_module.returnOnEquity) === null || _financialData_module4 === void 0 ? void 0 : _financialData_module4.raw) || "0");
          companyRevenueTTM = parseFloat(((_financialData_module5 = financialData_module.totalRevenue) === null || _financialData_module5 === void 0 ? void 0 : _financialData_module5.raw) || latestData.operatingRevenue || "0");
          companyGrossProfitTTM = parseFloat(latestData.grossProfit || "0");
          companySharesOutstanding = parseFloat(((_keyStats$sharesOutst = keyStats.sharesOutstanding) === null || _keyStats$sharesOutst === void 0 ? void 0 : _keyStats$sharesOutst.raw) || latestData.ordinarySharesNumber || "0");
          companyBeta = parseFloat(((_keyStats$beta = keyStats.beta) === null || _keyStats$beta === void 0 ? void 0 : _keyStats$beta.raw) || "0");
          companyLatestQuarter = ((_financialData_module6 = financialData_module.mostRecentQuarter) === null || _financialData_module6 === void 0 ? void 0 : _financialData_module6.fmt) || "";
          companyTaxRate = parseFloat(latestData.taxRateForCalcs || "0.25"); //BALANCE SHEET DATA (from latest Yahoo Finance data)
          companyTotalAssets = parseFloat(latestData.totalAssets || "0");
          companyTotalLiabilities = parseFloat(latestData.totalLiabilitiesNetMinorityInterest || "0");
          companyTotalShareholderEquity = parseFloat(latestData.totalEquityGrossMinorityInterest || "0");
          companyTotalCurrentAssets = parseFloat(latestData.currentAssets || "0");
          companyTotalCurrentLiabilities = parseFloat(latestData.currentLiabilities || "0");
          companyCashAndCashEquivalentsAtCarryingValue = parseFloat(latestData.cashAndCashEquivalents || "0");
          companyShortTermInvestments = parseFloat(latestData.otherShortTermInvestments || "0");
          companyNetReceivables = parseFloat(latestData.accountsReceivable || "0");
          companyInventory = parseFloat(latestData.inventory || "0");
          companyOtherCurrentAssets = parseFloat(latestData.otherCurrentAssets || "0");
          companyLongTermInvestments = parseFloat(latestData.investmentsAndAdvances || "0");
          companyPropertyPlantEquipment = parseFloat(latestData.netPPE || "0");
          companyGoodwill = parseFloat(latestData.goodwill || "0");
          companyIntangibleAssets = parseFloat(latestData.otherIntangibleAssets || "0");
          companyShortTermDebt = parseFloat(latestData.currentDebt || "0");
          companyLongTermDebt = parseFloat(latestData.longTermDebt || "0");
          companyOtherCurrentLiabilities = parseFloat(latestData.otherCurrentLiabilities || "0");
          companyOtherLiabilities = parseFloat(latestData.otherNonCurrentLiabilities || "0");
          companyCommonStock = parseFloat(latestData.commonStock || "0");
          companyRetainedEarnings = parseFloat(latestData.retainedEarnings || "0");
          companyOtherShareholderEquity = parseFloat(latestData.otherEquityAdjustments || "0");
          companyTotalDebt = parseFloat(latestData.totalDebt || companyShortTermDebt + companyLongTermDebt); //INCOME STATEMENT DATA (from latest Yahoo Finance data)
          companyTotalRevenue = parseFloat(latestData.operatingRevenue || latestData.totalRevenue || "0");
          companyCostOfRevenue = parseFloat(latestData.reconciledCostOfRevenue || "0");
          companyGrossProfit = parseFloat(latestData.grossProfit || "0");
          companyOperatingIncome = parseFloat(latestData.operatingIncome || "0");
          companyNetIncome = parseFloat(latestData.netIncome || "0");
          companyOperatingExpenses = parseFloat(latestData.operatingExpense || "0");
          companySellingGeneralAndAdministrative = parseFloat(latestData.sellingGeneralAndAdministration || "0");
          companyResearchAndDevelopment = parseFloat(latestData.researchAndDevelopment || "0");
          companyInterestExpense = parseFloat(latestData.interestExpense || "0");
          companyIncomeTaxExpense = parseFloat(latestData.taxProvision || "0");
          companyDepreciationAndAmortization = parseFloat(latestData.depreciationAndAmortization || "0");
          companyIncomeBeforeTax = parseFloat(latestData.pretaxIncome || "0");
          companyCostOfGoodsAndServicesSold = parseFloat(latestData.reconciledCostOfRevenue || "0"); //CASH FLOW DATA (from latest Yahoo Finance data)
          companyOperatingCashFlow = parseFloat(latestData.operatingCashFlow || "0");
          companyCapitalExpenditures = Math.abs(parseFloat(latestData.capitalExpenditure || "0"));
          companyChangeInOperatingAssets = parseFloat(latestData.changeInOtherCurrentAssets || "0");
          companyChangeInOperatingLiabilities = parseFloat(latestData.changeInOtherCurrentLiabilities || "0");
          companyDividendPayouts = parseFloat(latestData.cashDividendsPaid || "0"); // Calculate derived metrics
          companyNetDebt = parseFloat(latestData.netDebt || companyTotalDebt - companyCashAndCashEquivalentsAtCarryingValue);
          companyEBITMargin = companyOperatingIncome / companyTotalRevenue || 0;
          companyEBITDAMargin = companyEBITDA / companyTotalRevenue || 0;
          companyGrossMargin = companyGrossProfit / companyTotalRevenue || 0;
          companyNetWorkingCapital = parseFloat(latestData.workingCapital || companyTotalCurrentAssets - companyTotalCurrentLiabilities);
          companyFreeCashFlow = parseFloat(latestData.freeCashFlow || companyOperatingCashFlow - companyCapitalExpenditures); // Financial calculations
          MarketPremium = 0.06;
          RiskFreeRate = 0.04;
          companyCostOfEquity = RiskFreeRate + companyBeta * MarketPremium;
          companyCostOfDebt = companyTotalDebt > 0 ? Math.abs(companyInterestExpense) / companyTotalDebt : 0;
          companyWeightDebt = companyTotalDebt > 0 ? companyTotalDebt / (companyTotalDebt + companySharesOutstanding * price) : 0;
          companyWeightEquity = 1 - companyWeightDebt;
          companyWACC = companyWeightDebt * companyCostOfDebt * (1 - companyTaxRate) + companyWeightEquity * companyCostOfEquity; // Calculate revenue growth from multiple years of data
          growthRates = [];
          for (i = 0; i < sortedData.length - 1; i++) {
            currentRevenue = parseFloat(sortedData[i].operatingRevenue || sortedData[i].totalRevenue || "0");
            previousRevenue = parseFloat(sortedData[i + 1].operatingRevenue || sortedData[i + 1].totalRevenue || "0");
            if (previousRevenue > 0 && currentRevenue > 0) {
              growth = (currentRevenue - previousRevenue) / previousRevenue;
              growthRates.push(growth);
            }
          }
          companyRevenueGrowth = growthRates.length > 0 ? growthRates.reduce(function (a, b) {
            return a + b;
          }, 0) / growthRates.length : 0;
          companyDAPercentage = companyTotalRevenue > 0 ? companyDepreciationAndAmortization / companyTotalRevenue : 0;
          companyCapExPercentage = companyTotalRevenue > 0 ? companyCapitalExpenditures / companyTotalRevenue : 0;
          companyNWCPercentage = companyTotalRevenue > 0 ? companyNetWorkingCapital / companyTotalRevenue : 0; // Create Excel sheets with the processed data
          _context3.n = 5;
          return Excel.run(/*#__PURE__*/function () {
            var _ref = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee2(context) {
              var workbook, sheet, headerRange, basicInfoStartRow, basicInfoData, basicInfoRange, tradingStartRow, tradingData, tradingRange, metricsStartRow, metricsData, metricsRange, ratiosStartRow, ratiosData, ratiosRange, analystStartRow, analystData, analystRange, ownershipStartRow, ownershipData, ownershipRange, descStartRow, descRange, tradingPriceRange, metricsValueRange, metricsDecimalRange, sharesRange, ratiosDecimalRange, ratiosPercentRange, ownershipPercentRange;
              return _regenerator().w(function (_context2) {
                while (1) switch (_context2.n) {
                  case 0:
                    workbook = context.workbook;
                    sheet = workbook.worksheets.getActiveWorksheet(); // Clear existing content
                    try {
                      sheet.getUsedRange().clear();
                    } catch (e) {
                      // No used range to clear
                    }

                    // Company Header
                    headerRange = sheet.getRange("A1:H1");
                    headerRange.merge();
                    headerRange.values = [["".concat(companyName || symbol, " (").concat(symbol, ") - Company Overview")]];
                    headerRange.format.font.bold = true;
                    headerRange.format.font.size = 16;
                    headerRange.format.horizontalAlignment = "Center";
                    headerRange.format.fill.color = "#4472C4";
                    headerRange.format.font.color = "white";

                    // Basic Information Section
                    basicInfoStartRow = 3;
                    sheet.getRange("A".concat(basicInfoStartRow, ":B").concat(basicInfoStartRow)).values = [["BASIC INFORMATION", ""]];
                    sheet.getRange("A".concat(basicInfoStartRow, ":B").concat(basicInfoStartRow)).format.font.bold = true;
                    sheet.getRange("A".concat(basicInfoStartRow, ":B").concat(basicInfoStartRow)).format.fill.color = "#D9E2F3";
                    basicInfoData = [["Symbol", symbol], ["Company Name", overview.longName || ""], ["Asset Type", overview.AssetType || ""], ["Exchange", overview.Exchange || ""], ["Currency", overview.Currency || ""], ["Country", overview.Country || ""], ["Sector", overview.Sector || ""], ["Industry", overview.Industry || ""], ["CIK", overview.CIK || ""], ["Official Site", overview.OfficialSite || ""], ["Address", overview.Address || ""], ["Fiscal Year End", overview.FiscalYearEnd || ""], ["Latest Quarter", overview.LatestQuarter || ""]];
                    basicInfoRange = sheet.getRangeByIndexes(basicInfoStartRow, 0, basicInfoData.length, 2);
                    basicInfoRange.values = basicInfoData;
                    basicInfoRange.getColumn(0).format.font.bold = true;

                    // Stock Price & Trading Section
                    tradingStartRow = basicInfoStartRow + basicInfoData.length + 2;
                    sheet.getRange("A".concat(tradingStartRow, ":B").concat(tradingStartRow)).values = [["STOCK PRICE & TRADING", ""]];
                    sheet.getRange("A".concat(tradingStartRow, ":B").concat(tradingStartRow)).format.font.bold = true;
                    sheet.getRange("A".concat(tradingStartRow, ":B").concat(tradingStartRow)).format.fill.color = "#D9E2F3";
                    tradingData = [["Current Price", price || ""], ["Previous Close", quote["08. previous close"] || ""], ["Change", quote["09. change"] || ""], ["Change %", quote["10. change percent"] || ""], ["Volume", volume || ""], ["52 Week High", overview["52WeekHigh"] || ""], ["52 Week Low", overview["52WeekLow"] || ""], ["50 Day Moving Average", overview["50DayMovingAverage"] || ""], ["200 Day Moving Average", overview["200DayMovingAverage"] || ""]];
                    tradingRange = sheet.getRangeByIndexes(tradingStartRow, 0, tradingData.length, 2);
                    tradingRange.values = tradingData;
                    tradingRange.getColumn(0).format.font.bold = true;

                    // Financial Metrics Section (Column D-E)
                    metricsStartRow = 3;
                    sheet.getRange("D".concat(metricsStartRow, ":E").concat(metricsStartRow)).values = [["FINANCIAL METRICS", ""]];
                    sheet.getRange("D".concat(metricsStartRow, ":E").concat(metricsStartRow)).format.font.bold = true;
                    sheet.getRange("D".concat(metricsStartRow, ":E").concat(metricsStartRow)).format.fill.color = "#D9E2F3";
                    metricsData = [["Market Cap", parseFloat(overview.MarketCapitalization || "0")], ["Enterprise Value", parseFloat(overview.MarketCapitalization || "0") + companyNetDebt], ["Revenue (TTM)", parseFloat(overview.RevenueTTM || "0")], ["Gross Profit (TTM)", parseFloat(overview.GrossProfitTTM || "0")], ["EBITDA", parseFloat(overview.EBITDA || "0")], ["Net Income (TTM)", parseFloat(overview.RevenueTTM || "0") * parseFloat(overview.ProfitMargin || "0")], ["EPS", parseFloat(overview.EPS || "0")], ["Diluted EPS (TTM)", parseFloat(overview.DilutedEPSTTM || "0")], ["Revenue Per Share (TTM)", parseFloat(overview.RevenuePerShareTTM || "0")], ["Book Value", parseFloat(overview.BookValue || "0")], ["Dividend Per Share", parseFloat(overview.DividendPerShare || "0")], ["Dividend Yield", parseFloat(overview.DividendYield || "0")], ["Shares Outstanding", parseFloat(overview.SharesOutstanding || "0")], ["Shares Float", parseFloat(overview.SharesFloat || "0")]];
                    metricsRange = sheet.getRangeByIndexes(metricsStartRow, 3, metricsData.length, 2);
                    metricsRange.values = metricsData;
                    metricsRange.getColumn(0).format.font.bold = true;

                    // Valuation & Ratios Section (Column G-H)  
                    ratiosStartRow = 3;
                    sheet.getRange("G".concat(ratiosStartRow, ":H").concat(ratiosStartRow)).values = [["VALUATION & RATIOS", ""]];
                    sheet.getRange("G".concat(ratiosStartRow, ":H").concat(ratiosStartRow)).format.font.bold = true;
                    sheet.getRange("G".concat(ratiosStartRow, ":H").concat(ratiosStartRow)).format.fill.color = "#D9E2F3";
                    ratiosData = [["P/E Ratio", parseFloat(overview.PERatio || "0")], ["Trailing P/E", parseFloat(overview.TrailingPE || "0")], ["Forward P/E", parseFloat(overview.ForwardPE || "0")], ["PEG Ratio", parseFloat(overview.PEGRatio || "0")], ["Price to Sales (TTM)", parseFloat(overview.PriceToSalesRatioTTM || "0")], ["Price to Book", parseFloat(overview.PriceToBookRatio || "0")], ["EV to Revenue", parseFloat(overview.EVToRevenue || "0")], ["EV to EBITDA", parseFloat(overview.EVToEBITDA || "0")], ["Profit Margin", parseFloat(overview.ProfitMargin || "0")], ["Operating Margin (TTM)", parseFloat(overview.OperatingMarginTTM || "0")], ["Return on Assets (TTM)", parseFloat(overview.ReturnOnAssetsTTM || "0")], ["Return on Equity (TTM)", parseFloat(overview.ReturnOnEquityTTM || "0")], ["Beta", parseFloat(overview.Beta || "0")], ["Quarterly Earnings Growth", parseFloat(overview.QuarterlyEarningsGrowthYOY || "0")], ["Quarterly Revenue Growth", parseFloat(overview.QuarterlyRevenueGrowthYOY || "0")]];
                    ratiosRange = sheet.getRangeByIndexes(ratiosStartRow, 6, ratiosData.length, 2);
                    ratiosRange.values = ratiosData;
                    ratiosRange.getColumn(0).format.font.bold = true;

                    // Analyst Information Section
                    analystStartRow = tradingStartRow + tradingData.length + 2;
                    sheet.getRange("A".concat(analystStartRow, ":B").concat(analystStartRow)).values = [["ANALYST INFORMATION", ""]];
                    sheet.getRange("A".concat(analystStartRow, ":B").concat(analystStartRow)).format.font.bold = true;
                    sheet.getRange("A".concat(analystStartRow, ":B").concat(analystStartRow)).format.fill.color = "#D9E2F3";
                    analystData = [["Analyst Target Price", parseFloat(overview.AnalystTargetPrice || "0")], ["Strong Buy Ratings", parseInt(overview.AnalystRatingStrongBuy || "0")], ["Buy Ratings", parseInt(overview.AnalystRatingBuy || "0")], ["Hold Ratings", parseInt(overview.AnalystRatingHold || "0")], ["Sell Ratings", parseInt(overview.AnalystRatingSell || "0")], ["Strong Sell Ratings", parseInt(overview.AnalystRatingStrongSell || "0")]];
                    analystRange = sheet.getRangeByIndexes(analystStartRow, 0, analystData.length, 2);
                    analystRange.values = analystData;
                    analystRange.getColumn(0).format.font.bold = true;

                    // Ownership & Dates Section (Column D-E)
                    ownershipStartRow = analystStartRow;
                    sheet.getRange("D".concat(ownershipStartRow, ":E").concat(ownershipStartRow)).values = [["OWNERSHIP & DATES", ""]];
                    sheet.getRange("D".concat(ownershipStartRow, ":E").concat(ownershipStartRow)).format.font.bold = true;
                    sheet.getRange("D".concat(ownershipStartRow, ":E").concat(ownershipStartRow)).format.fill.color = "#D9E2F3";
                    ownershipData = [["Percent Insiders", parseFloat(overview.PercentInsiders || "0")], ["Percent Institutions", parseFloat(overview.PercentInstitutions || "0")], ["Dividend Date", overview.DividendDate || ""], ["Ex-Dividend Date", overview.ExDividendDate || ""]];
                    ownershipRange = sheet.getRangeByIndexes(ownershipStartRow, 3, ownershipData.length, 2);
                    ownershipRange.values = ownershipData;
                    ownershipRange.getColumn(0).format.font.bold = true;

                    // Company Description Section
                    descStartRow = Math.max(analystStartRow + analystData.length, ownershipStartRow + ownershipData.length) + 2;
                    sheet.getRange("A".concat(descStartRow, ":H").concat(descStartRow)).values = [["COMPANY DESCRIPTION", "", "", "", "", "", "", ""]];
                    sheet.getRange("A".concat(descStartRow, ":H").concat(descStartRow)).format.font.bold = true;
                    sheet.getRange("A".concat(descStartRow, ":H").concat(descStartRow)).format.fill.color = "#D9E2F3";
                    descRange = sheet.getRange("A".concat(descStartRow + 1, ":H").concat(descStartRow + 6));
                    descRange.merge();
                    descRange.values = [[overview.Description || ""]];
                    descRange.format.wrapText = true;
                    descRange.format.verticalAlignment = "Top";
                    _context2.n = 1;
                    return context.sync();
                  case 1:
                    // Apply number formatting after data is populated
                    // Price formatting for trading data
                    tradingPriceRange = sheet.getRangeByIndexes(tradingStartRow, 1, 5, 1); // B column for first 5 rows
                    tradingPriceRange.numberFormat = [["$0.00"]];

                    // Volume formatting 
                    sheet.getRange("B".concat(tradingStartRow + 4)).numberFormat = [["#,##0"]];

                    // Large number formatting for financial metrics
                    metricsValueRange = sheet.getRangeByIndexes(metricsStartRow, 4, 6, 1); // E column for first 6 rows (large numbers)
                    metricsValueRange.numberFormat = [["#,##0"]];

                    // Decimal formatting for per-share metrics
                    metricsDecimalRange = sheet.getRangeByIndexes(metricsStartRow + 6, 4, 4, 1); // E column for EPS, etc.
                    metricsDecimalRange.numberFormat = [["0.00"]];

                    // Share count formatting
                    sharesRange = sheet.getRangeByIndexes(metricsStartRow + 12, 4, 2, 1); // E column for share counts
                    sharesRange.numberFormat = [["#,##0"]];

                    // Ratio formatting (first 8 ratios as decimals)
                    ratiosDecimalRange = sheet.getRangeByIndexes(ratiosStartRow, 7, 8, 1); // H column
                    ratiosDecimalRange.numberFormat = [["0.00"]];

                    // Percentage formatting for margins and returns
                    ratiosPercentRange = sheet.getRangeByIndexes(ratiosStartRow + 8, 7, 7, 1); // H column
                    ratiosPercentRange.numberFormat = [["0.00%"]];

                    // Analyst target price
                    sheet.getRange("B".concat(analystStartRow)).numberFormat = [["$0.00"]];

                    // Ownership percentages
                    ownershipPercentRange = sheet.getRangeByIndexes(ownershipStartRow, 4, 2, 1); // E column
                    ownershipPercentRange.numberFormat = [["0.00%"]];

                    // Auto-fit columns
                    sheet.getUsedRange().format.autofitColumns();

                    // ---------------- (2) CREATE FINANCIAL STATEMENT SHEETS ----------------
                    _context2.n = 2;
                    return createIncomeStatementSheet(context, symbol, companyName || symbol, sortedData);
                  case 2:
                    _context2.n = 3;
                    return createBalanceSheetSheet(context, symbol, companyName || symbol, sortedData);
                  case 3:
                    _context2.n = 4;
                    return createCashFlowSheet(context, symbol, companyName || symbol, sortedData);
                  case 4:
                    _context2.n = 5;
                    return context.sync();
                  case 5:
                    statusElement.textContent = "Complete company overview and financial statements created for \"".concat(symbol, "\".");
                  case 6:
                    return _context2.a(2);
                }
              }, _callee2);
            }));
            return function (_x12) {
              return _ref.apply(this, arguments);
            };
          }());
        case 5:
          _context3.n = 7;
          break;
        case 6:
          _context3.p = 6;
          _t2 = _context3.v;
          console.error(_t2);
          statusElement.textContent = "Error: ".concat(_t2.message || _t2);
        case 7:
          return _context3.a(2);
      }
    }, _callee3, null, [[2, 6]]);
  }));
  return _onGetDataClick.apply(this, arguments);
}
function createIncomeStatementSheet(_x2, _x3, _x4, _x5) {
  return _createIncomeStatementSheet.apply(this, arguments);
}
function _createIncomeStatementSheet() {
  _createIncomeStatementSheet = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee4(context, symbol, companyName, financialData) {
    var sheetName, sheet, used, reports, years, headerRow, lineItems, i, row, _lineItems$i, label, field, rowData, j, _iterator, _step, report, value, dataRange, tableRange, _t3;
    return _regenerator().w(function (_context4) {
      while (1) switch (_context4.p = _context4.n) {
        case 0:
          sheetName = "Income_Statement_".concat(symbol);
          sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
          sheet.load("name");
          _context4.n = 1;
          return context.sync();
        case 1:
          if (!sheet.isNullObject) {
            _context4.n = 2;
            break;
          }
          sheet = context.workbook.worksheets.add(sheetName);
          _context4.n = 5;
          break;
        case 2:
          _context4.p = 2;
          used = sheet.getUsedRange(false);
          used.load("address");
          _context4.n = 3;
          return context.sync();
        case 3:
          used.clear();
          _context4.n = 5;
          break;
        case 4:
          _context4.p = 4;
          _t3 = _context4.v;
        case 5:
          // Get up to 5 years of data
          reports = financialData.slice(0, 5);
          years = reports.map(function (r) {
            return new Date(r.date).getFullYear();
          }); // Header
          sheet.getRange("A1:F1").merge();
          sheet.getRange("A1").values = [["".concat(companyName, " (").concat(symbol, ") - Income Statement")]];
          sheet.getRange("A1").format.font.bold = true;
          sheet.getRange("A1").format.font.size = 14;
          sheet.getRange("A1").format.horizontalAlignment = "Center";

          // Column headers (years)
          headerRow = ["Line Item"].concat(_toConsumableArray(years));
          sheet.getRange("A3:".concat(String.fromCharCode(65 + years.length), "3")).values = [headerRow];
          sheet.getRange("A3:".concat(String.fromCharCode(65 + years.length), "3")).format.font.bold = true;
          sheet.getRange("A3:".concat(String.fromCharCode(65 + years.length), "3")).format.fill.color = "#D9D9D9";

          // Yahoo Finance field mapping for income statement
          lineItems = [["REVENUE & GROSS PROFIT", ""], ["Total Revenue", "operatingRevenue"], ["Cost of Revenue", "reconciledCostOfRevenue"], ["Gross Profit", "grossProfit"], ["", ""], ["OPERATING EXPENSES", ""], ["Selling General & Administrative", "sellingGeneralAndAdministrative"], ["Research and Development", "researchAndDevelopment"], ["Operating Expenses", "operatingExpense"], ["Depreciation and Amortization", "reconciledDepreciation"], ["Operating Income", "operatingIncome"], ["", ""], ["NON-OPERATING & OTHER", ""], ["Interest Income", "interestIncome"], ["Interest Expense", "interestExpense"], ["Other Income Expense", "otherIncomeExpense"], ["", ""], ["EARNINGS METRICS", ""], ["EBIT", "EBIT"], ["EBITDA", "EBITDA"], ["Income Before Tax", "pretaxIncome"], ["Income Tax Expense", "taxProvision"], ["Net Income", "netIncome"]]; // Fill data
          for (i = 0; i < lineItems.length; i++) {
            row = 4 + i;
            _lineItems$i = _slicedToArray(lineItems[i], 2), label = _lineItems$i[0], field = _lineItems$i[1];
            rowData = [label];
            if (field === "") {
              // Header rows or empty rows
              for (j = 0; j < years.length; j++) {
                rowData.push("");
              }
              if (label !== "") {
                sheet.getRange("A".concat(row, ":").concat(String.fromCharCode(65 + years.length)).concat(row)).format.font.bold = true;
                sheet.getRange("A".concat(row, ":").concat(String.fromCharCode(65 + years.length)).concat(row)).format.fill.color = "#F2F2F2";
              }
            } else {
              _iterator = _createForOfIteratorHelper(reports);
              try {
                for (_iterator.s(); !(_step = _iterator.n()).done;) {
                  report = _step.value;
                  value = parseFloat(report[field] || "0");
                  rowData.push(value === 0 ? "" : value);
                }
              } catch (err) {
                _iterator.e(err);
              } finally {
                _iterator.f();
              }
            }
            sheet.getRange("A".concat(row, ":").concat(String.fromCharCode(65 + years.length)).concat(row)).values = [rowData];
          }

          // Format numbers as currency
          dataRange = sheet.getRange("B4:".concat(String.fromCharCode(65 + years.length)).concat(3 + lineItems.length));
          dataRange.numberFormat = [["#,##0"]];

          // Bold the line item labels
          sheet.getRange("A4:A".concat(3 + lineItems.length)).format.font.bold = true;

          // Add borders
          tableRange = sheet.getRange("A3:".concat(String.fromCharCode(65 + years.length)).concat(3 + lineItems.length));
          tableRange.format.borders.getItem("EdgeBottom").style = "Continuous";
          tableRange.format.borders.getItem("EdgeTop").style = "Continuous";
          tableRange.format.borders.getItem("EdgeLeft").style = "Continuous";
          tableRange.format.borders.getItem("EdgeRight").style = "Continuous";
          tableRange.format.borders.getItem("InsideHorizontal").style = "Continuous";
          tableRange.format.borders.getItem("InsideVertical").style = "Continuous";
          sheet.getUsedRange().format.autofitColumns();
        case 6:
          return _context4.a(2);
      }
    }, _callee4, null, [[2, 4]]);
  }));
  return _createIncomeStatementSheet.apply(this, arguments);
}
function createBalanceSheetSheet(_x6, _x7, _x8, _x9) {
  return _createBalanceSheetSheet.apply(this, arguments);
}
function _createBalanceSheetSheet() {
  _createBalanceSheetSheet = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee5(context, symbol, companyName, financialData) {
    var sheetName, sheet, used, reports, years, lineItems, i, row, _lineItems$i2, label, field, rowData, j, _iterator2, _step2, report, value, dataRange, tableRange, _t4;
    return _regenerator().w(function (_context5) {
      while (1) switch (_context5.p = _context5.n) {
        case 0:
          sheetName = "Balance_Sheet_".concat(symbol);
          sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
          sheet.load("name");
          _context5.n = 1;
          return context.sync();
        case 1:
          if (!sheet.isNullObject) {
            _context5.n = 2;
            break;
          }
          sheet = context.workbook.worksheets.add(sheetName);
          _context5.n = 5;
          break;
        case 2:
          _context5.p = 2;
          used = sheet.getUsedRange(false);
          used.load("address");
          _context5.n = 3;
          return context.sync();
        case 3:
          used.clear();
          _context5.n = 5;
          break;
        case 4:
          _context5.p = 4;
          _t4 = _context5.v;
        case 5:
          reports = financialData.slice(0, 5);
          years = reports.map(function (r) {
            return new Date(r.date).getFullYear();
          }); // Yahoo Finance field mapping for balance sheet
          lineItems = [["ASSETS", ""], ["Total Assets", "totalAssets"], ["", ""], ["CURRENT ASSETS", ""], ["Total Current Assets", "currentAssets"], ["Cash and Cash Equivalents", "cashAndCashEquivalents"], ["Short Term Investments", "otherShortTermInvestments"], ["Accounts Receivable", "accountsReceivable"], ["Inventory", "inventory"], ["Other Current Assets", "otherCurrentAssets"], ["", ""], ["NON-CURRENT ASSETS", ""], ["Property Plant Equipment", "netPPE"], ["Investments", "investmentsAndAdvances"], ["Goodwill", "goodwill"], ["Other Intangible Assets", "otherIntangibleAssets"], ["Other Non Current Assets", "otherNonCurrentAssets"], ["", ""], ["LIABILITIES", ""], ["Total Liabilities", "totalLiabilitiesNetMinorityInterest"], ["", ""], ["CURRENT LIABILITIES", ""], ["Total Current Liabilities", "currentLiabilities"], ["Accounts Payable", "accountsPayable"], ["Current Debt", "currentDebt"], ["Other Current Liabilities", "otherCurrentLiabilities"], ["", ""], ["NON-CURRENT LIABILITIES", ""], ["Long Term Debt", "longTermDebt"], ["Other Non Current Liabilities", "otherNonCurrentLiabilities"], ["", ""], ["SHAREHOLDERS' EQUITY", ""], ["Total Shareholder Equity", "totalEquityGrossMinorityInterest"], ["Common Stock", "commonStock"], ["Retained Earnings", "retainedEarnings"]]; // Fill data
          for (i = 0; i < lineItems.length; i++) {
            row = 4 + i;
            _lineItems$i2 = _slicedToArray(lineItems[i], 2), label = _lineItems$i2[0], field = _lineItems$i2[1];
            rowData = [label];
            if (field === "") {
              // Header rows or empty rows
              for (j = 0; j < years.length; j++) {
                rowData.push("");
              }
              if (label !== "") {
                sheet.getRange("A".concat(row, ":").concat(String.fromCharCode(65 + years.length)).concat(row)).format.font.bold = true;
                sheet.getRange("A".concat(row, ":").concat(String.fromCharCode(65 + years.length)).concat(row)).format.fill.color = "#F2F2F2";
              }
            } else {
              _iterator2 = _createForOfIteratorHelper(reports);
              try {
                for (_iterator2.s(); !(_step2 = _iterator2.n()).done;) {
                  report = _step2.value;
                  value = parseFloat(report[field] || "0");
                  if (report[field] === "None" || isNaN(value)) value = 0;
                  rowData.push(value === 0 ? "" : value);
                }
              } catch (err) {
                _iterator2.e(err);
              } finally {
                _iterator2.f();
              }
            }
            sheet.getRange("A".concat(row, ":").concat(String.fromCharCode(65 + years.length)).concat(row)).values = [rowData];
          }

          // Format numbers as currency (skip share count)
          dataRange = sheet.getRange("B4:".concat(String.fromCharCode(65 + years.length)).concat(3 + lineItems.length));
          dataRange.numberFormat = [["#,##0"]];

          // Bold the line item labels
          sheet.getRange("A4:A".concat(3 + lineItems.length)).format.font.bold = true;

          // Add borders
          tableRange = sheet.getRange("A3:".concat(String.fromCharCode(65 + years.length)).concat(3 + lineItems.length));
          tableRange.format.borders.getItem("EdgeBottom").style = "Continuous";
          tableRange.format.borders.getItem("EdgeTop").style = "Continuous";
          tableRange.format.borders.getItem("EdgeLeft").style = "Continuous";
          tableRange.format.borders.getItem("EdgeRight").style = "Continuous";
          tableRange.format.borders.getItem("InsideHorizontal").style = "Continuous";
          tableRange.format.borders.getItem("InsideVertical").style = "Continuous";
          sheet.getUsedRange().format.autofitColumns();
        case 6:
          return _context5.a(2);
      }
    }, _callee5, null, [[2, 4]]);
  }));
  return _createBalanceSheetSheet.apply(this, arguments);
}
function createCashFlowSheet(_x0, _x1, _x10, _x11) {
  return _createCashFlowSheet.apply(this, arguments);
}
function _createCashFlowSheet() {
  _createCashFlowSheet = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee6(context, symbol, companyName, financialData) {
    var sheetName, sheet, used, reports, years, lineItems, i, row, _lineItems$i3, label, field, rowData, j, _iterator3, _step3, report, value, dataRange, tableRange, _t5;
    return _regenerator().w(function (_context6) {
      while (1) switch (_context6.p = _context6.n) {
        case 0:
          sheetName = "Cash_Flow_".concat(symbol);
          sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
          sheet.load("name");
          _context6.n = 1;
          return context.sync();
        case 1:
          if (!sheet.isNullObject) {
            _context6.n = 2;
            break;
          }
          sheet = context.workbook.worksheets.add(sheetName);
          _context6.n = 5;
          break;
        case 2:
          _context6.p = 2;
          used = sheet.getUsedRange(false);
          used.load("address");
          _context6.n = 3;
          return context.sync();
        case 3:
          used.clear();
          _context6.n = 5;
          break;
        case 4:
          _context6.p = 4;
          _t5 = _context6.v;
        case 5:
          reports = financialData.slice(0, 5);
          years = reports.map(function (r) {
            return new Date(r.date).getFullYear();
          }); // Yahoo Finance field mapping for cash flow
          lineItems = [["OPERATING ACTIVITIES", ""], ["Operating Cash Flow", "operatingCashFlow"], ["Net Income", "netIncome"], ["Depreciation and Amortization", "depreciationAndAmortization"], ["Change in Working Capital", "changeInWorkingCapital"], ["", ""], ["INVESTING ACTIVITIES", ""], ["Cash Flow from Investment", "investingCashFlow"], ["Capital Expenditures", "capitalExpenditure"], ["", ""], ["FINANCING ACTIVITIES", ""], ["Cash Flow from Financing", "financingCashFlow"], ["Cash Dividends Paid", "cashDividendsPaid"], ["Repurchase of Capital Stock", "repurchaseOfCapitalStock"], ["", ""], ["NET CHANGE", ""], ["Change in Cash", "changesInCash"], ["Free Cash Flow", "freeCashFlow"]]; // Fill data
          for (i = 0; i < lineItems.length; i++) {
            row = 4 + i;
            _lineItems$i3 = _slicedToArray(lineItems[i], 2), label = _lineItems$i3[0], field = _lineItems$i3[1];
            rowData = [label];
            if (field === "") {
              // Header rows or empty rows
              for (j = 0; j < years.length; j++) {
                rowData.push("");
              }
              if (label !== "") {
                sheet.getRange("A".concat(row, ":").concat(String.fromCharCode(65 + years.length)).concat(row)).format.font.bold = true;
                sheet.getRange("A".concat(row, ":").concat(String.fromCharCode(65 + years.length)).concat(row)).format.fill.color = "#F2F2F2";
              }
            } else {
              _iterator3 = _createForOfIteratorHelper(reports);
              try {
                for (_iterator3.s(); !(_step3 = _iterator3.n()).done;) {
                  report = _step3.value;
                  value = parseFloat(report[field] || "0");
                  if (report[field] === "None" || isNaN(value)) value = 0;
                  rowData.push(value === 0 ? "" : value);
                }
              } catch (err) {
                _iterator3.e(err);
              } finally {
                _iterator3.f();
              }
            }
            sheet.getRange("A".concat(row, ":").concat(String.fromCharCode(65 + years.length)).concat(row)).values = [rowData];
          }

          // Format numbers as currency
          dataRange = sheet.getRange("B4:".concat(String.fromCharCode(65 + years.length)).concat(3 + lineItems.length));
          dataRange.numberFormat = [["#,##0"]];

          // Bold the line item labels
          sheet.getRange("A4:A".concat(3 + lineItems.length)).format.font.bold = true;

          // Add borders
          tableRange = sheet.getRange("A3:".concat(String.fromCharCode(65 + years.length)).concat(3 + lineItems.length));
          tableRange.format.borders.getItem("EdgeBottom").style = "Continuous";
          tableRange.format.borders.getItem("EdgeTop").style = "Continuous";
          tableRange.format.borders.getItem("EdgeLeft").style = "Continuous";
          tableRange.format.borders.getItem("EdgeRight").style = "Continuous";
          tableRange.format.borders.getItem("InsideHorizontal").style = "Continuous";
          tableRange.format.borders.getItem("InsideVertical").style = "Continuous";
          sheet.getUsedRange().format.autofitColumns();
        case 6:
          return _context6.a(2);
      }
    }, _callee6, null, [[2, 4]]);
  }));
  return _createCashFlowSheet.apply(this, arguments);
}