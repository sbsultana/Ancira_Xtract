import { Routes } from '@angular/router';

export const routes: Routes = [
  { path: 'Header', loadComponent: () => import('./Layout/header/header').then(m => m.Header) },
  // ACCOUNTING BLOCK
  { path: 'AccountMapping', loadComponent: () => import('./Reports/Accounting/AccountMapping/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'FinancialStatement', loadComponent: () => import('./Reports/Accounting/FinancialStatement/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'FinancialSummary', loadComponent: () => import('./Reports/Accounting/FinancialSummary/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'ExpenseTrend', loadComponent: () => import('./Reports/Accounting/ExpenseTrend/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'FinancialStatementDetails', loadComponent: () => import('./Reports/Accounting/FinancialStatementDetails/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'IncomeStatementStore', loadComponent: () => import('./Reports/Accounting/IncomeStatementStoreComposite/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'IncomeStatementTrend', loadComponent: () => import('./Reports/Accounting/IncomeStatementTrend/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'LendersMapping', loadComponent: () => import('./Reports/Accounting/LenderMapping/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'FixedIncomeByExpense', loadComponent: () => import('./Reports/Accounting/EnterpriseIncomeByExpense/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'FixedIncomeByExpenseWithOutOverhead', loadComponent: () => import('./Reports/Accounting/EnterpriseIncomeByExpenseWithOutOverhead/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'VariableIncomeByExpense', loadComponent: () => import('./Reports/Accounting/EnterpriseIncomeByExpense/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'VariableIncomeByExpenseWithOutOverhead', loadComponent: () => import('./Reports/Accounting/EnterpriseIncomeByExpenseWithOutOverhead/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'EnterpriseTracking', loadComponent: () => import('./Reports/Accounting/EnterpriseTracking/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'LenderSummary', loadComponent: () => import('./Reports/Accounting/LenderSummary/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'EnterpriseIncomeByExpense', loadComponent: () => import('./Reports/Accounting/EnterpriseIncomeByExpense/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'EnterpriseIncomeByExpenseWithOutOverhead', loadComponent: () => import('./Reports/Accounting/EnterpriseIncomeByExpenseWithOutOverhead/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'EnterpriseExpense', loadComponent: () => import('./Reports/Accounting/EnterpriseExpense/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'EnterpriseIncomeByExpenseTrend', loadComponent: () => import('./Reports/Accounting/EnterpriseIncomeByExpenseTrend/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'NetAddsDeducts', loadComponent: () => import('./Reports/Accounting/NetAddsByDeducts/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'GLLookup', loadComponent: () => import('./Reports/Accounting/GLLookUp/dashboard/dashboard').then(m => m.Dashboard) },

  // RECEIVABLES BLOCK
  { path: 'Receivables', loadComponent: () => import('./Reports/Receivables/Receivables/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'BookedDeals', loadComponent: () => import('./Reports/Receivables/BookedDeals/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'CIT', loadComponent: () => import('./Reports/Receivables/CITFloorplan/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'Wholesale', loadComponent: () => import('./Reports/Receivables/Receivables/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'FinanceReserve', loadComponent: () => import('./Reports/Receivables/Receivables/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'Employee', loadComponent: () => import('./Reports/Receivables/Receivables/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'OpenAccountReceivables', loadComponent: () => import('./Reports/Receivables/OpenAccountsReceivables/dashboard/dashboard').then(m => m.Dashboard) },

  //LIABILITIES
  { path: 'Liabilities', loadComponent: () => import('./Reports/Liabilities/Liabilities/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'NewFlooring', loadComponent: () => import('./Reports/Liabilities/Liabilities/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'RentalInventory', loadComponent: () => import('./Reports/Liabilities/Liabilities/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'TT&L', loadComponent: () => import('./Reports/Liabilities/Liabilities/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'LienPayoffs', loadComponent: () => import('./Reports/Liabilities/Liabilities/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'WeOwe', loadComponent: () => import('./Reports/Liabilities/Liabilities/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'TitleTracking', loadComponent: () => import('./Reports/Liabilities/TitleTracking/dashboard/dashboard').then(m => m.Dashboard) },


  // SALES BLOCK
  { path: 'SalesGross', loadComponent: () => import('./Reports/Sales/SalesGross/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'VariableGrossGL', loadComponent: () => import('./Reports/Sales/VariableGrossGL/dashboard/dashboard').then(m => m.Dashboard) },

  { path: 'FandIProductPenetration', loadComponent: () => import('./Reports/Sales/FandIProductPenetration/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'FandIProductPenetrationV2', loadComponent: () => import('./Reports/Others/FandIProductPenetration/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'FandISummary', loadComponent: () => import('./Reports/Sales/FandISummary/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'SalespersonCommissions', loadComponent: () => import('./Reports/Sales/SalespersonCommissions/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'CarDeals', loadComponent: () => import('./Reports/Sales/CarDeals/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'Appointments', loadComponent: () => import('./Reports/Sales/Appointments/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'SalesContest', loadComponent: () => import('./Reports/Sales/SalesContest/dashboard/dashboard').then(m => m.Dashboard) },



  // SERVICE BLOCK
  { path: 'ServiceAppointments', loadComponent: () => import('./Reports/Services/ServiceAppointments/dashboard/dashboard').then(m => m.Dashboard) },


  //PARTS BLOCK
  { path: 'SearchParts', loadComponent: () => import('./Reports/Parts/SearchParts/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'PartsAging', loadComponent: () => import('./Reports/Parts/PartsAging/dashboard/dashboard').then(m => m.Dashboard) },

  // OTHERS

  { path: 'BudgetForecastInput', loadComponent: () => import('./Reports/Others/BudgetForecastInput/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'BudgetForecastInputAdd', loadComponent: () => import('./Reports/Others/BudgetForecastInput/budget-forecast-input-variables/budget-forecast-input-variables').then(m => m.BudgetForecastInputVariables) },
  { path: 'LeadSourceReport', loadComponent: () => import('./Reports/Others/LeadSourceReport/dashboard/dashboard').then(m => m.Dashboard) },

  //INVENTORY BLOCK
  { path: 'InventoryBrowser', loadComponent: () => import('./Reports/Inventory/InventoryBrowser/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'InventoryBook', loadComponent: () => import('./Reports/Inventory/InventoryBook/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'InventorySummary', loadComponent: () => import('./Reports/Inventory/InventorySummary/dashboard/dashboard').then(m => m.Dashboard) },



  //Rankings
  { path: 'SPranking', loadComponent: () => import('./Reports/Rankings/SalesPersonRanking/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'FandIManagerRanking', loadComponent: () => import('./Reports/Rankings/FandIManagerRanking/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'SalesManagerRanking', loadComponent: () => import('./Reports/Rankings/SalesManangerRanking/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'ServiceAdvisorRanking', loadComponent: () => import('./Reports/Rankings/ServiceAdvisorRanking/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'PartsCounterPersonRanking', loadComponent: () => import('./Reports/Rankings/PartsCounterPersonRanking/dashboard/dashboard').then(m => m.Dashboard) },




];

