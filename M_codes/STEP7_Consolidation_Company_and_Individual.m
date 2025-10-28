// ============================================================
// PROJECT: NOBO List Cleaning and Consolidation
// STEP 7 - Final Consolidation of Company and Individual Records
// ============================================================
// Author: Jaqueline F. Filogonio
// Tool: Microsoft Power Query (Excel)
// Description:
// This final step consolidates both company-level and individual-level 
// investor datasets into a single unified table.
//
// Key Objectives:
// - Combine the refined datasets from STEP 5 (Company) and STEP 6 (Individual).
// - Ensure structural consistency across both tables.
// - Produce the final, standardized dataset for analysis and reporting.
//
// Note:
// This script contains only transformation logic and structural references.
// No real or confidential investor data are included.
// ============================================================

let
    // ============================================================
    // STEP 1 — Combine company and individual datasets
    // ============================================================
    Source = Table.Combine({
        Investor_Expanded_Company_STEP5,
        Investor_Expanded_Individual_STEP6
    })

in
    Source