// ============================================================
// PROJECT: NOBO List Cleaning and Consolidation
// STEP 5 - Company-Level Refinement and Structure Alignment
// ============================================================
// Author: Jaqueline F. Filogonio
// Tool: Microsoft Power Query (Excel)
// Description:
// This step continues from STEP 4 (Investor_Expanded_General_STEP4) and
// focuses on refining and isolating company-type investor records.
//
// Key Objectives:
// - Filter company-type investors identified in previous step.
// - Generate a new composite identifier (ROW_ID_COMPOSITE_3) 
//   for company-level traceability.
// - Prepare a clean, standardized schema for company analysis.
// - Add placeholder fields (e.g., TITLE_EXTRACTED) for future enrichment.
//
// Note:
// This script contains only transformation logic and structural references.
// No real or confidential investor data are included.
// ============================================================

let
    // ============================================================
    // STEP 1 — Load data from previous step
    // ============================================================
    Source = Investor_Expanded_General_STEP4,

    // ============================================================
    // STEP 2 — Filter only company-type investor rows
    // ============================================================
    FilteredCompanyRows = Table.SelectRows(
        Source,
        each ([INVESTOR_TYPE] = "COMPANY")
    ),

    // ============================================================
    // STEP 3 — Add composite index for hierarchical traceability
    // ============================================================
    AddCompositeIndex = Table.AddColumn(
        FilteredCompanyRows,
        "ROW_ID_COMPOSITE_3",
        each [ROW_ID_COMPOSITE_2] & "_1",
        type text
    ),

    // ============================================================
    // STEP 4 — Remove obsolete identifiers
    // ============================================================
    RemovedColumns = Table.RemoveColumns(
        AddCompositeIndex,
        {"ROW_ID_COMPOSITE_2"}
    ),

    // ============================================================
    // STEP 5 — Add placeholder column for title extraction
    // ============================================================
    AddNullTitleColumn = Table.AddColumn(
        RemovedColumns,
        "TITLE_EXTRACTED",
        each null,
        type any
    ),

    // ============================================================
    // STEP 6 — Enforce data types and reorder columns
    // ============================================================
    ChangedType = Table.TransformColumnTypes(
        AddNullTitleColumn,
        {{"TITLE_EXTRACTED", type text}}
    ),

    ReorderedColumns = Table.ReorderColumns(
        ChangedType,
        {
            "ROW_ID_COMPOSITE_3",
            "INVESTOR_NAME",
            "INVESTOR_TYPE",
            "COMPANY_TAG",
            "TITLE_EXTRACTED"
        }
    ),

    // ============================================================
    // STEP 7 — Rename columns for final schema alignment
    // ============================================================
    RenamedColumns = Table.RenameColumns(
        ReorderedColumns,
        {
            {"INVESTOR_NAME", "INVESTOR"},
            {"INVESTOR_TYPE", "INVESTOR_TYPE_FINAL"}
        }
    )

in
    RenamedColumns
