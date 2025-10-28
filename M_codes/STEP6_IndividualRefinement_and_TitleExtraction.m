// ============================================================
// PROJECT: NOBO List Cleaning and Consolidation
// STEP 6 - Individual Investor Refinement and Title Extraction
// ============================================================
// Author: Jaqueline F. Filogonio
// Tool: Microsoft Power Query (Excel)
// Description:
// This step continues from STEP 4 (Investor_Expanded_General_STEP4) and
// focuses on refining individual investor records, handling naming patterns
// and extracting honorific titles.
//
// Key Objectives:
// - Filter individual-type investors identified in previous steps.
// - Normalize encoded characters and split composite names by “&”.
// - Generate hierarchical composite identifiers (ROW_ID_COMPOSITE_3).
// - Detect and extract honorific titles (e.g., MR, MS, DR, MRS, SIR).
// - Produce a clean, standardized list of individual investor names.
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
    // STEP 2 — Filter only individual-type investor rows
    // ============================================================
    Filtered_Individual_Rows = Table.SelectRows(
        Source,
        each [INVESTOR_TYPE] = "INDIVIDUAL"
    ),

    // ============================================================
    // STEP 3 — Normalize encoded ampersands and enforce text type
    // ============================================================
    NormalizeAmp = Table.TransformColumns(
        Filtered_Individual_Rows,
        {{"INVESTOR_NAME", each Text.Replace(Text.From(_), "&amp;", "&"), type text}}
    ),

    // ============================================================
    // STEP 4 — Split names by ampersand (“&”) and trim whitespace
    // ============================================================
    AddSplitList = Table.AddColumn(
        NormalizeAmp,
        "INVESTOR_NAME_FINAL",
        each List.Transform(Text.Split([INVESTOR_NAME], "&"), Text.Trim),
        type list
    ),

    Expanded_NAME_SPLIT_LIST = Table.ExpandListColumn(
        AddSplitList,
        "INVESTOR_NAME_FINAL"
    ),

    ChangedType = Table.TransformColumnTypes(
        Expanded_NAME_SPLIT_LIST,
        {{"INVESTOR_NAME_FINAL", type text}}
    ),

    // ============================================================
    // STEP 5 — Add investor index for hierarchical tracking
    // ============================================================
    GroupedRows = Table.Group(
        ChangedType,
        {"ROW_ID_COMPOSITE_2"},
        {
            {"AllRows", each Table.AddIndexColumn(_, "INVESTOR_INDEX_LEVEL_3", 1, 1, Int64.Type)}
        }
    ),

    ExpandedRows = Table.ExpandTableColumn(
        GroupedRows,
        "AllRows",
        {"INVESTOR_NAME_FINAL", "INVESTOR_INDEX_LEVEL_3"}
    ),

    // ============================================================
    // STEP 6 — Convert index to text
    // ============================================================
    ConvertedIndex = Table.TransformColumnTypes(
        ExpandedRows,
        {
            {"INVESTOR_INDEX_LEVEL_3", type text},
            {"INVESTOR_NAME_FINAL", type text}
        }
    ),

    // ============================================================
    // STEP 7 — Add composite identifier
    // ============================================================
    AddCompositeKey = Table.AddColumn(
        ConvertedIndex,
        "ROW_ID_COMPOSITE_3",
        each Text.From([ROW_ID_COMPOSITE_2]) & "_" & [INVESTOR_INDEX_LEVEL_3],
        type text
    ),

    // ============================================================
    // STEP 8 — Remove intermediate identifiers
    // ============================================================
    RemovedColumns = Table.RemoveColumns(
        AddCompositeKey,
        {"ROW_ID_COMPOSITE_2", "INVESTOR_INDEX_LEVEL_3"}
    ),

    // ============================================================
    // STEP 9 — Extract honorific titles
    // Titles sorted by length to avoid partial matches
    // ============================================================
    AddTitleColumn = Table.AddColumn(
        RemovedColumns,
        "TITLE_EXTRACTED",
        each 
            let
                name = [INVESTOR_NAME_FINAL],
                upperName = Text.Upper(name),
                titles = {"MADAM", "MISS", "MRS", "DR", "SIR", "MS", "MR"},
                found = List.First(
                    List.Select(titles, each Text.StartsWith(upperName, _)),
                    null
                )
            in
                found,
        type text
    ),

    // ============================================================
    // STEP 10 — Remove title prefix from investor name
    // ============================================================
    AddNameWithoutTitle = Table.AddColumn(
        AddTitleColumn,
        "INVESTOR_NAME_FINAL_NO_TITLE",
        each 
            if [TITLE_EXTRACTED] <> null 
            then Text.Trim(Text.Range([INVESTOR_NAME_FINAL], Text.Length([TITLE_EXTRACTED]) + 1)) 
            else [INVESTOR_NAME_FINAL],
        type text
    ),

    // ============================================================
    // STEP 11 — Add placeholder columns and finalize structure
    // ============================================================
    AddNullTagCompanyColumn = Table.AddColumn(
        AddNameWithoutTitle,
        "COMPANY_TAG",
        each null,
        type text
    ),

    AddInvestorTypeFinal = Table.AddColumn(
        AddNullTagCompanyColumn,
        "INVESTOR_TYPE_FINAL",
        each "INDIVIDUAL",
        type text
    ),

    RemovedHelperColumns = Table.RemoveColumns(
        AddInvestorTypeFinal,
        {"INVESTOR_NAME_FINAL"}
    ),

    ReorderedColumns = Table.ReorderColumns(
        RemovedHelperColumns,
        {
            "ROW_ID_COMPOSITE_3",
            "INVESTOR_NAME_FINAL_NO_TITLE",
            "INVESTOR_TYPE_FINAL",
            "COMPANY_TAG",
            "TITLE_EXTRACTED"
        }
    ),

    RenamedColumns = Table.RenameColumns(
        ReorderedColumns,
        {{"INVESTOR_NAME_FINAL_NO_TITLE", "INVESTOR"}}
    )

in
    RenamedColumns
