// ============================================================
// PROJECT: NOBO List Cleaning and Consolidation
// STEP 4 - Investor Type Classification and Company Detection
// ============================================================
// Author: Jaqueline F. Filogonio
// Tool: Microsoft Power Query (Excel)
// Description:
// This step continues from STEP 3 (Investor_Expanded_General_STEP3) and 
// performs the final normalization and investor type classification process.
//
// Key Objectives:
// - Split investor names into multiple levels of granularity.
// - Normalize connectors like “AND” and “OR” to unify name parsing.
// - Create hierarchical composite identifiers (ROW_ID_COMPOSITE_2).
// - Detect company-related entities based on keywords and patterns.
// - Tag investor records as “COMPANY” or “INDIVIDUAL” for subsequent
//   analytical segmentation.
//
// Note:
// This script contains only transformation logic and structural references.
// No real or confidential investor data are included.
// ============================================================

let
    // ============================================================
    // STEP 1 — Load data from previous step
    // ============================================================
    Source = Investor_Expanded_General_STEP3,

    // ============================================================
    // STEP 2 — Split into individual names + cleaning
    // Creates a new level of split using normalized connectors
    // ============================================================
    connectors = {" AND/OR ", " AND OR ", " ANDOR ", " AND ", " OR "},

    ReplaceConnectors = (text as text, list as list) as text =>
        List.Accumulate(
            list,
            Text.Upper(text),
            (state, current) => Text.Replace(state, current, "|")
        ),

    NormalizedConnectors = Table.TransformColumns(
        Source,
        {{"INVESTOR_NAME_LIST", each ReplaceConnectors(_, connectors), type text}}
    ),

    SplitList = Table.AddColumn(
        NormalizedConnectors,
        "INVESTOR_NAME_LIST_SPLIT",
        each Text.Split([INVESTOR_NAME_LIST], "|"),
        type list
    ),

    ExpandRows = Table.ExpandListColumn(SplitList, "INVESTOR_NAME_LIST_SPLIT"),
    ConvertToText = Table.TransformColumnTypes(
        ExpandRows,
        {{"INVESTOR_NAME_LIST_SPLIT", type text}}
    ),

    // ============================================================
    // STEP 3 — Add the investor index for traceability
    // ============================================================
    GroupedRows = Table.Group(
        ConvertToText,
        {"ROW_ID_COMPOSITE_1"},
        {
            {"AllRows", each Table.AddIndexColumn(_, "INVESTOR_INDEX_LEVEL_2", 1, 1, Int64.Type)}
        }
    ),

    ExpandedRows = Table.ExpandTableColumn(
        GroupedRows,
        "AllRows",
        {"INVESTOR_NAME_LIST_SPLIT", "INVESTOR_INDEX_LEVEL_2"}
    ),

    ConvertedIndex = Table.TransformColumnTypes(
        ExpandedRows,
        {{"INVESTOR_INDEX_LEVEL_2", type text}}
    ),

    // ============================================================
    // STEP 4 — Add composite key for hierarchical tracking
    // ============================================================
    AddCompositeKey = Table.AddColumn(
        ConvertedIndex,
        "ROW_ID_COMPOSITE_2",
        each Text.From([ROW_ID_COMPOSITE_1]) & "_" & [INVESTOR_INDEX_LEVEL_2],
        type text
    ),

    // ============================================================
    // STEP 5 — Clean and normalize investor names
    // ============================================================
    CleanedNames = Table.TransformColumns(
        AddCompositeKey,
        {{"INVESTOR_NAME_LIST_SPLIT", each Text.Remove(Text.Trim(_), {"."}), type text}}
    ),

    // ============================================================
    // STEP 6 — Identify company vs. individual investors
    // ============================================================
    CompanyKeywords = {
        "INC", "LTD", "CORP", "ENTERPRISES", "INVESTMENTS", "LLP",
        "L P", "FOUND", "CAPITAL", "CORPORATION", "FOUNDATION",
        "VALUE", "FUND", "LP"
    },

    AddCompanyFlag = Table.AddColumn(
        CleanedNames,
        "IS_COMPANY",
        each
            let
                nameUpper = Text.Upper([INVESTOR_NAME_LIST_SPLIT]),
                nameParts = Text.Split(nameUpper, " "),
                hasKeyword = List.AnyTrue(
                    List.Transform(CompanyKeywords, (kw) => List.Contains(nameParts, kw))
                ),
                isShortCode =
                    Text.Length(nameUpper) <= 8 and
                    List.AnyTrue(List.Transform({"0".."9"}, each Text.Contains(nameUpper, _))) and
                    List.AnyTrue(List.Transform({"A".."Z"}, each Text.Contains(nameUpper, _)))
            in
                hasKeyword or isShortCode,
        type logical
    ),

    // ============================================================
    // STEP 7 — Extract matched company keyword (if any)
    // ============================================================
    AddCompanyTag = Table.AddColumn(
        AddCompanyFlag,
        "COMPANY_TAG",
        each
            let
                words = Text.Split(Text.Upper([INVESTOR_NAME_LIST_SPLIT]), " "),
                match = List.First(List.Intersect({words, CompanyKeywords}), null)
            in
                match,
        type text
    ),

    // ============================================================
    // STEP 8 — Final classification by investor type
    // ============================================================
    AddInvestorType = Table.AddColumn(
        AddCompanyTag,
        "INVESTOR_TYPE",
        each if [IS_COMPANY] then "COMPANY" else "INDIVIDUAL",
        type text
    ),

    // ============================================================
    // STEP 9 — Cleanup and reorder columns
    // ============================================================
    RemovedColumns = Table.RemoveColumns(
        AddInvestorType,
        {"ROW_ID_COMPOSITE_1", "INVESTOR_INDEX_LEVEL_2", "IS_COMPANY"}
    ),
    ReorderedColumns = Table.ReorderColumns(
        RemovedColumns,
        {"ROW_ID_COMPOSITE_2", "INVESTOR_NAME_LIST_SPLIT", "COMPANY_TAG", "INVESTOR_TYPE"}
    ),
    RenamedColumns = Table.RenameColumns(
        ReorderedColumns,
        {{"INVESTOR_NAME_LIST_SPLIT", "INVESTOR_NAME"}}
    )

in
    RenamedColumns
