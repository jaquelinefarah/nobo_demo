// ============================================================
// PROJECT: NOBO List Cleaning and Consolidation
// STEP 3 - Investor Name Normalization and Aggregation
// ============================================================
// Author: Jaqueline F. Filogonio
// Tool: Microsoft Power Query (Excel)
// Description:
// This step continues from STEP 2 (Investor_Expanded_General_STEP2) and
// performs normalization, aggregation, and indexation of investor names.
//
// Key Objectives:
// - Aggregate split name fragments into unified investor name strings.
// - Normalize connectors like "AND/OR" and "OR" into consistent separators.
// - Expand concatenated names into multiple rows for analysis.
// - Create hierarchical row identifiers (ROW_ID_COMPOSITE_1) 
//   for traceability and relational integrity.
// - Prepare a clean and standardized investor name list for subsequent
//   entity classification and consolidation steps.
//
// Note:
// This script contains only transformation logic and structural references.
// No real or confidential investor data are included.
// ============================================================

let
    // ============================================================
    // STEP 1 — Ensure INVESTOR_NAME is of type text
    // ============================================================
    Source = Investor_Expanded_General_STEP2,
    EnsureTextType = Table.TransformColumnTypes(
        Source,
        {{"TEMPORARY_INVESTOR_NAME_PART_3", type text}}
    ),

    // ============================================================
    // STEP 2 — Group by INVESTOR_ROW_ID
    // ============================================================
    Grouped = Table.Group(
        EnsureTextType,
        {"INVESTOR_ROW_ID"},
        {{"AllRows", each _, type table}}
    ),

    // ============================================================
    // STEP 3 — Expand INVESTOR_NAME and INVESTOR_POSITION
    // ============================================================
    Expanded = Table.ExpandTableColumn(
        Grouped,
        "AllRows",
        {"TEMPORARY_INVESTOR_NAME_PART_3", "INVESTOR_INDEX"}
    ),

    // ============================================================
    // STEP 4 — Create dynamic attribute name based on investor position
    // ============================================================
    AddAttribute = Table.AddColumn(
        Expanded,
        "Attribute_Indexed",
        each "TEMPORARY_INVESTOR_NAME_PART_3" & Text.From([INVESTOR_INDEX]),
        type text
    ),

    // ============================================================
    // STEP 5 — Replace nulls with empty text to avoid pivot errors
    // ============================================================
    CleanNames = Table.TransformColumns(
        AddAttribute,
        {{"TEMPORARY_INVESTOR_NAME_PART_3", each if _ = null then "" else _, type text}}
    ),

    // ============================================================
    // STEP 6 — Get list of unique attribute names to use as pivot columns
    // ============================================================
    AttributeList = List.Distinct(CleanNames[Attribute_Indexed]),

    // ============================================================
    // STEP 7 — Pivot to transpose investor names by position
    // ============================================================
    PivotNames = Table.Pivot(
        Table.RemoveColumns(CleanNames, {"INVESTOR_INDEX"}),
        AttributeList,
        "Attribute_Indexed",
        "TEMPORARY_INVESTOR_NAME_PART_3"
    ),

    // ============================================================
    // STEP 8 — Combine names into a single cleaned string
    // ============================================================
    AddCombinedName = Table.AddColumn(
        PivotNames,
        "TEMPORARY_INVESTOR_NAME_AGREGATED",
        each Text.Combine(
            List.RemoveNulls({
                [TEMPORARY_INVESTOR_NAME_PART_31],
                [TEMPORARY_INVESTOR_NAME_PART_32],
                [TEMPORARY_INVESTOR_NAME_PART_33]
            }),
            " "
        ),
        type text
    ),

    // ============================================================
    // STEP 9 — Define connectors to normalize (e.g., AND/OR)
    // ============================================================
    connectors = {" AND/O ", " AND/OR ", " AND OR ", " AND / OR ", " ANDOR ", " OR "},

    // ============================================================
    // STEP 10 — Replace connectors like "AND/OR" with "|"
    // ============================================================
    ReplaceConnectors = (text as text, list as list) as text =>
        List.Accumulate(
            list,
            Text.Upper(text),
            (state, current) => Text.Replace(state, current, "|")
        ),

    NormalizedConnectors = Table.TransformColumns(
        AddCombinedName,
        {{"TEMPORARY_INVESTOR_NAME_AGREGATED", each ReplaceConnectors(_, connectors), type text}}
    ),

    // ============================================================
    // STEP 11 — Split combined name into individual list elements
    // ============================================================
    SplitList = Table.AddColumn(
        NormalizedConnectors,
        "INVESTOR_NAME_LIST",
        each Text.Split([TEMPORARY_INVESTOR_NAME_AGREGATED], "|"),
        type list
    ),

    // ============================================================
    // STEP 12 — Expand name list into individual rows
    // ============================================================
    ExpandRows = Table.ExpandListColumn(SplitList, "INVESTOR_NAME_LIST"),
    ConvertToText = Table.TransformColumnTypes(ExpandRows, {{"INVESTOR_NAME_LIST", type text}}),

    // ============================================================
    // STEP 13 — Group by INVESTOR_ROW_ID and add sequential index
    // ============================================================
    GroupedRows = Table.Group(
        ConvertToText,
        {"INVESTOR_ROW_ID"},
        {{"AllRows", each Table.AddIndexColumn(_, "INVESTOR_INDEX_LEVEL_1", 1, 1, Int64.Type)}}
    ),

    // ============================================================
    // STEP 14 — Expand grouped table
    // ============================================================
    ExpandedRows = Table.ExpandTableColumn(
        GroupedRows,
        "AllRows",
        {"INVESTOR_NAME_LIST", "INVESTOR_INDEX_LEVEL_1"}
    ),

    // ============================================================
    // STEP 15 — Convert index to text
    // ============================================================
    ConvertedIndex = Table.TransformColumnTypes(
        ExpandedRows,
        {{"INVESTOR_INDEX_LEVEL_1", type text}}
    ),

    // ============================================================
    // STEP 16 — Add composite ID
    // ============================================================
    AddCompositeKey = Table.AddColumn(
        ConvertedIndex,
        "ROW_ID_COMPOSITE_1",
        each Text.From([INVESTOR_ROW_ID]) & "_" & [INVESTOR_INDEX_LEVEL_1],
        type text
    ),

    // ============================================================
    // STEP 17 — Trim and remove dots from each name
    // ============================================================
    CleanedNames = Table.TransformColumns(
        AddCompositeKey,
        {{"INVESTOR_NAME_LIST", each Text.Remove(Text.Trim(_), {"."}), type text}}
    ),

    // ============================================================
    // STEP 18 — Remove helper columns and reorder
    // ============================================================
    RemovedColumns = Table.RemoveColumns(
        CleanedNames,
        {"INVESTOR_ROW_ID", "INVESTOR_INDEX_LEVEL_1"}
    ),
    ReorderedColumns = Table.ReorderColumns(
        RemovedColumns,
        {"ROW_ID_COMPOSITE_1", "INVESTOR_NAME_LIST"}
    )

in
    ReorderedColumns
