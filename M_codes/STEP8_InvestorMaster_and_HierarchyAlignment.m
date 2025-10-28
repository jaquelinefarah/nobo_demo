// ============================================================
// PROJECT: NOBO List Cleaning and Consolidation
// STEP 8 - Investor Master Construction and Hierarchical Alignment
// ============================================================
// Author: Jaqueline F. Filogonio
// Tool: Microsoft Power Query (Excel)
// Description:
// This step continues from STEP 7 (Investor_Expanded_All_STEP7) and performs
// hierarchical decomposition of composite IDs, transposition of investor name
// levels, and creation of the consolidated “Investor Master” table.
//
// Key Objectives:
// - Decompose hierarchical composite identifiers (ROW_ID_COMPOSITE_3).
// - Reconstruct relational investor position indexes.
// - Dynamically pivot and align multiple investor name fields by position.
// - Build a unified “INVESTOR_MASTER” field aggregating all investor aliases.
// - Merge investor type and account type metadata for final integration.
//
// Note:
// This script contains only transformation logic and structural references.
// No real or confidential investor data are included.
// ============================================================

let
    // ============================================================
    // STEP 1 — Load data from previous step
    // ============================================================
    Source = Investor_Expanded_All_STEP7,

    // ============================================================
    // STEP 2 — Split composite hierarchical ID into levels
    // ============================================================
    SplitComposite = Table.SplitColumn(
        Source,
        "ROW_ID_COMPOSITE_3",
        Splitter.SplitTextByDelimiter("_", QuoteStyle.Csv),
        {"INVESTOR_ROW_ID", "ROW_ID_LEVEL_1", "ROW_ID_LEVEL_2", "ROW_ID_LEVEL_3"}
    ),

    // ============================================================
    // STEP 3 — Add incremental index for traceability
    // ============================================================
    AddIndex = Table.AddIndexColumn(SplitComposite, "INDEX", 1, 1, Int64.Type),

    // ============================================================
    // STEP 4 — Convert hierarchy levels to numeric types
    // ============================================================
    ConvertedLevels = Table.TransformColumnTypes(
        SplitComposite,
        {
            {"ROW_ID_LEVEL_1", Int64.Type},
            {"ROW_ID_LEVEL_2", Int64.Type},
            {"ROW_ID_LEVEL_3", Int64.Type}
        }
    ),

    // ============================================================
    // STEP 5 — Calculate investor positional hierarchy
    // ============================================================
    AddInvestorPosition = Table.AddColumn(
        ConvertedLevels,
        "INVESTOR_POSITION",
        each 
            [ROW_ID_LEVEL_1] +
            (if [ROW_ID_LEVEL_2] > 1 then 1 else 0) +
            (if [ROW_ID_LEVEL_3] > 1 then 1 else 0),
        Int64.Type
    ),

    // ============================================================
    // STEP 6 — Convert position to text type for pivot operations
    // ============================================================
    ChangedType = Table.TransformColumnTypes(
        AddInvestorPosition,
        {{"INVESTOR_POSITION", type text}}
    ),

    // ============================================================
    // STEP 7 — Ensure investor name field type consistency
    // ============================================================
    EnsureTextType = Table.TransformColumnTypes(
        ChangedType,
        {{"INVESTOR", type text}}
    ),

    // ============================================================
    // STEP 8 — Group and prepare for transposition
    // ============================================================
    Grouped = Table.Group(
        EnsureTextType,
        {"INVESTOR_ROW_ID"},
        {{"AllRows", each _, type table}}
    ),

    Expanded = Table.ExpandTableColumn(
        Grouped,
        "AllRows",
        {"INVESTOR", "INVESTOR_POSITION"}
    ),

    // ============================================================
    // STEP 9 — Create dynamic attribute names by investor position
    // ============================================================
    AddAttribute = Table.AddColumn(
        Expanded,
        "Attribute_Indexed",
        each "INVESTOR_" & Text.From([INVESTOR_POSITION]),
        type text
    ),

    // ============================================================
    // STEP 10 — Replace nulls with blanks to avoid pivot issues
    // ============================================================
    CleanNames = Table.TransformColumns(
        AddAttribute,
        {{"INVESTOR", each if _ = null then "" else _, type text}}
    ),

    // ============================================================
    // STEP 11 — Prepare attribute list for dynamic pivoting
    // ============================================================
    AttributeList = List.Distinct(CleanNames[Attribute_Indexed]),

    // ============================================================
    // STEP 12 — Pivot investor names by position
    // ============================================================
    PivotNames = Table.Pivot(
        Table.RemoveColumns(CleanNames, {"INVESTOR_POSITION"}),
        AttributeList,
        "Attribute_Indexed",
        "INVESTOR"
    ),

    // ============================================================
    // STEP 13 — Build consolidated investor master name
    // ============================================================
    AddInvestorMaster = Table.AddColumn(
        PivotNames,
        "INVESTOR_MASTER",
        each 
            Text.Combine(
                List.Select(
                    Record.ToList(
                        Record.SelectFields(
                            _,
                            List.Select(
                                Record.FieldNames(_),
                                each Text.StartsWith(_, "INVESTOR_") and _ <> "INVESTOR_ROW_ID"
                            )
                        )
                    ),
                    each _ <> null and _ <> ""
                ),
                " | "
            ),
        type text
    ),

    // ============================================================
    // STEP 14 — Merge Account Type metadata
    // ============================================================
    Merged_AccountType = Table.NestedJoin(
        AddInvestorMaster,
        {"INVESTOR_ROW_ID"},
        Account_Type,
        {"INVESTOR_ROW_ID"},
        "Account_Type",
        JoinKind.LeftOuter
    ),

    Expanded_AccountType = Table.ExpandTableColumn(
        Merged_AccountType,
        "Account_Type",
        {"ACCOUNT_TYPE_CONSOLIDATED"},
        {"ACCOUNT_TYPE_CONSOLIDATED"}
    ),

    // ============================================================
    // STEP 15 — Merge Investor Type metadata
    // ============================================================
    Merged_InvestorType = Table.NestedJoin(
        Expanded_AccountType,
        {"INVESTOR_ROW_ID"},
        Investor_Type,
        {"INVESTOR_ROW_ID"},
        "Investor_Type",
        JoinKind.LeftOuter
    ),

    Expanded_InvestorType = Table.ExpandTableColumn(
        Merged_InvestorType,
        "Investor_Type",
        {"INVESTOR_TYPE"},
        {"INVESTOR_TYPE.1"}
    )

in
    Expanded_InvestorType
