// ============================================================
// PROJECT: NOBO List Cleaning and Consolidation
// STEP 2 - Investor Name Decomposition and Classification
// ============================================================
// Author: Jaqueline F. Filogonio
// Tool: Microsoft Power Query (Excel)
// Description:
// This step continues from STEP 1 (Investor_Unique_Row) and performs 
// the decomposition and classification of investor name fragments.
//
// Key Objectives:
// - Split consolidated investor names ("TEMPORARY_INVESTOR_NAME_FINAL")
//   into multiple rows for granular analysis.
// - Identify and flag special markers (e.g., "C/O", "A/S", "RRSP", "TRUST").
// - Separate representative names (e.g., "ATTN", "C/O") from investor names.
// - Detect and remove account types (e.g., "RRSP", "TFSA", "ESTATE").
// - Prepare cleaned and structured name parts for further entity 
//   classification in STEP 3.
//
// Note:
// This script contains only transformation logic and structural references. 
// No real or confidential investor data are included.
// ============================================================


let

// ==============================================
// STEP 0: LOAD SOOURCE
// This query is a continuite from the Investor_Base_Unique_Row
// ==============================================
    
Source = Investor_Unique_Row_STEP1,

// ============================================================
// STEP 1: SPLIT CONSOLIDATED NAME FIELD INTO MULTIPLE ROWS
// Description: Creates a row for each item separated by " | "
//              in TEMPORARY_INVESTOR_NAME_FINAL. 
//              Preserves source row identity using INVESTOR_ROW_ID.
//
// Reason: Enables precise classification of mixed content like 
//         "Company", "Representative", or "Investor Name"
//         which may coexist in the same record.
//
// Output: Each fragment becomes one row in TEMPORARY_INVESTOR_NAME_PART_0
// ============================================================


// Split the final investor name column into a list using " | " separator
AddSplitColumn = Table.AddColumn(Source, "TEMPORARY_INVESTOR_NAME_PART_0", each 
    Text.Split([TEMPORARY_INVESTOR_NAME_FINAL], " | "), type list),

// Expand list into multiple rows — each row is one name fragment
ExpandedInvestorParts = Table.ExpandListColumn(AddSplitColumn, "TEMPORARY_INVESTOR_NAME_PART_0"),



// ==============================================
// STEP 1: ADD THE INVESTOR_INDEX
// Add the second level of the identificator rows
// ==============================================
    
Grouped = Table.Group(
    ExpandedInvestorParts,
    {"INVESTOR_ROW_ID"},
    {
        {"AllRows", each Table.AddIndexColumn(_, "INVESTOR_INDEX", 1, 1, Int64.Type)}
    }
),

// Step — Expand grouped table
Expanded = Table.ExpandTableColumn(
    Grouped,
    "AllRows",
    {"TEMPORARY_INVESTOR_NAME_PART_0", "INVESTOR_INDEX"}
),


ConvertedIndex = Table.TransformColumnTypes(Expanded, {{"INVESTOR_INDEX", type text}}),

// ==============================================
// STEP 2: ADD THE INVESTOR_INDEX
// FIND A SPECIAL CARACTER
// ==============================================

// === STEP: Extract special markers from investor name ===

SPECIAL_CODES = {
    "4E9", "4F3", "2C1", "A/S", "C/O", "CO", "TR", "RR", "ITF", "FBO", "U/A", "U/T", "**", "*"
},
    AddSpecialMarkerAfter = Table.AddColumn(
    ConvertedIndex,
    "SPECIAL_MARKER_AFTER_NAME",
    each
        let
            txt = Text.Upper(Text.Trim([TEMPORARY_INVESTOR_NAME_PART_0])),
            parts = Text.Split(txt, " "),
            last = if List.Count(parts) > 0 then List.Last(parts) else null,

            // Regra 1: está na lista conhecida
            isListedCode = List.Contains(SPECIAL_CODES, last),

            // Regra 2: código curto com caracteres misturados (incluindo '*')
            isCodePattern =
                Text.Length(last) <= 5 and (
                    (
                        Text.Length(Text.Select(last, {"A".."Z"})) > 0 and 
                        Text.Length(Text.Select(last, {"0".."9"})) > 0
                    )
                    or Text.Contains(last, "*")
                ),

            // Novo filtro: se for tudo letras e números, mas com 5+ caracteres, evita classificar como código
            isCorporateLike = Text.Length(last) > 4 and Text.Length(Text.Select(last, {"A".."Z","0".."9"})) = Text.Length(last),

            result = if (isListedCode or isCodePattern) and not isCorporateLike then last else null
        in
            result,
    type nullable text
),

AddSpecialBefore = Table.AddColumn(
    AddSpecialMarkerAfter,
    "SPECIAL_MARKER_BEFORE_NAME",
    each
        let
            txt = Text.Upper(Text.Trim([TEMPORARY_INVESTOR_NAME_PART_0]))
        in
            if Text.StartsWith(txt, ">") then ">" else null,
    type nullable text
),

// ============================================================
// STEP 3: CLEAN SPECIAL CHARACTERS FROM INVESTOR NAME PART
// Description: Creates TEMPORARY_INVESTOR_NAME_PART_1 by 
//              removing ">" prefix and "**" markers like "50**".
//              Preserves the original in PART_0 for traceability.
//
// Output: Cleaned name string in STEP 2
// ============================================================

AddCleanedInvestorName = Table.AddColumn(
    AddSpecialBefore,
    "TEMPORARY_INVESTOR_NAME_PART_1",
    each 
        let
            original = [TEMPORARY_INVESTOR_NAME_PART_0],
            trimmed = Text.Trim(original),

            // Remove marcador de prefixo ">"
            removePrefix = if [SPECIAL_MARKER_BEFORE_NAME] <> null then Text.Middle(trimmed, 1) else trimmed,

            // Remove última palavra se for um marcador identificado
            words = Text.Split(removePrefix, " "),
            lastWord = if List.Count(words) > 0 then List.Last(words) else null,
            cleanedWords = if lastWord <> null and lastWord = [SPECIAL_MARKER_AFTER_NAME]
                           then List.RemoveLastN(words, 1)
                           else words,

            result = Text.Trim(Text.Combine(cleanedWords, " "))
        in
            result,
    type nullable text
),


// ==============================================
// STEP 4: EXTRACT REPRESENTATIVE FROM NAME PART 1
// ==============================================

AddWordList = Table.AddColumn(AddCleanedInvestorName, "WORD_LIST", each
    if [TEMPORARY_INVESTOR_NAME_PART_1] = null then {} 
    else 
        Text.Split(
            Text.Upper(
                Text.Remove(
                    Text.Trim([TEMPORARY_INVESTOR_NAME_PART_1]), 
                    {":", ",", ".", "(", ")", "-", "_"}
                )
            ), 
            " "
        ), 
    type list
),

AddIsRepresentative = Table.AddColumn(AddWordList, "IS_REPRESENTATIVE", each 
    List.Count(
        List.Intersect({
            [WORD_LIST], 
            {"ATTN", "C/O", "CO", "FOR", "A/S", "IN TRUST FOR", "AS", "ATT"}
        })
    ) > 0, 
    type logical
),

AddRepresentative = Table.AddColumn(AddIsRepresentative, "REPRESENTATIVE", each 
    if [IS_REPRESENTATIVE] then [TEMPORARY_INVESTOR_NAME_PART_1] else null, 
    type nullable text
),

AddTempPart2 = Table.AddColumn(AddRepresentative, "TEMPORARY_INVESTOR_NAME_PART_2", each 
    if [IS_REPRESENTATIVE] then null else [TEMPORARY_INVESTOR_NAME_PART_1], 
    type nullable text
),

RemoveRepHelperCols = Table.RemoveColumns(AddTempPart2, {"WORD_LIST"}),

// ==============================================
// STEP 5: CLASSIFY AND REMOVE ACCOUNT TYPE
// ==============================================

ACCOUNT_TYPES = {
    "JTWROS", "JT/WROS", "SPOUSAL", "SPOUSAL PLAN", 
    "TRUST", "RRSP", "TFSA", "ESTATE OF", "ESTATE", 
    "DIVIDEND REINVESTMENT", "REINVESTMENT", "DIVIDEND"
},

AddCleanForMatch = Table.AddColumn(RemoveRepHelperCols, "CLEANED_FOR_MATCH", each 
    if [TEMPORARY_INVESTOR_NAME_PART_2] = null then null 
    else 
        Text.Upper(
            Text.Remove([TEMPORARY_INVESTOR_NAME_PART_2], {"(", ")", ".", ",", "-", "_"})
        ), 
    type nullable text
),

AddIsAcctType = Table.AddColumn(AddCleanForMatch, "IS_ACCOUNT_TYPE", each 
    List.AnyTrue(List.Transform(ACCOUNT_TYPES, (kw) => Text.Contains([CLEANED_FOR_MATCH], kw))), 
    type logical
),

AddAcctType = Table.AddColumn(AddIsAcctType, "ACCOUNT_TYPE", each 
    Text.Combine(
        List.Select(ACCOUNT_TYPES, (kw) => Text.Contains([CLEANED_FOR_MATCH], kw)), 
        ", "
    ), 
    type nullable text
),

CleanNamePart3 = Table.AddColumn(AddAcctType, "TEMPORARY_INVESTOR_NAME_PART_3", each 
    if [TEMPORARY_INVESTOR_NAME_PART_2] = null then null 
    else
        let
            original = Text.Upper([TEMPORARY_INVESTOR_NAME_PART_2]),
            cleaned = List.Accumulate(ACCOUNT_TYPES, original, (state, kw) => Text.Replace(state, kw, "")),
            final = Text.Trim(cleaned)
        in
            final,
    type nullable text
)
in
    CleanNamePart3