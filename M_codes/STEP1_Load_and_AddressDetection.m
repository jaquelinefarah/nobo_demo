let
// ============================================================
// PROJECT: NOBO List Cleaning and Consolidation
// STEP 1 - Raw Data Load & Address Detection
// ============================================================
// Author: Jaqueline F. Filogonio
// Tool: Microsoft Power Query (Excel)
// Description:
// Script demonstrating advanced text cleaning, address detection,
// and investor name reconstruction logic.
// All data references are structural; no real data are included.
// ============================================================

//This step performs advanced cleaning and classification of investor data,
//including:

//- Loading and promoting headers from the original Excel sheet
//- Adding unique `INVESTOR_ROW_ID` for traceability
//- Cleaning 7 address columns using a custom Power Query function `CLEAN_ADDRESS`
//- Detecting valid address patterns with `IS_ADDRESS()`
//- Propagating address flags across columns to identify the most reliable field
//- Reconstructing temporary investor names extracted from non-address cells

// ============================================================
// STEP 1: LOAD EXCEL SHEET
// ============================================================

source = Excel.Workbook(File.Contents("data/Nobo_List_Raw_Sample.xlsx"), null, true)), null, true),
Sheet1_Sheet = source{[Item="Sheet1",Kind="Sheet"]}[Data],
removed_top_rows = Table.Skip(Sheet1_Sheet,5),
promoted_headers = Table.PromoteHeaders(removed_top_rows, [PromoteAllScalars=true]),

// ============================================================
// STEP 2: ADD A UNIQUE ROW ID FOR TRACEABILITY
// ============================================================

// Add a unique row ID for traceability
AddRowID = Table.AddIndexColumn(promoted_headers, "INVESTOR_ROW_ID", 1, 1, Int64.Type),

ReorderedColumns = Table.ReorderColumns(AddRowID,{"INVESTOR_ROW_ID", "NAME", "ADDRESS 1", "ADDRESS 2", "ADDRESS 3", "ADDRESS 4", "ADDRESS 5", "ADDRESS 6", "ADDRESS 7", "POSTAL CODE", "POSTAL REGION", "NOTICE AND ACCESS", "E-MAIL ADDRESS", "LANGUAGE", "NUMBER OF SHARES", "RECEIVE ALL MATERIALS", "AGREE TO ELECTRONIC DELIVERY", "CUID"}),


// ============================================================
// STEP 3: FIX COLUMN TYPES
// ============================================================

typecorrect = Table.TransformColumnTypes(ReorderedColumns, {
    {"NAME", type text}, {"ADDRESS 1", type text}, {"ADDRESS 2", type text}, {"ADDRESS 3", type text},
    {"ADDRESS 4", type text}, {"POSTAL CODE", type text}, {"POSTAL REGION", type text},
    {"NOTICE AND ACCESS", type text}, {"E-MAIL ADDRESS", type text}, {"LANGUAGE", type text},
    {"NUMBER OF SHARES", Int64.Type}, {"RECEIVE ALL MATERIALS", type text},
    {"AGREE TO ELECTRONIC DELIVERY", type text}, {"CUID", type text},
    {"ADDRESS 5", type text}, {"ADDRESS 6", type text}, {"ADDRESS 7", type text}
}),


// ============================================================
// STEP 4: DUPLICATE NAME COLUMN
// ============================================================

duplicatename = Table.DuplicateColumn(typecorrect, "NAME", "TEMPORARY_INVESTOR_NAME_0"),
  

// ============================================================
// FUNCTION: CLEAN_ADDRESS
// Description: Standardizes address by removing accents, symbols, and formatting text
// ============================================================

CLEAN_ADDRESS = (ADDRESS as nullable text) as nullable text =>
    if ADDRESS = null then null else
    let
        UPPERCASE = Text.Upper(ADDRESS),
        WITH_AND = Text.Replace(UPPERCASE, "&", " AND"),
        NO_ACCENTS = Text.Remove(WITH_AND, {"Á".."Ú", "À".."Ù", "Â".."Û", "Ã", "Õ", "É", "Ê", "Í", "Ó", "Ô", "Ç"}),
        VALID_ONLY = Text.Select(NO_ACCENTS, {"A".."Z", "0".."9", " "}),
        WORDS = List.Select(Text.Split(VALID_ONLY, " "), each _ <> ""),
        RESULT = Text.Combine(WORDS, " ")
    in
        RESULT,

// 1. Clean each original address column using the CLEAN_ADDRESS function
    CleanAddress1 = Table.AddColumn(duplicatename, "ADDRESS1", each CLEAN_ADDRESS([ADDRESS 1]), type nullable text),
    CleanAddress2 = Table.AddColumn(CleanAddress1, "ADDRESS2", each CLEAN_ADDRESS([ADDRESS 2]), type nullable text),
    CleanAddress3 = Table.AddColumn(CleanAddress2, "ADDRESS3", each CLEAN_ADDRESS([ADDRESS 3]), type nullable text),
    CleanAddress4 = Table.AddColumn(CleanAddress3, "ADDRESS4", each CLEAN_ADDRESS([ADDRESS 4]), type nullable text),
    CleanAddress5 = Table.AddColumn(CleanAddress4, "ADDRESS5", each CLEAN_ADDRESS([ADDRESS 5]), type nullable text),
    CleanAddress6 = Table.AddColumn(CleanAddress5, "ADDRESS6", each CLEAN_ADDRESS([ADDRESS 6]), type nullable text),
    CleanAddress7 = Table.AddColumn(CleanAddress6, "ADDRESS7", each CLEAN_ADDRESS([ADDRESS 7]), type nullable text),


// ============================================================
// CONSTANTS FOR ADDRESS DETECTION
// ============================================================

ADDRESS_WORDS = {
    "ST", "STREET", "AVE", "AVENUE", "BLVD", "ROAD", "RD", "PL", "PLACE", "DR", "DRIVE", "UNIT",
    "PKWY", "PARKWAY", "LANE", "LN", "CRT", "COURT", "CRES", "CRESCENT", "BAY", "STR",
    "HOUSE", "FLOOR", "WAY", "TERRACE", "TRAIL", "COVE", "APT", "SUITE", "BOX", "RR", "SQUARE", 
    "HONG KONG", "HONG KONG HONG KONG"
},

PROVINCE_CODES = {
    "AB", "BC", "MB", "NB", "NL", "NS", "NT", "NU", "ON", "PE", "QC", "SK", "YT"
},


// ============================================================
// FUNCTION: IS_ADDRESS
// Description: Determines if a cleaned text field likely represents a valid address
// ============================================================

IS_ADDRESS = (TEXT as nullable text) as logical =>
    let
        CLEANED = if TEXT = null then "" else CLEAN_ADDRESS(TEXT),
        WORD_LIST = List.Select(Text.Split(CLEANED, " "), each _ <> ""),
        HAS_NUMBERS = Text.Length(Text.Select(CLEANED, {"0".."9"})) > 0,
        HAS_ADDRESS_KEYWORDS = List.AnyTrue(List.Transform(ADDRESS_WORDS, each List.Contains(WORD_LIST, _))),
        HAS_PROVINCE = List.AnyTrue(List.Transform(PROVINCE_CODES, each List.Contains(WORD_LIST, _))),
        STARTS_WITH_NUMBER =
            let first = List.First(WORD_LIST, null)
            in try Value.Is(Value.FromText(first), type number) otherwise false
    in
        HAS_NUMBERS or HAS_ADDRESS_KEYWORDS or HAS_PROVINCE or STARTS_WITH_NUMBER,

// ============================================================
// STEP 5: APPLY IS_ADDRESS TO EACH ADDRESS FIELD (1–7)
// Creates boolean flags: IS_ADDRESS_1 through IS_ADDRESS_7
// ============================================================

Classify_Adrress1 = Table.AddColumn(CleanAddress7, "IS_ADDRESS_1", each IS_ADDRESS([ADDRESS1]), type logical),
Classify_Adrress2 = Table.AddColumn(Classify_Adrress1, "IS_ADDRESS_2", each IS_ADDRESS([ADDRESS2]), type logical),
Classify_Adrress3= Table.AddColumn(Classify_Adrress2, "IS_ADDRESS_3", each IS_ADDRESS([ADDRESS3]), type logical),
Classify_Adrress4 = Table.AddColumn(Classify_Adrress3, "IS_ADDRESS_4", each IS_ADDRESS([ADDRESS4]), type logical),
Classify_Adrress5 = Table.AddColumn(Classify_Adrress4, "IS_ADDRESS_5", each IS_ADDRESS([ADDRESS5]), type logical),
Classify_Adrress6 = Table.AddColumn(Classify_Adrress5, "IS_ADDRESS_6", each IS_ADDRESS([ADDRESS6]), type logical),
Classify_Adrress7 = Table.AddColumn(Classify_Adrress6, "IS_ADDRESS_7", each IS_ADDRESS([ADDRESS7]), type logical),



// ============================================================
// STEP 8: PROPAGATE TRUE FOR FIRST VALID ADDRESS FOUND
// Description: Once a "Likely Address" is found, all subsequent flags become TRUE
// ============================================================

step1_final = Table.AddColumn(Classify_Adrress7, "IS_ADDRESS_1_FINAL", each [IS_ADDRESS_1]),

step2_final = Table.AddColumn(step1_final, "IS_ADDRESS_2_FINAL", each 
    if [IS_ADDRESS_1_FINAL] then true else [IS_ADDRESS_2]
),

step3_final = Table.AddColumn(step2_final, "IS_ADDRESS_3_FINAL", each 
    if [IS_ADDRESS_2_FINAL] then true else [IS_ADDRESS_3]
),

step4_final = Table.AddColumn(step3_final, "IS_ADDRESS_4_FINAL", each 
    if [IS_ADDRESS_3_FINAL] then true else [IS_ADDRESS_4]
),

step5_final = Table.AddColumn(step4_final, "IS_ADDRESS_5_FINAL", each 
    if [IS_ADDRESS_4_FINAL] then true else [IS_ADDRESS_5]
),

step6_final = Table.AddColumn(step5_final, "IS_ADDRESS_6_FINAL", each 
    if [IS_ADDRESS_5_FINAL] then true else [IS_ADDRESS_6]
),

step7_final = Table.AddColumn(step6_final, "IS_ADDRESS_7_FINAL", each 
    if [IS_ADDRESS_6_FINAL] then true else [IS_ADDRESS_7]
),

// ============================================================
// STEP 9: EXTRACT TEMPORARY INVESTOR NAMES FROM ADDRESS COLUMNS
// Description: Creates TEMPORARY_INVESTOR_NAME_1 to 7 by 
//              isolating non-address content from ADDRESS 1–7 
//              using IS_ADDRESS_X_FINAL logic
// Output: Returns ADDRESSX value only when IS_ADDRESS_X_FINAL = false
// ============================================================

AddTempInvestorName1 = Table.AddColumn(step7_final, "TEMPORARY_INVESTOR_NAME_1", each 
    if [IS_ADDRESS_1_FINAL] then null else [ADDRESS 1], type nullable text),

AddTempInvestorName2 = Table.AddColumn(AddTempInvestorName1, "TEMPORARY_INVESTOR_NAME_2", each 
    if [IS_ADDRESS_2_FINAL] then null else [ADDRESS 2], type nullable text),

AddTempInvestorName3 = Table.AddColumn(AddTempInvestorName2, "TEMPORARY_INVESTOR_NAME_3", each 
    if [IS_ADDRESS_3_FINAL] then null else [ADDRESS 3], type nullable text),

AddTempInvestorName4 = Table.AddColumn(AddTempInvestorName3, "TEMPORARY_INVESTOR_NAME_4", each 
    if [IS_ADDRESS_4_FINAL] then null else [ADDRESS 4], type nullable text),

AddTempInvestorName5 = Table.AddColumn(AddTempInvestorName4, "TEMPORARY_INVESTOR_NAME_5", each 
    if [IS_ADDRESS_5_FINAL] then null else [ADDRESS 5], type nullable text),

AddTempInvestorName6 = Table.AddColumn(AddTempInvestorName5, "TEMPORARY_INVESTOR_NAME_6", each 
    if [IS_ADDRESS_6_FINAL] then null else [ADDRESS 6], type nullable text),

AddTempInvestorName7 = Table.AddColumn(AddTempInvestorName6, "TEMPORARY_INVESTOR_NAME_7", each 
    if [IS_ADDRESS_7_FINAL] then null else [ADDRESS 7], type nullable text),

// ============================================================
// STEP: CONCATENATE TEMPORARY INVESTOR NAME FIELDS
// Description: Merges all TEMPORARY_INVESTOR_NAME_0 to _7 into 
//              a single consolidated text field, separated by " | "
//              This field will be used to analyze hidden investor names
// Output: TEMPORARY_INVESTOR_NAME_FINAL
// ============================================================

AddTemporaryInvestorFinal = Table.AddColumn(AddTempInvestorName7, "TEMPORARY_INVESTOR_NAME_FINAL", each 
    Text.Combine(List.RemoveNulls({
        [TEMPORARY_INVESTOR_NAME_0],
        [TEMPORARY_INVESTOR_NAME_1],
        [TEMPORARY_INVESTOR_NAME_2],
        [TEMPORARY_INVESTOR_NAME_3],
        [TEMPORARY_INVESTOR_NAME_4],
        [TEMPORARY_INVESTOR_NAME_5],
        [TEMPORARY_INVESTOR_NAME_6],
        [TEMPORARY_INVESTOR_NAME_7]
    }), " | "), type nullable text),
    #"Changed Type" = Table.TransformColumnTypes(AddTemporaryInvestorFinal,{{"INVESTOR_ROW_ID", type text}})
in
    #"Changed Type"