// ============================================================
// PROJECT: NOBO List Cleaning and Consolidation
// STEP 9 - Address Cleaning, Standardization and Consolidation
// ============================================================
// Author: Jaqueline F. Filogonio
// Tool: Microsoft Power Query (Excel)
// Description:
// This step performs comprehensive address cleaning and postal data
// consolidation based on prior address validation flags.
//
// Key Objectives:
// - Reconstruct full addresses using IS_ADDRESS_X_FINAL logic.
// - Extract and standardize postal codes and FSAs.
// - Identify and map country information.
// - Define unified columns for “ADDRESS_CONSOLIDATED”,
//   “POSTAL_CODE_CONSOLIDATED”, and “COUNTRY_CONSOLIDATED”.
// - Remove redundant intermediate columns for a clean output.
//
// Note:
// This script contains only transformation logic and structural references.
// No real or confidential investor data are included.
// ============================================================

let
    // ============================================================
    // STEP 1 — Load data from initial address-flagged table
    // ============================================================
    Source = Investor_Unique_Row_STEP1,

    // ============================================================
    // STEP 2 — Create cleaned address columns (1–7)
    // ============================================================
    AddCol1 = Table.AddColumn(Source, "ADDRESS1_CLEANED", each if [IS_ADDRESS_1_FINAL] then [ADDRESS1] else null, type text),
    AddCol2 = Table.AddColumn(AddCol1, "ADDRESS2_CLEANED", each if [IS_ADDRESS_2_FINAL] then [ADDRESS2] else null, type text),
    AddCol3 = Table.AddColumn(AddCol2, "ADDRESS3_CLEANED", each if [IS_ADDRESS_3_FINAL] then [ADDRESS3] else null, type text),
    AddCol4 = Table.AddColumn(AddCol3, "ADDRESS4_CLEANED", each if [IS_ADDRESS_4_FINAL] then [ADDRESS4] else null, type text),
    AddCol5 = Table.AddColumn(AddCol4, "ADDRESS5_CLEANED", each if [IS_ADDRESS_5_FINAL] then [ADDRESS5] else null, type text),
    AddCol6 = Table.AddColumn(AddCol5, "ADDRESS6_CLEANED", each if [IS_ADDRESS_6_FINAL] then [ADDRESS6] else null, type text),
    AddCol7 = Table.AddColumn(AddCol6, "ADDRESS7_CLEANED", each if [IS_ADDRESS_7_FINAL] then [ADDRESS7] else null, type text),

    // ============================================================
    // STEP 3 — Consolidate all address parts into a single string
    // ============================================================
    AddConsolidatedAddress = Table.AddColumn(
        AddCol7,
        "ADDRESS_CONSOLIDATED",
        each Text.Combine(
            List.Select(
                {
                    [ADDRESS1_CLEANED],
                    [ADDRESS2_CLEANED],
                    [ADDRESS3_CLEANED],
                    [ADDRESS4_CLEANED],
                    [ADDRESS5_CLEANED],
                    [ADDRESS6_CLEANED],
                    [ADDRESS7_CLEANED]
                },
                each _ <> null and _ <> ""
            ),
            " "
        ),
        type text
    ),

    // ============================================================
    // STEP 4 — Extract Postal Code from consolidated address
    // ============================================================
    AddPostalCode = Table.AddColumn(
        AddConsolidatedAddress,
        "POSTAL_CODE_EXTRACTED",
        each 
            let
                txt = Text.Upper(Text.Trim([ADDRESS_CONSOLIDATED])),
                words = Text.Split(txt, " "),
                wordCount = List.Count(words),

                provinceAbbr = {"AB", "BC", "MB", "NB", "NL", "NS", "NT", "NU", "ON", "PE", "QC", "SK", "YT"},
                lastWords = if wordCount >= 3 then List.LastN(words, 3) else words,

                siglaIndex = List.PositionOfAny(lastWords, provinceAbbr),
                siglaPostal =
                    if siglaIndex <> -1 then
                        let
                            afterSigla = List.Skip(lastWords, siglaIndex + 1),
                            combo =
                                if List.Count(afterSigla) = 2 then afterSigla{0} & afterSigla{1}
                                else if List.Count(afterSigla) = 1 then afterSigla{0}
                                else null
                        in
                            if combo <> null and Text.Length(combo) = 6 then Text.Insert(combo, 3, " ") else null
                    else null,

                canadaIndex = List.PositionOf(words, "CANADA"),
                afterCanada =
                    if canadaIndex <> -1 and canadaIndex + 1 < wordCount then
                        let
                            possible1 = words{canadaIndex + 1},
                            possible2 = if canadaIndex + 2 < wordCount then words{canadaIndex + 2} else "",
                            combo =
                                if Text.Length(possible1) = 3 and Text.Length(possible2) = 3 then possible1 & possible2
                                else if Text.Length(possible1) = 6 then possible1
                                else null
                        in
                            if combo <> null and Text.Length(combo) = 6 then Text.Insert(combo, 3, " ") else null
                    else null,

                finalPostal =
                    if siglaPostal <> null then siglaPostal
                    else if canadaIndex <> -1 and afterCanada <> null then afterCanada
                    else null
            in
                finalPostal,
        type text
    ),

    // ============================================================
    // STEP 5 — Detect country name based on address text
    // ============================================================
    AddCountry = Table.AddColumn(
        AddPostalCode,
        "COUNTRY",
        each 
            let
                txt = Text.Upper(Text.Trim([ADDRESS_CONSOLIDATED])),
                CountryMap = [
                    USA = "UNITED STATES",
                    #"UNITED STATES" = "UNITED STATES",
                    UK = "UNITED KINGDOM",
                    #"UNITED KINGDOM" = "UNITED KINGDOM",
                    GERMANY = "GERMANY",
                    AUSTRALIA = "AUSTRALIA",
                    UAE = "UNITED ARAB EMIRATES",
                    #"UNITED ARAB EMIRATES" = "UNITED ARAB EMIRATES",
                    #"HONG KONG" = "HONG KONG",
                    BRAZIL = "BRAZIL",
                    FRANCE = "FRANCE",
                    JAPAN = "JAPAN",
                    CHINA = "CHINA",
                    SPAIN = "SPAIN",
                    INDIA = "INDIA",
                    CAYMAN = "CAYMAN ISLANDS",
                    #"ABU DHABI" = "UNITED ARAB EMIRATES",
                    QATAR = "QATAR",
                    SWITZERLAND = "SWITZERLAND",
                    NETHERLANDS = "NETHERLANDS"
                ],
                match = List.First(List.Select(Record.FieldNames(CountryMap), each Text.Contains(txt, _)), null),
                countryName = if match <> null then Record.Field(CountryMap, match) else null
            in
                countryName,
        type text
    ),

    // ============================================================
    // STEP 6 — Clean and standardize postal code formats
    // ============================================================
    CleanPostalCode = Table.AddColumn(
        AddCountry,
        "POSTAL_CODE_CLEANED",
        each 
            let
                raw = Text.Upper([POSTAL CODE]),
                noSpaces = Text.Remove(raw, {" ", "-"}),
                only6 = if Text.Length(noSpaces) = 6 then noSpaces else null
            in
                only6,
        type text
    ),

    AddPostalCodeConsolidated = Table.AddColumn(
        CleanPostalCode,
        "POSTAL_CODE_CONSOLIDATED",
        each if [POSTAL_CODE_CLEANED] <> null then [POSTAL_CODE_CLEANED] else [POSTAL_CODE_EXTRACTED],
        type text
    ),

    // ============================================================
    // STEP 7 — Identify Canadian records and derive FSA code
    // ============================================================
    AddIsCanadian = Table.AddColumn(
        AddPostalCodeConsolidated,
        "IS_CANADIAN",
        each 
            let
                country = Text.Upper(Text.Trim([COUNTRY])),
                postal = [POSTAL_CODE_CONSOLIDATED]
            in
                (country = "CANADA") or (postal <> null),
        type logical
    ),

    ADD_FSA = Table.AddColumn(
        AddIsCanadian,
        "POSTAL_CODE_FSA",
        each Text.Start([POSTAL_CODE_CONSOLIDATED], 3),
        type text
    ),

    // ============================================================
    // STEP 8 — Consolidate country information
    // ============================================================
    AddCountryConsolidated = Table.AddColumn(
        ADD_FSA,
        "COUNTRY_CONSOLIDATED",
        each if [IS_CANADIAN] then "CANADA" else [COUNTRY],
        type text
    ),

    // ============================================================
    // STEP 9 — Remove redundant and intermediate columns
    // ============================================================
    Removed_Columns = Table.RemoveColumns(
        AddCountryConsolidated,
        {
            "NAME", "ADDRESS 1", "ADDRESS 2", "ADDRESS 3", "ADDRESS 4", "ADDRESS 5", "ADDRESS 6", "ADDRESS 7",
            "POSTAL CODE", "POSTAL REGION", "NOTICE AND ACCESS", "E-MAIL ADDRESS", "LANGUAGE", "NUMBER OF SHARES",
            "RECEIVE ALL MATERIALS", "AGREE TO ELECTRONIC DELIVERY", "CUID",
            "TEMPORARY_INVESTOR_NAME_0", "ADDRESS1", "ADDRESS2", "ADDRESS3", "ADDRESS4", "ADDRESS5", "ADDRESS6", "ADDRESS7",
            "IS_ADDRESS_1", "IS_ADDRESS_2", "IS_ADDRESS_3", "IS_ADDRESS_4", "IS_ADDRESS_5", "IS_ADDRESS_6", "IS_ADDRESS_7",
            "IS_ADDRESS_1_FINAL", "IS_ADDRESS_2_FINAL", "IS_ADDRESS_3_FINAL", "IS_ADDRESS_4_FINAL", "IS_ADDRESS_5_FINAL",
            "IS_ADDRESS_6_FINAL", "IS_ADDRESS_7_FINAL",
            "TEMPORARY_INVESTOR_NAME_1", "TEMPORARY_INVESTOR_NAME_2", "TEMPORARY_INVESTOR_NAME_3", "TEMPORARY_INVESTOR_NAME_4",
            "TEMPORARY_INVESTOR_NAME_5", "TEMPORARY_INVESTOR_NAME_6", "TEMPORARY_INVESTOR_NAME_7",
            "TEMPORARY_INVESTOR_NAME_FINAL",
            "ADDRESS1_CLEANED", "ADDRESS2_CLEANED", "ADDRESS3_CLEANED", "ADDRESS4_CLEANED", "ADDRESS5_CLEANED",
            "ADDRESS6_CLEANED", "ADDRESS7_CLEANED",
            "POSTAL_CODE_EXTRACTED", "COUNTRY", "POSTAL_CODE_CLEANED"
        }
    )

in
    Removed_Columns
