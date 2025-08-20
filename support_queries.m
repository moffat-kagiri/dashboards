// Requirements_Processing Query
let
    Source = Folder.Files(RequirementsFolderPath),

    // 1) Keep only Excel files
    FilteredExcel = Table.SelectRows(Source, each [Extension] = ".xlsx"),

    // 2) Robust WeekNum extraction as integer (handles "Week"/"Wk", case-insensitive)
    AddWeekNum =
        Table.AddColumn(
            FilteredExcel, "WeekNum",
            each
                let
                    nameText   = [Name],
                    tokens     = {"Week","Wk"},
                    positions  = List.Transform(tokens, each Text.PositionOf(nameText, _, Occurrence.First, Comparer.OrdinalIgnoreCase)),
                    validPos   = List.Select(positions, each _ >= 0),
                    posWeek    = if List.IsEmpty(validPos) then -1 else List.Min(validPos),
                    tail       = if posWeek >= 0 then Text.Range(nameText, posWeek + 4) else nameText, // skip "Week" (4 chars); works for "Wk" too
                    // Find first digit after token
                    digStart   = Text.PositionOfAny(tail, {"0".."9"}),
                    segment    = if digStart >= 0 then Text.Range(tail, digStart) else null,
                    // Take only the first contiguous run of digits
                    digitRun   =
                        if segment <> null then
                            Text.Combine(
                                List.FirstN(
                                    Text.ToList(segment),
                                    each _ >= "0" and _ <= "9"
                                )
                            )
                        else
                            "",
                    weekNum    = if Text.Length(digitRun) > 0 then Int64.From(digitRun) else null
                in
                    weekNum,
            Int64.Type
        ),

    // 3) Load each file
    ProcessFiles = Table.AddColumn(AddWeekNum, "ExcelData", each Excel.Workbook([Content], true)),

    // 4) Expand sheets
    ExpandedSheets =
        Table.ExpandTableColumn(
            ProcessFiles,
            "ExcelData",
            {"Name", "Item", "Data"},
            {"SheetName", "SheetType", "SheetData"}
        ),

    // 5) Classify sheet types
    ClassifySheets =
        Table.AddColumn(
            ExpandedSheets,
            "SheetCategory",
            each if Text.Contains([SheetType], "ADA", Comparer.OrdinalIgnoreCase) then "ADA"
                 else if Text.Contains([SheetType], "Payment", Comparer.OrdinalIgnoreCase) then "Payment"
                 else "Other"
        ),

    // 6) Keep only ADA/Payment sheets
    FilterRelevant = Table.SelectRows(ClassifySheets, each [SheetCategory] <> "Other"),

    // 7) ADA sheets → expand + preserve WeekNum from outer row
    // Keep only ADA rows
    ADA_Data = Table.SelectRows(FilterRelevant, each [SheetCategory] = "ADA"),

    // Build a per-row expanded table that carries WeekNum down
    ADA_Prepared =
        Table.AddColumn(
            ADA_Data,
            "ADA_Rows",
            (outer) =>
                let
                    t0 = outer[SheetData],
                    // Guard: handle null or non-table SheetData
                    t1 = if (t0 is table) then t0 else #table({},{}),
                    // Select only needed columns; create nulls if missing
                    t2 = Table.SelectColumns(t1, {"RefNo", "Reason for Pending Activation"}, MissingField.UseNull),
                    // Rename to standard names (ignore if already renamed/missing)
                    t3 = Table.RenameColumns(t2, {{"Reason for Pending Activation", "PendingActivationReason"}}, MissingField.Ignore),
                    // Add WeekNum from the OUTER row
                    t4 = Table.AddColumn(t3, "WeekNum", (inner) => outer[WeekNum], Int64.Type),
                    // Type columns
                    t5 = Table.TransformColumnTypes(t4, {{"RefNo", Int64.Type}, {"PendingActivationReason", type text}, {"WeekNum", Int64.Type}})
                in
                    t5
        ),

    // Combine all per-row tables
    ADA_Final = Table.Combine(ADA_Prepared[ADA_Rows]),

    ADA_Final_Typed =
        Table.TransformColumnTypes(
            ADA_Final,
            {{"RefNo", Int64.Type}, {"WeekNum", Int64.Type}, {"PendingActivationReason", type text}}
        ),

    // 8) Payment sheets → expand + preserve WeekNum from outer row
    // Keep only Payment rows
    Payment_Data = Table.SelectRows(FilterRelevant, each [SheetCategory] = "Payment"),

    Payment_Prepared =
        Table.AddColumn(
            Payment_Data,
            "Payment_Rows",
            (outer) =>
                let
                    t0 = outer[SheetData],
                    t1 = if (t0 is table) then t0 else #table({},{}),
                    t2 = Table.SelectColumns(t1, {"RefNo", "Payment Status", "Pending Premium"}, MissingField.UseNull),
                    t3 = Table.RenameColumns(t2, {{"Payment Status", "PaymentStatus"}, {"Pending Premium", "PendingPremium"}}, MissingField.Ignore),
                    t4 = Table.AddColumn(t3, "WeekNum", (inner) => outer[WeekNum], Int64.Type),
                    t5 = Table.TransformColumnTypes(t4, {{"RefNo", Int64.Type}, {"PaymentStatus", type text}, {"PendingPremium", type text}, {"WeekNum", Int64.Type}})
                in
                    t5
        ),

    Payment_Final = Table.Combine(Payment_Prepared[Payment_Rows]),
    Payment_Final_Typed =
        Table.TransformColumnTypes(
            Payment_Final,
            {{"RefNo", Int64.Type}, {"WeekNum", Int64.Type}, {"PaymentStatus", type text}, {"PendingPremium", type text}}
        ),

    // 9) Output
    FinalOutput = [
        ADA_Requirements = ADA_Final_Typed,
        Payment_Details  = Payment_Final_Typed
    ]
in
    FinalOutput
// Payment_Lookup
let
    Source = Requirements_AllFiles,
    Payment_Table = Source[Payment_Details],
    // Add Payment-specific transformations here
    SetTypes = Table.TransformColumnTypes(Payment_Table, {
        {"RefNo", Int64.Type},
        {"PaymentStatus", type text},
        {"PendingPremium", Int64.Type},  // 1=unpaid, 0=paid
        {"WeekNum", Int64.Type}
    }),
    // Set default for PendingPremium
    FinalTable = Table.ReplaceValue(
        SetTypes,
        null, 0, Replacer.ReplaceValue,
        {"PendingPremium"}
    )
in
    FinalTable
