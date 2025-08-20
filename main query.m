let
    Source = Folder.Files(ASReportsFolderPath),
    #"Filtered Hidden Files1" = Table.SelectRows(Source, each [Attributes]?[Hidden]? <> true),
    #"Invoke Custom Function1" = Table.AddColumn(#"Filtered Hidden Files1", "Transform File", each #"Transform File"([Content])),
    #"Renamed Columns1" = Table.RenameColumns(#"Invoke Custom Function1", {"Name", "Source.Name"}),
    #"Removed Other Columns1" = Table.SelectColumns(#"Renamed Columns1", {"Source.Name", "Transform File"}),
    #"Expanded Table Column1" = Table.ExpandTableColumn(#"Removed Other Columns1", "Transform File", Table.ColumnNames(#"Transform File"(#"Sample File"))),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded Table Column1",{{"Source.Name", type text}, {"RefNo", Int64.Type}, {"Product_desc", type text}, {"StatusID", Int64.Type}, {"CommenceDate", type datetime}, {"EstimatedPremium", type number}, {"PayType", type text}, {"CreatedDate", type datetime}, {"PaymentFrequencyID", Int64.Type}, {"Lives", Int64.Type}, {"LevelCode", Int64.Type}, {"AgencyCode", Int64.Type}, {"AgencyName", type text}, {"UnitCode", Int64.Type}, {"UnitName", type text}, {"Pay_Frequency", type text}, {"PolicySuspense", type number}, {"Description", type text}, {"PendingActivation", Int64.Type}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type", "Conversion", each if [Description] = "New_Business_Captured" then "New Business" 
else if [Description] = "Lapse_Automatic" or [Description] = "Paid_Up_Automatic" then "Lapse or Paid Up" 
else if [Description] = "NTU" or [Description] = "Cancel" then "Lost" 
else "Converted"),
    #"Added Custom1" = Table.AddColumn(#"Added Custom", "WeekNumber", each let
    CurrentDate = Date.From([CreatedDate]),
    YearStart = #date(Date.Year(CurrentDate), 1, 1),
    DaysOffset = Number.Mod(Date.DayOfWeek(YearStart, Day.Sunday) + 5, 7),
    AdjustedDate = Date.AddDays(CurrentDate, DaysOffset),
    WeekNum = Date.WeekOfYear(AdjustedDate)
in
    "Week " & Text.From(WeekNum)),
    #"Added Custom2" = Table.AddColumn(#"Added Custom1", "AnnualPremium", each if [Pay_Frequency] = null or [EstimatedPremium] = null then null
else [EstimatedPremium] * 
    (if [Pay_Frequency] = "Monthly" then 12
     else if [Pay_Frequency] = "Quarterly" then 4
     else if [Pay_Frequency] = "Half-yearly" then 2
     else 1)),
    #"Added Custom3" = Table.AddColumn(#"Added Custom2", "WeekNum", each Text.Middle([Source.Name], Text.PositionOf([Source.Name], "Week ")+5, 2)),
    #"TypeRefine" =
        Table.TransformColumns(
            #"Added Custom3",
            {{"RefNo", Int64.From}, {"WeekNum", Int64.From}}
        ),

    // Buffer lookup with proper types
    Buffered_ADA_Data =
        Table.Buffer(
            Table.TransformColumns(
                ADA_Lookup,
                {{"RefNo", Int64.From}, {"WeekNum", Int64.From}}
            )
        ),

    // Step 1: Merge with ADA_Lookup (like XLOOKUP)
    #"Merged ADA" =
        Table.NestedJoin(
            #"TypeRefine",
            {"RefNo", "WeekNum"},
            Buffered_ADA_Data,
            {"RefNo", "WeekNum"},
            "ADA_Match",
            JoinKind.LeftOuter
        ),

    // Step 2: Expand PendingActivationReason
    #"Expanded ADA" =
        Table.ExpandTableColumn(
            #"Merged ADA",
            "ADA_Match",
            {"PendingActivationReason"},
            {"UndRequirements"} // Directly name it in expansion
        ),

    // Step 3: Replace nulls with default
    #"With Requirements" =
        Table.ReplaceValue(
            #"Expanded ADA",
            null,
            "Requirements met",
            Replacer.ReplaceValue,
            {"UndRequirements"}
        ),

    // Buffer Payment_Lookup with proper types
    Buffered_Payment_Data =
        Table.Buffer(
            Table.TransformColumns(
                Payment_Lookup,
                {{"RefNo", Int64.From}, {"WeekNum", Int64.From}}
            )
        ),

    // Step 1: Merge with Payment_Lookup starting from With Requirements table
    #"Merged Payment" =
        Table.NestedJoin(
            #"With Requirements",              // ✅ Start from table that already has UndRequirements
            {"RefNo", "WeekNum"},
            Buffered_Payment_Data,
            {"RefNo", "WeekNum"},
            "Payment_Match",
            JoinKind.LeftOuter
        ),

    // Step 2: Expand PaymentStatus
    #"Expanded Payment" =
        Table.ExpandTableColumn(
            #"Merged Payment",
            "Payment_Match",
            {"PaymentStatus"},
            {"PaymentMode"} // Directly name it in expansion
        ),

    // Step 3: Replace nulls with default
    #"With PaymentMode" =
        Table.ReplaceValue(
            #"Expanded Payment",
            null,
            "Unspecified",
            Replacer.ReplaceValue,
            {"PaymentMode"}
        ),
    // Left join to bring in PendingPremium
    MergedReports =
        Table.NestedJoin(
            #"With PaymentMode",
            {"RefNo", "WeekNum"},
            Payment_Lookup,
            {"RefNo", "WeekNum"},
            "PaymentMatch",
            JoinKind.LeftOuter
        ),

    // Expand PendingPremium
    ExpandedReports =
        Table.ExpandTableColumn(MergedReports, "PaymentMatch", {"PendingPremium"}, {"PendingPremium"}),

    // Replace nulls with 0
    FinalReports =
        Table.ReplaceValue(
            ExpandedReports,
            null,
            0,
            Replacer.ReplaceValue,
            {"PendingPremium"}
        ),
    #"Changed Type1" = Table.TransformColumnTypes(FinalReports,{{"PendingPremium", Int64.Type}})
in
    #"Changed Type1"
