// =============================================================================
//  MS Project -> Milestone Trend Analysis  |  Power Query (M) import
//
//  Create ONE query per block below, named exactly as in the block header:
//    Data > Get Data > From Other Sources > Blank Query > Advanced Editor > paste.
//
//    pMSPFolder, pDateCulture, pDelimiter, pEncoding  -> parameters (Connection only)
//    fnMSPDate, fnMSPHeader                            -> functions  (Connection only)
//    MSP_Files, MSP_Tasks                              -> staging    (Connection only)
//    MSP Observations  -> Load To: Table, existing worksheet 'MSP Observations'!$A$1
//    MSP Milestones    -> Load To: Table, existing worksheet 'MSP Milestones'!$A$1
//
//  Folder convention: one CSV per project per reporting cycle, named
//      <ProjectID>_<yyyy-mm-dd>.csv        e.g.  PRJ-001_2026-11-30.csv
//  The date in the file name is the reporting (status) date of that snapshot.
//  If it cannot be read, the file's Date modified is used instead.
// =============================================================================


// ===== Query: pMSPFolder =====================================================
// Full path of the folder that holds the MS Project CSV exports.
"C:\MTA\MSP Exports" meta [IsParameterQuery = true, Type = "Text", IsParameterQueryRequired = true]


// ===== Query: pDateCulture ===================================================
// Culture MS Project used to write dates: "en-GB" for 30/11/26, "en-US" for 11/30/26.
"en-GB" meta [IsParameterQuery = true, Type = "Text", IsParameterQueryRequired = true]


// ===== Query: pDelimiter =====================================================
// MS Project writes the Windows list separator: "," in most English locales, ";" in many European ones.
"," meta [IsParameterQuery = true, Type = "Text", IsParameterQueryRequired = true]


// ===== Query: pEncoding ======================================================
// 1252 = Windows ANSI (MS Project default CSV export). Use 65001 for UTF-8 files.
1252 meta [IsParameterQuery = true, Type = "Number", IsParameterQueryRequired = true]


// ===== Query: fnMSPDate ======================================================
// Converts an MS Project date text ("Mon 30/11/26", "30/11/26 17:00", "November 30, 2026 5:00 PM",
// "NA", "") into a date, or null.
(value as any) as nullable date =>
let
    AsDate =
        if value = null then null
        else if Value.Is(value, type date) then value
        else if Value.Is(value, type datetime) then Date.From(value)
        else null,
    Txt = if value = null then "" else Text.Trim(Text.From(value)),
    Words = List.Select(Text.Split(Txt, " "), each _ <> ""),
    W0 = if List.Count(Words) > 0 then Text.Lower(Text.Remove(Words{0}, {",", "."})) else "",
    IsWeekday =
        List.Count(Words) > 1
        and List.AnyTrue(List.Transform({"mon", "tue", "wed", "thu", "fri", "sat", "sun"}, (d) => Text.StartsWith(W0, d)))
        and Text.Length(Text.Select(W0, {"0".."9"})) = 0,
    NoWeekday = if IsWeekday then List.Skip(Words, 1) else Words,
    DateWords = List.Select(NoWeekday, each not Text.Contains(_, ":") and not List.Contains({"am", "pm"}, Text.Lower(_))),
    Clean = Text.Combine(DateWords, " "),
    Parsed =
        if AsDate <> null then AsDate
        else if Clean = "" or Text.Upper(Clean) = "NA" then null
        else try Date.FromText(Clean, pDateCulture)
             otherwise try Date.From(DateTime.FromText(Txt, pDateCulture))
             otherwise null
in
    Parsed


// ===== Query: fnMSPHeader ====================================================
// Normalises a column header: "Baseline_Finish", "Baseline Finish" and "% Complete" -> "baselinefinish", "complete".
(header as text) as text => Text.Lower(Text.Select(header, {"a".."z", "A".."Z", "0".."9"}))


// ===== Query: MSP_Files ======================================================
let
    Source = Folder.Files(pMSPFolder),
    CsvOnly = Table.SelectRows(Source, each Text.Lower([Extension]) = ".csv" and not Text.StartsWith([Name], "~")),
    WithMeta = Table.AddColumn(CsvOnly, "Meta", each
        let
            Stem = Text.BeforeDelimiter([Name], ".", {0, RelativePosition.FromEnd}),
            Parts = Text.Split(Stem, "_"),
            ProjectID = Text.Trim(Parts{0}),
            DateText = if List.Count(Parts) > 1 then Text.Trim(List.Last(Parts)) else "",
            RepDate = try Date.FromText(DateText) otherwise Date.From([Date modified])
        in
            [ProjectID = ProjectID, ReportingDate = RepDate]),
    Expanded = Table.ExpandRecordColumn(WithMeta, "Meta", {"ProjectID", "ReportingDate"}),
    WithData = Table.AddColumn(Expanded, "Data", each
        let
            Raw = Csv.Document([Content], [Delimiter = pDelimiter, Encoding = pEncoding, QuoteStyle = QuoteStyle.Csv]),
            Promoted = Table.PromoteHeaders(Raw, [PromoteAllScalars = true]),
            Normalised = Table.TransformColumnNames(Promoted, fnMSPHeader)
        in
            Normalised),
    Kept = Table.SelectColumns(WithData, {"Name", "Date modified", "ProjectID", "ReportingDate", "Data"}),
    Renamed = Table.RenameColumns(Kept, {{"Name", "FileName"}, {"Date modified", "FileModified"}})
in
    Renamed


// ===== Query: MSP_Tasks ======================================================
// One row per milestone task per file, with standard columns whatever the export map called them.
let
    Aliases = [
        uniqueid        = {"uniqueid", "uid"},
        name            = {"name", "taskname"},
        milestone       = {"milestone"},
        summary         = {"summary"},
        baselinefinish  = {"baselinefinish", "baseline0finish"},
        finish          = {"finish", "finishdate"},
        actualfinish    = {"actualfinish"},
        percentcomplete = {"percentcomplete", "complete", "pctcomplete"},
        resourcenames   = {"resourcenames", "resources"},
        priority        = {"priority"},
        category        = {"text1", "category"},
        notes           = {"notes"}
    ],
    Fields = Record.FieldNames(Aliases),
    Standardise = (t as table) as table =>
        let
            Cols = Table.ColumnNames(t),
            Picks = List.Transform(Fields, (f) => List.First(List.Select(Record.Field(Aliases, f), each List.Contains(Cols, _)), null)),
            Added = List.Accumulate(
                List.Positions(Fields), t,
                (state, i) =>
                    if Picks{i} = null then Table.AddColumn(state, "__" & Fields{i}, each null)
                    else Table.AddColumn(state, "__" & Fields{i}, (r) => Record.Field(r, Picks{i}))),
            Selected = Table.SelectColumns(Added, List.Transform(Fields, each "__" & _)),
            Renamed = Table.RenameColumns(Selected, List.Transform(Fields, each {"__" & _, _}))
        in
            Renamed,
    Files = MSP_Files,
    Std = Table.AddColumn(Files, "Tasks", each Standardise([Data])),
    Expanded = Table.ExpandTableColumn(Table.RemoveColumns(Std, {"Data"}), "Tasks", Fields),
    IsYes = (v) => List.Contains({"yes", "true", "1", "-1", "ja", "oui", "si", "sí"}, Text.Lower(Text.Trim(if v = null then "" else Text.From(v)))),
    MilestonesOnly = Table.SelectRows(Expanded, each
        IsYes([milestone]) and not IsYes([summary])
        and [uniqueid] <> null and Text.Trim(Text.From([uniqueid])) <> ""),
    AddID = Table.AddColumn(MilestonesOnly, "Milestone ID", each [ProjectID] & "-" & Text.Trim(Text.From([uniqueid])), type text),
    AddForecast = Table.AddColumn(AddID, "Forecast Completion Date", each fnMSPDate([finish]), type nullable date),
    AddPct = Table.AddColumn(AddForecast, "Pct", each try Number.From(Text.Select(Text.From([percentcomplete]), {"0".."9", "."})) otherwise null, type nullable number),
    AddActual = Table.AddColumn(AddPct, "Actual Completion Date", each
        let a = fnMSPDate([actualfinish]) in
        if a <> null then a
        else if [Pct] <> null and [Pct] >= 100 then [Forecast Completion Date]
        else null, type nullable date),
    AddBaseline = Table.AddColumn(AddActual, "Baseline Date", each fnMSPDate([baselinefinish]), type nullable date),
    AddComments = Table.AddColumn(AddBaseline, "Comments", each
        if [notes] = null then null else Text.Start(Text.Clean(Text.From([notes])), 250), type nullable text),
    AddOwner = Table.AddColumn(AddComments, "Responsible Person", each
        if [resourcenames] = null then null
        else Text.Combine(List.Transform(Text.Split(Text.From([resourcenames]), ","), each Text.Trim(Text.BeforeDelimiter(_, "["))), ", "),
        type nullable text),
    Valid = Table.SelectRows(AddOwner, each [Forecast Completion Date] <> null and [ReportingDate] <> null)
in
    Valid


// ===== Query: MSP Observations ===============================================
// One row per milestone per reporting date -> feeds tblMilestoneData.
// Column ORDER matters: the workbook reads columns A-E of sheet 'MSP Observations'.
let
    Src = MSP_Tasks,
    // if two files give the same milestone + reporting date, keep the most recently saved file
    Sorted = Table.Buffer(Table.Sort(Src, {{"Milestone ID", Order.Ascending}, {"ReportingDate", Order.Ascending}, {"FileModified", Order.Descending}})),
    Dedup = Table.Distinct(Sorted, {"Milestone ID", "ReportingDate"}),
    // MTA convention: a milestone's line ends at the first report that shows it complete
    FirstDone = Table.Group(
        Table.SelectRows(Dedup, each [Actual Completion Date] <> null),
        {"Milestone ID"}, {{"FirstDone", each List.Min([ReportingDate]), type date}}),
    Joined = Table.NestedJoin(Dedup, {"Milestone ID"}, FirstDone, {"Milestone ID"}, "D", JoinKind.LeftOuter),
    WithDone = Table.ExpandTableColumn(Joined, "D", {"FirstDone"}),
    UpToDone = Table.SelectRows(WithDone, each [FirstDone] = null or [ReportingDate] <= [FirstDone]),
    Selected = Table.SelectColumns(UpToDone, {"Milestone ID", "ReportingDate", "Forecast Completion Date", "Actual Completion Date", "Comments"}),
    Renamed = Table.RenameColumns(Selected, {{"ReportingDate", "Reporting Date"}}),
    Typed = Table.TransformColumnTypes(Renamed, {
        {"Milestone ID", type text}, {"Reporting Date", type date}, {"Forecast Completion Date", type date},
        {"Actual Completion Date", type date}, {"Comments", type text}}),
    Final = Table.Sort(Typed, {{"Reporting Date", Order.Ascending}, {"Milestone ID", Order.Ascending}})
in
    Final


// ===== Query: MSP Milestones =================================================
// One row per milestone -> feeds tblMilestones.
// Column ORDER matters: the workbook reads columns A-I of sheet 'MSP Milestones'.
let
    Src = MSP_Tasks,
    // latest file describes the milestone (name, owner, priority, category)
    Latest = Table.Distinct(Table.Buffer(Table.Sort(Src, {{"Milestone ID", Order.Ascending}, {"ReportingDate", Order.Descending}, {"FileModified", Order.Descending}})), {"Milestone ID"}),
    // baseline = latest Baseline Finish recorded; if MS Project was never baselined, the first forecast
    BaseKnown = Table.Group(Table.SelectRows(Src, each [Baseline Date] <> null), {"Milestone ID"},
        {{"KnownBaseline", each List.Last(Table.Sort(_, {{"ReportingDate", Order.Ascending}})[Baseline Date]), type date}}),
    FirstRows = Table.Distinct(Table.Buffer(Table.Sort(Src, {{"Milestone ID", Order.Ascending}, {"ReportingDate", Order.Ascending}})), {"Milestone ID"}),
    FirstFc = Table.RenameColumns(Table.SelectColumns(FirstRows, {"Milestone ID", "Forecast Completion Date"}), {{"Forecast Completion Date", "FirstForecast"}}),
    J1 = Table.ExpandTableColumn(Table.NestedJoin(Latest, {"Milestone ID"}, BaseKnown, {"Milestone ID"}, "B", JoinKind.LeftOuter), "B", {"KnownBaseline"}),
    J2 = Table.ExpandTableColumn(Table.NestedJoin(J1, {"Milestone ID"}, FirstFc, {"Milestone ID"}, "F", JoinKind.LeftOuter), "F", {"FirstForecast"}),
    AddBase = Table.AddColumn(J2, "Baseline", each if [KnownBaseline] <> null then [KnownBaseline] else [FirstForecast], type date),
    AddPriority = Table.AddColumn(AddBase, "PriorityText", each
        let p = try Number.From([priority]) otherwise 500 in
        if p = null then "Medium" else if p >= 800 then "Critical" else if p >= 600 then "High" else if p >= 400 then "Medium" else "Low", type text),
    AddCategory = Table.AddColumn(AddPriority, "CategoryText", each
        if [category] = null or Text.Trim(Text.From([category])) = "" then "Milestone" else Text.Trim(Text.From([category])), type text),
    AddDesc = Table.AddColumn(AddCategory, "Description", each "MS Project task UID " & Text.Trim(Text.From([uniqueid])), type text),
    AddDept = Table.AddColumn(AddDesc, "Department", each null, type nullable text),
    Selected = Table.SelectColumns(AddDept, {"ProjectID", "Milestone ID", "name", "Description", "Baseline", "Responsible Person", "Department", "PriorityText", "CategoryText"}),
    Renamed = Table.RenameColumns(Selected, {
        {"ProjectID", "Project ID"}, {"name", "Milestone Name"}, {"Description", "Milestone Description"},
        {"Baseline", "Baseline Date"}, {"PriorityText", "Priority"}, {"CategoryText", "Category"}}),
    Typed = Table.TransformColumnTypes(Renamed, {{"Project ID", type text}, {"Milestone Name", type text}, {"Baseline Date", type date}}),
    Final = Table.Sort(Typed, {{"Project ID", Order.Ascending}, {"Baseline Date", Order.Ascending}, {"Milestone ID", Order.Ascending}})
in
    Final
