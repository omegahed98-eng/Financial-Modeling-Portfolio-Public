let
    Params = Excel.CurrentWorkbook(){[Name="tblSourceParams"]}[Content],
    ReadParam = (k as text, fallback as text) as text =>
        let found = Table.SelectRows(Params, each Text.Upper(Text.From([Parameter])) = Text.Upper(k))
        in if Table.RowCount(found)=0 then fallback else Text.From(found{0}[Value]),
    Root = ReadParam("RootFolder", "C:\\PortfolioData"),
    SelectID = ReadParam("ActiveSourceID", ""),
    Mode = Text.Upper(ReadParam("LoadMode", "SINGLE")),
    Sources = Excel.CurrentWorkbook(){[Name="tblSourceRegister"]}[Content],
    Headers = Excel.CurrentWorkbook(){[Name="tblHeaderMap"]}[Content],
    Maps = Excel.CurrentWorkbook(){[Name="tblDimensionMap"]}[Content],
    PeriodMaps = Excel.CurrentWorkbook(){[Name="tblPeriodMap"]}[Content],
    Canon = {"RecordID","SourceID","BusinessLineID","CostCenterID","AccountID","ProductID","PeriodKey","Version","Scenario","SnapshotID","AmountCAD","Quantity","UnitPriceCAD","EventDate","DueDate","CounterpartyID","DocumentID","SourceRef","SourceFormat","QualityStatus","QualityMessage"},
    SourceFields = {"RawBusinessLine","RawCostCenter","RawAccount","RawProduct","PeriodKey","Version","Scenario","SnapshotID","AmountCAD","Quantity","UnitPriceCAD","EventDate","DueDate","CounterpartyID","DocumentID","SourceRef","MoneyInCAD","MoneyOutCAD"},
    CleanText = (v as any) as nullable text => if v=null then null else let x=Text.Trim(Text.From(v)) in if x="" then null else x,
    ParseNum = (v as any) as nullable number => try if v=null or CleanText(v)=null then null else Number.From(v) otherwise null,
    ParseDate = (v as any) as nullable date => try if v=null or CleanText(v)=null then null else Date.From(v) otherwise null,
    MapDim = (dim as text, source as text, raw as any, def as any) as nullable text =>
        let value = CleanText(raw), candidate = if value=null then CleanText(def) else value,
            rows = if candidate=null then #table(type table [CanonicalID=text], {})
                else Table.SelectRows(Maps, each [Dimension]=dim and (Text.From([SourceID])=source or Text.From([SourceID])="*") and Text.Upper(Text.Trim(Text.From([RawValue])))=Text.Upper(candidate)),
            ordered = if Table.RowCount(rows)=0 then rows else Table.AddColumn(rows,"Priority", each if Text.From([SourceID])=source then 0 else 1, Int64.Type),
            sorted = if Table.RowCount(ordered)=0 then ordered else Table.Sort(ordered,{{"Priority",Order.Ascending}})
        in if Table.RowCount(sorted)=0 then null else CleanText(sorted{0}[CanonicalID]),
    MapPeriod = (source as text, raw as any, event as nullable date) as nullable text =>
        let t=CleanText(raw), matched = if t=null then #table(type table [PeriodKey=text], {}) else
                Table.SelectRows(PeriodMaps, each (Text.From([SourceID])=source or Text.From([SourceID])="*") and Text.Upper(Text.Trim(Text.From([RawPeriod])))=Text.Upper(t)),
            result = if Table.RowCount(matched)>0 then CleanText(matched{0}[PeriodKey]) else if t<>null and Text.Length(t)=7 and Text.Middle(t,4,1)="-" then t
                     else if event<>null then Date.ToText(event,"yyyy-MM") else null
        in result,
    LoadSource = (r as record) as table =>
        let sid=Text.From(r[SourceID]), fmt=Text.Upper(Text.From(r[Format])),
            rel=Text.From(r[RelativePath]), path=Root & (if Text.EndsWith(Root,"\") or Text.EndsWith(Root,"/") then "" else "\") & Text.Replace(rel,"/","\"),
            objectName=CleanText(r[ObjectName]),
            raw = if fmt="CSV" then Table.PromoteHeaders(Csv.Document(File.Contents(path),[Delimiter=",",Encoding=65001,QuoteStyle=QuoteStyle.Csv]),[PromoteAllScalars=true])
                  else if fmt="XLSX" or fmt="EXCEL" then
                      let w=Excel.Workbook(File.Contents(path),null,true), opt=Table.SelectRows(w,each [Item]=objectName and ([Kind]="Sheet" or [Kind]="Table"))
                      in if Table.RowCount(opt)=0 then error "Missing Excel object: " & objectName else Table.PromoteHeaders(opt{0}[Data],[PromoteAllScalars=true])
                  else error "Source format unsupported: " & fmt,
            hrows=Table.SelectRows(Headers, each Text.From([SourceID])=sid or Text.From([SourceID])="*"),
            renamed=List.Accumulate(Table.ToRecords(hrows),raw,(state,x)=>
                let match=List.Select(Table.ColumnNames(state), each Text.Upper(Text.Trim(_))=Text.Upper(Text.Trim(Text.From(x[RawHeader])))),
                    dst=Text.From(x[CanonicalHeader])
                in if List.Count(match)=0 or List.Contains(Table.ColumnNames(state),dst) then state
                   else Table.RenameColumns(state,{{match{0},dst}},MissingField.Ignore)),
            isWide=Text.Upper(Text.From(r[ValueMode]))="WIDE",
            keepDims=List.Intersect({Table.ColumnNames(renamed),List.RemoveItems(SourceFields,{"PeriodKey","AmountCAD","MoneyInCAD","MoneyOutCAD"})}),
            validPeriods=List.Transform(PeriodMaps[RawPeriod],each Text.Upper(Text.Trim(Text.From(_)))),
            periodColumns=List.Select(Table.ColumnNames(renamed),each List.Contains(validPeriods,Text.Upper(Text.Trim(_))) or (Text.Length(_)=7 and Text.Middle(_,4,1)="-")),
            wideClean=Table.SelectColumns(renamed,List.Combine({keepDims,periodColumns}),MissingField.Ignore),
            normalizedWide=if isWide then Table.UnpivotOtherColumns(wideClean,keepDims,"PeriodKey","AmountCAD") else renamed,
            selected=Table.SelectColumns(normalizedWide,SourceFields,MissingField.UseNull),
            numbered=Table.AddIndexColumn(selected,"_RowNumber",1,1,Int64.Type),
            withRow=Table.AddColumn(numbered,"_Normalized",(x)=>
                let event=ParseDate(x[EventDate]), due=ParseDate(x[DueDate]),
                    amount = if x[AmountCAD]<>null then ParseNum(x[AmountCAD]) else if x[MoneyInCAD]<>null or x[MoneyOutCAD]<>null then
                       (if ParseNum(x[MoneyInCAD])=null then 0 else ParseNum(x[MoneyInCAD]))-(if ParseNum(x[MoneyOutCAD])=null then 0 else ParseNum(x[MoneyOutCAD])) else null,
                    bl=MapDim("BusinessLine",sid,x[RawBusinessLine],r[BLDefault]),
                    cc=MapDim("CostCenter",sid,x[RawCostCenter],r[CCDefault]),
                    acct=MapDim("Account",sid,x[RawAccount],r[AccountDefault]),
                    prod=if CleanText(x[RawProduct])=null then "ALL" else MapDim("Product",sid,x[RawProduct],null),
                    period=MapPeriod(sid,x[PeriodKey],if due<>null then due else event),
                    version=if CleanText(x[Version])=null then CleanText(r[VersionDefault]) else CleanText(x[Version]),
                    scenario=if CleanText(x[Scenario])=null then CleanText(r[ScenarioDefault]) else CleanText(x[Scenario]),
                    doc=CleanText(x[DocumentID]),
                    errs=List.RemoveNulls({if bl=null then "BL_UNMAPPED" else null,if cc=null then "CC_UNMAPPED" else null,
                          if acct=null then "ACCOUNT_UNMAPPED" else null,if prod=null then "PRODUCT_UNMAPPED" else null,
                          if period=null or (try Date.FromText(period & "-01") otherwise null)=null then "PERIOD_INVALID" else null,
                          if amount=null then "AMOUNT_INVALID" else null,if not List.Contains({"Actual","Budget","Forecast"},version) then "VERSION_INVALID" else null,
                          if not List.Contains({"Base","Upside","Downside"},scenario) then "SCENARIO_INVALID" else null}),
                    qa=Text.Combine(errs,";"),
                    rec=[RecordID=sid & ":" & (if doc=null then "ROW" else doc) & ":" & Text.From(x[_RowNumber]),
                         SourceID=sid,BusinessLineID=bl,CostCenterID=cc,AccountID=acct,ProductID=prod,PeriodKey=period,
                         Version=version,Scenario=scenario,SnapshotID=if CleanText(x[SnapshotID])=null then "FY27_Q2" else CleanText(x[SnapshotID]),
                         AmountCAD=amount,Quantity=ParseNum(x[Quantity]),UnitPriceCAD=ParseNum(x[UnitPriceCAD]),
                         EventDate=event,DueDate=due,CounterpartyID=CleanText(x[CounterpartyID]),DocumentID=doc,
                         SourceRef=if CleanText(x[SourceRef])=null then rel else CleanText(x[SourceRef]),SourceFormat=fmt,
                         QualityStatus=if qa="" then "PASS" else "FAIL",QualityMessage=qa]
                in rec, type record),
            noCollisions=Table.SelectColumns(withRow,{"_Normalized"}),
            expanded=Table.ExpandRecordColumn(noCollisions,"_Normalized",Canon,Canon),
            output=Table.SelectColumns(expanded,Canon,MissingField.UseNull)
        in output,
    target = if Mode="ALL" then Table.SelectRows(Sources,each Text.Upper(Text.From([Enabled]))="Y")
             else Table.SelectRows(Sources,each Text.From([SourceID])=SelectID),
    allRows=if Table.RowCount(target)=0 then #table(Canon,{}) else Table.Combine(List.Transform(Table.ToRecords(target),each LoadSource(_))),
    duplicates=Table.Group(allRows,{"SourceID","RecordID"},{{"KeyCount",each Table.RowCount(_),Int64.Type}}),
    joined=Table.NestedJoin(allRows,{"SourceID","RecordID"},duplicates,{"SourceID","RecordID"},"_Dup",JoinKind.LeftOuter),
    expandedDup=Table.ExpandTableColumn(joined,"_Dup",{"KeyCount"},{"KeyCount"}),
    qa=Table.AddColumn(expandedDup,"_QA", each if [KeyCount]>1 then Text.Combine(List.Select({[QualityMessage],"DUPLICATE_RECORD"},each _<>null and _<>""),";") else [QualityMessage],type text),
    drop=Table.RemoveColumns(qa,{"KeyCount","QualityMessage","QualityStatus"}),
    setMsg=Table.RenameColumns(drop,{{"_QA","QualityMessage"}}),
    setStatus=Table.AddColumn(setMsg,"QualityStatus",each if [QualityMessage]="" then "PASS" else "FAIL",type text),
    ordered=Table.SelectColumns(setStatus,Canon,MissingField.UseNull),
    typed=Table.TransformColumnTypes(ordered,{{"RecordID",type text},{"SourceID",type text},{"BusinessLineID",type text},{"CostCenterID",type text},{"AccountID",type text},{"ProductID",type text},{"PeriodKey",type text},{"Version",type text},{"Scenario",type text},{"SnapshotID",type text},{"AmountCAD",type number},{"Quantity",type number},{"UnitPriceCAD",type number},{"EventDate",type date},{"DueDate",type date},{"CounterpartyID",type text},{"DocumentID",type text},{"SourceRef",type text},{"SourceFormat",type text},{"QualityStatus",type text},{"QualityMessage",type text}})
in
    typed