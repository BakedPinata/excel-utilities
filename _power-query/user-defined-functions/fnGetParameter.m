// fnGetParameter
let
    Source = (ParameterName as text, optional TableName as text, optional NameColIndex as number, optional ValueColIndex as number) as nullable text =>
        let
            // Paste the content of this inner let function into another query then uncomment the following lines for testing
            /*
            Source = "",
            ParameterName = "PTH",
            TableName = "tblParams",
            NameColIndex = 1,
            ValueColIndex = 2,
            */
            _TableName = if TableName is null
                then "Parameters"
                else TableName,
            _NameColIndex = if NameColIndex is null
                then 0
                else NameColIndex,
            _ValueColIndex = if ValueColIndex is null
                then 1
                else ValueColIndex,
            ParamSource = Excel.CurrentWorkbook(){[Name=_TableName]}[Content],
            ParamRow = Table.SelectRows(ParamSource, each (Record.Field(_,Table.ColumnNames(ParamSource){_NameColIndex}) = ParameterName)),
            Value = if Table.IsEmpty(ParamRow)=true
                then ""
                else Record.Field(ParamRow{0},Table.ColumnNames(ParamRow){_ValueColIndex}),
            OutValue = if Value.Type(Value)=Number.Type
                then Number.ToText(Value)
                else Value
        in
            OutValue
in
    Source