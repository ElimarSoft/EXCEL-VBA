Public Function GetFieldData(DataBaseFile As String, TableName As String, IndexField As String)

    'Get data from an access file
    'The index is the first column and field names are in the first row
    
    Dim myRange As Range: Set myRange = Application.ThisCell
   
    Dim row1 As Integer: row1 = myRange.row
    Dim col1 As Integer: col1 = myRange.Column

    Dim conn As ADODB.Connection
    Dim rs1 As ADODB.Recordset
    Dim connStr As String
    Dim strSQL As String
    
    Dim tag1 As String: tag1 = Cells(row1, 1)
    
    Dim FieldNames As String
    
    FieldNames = Cells(1, col1)
        
    Dim i As Integer: i = 1
    Do
        If Cells(1, col1 + i) = vbNullString Then Exit Do
        FieldNames = FieldNames + "," + Cells(1, col1 + i)
        i = i + 1
        
    Loop
    
    tag1 = "'" + tag1 + "'"
    
    'Reference Microsoft ActiveX Data Objects 2.8 Library
    connStr = "Data Source = " + DataBaseFile
    Set conn = New ADODB.Connection
    conn.Provider = "Microsoft.ACE.OLEDB.12.0"
    conn.Open (connStr)

    Set rs1 = New ADODB.Recordset
    
    Dim Filter As String: Filter = IndexField + " Like " + tag1
    Dim Query As String: Query = "Select " + FieldNames + " from " + TableName + " where " + Filter
    rs1.Open Query, conn
    
    'Get Record Data, transpose and filter null items
    Dim Data1() As Variant: Data1 = rs1.GetRows
    Dim Data2() As Variant: ReDim Data2(UBound(Data1, 2), UBound(Data1, 1))
    
    Dim n As Integer, m As Integer
    For n = 0 To UBound(Data1, 1)
    For m = 0 To UBound(Data1, 2)
        If IsNull(Data1(n, m)) Then
            Data2(m, n) = vbNullString
        Else
            Data2(m, n) = Data1(n, m)
        End If
    Next m
    Next n

    GetFieldData = Data2

conn.Close


End Function
