Attribute VB_Name = "modMain"
'***************************************************************************
' MMMM  MMMMM  OMMM   MMMO    OMMM    OMMM    OMMMMO     OMMMMO    OMMMMO  '
'  MM    MM   MM MM    MMMO  OMMM    MM MM    MM   MO   OM    MO  OM    MO '
'  MM  MM    MM  MM    MM  OO  MM   MM  MM    MM   MO   OM    MO       OMO '
'  MMMM     MMMMMMMM   MM  MM  MM  MMMMMMMM   MMMMMO     OMMMMO      OMO   '
'  MM  MM        MM    MM      MM       MM    MM   MO   OM    MO   OMO     '
'  MM    MM      MM    MM      MM       MM    MM    MO  OM    MO  OM   MM  '
' MMMM  MMMM    MMMM  MMMM    MMMM     MMMM  MMMM  MMMM  OMMMMO   MMMMMMM  '
'                                                                          '
' K4m4r82's Laboratory                                                     '
' http://coding4ever.wordpress.com                                         '
'***************************************************************************

Option Explicit

Public conn     As ADODB.Connection
Public strSql   As String

Public Function dbGetValue(ByVal query As String, ByVal defValue As Variant) As Variant
    Dim rsDbGetValue  As ADODB.Recordset
    
    On Error GoTo errHandle
    
    Set rsDbGetValue = New ADODB.Recordset
    rsDbGetValue.Open strSql, conn, adOpenForwardOnly, adLockReadOnly
    If Not rsDbGetValue.EOF Then
        If Not IsNull(rsDbGetValue(0).Value) Then
            dbGetValue = rsDbGetValue(0).Value
        Else
            dbGetValue = defValue
        End If
    Else
        dbGetValue = defValue
    End If
    rsDbGetValue.Close
    Set rsDbGetValue = Nothing
    
    Exit Function
errHandle:
    dbGetValue = defValue
End Function

Public Function getJumlahHariByBulan(ByVal bulan As Integer, ByVal tahun As Long) As Integer
    getJumlahHariByBulan = Day(DateSerial(tahun, bulan + 1, 0))
End Function

Private Function getFirstSunday() As Integer
    Dim firstDay As String
    
    firstDay = Year(Now) & "/" & Month(Now) & "/1"
    firstDay = Weekday(firstDay)
    If Val(firstDay) > 1 Then
        getFirstSunday = 9 - Val(firstDay)
    Else
        getFirstSunday = Val(firstDay)
    End If
End Function

Private Sub openDb()
    Set conn = New ADODB.Connection
    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\sampleDB.mdb"
    conn.Open
End Sub

Private Sub addHariMinggu()
    Dim i           As Integer
    Dim firstDay    As Integer
    Dim ret         As Integer

    Dim tgl         As String

    firstDay = getFirstSunday 'ambil tgl hari minggu pertama
    For i = firstDay To getJumlahHariByBulan(Month(Now), Year(Now)) Step 7
        tgl = Year(Now) & "/" & Month(Now) & "/" & i

        strSql = "SELECT COUNT(*) FROM hari_libur " & _
                 "WHERE tanggal = #" & tgl & "# AND keterangan = 'Minggu'"
        ret = CInt(dbGetValue(strSql, 0))
        If ret = 0 Then
            strSql = "INSERT INTO hari_libur(tanggal, keterangan) VALUES (#" & tgl & "#, 'Minggu')"
            conn.Execute strSql
        End If
    Next
End Sub

Public Sub Main()
    Call openDb
    
    'prosedur otomatis untuk mengisikan tgl libur khusus hari minggu
    Call addHariMinggu
    frmHariLibur.Show
End Sub

