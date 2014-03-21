VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmHariLibur 
   Caption         =   "Demo Form Input Hari Libur"
   ClientHeight    =   3690
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   10305
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   10305
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstKetHariLibur 
      Appearance      =   0  'Flat
      Height          =   1785
      Left            =   5205
      TabIndex        =   5
      Top             =   375
      Width           =   4965
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">>"
      Height          =   285
      Left            =   4710
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "<<"
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtBulan 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   480
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   4225
   End
   Begin MSFlexGridLib.MSFlexGrid gridKalender 
      Height          =   3495
      Left            =   120
      TabIndex        =   2
      Top             =   375
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   6165
      _Version        =   393216
      Rows            =   7
      Cols            =   3
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   2
      ScrollBars      =   0
      Appearance      =   0
   End
   Begin VB.Label Label1 
      Caption         =   "Keterangan Hari Libur"
      Height          =   255
      Left            =   5205
      TabIndex        =   4
      Top             =   120
      Width           =   4965
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "Pop Up Menu"
      Begin VB.Menu mnuHariLibur 
         Caption         =   "Hari Libur"
      End
      Begin VB.Menu mnuSpr 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHapusHariLibur 
         Caption         =   "Hapus Hari Libur"
      End
   End
End
Attribute VB_Name = "frmHariLibur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Dim arrHari(6)  As String
Dim setMonth    As Date
Dim leap        As Boolean

Private Sub initHari()
    arrHari(0) = "Minggu"
    arrHari(1) = "Senin"
    arrHari(2) = "Selasa"
    arrHari(3) = "Rabu"
    arrHari(4) = "Kamis"
    arrHari(5) = "Jum'at"
    arrHari(6) = "Sabtu"
End Sub

Private Sub initGrid()
    Dim x As Integer
    
    With gridKalender
        .Cols = 7
        .Rows = 7
        .FixedRows = 1
        .FixedCols = 0
        
        For x = 0 To .Cols - 1 'looping untuk pengaturan judul tabel
            .Col = x
            .Row = 0

            .CellFontBold = True
            .FixedAlignment(x) = flexAlignCenterCenter
            
            .ColWidth(x) = 700
            .ColAlignment(x) = flexAlignCenterCenter
        Next x
        
        For x = 0 To .Cols - 1
            .TextMatrix(0, x) = arrHari(x)
        Next
        
        For x = 0 To .Rows - 1
            .RowHeight(x) = 500
        Next
        
        .GridLines = flexGridFlat
        .GridLinesFixed = flexGridFlat
        
        .ForeColorFixed = &H0& 'WARNA_HITAM
        .BackColorSel = &HED9564 'WARNA_BIRU
    End With
End Sub

Private Function setNewMonth(ByVal incrementMonth As Boolean) As String
    Dim tahun   As Long
    Dim bulan   As Long
    
    tahun = Year(setMonth)
    bulan = Month(setMonth)
    
    If incrementMonth Then
        bulan = bulan + 1
        
        If bulan = 13 Then
            bulan = 1
            tahun = tahun + 1
        End If
        
    Else
        bulan = bulan - 1
        If bulan = 0 Then
            bulan = 12
            tahun = tahun - 1
        End If
    End If
    
    setNewMonth = tahun & "/" & bulan & "/1"
End Function

Private Function getBulanIndonesia(ByVal bulan As Integer) As String
    Select Case bulan
        Case 1: getBulanIndonesia = "Januari"
        Case 2: getBulanIndonesia = "Februari"
        Case 3: getBulanIndonesia = "Maret"
        Case 4: getBulanIndonesia = "April"
        Case 5: getBulanIndonesia = "Mei"
        Case 6: getBulanIndonesia = "Juni"
        Case 7: getBulanIndonesia = "Juli"
        Case 8: getBulanIndonesia = "Agustus"
        Case 9: getBulanIndonesia = "September"
        Case 10: getBulanIndonesia = "Oktober"
        Case 11: getBulanIndonesia = "November"
        Case 12: getBulanIndonesia = "Desember"
    End Select
End Function

Private Sub refreshBulan(ByVal bulan As Date)
    txtBulan.Text = getBulanIndonesia(Month(bulan)) & " " & Year(bulan)
End Sub

Private Sub setHariLibur(ByVal hari As Integer)
    Dim x   As Integer
    Dim y   As Integer
    
    With gridKalender
        For x = 0 To .Cols - 1
            For y = 1 To .Rows - 1
                If Val(.TextMatrix(y, x)) = hari Then
                    .Col = x
                    .Row = y
                    
                    If Day(Now) = hari Then
                        .CellPictureAlignment = flexAlignCenterTop
                    Else
                        .CellPictureAlignment = flexAlignLeftTop
                    End If
                    
                    Set .CellPicture = LoadPicture(App.Path & "\smile.bmp")
                    
                    .CellFontBold = True
                    .CellForeColor = vbRed
                End If
            Next y
        Next x
    End With
End Sub

Private Function getAngkaByHari(ByVal hari As String) As Integer
    Select Case hari
        Case "Minggu": getAngkaByHari = 0
        Case "Senin": getAngkaByHari = 1
        Case "Selasa": getAngkaByHari = 2
        Case "Rabu": getAngkaByHari = 3
        Case "Kamis": getAngkaByHari = 4
        Case "Jum'at": getAngkaByHari = 5
        Case "Sabtu": getAngkaByHari = 6
    End Select
End Function

Private Function getRowByCell(ByVal cell As Integer) As Integer
    Select Case cell
        Case 1 To 7: getRowByCell = 1
        Case 8 To 14: getRowByCell = 2
        Case 15 To 21: getRowByCell = 3
        Case 22 To 28: getRowByCell = 4
        Case 29 To 35: getRowByCell = 5
        Case 36 To 42: getRowByCell = 6
        Case Else: getRowByCell = 1
    End Select
End Function

Private Function getColByCell(ByVal cell As Integer) As Integer
    Select Case cell
        Case 1, 8, 15, 22, 29, 36
            getColByCell = 0
            
        Case 2, 9, 16, 23, 30, 37
            getColByCell = 1
            
        Case 3, 10, 17, 24, 31, 38
            getColByCell = 2
            
        Case 4, 11, 18, 25, 32, 39
            getColByCell = 3
            
        Case 5, 12, 19, 26, 33, 40
            getColByCell = 4
            
        Case 6, 13, 20, 27, 34, 41
            getColByCell = 5
            
        Case 7, 14, 21, 28, 35, 42
            getColByCell = 6
    End Select
End Function

Private Sub setToDay(ByVal Col As Integer, ByVal Row As Integer)
    With gridKalender
        .Col = Col
        .Row = Row
        
        .CellPictureAlignment = flexAlignCenterTop
        Set .CellPicture = LoadPicture(App.Path & "\today.bmp")
        
        .CellFontBold = True
    End With
End Sub

Private Sub makeCalendar(ByVal jumlahHari As Integer, ByVal bulan As Integer, ByVal tahun As Integer)
    Dim hari        As Integer
    Dim y           As Integer
    Dim Index       As Integer
    Dim cell        As Integer
    
    Dim baris       As Integer
    Dim kolom       As Integer
    Dim ret         As Integer
    
    Dim str         As String
    Dim ketLibur    As String
    
    cell = 0
    lstKetHariLibur.Clear
    For hari = 1 To jumlahHari
        str = DOTW(hari, bulan, tahun)
        y = getAngkaByHari(str)
        
        For Index = cell To 41
            baris = getRowByCell(cell)
            kolom = getColByCell(cell)
            
            If kolom = y Then
                Index = 41
                gridKalender.TextMatrix(baris, kolom) = hari
                                                                    
                If Day(Now) = hari And Month(Now) = bulan Then Call setToDay(kolom, baris) 'setToDay -> prosedur untuk menampilkan icon today
                
                If kolom = 0 Then
                    Call setHariLibur(hari)
                Else
                    strSql = "SELECT COUNT(*) FROM hari_libur " & _
                             "WHERE DAY(tanggal) = " & hari & " AND " & _
                             "MONTH(tanggal) = " & bulan & " AND YEAR(tanggal) = " & tahun & ""
                    ret = CInt(dbGetValue(strSql, 0))
                    If ret > 0 Then
                        Call setHariLibur(hari)

                        strSql = "SELECT keterangan FROM hari_libur " & _
                                 "WHERE DAY(tanggal) = " & hari & " AND " & _
                                 "MONTH(tanggal) = " & bulan & " AND YEAR(tanggal) = " & tahun & ""
                        ketLibur = CStr(dbGetValue(strSql, ""))
                        lstKetHariLibur.AddItem hari & " : " & ketLibur
                    End If
                End If
                
            Else
                If baris > 0 And kolom > 0 Then gridKalender.TextMatrix(baris, kolom) = ""
            End If
            
            cell = cell + 1
        Next
    Next
End Sub

Private Sub resetKalender()
    Dim x   As Integer
    Dim y   As Integer
    
    With gridKalender
        For x = 0 To .Cols - 1
            For y = 1 To .Rows - 1
                .TextMatrix(y, x) = ""
                
                .Col = x
                .Row = y
                Set .CellPicture = Nothing
                
                .CellFontBold = False
                .CellForeColor = &H0& 'WARNA_HITAM
                .CellBackColor = &H80000005 'WARNA_PUTIH
            Next
        Next
    End With
End Sub

Private Function roundOff(ByVal num As Double) As Integer
    Dim str     As String
    Dim str2    As String
    Dim ctr     As Integer
    
    str = CStr(num)
    For ctr = 1 To Len(str)
        If Mid(str, ctr, 1) = "." Then
            roundOff = CInt(str2)
            Exit Function
        Else
            str2 = str2 & Mid(str, ctr, 1)
        End If
    Next
    
    roundOff = CInt(str2)
End Function

Private Function detrmMonth(ByVal bulan As Integer) As Integer
    Select Case bulan
        Case 1 'January
            If leap = True Then
                detrmMonth = 6
            Else
                detrmMonth = 0
            End If
            
        Case 2 'Febuary
            If leap = True Then
                detrmMonth = 2
            Else
                detrmMonth = 3
            End If
            
        Case 3 'March
            detrmMonth = 3
            
        Case 4 'April
            detrmMonth = 6
            
        Case 5 'May
            detrmMonth = 1
            
        Case 6 'June
            detrmMonth = 4
            
        Case 7 'July
            detrmMonth = 6
            
        Case 8 'August
            detrmMonth = 2
            
        Case 9 'September
            detrmMonth = 5
            
        Case 10 'October
            detrmMonth = 0
            
        Case 11 'November
            detrmMonth = 3
            
        Case 12 'December
            detrmMonth = 5
    End Select
End Function

Private Function getHariByAngka(ByVal hari As Integer) As String
    Select Case hari
        Case 0: getHariByAngka = "Minggu"
        Case 1: getHariByAngka = "Senin"
        Case 2: getHariByAngka = "Selasa"
        Case 3: getHariByAngka = "Rabu"
        Case 4: getHariByAngka = "Kamis"
        Case 5: getHariByAngka = "Jum'at"
        Case 6: getHariByAngka = "Sabtu"
    End Select
End Function

Private Function DOTW(ByVal hari As Integer, ByVal bulan As Integer, ByVal tahun As Integer) As String
    Dim yr      As Double
    Dim result  As Integer
    
    yr = tahun / 4
    result = roundOff(yr) + tahun
    
    yr = tahun / 100
    result = result - roundOff(yr)
    
    yr = tahun / 400
    result = result + roundOff(yr)
    result = result + hari
    result = result + detrmMonth(bulan)
    result = result - 1
    result = result Mod 7
    
    DOTW = getHariByAngka(result)
End Function

Private Sub genKalender()
    Dim jumlahHariByBulan   As Integer
    Dim num                 As Integer
            
    num = Year(setMonth) Mod 4
    If num = 0 Then
        leap = True
    Else
        leap = False
    End If
    
    Call resetKalender
    
    jumlahHariByBulan = getJumlahHariByBulan(Month(setMonth), Year(setMonth))
    Call makeCalendar(jumlahHariByBulan, Month(setMonth), Year(setMonth))
End Sub

Private Sub cmdNext_Click()
    setMonth = setNewMonth(True)
    Call refreshBulan(setMonth)
    Call genKalender
End Sub

Private Sub cmdPrev_Click()
    setMonth = setNewMonth(False)
    Call refreshBulan(setMonth)
    Call genKalender
End Sub

Private Sub Form_Load()
    mnuPopUp.Visible = False
    
    Call initHari
    Call initGrid
    
    setMonth = Date
    Call refreshBulan(setMonth)
    Call genKalender
End Sub

Private Sub gridKalender_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        With gridKalender
            If .MouseCol = 0 Then 'kolom hari minggu, semua menu dinonaktifkan
                mnuHariLibur.Enabled = False
                mnuHapusHariLibur.Enabled = False
                
            Else
                If Val(.TextMatrix(.MouseRow, .MouseCol)) > 0 Then
                    .Row = .MouseRow
                    .Col = .MouseCol
                
                    If .CellForeColor > 0 Then 'font warna merah, berarti hari libur
                        mnuHariLibur.Enabled = False 'menu hari libur dinonaktifkan
                        mnuHapusHariLibur.Enabled = True

                    Else
                        mnuHariLibur.Enabled = True
                        mnuHapusHariLibur.Enabled = False
                    End If

                Else
                    mnuHariLibur.Enabled = True
                    mnuHapusHariLibur.Enabled = False
                End If
            End If
        End With
        
        PopupMenu mnuPopUp
    End If
End Sub

Private Sub mnuHapusHariLibur_Click()
    Dim tanggal As String
    
    If MsgBox("Apakan Anda ingin menghapus hari libur ???", vbExclamation + vbYesNo, "Konfirmasi") = vbYes Then
        If Val(gridKalender.TextMatrix(gridKalender.Row, gridKalender.Col)) > 0 Then
            tanggal = Year(setMonth) & "/" & Month(setMonth) & "/" & Val(gridKalender.TextMatrix(gridKalender.Row, gridKalender.Col))
            
            strSql = "DELETE FROM hari_libur " & _
                     "WHERE tanggal = #" & tanggal & "#"
            conn.Execute strSql
            
            Call genKalender
            cmdNext.SetFocus
        End If
    End If
End Sub

Private Sub mnuHariLibur_Click()
    Dim inputKetLibur   As String
    Dim tanggal         As String
    Dim ret             As Integer
    
    inputKetLibur = InputBox("Keterangan Hari Libur", "Hari Libur")
    If Len(inputKetLibur) > 0 Then
        tanggal = Year(setMonth) & "/" & Month(setMonth) & "/" & Val(gridKalender.TextMatrix(gridKalender.Row, gridKalender.Col))
        
        strSql = "SELECT COUNT(*) FROM hari_libur " & _
                 "WHERE tanggal = #" & tanggal & "#"
        ret = CInt(dbGetValue(strSql, 0))
        If ret = 0 Then
            strSql = "INSERT INTO hari_libur(tanggal, keterangan) VALUES (#" & _
                     tanggal & "#,'" & inputKetLibur & "')"
            conn.Execute strSql
        End If
        
        Call genKalender
        cmdNext.SetFocus
    End If
End Sub
