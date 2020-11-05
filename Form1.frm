VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9990
   LinkTopic       =   "Form1"
   ScaleHeight     =   5625
   ScaleWidth      =   9990
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Hapus"
      Height          =   375
      Left            =   4200
      TabIndex        =   13
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Edit"
      Height          =   375
      Left            =   2880
      TabIndex        =   12
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Input"
      Height          =   375
      Left            =   1440
      TabIndex        =   11
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   1440
      TabIndex        =   10
      Text            =   "Text5"
      Top             =   1560
      Width           =   4335
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Text            =   "Text4"
      Top             =   1200
      Width           =   4335
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Text            =   "Text3"
      Top             =   840
      Width           =   4335
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   480
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   120
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   5106
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   2
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   7800
      Top             =   2280
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Telepon"
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Alamat"
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Jenis Kelamin"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nama"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NIK"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim KoneksiDB As New ADODB.Connection
Dim RSKaryawan As ADODB.Recordset
Dim RSJabatan As ADODB.Recordset
Sub BukaDB()
Set KoneksiDB = New ADODB.Connection
Set RSKaryawan = New ADODB.Recordset
KoneksiDB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DBPenggajian.mdb;"
End Sub

Private Sub Command1_Click()
    If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Then
        MsgBox "Silahkan isi data terlebih dahulu!"
    Else
        Call BukaDB
        Dim TambahData
        TambahData = "Insert into TBL_KARYAWAN values('" & Text1 & "','" & Text2 & "','" & Text3 & "','" & Text4 & "','" & Text5 & "')"
        KoneksiDB.Execute TambahData
        MsgBox "Tambah Data Berhasil"
        Call KosongkanData
        Call AmbilData
    End If
End Sub

Private Sub Command2_Click()
    If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Then
        MsgBox "Pastikan data tidak kosong"
    Else
        Call BukaDB
        Dim UpdateData
        UpdateData = "update TBL_KARYAWAN SET NamaKaryawan = '" & Text2 & "' where NIK = '" & Text1 & "' "
        KoneksiDB.Execute UpdateData
        MsgBox "Update Data Berhasil"
        Call KosongkanData
        Call AmbilData
    End If
End Sub

Private Sub Command3_Click()
    Call BukaDB
    Dim HapusData
    Dim Jawab As String
    Jawab = MsgBox("Apa anda yakin?", vbQuestion + vbYesNo)
    If Jawab = vbYes Then
        HapusData = "delete from TBL_KARYAWAN where NIK = '" & Text1 & "'"
        KoneksiDB.Execute HapusData
        MsgBox "Hapus Data Berhasil"
        Call KosongkanData
        Call AmbilData
    End If
End Sub

Private Sub Form_Load()
Call KosongkanData
Call AmbilData
End Sub
Sub AmbilData()
    Call BukaDB
    Adodc1.ConnectionString = KoneksiDB
    Adodc1.RecordSource = "TBL_KARYAWAN"
    Adodc1.Refresh
    Set DataGrid1.DataSource = Adodc1
End Sub
Sub KosongkanData()
    Text1 = ""
    Text2 = ""
    Text3 = ""
    Text4 = ""
Text5 = ""
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call BukaDB
        RSKaryawan.Open "Select * From TBL_KARYAWAN where NIK = '" & Text1 & "'", KoneksiDB
        If Not RSKaryawan.EOF Then
            Text2 = RSKaryawan!NamaKaryawan
            Text3 = RSKaryawan!JenisKelamin
            Text4 = RSKaryawan!AlamatKaryawan
            Text5 = RSKaryawan!TeleponKaryawan
        Else
        End If
    End If
End Sub

Private Sub DataGrid1_Click()
    Text1.Text = DataGrid1.Columns(0).Text
    Text2.Text = DataGrid1.Columns(1).Text
    Text3.Text = DataGrid1.Columns(2).Text
    Text4.Text = DataGrid1.Columns(3).Text
    Text5.Text = DataGrid1.Columns(4).Text
End Sub
