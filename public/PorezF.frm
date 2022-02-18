VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPorez 
   Caption         =   "Form1"
   ClientHeight    =   7875
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11115
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7875
   ScaleWidth      =   11115
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   7575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   13361
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   873
      BackColor       =   8388608
      ForeColor       =   16711680
      TabCaption(0)   =   "Pregled"
      TabPicture(0)   =   "PorezF.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Image2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Data1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Obrada"
      TabPicture(1)   =   "PorezF.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Image1"
      Tab(1).Control(1)=   "Label1"
      Tab(1).Control(2)=   "IDTARIFA"
      Tab(1).Control(3)=   "PDV"
      Tab(1).ControlCount=   4
      Begin VB.TextBox PDV 
         Height          =   495
         Left            =   -72360
         TabIndex        =   4
         Top             =   2160
         Width           =   2655
      End
      Begin VB.TextBox IDTARIFA 
         Height          =   495
         Left            =   -72360
         TabIndex        =   2
         Top             =   1440
         Width           =   2535
      End
      Begin MSDataGridLib.DataGrid Data1 
         Height          =   5415
         Left            =   0
         TabIndex        =   1
         Top             =   600
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   9551
         _Version        =   393216
         AllowUpdate     =   -1  'True
         Appearance      =   0
         ForeColor       =   16711680
         HeadLines       =   1
         RowHeight       =   19
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial CE"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "idtarifa"
            Caption         =   "idtarifa"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1050
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "pdv"
            Caption         =   "pdv"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1050
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
      Begin MSForms.Image Image2 
         Height          =   6735
         Left            =   0
         Top             =   600
         Width           =   10215
         SizeMode        =   1
         Size            =   "18018;11880"
         Picture         =   "PorezF.frx":0038
      End
      Begin VB.Label Label1 
         Caption         =   "OZNAKA"
         Height          =   495
         Left            =   -74520
         TabIndex        =   3
         Top             =   1440
         Width           =   1815
      End
      Begin MSForms.Image Image1 
         Height          =   6495
         Left            =   -74760
         Top             =   840
         Width           =   10095
         BorderStyle     =   0
         SizeMode        =   1
         Size            =   "17806;11456"
         Picture         =   "PorezF.frx":57E0
         PictureAlignment=   0
      End
   End
End
Attribute VB_Name = "frmPorez"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rec1 As ADODB.Recordset

Private Sub Data1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Call VuciPod

End Sub

Private Sub Form_Activate()
Dim strSQL As String


strSQL = "Select * From porezneg"

Set Rec1 = New ADODB.Recordset
Rec1.Open strSQL, Con, adOpenStatic, adLockOptimistic

If Not Rec1.EOF Then

Set Data1.DataSource = Rec1
Data1.ReBind

Data1.Height = (Rec1.RecordCount + 1) * 290

End If




End Sub

Private Sub Form_Load()
Me.SSTab1.Top = 0
Me.SSTab1.Left = 0
Me.SSTab1.Width = MDIForm1.Width - 400

Me.SSTab1.Height = MDIForm1.Height - 3000



End Sub

Private Function VuciPod()
Me.IDTARIFA = Rec1!IDTARIFA
Me.PDV = Rec1!PDV


 
End Function

Private Sub Form_Resize()
'MsgBox Me.SSTab1.Width
Me.Image2.Width = Me.SSTab1.Width
Me.Image2.Height = Me.SSTab1.Height




End Sub
