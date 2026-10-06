VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmMescola 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Fattura"
   ClientHeight    =   7425
   ClientLeft      =   1350
   ClientTop       =   1530
   ClientWidth     =   10860
   Icon            =   "frmFattura.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7425
   ScaleWidth      =   10860
   Begin VB.ComboBox cmbMescola 
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   56
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CheckBox chkIVA 
      BackColor       =   &H00E0E0E0&
      Caption         =   "IVA 0%"
      Height          =   255
      Left            =   7200
      TabIndex        =   55
      Top             =   5400
      Width           =   855
   End
   Begin VB.TextBox txtNote 
      Height          =   735
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Text            =   "frmFattura.frx":0CCA
      Top             =   5760
      Width           =   7335
   End
   Begin VB.TextBox txtSconto 
      Height          =   285
      Left            =   1560
      TabIndex        =   7
      Text            =   "Sconto"
      Top             =   1800
      Width           =   585
   End
   Begin VB.TextBox txtIdCliente 
      Height          =   285
      Left            =   720
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   1560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox cmbProdotto 
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   8640
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   840
      Width           =   1935
   End
   Begin VB.ComboBox cmbCliente 
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   960
      Width           =   2295
   End
   Begin VB.Data dDettagli 
      Caption         =   "Righe"
      Connect         =   "Access"
      DatabaseName    =   "C:\Programmi\fatture\FattureRidotto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5040
      Width           =   2895
   End
   Begin MSDBGrid.DBGrid dbgDettagli 
      Bindings        =   "frmFattura.frx":0CD0
      Height          =   1575
      Left            =   2160
      OleObjectBlob   =   "frmFattura.frx":0CE8
      TabIndex        =   9
      Top             =   2880
      Width           =   8535
   End
   Begin VB.TextBox txtPagamento 
      Height          =   195
      Left            =   8400
      TabIndex        =   41
      Text            =   "Text1"
      Top             =   6360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox cmbPagamento 
      Height          =   315
      Left            =   7680
      TabIndex        =   11
      Top             =   6120
      Width           =   2415
   End
   Begin VB.TextBox txtOrdine 
      Height          =   285
      Left            =   10080
      TabIndex        =   5
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox txtRifORd 
      Height          =   285
      Left            =   7200
      TabIndex        =   4
      Text            =   "Text8"
      Top             =   360
      Width           =   2775
   End
   Begin VB.TextBox txtAnnoD 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "2001"
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox txtDDT 
      Height          =   285
      Left            =   4560
      TabIndex        =   2
      Text            =   "9999"
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox txtAnnoF 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "2001"
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox txtIdFattura 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "9999"
      Top             =   360
      Width           =   495
   End
   Begin MSMask.MaskEdBox mskData 
      Height          =   285
      Left            =   3000
      TabIndex        =   1
      Top             =   360
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yy"
      Mask            =   "99/99/99"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskDataDDT 
      Height          =   285
      Left            =   5880
      TabIndex        =   3
      Top             =   360
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yy"
      Mask            =   "99/99/99"
      PromptChar      =   "_"
   End
   Begin MSComctlLib.Toolbar tbCommand 
      Align           =   2  'Align Bottom
      Height          =   630
      Left            =   0
      TabIndex        =   42
      Top             =   6795
      Width           =   10860
      _ExtentX        =   19156
      _ExtentY        =   1111
      ButtonWidth     =   1244
      ButtonHeight    =   953
      Wrappable       =   0   'False
      Appearance      =   1
      ImageList       =   "ilToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Nuovo"
            Key             =   "N"
            Object.ToolTipText     =   "Nuova Fattura"
            ImageIndex      =   2
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "prova"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   300
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salva"
            Key             =   "S"
            Object.ToolTipText     =   "Salva Fattura Corrente"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   300
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Annulla"
            Key             =   "A"
            Object.ToolTipText     =   "Annulla modifiche su fattura corrente"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   300
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Modifica"
            Key             =   "M"
            Object.ToolTipText     =   "Modifica la fattura corrente"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   300
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Chiudi"
            Key             =   "C"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   300
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Stampa"
            Key             =   "P"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "phRight"
            Style           =   4
            Object.Width           =   500
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Esci"
            Key             =   "X"
            Object.ToolTipText     =   "Esci"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilToolBar 
      Left            =   10200
      Top             =   6000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFattura.frx":16BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFattura.frx":181B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFattura.frx":1D5F
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFattura.frx":22A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFattura.frx":27E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFattura.frx":2D2B
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFattura.frx":326F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Crystal.CrystalReport cryFattura 
      Left            =   10080
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin MSMask.MaskEdBox mskTotale 
      Height          =   285
      Left            =   120
      TabIndex        =   46
      Top             =   5040
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   393216
      Format          =   "L#,##0;(L#,##0)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskEuro 
      Height          =   285
      Left            =   9120
      TabIndex        =   47
      Top             =   5400
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   393216
      Format          =   "€#,##0.00;(€#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskScontato 
      Height          =   285
      Left            =   1800
      TabIndex        =   48
      Top             =   5040
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   393216
      Format          =   "L#,##0;(L#,##0)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskIVA 
      Height          =   285
      Left            =   7200
      TabIndex        =   50
      Top             =   5040
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   393216
      Format          =   "L#,##0;(L#,##0)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskFinale 
      Height          =   285
      Left            =   9120
      TabIndex        =   52
      Top             =   5040
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   393216
      Format          =   "L#,##0;(L#,##0)"
      PromptChar      =   "_"
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "per questa fattura"
      Height          =   195
      Left            =   240
      TabIndex        =   54
      Top             =   1845
      Width           =   1245
   End
   Begin VB.Label lblTotale 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Totale Fattura"
      Height          =   195
      Left            =   9120
      TabIndex        =   53
      Top             =   4800
      Width           =   990
   End
   Begin VB.Label lblIVA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IVA"
      Height          =   195
      Left            =   7320
      TabIndex        =   51
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label lblScontato 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Scontato"
      Height          =   195
      Left            =   1920
      TabIndex        =   49
      Top             =   4800
      Width           =   645
   End
   Begin VB.Label lblImponibile 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imponibile"
      Height          =   195
      Left            =   360
      TabIndex        =   45
      Top             =   4800
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente"
      Height          =   195
      Left            =   120
      TabIndex        =   44
      Top             =   720
      Width           =   480
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Note"
      Height          =   195
      Left            =   360
      TabIndex        =   43
      Top             =   5520
      Width           =   345
   End
   Begin VB.Label lblPagamento 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pagamento"
      Height          =   195
      Left            =   7680
      TabIndex        =   40
      Top             =   5880
      Width           =   810
   End
   Begin VB.Label lblESconto 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " % sconto"
      Height          =   195
      Left            =   240
      TabIndex        =   39
      Top             =   1470
      Width           =   690
   End
   Begin VB.Label lblCParentesi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ")"
      Height          =   195
      Left            =   8880
      TabIndex        =   38
      Top             =   1680
      Width           =   45
   End
   Begin VB.Label lblAParentesi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "("
      Height          =   195
      Left            =   8280
      TabIndex        =   37
      Top             =   1680
      Width           =   45
   End
   Begin VB.Label lblECitta 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Città"
      Height          =   195
      Left            =   6960
      TabIndex        =   36
      Top             =   1440
      Width           =   315
   End
   Begin VB.Label lblVirgola 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ","
      Height          =   195
      Left            =   5040
      TabIndex        =   35
      Top             =   1680
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Via"
      Height          =   195
      Left            =   2760
      TabIndex        =   34
      Top             =   1440
      Width           =   225
   End
   Begin VB.Image iLogo 
      Height          =   450
      Left            =   0
      Picture         =   "frmFattura.frx":37B3
      Top             =   0
      Width           =   450
   End
   Begin VB.Label lblSconto 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1560
      TabIndex        =   33
      Top             =   1440
      Width           =   585
   End
   Begin VB.Label lblProv 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Prov"
      Height          =   195
      Left            =   8400
      TabIndex        =   32
      Top             =   1680
      Width           =   330
   End
   Begin VB.Label lblCAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CAP"
      Height          =   195
      Left            =   9240
      TabIndex        =   31
      Top             =   1680
      Width           =   315
   End
   Begin VB.Label lblCitta 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Città"
      Height          =   195
      Left            =   6600
      TabIndex        =   30
      Top             =   1680
      Width           =   1395
   End
   Begin VB.Label lblNazione 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nazione"
      Height          =   195
      Left            =   9960
      TabIndex        =   29
      Top             =   1680
      Width           =   585
   End
   Begin VB.Label lblNumero 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Numero"
      Height          =   195
      Left            =   5400
      TabIndex        =   28
      Top             =   1680
      Width           =   555
   End
   Begin VB.Label lblCF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CF"
      Height          =   255
      Left            =   8400
      TabIndex        =   27
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label lblPIVA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PIVA"
      Height          =   255
      Left            =   5520
      TabIndex        =   26
      Top             =   1080
      Width           =   420
   End
   Begin VB.Label lblVia 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Via"
      Height          =   195
      Left            =   2760
      TabIndex        =   25
      Top             =   1680
      Width           =   225
   End
   Begin VB.Label lblCliente 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente"
      Height          =   315
      Left            =   2760
      TabIndex        =   24
      Top             =   1080
      Width           =   2520
   End
   Begin VB.Label lblBarra3 
      BackStyle       =   0  'Transparent
      Caption         =   "\"
      Height          =   255
      Left            =   9960
      TabIndex        =   23
      Top             =   375
      Width           =   255
   End
   Begin VB.Label lblOrdine 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rif.Vs. Ordine"
      Height          =   195
      Left            =   9000
      TabIndex        =   22
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblArticoli 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Elementi da produrre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4320
      TabIndex        =   21
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label lblDdtdat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data"
      Height          =   195
      Left            =   6120
      TabIndex        =   20
      Top             =   120
      Width           =   345
   End
   Begin VB.Label lblBarra2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "\"
      Height          =   195
      Left            =   5040
      TabIndex        =   19
      Top             =   375
      Width           =   75
   End
   Begin VB.Label lblDdtnum 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DDT"
      Height          =   195
      Left            =   4800
      TabIndex        =   18
      Top             =   120
      Width           =   345
   End
   Begin VB.Label lblFattdat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data"
      Height          =   195
      Left            =   3240
      TabIndex        =   17
      Top             =   120
      Width           =   345
   End
   Begin VB.Label lblBarra1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "\"
      Height          =   195
      Left            =   2160
      TabIndex        =   16
      Top             =   375
      Width           =   75
   End
   Begin VB.Label lblFattnum 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fattura n°"
      Height          =   195
      Left            =   1920
      TabIndex        =   14
      Top             =   120
      Width           =   690
   End
   Begin VB.Shape shpCliente 
      BackColor       =   &H80000004&
      BackStyle       =   1  'Opaque
      Height          =   975
      Left            =   2520
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Width           =   8175
   End
   Begin VB.Shape shpProdotto 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   2415
      Left            =   2040
      Shape           =   4  'Rounded Rectangle
      Top             =   2280
      Width           =   8775
   End
End
Attribute VB_Name = "frmMescola"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mKey As String, rsFattura As Recordset
Private fNuovoRecord As Boolean, sStato As String, fNuovaRiga As Boolean
Private mStatus As Integer, rsDettagli As Recordset
Private lInitW As Long, lInitH As Long, lDh As Long
Private fGetFromLoad As Boolean
Private lDx As Long
Const dimDettagli = 10215
Const DimDescrizione = 2800
Private Enum CMDS
   cNuovo = 1
   cSalva = 3
   cAnnulla = 5
   cModifica = 7
   cChiudi = 9
   cStampa = 11
   cEsci = 13
End Enum
Private Enum Status
   cNew = 0
   cMod = 1
   cExists = 2
   cNotExists = 3
End Enum
'  -----------------------------------------
Const lCambio = 1927.36
' ------------------------------------------
Dim rsMescola As Recordset







Public Property Let Status(vValue As Integer)
   If vValue >= cNew And vValue <= cNotExists Then
      mStatus = vValue
   End If
   With tbCommand
      Select Case mStatus
         Case cNew, cMod
            .Buttons(cNuovo).Enabled = False
            .Buttons(cSalva).Enabled = True
            .Buttons(cAnnulla).Enabled = True
            .Buttons(cModifica).Enabled = False
'            .Buttons(cChiudi).Enabled = False
            .Buttons(cStampa).Enabled = False
            .Buttons(cEsci).Enabled = False
         Case cExists
            .Buttons(cNuovo).Enabled = False
            .Buttons(cSalva).Enabled = False
            .Buttons(cAnnulla).Enabled = False
            .Buttons(cModifica).Enabled = True
'            .Buttons(cChiudi).Enabled = (sStato = "V")
            .Buttons(cStampa).Enabled = True
            .Buttons(cEsci).Enabled = True
         Case cNotExists
            .Buttons(cNuovo).Enabled = True
           
           ' Nascondo i pulsanti per caricare le righe della fattura
             Righe False

            .Buttons(cSalva).Enabled = False
            .Buttons(cAnnulla).Enabled = False
            .Buttons(cModifica).Enabled = False
'            .Buttons(cChiudi).Enabled = False
            .Buttons(cStampa).Enabled = False
            .Buttons(cEsci).Enabled = True
      End Select
   End With
'Private Sub NuovoReclamo()
'End Sub
End Property
Public Property Get Status() As Integer
   Status = mStatus
End Property

Public Property Let Key(vValue As String)
    mKey = Trim$(vValue)
'      CercaFattura
'        CercaMescola
    
    CaricodDettagli
   
End Property
Public Property Get Key() As String
    Key = mKey
End Property

Private Sub SalvaRec()
   Dim sCod As Variant
   If fNuovoRecord Then
      rsFattura.AddNew
      txtIdFattura.Text = CalcUltimaFattura()
      rsFattura("IdFattura") = txtIdFattura.Text
''      rsFattura("ODL") = Me.Key
   Else
      rsFattura.Edit
   End If
   If Not (IsNull(txtRifORd.Text) Or txtRifORd.Text = "") Then
        rsFattura("RifOrd") = txtRifORd.Text
   Else
         rsFattura("RifOrd") = Null
   End If
   If Not (IsNull(txtIdCliente.Text) Or txtIdCliente.Text = "") Then
        rsFattura("IdCliente") = txtIdCliente.Text
   Else
         rsFattura("IdCliente") = Null
   End If
   If Not (IsNull(mskData.Text) Or mskData.Text = "__/__/__") Then
      rsFattura("Data") = mskData.Text
   Else
       rsFattura("Data") = Null
   End If
   If Not (IsNull(txtSconto.Text) Or txtSconto.Text = "") Then
       rsFattura("PctSconto") = txtSconto
   Else
       rsFattura("PctSconto") = Null
   End If
   If Not (IsNull(txtPagamento.Text) Or txtPagamento.Text = "") Then
       rsFattura("Pagamento") = txtPagamento
   Else
       rsFattura("Pagamento") = Null
   End If
   If Not (IsNull(txtDDT.Text) Or txtDDT.Text = "") Then
       rsFattura("DDT") = txtDDT
   Else
       rsFattura("DDT") = Null
   End If
   If Not (IsNull(mskDataDDT.Text) Or mskDataDDT.Text = "__/__/__") Then
      rsFattura("DataDDT") = mskDataDDT.Text
   Else
       rsFattura("DataDDT") = Null
   End If
   If Not (IsNull(txtNote.Text) Or txtNote.Text = "") Then
       rsFattura("Note") = txtNote
   Else
       rsFattura("Note") = Null
   End If
   
   rsFattura.Update
'************************************************
   CercaFattura
'************** ***********************************+

'     disabilito i campi
   Campi False
   fNuovoRecord = False
End Sub

Private Sub chkIVA_Click()
'CalcolaSconto
End Sub

Private Sub cmbCliente_Click()
'Dim sCli As String
'sCli = Val(Left((cmbCliente.List(cmbCliente.ListIndex)), 4))
'Cliente (sCli)

'Reimposto la dimensione del combo
'cmbCliente.Width = 2300

End Sub
Private Sub Cliente(sTmp As String)
'   Dim rsTemp As Recordset, sqlc As String
''   CancellaCampi
'   lblCliente.Caption = sTmp
'   sqlc = "Select *   from [Cliente] where IdCliente=" & sTmp
'   Set rsTemp = db.OpenRecordset(sqlc, dbOpenSnapshot)'''

'   'Carico i campi del cliente
'   If Not rsTemp.EOF Then
'      If Not IsNull(rsTemp("Ragione")) Then
'         lblCliente.Caption = rsTemp("Ragione")
'      Else
'          lblCliente.Caption = ""
'      End If
'
'      If Not IsNull(rsTemp("PIVA")) Then
'         lblPIVA.Caption = rsTemp("PIVA")
'      Else
'          lblPIVA.Caption = ""
'      End If
'      lblPIVA.Left = lblCliente.Left + lblCliente.Width + 100
'
'      If Not IsNull(rsTemp("CF")) Then
'
'      lblCF.Caption = rsTemp("CF")
'      Else
'         lblCF.Caption = ""
'      End If
'      lblCF.Left = lblPIVA.Left + lblPIVA.Width + 100
'
'      If Not IsNull(rsTemp("Via")) Then
'         lblVia.Caption = rsTemp("Via")
'      Else
'         lblVia.Caption = ""
'      End If'

'      If Not IsNull(rsTemp("Numero")) Then
'         lblNumero.Caption = rsTemp("Numero")
'      Else
'          lblNumero.Caption = ""
'      End If
'      lblNumero.Left = lblVia.Left + lblVia.Width + 100
'      lblVirgola.Left = lblNumero.Left - 100
'
'      If Not IsNull(rsTemp("Citta")) Then
'         lblCitta.Caption = rsTemp("Citta")
'      Else
'         lblCitta.Caption = ""
'      End If
'      lblCitta.Left = lblNumero.Left + lblNumero.Width + 200
'      lblECitta.Left = lblCitta.Left
'
'      If Not IsNull(rsTemp("Provincia")) Then
'         lblProv.Caption = rsTemp("Provincia")
'      Else
'         lblProv.Caption = ""
'      End If
'
'      lblProv.Left = lblCitta.Left + lblCitta.Width + 200
'      lblAParentesi.Left = lblProv.Left - 100
'      lblCParentesi.Left = lblProv.Left + lblProv.Width + 100
'      If Not IsNull(rsTemp("CAP")) Then
''         lblCAP.Caption = rsTemp("CAP")
'      Else
'         lblCAP.Caption = ""
'      End If
'      lblCAP.Left = lblCParentesi.Left + 200
'      If Not IsNull(rsTemp("Nazione")) Then
'         lblNazione.Caption = rsTemp("Nazione")
'      Else
'         lblNazione.Caption = ""
'      End If
'      lblNazione.Left = lblCAP.Left + lblCAP.Width + 100
'
'      If Not IsNull(rsTemp("PctSconto")) Then
'         lblSconto.Caption = rsTemp("PctSconto")
'      Else
'         lblSconto.Caption = "0"
'      End If
'
'      If txtSconto.Text = "" Then
'         txtSconto.Text = lblSconto.Caption
'      End If
'      ' Carico i dati nella fattura
'      txtIdCliente.Text = rsTemp("IdCliente")
'   End If'

End Sub
Private Sub cmbCliente_DropDown()
'cmbCliente.Width = shpCliente.Width

End Sub

Private Sub cmbMescola_Click()
Dim sqlc As String, sDesc As String
Dim iRes As Integer, Elemento
   
   sDesc = Trim$(cmbMescola.Text)
   sqlc = "Select Caricata, Mescola, Data, Odl, ID, Qta  from Formule " & _
            "where Caricata = 'N' and  Mescola='" & sDesc & "'"
   Set rsMescola = db.OpenRecordset(sqlc, dbOpenDynaset)
    If Not rsMescola.EOF Then
        Set dDettagli.Recordset = rsMescola
    End If
    
End Sub

Private Sub cmbPagamento_Click()
'    txtPagamento.Text = cmbPagamento.Text
End Sub

Private Sub cmbProdotto_Click()
'Dim sqlc As String, rsProdotto As Recordset, sDesc As String
'Dim iRes As Integer
'
'   sDesc = Val(Left((cmbProdotto.Text), 3))
'   sqlc = "Select * from Prodotto where IdProdotto=" & sDesc
'   Set rsProdotto = db.OpenRecordset(sqlc, dbOpenDynaset)
'   If Not rsProdotto.EOF Then
'      dbgDettagli.AllowAddNew = True
'      dbgDettagli.Columns(2) = rsProdotto("IdProdotto")
'      dbgDettagli.Columns(3) = rsProdotto("Descrizione")
'      dbgDettagli.Columns(4) = rsProdotto("UM")
'      dbgDettagli.Columns(6) = rsProdotto("Prezzo")
'      dbgDettagli.Columns(7) = rsProdotto("IVA")
'   End If

End Sub



Private Sub dbgDettagli_AfterColEdit(ByVal ColIndex As Integer)
   Select Case ColIndex
      Case 0, 1, 2, 3, 4, 5, 6, 7
         Somme
      End Select
End Sub

Private Sub dbgDettagli_AfterUpdate()
   Totale
'         With dbgDettagli
 '           .Columns(8) = .Columns(5) * .Columns(6)
 '           .Columns(9) = .Columns(8) + (.Columns(8) * .Columns(7) / 100)
 '        End With
End Sub
Private Sub Totale()
Dim sqlc As String, rsTotale As Recordset, sVal As String

   mskTotale.Text = "0"
   mskIVA.Text = "0"
   If Me.Key > 0 Then
      sVal = Me.Key
   Else
      sVal = txtIdFattura.Text
   End If



     sqlc = "SELECT FD.Totale, FD.Ivato " & _
              "FROM FatturaDettagli AS FD where idFattura=" & sVal
    Set rsTotale = db.OpenRecordset(sqlc, dbOpenSnapshot)
       Do While Not rsTotale.EOF
            If Not (rsTotale("Totale") = 0 Or rsTotale("Ivato") = 0) Then
                mskTotale.Text = Val(mskTotale.Text) + Format((rsTotale("Totale")), "##########")
                mskIVA.Text = Val(mskIVA.Text) + Format(rsTotale("Ivato"), "###############")
    '    Stop
             rsTotale.MoveNext
             Else
                Exit Sub
            End If
             Loop
           
 
   
'   mskTotale.Text = Val(mskTotale.Text) + Val(dbgDettagli.Columns(8).Text)
   mskIVA.Text = Val(mskIVA.Text) - ((Val(mskIVA.Text) / 100) * (Val(txtSconto.Text))) - Val(mskScontato.Text)
End Sub

Private Sub dbgDettagli_Change()
   Somme
   Totale
End Sub
Private Sub Somme()
            With dbgDettagli
            If IsNull(.Columns(5)) Or .Columns(5) = "" Or IsNull(.Columns(6)) Or .Columns(6) = "" Then
               .Columns(8) = 1 * 1
            Else
               If IsNull(.Columns(5)) Or .Columns(5) = "" Then
                          .Columns(8) = 1 * .Columns(6)
               Else
                  If IsNull(.Columns(6)) Or .Columns(6) = "" Then
                    .Columns(8) = .Columns(5) * 1
                  Else
                    .Columns(8) = .Columns(5) * .Columns(6)
                  End If
               End If
               
                  If IsNull(.Columns(7)) Or .Columns(7) = "" Then
                     .Columns(9) = .Columns(8) + (.Columns(8) * 1 / 100)
                  Else
                  .Columns(9) = .Columns(8) + (.Columns(8) * .Columns(7) / 100)
            
                  End If
            End If

         End With

End Sub

Private Sub dbgDettagli_Click()
   ContaRighe
   'Somme
   Totale
End Sub

Private Sub dbgDettagli_OnAddNew()
   dbgDettagli.Columns(0) = txtIdFattura.Text
   Prog
   Somme
   Totale
'   mskTotale.Text = Val(mskTotale.Text) + Val(dbgDettagli.Columns(8).Text)
End Sub
Private Sub Prog()
    Dim sqlc As String, rsNum As Recordset, sVal As String
    If Me.Key > 0 Then
      sVal = Me.Key
   Else
      sVal = txtIdFattura.Text
    End If
    sqlc = "Select max([Prog]) from FatturaDettagli where IdFattura = " & sVal
    Set rsNum = db.OpenRecordset(sqlc, dbOpenSnapshot)
    If Not IsNull(rsNum(0)) Then
        dbgDettagli.Columns(1) = rsNum(0) + 1
    Else
        dbgDettagli.Columns(1) = 1
    End If

End Sub

Private Sub dDettagli_Reposition()
   ContaRighe
End Sub

Private Sub Form_Load()
   Dim btn As Button
   Dim sqlc As String, rsFattura As Recordset, rsMescola As Recordset
'Carico i dati per fare il resize
   fGetFromLoad = True
   lDx = dbgDettagli.Left - shpCliente.Left
   With Me
      lDh = shpCliente.Top
'      .Height = lDh + shpCliente.Height + 2 * lDh + _
            (.Height - .ScaleHeight)
 '     .Width = shpCliente.Left + shpCliente.Width + 2 * (.Width - .ScaleWidth)
      lInitW = .Width
      lInitH = .Height
      .MousePointer = vbArrowHourglass
      .Show
      .Refresh
   End With
   fGetFromLoad = False
   ShowWait "Attendere !!" & vbCrLf & vbCrLf & "Caricamento Clienti in corso"
'   sqlc = "Select *   from [Cliente] order by Ragione"
'   Set rsFattura = db.OpenRecordset(sqlc, dbOpenSnapshot)
'   Do While Not rsFattura.EOF
'      lstCliente.AddItem rsFattura(1)
'      cmbCliente.AddItem rsFattura(0) & "      " & rsFattura(1)
'      rsFattura.MoveNext
'   Loop
 '   cmbCliente.ListIndex = 0
'    cmbCliente.SetFocus
   Campi False
'___________________________________________________________________
' Carico il combo delle mescole
'___________________________________________________________________
   sqlc = "SELECT DISTINCT Formule.Mescola, Formule.Caricata " & _
            "From Formule " & _
            "Where Caricata = 'N' " & _
            "ORDER BY Formule.Mescola"

   Set rsMescola = db.OpenRecordset(sqlc, dbOpenSnapshot)
   Do While Not rsMescola.EOF
      cmbMescola.AddItem (rsMescola(0))
      rsMescola.MoveNext
   Loop
   cmbMescola.ListIndex = -1
'___________________________________________________________________
'Fine combo mescole
'___________________________________________________________________
   For Each btn In tbCommand.Buttons
     If btn.Key <> "X" Then
        btn.Enabled = False
     End If
   Next
   Me.MousePointer = vbNormal
   HideWait
End Sub

Private Sub Righe(Flag As Boolean)
   lblArticoli.Visible = Flag
   cmbProdotto.Visible = Flag
   dbgDettagli.Visible = Flag
   dDettagli.Visible = Flag

End Sub
Private Sub CaricodDettagli()
Dim sqlc As String, sPosition As String, sVal As String


   If Me.Key <> "" Then
      sVal = Me.Key
   Else
      sVal = txtIdFattura.Text
   End If
   
'      sqlc = "SELECT FD.IdFattura, FD.Prog, FD.IdProdotto, FD.Descrizione, " & _
               "FD.UM, FD.Qta, FD.Prezzo, FD.Iva, FD.Totale, FD.Ivato " & _
               "FROM FatturaDettagli AS FD where idFattura=" & sVal
 
'    sqlc = "Select * from Formule " & _
            "Where caricata='" & Me.Key
      sqlc = "SELECT Formule.Mescola, Formule.Data, Formule.OdL, " & _
                    "Formule.ID, Formule.CodArt, Formule.Qta " & _
                "FROM Formule INNER JOIN Descrizione ON " & _
                    "Formule.Caricata = Descrizione.Caricata " & _
                "Where Descrizione & 'A' & [Mescola]='" & Me.Key & "' " & _
                "Order by 1 asc, ODL asc, ID asc "
    
    Set rsDettagli = db.OpenRecordset(sqlc, dbOpenDynaset)
    If Not rsDettagli.EOF Then
        Set dDettagli.Recordset = rsDettagli
'      dDettagli.RecordSource = sqlc
'   Else
'      rsDettagli.AddNew
'      If Me.Key > 0 Then
'         rsDettagli("IdFattura") = Me.Key
'      Else
'         If fNuovaRiga Then
'            rsDettagli("IdFattura") = txtIdFattura.Text
'         Else
'            Exit Sub
'         End If
'      End If
'
'      rsDettagli("Prog") = "1"
'      rsDettagli("IdProdotto") = "1"
'      rsDettagli("Descrizione") = "Inserire il prodotto"
'      rsDettagli("UM") = "Pz"
'      rsDettagli("Qta") = "1"
'      rsDettagli("Prezzo") = "10"
'      rsDettagli("IVA") = "20"
'      rsDettagli.Update
'
'      Set dDettagli.Recordset = rsDettagli
''      dDettagli.RecordSource = sqlc
'   '      dbgDettagli.Columns(0).Caption = "test"
'
    End If

  



'   dbgDettagli.Columns(0).Visible = False
'   dbgDettagli.Columns(1).Width = 420
'   dbgDettagli.Columns(3).Width = 6300
'   dbgDettagli.Columns(4).Width = 345
'   dbgDettagli.Columns(5).Width = 420
'   dbgDettagli.Columns(6).Width = 950
'   dbgDettagli.Columns(6).NumberFormat = "currency"
'   dbgDettagli.Columns(2).Visible = False
'   dbgDettagli.Columns(8).Width = 950
''   dbgDettagli.Columns(8).NumberFormat = "currency"
'   dbgDettagli.Columns(9).Width = 950
'   dbgDettagli.Columns(9).NumberFormat = "currency"

'   dbgDettagli.Columns(8).Caption = dbgDettagli.Columns(5) * dbgDettagli.Columns(6)
'Trovo la posizione
   fNuovaRiga = False
   Somme
   Totale
   
End Sub
Private Sub ContaRighe()
 'rsDettagli.MoveLast
' rsDettagli.MoveFirst
 
 dDettagli.Caption = "Riga n." & rsMescola.AbsolutePosition & " di" & rsMescola.RecordCount
      
End Sub

Private Sub Form_Resize()
   Dim b As Button, l As Long, idx As Integer
      
   If Not fGetFromLoad Then
      With Me
         If .Width < lInitW Then .Width = lInitW
         If .Height < lInitH Then .Height = lInitH
         shpCliente.Width = .ScaleWidth - shpCliente.Left - 100
        
      End With
      With dbgDettagli
'         .Height = dDettagli.Top - .Top
         .Height = Me.Height - tbCommand.Height - .Top - (2.5 * txtNote.Height) - (2 * dDettagli.Height)
         .Width = shpCliente.Width + shpCliente.Left - .Left - 200
         If Me.Key > 0 Then
            .Columns(3).Width = .Columns(3).Width + (.Width - dimDettagli)
         End If
      End With
      If Me.Key > 0 Then
         If dbgDettagli.Columns(3).Width < DimDescrizione Then
            dbgDettagli.Columns(3).Width = DimDescrizione
         Else
            dbgDettagli.Columns(3).Width = dbgDettagli.Width - (dbgDettagli.Columns(0).Width + dbgDettagli.Columns(4).Width + dbgDettagli.Columns(5).Width + dbgDettagli.Columns(6).Width + dbgDettagli.Columns(7).Width + dbgDettagli.Columns(8).Width)
                                          
         End If
      End If
      With shpProdotto
         .Height = dbgDettagli.Height + (dbgDettagli.Top - .Top) + 200
         .Width = shpCliente.Width + shpCliente.Left - .Left
         
         
      End With
      
      
      With txtNote
         .Top = Me.Height - tbCommand.Height - (2 * .Height)
      End With
      lblNote.Top = txtNote.Top - (1.5 * lblNote.Height)
      
      cmbPagamento.Top = txtNote.Top + (txtNote.Height / 2)
      lblPagamento.Top = cmbPagamento.Top - (1.5 * lblPagamento.Height)
      
      
      With dDettagli
         .Top = Me.Height - .Height - tbCommand.Height - (2.5 * txtNote.Height)
         .Left = (dbgDettagli.Width / 2) - (.Width / 2)
      End With

      mskTotale.Top = Me.Height - mskTotale.Height - tbCommand.Height - (2.5 * txtNote.Height)
      lblImponibile.Top = mskTotale.Top - lblImponibile.Height - 50
      
      mskScontato.Top = Me.Height - mskScontato.Height - tbCommand.Height - (2.5 * txtNote.Height)
      lblScontato.Top = mskScontato.Top - lblScontato.Height - 50
      
      mskIVA.Top = Me.Height - mskIVA.Height - tbCommand.Height - (2.5 * txtNote.Height)
      mskIVA.Left = dbgDettagli.Width - mskFinale.Width - mskIVA.Width - 200
      lblIVA.Left = mskIVA.Left
      lblIVA.Top = mskIVA.Top - lblIVA.Height - 50
      
      mskFinale.Top = Me.Height - mskFinale.Height - tbCommand.Height - (2.5 * txtNote.Height)
      mskFinale.Left = dbgDettagli.Width - mskFinale.Width
      lblTotale.Top = mskFinale.Top - lblTotale.Height - 50
      lblTotale.Left = mskFinale.Left
      
      mskEuro.Top = Me.Height - mskEuro.Height + (2 * mskFinale.Height) + 100 - tbCommand.Height - (2.5 * txtNote.Height) - mskFinale.Height
      mskEuro.Left = dbgDettagli.Width - mskFinale.Width
      chkIVA.Top = mskEuro.Top
      chkIVA.Left = mskIVA.Left
      
      
      
      
      Me.Refresh
      
      
'Dispongo i pulsanti

      For Each b In tbCommand.Buttons
         If b.Caption <> "phRight" Then
            l = l + b.Width
         Else
            idx = b.Index
         End If
      Next
      Set b = tbCommand.Buttons(idx)
      b.Width = Me.ScaleWidth - l
      Set b = Nothing
   End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
'   rsFattura.Close
'   Set rsFattura = Nothing
   rsMescola.Close
    Set rsMescola = Nothing
End Sub
Private Sub CercaFattura()
    Dim sqlc As String, i As Integer
   If Me.Key > 0 Then
   
      sqlc = "SELECT * from Fattura where IdFattura=" & Me.Key
   Else
      sqlc = "SELECT * from Fattura where IdFattura=" & txtIdFattura.Text
   End If
   Set rsFattura = db.OpenRecordset(sqlc, dbOpenDynaset)
    
   If Not rsFattura.EOF Then
      FoundFattura
      Cliente (txtIdCliente.Text)
      
      Me.Status = cExists
    Else
      Me.Status = cNotExists
    End If
End Sub

Private Sub FoundFattura()
Dim sqlc As String, rsCliente As Recordset

    If Not IsNull(rsFattura("IdFattura")) Then
        txtIdFattura.Text = rsFattura("IdFattura")
    Else
        txtIdFattura.Text = ""
    End If
    
    If Not IsNull(rsFattura("RifOrd")) Then
        txtRifORd.Text = rsFattura("RifOrd")
    Else
        txtRifORd.Text = ""
    End If
    
    If Not IsNull(rsFattura("IdCliente")) Then
        txtIdCliente.Text = rsFattura("IdCliente")
    Else
        txtIdCliente.Text = ""
    End If
    
    If Not IsNull(rsFattura("Data")) Then
        mskData.Text = rsFattura("Data")
    Else
        mskData.Text = "__/__/__"
    End If
   If Not IsNull(rsFattura("PctSconto")) Then
        txtSconto.Text = rsFattura("PctSconto")
    Else
        txtSconto.Text = lblSconto.Caption
    End If
    If Not IsNull(rsFattura("Note")) Then
        txtNote.Text = rsFattura("Note")
    Else
        txtNote.Text = ""
    End If
 
    If Not IsNull(rsFattura("Pagamento")) Then
        txtPagamento.Text = rsFattura("Pagamento")
    Else
        txtPagamento.Text = ""
    End If
    'Carico il combo del pagamento
      cmbPagamento.Text = txtPagamento.Text
   
    If Not IsNull(rsFattura("DDT")) Then
        txtDDT.Text = rsFattura("DDT")
    Else
        txtDDT.Text = ""
    End If
    If Not IsNull(rsFattura("DataDDT")) Then
        mskDataDDT.Text = rsFattura("DataDDT")
    Else
        mskDataDDT.Text = "__/__/__"
    End If
      
 
    
End Sub
Private Sub Campi(Flag As Boolean)
    
      txtIdFattura.Enabled = False
      mskData.Enabled = Flag
      txtDDT.Enabled = Flag
      mskDataDDT.Enabled = Flag
      txtRifORd.Enabled = Flag
      txtOrdine.Enabled = Flag
      txtIdCliente.Enabled = Flag
      txtSconto.Enabled = Flag
      txtNote.Enabled = Flag
      txtPagamento.Enabled = Flag
      cmbProdotto.Enabled = Flag
      cmbCliente.Enabled = Flag
      cmbPagamento.Enabled = Flag
      dDettagli.Enabled = Flag
      dbgDettagli.AllowDelete = Flag
      dbgDettagli.AllowAddNew = Flag
      dbgDettagli.AllowUpdate = Flag
      dbgDettagli.Enabled = True
      
      
      chkIVA.Enabled = Flag
      
      mskTotale.Enabled = False
      mskScontato.Enabled = False
      mskIVA.Enabled = False
      mskFinale.Enabled = False
      mskEuro.Enabled = False
      
      
      txtIdFattura.Locked = Not Flag
      txtRifORd.Locked = Not Flag
      txtOrdine.Locked = Not Flag
      txtIdCliente.Locked = Not Flag
      txtSconto.Locked = Not Flag
      txtNote.Locked = Not Flag
      txtPagamento.Locked = Not Flag
      txtDDT.Locked = Not Flag
      cmbProdotto.Locked = Not Flag
      cmbCliente.Locked = Not Flag
      cmbPagamento.Locked = Not Flag
      
End Sub

Private Sub CancellaCampi()
    txtIdFattura.Text = ""
    txtRifORd.Text = ""
    txtIdCliente.Text = ""
    mskData.Text = "__/__/__"
    txtSconto.Text = ""
    txtNote.Text = ""
    txtPagamento.Text = ""
    txtDDT.Text = ""
    mskDataDDT.Text = "__/__/__"
    mskTotale.Text = ""
    mskScontato.Text = ""
    mskIVA.Text = ""
    mskFinale.Text = ""
    mskEuro.Text = ""
End Sub

Private Sub mskFinale_Change()
mskEuro.Text = Val(mskFinale.Text) / lCambio
End Sub

Private Sub mskIVA_Change()
mskFinale.Text = Val(mskScontato.Text) + Val(mskIVA.Text)
End Sub

Private Sub mskScontato_Change()
Dim iVal As Integer, sqlc As String, rsIva As Recordset

If chkIVA Then
   iVal = 0
Else
   iVal = 20
End If

''





'mskIVA.Text = Val(mskIVA.Text) - Val(mskScontato.Text)
End Sub

Private Sub mskTotale_Change()
CalcolaSconto

End Sub
Private Sub CalcolaSconto()
Dim lSconto As Long
If Val(txtSconto.Text) < 100 Then
   lSconto = Val(mskTotale.Text) / 100 * Val(txtSconto.Text)
Else
   MsgBox "Lo sconto non può essere superiore al 100%"
End If

mskScontato.Text = Val(mskTotale.Text) - lSconto

End Sub
Private Sub tbCommand_ButtonClick(ByVal Button As MSComctlLib.Button)
   Dim iRes As Integer, rsTmp As Recordset, sqlc As String
    Select Case Button.Key
        Case "X"
             Unload Me
        Case "N"
             fNuovoRecord = True
             Campi True
             CancellaCampi
             Me.Status = cNew
             fNuovaRiga = True
        Case "S"
            If CheckCampi Then
                  
               SalvaRec
               Righe True
               
                  If fNuovaRiga Then
                    Me.Status = cMod
                    Campi True
  '                  CaricodDettagli
' Provo ad inserire questa riga per verificare se così mi salva i dati
  '                  CaricodDettagli
' Funziona ************************************************************
                  Else
                     Me.Status = cExists
                     CaricodDettagli
                  End If
                
            End If
        Case "A"
            If fNuovoRecord Then
                CancellaCampi
                Me.Status = cNotExists
            
             Else
                FoundFattura
                Me.Status = cExists
             End If
             fNuovaRiga = False 'Per poter inserire nuovamente una nuova fattura
             
             Campi False
        Case "M"
             fNuovoRecord = False
             Campi True
             Me.Status = cMod
      Case "P"
         fNuovoRecord = False
         StampaFattura
         Me.Status = cExists
   End Select
End Sub
Private Function CalcUltimaFattura()
   Dim rsT As Recordset, sqlc As String
   sqlc = "update UltimaFattura " & _
               "set UltimaFattura = UltimaFattura + 1 " & _
               "Where dummy=0"
   db.Execute (sqlc)
   sqlc = "Select UltimaFattura from UltimaFattura " & _
               "Where dummy = 0"
   Set rsT = db.OpenRecordset(sqlc, dbOpenSnapshot)
   If Not rsT.EOF Then
      CalcUltimaFattura = rsT(0)
   Else
      Err.Raise 555, , "Errore nel calcolo del numero Risposta"
   End If
   rsT.Close
   Set rsT = Nothing
End Function
Private Sub StampaFattura()
   Dim sSelFormula As String, sIVA As String, i As Byte
   sIVA = Int(mskIVA.Text)
   
   With cryFattura
      On Error Resume Next
      .DataFiles(0) = db.Name
      .Destination = crptToWindow
      .ReportFileName = App.Path & "\fatturaI.rpt"
      .Formulas(i) = "IVA2= totext(" & sIVA & ",0)" 'Chr$(34)"
      .WindowState = crptNormal
      sSelFormula = "{Fattura.IdFattura} = " & txtIdFattura
'      .Formulas(1) = "DataOra=" & Chr$(34) & _
'          Format(Now, "dd/mm/yy  -  hh:mm") & Chr$(34)
   .SelectionFormula = sSelFormula
      .Action = 1
      If Err.Number <> 0 Then
          MsgBox "Errore nel produrre il report " & vbCrLf & _
              Err.Number & " - " & Err.Description
          Err.Clear
      End If
   End With
'   rsRete.Close
'   Set rsRete = Nothing
End Sub
Private Function CheckCampi() As Boolean
   
'Devo Controllare tutti i campi per la fattura

'*++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Dim fExit As Boolean, rsMax As Recordset
  Dim sqlc As String, rsData As Recordset, sData As Date, sFattura As String
   fExit = True
   If IsNull(txtIdCliente.Text) Or txtIdCliente.Text = "" Then
         MsgBox "Inserire il cliente a cui fatturare la merce"
         cmbCliente.SetFocus
         fExit = False
   End If
   If fExit Then
   
   
      If (IsNull(mskData.Text) Or mskData.Text = "__/__/__") _
            Or Not IsDate(mskData.Text) Then
         MsgBox "Data Errata"
         mskData.SetFocus
         fExit = False
       End If
   End If
  'Controllo l'esistenza della fattura altrimenti vedo il numero più alto
   If fExit Then
      If txtIdFattura <> "" Then
         sFattura = Val(txtIdFattura.Text) - 1
      Else
         sqlc = "SELECT Max(Fattura.IdFattura) AS MaxFattura " & _
                  "From Fattura"
         Set rsMax = db.OpenRecordset(sqlc, dbOpenSnapshot)
            If Not rsMax.EOF Then
               If Not IsNull(rsMax("MaxFattura")) Then
               
                  sFattura = rsMax("MaxFattura")
               Else
                  sFattura = 1
               End If
            End If
      End If
      
      sqlc = "SELECT Max(Fattura.Data) AS MaxDiData " & _
               "from Fattura where IdFattura=" & sFattura
      Set rsData = db.OpenRecordset(sqlc, dbOpenSnapshot)
         If Not rsData.EOF Then
            If Not IsNull(rsData("MaxDiData")) Then
               sData = rsData("MaxDiData")
            Else
               sData = "01/01/00"
            End If
            
         End If
         
     
     If mskData < sData Then
            MsgBox "La data non può essere anteriore alla fattura precedente"
            mskData.SetFocus
            fExit = False
         End If
   End If

   CheckCampi = fExit
End Function



Private Sub txtSconto_Change()
CalcolaSconto
End Sub

Private Sub CercaMescola()
    Dim sqlc As String, i As Integer
    If Me.Key > 0 Then
        sqlc = "Devo caricare l'sql dei componenti della formula e dei loro controtipi"
        
'      sqlc = "SELECT Formule.Mescola, Formule.Data, Formule.OdL, " & _
                    "Formule.ID, Formule.CodArt, Formule.Qta " & _
                "FROM Formule INNER JOIN Descrizione ON " & _
                    "Formule.Caricata = Descrizione.Caricata " & _
                "Where Descrizione & 'A' & [Mescola]='" & Me.Key & "' " & _
                "Order by 1 asc, ODL asc, ID asc "
   Else
      MsgBox " Questo non è un messaggio! CercaMescola con Me.key <0"
      
   End If
   Set rsFattura = db.OpenRecordset(sqlc, dbOpenDynaset)
    
   If Not rsFattura.EOF Then
      
      FoundFattura
      Cliente (txtIdCliente.Text)
      
      Me.Status = cExists
    Else
      Me.Status = cNotExists
    End If

End Sub
