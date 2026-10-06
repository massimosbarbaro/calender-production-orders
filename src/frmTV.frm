VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTV 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   7050
   ClientLeft      =   405
   ClientTop       =   1170
   ClientWidth     =   11250
   Icon            =   "frmTV.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   11250
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEsci 
      Height          =   585
      Left            =   5160
      Picture         =   "frmTV.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6240
      Width           =   585
   End
   Begin MSComctlLib.ImageList imglView 
      Left            =   10680
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTV.frx":0E14
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTV.frx":0F6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTV.frx":10C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTV.frx":12EA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10320
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTV.frx":173E
            Key             =   "D"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTV.frx":1B92
            Key             =   "F"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTV.frx":1FE6
            Key             =   "R"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTV.frx":243A
            Key             =   "V"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTV.frx":288E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTV.frx":2CE2
            Key             =   "C"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTV.frx":3142
            Key             =   "A"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTV.frx":35A2
            Key             =   "N"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTV.frx":3742
            Key             =   "P"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvReclami 
      Height          =   3015
      Left            =   6000
      TabIndex        =   1
      Top             =   1680
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   5318
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.TreeView tvReclami 
      Height          =   5655
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   9975
      _Version        =   393217
      HideSelection   =   0   'False
      Style           =   7
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.Image iLogo 
      Height          =   450
      Left            =   0
      Picture         =   "frmTV.frx":38E2
      Top             =   0
      Width           =   450
   End
   Begin VB.Label lblSaveODL 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblCursor 
      BackColor       =   &H000000C0&
      Height          =   3375
      Left            =   4680
      MousePointer    =   9  'Size W E
      TabIndex        =   2
      ToolTipText     =   "Click & Drag to resize"
      Top             =   1920
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      FillColor       =   &H00E0E0E0&
      Height          =   6135
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   10935
   End
End
Attribute VB_Name = "frmTV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lDx As Long
Private Const MinLblCursor = 75
Private lInitW As Long, lInitH As Long, lDh As Long
Private fGetFromLoad As Boolean
Private LastNode As Node
Private mChangedTV As Boolean, mParam As String
Private sNodeType As String, sWhereStato As String
' Aggiunta x Sort
Private sOrderBy As String
Private Const cDefOrderBy = " Order by 1 asc, ODL asc, ID asc "
Private mValore As String, sVal As String

' Fine Aggiunta x Sort


Public Property Get ChangedTV() As Boolean
   ChangedTV = mChangedTV
End Property
Public Property Let ChangedTV(vValue As Boolean)
   mChangedTV = vValue
   If mChangedTV Then
      SetTV
     mChangedTV = False
   End If
End Property
Public Property Let Param(vValue As String)
   If vValue = "ATC" Or vValue = "MKG" Or vValue = "ATC/MKG" Or vValue = "VIEW" Then
      mParam = vValue
      With Me
         .ChangedTV = True
         .Caption = "Visualizzazione del materiale in produzione "
      End With
   Else
      Err.Raise 17, , "Errore nel passaggio del paramtero al TV"
   End If
End Property
Public Property Get Param() As String
   Param = mParam
End Property

Private Sub SetTV()
   Dim sqlc As String, rsTV As Recordset, i As Byte
   Dim nodX As Node, nodY As Node
   Dim sStato As String
   Dim sLivello2 As String
   
   Screen.MousePointer = vbArrowHourglass


'm
    Dim rsLinea As Recordset

   Set LastNode = Nothing     ' Inizializzo comunque lastnode a niente
   
   With tvReclami
      .Nodes.Clear
      .ImageList = ImageList1
'
'  Crea i Nodi Principali (Root, ATC / MKG  e Stato)
'
      i = 1
      Set nodX = .Nodes.Add(, , "R", "Produzione", i)
      
'M Costriusco il nodo Linea in automatico

sqlc = "SELECT Descrizione, Ordine " & _
        "FROM Formule INNER JOIN Descrizione " & _
            "ON Formule.Caricata = Descrizione.Caricata " & _
        "GROUP BY Descrizione, Ordine " & _
        "ORDER BY Ordine"

Set rsLinea = db.OpenRecordset(sqlc, dbOpenSnapshot)
    'Costruisco i sottonodi livello2
    Do While Not rsLinea.EOF
 '       i = i + 1      'Questo serve per le immagini
                        'Ricordarsi di farlo alla fine!!!!!!!!!!!!!!
                        'Aggiungendo , i) alla fine della riga successiva
        Set nodX = .Nodes.Add("R", tvwChild, rsLinea(0) & "A", rsLinea(0), 8)
        rsLinea.MoveNext
    Loop
rsLinea.Close

'Inizio nodo livello 3 Famiglie
'sqlc = "SELECT Caricata & 'A' as Stato, Mescola " & _
        "From Formule " & _
        "GROUP BY Caricata & 'A', Mescola "
    
sqlc = "SELECT [Descrizione] & 'A' AS Stato, ucase(Formule.Mescola)as Mescola " & _
        "FROM Formule INNER JOIN Descrizione " & _
            "ON Formule.Caricata = Descrizione.Caricata " & _
        "GROUP BY [Descrizione] & 'A', Formule.Mescola "
    
    Set rsLinea = db.OpenRecordset(sqlc, dbOpenSnapshot)
'riempio il nodo famiglia
    Do While Not rsLinea.EOF
        sStato = rsLinea("Stato")
        sLivello2 = UCase(sStato & rsLinea("Mescola"))
'        If Not IsNull(rsLinea("Desc")) Then
           Set nodX = .Nodes.Add(sStato, tvwChild, sLivello2, rsLinea("Mescola"), 9)
'        End If
'        Debug.Print nodX.Key
        rsLinea.MoveNext
    Loop
rsLinea.Close


'---------------------------------------------------------------------------------------------------
'Qui si construiscono i nodi ATC e MKG

'  Nodo "ATC"
'      Set nodX = .Nodes.Add("R", tvwChild, "Linea", "Linea", 5)
      
'      Set nodX = .Nodes.Add("R", tvwChild, "ATC", "Linea", 5)
'
'      sqlc = "Select Stato, DescrStato " & _
'               "From StatoReclamo " & _
'               "ORDER BY Ordine"
'
'      Set rsTV = db.OpenRecordset(sqlc, dbOpenSnapshot)
'                        '
'                        '   Costruisco la condizione Where per gli stati (dipende dall'INI)
'                        '
'                        'M Carico nuovo, visto, chiuso di ATC????
'                        '''
'                        '''
'                            sWhereStato = " Where rv.Stato in ("
'                              Do While Not rsTV.EOF
'                                 i = i + 1
'                                 Set nodX = .Nodes.Add("ATC", tvwChild, rsTV(0) & "A", rsTV(1), i)
'                                 sWhereStato = sWhereStato & "'" & rsTV(0) & "',"
'                                 rsTV.MoveNext
'                              Loop
'                              sWhereStato = Left$(sWhereStato, Len(sWhereStato) - 1) & ") "
'                        ' La costringo a nullo per evitare di farla vedere
'                              sWhereStato = ""
'
'  Nodo "MKG"
'      Set nodX = .Nodes.Add("R", tvwChild, "MKG", "MKG", 5)
'
'      rsTV.MoveFirst
'      i = 1
'      Do While Not rsTV.EOF
'         i = i + 1
'         Set nodX = .Nodes.Add("MKG", tvwChild, rsTV(0) & "M", rsTV(1), i)
'         rsTV.MoveNext
'      Loop
'
'      rsTV.Close
'Fine Nodi ATC e MKG
'------------------------------------------------------------------------------------------
    
'Inizio nodo Venditore


'      Select Case Me.Param
'         Case "ATC"
'            sqlc = "SELECT Stato & 'A' as StatoR, Venditore " & _
'                     "From ReclamoVenditore as RV " & _
'                     sWhereStato & _
'                     "GROUP BY Stato, Venditore"
'         Case "MKG"
'            sqlc = "SELECT ATC.StatoReclamo  & 'M' as StatoR, rv.Venditore " & _
'                     "FROM ReclamoVenditore AS rv " & _
'                     "INNER JOIN ATC ON rv.ODL = ATC.ODL " & _
'                     sWhereStato & _
'                     "GROUP BY ATC.StatoReclamo, rv.Venditore " & _
'                     "Having ((Not (ATC.StatoReclamo) Is Null)) "
'         Case "ATC/MKG"
'            sqlc = "SELECT 'A' as AM, ([Stato] & 'A') as StatoR, Venditore " & _
'                     "From ReclamoVenditore AS rv " & _
'                     sWhereStato & _
'                     "GROUP BY Stato, Venditore " & _
'                     "Union All " & _
'                     "SELECT 'M' as AM, ATC.StatoReclamo  & 'M' as StatoR, " & _
'                     "rv.Venditore " & _
'                     "FROM ReclamoVenditore AS rv " & _
'                     sWhereStato & _
'                     "INNER JOIN ATC ON rv.ODL = ATC.ODL " & _
'                     "GROUP BY ATC.StatoReclamo, rv.Venditore " & _
'                     "Having ((Not (ATC.StatoReclamo) Is Null)) " & _
'                     "order by 1, 2, 3"
'      End Select
      
'      Set rsTV = db.OpenRecordset(sqlc, dbOpenSnapshot)
'
'     Aggiunge il nodo "Venditore"
'
'
'Fine nodo Livello3

'Inizio nodo livello 3 Formula
'sqlc = "SELECT [LdP] & 'A' & [ProdFin] & [desc] AS Linea, [Mescola] as formula , Standard " & _
        "From Generale " & _
        "GROUP BY [LdP] & 'A' & [ProdFin] & [desc], Generale.[Mescola], Generale.Standard"
    '
'    Set rsLinea = db.OpenRecordset(sqlc, dbOpenSnapshot)
'riempio il nodo famiglia
'    Do While Not rsLinea.EOF
'        sStato = UCase(rsLinea("Linea"))
'        If Not IsNull(rsLinea("Formula")) Then
'           On Error Resume Next
'           Set nodY = .Nodes.Add(sStato, tvwChild, , rsLinea("Formula"), 6)
'        End If
'        rsLinea.MoveNext
'    Loop
    
'    rsLinea.Close
    Set rsLinea = Nothing




'Fine Nodo Famiglia Livello3


'      .BorderStyle = 1
      
      .BorderStyle = vbFixedSingle
      .Checkboxes = False
      .FullRowSelect = True
      .HideSelection = False
      .HotTracking = True
      .Indentation = 100
      .LabelEdit = tvwManual
      .LineStyle = tvwTreeLines
      .Scroll = True
'      .SingleSel = True   Non so se mi piace !!
      'Se allo Style non metti niente il programma funziona con + e -
      .Style = tvwTreelinesPlusMinusPictureText
      
      With .Nodes("R")
         .Expanded = True
         .BackColor = vbYellow
         .Bold = True
      End With
'      With .Nodes("ATC")
'         .Expanded = True
'         .Bold = True
'         .BackColor = vbBlue
'      End With
'      With .Nodes("MKG")
'         .Expanded = True
'         .Bold = True
'         .BackColor = vbGreen
'      End With
'      With .Nodes("NA")
'         .Bold = True
'         .ForeColor = vbRed
'         .Selected = True
'      End With
   End With
'
'  Imposta la Header del List View
'
   SetLVHeader
   Screen.MousePointer = vbNormal

'   tvReclami.Nodes("NA").Selected = True
End Sub
Private Sub SetLVHeader()
   Dim clX As ColumnHeader
   With lvReclami
      .ColumnHeaders.Clear
      .View = lvwReport
      .Icons = imglView
      .SmallIcons = imglView
      Select Case Me.Param
         Case "ATC"
            Set clX = .ColumnHeaders.Add(, , "ID", 1000, lvwColumnLeft)
            Set clX = .ColumnHeaders.Add(, , "Data", 1000, lvwColumnCenter)
            Set clX = .ColumnHeaders.Add(, , "ODL", 1000, lvwColumnLeft)
            Set clX = .ColumnHeaders.Add(, , "SeId", 1000, lvwColumnLeft)
            Set clX = .ColumnHeaders.Add(, , "CodArt", 1565, lvwColumnRight)
            Set clX = .ColumnHeaders.Add(, , "Qta", 1000, lvwColumnRight)
            Set clX = .ColumnHeaders.Add(, , "Mescola", 1000, lvwColumnRight)
'            Set clX = .ColumnHeaders.Add(, , "Standard", 765, lvwColumnRight)
'            Set clX = .ColumnHeaders.Add(, , "Corrente", 765, lvwColumnRight)
'            Set clX = .ColumnHeaders.Add(, , "DCosto", 765, lvwColumnRight)
''            Set clX = .ColumnHeaders.Add(, , "DMA514£", 1065, lvwColumnRight)
'            Set clX = .ColumnHeaders.Add(, , "Data", 1000, lvwColumnRight)
'            Set clX = .ColumnHeaders.Add(, , "Altro", 1265, lvwColumnRight)


         Case Else
            Set clX = .ColumnHeaders.Add(, , "LdP", 1261, lvwColumnLeft)
            Set clX = .ColumnHeaders.Add(, , "ProdFin", 1200, lvwColumnCenter)
            Set clX = .ColumnHeaders.Add(, , "Desc", 2190, lvwColumnLeft)
            Set clX = .ColumnHeaders.Add(, , "Codice Art", 1000, lvwColumnLeft)
            Set clX = .ColumnHeaders.Add(, , "%MA514", 1065, lvwColumnRight)
            Set clX = .ColumnHeaders.Add(, , "Standard", 1065, lvwColumnRight)
            Set clX = .ColumnHeaders.Add(, , "Altro", 1265, lvwColumnRight)
'         Case "ATC/MKG"
'            Set clX = .ColumnHeaders.Add(, , "ATC/MKG ODL", 1261, lvwColumnLeft)
'            Set clX = .ColumnHeaders.Add(, , "Data", 1200, lvwColumnCenter)
'            Set clX = .ColumnHeaders.Add(, , "Cliente", 2190, lvwColumnLeft)
'            Set clX = .ColumnHeaders.Add(, , "Prodotto", 2190, lvwColumnLeft)
'            Set clX = .ColumnHeaders.Add(, , "Qta Spedita", 1065, lvwColumnRight)
'           Set clX = .ColumnHeaders.Add(, , "Qta Reclamata", 1265, lvwColumnRight)
        End Select
      .Arrange = lvwAutoLeft
   End With
End Sub

Private Sub cmdEsci_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
   Dim nodX As Node
'Per Giudo Così Funziona
   With tvReclami
         On Error Resume Next
   
      If .SelectedItem <> "" Then
'      On Error Resume Next
         Set nodX = .Nodes(.SelectedItem.Index)
      Else
         Set nodX = .Nodes(1)    '  Nuovi
      End If
   End With
   tvReclami_NodeClick nodX
End Sub

Private Sub Form_Click()
'   Dim i As Integer
'   Dim strNodes As String
'   For i = 1 To tvReclami.Nodes.Count
'   strNodes = strNodes & tvReclami.Nodes(i).Index & " " & _
'               "Key: " & tvReclami.Nodes(i).Key & " " & _
'               "Text: " & tvReclami.Nodes(i).Text & vbLf
'   Next i
'   MsgBox strNodes
End Sub


Private Sub Form_Load()
   Dim sqlc As String
   
   Screen.MousePointer = vbArrowHourglass
   
   fGetFromLoad = True
   
   lblSaveODL.Visible = False
   lDx = tvReclami.Left - Shape1.Left
   With Me
      lDh = Shape1.Top
      .Height = lDh + Shape1.Height + 2 * lDh + cmdEsci.Height + _
            (.Height - .ScaleHeight)
      .Width = Shape1.Left + Shape1.Width + 2 * (.Width - .ScaleWidth)
      lInitW = .Width
      lInitH = .Height
   End With
   fGetFromLoad = False
'
'   Me.ChangedTV = True
'
'   Me.Show
'
'  Simulo un Click sul nodo "Nuovi"
'
'   Set nodX = tvReclami.Nodes("N")
'   tvReclami_NodeClick nodX
   Screen.MousePointer = vbNormal
' Aggiunta x Sort
' "order by  rv.odl asc, Soluzione asc, DataReclamo desc "
   sOrderBy = cDefOrderBy
' Fine Aggiunta x Sort
End Sub

Private Sub Form_Resize()
   If Not fGetFromLoad Then
      With Me
         If .Width < lInitW Then .Width = lInitW
         If .Height < lInitH Then .Height = lInitH
         Shape1.Width = .ScaleWidth - 2 * Shape1.Left
         Shape1.Height = .ScaleHeight - (3 * lDh + cmdEsci.Height)
         lvReclami.Width = Shape1.Width - lDx - lvReclami.Left
      End With
      With tvReclami
         .Height = Shape1.Height - (2 * (tvReclami.Top - Shape1.Top))
      End With
      With lvReclami
         .Left = tvReclami.Left + tvReclami.Width + MinLblCursor
         .Top = tvReclami.Top
         .Height = tvReclami.Height
         .Width = Shape1.Width + Shape1.Left - lDx - .Left
      End With
      With cmdEsci
         .Top = Me.ScaleHeight - (.Height + lDh)
         .Left = Shape1.Left + Shape1.Width - .Width
      End With
'      With lblExplain
'         .Top = cmdEsci.Top + (cmdEsci.Height - .Height) \ 2
'         .Left = lvReclami.Left + (lvReclami.Width - .Width) \ 2
'      End With
   End If
   With lblCursor
'      .Height = tvReclami.Height * 1.01
     .Height = tvReclami.Height * 0.99
     .Width = (tvReclami.Width / 2) + (lvReclami.Width / 2) + MinLblCursor
     .Top = tvReclami.Top * 1.01
     .Left = tvReclami.Left + (tvReclami.Width / 2)
     .BackColor = Shape1.BackColor
   End With
End Sub

Private Sub lblCursor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = vbLeftButton Then
      If X < MinLblCursor Then
         X = MinLblCursor
      End If
      If X > lblCursor.Width Then
         X = lblCursor.Width
      End If
      lvReclami.Left = lblCursor.Left + X
      lvReclami.Width = Shape1.Width - lDx - lvReclami.Left
      tvReclami.Width = (lblCursor.Left + X) - tvReclami.Left - MinLblCursor
   End If
End Sub


Private Sub lvReclami_BeforeLabelEdit(Cancel As Integer)
'   MsgBox "Non è possibile modificare i valori.", vbOKOnly + vbExclamation, ""
'   Beep
   Cancel = True
End Sub

Private Sub lvReclami_DblClick()
   Dim sEntrata As String
   If lblSaveODL.Caption <> "" Then
    Select Case sNodeType
        Case "Prod"
            Load frmSE
            frmSE.Key = lblSaveODL.Caption & ", " & mValore
            frmSE.Show
'            Load frmID
'            frmID.Key = lblSaveODL.Caption & mValore
'            frmID.Show
        
        Case "Linea"
            MsgBox "Selezionare prima il nome di una mescola.", _
            vbOKOnly + vbExclamation
'            Load frmCosti
'            frmCosti.Key = lblSaveODL.Caption & mValore
'            frmCosti.Show
        Case "Famiglia"
'            sEntrata = Left$(mValore, 5)
            sEntrata = Left$(mValore, Len(mValore) - 5)
            'Per ADP il codice delle formule è composto da 5 lettere
            'perciò il valore deve esser abbassato di 6
            'sEntrata = Left$(mValore, Len(mValore) - 6)

            Select Case sEntrata
                Case "NUOVA"
                    SpostaID
              
                Case "DA FARE"
                    Load frmID
                    frmID.Key = lblSaveODL.Caption & ", " & mValore
                    frmID.Show
                
                Case "PRONTA"
                    Load frmBilance
                    frmBilance.Key = lblSaveODL.Caption & ", " & mValore
                    frmBilance.Show
                Case "FATTA"
                    Load frmID
                    frmID.Key = lblSaveODL.Caption & ", " & mValore
                    frmID.Show
                Case "TRASFERITA"
                    Load frmID
                    frmID.Key = lblSaveODL.Caption & ", " & mValore
                    frmID.Show
             '       MsgBox "FineTrasferita"
            End Select
            
'            If sEntrata = "NUOVA" Then
'
'            Else
'                Load frmID
'                frmID.Key = lblSaveODL.Caption & mValore
'                frmID.Show
'
'            End If
'Da attivare solo per la lettura da parte degli operatori alla bilancia

'            Load frmMescola
'            frmMescola.Key = lblSaveODL 'mValore
'            frmMescola.Show
'_____________________________________________________________________
        Case Else
            MsgBox "Questo non è un messaggio"
    End Select
'     Case "ATC"
'      If Me.Param = "ATC" Then
'       Debug.Print
'         Load frmCosti
'         frmCosti.Key = lblSaveODL.Caption
'         frmCosti.Show
 ''     Case "MKG"
 '        Load frmMKG
 '        frmMKG.Key = lblSaveODL.Caption
 '        frmMKG.Show
 '     Case Else
 '      MsgBox "Questo non è un messaggio"
 '     End Select
      
   Else
      MsgBox "Nessun ODL Selezionato.", _
            vbOKOnly + vbExclamation
   End If
End Sub

Private Sub lvReclami_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'    Stop
' Aggiunta x Sort
    Select Case ColumnHeader.Index
        Case 1, 2  ' ordinamenti possibili: colonne 1 e 2 (ODL, Data)
            sOrderBy = " Order by " & Trim$(Str$(ColumnHeader.Index)) & " "
'   Se si vuole lasciare l'ordinamento anche per le altre colonne
        Case 3 To 100  ' ordinamenti possibili: colonne 1 e 2 (ODL, Data)
            sOrderBy = " Order by " & Trim$(Str$(ColumnHeader.Index)) & " "
        Case Else   ' se seleziono un'altra colonna --> reset ordine iniziale (1)
            sOrderBy = cDefOrderBy
    End Select
    'Modifica per errore icona nodo
    
'    tvReclami_NodeClick LastNode
    tvReclami_NodeClick tvReclami.SelectedItem
'   Fine modifica per errore icona
' Fine Aggiunta x Sort
    
End Sub

Private Sub lvReclami_ItemClick(ByVal Item As MSComctlLib.ListItem)
   lblSaveODL.Caption = Trim$(Item.Text)
   

End Sub


Private Sub tvReclami_Collapse(ByVal Node As MSComctlLib.Node)
   If Node.Index = 1 Then
      Node.Expanded = True   ' Expand the node again.
      tvReclami_NodeClick Node
   End If
End Sub

Private Sub tvReclami_Expand(ByVal Node As MSComctlLib.Node)
'   Debug.Print "Expanded " & Node.Text
End Sub

Private Sub tvReclami_NodeClick(ByVal Node As MSComctlLib.Node)
   Dim sqlc As String, rsLV As Recordset, iImg As Integer
   Dim itmX As ListItem, i As Integer
   Dim nodeX As Node
   Dim sTabella As String
   lblSaveODL.Caption = ""    '  Elimino l'ultimo odl selezionato
'
'   Pulisco comunque la parte ListView
'
    lvReclami.ListItems.Clear

   If Node.Index = 1 Then  ' Si tratta del nodo Root
'      Set nodeX = tvReclami.Nodes("NA")
'      nodeX.Selected = True
'      tvReclami_NodeClick nodeX
   Else
      If Node.Key = "" Then      '  Nodo Foglia
         Node.Image = "A"
         If Not (LastNode Is Nothing) Then
            If Node <> LastNode Then
               LastNode.Image = "C"
            End If
         End If
         Set LastNode = Node
         
         
         
         
'M Seleziono se fai il clic sul nodo ATC=A oppure MKG
         
         If Len(Node.Parent.Key) < 7 And Right$(Node.Parent.Key, 1) = "A" Then
'         If Me.Param = "ATC" Then
'            sqlc = "Select rv.ODL, DataReclamo, Cliente, Prodotto, " & _
                     "Qta, QtaReclamata, " & _
                     "Stato, Soluzione " & _
                     "FROM (MFG INNER JOIN ReclamoVenditore AS rv ON MFG.ODL = rv.ODL) " & _
                     "LEFT JOIN ATC ON MFG.ODL = ATC.ODL " & _
                     "Where Stato = '" & Left$(Trim$(Node.Parent.Key), 1) & "' " & _
                     "and Venditore = '" & Trim$(Node.Text) & "' " & _
                     sOrderBy
'                     "order by  rv.odl asc, Soluzione asc, DataReclamo desc "
            sNodeType = "ATC"
         Else
            
            'SQLC per MKG
           'M uso qusto come se fosse un if singolo
           'perchè node.parent.key non sarà mai A
            sqlc = "SELECT LdP, ProdFin, Desc, [Codice Art], CstTot, " & _
                         "[LdP] & 'A' & [ProdFin] & [desc] AS Famiglia " & _
                        "From qTot " & _
                        "WHERE ([LdP] & 'A' & [ProdFin] & [desc])='" & Trim$(Node.Parent.Key) & "' " & _
                        "and [Codice art]='" & Trim$(Node.Text) & "' " & _
                        sOrderBy
            'M Applico i costi delle formule
'
'            sqlc = "SELECT Generale.Mescola, Generale.ProdFin, Generale.Desc, " & _
                        " Generale.DescMescola, Spessore.DA, " & _
                        "Spessore.A, Generale.KgMa514, Generale.Standard, " & _
                        "MescoleMaxCostoCRR.Corrente, Format([Corrente]-[Standard],'0.00') " & _
                        "AS DMa514, MescoleMaxCostoCRR.MaxDiData AS Data " & _
                    "FROM (Generale LEFT JOIN MescoleMaxCostoCRR ON " & _
                        "Generale.Mescola = MescoleMaxCostoCRR.Mescola) " & _
                        "LEFT JOIN Spessore ON Generale.Mescola = Spessore.Mesc " & _
                     "WHERE ([LdP] & 'A' & [ProdFin] & [desc])='" & Trim$(Node.Parent.Key) & "' " & _
                        "and Generale.Mescola='" & Trim$(Node.Text) & "' " & _
                    sOrderBy
             mValore = Trim$(Node.Parent.Key)
          sNodeType = "Prod"
         End If
'                  "order by DataReclamo desc "
'  Ora è ordinato per odl in modo tale da poter digitare l'odl
'


      Else
         If Node.Key <> "ATC" And Node.Key <> "MKG" Then
            If Len(Node.Key) < 12 And Right$(Node.Key, 1) = "A" Then
            ' M Inserisco l'sqlc della linea di prodotto
'                sqlc = "SELECT Generale.Mescola, Generale.ProdFin, Generale.Desc, " & _
                            " Generale.DescMescola, " & _
                            "Spessore.DA, Spessore.A, Generale.KgMa514, Generale.Standard " & _
                        "FROM Generale LEFT JOIN Spessore ON Generale.Mescola = Spessore.Mesc " & _
                        "Where LdP = '" & Left(Trim$(Node.Key), 4) & "' " & _
                        sOrderBy
                sqlc = "SELECT Formule.ID, Formule.Data, Formule.OdL, " & _
                            "Formule.SeID, Formule.CodArt, Formule.Qta, ucase(Formule.Mescola) as Mescola " & _
                        "FROM Formule INNER JOIN Descrizione " & _
                            "ON Formule.Caricata = Descrizione.Caricata " & _
                        "Where Descrizione='" & Left(Trim$(Node.Key), Len(Node.Key) - 1) & "' " & _
                        sOrderBy
            
            'M Fine ____________________________________________
             mValore = Trim$(Node.Key)
             sNodeType = "Linea"
            Else
            'M Inserisco l'sqlc della Famiglia
'                sqlc = "SELECT Generale.Mescola, Generale.ProdFin, Generale.Desc, " & _
                        " Generale.DescMescola, Spessore.DA, " & _
                        "Spessore.A, Generale.KgMa514, Generale.Standard, " & _
                        "MescoleMaxCostoCRR.Corrente, Format([Corrente]-[Standard],'0.00') " & _
                        "AS DMa514, MescoleMaxCostoCRR.MaxDiData AS Data " & _
                    "FROM (Generale LEFT JOIN MescoleMaxCostoCRR ON " & _
                        "Generale.Mescola = MescoleMaxCostoCRR.Mescola) " & _
                        "LEFT JOIN Spessore ON Generale.Mescola = Spessore.Mesc " & _
                        "WHERE ([LdP] & 'A' & [ProdFin] & [desc])='" & Trim$(Node.Key) & "' " & _
                        sOrderBy
                
                sqlc = "SELECT Formule.ID, Formule.Data, Formule.OdL, Formule.SeID, " & _
                            " Formule.CodArt, Formule.Qta, ucase(Formule.Mescola) as Mescola " & _
                        "FROM Formule INNER JOIN Descrizione " & _
                            "ON Formule.Caricata = Descrizione.Caricata " & _
                        "Where Descrizione& 'A' & [Mescola]='" & (Node.Key) & "' " & _
                        sOrderBy
            
                
            'M Fine ___________________________________________________
               
'               sqlc = "SELECT rv.ODL, ATC.DataRisposta, " & _
                        "MFG.Cliente, MFG.Prodotto, " & _
                        "Qta, QtaReclamata, " & _
                        "ATC.StatoReclamo AS Stato, Soluzione " & _
                        "FROM (ATC INNER JOIN MFG ON ATC.ODL = MFG.ODL) " & _
                        "INNER JOIN ReclamoVenditore AS rv ON ATC.ODL = rv.ODL " & _
                        "Where ATC.StatoReclamo = '" & Left(Trim$(Node.Key), 1) & "' " & _
                        sOrderBy
'                        "order by  rv.odl asc, Soluzione asc, DataReclamo desc "
             mValore = Trim$(Node.Key)
             sNodeType = "Famiglia"
             
            End If
         Else
            sqlc = ""
         End If
      End If


      With lvReclami
         If sqlc <> "" Then
            Set rsLV = db.OpenRecordset(sqlc, dbOpenDynaset)
            Do While Not rsLV.EOF
               Select Case rsLV("Mescola")
                  Case "C"
                     iImg = 1
                  Case "N"
                     iImg = 2
                  Case "V"
                     iImg = 3
                  Case Else
                    iImg = 3
               End Select
' Carico una icona differente per le segnalazioni
   
               'If rsLV("Soluzione") = 2 Then
               '   iImg = 4
               'End If
'**************************************************************
               Set itmX = .ListItems.Add(, , rsLV(0), , iImg)
               With itmX
'M Seleziono chi entra per far vedere una cosa o l'altra
            Select Case Me.Param
                Case "ATC"
                    sTabella = rsLV.Fields.Count - 1
                    
                    Select Case sNodeType
                        Case "Prod"
                        sTabella = rsLV.Fields.Count - 1
                        Case "Linea"
                        sTabella = rsLV.Fields.Count - 1
                        Case "Famiglia"
                        sTabella = rsLV.Fields.Count - 1
                    End Select
                  Case Else
                    If sNodeType <> "Prod" Then
                        sTabella = rsLV.Fields.Count - 1
                    Else
                        sTabella = rsLV.Fields.Count - 5
                    End If
                
            End Select
                For i = 1 To sTabella
                     If Not IsNull(rsLV(i)) Then
                        itmX.SubItems(i) = CStr(rsLV(i))
                     Else
                        itmX.SubItems(i) = " "
                     End If
                  Next i
               .Ghosted = False
               
               End With
               rsLV.MoveNext
            Loop
            rsLV.Close
            Set rsLV = Nothing
         End If
      End With
   End If
   
' Modifica per errore icona nel nodo
'Set LastNode = Node
'   Fine modifica


End Sub
Private Sub SpostaID()
'Si vuole fare in modo che al doppio clic su un Id di una mescola si sposti da
'Nuovo in Da Fare
'Dim sqlc As String, rsId As Recordset
Dim iRes As Integer, sqlc As String
        iRes = MsgBox("Volete mettere in preparazione l'ID" & lblSaveODL.Caption, _
               vbYesNo + vbDefaultButton2 + vbQuestion)
        If iRes = vbYes Then
            sqlc = "Update Formule set Caricata = 'S' " & _
                     "Where ID = '" & lblSaveODL.Caption & "' and " & _
                     "Caricata = 'N'"
            db.Execute (sqlc)
            AssegnaSe
'            sStatoReclamo = "C"
'            Me.Status = cExists
            frmTV.ChangedTV = True
        End If

    
    
End Sub

Private Sub AssegnaSe()
Dim sqlc As String, rsMp As Recordset, rsS As Recordset
Dim sMP As String, sSe As String

sqlc = "SELECT Formule.Mescola, Formule.ID " & _
        "From Formule " & _
        "WHERE ID = '" & lblSaveODL.Caption & "'"
Set rsMp = db.OpenRecordset(sqlc, dbOpenSnapshot)
    If Not rsMp.EOF Then
        sMP = rsMp("Mescola")
    End If
sqlc = "SELECT Formule.Mescola, Formule.ID, Formule.Caricata, Formule.SeID " & _
        "From Formule " & _
        "Where Formule.Mescola = '" & sMP & "' And Formule.Caricata = 's'"

Set rsS = db.OpenRecordset(sqlc, dbOpenDynaset)
    Do While Not rsS.EOF
        If Not IsNull(rsS("SeID")) And rsS("SeId") <> "" Then
            sSe = rsS("SeID")
        End If
    rsS.MoveNext
    Loop
    
    rsS.MoveFirst
        
    If IsNull(sSe) Or sSe = "" Then
        CreaSe
        sSe = sVal
    End If
    Do While Not rsS.EOF
        If IsNull(rsS("SeID")) Then
            rsS.Edit
            rsS("SeID") = sSe
            rsS.Update
        End If
        
    rsS.MoveNext
    Loop



'Chiudo i rs

rsMp.Close
Set rsMp = Nothing
rsS.Close
Set rsS = Nothing


End Sub


Private Sub CreaSe()
Dim rsNum As Recordset, sqlc As String

'   sqlc = "update UltimaRisposta " & _
               "set UltimaRisposta = UltimaRisposta + 1 " & _
               "Where dummy=0"
'   db.Execute (sqlc)
'    Dim sqlc As String, rsNum As Recordset
    sqlc = "Select max(SeId) as MSE  from SE "
    Set rsNum = db.OpenRecordset(sqlc, dbOpenSnapshot)
    If Not rsNum.EOF Then
        If Not IsNull(rsNum("MSE")) And rsNum("MSE") <> "" Then
            sVal = rsNum("MSE")
        Else
            sVal = 0
        End If
        
    Else
        sVal = 0
    End If
    sVal = sVal + 1
    sqlc = "insert into SE " & _
            "(seID, SeS,Data, Qta)  " & _
            "values ('" & sVal & " ', 'N', Now, '0')"
            

   db.Execute (sqlc)
    
   rsNum.Close
   Set rsNum = Nothing
End Sub






