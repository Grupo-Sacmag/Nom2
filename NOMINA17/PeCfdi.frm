VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form PeCfdi 
   Caption         =   "Datos complementarios de Personal"
   ClientHeight    =   5952
   ClientLeft      =   144
   ClientTop       =   444
   ClientWidth     =   9564
   LinkTopic       =   "Form9"
   ScaleHeight     =   5952
   ScaleWidth      =   9564
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox LTXT 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   7092
   End
   Begin MSFlexGridLib.MSFlexGrid PCpCfdi 
      Height          =   4932
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   9012
      _ExtentX        =   15896
      _ExtentY        =   8700
      _Version        =   393216
   End
   Begin VB.Menu ArCf 
      Caption         =   "&Archivo"
      Begin VB.Menu ArCfAbr 
         Caption         =   "&Abrir"
      End
      Begin VB.Menu ArCfGuarda 
         Caption         =   "&Guardar"
      End
      Begin VB.Menu ArCfsep1 
         Caption         =   "-"
      End
      Begin VB.Menu ArCfSale 
         Caption         =   "&Salida"
      End
   End
   Begin VB.Menu EdCf 
      Caption         =   "&Editar"
      Begin VB.Menu EdCfPegar 
         Caption         =   "&Pegar"
         Shortcut        =   ^V
      End
      Begin VB.Menu EdCfCopia 
         Caption         =   "&Copiar"
         Shortcut        =   ^C
      End
      Begin VB.Menu EdCfSep1 
         Caption         =   "-"
      End
      Begin VB.Menu EdCfInsertar 
         Caption         =   "&Insertar"
      End
      Begin VB.Menu EdCfEliminar 
         Caption         =   "&Eliminar"
      End
      Begin VB.Menu EdCfSep2 
         Caption         =   "-"
      End
      Begin VB.Menu EdCfSelT 
         Caption         =   "&Seleccionar Todo"
      End
   End
End
Attribute VB_Name = "PeCfdi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim temporal, Temporal1
Sub Col_Def()
PCpCfdi.Clear: PCpCfdi.Rows = 200: PCpCfdi.Cols = 9: PCpCfdi.FixedCols = 0
   PCpCfdi.FixedRows = 1
   PCpCfdi.Row = 0
   PCpCfdi.Col = 0: PCpCfdi.CellAlignment = 4: PCpCfdi.ColWidth(0) = 800: PCpCfdi.Text = "Num"
   PCpCfdi.Col = 1: PCpCfdi.CellAlignment = 4: PCpCfdi.ColWidth(1) = 3600: PCpCfdi.Text = "Nombre"
   PCpCfdi.Col = 2: PCpCfdi.CellAlignment = 4: PCpCfdi.ColWidth(2) = 2400: PCpCfdi.Text = "Direccion"
   PCpCfdi.Col = 3: PCpCfdi.CellAlignment = 4: PCpCfdi.ColWidth(3) = 2400: PCpCfdi.Text = "Colonia"
   PCpCfdi.Col = 4: PCpCfdi.CellAlignment = 4: PCpCfdi.ColWidth(4) = 2400: PCpCfdi.Text = "Ciudad"
   PCpCfdi.Col = 5: PCpCfdi.CellAlignment = 4: PCpCfdi.ColWidth(5) = 2400: PCpCfdi.Text = "Estado"
   PCpCfdi.Col = 6: PCpCfdi.CellAlignment = 4: PCpCfdi.ColWidth(6) = 2400: PCpCfdi.Text = "Delegacion"
   PCpCfdi.Col = 7: PCpCfdi.CellAlignment = 4: PCpCfdi.ColWidth(7) = 800: PCpCfdi.Text = "Codigo"
   PCpCfdi.Col = 8: PCpCfdi.CellAlignment = 4: PCpCfdi.ColWidth(8) = 2400: PCpCfdi.Text = "Correo "
   'pcpcfdi.Col = 9: pcpcfdi.CellAlignment = 4: pcpcfdi.ColWidth(9) = 600: pcpcfdi.Text = "Cons."
   
   
End Sub

Private Sub ArCfAbr_Click()
   Dim EmpNum As Long, Nombre As String, Cadena1
   PCpCfdi.Rows = 1
    Close 1, 2
    Open "Personal.dno" For Random As 2 Len = Len(personal)
    cm = LOF(2) / Len(personal)
    Open "Perscfdi.dno" For Random As 1 Len = Len(Empleado_1)
    
    For i = 1 To cm: Get 2, i, personal: Get 1, i, Empleado_1
    
    If IsNumeric(Empleado_1.Cpostal) Then
           Nombre = (Trim(personal.nom) + " " + Trim(personal.ape1) + " " + Trim(personal.ape2))
           Cadena1 = Format(i, "####0") & Chr(9) & Nombre & Chr(9) & Empleado_1.Direccion _
                     & Chr(9) & Empleado_1.Colonia & Chr(9) & Empleado_1.Ciudad _
                     & Chr(9) & Empleado_1.Estado & Chr(9) & Empleado_1.Delegacion _
                     & Chr(9) & Empleado_1.Cpostal & Chr(9) & Empleado_1.correo
          PCpCfdi.AddItem Cadena1
    End If
   Next i
End Sub

Private Sub ArCfGuarda_Click()
   Dim EmpNum As Long
    Close 1
    Open "Perscfdi.dno" For Random As 1 Len = Len(Empleado_1)
    cm = LOF(1) / Len(Empleado_1)
    For renglon = 1 To PCpCfdi.Rows - 1
    If IsNumeric(PCpCfdi.TextMatrix(renglon, 0)) Then
                EmpNum = PCpCfdi.TextMatrix(renglon, 0)
                Empleado_1.Direccion = PCpCfdi.TextMatrix(renglon, 2)
                Empleado_1.Colonia = PCpCfdi.TextMatrix(renglon, 3)
                Empleado_1.Ciudad = PCpCfdi.TextMatrix(renglon, 4)
                Empleado_1.Estado = PCpCfdi.TextMatrix(renglon, 5)
                Empleado_1.Delegacion = PCpCfdi.TextMatrix(renglon, 6)
                Empleado_1.Cpostal = PCpCfdi.TextMatrix(renglon, 7)
                Empleado_1.correo = PCpCfdi.TextMatrix(renglon, 8)
                Put 1, EmpNum, Empleado_1
                
    End If
    Next renglon
    Close 1
    
End Sub

Private Sub ArCfSale_Click()
   Close
   Unload PeCfdi
End Sub

Private Sub EdCfCopia_Click()
      Dim Temporal1
 Clipboard.Clear
   
   difer = PCpCfdi.RowSel - PCpCfdi.Row
   For i = PCpCfdi.Row To PCpCfdi.RowSel
      
      For f = PCpCfdi.Col To PCpCfdi.ColSel
            Temporal1 = Temporal1 + PCpCfdi.TextMatrix(i, f)
            If f < PCpCfdi.ColSel Then
                Temporal1 = Temporal1 & Chr(9)
            End If
      Next f
      Temporal1 = Temporal1 & Chr(13) & Chr(10)
      
   Next i
    Clipboard.SetText Temporal1

End Sub

Private Sub EdCfEliminar_Click()
 Dim w1 As Long
  For w1 = PCpCfdi.Row To PCpCfdi.RowSel
            PCpCfdi.RemoveItem PCpCfdi.Row
   Next w1

End Sub

Private Sub EdCfInsertar_Click()
  Dim W As Long
      For W = PCpCfdi.Row To PCpCfdi.RowSel
     PCpCfdi.AddItem (""), PCpCfdi.Row
    Next W

End Sub

Private Sub EdCfPegar_Click()
  Dim DeAqui As Integer, RetornoCarro As Long, InicioCopia As Long
  Temporal1 = ""
  temporal = Clipboard.GetText(vbCFText)

  RetornoCarro = PCpCfdi.Col
  InicioCopia = PCpCfdi.Row: DeAqui = InicioCopia
  InIPGr = InicioCopia: InIPGc = RetornoCarro
If temporal <> "" Then
  Clipboard.Clear
  DeAqui = 1
For i = 1 To Len(temporal)
    Select Case Mid(temporal, i, 1)
          Case Chr(9)
          LTXT.Text = Mid(temporal, DeAqui, (i - DeAqui))
          Temporal1 = Temporal1 + PCpCfdi.TextMatrix(PCpCfdi.Row, PCpCfdi.Col) & Chr(9)
          PCpCfdi.Text = Mid(temporal, DeAqui, (i - DeAqui))
          PCpCfdi.Col = PCpCfdi.Col + 1
          DeAqui = i + 1
          Case Chr(13)
          Temporal1 = Temporal1 + PCpCfdi.TextMatrix(PCpCfdi.Row, PCpCfdi.Col) & Chr(13)
          LTXT.Text = Mid(temporal, DeAqui, (i - DeAqui))
          PCpCfdi.Text = Mid(temporal, DeAqui, (i - DeAqui))
          If (PCpCfdi.Rows - 1) <= PCpCfdi.Row Then PCpCfdi.Rows = PCpCfdi.Rows + 1:
          PCpCfdi.Row = PCpCfdi.Row + 1: PCpCfdi.TopRow = 1
          DeAqui = i + 1
          Case Chr(10)
          Temporal1 = Temporal1 & Chr(10)
          PCpCfdi.Col = RetornoCarro
          DeAqui = i + 1
          Case Else
          Rem nada
    End Select
 Next i
 Rem pcpcfdi.Row = InicioCopia:
 PCpCfdi.Col = RetornoCarro
End If
 
 Clipboard.Clear
 Clipboard.SetText Clipboard.GetText + Temporal1
End Sub

Private Sub EdCfSelT_Click()
    Dim limite As Long
    Clipboard.Clear
    PCpCfdi.Row = 1: PCpCfdi.Col = 0
   For limite = 1 To PCpCfdi.Rows - 1
       renglon = limite
    If IsNumeric(PCpCfdi.TextMatrix(renglon, 0)) Then
           PCpCfdi.RowSel = PCpCfdi.Row
    End If
   Next limite
    PCpCfdi.ColSel = PCpCfdi.Cols - 1
End Sub

Private Sub Form_Load()
    Col_Def
    
End Sub


Private Sub LTXT_Change()
    PCpCfdi.Text = LTXT.Text
End Sub
Private Sub pcpcfdi_EnterCell()
    PCpCfdi.CellBackColor = vbYellow
    valcelant = PCpCfdi.Text
    LTXT.Text = valcelant
    Rem If pcpcfdi.Row > 25 Then pcpcfdi.TopRow = pcpcfdi.TopRow + 1
    
End Sub

Private Sub pcpcfdi_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
            Case vbKeyDelete
               
               For Q = PCpCfdi.Row To PCpCfdi.RowSel
                 
                 For W = PCpCfdi.Col To PCpCfdi.ColSel
                    
                    PCpCfdi.TextMatrix(Q, W) = ""
                 Next W
               Next Q
               
                LTXT.Text = PCpCfdi.Text
            Case vbKeyF2
                If PCpCfdi.Text <> "" Then valcelant = PCpCfdi.Text
                LTXT.Text = LTrim(RTrim(PCpCfdi.Text))
                LTXT.SetFocus
               
       End Select

End Sub

Private Sub pcpcfdi_KeyPress(KeyAscii As Integer)
    valcelant = PCpCfdi.Text
    LTXT.Text = Chr(KeyAscii)
    LTXT.SetFocus
End Sub

Private Sub pcpcfdi_LeaveCell()
   PCpCfdi.CellBackColor = vbWhite
End Sub
