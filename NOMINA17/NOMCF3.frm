VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form NOMCF3 
   Caption         =   "Generar CFDI"
   ClientHeight    =   4995
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   11760
   LinkTopic       =   "Form9"
   ScaleHeight     =   4995
   ScaleWidth      =   11760
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox LTXT 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   0
      Width           =   2415
   End
   Begin MSFlexGridLib.MSFlexGrid NomCfdi 
      Height          =   4695
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   8281
      _Version        =   393216
      AllowUserResizing=   3
   End
   Begin VB.Menu Archivo 
      Caption         =   "Archivo"
      Begin VB.Menu Copiar 
         Caption         =   "Copiar"
         Checked         =   -1  'True
         Shortcut        =   ^C
      End
      Begin VB.Menu seleccionar 
         Caption         =   "Seleccionar todo"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu Excel 
      Caption         =   "Generar Excel"
   End
End
Attribute VB_Name = "NOMCF3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub IniCols()

    Open "DatCFDI.dno" For Random As 7 Len = Len(DATcf)
    cm = LOF(7) / Len(DATcf)
    'Close 7
   NomCfdi.Clear: NomCfdi.Rows = 2: NomCfdi.Cols = cm: NomCfdi.FixedCols = 0
   NomCfdi.FixedRows = 1
   NomCfdi.Rows = Form8.ConNom1.Rows
   NomCfdi.Row = 0
   For I7 = 0 To cm - 1: Get 7, (I7 + 1), DATcf
      
   NomCfdi.Col = I7: NomCfdi.CellAlignment = 4: NomCfdi.ColWidth(I7) = 1200: NomCfdi.Text = Trim(DATcf.Concepto)
        'nomcfdi.Col = 1: nomcfdi.CellAlignment = 4: nomcfdi.ColWidth(1) = 3600: nomcfdi.Text = "Nombre"
        'nomcfdi.Col = 2: nomcfdi.CellAlignment = 4: nomcfdi.ColWidth(2) = 2400: nomcfdi.Text = "Direccion"
        'nomcfdi.Col = 3: nomcfdi.CellAlignment = 4: nomcfdi.ColWidth(3) = 2400: nomcfdi.Text = "Colonia"
        'nomcfdi.Col = 4: nomcfdi.CellAlignment = 4: nomcfdi.ColWidth(4) = 2400: nomcfdi.Text = "Ciudad"
        'nomcfdi.Col = 5: nomcfdi.CellAlignment = 4: nomcfdi.ColWidth(5) = 2400: nomcfdi.Text = "Estado"
        'nomcfdi.Col = 6: nomcfdi.CellAlignment = 4: nomcfdi.ColWidth(6) = 2400: nomcfdi.Text = "Delegacion"
        'nomcfdi.Col = 7: nomcfdi.CellAlignment = 4: nomcfdi.ColWidth(7) = 800: nomcfdi.Text = "Codigo"
        'nomcfdi.Col = 8: nomcfdi.CellAlignment = 4: nomcfdi.ColWidth(8) = 2400: nomcfdi.Text = "Correo "
        'nomcfdi.Col = 9: nomcfdi.CellAlignment = 4: nomcfdi.ColWidth(9) = 600: nomcfdi.Text = "Cons."
  Next I7
  Close 7
End Sub

Private Sub Excel_Click()
    If Exportar_Excel(App.Path & "\libro1.xls", NomCfdi) Then
        MsgBox " Datos exportados en " & App.Path, vbInformation
    End If
End Sub

Public Function Exportar_Excel(sOutputPath As String, FlexGrid As Object) As Boolean
  
    On Error GoTo Error_Handler
  
    Dim o_Excel     As Object
    Dim o_Libro     As Object
    Dim o_Hoja      As Object
    Dim fila        As Long
    Dim Columna     As Long
      
    ' -- Crea el objeto Excel, el objeto workBook y el objeto sheet
    Set o_Excel = CreateObject("Excel.Application")
    Set o_Libro = o_Excel.Workbooks.Add
    Set o_Hoja = o_Libro.Worksheets.Add
      
    ' -- Bucle para Exportar los datos
    With FlexGrid
        For fila = 1 To .Rows - 1
            For Columna = 0 To .Cols - 1
                o_Hoja.Cells(fila, Columna + 1).Value = .TextMatrix(fila, Columna)
            Next
        Next
    End With
    o_Libro.Close True, sOutputPath
    ' -- Cerrar Excel
    o_Excel.Quit
    ' -- Terminar instancias
    Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
    Exportar_Excel = True
Exit Function
  
' -- Controlador de Errores
Error_Handler:
    ' -- Cierra la hoja y el la aplicaci�n Excel
    If Not o_Libro Is Nothing Then: o_Libro.Close False
    If Not o_Excel Is Nothing Then: o_Excel.Quit
    Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
    If Err.Number <> 1004 Then MsgBox Err.Description, vbCritical
End Function

Private Sub ReleaseObjects(o_Excel As Object, o_Libro As Object, o_Hoja As Object)
    If Not o_Excel Is Nothing Then Set o_Excel = Nothing
    If Not o_Libro Is Nothing Then Set o_Libro = Nothing
    If Not o_Hoja Is Nothing Then Set o_Hoja = Nothing
End Sub


Private Sub Form_Load()
    Dim Folio As Long, NumerodePersonal As Long, serie As String, Nombrey As String
    Dim Sub_total As Currency, Des_cto As Currency, Gravado As Currency, ConNomina As Currency
    Dim Ex_to As Currency, T_ingreso As Currency, T_desctos As Currency, T_isr As Currency, T_gravado As Currency
    Dim T_neto As Currency, T_exento As Currency, T_imss As Currency, Reg_Patr As String
    Dim Met_pago As String, MiFechaAlta, RiesgoImss, VMiFechaAlta, VMiFecha, VMiFecha1, VFal, ViaTico As Currency
    IniCols
    Close 10
    Open "EMP_CFDI.DNO" For Random As 10 Len = Len(EmpCFDI)
    Get 10, 1, EmpCFDI
    Open "Perscfdi.dno" For Random As 7 Len = Len(Empleado_1)
    Open "PerOtre.dno" For Random As 9 Len = Len(Otros_Rgtros)
    
    Folio = InputBox("Dame el numero de Folio Anterior", "CFDI NOMINA", EmpCFDI.Folio)
    serie = InputBox("Dame la serie", "CFDI NOMINA", EmpCFDI.serie)
    ConNomina = InputBox("Numero de nomina consecutivo ", "CFDI NOMINA", (EmpCFDI.Consecutivo + 1))
    If Trim(Form8.ConNom1.TextMatrix(1, 24)) <> "" Then
             Met_pago = InputBox("Metodo de Pago ", "CFDI NOMINA", "DEBITO")
             Else
                Met_pago = InputBox("Metodo de Pago ", "CFDI NOMINA", "CHEQUE")
    End If
    
    Reg_Patr = InputBox("Registro Patronal Imss ", "CFDI NOMINA", EmpCFDI.RegPatr)
    RiesgoImss = InputBox("Riesgo Patronal", "CFDI NOMINA", EmpCFDI.RiesgoImss)
    
    EmpCFDI.Folio = Folio: EmpCFDI.Consecutivo = ConNomina: EmpCFDI.serie = serie:  EmpCFDI.RegPatr = Reg_Patr: EmpCFDI.RiesgoImss = RiesgoImss
    Put 10, 1, EmpCFDI
    
    For I7 = 1 To Form8.ConNom1.Rows - 2
         If IsNumeric(Form8.ConNom1.TextMatrix(I7, 0)) Then
             NumerodePersonal = Form8.ConNom1.TextMatrix(I7, 0)
             If IsNumeric(Form8.ConNom1.TextMatrix(I7, 20)) Then T_neto = Form8.ConNom1.TextMatrix(I7, 20) Else T_neto = 0
             If IsNumeric(Form8.ConNom1.TextMatrix(I7, 11)) Then T_ingreso = Form8.ConNom1.TextMatrix(I7, 11) Else T_ingreso = 0
             If IsNumeric(Form8.ConNom1.TextMatrix(I7, 10)) Then T_exento = Form8.ConNom1.TextMatrix(I7, 10) Else T_exento = 0
             If IsNumeric(Form8.ConNom1.TextMatrix(I7, 12)) Then T_isr = Form8.ConNom1.TextMatrix(I7, 12) Else T_isr = 0
             If T_isr < 0 Then T_isr = 0
             If IsNumeric(Form8.ConNom1.TextMatrix(I7, 14)) Then T_imss = Form8.ConNom1.TextMatrix(I7, 14) Else T_imss = 0
             If Trim(Form8.ConNom1.TextMatrix(I7, 24)) <> "" Then
             Met_pago = "DEBITO"
             Else
                Met_pago = "CHEQUE"
             End If

             NomCfdi.TextMatrix(I7, 0) = 1:                             '1 SUCURSAL
             Folio = Folio + 1: NomCfdi.TextMatrix(I7, 1) = Folio       '2 FOLIO
             NomCfdi.TextMatrix(I7, 2) = serie:                         '3 SERIE
             Get 2, NumerodePersonal, personal:
             MiFechaAlta = Trim(personal.fal)
             Nombrey = Trim(personal.nom) + " " + Trim(personal.ape1) + " " + Trim(personal.ape2)
             NomCfdi.TextMatrix(I7, 3) = Nombrey:                       '4 NOMBRE
             Get 7, NumerodePersonal, Empleado_1
             NomCfdi.TextMatrix(I7, 4) = Trim(Empleado_1.Direccion):     '5 DIRECCION
             NomCfdi.TextMatrix(I7, 5) = Trim(Empleado_1.Colonia):       '6 COLONIA
             NomCfdi.TextMatrix(I7, 6) = Trim(Empleado_1.Ciudad):        '7 CIUDAD
             NomCfdi.TextMatrix(I7, 7) = Trim(Empleado_1.Estado):        '8 ESTADO
             NomCfdi.TextMatrix(I7, 8) = Trim(Empleado_1.Delegacion):    '9 DELEGACION
             NomCfdi.TextMatrix(I7, 9) = Trim(Empleado_1.Cpostal):       '10 CP
             NomCfdi.TextMatrix(I7, 10) = Trim(personal.rfc):            '11 RFC
             NomCfdi.TextMatrix(I7, 11) = "MEXICO":                      '12 PAIS
             NomCfdi.TextMatrix(I7, 12) = Trim(Empleado_1.correo):       '13 CORREO
             NomCfdi.TextMatrix(I7, 13) = ""                             '14 OBSERVACIONES
             NomCfdi.TextMatrix(I7, 14) = "PESOS"                        '15 MONEDA
             NomCfdi.TextMatrix(I7, 15) = 1                              '16 TIPOCAMBIO
             NomCfdi.TextMatrix(I7, 16) = Format(Form8.ConNom1.TextMatrix(I7, 20), "###0.00"): '17 TOTAL
             Sub_total = T_ingreso
             NomCfdi.TextMatrix(I7, 17) = Format(Sub_total, "###0.00"):                      '18 SUBTOTAL
             Des_cto = T_ingreso - Sub_total
             NomCfdi.TextMatrix(I7, 18) = Format(Des_cto, "###0.00"):                        '19 DESCUENTO
             NomCfdi.TextMatrix(I7, 19) = "Deducciones Nomina":            '20 MOTIVODESCUENTO
             Gravado = T_ingreso - T_exento:
             NomCfdi.TextMatrix(I7, 20) = Format(Gravado, "###0.00"):                        '21 TOTALGRAVADOPERCEPCIONES
             NomCfdi.TextMatrix(I7, 21) = Format(T_exento, "###0.00"):                       '22 TOTALEXENTOPERCEPCIONES
             NomCfdi.TextMatrix(I7, 22) = Format(0, "###0.00"):                              '23 TOTALGRAVADODEDUCCIONES
             NomCfdi.TextMatrix(I7, 23) = Format(T_isr + T_imss, "###0.00"):                 '24 TOTALEXENTODEDUCCIONES
             NomCfdi.TextMatrix(I7, 24) = "ISR":                           '25 CONCEPTORETISR
             NomCfdi.TextMatrix(I7, 25) = Format(T_isr, "###0.00"):                          '26 IMPORTERETISR
             NomCfdi.TextMatrix(I7, 26) = ConNomina:                       '27 PARTIDA
             NomCfdi.TextMatrix(I7, 27) = "PAGO DE NOMINA, PRIMA VACACIONAL, AGUINALDO":  '28 DESCRIPCION
             NomCfdi.TextMatrix(I7, 28) = 1:                               '29 CANTIDAD
             NomCfdi.TextMatrix(I7, 29) = "SERVICIO":                      '30 UNIDAD
             NomCfdi.TextMatrix(I7, 30) = Format(T_ingreso, "###0.00"):                      '31 VALORUNITARIO
             NomCfdi.TextMatrix(I7, 31) = Format(T_ingreso, "###0.00"):                      '32 IMPORTE
             NomCfdi.TextMatrix(I7, 32) = "":                              '33 TOTALCONLETRA
             'If Met_pago = "DEBITO" Then
                'Met_pago = Left(Trim(Form8.ConNom1.TextMatrix(I7, 24)), 4)
                'Else
                    'NomCfdi.TextMatrix(I7, 33) = Met_pago:
             'End If
             NomCfdi.TextMatrix(I7, 33) = Met_pago:                        '34 METODOPAGO
             NomCfdi.TextMatrix(I7, 34) = "MEXICO, D.F.":                  '35 LUGAREXPEDICION
             NomCfdi.TextMatrix(I7, 35) = "SUELDOS Y SALARIOS":            '36 REGIMEN
             If Met_pago = "DEBITO" Then
                    NomCfdi.TextMatrix(I7, 36) = Right(Trim(Form8.ConNom1.TextMatrix(I7, 24)), 4)  '37 NUMCTAPAG
                   Else
                    NomCfdi.TextMatrix(I7, 36) = ""                        '37 NUMCTAPAG
             End If
             NomCfdi.TextMatrix(I7, 37) = Reg_Patr:                        '38 REGISTROPATRONAL
             NomCfdi.TextMatrix(I7, 38) = NumerodePersonal:                '39 NUMEMPLEADO
             Get 9, NumerodePersonal, Otros_Rgtros
             NomCfdi.TextMatrix(I7, 39) = Otros_Rgtros.curp:               '40 CURP
             NomCfdi.TextMatrix(I7, 40) = 2:                               '41 TIPOREGIMEN
             NomCfdi.TextMatrix(I7, 41) = Trim(personal.imss):             '42 NUMSEGURIDADSOCIAL
             VMiFecha = Mid(Trim(MiFecha), 7, 4) + "-" + Mid(Trim(MiFecha), 4, 2) + "-" + Mid(Trim(MiFecha), 1, 2)
             NomCfdi.TextMatrix(I7, 42) = VMiFecha:                         '43 FECHAPAGO
             DDINI = Left(MiFecha, 2)
             If DDINI < 16 Then
                  DDINI = 1
                  MiFecha1 = "01" + Mid(MiFecha, 3)
                  Else
                  DDINI = 16
                  MiFecha1 = "16" + Mid(MiFecha, 3)
             End If
             VMiFecha1 = Mid(Trim(MiFecha1), 7, 4) + "-" + Mid(Trim(MiFecha1), 4, 2) + "-" + Mid(Trim(MiFecha1), 1, 2)
             NomCfdi.TextMatrix(I7, 43) = VMiFecha1:                         '44 FECHAINICIALPAGO
             NomCfdi.TextMatrix(I7, 44) = VMiFecha:                          '45 FECHAFINALPAGO
             If N_ormal = 1 Then
                    NomCfdi.TextMatrix(I7, 45) = 0.01: Rem Form8.ConNom1.TextMatrix(I7, 2):  '46 NUMDIASPAGADOS
                    Else
                    NomCfdi.TextMatrix(I7, 45) = Form8.ConNom1.TextMatrix(I7, 2):  '46 NUMDIASPAGADOS
             End If
             NomCfdi.TextMatrix(I7, 46) = "OFICINA"                         '47 DEPARTAMENTO
             NomCfdi.TextMatrix(I7, 47) = ""                       '48 CLABE
             NomCfdi.TextMatrix(I7, 48) = ""                       '49 BANCO
               VFal = Mid(Trim(personal.fal), 7, 4) + "-" + Mid(Trim(personal.fal), 4, 2) + "-" + Mid(Trim(personal.fal), 1, 2)
             NomCfdi.TextMatrix(I7, 49) = VFal                      '50 FECHAINICIORELLABORAL
             NomCfdi.TextMatrix(I7, 50) = Year(MiFecha) - Year(MiFechaAlta) '51 ANTIGUEDAD
             NomCfdi.TextMatrix(I7, 51) = ""                       '52 PUESTO
             NomCfdi.TextMatrix(I7, 52) = ""                       '53 TIPOCONTRATO
             NomCfdi.TextMatrix(I7, 53) = ""                       '54 TIPOJORNADA
             If N_ormal = 1 Then
                    NomCfdi.TextMatrix(I7, 54) = "ANUAL"
                    Else
                    NomCfdi.TextMatrix(I7, 54) = "QUINCENAL"                       '55 PERIODICIDADPAGO
             End If
             NomCfdi.TextMatrix(I7, 55) = Format(personal.ingr, "###0.00")  '56 SALARIOBASECOTAPOR
             NomCfdi.TextMatrix(I7, 56) = RiesgoImss                        '57 RIESGOPUESTO
             NomCfdi.TextMatrix(I7, 57) = Format(personal.integrado, "###0.00")               '58 SALARIODIARIOINTEGRADO
            
         End If
    Next I7
    Get 10, 1, EmpCFDI: EmpCFDI.Folio = Folio: Put 10, 1, EmpCFDI
    Close 7, 9, 10
End Sub

Private Sub NCfEdCop_Click()
      Dim Temporal1
 Clipboard.Clear
   
   difer = NomCfdi.RowSel - NomCfdi.Row
   For i = NomCfdi.Row To NomCfdi.RowSel
      
      For f = NomCfdi.Col To NomCfdi.ColSel
            Temporal1 = Temporal1 + NomCfdi.TextMatrix(i, f)
            If f < NomCfdi.ColSel Then
                Temporal1 = Temporal1 & Chr(9)
            End If
      Next f
      Temporal1 = Temporal1 & Chr(13) & Chr(10)
      
   Next i
    Clipboard.SetText Temporal1
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> 1 Then
   NomCfdi.Height = ScaleHeight - 1000
   NomCfdi.Width = ScaleWidth - 600
 End If

End Sub

Private Sub LTXT_Change()
    NomCfdi.Text = LTXT.Text
End Sub

Private Sub NCfEdSel_Click()
     Dim limite As Long
    Clipboard.Clear
    NomCfdi.Row = 1: NomCfdi.Col = 0
   For limite = 1 To NomCfdi.Rows - 1
       renglon = limite
    If IsNumeric(NomCfdi.TextMatrix(renglon, 0)) Then
           NomCfdi.RowSel = renglon
    End If
   Next limite
    NomCfdi.ColSel = NomCfdi.Cols - 1
End Sub
Private Sub NomCfdi_EnterCell()
  If NomCfdi.Row > 0 Then
    NomCfdi.CellBackColor = vbYellow
  End If
    valcelant = NomCfdi.Text
    LTXT.Text = valcelant
    Rem If NomCfdi.Row > 25 Then NomCfdi.TopRow = NomCfdi.TopRow + 1
    
End Sub

Private Sub NomCfdi_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
            Case vbKeyDelete
               
               For Q = NomCfdi.Row To NomCfdi.RowSel
                 
                 For W = NomCfdi.Col To NomCfdi.ColSel
                    
                    NomCfdi.TextMatrix(Q, W) = ""
                 Next W
               Next Q
               
                LTXT.Text = NomCfdi.Text
            Case vbKeyF2
                If NomCfdi.Text <> "" Then valcelant = NomCfdi.Text
                LTXT.Text = LTrim(RTrim(NomCfdi.Text))
                LTXT.SetFocus
               
       End Select

End Sub

Private Sub NomCfdi_KeyPress(KeyAscii As Integer)
    valcelant = NomCfdi.Text
    LTXT.Text = Chr(KeyAscii)
    LTXT.SetFocus
End Sub

Private Sub NomCfdi_LeaveCell()
  If NomCfdi.Row > 0 Then
   NomCfdi.CellBackColor = vbWhite
  End If
End Sub


