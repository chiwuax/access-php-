Dim i, j, ultimafila, tc, bandera, bandera2, BT, BTe, BMc, BMo, contador  As Integer
Dim c As Range
Dim firstAddress, abc, TecnicosCod, MyValue As String
Dim Message, Title, default

Sub Test1()

     ultimafila = Cells(Rows.Count, 1).End(xlUp).Row ' importante recorre hoja y regresa la ultima fila con datos
     i = ultimafila
     
End Sub
Sub buscar()

    'On Error Resume Next
    
    If abc = "" Then
    Else
    
    Selection.Find(what:=abc, after:=ActiveCell, LookIn:=xlFormulas, MatchCase:=False).Select ' Busco en la columna seleccionada el valor de cbm y seleciono la celda donde esta
    i = ActiveCell.Row 'doy el valor a line de la fila donde ha encontrado anteriormente el valor de cbm
   ' MsgBox i
    
    End If
    
End Sub
Private Sub cb_tecnicos_Change()

    abc = CB_tecnicos.Text
    
    If abc = "" Then
        BT = BT + 1
    Else
    i = 0
    ThisWorkbook.Sheets("Tecnicos").Activate
    ThisWorkbook.Sheets("Tecnicos").Application.Range("B:B").Select ' Selecciona la columna donde buscar.
    Call buscar
    TecnicosCod = ActiveSheet.Cells(i, 1)
    BT = 0
'    MsgBox " EL codigo de programador es:  " & TecnicosCod & " y su nombre es  " & ActiveSheet.Cells(i, 2)
    End If

End Sub
Private Sub CB_TE_Change()

    abc = CB_TE.Text
    
    If abc = "" Then
        BTe = BTe + 1
    Else
    i = 0
    ThisWorkbook.Sheets("Tipos").Activate
    ThisWorkbook.Sheets("Tipos").Application.Range("B:B").Select ' Selecciona la columna donde buscar.
    
    Call buscar
    TecnicosCod = ActiveSheet.Cells(i, 1)
    BTe = 0
'    MsgBox " EL codigo de programador es:  " & TecnicosCod & " y su nombre es  " & ActiveSheet.Cells(i, 2)
     End If

End Sub
Private Sub cb_marca_Change()

    abc = CB_marca.Text
    
    If abc = "" Then
        BMc = BMc + 1
    Else
    i = 0
    ThisWorkbook.Sheets("marcas").Activate
    Sheets("Marcas").Application.Range("B:B").Select ' Selecciona la columna donde buscar.
    Call buscar
    TecnicosCod = ActiveSheet.Cells(i, 1)
    BMc = 0
'    MsgBox " EL codigo de programador es:  " & TecnicosCod & " y su nombre es  " & ActiveSheet.Cells(i, 2)
     End If

End Sub
Private Sub CB_modelo_Change()

    abc = CB_modelo.Text
    If abc = "" Then
        BMo = BMo + 1
    Else
    i = 0
    ThisWorkbook.Sheets("Modelos").Activate
    ThisWorkbook.Sheets("Modelos").Application.Range("B:B").Select ' Selecciona la columna donde buscar.
    Call buscar
    TecnicosCod = ActiveSheet.Cells(i, 1)
    BMo = 0
'    MsgBox " EL codigo de programador es:  " & TecnicosCod & " y su nombre es  " & ActiveSheet.Cells(i, 2)
    End If
     
End Sub
Sub llena()

ThisWorkbook.Sheets("Tecnicos").Activate
Call Test1
 For tc = 2 To i
    CB_tecnicos.AddItem (Cells(tc, 2))
    'MsgBox tc
 Next

ThisWorkbook.Sheets("Tipos").Activate
Call Test1
 For tc = 2 To i
    CB_TE.AddItem (Cells(tc, 2))
    'MsgBox tc
 Next

ThisWorkbook.Sheets("marcas").Activate
Call Test1
 For tc = 2 To i
    CB_marca.AddItem (Cells(tc, 2))
    'MsgBox tc
 Next
 
ThisWorkbook.Sheets("Modelos").Activate
Call Test1
 For tc = 2 To i
    CB_modelo.AddItem (Cells(tc, 2))
    'MsgBox tc
 Next

End Sub
Private Sub cmd_addC_Click()
    
    Unload Me
    clientes.Show
        
End Sub
Sub pregunta()
 
'    Message = "Enter a value between 1 and 3"    ' Set prompt.
    Title = "Introducir"
    default = " "
    
    
    ' Display message, title, and default value.
    MyValue = InputBox(Message, Title, default)
    

'MsgBox MyValue
End Sub
Private Sub cmd_BC_Click()

    Unload Me
    Buscar_cliente.Show
    
End Sub

Private Sub cmd_marca_Click()

    Message = "Marca"
    Call pregunta
    
    If StrPtr(MyValue) = 0 Then
        'MsgBox ("User canceled!")
    Else
        
    ThisWorkbook.Sheets("marcas").Activate
    ThisWorkbook.Sheets("marcas").Application.Range("B:B").Select ' Selecciona la columna donde buscar.
    
    Call Test1
      
    contador = ActiveSheet.Cells(i, 1) + 1
    'MsgBox contador
     i = i + 1
    ActiveSheet.Cells(i, 1) = contador
    ActiveSheet.Cells(i, 2) = MyValue
    
    CB_marca.Text = MyValue
'    CB_marca.Locked = True
    

    End If
    
End Sub

Private Sub cmd_mod_Click()

    Message = "Tipo de Equipo"
    Call pregunta
    
    If StrPtr(MyValue) = 0 Then
        'MsgBox ("User canceled!")
    Else
        
    ThisWorkbook.Sheets("modelos").Activate
    ThisWorkbook.Sheets("modelos").Application.Range("B:B").Select ' Selecciona la columna donde buscar.
    
    Call Test1
      
    contador = ActiveSheet.Cells(i, 1) + 1
    'MsgBox contador
     i = i + 1
    ActiveSheet.Cells(i, 1) = contador
    ActiveSheet.Cells(i, 2) = MyValue
    
    CB_modelo.Text = MyValue

    End If
End Sub

Private Sub cmd_TE_Click()

    Message = "Tipo de Equipo"
    Call pregunta
    
    If StrPtr(MyValue) = 0 Then
        'MsgBox ("User canceled!")
    Else
        
    ThisWorkbook.Sheets("Tipos").Activate
    ThisWorkbook.Sheets("Tipos").Application.Range("B:B").Select ' Selecciona la columna donde buscar.
    
    Call Test1
      
    contador = ActiveSheet.Cells(i, 1) + 1
    'MsgBox contador
     i = i + 1
    ActiveSheet.Cells(i, 1) = contador
    ActiveSheet.Cells(i, 2) = MyValue
    
    CB_TE.Text = MyValue

    End If
    
End Sub
Private Sub UserForm_Initialize()
     
    Call llena
    Call busqueda
    
    Txt_fecha.Value = Date
    Txt_fecha.Locked = True
    
   
  
End Sub
Sub busqueda()
If bandera_busqueda = 0 Then
Txt_cliente.Text = nombre
    Txt_dir.Text = dir
    Txt_tel.Text = tel
Else

    abc = nombre_cliente
    ThisWorkbook.Sheets("clientes").Activate
    ThisWorkbook.Sheets("clientes").Application.Range("B:B").Select
    Call buscar
    Txt_cliente.Text = ActiveSheet.Cells(i, 2)
    Txt_dir.Text = ActiveSheet.Cells(i, 3)
    Txt_tel.Text = ActiveSheet.Cells(i, 4)
    bandera_busqueda = 0

End If

End Sub
