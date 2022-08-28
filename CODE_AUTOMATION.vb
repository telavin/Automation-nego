Sub egan()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim archivos
Dim goku As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String
archivos = Dir("D:\AUTOMA_FULL_NEGOCIOS\HC\*.xlsx")
Do While archivos <> ""
Workbooks.Open "D:\AUTOMA_FULL_NEGOCIOS\HC\" & archivos
archivos = Dir
Loop
bills = ActiveWorkbook.Name
tt = bills
Set goku = Workbooks.Open("D:\AUTOMA_FULL_NEGOCIOS\PLANTILLAS\PLANTILLA_HC")
xiaomi = "PLANTILLA_HC"
ss = xiaomi & ".xlsx"
Windows(tt).Activate
RATA = Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
Sheets(1).Range("A1:U" & RATA).Copy
Windows(ss).Activate
Sheets("HC").Select
Range("A2").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Windows(tt).Activate
ActiveWorkbook.Close
MsgBox "proceso de pegue completado", vbInformation
Windows(ss).Activate
Sheets("HC").Select
Set rangodatos = Sheets("HC").UsedRange
rangodatos.AutoFilter Field:=19, Criteria1:="<>"
Range("E1").Select
ActiveCell.FormulaR1C1 = "=RC[14]"
Selection.Copy
Range(Selection, Selection.End(xlDown)).Select
ActiveSheet.Paste
Application.CutCopyMode = False
Range("E1") = "Cod Cargo"
Range("G1").Select
ActiveCell.FormulaR1C1 = "=RC[13]"
Selection.Copy
Range(Selection, Selection.End(xlDown)).Select
ActiveSheet.Paste
Application.CutCopyMode = False
Range("G1") = "Sueldo Variable"
ActiveSheet.ShowAllData
ActiveWorkbook.Save
ActiveWorkbook.Close
Application.ScreenUpdating = True
Application.DisplayAlerts = True
MsgBox "proceso FINAL FINAL completado CAPO, puede continuar con el siguiente paso", vbInformation
End Sub
Sub garan()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
archivos = Dir("D:\AUTOMA_FULL_NEGOCIOS\GARANTIZADOS\*.xlsx")
Do While archivos <> ""
Workbooks.Open "D:\AUTOMA_FULL_NEGOCIOS\GARANTIZADOS\" & archivos
archivos = Dir
Loop
goten = ActiveWorkbook.Name
pp = goten
Windows(pp).Activate
Sheets(1).Select
Range("A1").Select
las = Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
Sheets(1).Range("A2:E" & las).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Garantizados").Select
Range("A2").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
:=False, Transpose:=False
Application.CutCopyMode = False
Windows(pp).Activate
ActiveWorkbook.Close SaveChanges:=False
Range("H1").Select
ActiveCell.FormulaR1C1 = "=COUNTA(C[-7])-1"
claro = Range("H1").Value
Range("F2").Select
For i = 1 To claro
ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-3],HC!C1:C7,5,FALSE)"
ActiveCell.Offset(1, 0).Select
Next i
MsgBox "proceso FINAL FINAL completado CAPO, puede continuar con el siguiente paso", vbInformation
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub maquillaje()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("LIQUIDADOR").Select
Range("B3").Select
fav = Selection.End(xlDown).Row - 3
Range("A4").Select
For h = 1 To fav
ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[1],HC!C:C[17],2,0)"
ActiveCell.Offset(1, 0).Select
Next h
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic
MsgBox "proceso FINAL FINAL completado CAPO, puede continuar con el siguiente paso", vbInformation
MsgBox "Por favor espere DOS minutos mientras se enfría el procesador, de lo contrario la memoria se desboradará", vbInformation
End Sub
Sub blabla()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim archivos
Dim goku As Excel.Workbook
Dim vegueta As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String
archivos = Dir("D:\AUTOMA_FULL_NEGOCIOS\AUSENTISMOS\*.xlsx")
Do While archivos <> ""
Workbooks.Open "D:\AUTOMA_FULL_NEGOCIOS\AUSENTISMOS\" & archivos
archivos = Dir
Loop
bills = ActiveWorkbook.Name
tt = bills
Windows(tt).Activate
Sheets(1).Select
Set rangodatos = Sheets(1).UsedRange
rangodatos.AutoFilter Field:=5, Criteria1:="=ret*", _
Operator:=xlAnd
Range("E1").Select
With Selection.Interior
 .Pattern = xlNone
 .TintAndShade = 0
 .PatternTintAndShade = 0
End With
ActiveCell = "retiro"
Selection.Copy
Range(Selection, Selection.End(xlDown)).Select
ActiveSheet.Paste
Application.CutCopyMode = False
Range("E1") = "CLASE AUSENTISMO"
ActiveSheet.ShowAllData
MsgBox "listos retiros", vbInformation
Range("I1") = "HC ACTUAL"
Windows(tt).Activate
Range("M1").Select
ActiveCell.FormulaR1C1 = "=COUNTA(C[-12])-1"
Set goku = Workbooks.Open("D:\AUTOMA_FULL_NEGOCIOS\PLANTILLAS\PLANTILLA_HC")
xiaomi = "PLANTILLA_HC"
ss = xiaomi & ".xlsx"
Windows(ss).Activate
Windows(tt).Activate
hola = Range("M1").Value
Range("I2").Select
For i = 1 To hola
ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-8],[PLANTILLA_HC.xlsx]HC!C1:C5,5,FALSE)"
ActiveCell.Offset(1, 0).Select
Next i
Range("J1") = "HC VIEJO"
Windows(ss).Activate
ActiveWorkbook.Save
ActiveWorkbook.Close
MsgBox "LISTOS EL ACTUAL, CAPO", vbInformation
Set vegueta = Workbooks.Open("D:\AUTOMA_FULL_NEGOCIOS\VIEJA\HC_VIEJO")
nevado = "HC_VIEJO"
cr7 = nevado & ".xlsx"
Windows(cr7).Activate
Windows(tt).Activate
Sheets(1).Select
carita = Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
ActiveSheet.ListObjects.Add(xlSrcRange, Range("A1:J" & carita), , xlYes).Name = _
"Tabla1"
Range("Tabla1[#All]").Select
Range("j2").Select
For i = 1 To hola
ActiveCell.FormulaR1C1 = "=VLOOKUP([@CEDULA],[HC_VIEJO.xlsx]HC!C3:C7,5,FALSE)"
ActiveCell.Offset(1, 0).Select
Next i
MsgBox "listo viejo", vbInformation
Range("Tabla1").AutoFilter Field:=9, Criteria1:="#N/A"
Range("Tabla1").AutoFilter Field:=10, Criteria1:="<>#N/A" _
, Operator:=xlAnd
Range("A1").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Windows(cr7).Activate
Worksheets.Add.Name = "NUEVA"
Sheets("NUEVA").Select
Range("A1").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Windows(tt).Activate
ActiveWorkbook.Save
ActiveWorkbook.Close
Windows(cr7).Activate
Sheets("HC").Select
Range("AI1").Select
ActiveCell.FormulaR1C1 = "=COUNTA(C[-32])-1"
CISCO = Range("AI1").Value
Range("AB2").Select
For i = 1 To CISCO
ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-25],NUEVA!C[-27],1,FALSE)"
ActiveCell.Offset(1, 0).Select
Next i
Sheets("HC").Select
Set rangodatos = Sheets("HC").UsedRange
rangodatos.AutoFilter Field:=28, Criteria1:="<>#N/A" _
, Operator:=xlAnd
rangodatos.AutoFilter Field:=7, Criteria1:=Array("Consultor Negocios", "Especialista Comercial Negocios", "Especialista Canales Negocios", "Especialista Ventas Sinergia Empresas", "Especialista Comercial Telemercadeo Empresas Y Negocios", "Especialista Telemercadeo Empresas Y Negocios", "Coordinador Comercial Negocios", "Coordinador Ventas Negocios", "Jefe Comercial Negocios", "Jefe Telemercadeo Pyme", "Jefe Canales Negocios", "Gerente Comercial Telemercadeo Empresas Y Negocios", "Gerente Comercial Negocios 1", "Gerente Comercial Negocios 2", "Gerente Comercial Negocios 3", "Director Negocios"), _
Operator:=xlFilterValues
Range("AF1").Select
ActiveCell.FormulaR1C1 = "=SUBTOTAL(103,C3)-1"
lenovo = Range("AF1").Value
If lenovo >= 1 Then
Range("C1").Select
Range(Selection, Selection.End(xlDown)).Select
With Selection.Font
 .Color = -16776961
 .TintAndShade = 0
End With
dedo = Sheets("HC").Range("C" & Rows.Count).End(xlUp).Row
Sheets("HC").Select
Sheets("HC").Range("C1:W" & dedo).Copy
Set goku = Workbooks.Open("D:\AUTOMA_FULL_NEGOCIOS\PLANTILLAS\PLANTILLA_HC")
xiaomi = "PLANTILLA_HC"
ss = xiaomi & ".xlsx"
Windows(ss).Activate
enero = Sheets("HC").Range("A" & Rows.Count).End(xlUp).Row
Sheets("HC").Select
Range("A" & enero + 1).Select
ActiveSheet.Paste
Application.CutCopyMode = False
Sheets("HC").Select
Set rangodatos = Sheets("HC").UsedRange
rangodatos.AutoFilter Field:=1, Criteria1:="Expediente"
With Range("A1")
Range(Cells(Rows.Count, .Column).End(xlUp), Cells(.Row + 1, Columns.Count).End(xlToLeft)).SpecialCells(12).Delete
End With
Range("A1").Select
ActiveSheet.ShowAllData
Windows(ss).Activate
ActiveWorkbook.Save
ActiveWorkbook.Close
Windows(cr7).Activate
ActiveWorkbook.Save
ActiveWorkbook.Close
Application.ScreenUpdating = True
Application.DisplayAlerts = True
MsgBox "proceso FINAL FINAL completado CAPO, puede continuar con el siguiente paso", vbInformation
Else
Windows(cr7).Activate
ActiveWorkbook.Save
ActiveWorkbook.Close
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic
MsgBox "proceso FINAL FINAL completado CAPO, puede continuar con el siguiente paso", vbInformation
MsgBox "Por favor espere DOS minutos mientras se enfría el procesador, de lo contrario la memoria se desboradará", vbInformation
End If
End Sub
Sub functlon()
Dim elon As Integer
Dim crew As String
Dim musk As String
crew = "D:\AUTOMA_FULL_NEGOCIOS\ARCHIVO PLANO\CONSULTA_PERMANENCIA.txt"
elon = FreeFile
Open crew For Input As elon
musk = Input(LOF(elon), elon)
Close elon
musk = Replace(musk, "a", "'")
elon = FreeFile
Open crew For Output As elon
Print #elon, musk
Close elon
End Sub
Sub CARGUE_HC()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim goku As Excel.Workbook
Dim vegueta As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String
Set goku = Workbooks.Open("D:\AUTOMA_FULL_NEGOCIOS\PLANTILLAS\PLANTILLA_HC")
xiaomi = "PLANTILLA_HC"
ss = xiaomi & ".xlsx"
Windows(ss).Activate
col = Sheets("HC").Range("A" & Rows.Count).End(xlUp).Row
Sheets("HC").Range("A2:U" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("HC").Select
Range("A2").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Range("A2").Select
Windows(ss).Activate
ActiveWorkbook.Close SaveChanges:=False
Application.ScreenUpdating = True
Application.DisplayAlerts = True
MsgBox "proceso FINAL FINAL completado CAPO, puede continuar con el siguiente paso", vbInformation
End Sub
Sub vacaciones()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim archivos
Dim goku As Excel.Workbook
Dim vegueta As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String
archivos = Dir("D:\AUTOMA_FULL_NEGOCIOS\VACACIONES\*.csv")
Do While archivos <> ""
Workbooks.Open "D:\AUTOMA_FULL_NEGOCIOS\VACACIONES\" & archivos
archivos = Dir
Loop
bills = ActiveWorkbook.Name
tt = bills
Windows(tt).Activate
Columns("A:A").Select
Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
Semicolon:=True, Comma:=False, Space:=False, Other:=False, FieldInfo _
:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), _
Array(7, 1)), TrailingMinusNumbers:=True
Range("A1").Select
Sheets(1).Select
Range("A1").Select
col = Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
Sheets(1).Range("A1:G" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Vacaciones").Select
Range("A2").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
:=False, Transpose:=False
Application.CutCopyMode = False
Range("A2").Select
Windows(tt).Activate
ActiveWorkbook.Close SaveChanges:=False
MsgBox "proceso FINAL FINAL completado CAPO, puede continuar con el siguiente paso", vbInformation
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub auto_open()
UserForm8.Show
End Sub
Sub spacex()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim archivos
Dim goku As Excel.Workbook
Dim vegueta As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String
archivos = Dir("D:\AUTOMA_FULL_NEGOCIOS\METAS\*.xlsx")
Do While archivos <> ""
Workbooks.Open "D:\AUTOMA_FULL_NEGOCIOS\METAS\" & archivos
archivos = Dir
Loop
bills = ActiveWorkbook.Name
tt = bills
Windows(tt).Activate
Sheets("LIDERES NEGOCIOS-INTERMEDIAS").Select
Range("A1").Select
col = Sheets("LIDERES NEGOCIOS-INTERMEDIAS").Range("A" & Rows.Count).End(xlUp).Row
Sheets("LIDERES NEGOCIOS-INTERMEDIAS").Range("B2:D" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Meta General").Select
Range("A2").Select
ActiveSheet.Paste
Windows(tt).Activate
Sheets("LIDERES NEGOCIOS-INTERMEDIAS").Range("F2:M" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Range("D2").Select
ActiveSheet.Paste
Windows(tt).Activate
Sheets("LIDERES NEGOCIOS-INTERMEDIAS").Range("N2:N" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Range("M2").PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("META GENERAL NEGOCIOS").Select
Range("A1").Select
musk = Sheets("META GENERAL NEGOCIOS").Range("A" & Rows.Count).End(xlUp).Row
Sheets("META GENERAL NEGOCIOS").Range("A2:C" & musk).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Range("A1").Select
dragon = Selection.End(xlDown).Row
Range("A" & dragon + 1).Select
ActiveSheet.Paste
Windows(tt).Activate
Sheets("META GENERAL NEGOCIOS").Range("E2:L" & musk).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Range("D1").Select
crew = Selection.End(xlDown).Row
Range("D" & crew + 1).Select
ActiveSheet.Paste
Windows(tt).Activate
Sheets("META GENERAL NEGOCIOS").Range("M2:M" & musk).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Range("M1").Select
crew = Selection.End(xlDown).Row
Range("M" & crew + 1).PasteSpecial xlPasteValues
Windows(tt).Activate
ActiveWorkbook.Close SaveChanges:=False
MsgBox "listo el pegue", vbInformation
Range("S1").Select
ActiveCell.FormulaR1C1 = "=COUNTA(C1)-1"
elon = Range("S1").Value
Range("L2").Select
For i = 1 To elon
ActiveCell.FormulaR1C1 = "=SUM(RC[-5],RC[-1])"
ActiveCell.Offset(1, 0).Select
Next i
MsgBox "proceso FINAL FINAL completado CAPO, puede continuar con el siguiente paso", vbInformation
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub nasa()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim archivos
Dim goku As Excel.Workbook
Dim vegueta As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String
archivos = Dir("D:\AUTOMA_FULL_NEGOCIOS\FIJOS\*.xlsx")
Do While archivos <> ""
Workbooks.Open "D:\AUTOMA_FULL_NEGOCIOS\FIJOS\" & archivos
archivos = Dir
Loop
bills = ActiveWorkbook.Name
tt = bills
Windows(tt).Activate
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Meses").Select
yoyis = Range("H1").Value
Windows(tt).Activate
Sheets("Base_Comisiones").Select
Set rangodatos = Sheets("Base_Comisiones").UsedRange
rangodatos.AutoFilter Field:=4, Criteria1:=yoyis
Sheets("Base_Comisiones").Select
Range("A1").Select
col = Sheets("Base_Comisiones").Range("A" & Rows.Count).End(xlUp).Row
Sheets("Base_Comisiones").Range("A2:B" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("A2").PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("O2:P" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("D2").PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("CB2:CC" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("F2").PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("K2:M" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("H2").PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("W2:X" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("K2").PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("J2:J" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("M2").PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("AC2:AC" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("O2").PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("AD2:AD" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("Q2").PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("V2:V" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("R2").PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("S2:S" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("S2").PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("AR2:AR" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("T2").PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("H2:I" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("V2").PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("G2:G" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("X2").PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("AW2:AW" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("AB2").PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("CF2:CI" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("AC2").PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("AU2:AU" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("AG2").PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("AS2:AT" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("AH2").PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("U2:U" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("AJ2").PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("BZ2:CA" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("AK2").PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("CZ2:CZ" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("AM2").PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("DD2:DD" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("AN2").PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("BJ2:BJ" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("AO2").PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("CD2:CE" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("AP2").PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("CR2:CS" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("AS2").PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("CW2:CW" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("AU2").PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("CP2:CQ" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("AV2").PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("CT2:CU" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("AX2").PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("DA2:DA" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("AZ2").PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("D2:D" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("C2").PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("Z2:Z" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("P2").PasteSpecial xlPasteValues
Windows(tt).Activate
ActiveWorkbook.Close SaveChanges:=False
MsgBox "LISTO EL PEGUE 1", vbInformation
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub golovin()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim archivos
Dim goku As Excel.Workbook
Dim vegueta As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String
Set goku = Workbooks.Open("D:\AUTOMA_FULL_NEGOCIOS\PLANTILLAS\LEGALI")
xiaomi = "LEGALI"
ss = xiaomi & ".xlsx"
Sheets(1).Select
Range("A1").Select
coll = Selection.End(xlDown).Row
Sheets(1).Range("A1:H" & coll).Select
Selection.ClearContents
Range("A1").Select
Windows(ss).Activate
ActiveWorkbook.Close SaveChanges:=True
Set goku = Workbooks.Open("D:\AUTOMA_FULL_NEGOCIOS\PLANTILLAS\PERMA")
xiaom = "PERMA"
st = xiaom & ".xlsx"
Sheets(1).Select
Range("A1").Select
cooll = Selection.End(xlDown).Row
Sheets(1).Range("A1:J" & cooll).Select
Selection.ClearContents
Range("A1").Select
Windows(st).Activate
ActiveWorkbook.Close SaveChanges:=True
Set goku = Workbooks.Open("D:\AUTOMA_FULL_NEGOCIOS\PLANTILLAS\PLANTILLA_HC")
xiom = "PLANTILLA_HC"
tt = xiom & ".xlsx"
Sheets(1).Select
Range("A1").Select
coooll = Selection.End(xlDown).Row
Sheets(1).Range("A2:U" & coooll).Select
Selection.ClearContents
Range("A1").Select
Windows(tt).Activate
ActiveWorkbook.Close SaveChanges:=True
Set goku = Workbooks.Open("D:\AUTOMA_FULL_NEGOCIOS\MOVIL VIEJO\BACKLOG")
xiomm = "BACKLOG"
pp = xiomm & ".xlsx"
Sheets(1).Select
Range("A1").Select
tuy = Selection.End(xlDown).Row
Sheets(1).Range("A2:AO" & tuy).Select
Selection.ClearContents
Range("A1").Select
Windows(pp).Activate
ActiveWorkbook.Close SaveChanges:=True
Set goku = Workbooks.Open("D:\AUTOMA_FULL_NEGOCIOS\VIEJA\HC_VIEJO")
eg = "HC_VIEJO"
xx = eg & ".xlsx"
Sheets("HC").Select
Range("A1").Select
ActiveSheet.ShowAllData
ty = Selection.End(xlDown).Row
Sheets("HC").Range("A2:AB" & ty).Select
Selection.ClearContents
Range("AF1").ClearContents
Range("AI1").ClearContents
Range("A1").Select
Sheets("NUEVA").Delete
Windows(xx).Activate
ActiveWorkbook.Close SaveChanges:=True
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub franco()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim archivos
Dim goku As Excel.Workbook
Dim vegueta As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String
archivos = Dir("D:\AUTOMA_FULL_NEGOCIOS\NETO MOVIL\*.xlsx")
Do While archivos <> ""
Workbooks.Open "D:\AUTOMA_FULL_NEGOCIOS\NETO MOVIL\" & archivos
archivos = Dir
Loop
bills = ActiveWorkbook.Name
tt = bills
Windows(tt).Activate
Range("A3").Select
cool = Selection.End(xlDown).Row
Sheets(1).Range("A3:G" & cool).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Neto Movil").Select
Range("A3").PasteSpecial xlPasteValues
Range("A3").Select
MsgBox "LISTO MOVIL", vbInformation
Windows(tt).Activate
ActiveWorkbook.Close SaveChanges:=False
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub funtion()
Dim elon As Integer
Dim crew As String
Dim musk As String
crew = "D:\AUTOMA_FULL_NEGOCIOS\ARCHIVO PLANO\CONSULTA_PERMANENCIA.txt"
elon = FreeFile
Open crew For Input As elon
musk = Input(LOF(elon), elon)
Close elon
musk = Replace(musk, "b", "',")
elon = FreeFile
Open crew For Output As elon
Print #elon, musk
Close elon
End Sub
Sub falcon()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim archivos
Dim goku As Excel.Workbook
Dim vegueta As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String
Sheets("Meses").Select
KING = WorksheetFunction.VLookup(Range("H1").Value, Sheets("Meses").Range("A:B"), 2, 0) - 30
bongo = Format(KING, "mmmm")
archivos = Dir("D:\AUTOMA_FULL_NEGOCIOS\FIJOS\*.xlsx")
Do While archivos <> ""
Workbooks.Open "D:\AUTOMA_FULL_NEGOCIOS\FIJOS\" & archivos
archivos = Dir
Loop
bills = ActiveWorkbook.Name
tt = bills
Windows(tt).Activate
Sheets("Base_Comisiones").Select
Range("A1").Select
Set rangodatos = Sheets("Base_Comisiones").UsedRange
rangodatos.AutoFilter Field:=79, Operator:= _
xlFilterValues, Criteria2:=Array(0, "1/1/2021")
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
calamaro = WorksheetFunction.VLookup(Range("H1").Value, Sheets("Meses").Range("A:C"), 3, 0)
Windows(tt).Activate
Set rangodatos = Sheets("Base_Comisiones").UsedRange
rangodatos.AutoFilter Field:=79, Operator:= _
xlFilterValues, Criteria2:=Array(1, calamaro & "/1/2021")
Set rangodatos = Sheets("Base_Comisiones").UsedRange
rangodatos.AutoFilter Field:=4, Criteria1:=bongo
Sheets("Base_Comisiones").Select
Range("A1").Select
dragon = Selection.End(xlDown).Row
col = Sheets("Base_Comisiones").Range("A" & Rows.Count).End(xlUp).Row
Sheets("Base_Comisiones").Range("A2:B" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("A1").Select
dragon = Selection.End(xlDown).Row
Range("A" & dragon + 1).Select
Selection.PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("O2:P" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("D" & dragon + 1).Select
Selection.PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("CB2:CC" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("F" & dragon + 1).Select
Selection.PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("K2:M" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("H" & dragon + 1).Select
Selection.PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("W2:X" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("K" & dragon + 1).Select
Selection.PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("J2:J" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("M" & dragon + 1).Select
Selection.PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("AC2:AC" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("O" & dragon + 1).Select
Selection.PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("AD2:AD" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("Q" & dragon + 1).Select
Selection.PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("V2:V" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("R" & dragon + 1).Select
Selection.PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("S2:S" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("S" & dragon + 1).Select
Selection.PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("AR2:AR" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("T" & dragon + 1).Select
Selection.PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("H2:I" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("V" & dragon + 1).Select
Selection.PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("G2:G" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("X" & dragon + 1).Select
Selection.PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("AW2:AW" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("AB" & dragon + 1).Select
Selection.PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("CF2:CI" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("AC" & dragon + 1).Select
Selection.PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("AU2:AU" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("AG" & dragon + 1).Select
Selection.PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("AS2:AT" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("AH" & dragon + 1).Select
Selection.PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("U2:U" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("AJ" & dragon + 1).Select
Selection.PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("BZ2:CA" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("AK" & dragon + 1).Select
Selection.PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("CZ2:CZ" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("AM" & dragon + 1).Select
Selection.PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("DD2:DD" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("AN" & dragon + 1).Select
Selection.PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("BJ2:BJ" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("AO" & dragon + 1).Select
Selection.PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("CD2:CE" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("AP" & dragon + 1).Select
Selection.PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("CR2:CS" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("AS" & dragon + 1).Select
Selection.PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("CW2:CW" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("AU" & dragon + 1).Select
Selection.PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("CP2:CQ" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("AV" & dragon + 1).Select
Selection.PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("CT2:CU" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("AX" & dragon + 1).Select
Selection.PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("DA2:DA" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("AZ" & dragon + 1).Select
Selection.PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("D2:D" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("C" & dragon + 1).Select
Selection.PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("Base_Comisiones").Range("Z2:Z" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("P" & dragon + 1).Select
Selection.PasteSpecial xlPasteValues
Windows(tt).Activate
ActiveWorkbook.Close SaveChanges:=False
MsgBox "proceso FINAL FINAL completado CAPO, puede continuar con el siguiente paso", vbInformation
MsgBox "Por favor espere DOS minutos mientras se enfría el procesador, de lo contrario la memoria se desboradará", vbInformation
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub espagnolo()
Dim elon As Integer
Dim crew As String
Dim musk As String
crew = "D:\AUTOMA_FULL_NEGOCIOS\ARCHIVO PLANO\CONSULTA_ESTADOS.txt"
elon = FreeFile
Open crew For Input As elon
musk = Input(LOF(elon), elon)
Close elon
musk = Replace(musk, "b", "',")
elon = FreeFile
Open crew For Output As elon
Print #elon, musk
Close elon
End Sub
Sub rakuten()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Sheets("Base Fijos").Select
Range("BI1").Select
ActiveCell.FormulaR1C1 = "=COUNTA(C1)-1"
luna = Range("BI1").Value
Range("A1").Select
dragon = Selection.End(xlDown).Row
Range("Y2").Select
For i = 1 To luna
ActiveCell.FormulaR1C1 = _
"=IF(RC[-1]=""UP SELLING"",0,IF(RC[11]=""HFC"", IF(RC[-7]>=36, RC[-8]*0.2, IF(RC[-7]>=24, RC[-8]*0.15, IF(RC[-7]>=18, RC[-8]*0.1,0))), IF(RC[11]=""FO"", IF(RC[-7]>=36, RC[-8]*0.3, IF(RC[-7]>=24, RC[-8]*0.2, IF(RC[-7]>=18, RC[-8]*0.15,0))),0)))"
ActiveCell.Offset(0, 1).Range("A1").Select
ActiveCell.FormulaR1C1 = _
"=IF(RC[-2]=""UP SELLING"",0,IF(RC[10]=""HFC"", IF(RC[27]>=50, RC[-9]*0.2, IF(RC[27]>=20, RC[-9]*0.1,0)), IF(RC[10]=""FO"", IF(RC[27]>=50, RC[-9]*0.3, IF(RC[27]>=20, RC[-9]*0.2,0)),0)))"
ActiveCell.Offset(0, 1).Range("A1").Select
ActiveCell.FormulaR1C1 = "=SUM(RC[-10],RC[-2],RC[-1])"
ActiveCell.Offset(1, 0).Select
ActiveCell.Offset(0, -2).Select
Next i
MsgBox "LISTO PRIMERA VALIDACIÓN", vbInformation
Range("BA2").Select
For j = 1 To luna
ActiveCell.FormulaR1C1 = _
"=IF(RC[-43]=""INTERNET"",IF(LEFT(RC[-44],1)=""M"",SUBSTITUTE(RC[-44],""M"","""")*1,IF(ISNUMBER(RC[-44]*1),RC[-44]*1/1000,0)),0)"
ActiveCell.Offset(0, 1).Range("A1").Select
ActiveCell.FormulaR1C1 = "=RC[-37]"
ActiveCell.Offset(1, 0).Select
ActiveCell.Offset(0, -1).Select
Next j
MsgBox "LISTO SEGUNDA VALIDACIÓN", vbInformation
Range("AR2") = 0
Range("AR2").Copy
Range("AR2:AR" & dragon).Select
Selection.PasteSpecial xlPasteValues
Application.CutCopyMode = False
Range("BA2").Select
MsgBox "LISTO TERCERA VALIDACIÓN", vbInformation
Sheets("Base Fijos").Select
Set rangodatos = Sheets("Base Fijos").UsedRange
rangodatos.AutoFilter Field:=39, Criteria1:=Array("IAAS", "SAAS", "PAAS"), _
Operator:=xlFilterValues
Range("J1").Select
With Selection.Interior
.Pattern = xlNone
.TintAndShade = 0
.PatternTintAndShade = 0
End With
With Selection.Font
.ColorIndex = xlAutomatic
.TintAndShade = 0
End With
Selection.Font.Bold = False
ActiveCell.Offset(0, 0) = "CLOUDb"
Range("J1").Copy
Range(Selection, Selection.End(xlDown)).Select
Selection.SpecialCells(xlCellTypeVisible).Select
ActiveSheet.Paste
Application.CutCopyMode = False
Range("J1").Select
With Selection.Interior
.Pattern = xlSolid
.PatternColorIndex = xlAutomatic
.ThemeColor = xlThemeColorLight1
.TintAndShade = 0
.PatternTintAndShade = 0
End With
With Selection.Font
.ThemeColor = xlThemeColorDark1
.TintAndShade = 0
End With
Selection.Font.Bold = True
Range("J1") = "LÍNEAS"
Range("A1").Select
ActiveSheet.ShowAllData
MsgBox "proceso FINAL FINAL completado CAPO, puede continuar con el siguiente paso", vbInformation
MsgBox "Por favor espere DOS minutos mientras se enfría el procesador, de lo contrario la memoria se desboradará", vbInformation
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub tesla()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim archivos
Dim goku As Excel.Workbook
Dim vegueta As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String
archivos = Dir("D:\AUTOMA_FULL_NEGOCIOS\MOVIL\*.xlsx")
Do While archivos <> ""
Workbooks.Open "D:\AUTOMA_FULL_NEGOCIOS\MOVIL\" & archivos
archivos = Dir
Loop
bills = ActiveWorkbook.Name
tt = bills
Windows(tt).Activate
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Meses").Select
calamaro = WorksheetFunction.VLookup(Range("H1").Value, Sheets("Meses").Range("A:I"), 9, 0)
odin = "2021" & calamaro
Windows(tt).Activate
Sheets("DETALLE").Select
Sheets("DETALLE").ListObjects("Tabla1").Unlist
Set rangodatos = Sheets("DETALLE").UsedRange
rangodatos.AutoFilter Field:=2, Criteria1:=odin
Set rangodatos = Sheets("DETALLE").UsedRange
rangodatos.AutoFilter Field:=1, Criteria1:=Array("Indirecto", "Directo", "TMK", "TMK_Upselling"), _
Operator:=xlFilterValues
Set rangodatos = Sheets("DETALLE").UsedRange
rangodatos.AutoFilter Field:=20, Criteria1:=Array("PAQUETE ALTAS", "PREPOS", "ALTAS", "UP"), _
Operator:=xlFilterValues
Set rangodatos = Sheets("DETALLE").UsedRange
rangodatos.AutoFilter Field:=23, Criteria1:="<>0", Operator:=xlAnd, Criteria2:="<>"
Range("A1").Select
col = Selection.End(xlDown).Row
Sheets("DETALLE").Range("A2:R" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Móvil").Select
Range("A2").PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("DETALLE").Range("T2:W" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Móvil").Select
Range("T2").PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("DETALLE").Range("AA2:AC" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Móvil").Select
Range("AB2").PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("DETALLE").Range("AE2:AE" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Móvil").Select
Range("AF2").PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("DETALLE").Range("AF2:AF" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Móvil").Select
Range("AH2").PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("DETALLE").Range("AK2:AK" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Móvil").Select
Range("AK2").PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("DETALLE").Range("AI2:AI" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Móvil").Select
Range("AL2").PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("DETALLE").Range("AH2:AH" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Móvil").Select
Range("AM2").PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("DETALLE").Range("S2:S" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Móvil").Select
Range("AO2").PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("DETALLE").Range("AS2:AS" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Móvil").Select
Range("AP2").PasteSpecial xlPasteValues
Application.CutCopyMode = False
Windows(tt).Activate
ActiveWorkbook.Close SaveChanges:=False
MsgBox "LISTO EL PEGUE 1", vbInformation
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub backtesla()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim archivos
Dim goku As Excel.Workbook
Dim vegueta As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String
archivos = Dir("D:\AUTOMA_FULL_NEGOCIOS\MOVIL VIEJO\*.xlsx")
Do While archivos <> ""
Workbooks.Open "D:\AUTOMA_FULL_NEGOCIOS\MOVIL VIEJO\" & archivos
archivos = Dir
Loop
bills = ActiveWorkbook.Name
tt = bills
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Meses").Select
KING = WorksheetFunction.VLookup(Range("H1").Value, Sheets("Meses").Range("A:B"), 2, 0) - 30
bongo = Format(KING, "mmmm")
xbox = WorksheetFunction.VLookup(bongo, Sheets("Meses").Range("A:I"), 9, 0)
thor = "2021" & xbox
Windows(tt).Activate
Sheets("DETALLE").Select
Set rangodatos = Sheets("DETALLE").UsedRange
rangodatos.AutoFilter Field:=2, Criteria1:=thor
Set rangodatos = Sheets("DETALLE").UsedRange
rangodatos.AutoFilter Field:=20, Criteria1:=Array("PREPOS", "ALTAS"), _
Operator:=xlFilterValues
Set rangodatos = Sheets("DETALLE").UsedRange
rangodatos.AutoFilter Field:=31, Criteria1:=Array("Documentos no legalizados", "Documentos no solucionados"), _
Operator:=xlFilterValues
Range("A1").Select
col = Selection.End(xlDown).Row
Sheets("DETALLE").Range("A2:E" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Móvil").Select
Range("A1").Select
dragon = Selection.End(xlDown).Row
Range("A" & dragon + 1).Select
Selection.PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("DETALLE").Range("J2:AD" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Móvil").Select
Range("J" & dragon + 1).Select
Selection.PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("DETALLE").Range("AF2:AP" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Móvil").Select
Range("AF" & dragon + 1).Select
Selection.PasteSpecial xlPasteValues
Windows(tt).Activate
ActiveWorkbook.Close SaveChanges:=False
MsgBox "proceso FINAL FINAL completado CAPO, puede continuar con el siguiente paso", vbInformation
MsgBox "Por favor espere DOS minutos mientras se enfría el procesador, de lo contrario la memoria se desboradará", vbInformation
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub PROCOLOMBIA()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim archivos, motorola As String
Dim goku As Excel.Workbook
Dim vegueta As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String
Workbooks.Add
bills = ActiveWorkbook.Name
tt = bills
Windows(tt).Activate
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Móvil").Select
Range("A1").Select
Set rangodatos = Sheets("Base Móvil").UsedRange
rangodatos.AutoFilter Field:=3, Criteria1:="<>0", _
Operator:=xlAnd, Criteria2:="<>"
rangodatos.AutoFilter Field:=20, Criteria1:=Array("ALTAS", "PREPOS"), Operator:=xlFilterValues
Range("A1").Select
dragon = Selection.End(xlDown).Row
Range("C1").Select
Sheets("Base Móvil").Range("C1:C" & dragon).Copy
Windows(tt).Activate
Sheets("Hoja1").Select
Range("A1").Select
ActiveSheet.Paste
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Range("A1").Select
ActiveSheet.ShowAllData
Windows(tt).Activate
Rows(1).EntireRow.Delete
Range("A1").Select
motorola = "D:\AUTOMA_FULL_NEGOCIOS\ARCHIVO PLANO\"
ActiveWorkbook.SaveAs Filename:= _
motorola & "CONSULTA_LEGALIZACION.txt" _
, FileFormat:=xlText, CreateBackup:=False
Windows("CONSULTA_LEGALIZACION.txt").Activate
ActiveWorkbook.Close SaveChanges:=False
MsgBox "proceso FINAL FINAL completado CAPO, puede continuar con el siguiente paso", vbInformation
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub mexico_86()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim archivos, motorola As String
Dim goku As Excel.Workbook
Dim vegueta As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String
Workbooks.Add
bills = ActiveWorkbook.Name
tt = bills
Windows(tt).Activate
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Móvil").Select
Range("A1").Select
Set rangodatos = Sheets("Base Móvil").UsedRange
rangodatos.AutoFilter Field:=3, Criteria1:="<>0", _
Operator:=xlAnd, Criteria2:="<>"
Range("A1").Select
dragon = Selection.End(xlDown).Row
Range("C1").Select
Sheets("Base Móvil").Range("C1:C" & dragon).Copy
Windows(tt).Activate
Sheets("Hoja1").Select
Range("A1").Select
ActiveSheet.Paste
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Range("A1").Select
ActiveSheet.ShowAllData
Windows(tt).Activate
Range("A1").Select
Rows(1).EntireRow.Delete
Range("A1").Select
genius = Selection.End(xlDown).Row
Columns("A:A").Select
ActiveSheet.Range("$A$1:A" & genius).RemoveDuplicates Columns:=1, Header:=xlNo
Range("A1").Select
genius = Selection.End(xlDown).Row
Range("B1").Select
ActiveCell.FormulaR1C1 = "a"
ActiveCell.Offset(0, 1).Select
ActiveCell.FormulaR1C1 = "b"
ActiveCell.Offset(0, 1).Select
ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[-2],RC[-3],RC[-1])"
MsgBox "LISTO PRIMER VALIDADOR", vbInformation
Range("B1:D1").Select
Selection.Copy
Range("A1").Select
Selection.End(xlDown).Select
ActiveCell.Offset(0, 1).Select
Range(Selection, Selection.End(xlUp)).Select
ActiveSheet.Paste
Range("A1").Select
Application.CutCopyMode = False
Columns("D:D").Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
:=False, Transpose:=False
Application.CutCopyMode = False
Range("B1").Select
Range("A:C").Columns.Delete
motorola = "D:\AUTOMA_FULL_NEGOCIOS\ARCHIVO PLANO\"
ActiveWorkbook.SaveAs Filename:= _
motorola & "CONSULTA_ESTADOS.txt" _
, FileFormat:=xlText, CreateBackup:=False
Windows("CONSULTA_ESTADOS.txt").Activate
ActiveWorkbook.Close SaveChanges:=False
MsgBox "LISTO SEGUNDO VALIDADOR", vbInformation
Call napoli
Call espagnolo
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Móvil").Select
Range("A1").Select
MsgBox "proceso FINAL FINAL completado CAPO, puede continuar con el siguiente paso", vbInformation
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub BOEING()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Móvil").Select
Range("A1").Select
dragon = Selection.End(xlDown).Row
Range("X2").Select
ActiveCell.FormulaR1C1 = _
"=IF(ISNUMBER(SEARCH(""Enlacedirecto"",RC[4])),0,IF(ISNUMBER(SEARCH(""Enl.Direc"",RC[4])),0,IF(AND(OR(RC[-4]=""ALTAS"",RC[-4]=""PREPOS""),RC[7]=""OK Fechas"",AND(RC[-1]>=120000,RC[-1]<149900)),RC[-1]*5%,IF(AND(OR(RC[-4]=""ALTAS"",RC[-4]=""PREPOS""),RC[7]=""OK Fechas"",RC[-1]>=150000),RC[-1]*10%,0))))"
ActiveCell.Offset(0, 1).Range("A1").Select
ActiveCell.FormulaR1C1 = _
"=IF(ISNUMBER(SEARCH(""Enlacedirecto"",RC[3])),0,IF(ISNUMBER(SEARCH(""Enl.Direc"",RC[3])),0,IF(AND(OR(RC[-5]=""ALTAS"",RC[-5]=""PREPOS""),RC[6]=""OK Fechas"",(IF(ISNUMBER(SEARCH(""PORTABILIDAD"",RC[5])),""Portabilidad"",0))=""Portabilidad""),(RC[-2]*5%),0)))"
ActiveCell.Offset(0, 1).Range("A1").Select
ActiveCell.FormulaR1C1 = _
"=IF(ISNUMBER(SEARCH(""Enlacedirecto"",RC[2])),0,IF(ISNUMBER(SEARCH(""Enl.Direc"",RC[2])),0,IF(AND(OR(RC[-6]=""ALTAS"",RC[-6]=""PREPOS""),RC[5]=""Ok Fechas"",RC[10]>=18,RC[10]<24),RC[-3]*10%,IF(AND(OR(RC[-6]=""ALTAS"",RC[-6]=""PREPOS""),RC[5]=""Ok Fechas"",RC[10]>=24,RC[10]<36),RC[-3]*15%,IF(AND(OR(RC[-6]=""ALTAS"",RC[-6]=""PREPOS""),RC[5]=""Ok Fechas"",RC[10]>=36),RC[-3]*20%,0)))))"
ActiveCell.Offset(0, 1).Range("A1").Select
ActiveCell.FormulaR1C1 = "=+RC[-4]+RC[-3]+RC[-2]+RC[-1]"
Range("X2:AA2").Select
Selection.Copy
Range("X3:X" & dragon).Select
ActiveSheet.Paste
Application.CutCopyMode = False
Range("X2").Select
MsgBox "LISTO SEGUNDA VALIDACIÓN", vbInformation
Range("AN2").Select
ActiveCell.FormulaR1C1 = _
        "=IF(ISNUMBER(SEARCH(""Enl.Direc"",RC[-12])),IF(IFERROR(VLOOKUP(RC[1],LIQUIDADOR!C2:C3,2,0),0)=""CONSULTOR"",RC[-17],0),IF(ISNUMBER(SEARCH(""Enlacedirecto"",RC[-12])),IF(IFERROR(VLOOKUP(RC[1],LIQUIDADOR!C2:C3,2,0),0)=""CONSULTOR"",RC[-17],0),IF(ISNUMBER(SEARCH(""Enlace Direc"",RC[-12])),IF(IFERROR(VLOOKUP(RC[1],LIQUIDADOR!C2:C3,2,0),0)=""CONSULTOR"",RC[-17],0),0)))"
ActiveCell.Copy
Range("AN3:AN" & dragon).Select
ActiveSheet.Paste
Application.CutCopyMode = False
Range("AN2").Select
MsgBox "proceso FINAL FINAL completado CAPO, puede continuar con el siguiente paso", vbInformation
MsgBox "Por favor espere DOS minutos mientras se enfría el procesador, de lo contrario la memoria se desboradará", vbInformation
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub marcolis()
Dim lala As Integer
Dim billions As Object
Set billions = CreateObject("Scripting.FileSystemObject")
lala = billions.CopyFile("D:\AUTOMA_FULL_NEGOCIOS\PLANTILLA_TOP_NEGOCIOS.xlsm", "D:\AUTOMA_FULL_NEGOCIOS\LIQ_FINAL\primeracopia.xlsm")
MsgBox "Validación ok", vbInformation
End Sub
Sub legacy()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim archivos
Dim goku As Excel.Workbook
Dim vegueta As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String
archivos = Dir("D:\AUTOMA_FULL_NEGOCIOS\LEGALIZACIÓN\*.xls")
If archivos = "" Then
MsgBox "NO HAY NINGÚN ARCHIVO DE PERMANENCIAS CARGADO, O LA EXTENSIÓN DEL ARCHIVO ESTÁ ERRADO", vbCritical
Else
Do While archivos <> ""
Workbooks.Open "D:\AUTOMA_FULL_NEGOCIOS\LEGALIZACIÓN\" & archivos
archivos = Dir
Loop
bills = ActiveWorkbook.Name
tt = bills
Windows(tt).Activate
Sheets("SQL Results").Select
Columns("C:C").Select
Selection.TextToColumns Destination:=Range("C1"), DataType:=xlDelimited, _
TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
:=Array(1, 1), TrailingMinusNumbers:=True
Range("C1").Select
Set goku = Workbooks.Open("D:\AUTOMA_FULL_NEGOCIOS\PLANTILLAS\LEGALI")
xiaomi = "LEGALI"
ss = xiaomi & ".xlsx"
Windows(tt).Activate
Range("B1").Select
dragon = Selection.End(xlDown).Row
Sheets("SQL Results").Range("B1:I" & dragon).Copy
Windows(ss).Activate
Range("A1").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Windows(tt).Activate
ActiveWorkbook.Close SaveChanges:=False
MsgBox "LISTO EL PEGUE", vbInformation
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Móvil").Select
Set rangodatos = Sheets("Base Móvil").UsedRange
rangodatos.AutoFilter Field:=20, Criteria1:=Array("ALTAS", "PREPOS"), Operator:=xlFilterValues
Range("a1").Select
ajo = Selection.End(xlDown).Row
Range("G1").Select
ActiveCell.FormulaR1C1 = _
"=IFERROR(VLOOKUP(RC[-4],'D:\AUTOMA_FULL_NEGOCIOS\PLANTILLAS\[LEGALI.xlsx]Hoja1'!C2:C8,7,FALSE),"""")"
ActiveCell.Offset(0, 1).Range("A1").Select
ActiveCell.FormulaR1C1 = _
"=IFERROR(VLOOKUP(RC[-5],'D:\AUTOMA_FULL_NEGOCIOS\PLANTILLAS\[LEGALI.xlsx]Hoja1'!C2:C4,3,FALSE),"""")"
ActiveCell.Offset(0, 1).Range("A1").Select
ActiveCell.FormulaR1C1 = _
"=IFERROR(VLOOKUP(RC[-6],'D:\AUTOMA_FULL_NEGOCIOS\PLANTILLAS\[LEGALI.xlsx]Hoja1'!C2:C6,5,FALSE),"""")"
Range("G1:I1").Select
Selection.Copy
Sheets("Base Móvil").Range("G1:I" & ajo).Select
Selection.SpecialCells(xlCellTypeVisible).Select
ActiveSheet.Paste
Range("A1").Select
ActiveSheet.ShowAllData
Windows(ss).Activate
ActiveWorkbook.Close SaveChanges:=True
MsgBox "EL SISTEMA SE CERRARÁ, POR FAVOR NO MANIPULE EL COMPUTADOR HASTA QUE EL EJECUTABLE DE EXCEL SE CIERRE, VUELVA ABRIR EL ARCHIVO Y EJECUTE EL BOTÓN CÁLCULOS LEGALIZACIÓN", vbInformation
ActiveWorkbook.Close SaveChanges:=True
End If
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub napoli()
Dim elon As Integer
Dim crew As String
Dim musk As String
crew = "D:\AUTOMA_FULL_NEGOCIOS\ARCHIVO PLANO\CONSULTA_ESTADOS.txt"
elon = FreeFile
Open crew For Input As elon
musk = Input(LOF(elon), elon)
Close elon
musk = Replace(musk, "a", "'")
elon = FreeFile
Open crew For Output As elon
Print #elon, musk
Close elon
End Sub
Sub lehman()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim archivos, motorola As String
Dim goku As Excel.Workbook
Dim vegueta As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String
Workbooks.Add
bills = ActiveWorkbook.Name
tt = bills
Windows(tt).Activate
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Móvil").Select
Range("A1").Select
Set rangodatos = Sheets("Base Móvil").UsedRange
rangodatos.AutoFilter Field:=3, Criteria1:="<>0", _
Operator:=xlAnd, Criteria2:="<>"
rangodatos.AutoFilter Field:=20, Criteria1:=Array("ALTAS", "PREPOS"), Operator:=xlFilterValues
Range("A1").Select
dragon = Selection.End(xlDown).Row
Range("C1").Select
Sheets("Base Móvil").Range("C1:C" & dragon).Copy
Windows(tt).Activate
Sheets("Hoja1").Select
Range("A1").Select
ActiveSheet.Paste
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Range("A1").Select
ActiveSheet.ShowAllData
Windows(tt).Activate
Range("A1").Select
Rows(1).EntireRow.Delete
Range("A1").Select
genius = Selection.End(xlDown).Row
Range("B1").Select
ActiveCell.FormulaR1C1 = "a"
ActiveCell.Offset(0, 1).Select
ActiveCell.FormulaR1C1 = "b"
ActiveCell.Offset(0, 1).Select
ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[-2],RC[-3],RC[-1])"
MsgBox "LISTO PRIMER VALIDADOR", vbInformation
Range("B1:D1").Select
Selection.Copy
Range("A1").Select
Selection.End(xlDown).Select
ActiveCell.Offset(0, 1).Select
Range(Selection, Selection.End(xlUp)).Select
ActiveSheet.Paste
Range("A1").Select
Application.CutCopyMode = False
Columns("D:D").Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
:=False, Transpose:=False
Application.CutCopyMode = False
Range("B1").Select
Range("A:C").Columns.Delete
motorola = "D:\AUTOMA_FULL_NEGOCIOS\ARCHIVO PLANO\"
ActiveWorkbook.SaveAs Filename:= _
motorola & "CONSULTA_PERMANENCIA.txt" _
, FileFormat:=xlText, CreateBackup:=False
Windows("CONSULTA_PERMANENCIA.txt").Activate
ActiveWorkbook.Close SaveChanges:=False
MsgBox "LISTO SEGUNDO VALIDADOR", vbInformation
Call functlon
Call funtion
MsgBox "proceso FINAL FINAL completado CAPO, puede continuar con el siguiente paso", vbInformation
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub gw()
Call rugerri
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Móvil").Select
Columns("F:I").Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
:=False, Transpose:=False
Application.CutCopyMode = False
Range("F1").Select
Range("F1").Select
Selection.Copy
Range("G1:I1").Select
Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
SkipBlanks:=False, Transpose:=False
Range("G1").Select
Application.CutCopyMode = False
Range("F1").Select
Selection.Copy
Range("G1:I1").Select
Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
SkipBlanks:=False, Transpose:=False
Range("G1").Select
Application.CutCopyMode = False
Set rangodatos = Sheets("Base Móvil").UsedRange
rangodatos.AutoFilter Field:=20, Criteria1:=Array("PAQUETE ALTAS", "UP"), Operator:=xlFilterValues
Range("A1").Select
favs = Selection.End(xlDown).Row
Sheets("Meses").Select
Range("XFD1").Copy
Sheets("Base Móvil").Select
Sheets("Base Móvil").Range("AE1:AE" & favs).Select
ActiveSheet.Paste
Range("A1").Select
ActiveSheet.ShowAllData
Set rangodatos = Sheets("Base Móvil").UsedRange
rangodatos.AutoFilter Field:=20, Criteria1:=Array("ALTAS", "PREPOS"), Operator:=xlFilterValues
Set rangodatos = Sheets("Base Móvil").UsedRange
rangodatos.AutoFilter Field:=7, Criteria1:="<>"
Range("A1").Select
fav = Selection.End(xlDown).Row
Sheets("Meses").Select
Range("XFD1").Copy
Sheets("Base Móvil").Select
Sheets("Base Móvil").Range("AE1:AE" & fav).Select
ActiveSheet.Paste
Range("A1").Select
ActiveSheet.ShowAllData
Set rangodatos = Sheets("Base Móvil").UsedRange
rangodatos.AutoFilter Field:=20, Criteria1:=Array("ALTAS", "PREPOS"), _
Operator:=xlFilterValues
Set rangodatos = Sheets("Base Móvil").UsedRange
rangodatos.AutoFilter Field:=7, Criteria1:="="
Set rangodatos = Sheets("Base Móvil").UsedRange
rangodatos.AutoFilter Field:=9, Criteria1:="<>"
Range("A1").Select
cr7 = Selection.End(xlDown).Row
Sheets("Meses").Select
Range("XFD2").Copy
Sheets("Base Móvil").Select
Sheets("Base Móvil").Range("AE1:AE" & cr7).Select
ActiveSheet.Paste
Range("A1").Select
ActiveSheet.ShowAllData
Set rangodatos = Sheets("Base Móvil").UsedRange
rangodatos.AutoFilter Field:=20, Criteria1:=Array("ALTAS", "PREPOS"), _
Operator:=xlFilterValues
Set rangodatos = Sheets("Base Móvil").UsedRange
rangodatos.AutoFilter Field:=7, Criteria1:="="
Set rangodatos = Sheets("Base Móvil").UsedRange
rangodatos.AutoFilter Field:=9, Criteria1:="="
Range("A1").Select
cvodi = Selection.End(xlDown).Row
Sheets("Meses").Select
Range("XFD3").Copy
Sheets("Base Móvil").Select
Sheets("Base Móvil").Range("AE1:AE" & cvodi).Select
ActiveSheet.Paste
Range("A1").Select
ActiveSheet.ShowAllData
Range("AE1") = "OBSERVACIÓN"
Range("AE1").Select
With Selection.Interior
.Pattern = xlSolid
.PatternColorIndex = xlAutomatic
.ThemeColor = xlThemeColorLight1
.TintAndShade = 0
.PatternTintAndShade = 0
End With
With Selection.Font
.ThemeColor = xlThemeColorDark1
.TintAndShade = 0
End With
Selection.Font.Bold = True
Range("G1") = "FECHA LEGALIZACION"
Range("H1") = "FECHA RECEPCION"
Range("I1") = "FECHA DEVOLUCION"
MsgBox "proceso FINAL FINAL completado CAPO, puede continuar con el siguiente paso", vbInformation
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub killbill()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim archivos
Dim goku As Excel.Workbook
Dim vegueta As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String
archivos = Dir("D:\AUTOMA_FULL_NEGOCIOS\MOVIL\*.xls")
Do While archivos <> ""
Workbooks.Open "D:\AUTOMA_FULL_NEGOCIOS\MOVIL\" & archivos
archivos = Dir
Loop
bills = ActiveWorkbook.Name
tt = bills
Windows(tt).Activate
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Meses").Select
calamaro = WorksheetFunction.VLookup(Range("H1").Value, Sheets("Meses").Range("A:I"), 9, 0)
odin = "2020" & calamaro
Windows(tt).Activate
Sheets("DETALLE").Select
Set rangodatos = Sheets("DETALLE").UsedRange
rangodatos.AutoFilter Field:=2, Criteria1:=odin
Set rangodatos = Sheets("DETALLE").UsedRange
rangodatos.AutoFilter Field:=1, Criteria1:=Array("Indirecto", "Directo", "TMK", "TMK_Upselling"), _
Operator:=xlFilterValues
Sheets("DETALLE").Select
Set rangodatos = Sheets("DETALLE").UsedRange
rangodatos.AutoFilter Field:=20, Criteria1:=Array("DOWN", "POSPRE_TRE", "BAJAS", "POSPRE", "PAQUETE BAJAS"), _
Operator:=xlFilterValues
Sheets("DETALLE").Select
Range("A1").Select
col = Selection.End(xlDown).Row
Sheets("DETALLE").Range("A2:Y" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Bajas Móvil").Select
Range("A2").PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("DETALLE").Range("Z2:AC" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Bajas Móvil").Select
Range("AA2").PasteSpecial xlPasteValues
Windows(tt).Activate
Sheets("DETALLE").Range("AF2:AK" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Bajas Móvil").Select
Range("AI2").PasteSpecial xlPasteValues
Windows(tt).Activate
ActiveWorkbook.Close SaveChanges:=False
MsgBox "LISTO PRIMER VALIDADOR", vbInformation
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Range("A1").Select
alien = Selection.End(xlDown).Row - 1
Range("AE2").Select
For i = 1 To alien
ActiveCell.FormulaR1C1 = "OK Fechas"
ActiveCell.Offset(0, 1).Range("A1").Select
ActiveCell.FormulaR1C1 = "OK Fechas"
ActiveCell.Offset(1, 0).Select
ActiveCell.Offset(0, -1).Select
Next i
MsgBox "proceso FINAL FINAL completado CAPO, puede continuar con el siguiente paso", vbInformation
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub rugerri()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Móvil").Select
Set rangodatos = Sheets("Base Móvil").UsedRange
rangodatos.AutoFilter Field:=7, Criteria1:="0"
Columns("G:G").Select
Selection.ClearContents
Range("G1").Select
Range("A1").Select
ActiveSheet.ShowAllData
Set rangodatos = Sheets("Base Móvil").UsedRange
rangodatos.AutoFilter Field:=8, Criteria1:="0"
Columns("H:H").Select
Selection.ClearContents
Range("G1").Select
Range("A1").Select
ActiveSheet.ShowAllData
Set rangodatos = Sheets("Base Móvil").UsedRange
rangodatos.AutoFilter Field:=9, Criteria1:="0"
Columns("I:I").Select
Selection.ClearContents
Range("G1").Select
Range("A1").Select
ActiveSheet.ShowAllData
MsgBox "LISTO PRIMER VALIDADOR", vbInformation
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub oud()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim archivos
Dim goku As Excel.Workbook
Dim vegueta As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String
archivos = Dir("D:\AUTOMA_FULL_NEGOCIOS\PERMANENCIA\*nen*")
If archivos = "" Then
MsgBox "NO HAY NINGÚN ARCHIVO DE PERMANENCIAS CARGADO, O LA EXTENSIÓN DEL ARCHIVO ESTÁ ERRADO", vbCritical
Else
Do While archivos <> ""
Workbooks.Open "D:\AUTOMA_FULL_NEGOCIOS\PERMANENCIA\" & archivos
archivos = Dir
Loop
bills = ActiveWorkbook.Name
tt = bills
Windows(tt).Activate
Columns("A:A").Select
Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
Semicolon:=False, Comma:=True, Space:=False, Other:=False, FieldInfo _
:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), _
Array(7, 1), Array(8, 1), Array(9, 1)), TrailingMinusNumbers:=True
Range("A1").Select
Set goku = Workbooks.Open("D:\AUTOMA_FULL_NEGOCIOS\PLANTILLAS\PERMA")
xiaomi = "PERMA"
ss = xiaomi & ".xlsx"
Windows(tt).Activate
col = Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
Sheets(1).Range("A1:I" & col).Copy
Windows(ss).Activate
Range("A1").Select
ActiveSheet.Paste
Windows(tt).Activate
ActiveWorkbook.Close SaveChanges:=False
MsgBox "LISTO EL PEGUE", vbInformation
Windows(ss).Activate
Range("H:H").Columns.Insert
Range("J:J").Columns.Insert
Range("A1").Select
core = Selection.End(xlDown).Row
Range("H2").Select
ActiveCell.FormulaR1C1 = "=IF(LEN(RC[-1])>5,0,RC[-1])"
ActiveCell.Copy
Range("H3:H" & core).Select
ActiveSheet.Paste
Application.CutCopyMode = False
MsgBox "listo primer largo", vbInformation
Range("J2").Select
ActiveCell.FormulaR1C1 = "=IF(LEN(RC[-1])>5,0,RC[-1])"
ActiveCell.Copy
Range("J3:J" & core).Select
ActiveSheet.Paste
Application.CutCopyMode = False
MsgBox "listo segundo largo", vbInformation
Range("H1") = "PACTADA"
Range("J1") = "PENDIENTE"
 Columns("H:H").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
Columns("J:J").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("J1").Select
MsgBox "OK TERMINADO", vbInformation
Columns("G:G").Delete
Columns("H:H").Delete
Range("J2").Select
ActiveCell.FormulaR1C1 = "=SUM(RC[-3],RC[-2])"
ActiveCell.Copy
Range("J3:J" & core).Select
ActiveSheet.Paste
Application.CutCopyMode = False

Range("J1") = "SUMATORIA"
MsgBox "LISTO VALIDADOR", vbInformation
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Móvil").Select
Range("A1").Select
fav = Selection.End(xlDown).Row
Range("AJ2").Select
ActiveCell.FormulaR1C1 = _
"=IFERROR(VLOOKUP(RC[-33],[PERMA.xlsx]DETALLE!C1:C10,10,FALSE),0)"
ActiveCell.Copy
Range("AJ3:AJ" & fav).Select
ActiveSheet.Paste
Application.CutCopyMode = False

Windows(ss).Activate
ActiveWorkbook.Close SaveChanges:=True
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Columns("AJ:AJ").Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
:=False, Transpose:=False
Application.CutCopyMode = False
Range("AJ1").Select
MsgBox "LISTO PRIMERA FASE", vbInformation
End If
Call copola
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub scribe()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim archivos
Dim goku As Excel.Workbook
Dim vegueta As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String
archivos = Dir("D:\AUTOMA_FULL_NEGOCIOS\BAJAS FIJO\*.xlsx")
Do While archivos <> ""
Workbooks.Open "D:\AUTOMA_FULL_NEGOCIOS\BAJAS FIJO\" & archivos
archivos = Dir
Loop
bills = ActiveWorkbook.Name
tt = bills
Windows(tt).Activate
Sheets("Detalle").Select
Range("A1").Select
col = Sheets("Detalle").Range("A" & Rows.Count).End(xlUp).Row
Sheets("Detalle").Range("A2:AC" & col).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Bajas Fijo").Select
Range("A2").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Windows(tt).Activate
ActiveWorkbook.Close SaveChanges:=True
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Range("L:L").Columns.Insert
Range("L1") = "COMISIONES"
Range("A1").Select
alien = Selection.End(xlDown).Row - 1
Range("L2").Select
For i = 1 To alien
ActiveCell.FormulaR1C1 = "=RC[-1]*-1"
ActiveCell.Offset(1, 0).Select
Next i
Range("L1").Select
With Selection.Interior
.Pattern = xlSolid
.Color = 65535
.TintAndShade = 0
.PatternTintAndShade = 0
End With
MsgBox "proceso FINAL FINAL completado CAPO, puede continuar con el siguiente paso", vbInformation
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub margin()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("HC").Select
Set rangodatos = Sheets("HC").UsedRange
rangodatos.AutoFilter Field:=5, Criteria1:=Array("Consultor Negocios", "Consultor Cuentas Corporativas Regional SN", "Consultor Intermedias Segmento Negocios"), _
Operator:=xlFilterValues
Range("A1").Select
ActiveCell.Offset(1, 0).Select
Range(Selection, Selection.End(xlDown)).Copy
Sheets("LIQUIDADOR").Select
Range("B4").Select
ActiveSheet.Paste
Sheets("HC").Select
Range("A1").Select
ActiveSheet.ShowAllData
Set rangodatos = Sheets("HC").UsedRange
rangodatos.AutoFilter Field:=5, Criteria1:=Array("Especialista Canales Negocios", "Especialista Comercial Negocios", "Especialista Comercial Telemercadeo Empresas Y Negocios", "Especialista Telemercadeo Empresas Y Negocios", "Especialista Ventas Sinergia Empresas"), _
Operator:=xlFilterValues
Range("A1").Select
ActiveCell.Offset(1, 0).Select
Range(Selection, Selection.End(xlDown)).Copy
Sheets("LIQUIDADOR").Select
Range("B3").Select
ad = Selection.End(xlDown).Row
Sheets("LIQUIDADOR").Range("B" & ad + 1).Select
ActiveSheet.Paste
Sheets("HC").Select
Range("A1").Select
ActiveSheet.ShowAllData
Set rangodatos = Sheets("HC").UsedRange
rangodatos.AutoFilter Field:=5, Criteria1:=Array("Coordinador Comercial Negocios", "Coordinador Ventas Negocios"), _
Operator:=xlFilterValues
Range("A1").Select
ActiveCell.Offset(1, 0).Select
Range(Selection, Selection.End(xlDown)).Copy
Sheets("LIQUIDADOR").Select
Range("B3").Select
od = Selection.End(xlDown).Row
Sheets("LIQUIDADOR").Range("B" & od + 1).Select
ActiveSheet.Paste
Sheets("HC").Select
Range("A1").Select
ActiveSheet.ShowAllData
Set rangodatos = Sheets("HC").UsedRange
rangodatos.AutoFilter Field:=5, Criteria1:=Array("Jefe Comercial Negocios", "Jefe Comercial Telemercadeo Empresas Y Negocios"), _
Operator:=xlFilterValues
Range("A1").Select
ActiveCell.Offset(1, 0).Select
Range(Selection, Selection.End(xlDown)).Copy
Sheets("LIQUIDADOR").Select
Range("B3").Select
ud = Selection.End(xlDown).Row
Sheets("LIQUIDADOR").Range("B" & ud + 1).Select
ActiveSheet.Paste
Sheets("HC").Select
Range("A1").Select
ActiveSheet.ShowAllData
Set rangodatos = Sheets("HC").UsedRange
rangodatos.AutoFilter Field:=5, Criteria1:=Array("Gerente Comercial Negocios 1", "Gerente Comercial Negocios 2", "Gerente Comercial Negocios 3"), _
Operator:=xlFilterValues
Range("A1").Select
ActiveCell.Offset(1, 0).Select
Range(Selection, Selection.End(xlDown)).Copy
Sheets("LIQUIDADOR").Select
Range("B3").Select
ud = Selection.End(xlDown).Row
Sheets("LIQUIDADOR").Range("B" & ud + 1).Select
ActiveSheet.Paste
Sheets("HC").Select
Range("A1").Select
ActiveSheet.ShowAllData
Set rangodatos = Sheets("HC").UsedRange
rangodatos.AutoFilter Field:=5, Criteria1:="Gerente Comercial Telemercadeo Empresas Y Negocios"
Range("A1").Select
ActiveCell.Offset(1, 0).Select
Range(Selection, Selection.End(xlDown)).Copy
Sheets("LIQUIDADOR").Select
Range("B3").Select
ud = Selection.End(xlDown).Row
Sheets("LIQUIDADOR").Range("B" & ud + 1).Select
ActiveSheet.Paste
Sheets("HC").Select
Range("A1").Select
ActiveSheet.ShowAllData
Set rangodatos = Sheets("HC").UsedRange
rangodatos.AutoFilter Field:=5, Criteria1:="Jefe Canales Negocios"
Range("A1").Select
ActiveCell.Offset(1, 0).Select
Range(Selection, Selection.End(xlDown)).Copy
Sheets("LIQUIDADOR").Select
Range("B3").Select
ud = Selection.End(xlDown).Row
Sheets("LIQUIDADOR").Range("B" & ud + 1).Select
ActiveSheet.Paste
Sheets("HC").Select
Range("A1").Select
ActiveSheet.ShowAllData
Set rangodatos = Sheets("HC").UsedRange
rangodatos.AutoFilter Field:=5, Criteria1:="Director Negocios"
Range("A1").Select
ActiveCell.Offset(1, 0).Select
Range(Selection, Selection.End(xlDown)).Copy
Sheets("LIQUIDADOR").Select
Range("B3").Select
ud = Selection.End(xlDown).Row
Sheets("LIQUIDADOR").Range("B" & ud + 1).Select
ActiveSheet.Paste
Sheets("HC").Select
Range("A1").Select
ActiveSheet.ShowAllData
Sheets("LIQUIDADOR").Select
Range("B3").Select
fav = Selection.End(xlDown).Row - 3
Range("B4").Select
For i = 1 To fav
ActiveCell.Offset(0, 1).Range("A1").Select
ActiveCell.FormulaR1C1 = _
"=IF(ISNUMBER(SEARCH(""CONSULTOR"",RC[1])),""CONSULTOR"",IF(ISNUMBER(SEARCH(""COORDINADOR"",RC[1])),""COORDINADOR"",IF(ISNUMBER(SEARCH(""ESPECIALISTA"",RC[1])),""ESPECIALISTA"",IF(ISNUMBER(SEARCH(""DIRECTOR"",RC[1])),""DIRECTOR"",IF(ISNUMBER(SEARCH(""Gerente Comercial n"",RC[1])),""GERENTE"",IF(ISNUMBER(SEARCH(""Jefe Comercial"",RC[1])),""JEFE"",IF(ISNUMBER(SEARCH(""Jefe canales"",RC[1])),""JEFECAN"",IF(ISNUMBER(SEARCH(""Gerente Comercial Telemercadeo"",RC[1])),""JEFETMK"",IF(ISNUMBER(SEARCH(""Jefe Telemercadeo"",RC[1])),""JEFE"")))))))))"
ActiveCell.Offset(0, 1).Range("A1").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-2],HC!C1:C5,5,0)"
ActiveCell.Offset(1, 0).Select
ActiveCell.Offset(0, -2).Select
Next i
MsgBox "proceso FINAL FINAL completado CAPO, puede continuar con el siguiente paso", vbInformation
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub poc()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Call marcolis
Dim archivos
Dim goku As Excel.Workbook
Dim vegueta As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String
archivos = Dir("D:\AUTOMA_FULL_NEGOCIOS\LIQ_FINAL\*.xlsm")
Do While archivos <> ""
Workbooks.Open "D:\AUTOMA_FULL_NEGOCIOS\LIQ_FINAL\" & archivos
archivos = Dir
Loop
bills = ActiveWorkbook.Name
tt = bills
Windows(tt).Activate
Sheets("Base Móvil").Select
Cells.Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False
Range("A1").Select
Sheets("Base Fijos").Select
Cells.Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False
Range("A1").Select
Sheets("FX").Delete
Sheets("LIQUIDADOR").Delete
Sheets("HC").Delete
Sheets("Garantizados").Delete
Sheets("Vacaciones").Delete
Sheets("Meses").Delete
Sheets("Neto Movil").Delete
Sheets("Neto Fijo").Delete
Sheets("DESARROLLO+PROYECTOS").Delete
Sheets("Proyectos-Desarrollo").Delete
Sheets("Meta General").Delete
Sheets("Base Fijos").Delete
Sheets("SIN_TURNOS").Delete
Sheets("Base Móvil").Select
Range("A1").Select
RUTA = "D:\AUTOMA_FULL_NEGOCIOS\LIQ_FINAL\"
GExcel = RUTA & "BASE MOVIL" & ".xlsx"
ActiveWorkbook.SaveAs Filename:= _
GExcel, FileFormat:= _
xlOpenXMLWorkbook, CreateBackup:=False
bill = ActiveWorkbook.Name
pp = bill
Windows(pp).Activate
ActiveWorkbook.Close SaveChanges:=True
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub vlog()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim archivos
Dim goku As Excel.Workbook
Dim vegueta As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String
archivos = Dir("D:\AUTOMA_FULL_NEGOCIOS\LIQ_FINAL\*.xlsm")
Do While archivos <> ""
Workbooks.Open "D:\AUTOMA_FULL_NEGOCIOS\LIQ_FINAL\" & archivos
archivos = Dir
Loop
bills = ActiveWorkbook.Name
tt = bills
Windows(tt).Activate
Sheets("Base Móvil").Select
Cells.Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False
Range("A1").Select
Sheets("Base Fijos").Select
Cells.Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False
Range("A1").Select
Sheets("FX").Delete
Sheets("LIQUIDADOR").Delete
Sheets("HC").Delete
Sheets("Garantizados").Delete
Sheets("Vacaciones").Delete
Sheets("Meta General").Delete
Sheets("Meses").Delete
Sheets("Neto Fijo").Delete
Sheets("DESARROLLO+PROYECTOS").Delete
Sheets("Proyectos-Desarrollo").Delete
Sheets("SIN_TURNOS").Delete
Sheets("Neto Movil").Delete
Sheets("Base Móvil").Delete
Sheets("Base Fijos").Select
Range("A1").Select
RUTA = "D:\AUTOMA_FULL_NEGOCIOS\LIQ_FINAL\"
GENERAR = RUTA & "BASE FIJO" & ".xlsx"
ActiveWorkbook.SaveAs Filename:= _
GENERAR, FileFormat:= _
xlOpenXMLWorkbook, CreateBackup:=False
bill = ActiveWorkbook.Name
pp = bill
Windows(pp).Activate
ActiveWorkbook.Close SaveChanges:=True
Kill ("D:\AUTOMA_FULL_NEGOCIOS\LIQ_FINAL\*.xlsm")
Call EUCLIDEES
Call ergo
MsgBox "proceso FINAL FINAL completado CAPO, puede continuar con el siguiente paso", vbInformation
MsgBox "Por favor espere DOS minutos mientras se enfría el procesador, de lo contrario la memoria se desboradará", vbInformation
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub EUCLIDES()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Columns("F:F").Select
Selection.TextToColumns Destination:=Range("F1"), DataType:=xlDelimited, _
TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
:=Array(1, 1), TrailingMinusNumbers:=True
Range("F1").Select
Columns("AC:AC").Select
Selection.TextToColumns Destination:=Range("AC1"), DataType:=xlDelimited, _
TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
:=Array(1, 1), TrailingMinusNumbers:=True
Range("AC1").Select
Columns("AE:AE").Select
Selection.TextToColumns Destination:=Range("AE1"), DataType:=xlDelimited, _
TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
:=Array(1, 1), TrailingMinusNumbers:=True
Range("AE1").Select
Columns("AP:AP").Select
Selection.TextToColumns Destination:=Range("AP1"), DataType:=xlDelimited, _
TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
:=Array(1, 1), TrailingMinusNumbers:=True
Range("AP1").Select
Columns("AV:AV").Select
Selection.TextToColumns Destination:=Range("AV1"), DataType:=xlDelimited, _
TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
:=Array(1, 1), TrailingMinusNumbers:=True
Range("AV1").Select
Columns("AX:AX").Select
Selection.TextToColumns Destination:=Range("AX1"), DataType:=xlDelimited, _
TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
:=Array(1, 1), TrailingMinusNumbers:=True
Range("AX1").Select
Columns("AS:AS").Select
Selection.TextToColumns Destination:=Range("AS1"), DataType:=xlDelimited, _
TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
:=Array(1, 1), TrailingMinusNumbers:=True
Range("AS1").Select
Sheets("Base Móvil").Select
Columns("M:M").Select
Selection.TextToColumns Destination:=Range("M1"), DataType:=xlDelimited, _
TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
:=Array(1, 1), TrailingMinusNumbers:=True
Range("M1").Select
Columns("O:O").Select
Selection.TextToColumns Destination:=Range("O1"), DataType:=xlDelimited, _
TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
:=Array(1, 1), TrailingMinusNumbers:=True
Range("O1").Select
Columns("Q:Q").Select
Selection.TextToColumns Destination:=Range("Q1"), DataType:=xlDelimited, _
TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
:=Array(1, 1), TrailingMinusNumbers:=True
Range("Q1").Select
Columns("S:S").Select
Selection.TextToColumns Destination:=Range("S1"), DataType:=xlDelimited, _
TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
:=Array(1, 1), TrailingMinusNumbers:=True
Range("S1").Select
Columns("AF:AF").Select
Selection.TextToColumns Destination:=Range("AF1"), DataType:=xlDelimited, _
TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
:=Array(1, 1), TrailingMinusNumbers:=True
Range("AF1").Select
Columns("AH:AH").Select
Selection.TextToColumns Destination:=Range("AH1"), DataType:=xlDelimited, _
TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
:=Array(1, 1), TrailingMinusNumbers:=True
Range("AH1").Select
Columns("AO:AO").Select
Selection.TextToColumns Destination:=Range("AO1"), DataType:=xlDelimited, _
TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
:=Array(1, 1), TrailingMinusNumbers:=True
Range("AO1").Select
Sheets("Base Móvil").Select
Columns("AI:AI").Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
:=False, Transpose:=False
Application.CutCopyMode = False
Range("A1").Select
MsgBox "EL SISTEMA SE CERRARÁ, POR FAVOR NO MANIPULE EL COMPUTADOR HASTA QUE EL EJECUTABLE DE EXCEL SE CIERRE, VUELVA ABRIR EL ARCHIVO Y EJECUTE EL BOTÓN EXPORTAR BASES", vbInformation
ActiveWorkbook.Close SaveChanges:=True
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub EUCLIDEES()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim archivos
Dim goku As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String
Set goku = Workbooks.Open("D:\AUTOMA_FULL_NEGOCIOS\LIQ_FINAL\BASE FIJO")
xiaomi = "BASE FIJO"
ss = xiaomi & ".xlsx"
Windows(ss).Activate
Sheets("Base Fijos").Select
Columns("F:F").Select
Selection.TextToColumns Destination:=Range("F1"), DataType:=xlDelimited, _
TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
:=Array(1, 1), TrailingMinusNumbers:=True
Range("F1").Select
Columns("AC:AC").Select
Selection.TextToColumns Destination:=Range("AC1"), DataType:=xlDelimited, _
TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
:=Array(1, 1), TrailingMinusNumbers:=True
Range("AC1").Select
Columns("AE:AE").Select
Selection.TextToColumns Destination:=Range("AE1"), DataType:=xlDelimited, _
TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
:=Array(1, 1), TrailingMinusNumbers:=True
Range("AE1").Select
Columns("AP:AP").Select
Selection.TextToColumns Destination:=Range("AP1"), DataType:=xlDelimited, _
TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
:=Array(1, 1), TrailingMinusNumbers:=True
Range("AP1").Select
Columns("AV:AV").Select
Selection.TextToColumns Destination:=Range("AV1"), DataType:=xlDelimited, _
TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
:=Array(1, 1), TrailingMinusNumbers:=True
Range("AV1").Select
Columns("AX:AX").Select
Selection.TextToColumns Destination:=Range("AX1"), DataType:=xlDelimited, _
TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
:=Array(1, 1), TrailingMinusNumbers:=True
Range("AX1").Select
Columns("AS:AS").Select
Selection.TextToColumns Destination:=Range("AS1"), DataType:=xlDelimited, _
TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
:=Array(1, 1), TrailingMinusNumbers:=True
Range("AS1").Select
Windows(ss).Activate
ActiveWorkbook.Close SaveChanges:=True
MsgBox "LISTOS FORMATOS_1", vbInformation
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub pelikan()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Sheets("Meses").Select
Range("H1").ClearContents
Range("F1").ClearContents
Range("G1").ClearContents
Sheets("LIQUIDADOR").Select
Range("B3").Select
col = Selection.End(xlDown).Row
Sheets("LIQUIDADOR").Range("A4:BI" & col).Select
Selection.ClearContents
Range("A1").Select
Sheets("DESARROLLO+PROYECTOS").Select
Range("C2").Select
col = Selection.End(xlDown).Row
Sheets("DESARROLLO+PROYECTOS").Range("C3:AN" & col).Select
Selection.ClearContents
Range("A1").Select
Sheets("Base Fijos").Select
Range("A1").Select
co = Selection.End(xlDown).Row
Sheets("Base Fijos").Range("A2:BB" & co).Select
Selection.ClearContents
Range("BI1").ClearContents
Range("A1").Select
Sheets("SIN_TURNOS").Select
Range("A1").Select
co = Selection.End(xlDown).Row
Sheets("SIN_TURNOS").Range("A2:B" & co).Select
Selection.ClearContents
Range("F1").Select
co = Selection.End(xlDown).Row
Sheets("SIN_TURNOS").Range("F2:G" & co).Select
Selection.ClearContents
Range("A1").Select
Sheets("Base Móvil").Select
Range("A1").Select
coo = Selection.End(xlDown).Row
Sheets("Base Móvil").Range("A2:AR" & coo).Select
Selection.ClearContents
Range("A1").Select
Sheets("HC").Select
Range("A1").Select
cool = Selection.End(xlDown).Row
Sheets("HC").Range("A2:V" & cool).Select
Selection.ClearContents
Range("A1").Select
Sheets("Proyectos-Desarrollo").Select
Range("A3").Select
Range(Selection, Selection.End(xlToRight)).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.ClearContents
Range("H3").Select
Range(Selection, Selection.End(xlToRight)).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.ClearContents
Range("D3").Select
Sheets("Garantizados").Select
Range("A1").Select
coool = Selection.End(xlDown).Row
Sheets("Garantizados").Range("A2:F" & coool).Select
Selection.ClearContents
Range("H1").ClearContents
Range("A1").Select
Sheets("Vacaciones").Select
Range("A1").Select
cooool = Selection.End(xlDown).Row
Sheets("Vacaciones").Range("A2:G" & cooool).Select
Selection.ClearContents
Range("A1").Select
Sheets("Meta General").Select
Range("A1").Select
coll = Selection.End(xlDown).Row
Sheets("Meta General").Range("A2:L" & coll).Select
Selection.ClearContents
Range("S1").ClearContents
Range("XFD1").ClearContents
Range("XFD2").ClearContents
Range("A1").Select
Sheets("Base Móvil").Select
Range("G1:I1").Select
Selection.ClearContents
Range("G3").Select
Selection.Copy
Range("G1:I1").Select
Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
SkipBlanks:=False, Transpose:=False
Range("G1").Select
Application.CutCopyMode = False
Sheets("Base Móvil").Select
Range("AE1").ClearContents
Range("AE1").Select
With Selection.Interior
.Pattern = xlNone
.TintAndShade = 0
.PatternTintAndShade = 0
End With
Range("AE2").Select
Sheets("Neto Fijo").Select
Range("A2").Select
roncero = Selection.End(xlDown).Row
Sheets("Neto Fijo").Range("A3:G" & roncero).Select
Selection.ClearContents
Range("A1").Select
Sheets("Neto Movil").Select
Range("A2").Select
real = Selection.End(xlDown).Row
Sheets("Neto Movil").Range("A3:G" & real).Select
Selection.ClearContents
Range("A1").Select
MsgBox "LISTO PRIMERA FASE", vbInformation
Call golovin
MsgBox "Plantilla de liquidación borrada con éxito, el ejecutable de excel se cerrará", vbInformation
ActiveWorkbook.Close SaveChanges:=True
Application.Quit
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub ergo()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim archivos
Dim goku As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String
Set goku = Workbooks.Open("D:\AUTOMA_FULL_NEGOCIOS\LIQ_FINAL\BASE MOVIL")
xiaomi = "BASE MOVIL"
ss = xiaomi & ".xlsx"
Windows(ss).Activate
Sheets("Base Móvil").Select
Columns("M:M").Select
Selection.TextToColumns Destination:=Range("M1"), DataType:=xlDelimited, _
TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
:=Array(1, 1), TrailingMinusNumbers:=True
Range("M1").Select
Columns("O:O").Select
Selection.TextToColumns Destination:=Range("O1"), DataType:=xlDelimited, _
TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
:=Array(1, 1), TrailingMinusNumbers:=True
Range("O1").Select
Columns("Q:Q").Select
Selection.TextToColumns Destination:=Range("Q1"), DataType:=xlDelimited, _
TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
:=Array(1, 1), TrailingMinusNumbers:=True
Range("Q1").Select
Columns("S:S").Select
Selection.TextToColumns Destination:=Range("S1"), DataType:=xlDelimited, _
TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
:=Array(1, 1), TrailingMinusNumbers:=True
Range("S1").Select
Columns("AF:AF").Select
Selection.TextToColumns Destination:=Range("AF1"), DataType:=xlDelimited, _
TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
:=Array(1, 1), TrailingMinusNumbers:=True
Range("AF1").Select
Columns("AH:AH").Select
Selection.TextToColumns Destination:=Range("AH1"), DataType:=xlDelimited, _
TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
:=Array(1, 1), TrailingMinusNumbers:=True
Range("AH1").Select
Columns("AO:AO").Select
Selection.TextToColumns Destination:=Range("AO1"), DataType:=xlDelimited, _
TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
:=Array(1, 1), TrailingMinusNumbers:=True
Windows(ss).Activate
ActiveWorkbook.Close SaveChanges:=True
MsgBox "LISTOS FORMATOS_2", vbInformation
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub dark()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim archivos
Dim goku As Excel.Workbook
Dim vegueta As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String
archivos = Dir("D:\AUTOMA_FULL_NEGOCIOS\NETO FIJO\*.xl*")
Do While archivos <> ""
Workbooks.Open "D:\AUTOMA_FULL_NEGOCIOS\NETO FIJO\" & archivos
archivos = Dir
Loop
bills = ActiveWorkbook.Name
tt = bills
Windows(tt).Activate
Range("A3").Select
cool = Selection.End(xlDown).Row
Sheets(1).Range("A3:G" & cool).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Neto Fijo").Select
Range("A3").PasteSpecial xlPasteValues
Range("A3").Select
MsgBox "LISTO FIJO", vbInformation
Windows(tt).Activate
ActiveWorkbook.Close SaveChanges:=False
Call franco
Sheets("Neto Fijo").Select
Range("A1").Select
Columns("A:A").Select
Range("A2").Activate
Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
:=Array(1, 1), TrailingMinusNumbers:=True
Range("A2").Select
Sheets("Neto Movil").Select
Range("A1").Select
Columns("A:A").Select
Range("A2").Activate
Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
:=Array(1, 1), TrailingMinusNumbers:=True
Range("A2").Select
MsgBox "LISTO PRIMERA FASE", vbInformation
Call brie
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub django()
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("LIQUIDADOR").Select
Range("B3").Select
fav = Selection.End(xlDown).Row - 4
Range("D4").Select
For h = 1 To fav
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-3],HC!C1:C3,3,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-4],'DESARROLLO+PROYECTOS'!C3:C9,7,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=Meses!R1C8"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-6],'Meta General'!C[-7]:C[4],5,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(IF(AND(RC[44]>=5,RC[44]<14,RC[1]/RC[-1]<0.6),RC[-1]*70%,IF(AND(RC[44]>=15,RC[1]/RC[-1]<0.6),RC[-1]*50%,RC[-1])),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
    "=IF(RC[-7]=""CONSULTOR"",SUMIFS('Base Fijos'!C[7],'Base Fijos'!C[-4],LIQUIDADOR!RC[-8]),IF(RC[-7]=""ESPECIALISTA"",SUMIFS('Base Fijos'!C[7],'Base Fijos'!C[21],LIQUIDADOR!RC[-8]),IF(RC[-7]=""COORDINADOR"",SUMIFS('Base Fijos'!C[7],'Base Fijos'!C[21],LIQUIDADOR!RC[-8])+SUMIFS('Base Fijos'!C[7],'Base Fijos'!C[35],LIQUIDADOR!RC[-8]),IF(RC[-7]=""JEFE"",SUMIFS('Base Fijos'!C[7],'Base Fijos'!C[19],LIQUIDADOR!RC[-8]),IF(RC[-7]=""JEFECAN"",SUMIFS('Base Fijos'!C[7],'Base Fijos'!C[38],LIQUIDADOR!RC[-8]),IF(RC[-7]=""JEFETMK"",SUMIFS('Base Fijos'!C[7],'Base Fijos'!C[32],LIQUIDADOR!RC[-8]),IF(RC[-7]=""GERENTE"",SUMIFS('Base Fijos'!C[7],'Base Fijos'!C[34],LIQUIDADOR!RC[-8])+SUMIFS('Base Fijos'!C[7],'Base Fijos'!C[32],LIQUIDADOR!RC[-8]),IF(RC[-7]=""DIRECTOR"",SUM('Base Fijos'!C[7],),0))))))))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(RC[-1]/RC[-2],0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
    "=IF(RC[-4]=RC[-3],IF(RC[-9]=""CONSULTOR"",SUMIFS('Base Fijos'!C[15],'Base Fijos'!C[-6],LIQUIDADOR!RC[-10]),IF(RC[-9]=""ESPECIALISTA"",SUMIFS('Base Fijos'!C[15],'Base Fijos'!C[19],LIQUIDADOR!RC[-10]),IF(RC[-9]=""COORDINADOR"",SUMIFS('Base Fijos'!C[15],'Base Fijos'!C[19],LIQUIDADOR!RC[-10])+SUMIFS('Base Fijos'!C[15],'Base Fijos'!C[33],LIQUIDADOR!RC[-10]),IF(RC[-9]=""JEFE"",SUMIFS('Base Fijos'!C[15],'Base Fijos'!C[17],LIQUIDADOR!RC[-10]),IF(RC[-9]=""JEFECAN"",SUMIFS('Base Fijos'!C[15],'Base Fijos'!C[36],LIQUIDADOR!RC[-10]),IF(RC[-9]=""JEFETMK"",SUMIFS('Base Fijos'!C[15],'Base Fijos'!C[30],LIQUIDADOR!RC[-10]),IF(RC[-9]=""GERENTE"",SUMIFS('Base Fijos'!C[15],'Base Fijos'!C[32],LIQUIDADOR!RC[-10])+IF(RC[-9]=""GERENTE"",SUMIFS('Base Fijos'!C[15],'Base Fijos'!C[30],LIQUIDADOR!RC[-10]),0),IF(RC[-9]=""DIRECTOR"",SUM('Base Fijos'!C[5],),0)))))))),RC[-2])"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(SUM((IF(RC[-2]>=0.8,RC[-2]+((RC[-1]-RC[-3])/RC[-4]),RC[-3]/RC[-4])),RC[21]),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
    "=IFERROR(IF(RC[-11]=""CONSULTOR"",IF(RC[-1]<60%,0%,IF(VLOOKUP(RC[-1],FX!R3C10:R30C12,3,TRUE)=""Lineal1"",RC[-1],IF(VLOOKUP(RC[-1],FX!R3C10:R30C12,3,TRUE)=""Lineal+5"",(RC[-1]+5%),IF(VLOOKUP(RC[-1],FX!R3C10:R30C12,3,TRUE)=""Lineal+10"",(RC[-1]+10%),IF(VLOOKUP(RC[-1],FX!R3C10:R30C12,3,TRUE)=""Lineal+15"",(RC[-1]+15%),IF(VLOOKUP(RC[-1],FX!R3C10:R30C12,3,TRUE)=""Lineal+20"",(RC[-1]+20%),IF(VLOOKUP(RC[-1],FX!R3C10:R30C12,3,TRUE)=""Lineal+25"",(RC[-1]+25%),IF(VLOOKUP(RC[-1],FX!R3C10:R30C12,3,TRUE)=""Lineal2"",RC[-1],VLOOKUP(RC[-1],FX!R3C10:R30C12,3,TRUE))))))))),IF(VLOOKUP(RC[-1],FX!R3C15:R11C17,3,TRUE)=""Lineal1"",RC[-1],IF(VLOOKUP(RC[-1],FX!R3C15:R11C17,3,TRUE)=""Lineal"",RC[-1],VLOOKUP(RC[-1],FX!R3C15:R11C17,3,TRUE)))),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF((IF(AND(RC[-2]>=100%,RC[9]>100%),IF(RC[-1]>RC[-2],RC[-1],RC[-2]),RC[-1]))>=400%,400%,IF(AND(RC[-2]>=100%,RC[9]>100%),IF(RC[-1]>RC[-2],RC[-1],RC[-2]),RC[-1]))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-13],FX!R3C2:R11C6,2,0)*RC[-1]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]*RC[-11]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-10]=RC[-9],VLOOKUP(RC[-7],FX!R40C3:R49C5,3,TRUE)*VLOOKUP(RC[-15],FX!R2C2:R11C6,2,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-17],'Meta General'!C[-18]:C[-7],9,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(IF(AND(RC[33]>=5,RC[33]<15,RC[1]/RC[-1]<60%),RC[-1]*70%,IF(AND(RC[33]>=15,RC[1]/RC[-1]<60%),RC[-1]*50%,RC[-1])),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC[-18]=""CONSULTOR"",RC[-18]=""ESPECIALISTA""),SUMIFS('Base Móvil'!C[2],'Base Móvil'!C[-2],LIQUIDADOR!RC[-19],'Base Móvil'!C[10],""Ok Fechas"",'Base Móvil'!C[14],""a""),IF(OR(RC[-18]=""COORDINADOR"",RC[-18]=""COORCAN""),SUMIFS('Base Móvil'!C[2],'Base Móvil'!C[-4],LIQUIDADOR!RC[-19],'Base Móvil'!C[10],""Ok Fechas"",'Base Móvil'!C[14],""a""),IF(RC[-18]=""JEFE"",SUMIFS('Base Móvil'!C[2]," & _
        "'Base Móvil'!C[-6],LIQUIDADOR!RC[-19],'Base Móvil'!C[10],""Ok Fechas"",'Base Móvil'!C[14],""a""),IF(RC[-18]=""JEFECAN"",SUMIFS('Base Móvil'!C[2],'Base Móvil'!C[13],LIQUIDADOR!RC[-19],'Base Móvil'!C[10],""Ok Fechas"",'Base Móvil'!C[14],""a""),IF(RC[-18]=""GERENTE"",SUMIFS('Base Móvil'!C[2],'Base Móvil'!C[-8],LIQUIDADOR!RC[-19],'Base Móvil'!C[10],""Ok Fechas"",'Base Móvil'!C[14],""a""),IF(RC[-18]=""DIRECTOR"",SUMIFS('Base Móvil'!C[2],'Base Móvil'!C[10],""Ok Fechas"",'Base Móvil'!C[14],""a""),IF(RC[-18]=""JEFETMK"",SUMIFS('Base Móvil'!C[2],'Base Móvil'!C[-8],LIQUIDADOR!RC[-19],'Base Móvil'!C[10],""Ok Fechas"",'Base Móvil'!C[14],""a""))))))))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(RC[-1]/RC[-2],0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-4]=RC[-3],IF(OR(RC[-20]=""CONSULTOR"",RC[-20]=""ESPECIALISTA""),SUMIFS('Base Móvil'!C[4],'Base Móvil'!C[-4],LIQUIDADOR!RC[-21],'Base Móvil'!C[8],""Ok Fechas"",'Base Móvil'!C[12],""a""),IF(OR(RC[-20]=""COORDINADOR"",RC[-20]=""COORCAN""),SUMIFS('Base Móvil'!C[4],'Base Móvil'!C[-6],LIQUIDADOR!RC[-21],'Base Móvil'!C[8],""Ok Fechas"",'Base Móvil'!C[12],""a""),IF(RC[-20]=""JEFE"",SUMIFS('Base Móvil'!C[4],'Base Móvil'!C[-8],LIQUIDADOR!RC[-21],'Base Móvil'!C[8],""Ok Fechas"",'Base Móvil'!C[12],""a""),IF(RC[-20]=""JEFECAN"",SUMIFS('Base Móvil'!C[4],'Base Móvil'!C[11],LIQUIDADOR!RC[-21],'Base Móvil'!C[8],""Ok Fechas"",'Base Móvil'!C[12],""a""),IF(RC[-20]=""GERENTE"",SUMIFS('Base Móvil'!C[4],'Base Móvil'!C[-10],LIQUIDADOR!RC[-21]," & _
        "'Base Móvil'!C[8],""Ok Fechas"",'Base Móvil'!C[12],""a""),IF(RC[-20]=""DIRECTOR"",SUMIFS('Base Móvil'!C[4],'Base Móvil'!C[8],""Ok Fechas"",'Base Móvil'!C[12],""a""),IF(RC[-20]=""JEFETMK"",SUMIFS('Base Móvil'!C[4],'Base Móvil'!C[-10],LIQUIDADOR!RC[-21],'Base Móvil'!C[8],""Ok Fechas"",'Base Móvil'!C[12],""a"")))))))),RC[-2])"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(SUM((IF(RC[-2]>=0.8,RC[-2]+((RC[-1]-RC[-3])/RC[-4]),RC[-3]/RC[-4]))),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(IF(RC[-22]=""CONSULTOR"",IF(RC[-1]<60%,0%,IF(VLOOKUP(RC[-1],FX!R3C10:R30C12,3,TRUE)=""Lineal1"",RC[-1],IF(VLOOKUP(RC[-1],FX!R3C10:R30C12,3,TRUE)=""Lineal+5"",(RC[-1]+5%),IF(VLOOKUP(RC[-1],FX!R3C10:R30C12,3,TRUE)=""Lineal+10"",(RC[-1]+10%),IF(VLOOKUP(RC[-1],FX!R3C10:R30C12,3,TRUE)=""Lineal+15"",(RC[-1]+15%),IF(VLOOKUP(RC[-1],FX!R3C10:R30C12,3,TRUE)=""Lineal+20"",(RC[-1]+20%),IF(VLOOKUP(RC[-1],FX!R3C10:R30C12,3,TRUE)=""Lineal+25"",(RC[-1]+25%),IF(VLOOKUP(RC[-1],FX!R3C10:R30C12,3,TRUE)=""Lineal2"",RC[-1],VLOOKUP(RC[-1],FX!R3C10:R30C12,3,TRUE))))))))),IF(VLOOKUP(RC[-1],FX!R3C15:R11C17,3,TRUE)=""Lineal1"",RC[-1],IF(VLOOKUP(RC[-1],FX!R3C15:R11C17,3,TRUE)=""Lineal"",RC[-1],VLOOKUP(RC[-1],FX!R3C15:R11C17,3,TRUE)))),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF((IF(AND(RC[-2]>=100%,RC[-13]>100%),IF(RC[-1]>RC[-2],RC[-1],RC[-2]),RC[-1]))>=400%,400%,(IF(AND(RC[-2]>=100%,RC[-13]>100%),IF(RC[-1]>RC[-2],RC[-1],RC[-2]),RC[-1])))+RC[39]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-24],FX!R3C2:R11C6,3,0)*RC[-1]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]*RC[-22]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-10]=RC[-9],VLOOKUP(RC[-7],FX!R40C3:R49C5,3,TRUE)*VLOOKUP(RC[-26],FX!R2C2:R11C6,3,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-28],'Meta General'!C[-29]:C[-18],4,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(IF(AND(RC[22]>=5,RC[22]<15,RC[1]/RC[-1]<60%),RC[-1]*70%,IF(AND(RC[22]>=15,RC[1]/RC[-1]<60%),RC[-1]*50%,RC[-1])),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=SUM(IF(RC[-29]=""CONSULTOR"",SUMIFS('Base Fijos'!C[-15],'Base Fijos'!C[-26],LIQUIDADOR!RC[-30],'Base Fijos'!C[-22],""RETO""),IF(RC[-29]=""ESPECIALISTA"",SUMIFS('Base Fijos'!C[-15],'Base Fijos'!C[-1],LIQUIDADOR!RC[-30],'Base Fijos'!C[-22],""RETO""),IF(RC[-29]=""COORDINADOR"",SUM(SUMIFS('Base Fijos'!C[-15],'Base Fijos'!C[-1],LIQUIDADOR!RC[-30],'Base Fijos'!C[-22],""RETO""),SUMIFS('Base Fijos'!C[-15],'Base Fijos'!C[13],LIQUIDADOR!RC[-30],'Base Fijos'!C[-22],""RETO"")),IF(RC[-29]=""JEFE"",SUMIFS('Base Fijos'!C[-15],'Base Fijos'!C[-3],LIQUIDADOR!RC[-30],'Base Fijos'!C[-22],""RETO""),IF(RC[-29]=""JEFECAN"",SUMIFS('Base Fijos'!C[-15],'Base Fijos'!C[16],LIQUIDADOR!RC[-30],'Base Fijos'!C[-22],""RETO""),IF(RC[-29]=""JEFETMK"",SUMIFS('Base Fijos'!C[-15]," & _
        "'Base Fijos'!C[10],LIQUIDADOR!RC[-30],'Base Fijos'!C[-22],""RETO""),IF(RC[-29]=""GERENTE"",SUM(SUMIFS('Base Fijos'!C[-15],'Base Fijos'!C[12],LIQUIDADOR!RC[-30],'Base Fijos'!C[-22],""RETO""),SUMIFS('Base Fijos'!C[-15],'Base Fijos'!C[10],LIQUIDADOR!RC[-30],'Base Fijos'!C[-22],""RETO"")),IF(RC[-29]=""DIRECTOR"",SUMIFS('Base Fijos'!C[-15],'Base Fijos'!C[-22],""RETO""),0)))))))),IF(OR(RC[-29]=""CONSULTOR"",RC[-29]=""ESPECIALISTA""),SUMIFS('Base Móvil'!C[-9],'Base Móvil'!C[-13],LIQUIDADOR!RC[-30],'Base Móvil'!C[-1],""Ok Fechas"",'Base Móvil'!C[3],""a"",'Base Móvil'!C[5],""RETO ESTRATEGICO""),IF(OR(RC[-29]=""COORDINADOR"",RC[-29]=""COORCAN""),SUMIFS('Base Móvil'!C[-9],'Base Móvil'!C[-15],LIQUIDADOR!RC[-30],'Base Móvil'!C[-1],""Ok Fechas""," & _
        "'Base Móvil'!C[3],""a"",'Base Móvil'!C[5],""RETO ESTRATEGICO""),IF(OR(RC[-29]=""JEFE"",RC[-29]=""JEFETMK""),SUMIFS('Base Móvil'!C[-9],'Base Móvil'!C[-17],LIQUIDADOR!RC[-30],'Base Móvil'!C[-1],""Ok Fechas"",'Base Móvil'!C[3],""a"",'Base Móvil'!C[5],""RETO ESTRATEGICO""),IF(RC[-29]=""JEFECAN"",SUMIFS('Base Móvil'!C[-9],'Base Móvil'!C[2],LIQUIDADOR!RC[-30],'Base Móvil'!C[-1],""Ok Fechas"",'Base Móvil'!C[3],""a"",'Base Móvil'!C[5],""RETO ESTRATEGICO""),IF(RC[-29]=""GERENTE"",SUMIFS('Base Móvil'!C[-9],'Base Móvil'!C[-19],LIQUIDADOR!RC[-30],'Base Móvil'!C[-1],""Ok Fechas"",'Base Móvil'!C[3],""a"",'Base Móvil'!C[5],""RETO ESTRATEGICO""),IF(RC[-29]=""DIRECTOR"",SUMIFS('Base Móvil'!C[-9],'Base Móvil'!C[-1],""Ok Fechas"",'Base Móvil'!C[3],""a""," & _
        "'Base Móvil'!C[5],""RETO ESTRATEGICO""),0)))))))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(RC[-1]/RC[-2],0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(IF(AND(RC[-1]>=1,RC[-31]=""DIRECTOR""),VLOOKUP(RC[-23],FX!R30C3:R36C5,3,1),0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(IF(RC[-32]=""CONSULTOR"",IF(RC[-2]<60%,0%,IF(VLOOKUP(RC[-2],FX!R3C10:R30C12,3,TRUE)=""Lineal1"",RC[-2],IF(VLOOKUP(RC[-2],FX!R3C10:R30C12,3,TRUE)=""Lineal+5"",(RC[-2]+5%),IF(VLOOKUP(RC[-2],FX!R3C10:R30C12,3,TRUE)=""Lineal+10"",(RC[-2]+10%),IF(VLOOKUP(RC[-2],FX!R3C10:R30C12,3,TRUE)=""Lineal+15"",(RC[-2]+15%),IF(VLOOKUP(RC[-2],FX!R3C10:R30C12,3,TRUE)=""Lineal+20"",(RC[-2]+20%),IF(VLOOKUP(RC[-2],FX!R3C10:R30C12,3,TRUE)=""Lineal+25"",(RC[-2]+25%),IF(VLOOKUP(RC[-2],FX!R3C10:R30C12,3,TRUE)=""Lineal2"",RC[-2],VLOOKUP(RC[-2],FX!R3C10:R30C12,3,TRUE))))))))),IF(VLOOKUP(RC[-2],FX!R3C15:R11C17,3,TRUE)=""Lineal1"",RC[-2],IF(VLOOKUP(RC[-2],FX!R3C15:R11C17,3,TRUE)=""Lineal"",RC[-2],VLOOKUP(RC[-2],FX!R3C15:R11C17,3,TRUE)))),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-33],FX!R3C2:R11C6,4,0)*RC[-1]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]*RC[-31]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-8]=RC[-7],VLOOKUP(RC[-5],FX!R40C3:R49C5,3,TRUE)*VLOOKUP(RC[-35],FX!R2C2:R11C6,4,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-37],'Meta General'!C[-38]:C[-27],12,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(IF(AND(RC[13]>=5,RC[13]<15,RC[1]/RC[-1]<60%),RC[-1]*70%,IF(AND(RC[13]>=15,RC[1]/RC[-1]<60%),RC[-1]*50%,RC[-1])),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-39],'Neto Fijo'!C1:C7,7,0),0)+IFERROR(VLOOKUP(RC[-39],'Neto Movil'!C1:C7,7,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(RC[-1]/RC[-2],0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-1]<5%,0,IF(VLOOKUP(RC[-1],FX!R64C3:R77C5,3,TRUE)=""Lineal"",RC[-1],VLOOKUP(RC[-1],FX!R64C3:R77C5,3,TRUE)))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-41],FX!R3C2:R11C6,5,0)*RC[-1]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]*RC[-39]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-7]=RC[-6],IFERROR(VLOOKUP(RC[-4],FX!R40C3:R49C5,3,TRUE)*VLOOKUP(RC[-43],FX!R2C2:R11C6,5,0),0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-31],RC[-20],RC[-11],RC[-3])"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]*RC[-42]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-31],RC[-20],RC[-11],RC[-3])"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=SUM(RC[-2]:RC[-1])+RC[8]+VLOOKUP(RC[-48],'DESARROLLO+PROYECTOS'!C3:C34,32,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(AND(RC[5]>0,RC[-1]>0,RC[-1]<RC[5]),0,IF(AND(RC[-1]>0,RC[-1]>0,RC[-1]>RC[5]),RC[-1]-RC[5],RC[-1]))"
    ActiveCell.Offset(0, 2).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-50]=""CONSULTOR"",IFERROR(VLOOKUP(RC[-51],Vacaciones!C[-52]:C[-46],6,0),0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = ""
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = ""
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-54],Garantizados!C[-53]:C[-50],3,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(AVERAGE(RC[-46],RC[-35]),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(IF(AND(RC[-1]>=60%,RC[-1]<=69.99%),(SUMIFS('Base Móvil'!C[-18],'Base Móvil'!C[-17],RC[-56]))*5%,IF(AND(RC[-1]>=70%,RC[-1]<=89.99%),(SUMIFS('Base Móvil'!C[-18],'Base Móvil'!C[-17],RC[-56]))*10%,IF(RC[-1]>=90%,(SUMIFS('Base Móvil'!C[-18],'Base Móvil'!C[-17],RC[-56]))*20%,0))),0)"
    ActiveCell.Offset(0, 2).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=AVERAGE(RC[-49],RC[-38],RC[-27],RC[-18])"
    ActiveCell.Offset(0, 2).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-60],'Meta General'!C1:C13,13,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC[-60]=""CONSULTOR"",RC[-60]=""ESPECIALISTA""),SUMIFS('Base Móvil'!C[-40],'Base Móvil'!C[-44],LIQUIDADOR!RC[-61],'Base Móvil'!C[-32],""Ok Fechas"",'Base Móvil'!C[-28],""a"",'Base Móvil'!C[-21],""CONVERGENCIA""),IF(OR(RC[-60]=""COORDINADOR"",RC[-60]=""COORCAN""),SUMIFS('Base Móvil'!C[-40],'Base Móvil'!C[-46],LIQUIDADOR!RC[-61],'Base Móvil'!C[-32],""Ok Fechas"",'Base Móvil'!C[-28],""a"",'Base Móvil'!C[-21],""CONVERGENCIA""),IF(RC[-60]=""JEFE"",SUMIFS('Base Móvil'!C[-40],'Base Móvil'!C[-48],LIQUIDADOR!RC[-61],'Base Móvil'!C[-32],""Ok Fechas"",'Base Móvil'!C[-28],""a"",'Base Móvil'!C[-21],""CONVERGENCIA""),IF(RC[-60]=""JEFECAN"",SUMIFS('Base Móvil'!C[-40],'Base Móvil'!C[-29],LIQUIDADOR!RC[-61],'Base Móvil'!C[-32],""Ok Fechas""," & _
        "'Base Móvil'!C[-28],""a"",'Base Móvil'!C[-21],""CONVERGENCIA""),IF(RC[-60]=""GERENTE"",SUMIFS('Base Móvil'!C[-40],'Base Móvil'!C[-50],LIQUIDADOR!RC[-61],'Base Móvil'!C[-32],""Ok Fechas"",'Base Móvil'!C[-28],""a"",'Base Móvil'!C[-21],""CONVERGENCIA""),IF(RC[-60]=""DIRECTOR"",SUMIFS('Base Móvil'!C[-40],'Base Móvil'!C[-32],""Ok Fechas"",'Base Móvil'!C[-28],""a"",'Base Móvil'!C[-21],""CONVERGENCIA""),IF(RC[-60]=""JEFETMK"",SUMIFS('Base Móvil'!C[-40],'Base Móvil'!C[-50],LIQUIDADOR!RC[-61],'Base Móvil'!C[-32],""Ok Fechas"",'Base Móvil'!C[-28],""a"",'Base Móvil'!C[-21],""CONVERGENCIA""))))))))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(RC[-1]/RC[-2],0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-43]>=60%,IF(AND(RC[-1]>=80%,RC[-1]<100%),5%,IF(RC[-1]>=100%,10%,0)),0%)"
    ActiveCell.Offset(1, 0).Range("A1").Select
    Selection.End(xlToLeft).Select
Next h
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic
MsgBox "LISTO PRIMERA FASE", vbInformation
End Sub
Sub tosh()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Call marcolis
Dim archivos
Dim goku As Excel.Workbook
Dim vegueta As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String
archivos = Dir("D:\AUTOMA_FULL_NEGOCIOS\LIQ_FINAL\*.xlsm")
Do While archivos <> ""
Workbooks.Open "D:\AUTOMA_FULL_NEGOCIOS\LIQ_FINAL\" & archivos
archivos = Dir
Loop
bills = ActiveWorkbook.Name
tt = bills
Windows(tt).Activate
Sheets("Meses").Select
RATA = Range("F1").Value
cr7 = Range("G1").Value
Range("A1").Select
Sheets("HC").Visible = False
Sheets("Neto Fijo").Select
Columns("C:C").Select
Selection.EntireColumn.Hidden = True
Sheets("Neto Fijo").Protect Password:="TOP"
Sheets("Neto Movil").Select
Columns("C:C").Select
Selection.EntireColumn.Hidden = True
Sheets("Neto Movil").Protect Password:="TOP"
Sheets("SIN_TURNOS").Protect Password:="TOP"
Sheets("Meses").Protect Password:="TOP"
Sheets("LIQUIDADOR").Select
Columns("E:E").Select
Selection.EntireColumn.Hidden = True
Sheets("LIQUIDADOR").Protect Password:="TOP"
Sheets("FX").Protect Password:="TOP"
Sheets("Base Fijos").Protect Password:="TOP"
Sheets("Base Móvil").Protect Password:="TOP"
Sheets("Garantizados").Select
Columns("D:D").Select
Selection.EntireColumn.Hidden = True
Sheets("Garantizados").Protect Password:="TOP"
Sheets("Vacaciones").Select
Columns("B:B").Select
Selection.EntireColumn.Hidden = True
Sheets("Vacaciones").Protect Password:="TOP"
Sheets("Meta General").Select
Columns("B:B").Select
Selection.EntireColumn.Hidden = True
Sheets("Meta General").Protect Password:="TOP"
ActiveWorkbook.Protect ("TOP")
Sheets("LIQUIDADOR").Select
RUTA = "D:\AUTOMA_FULL_NEGOCIOS\LIQ_FINAL\"
GENERAR_ARCHIVO_EXCEL = RUTA & "Liq_" & RATA & cr7 & "_Negocios_VAL" & ".xlsx"
ActiveWorkbook.SaveAs Filename:= _
GENERAR_ARCHIVO_EXCEL, FileFormat:= _
xlOpenXMLWorkbook, Password:="3392", CreateBackup:=False
simpson = ActiveWorkbook.Name
hh = simpson
Windows(hh).Activate
ActiveWorkbook.Close SaveChanges:=True
Call arm
Kill ("D:\AUTOMA_FULL_NEGOCIOS\LIQ_FINAL\*.xlsm")
Application.ScreenUpdating = True
Application.DisplayAlerts = True
MsgBox "proceso FINAL FINAL completado CAPO, puede continuar con el siguiente paso", vbInformation
MsgBox "Por favor espere DOS minutos mientras se enfría el procesador, de lo contrario la memoria se desboradará", vbInformation
End Sub
Sub arm()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim archivos
Dim goku As Excel.Workbook
Dim vegueta As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String
archivos = Dir("D:\AUTOMA_FULL_NEGOCIOS\LIQ_FINAL\*.xlsm")
Do While archivos <> ""
Workbooks.Open "D:\AUTOMA_FULL_NEGOCIOS\LIQ_FINAL\" & archivos
archivos = Dir
Loop
bills = ActiveWorkbook.Name
tt = bills
Windows(tt).Activate
Sheets("Meses").Select
RATA = Range("F1").Value
cr7 = Range("G1").Value
Range("A1").Select
Sheets("LIQUIDADOR").Select
RUTA = "D:\AUTOMA_FULL_NEGOCIOS\LIQ_FINAL\"
GENERAR_ARCHIVO_EXCEL = RUTA & "Liq_" & RATA & cr7 & "_Negocios" & ".xlsx"
ActiveWorkbook.SaveAs Filename:= _
GENERAR_ARCHIVO_EXCEL, FileFormat:= _
xlOpenXMLWorkbook, CreateBackup:=False
simpson = ActiveWorkbook.Name
hh = simpson
Windows(hh).Activate
ActiveWorkbook.Close SaveChanges:=True
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub msi()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("SIN_TURNOS").Select
Columns("F:F").Select
Selection.TextToColumns Destination:=Range("F1"), DataType:=xlDelimited, _
TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
:=Array(1, 1), TrailingMinusNumbers:=True
Range("F1").Select
Sheets("Base Móvil").Select
Range("A1").Select
fav = Selection.End(xlDown).Row
Set rangodatos = Sheets("Base Móvil").Range("A1:AR" & fav)
rangodatos.AutoFilter Field:=3, Criteria1:="<>"
rangodatos.AutoFilter Field:=20, Criteria1:=Array("ALTAS", "PREPOS"), Operator:=xlFilterValues
Range("AQ1").Select
ActiveCell.FormulaR1C1 = "=IF(IFERROR(VLOOKUP(RC[-40],SIN_TURNOS!C6,1,FALSE),"""")=RC[-40],""VENTAS SIN TURNO CAV"","""")"
ActiveCell.Copy
Sheets("Base Móvil").Range("AQ1:AQ" & fav).SpecialCells(xlCellTypeVisible).Select
Selection.PasteSpecial xlPasteFormulas
Application.CutCopyMode = False
ActiveSheet.ShowAllData
Range("AQ1") = "MARCA NO PAGO TURNO CAV"
Set rangodatos = Sheets("Base Móvil").Range("A1:AR" & fav)
rangodatos.AutoFilter Field:=43, Criteria1:="<>"
Range("AR1").Select
ActiveCell.FormulaR1C1 = "=RC[-21]"
ActiveCell.Copy
Sheets("Base Móvil").Range("AR1:AR" & fav).SpecialCells(xlCellTypeVisible).Select
Selection.PasteSpecial xlPasteFormulas
Application.CutCopyMode = False
Range("A1").Select
Range("AR1") = "CFM VENTAS SIN TURNO"
ActiveSheet.ShowAllData
ActiveCell.Columns("AR:AR").EntireColumn.Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False
Range("A1").Select
Set rangodatos = Sheets("Base Móvil").Range("A1:AR" & fav)
rangodatos.AutoFilter Field:=43, Criteria1:="<>"
Range("W1").Select
Range("W1") = "0"
ActiveCell.Copy
Sheets("Base Móvil").Range("W1:W" & fav).SpecialCells(xlCellTypeVisible).Select
Selection.PasteSpecial xlPasteValues
Application.CutCopyMode = False
Range("A1").Select
Range("W1") = "CFM_ACTUAL / CARGA"
ActiveSheet.ShowAllData
MsgBox "LISTO SEGUNDA FASE", vbInformation
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub oblak()
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim archivos, motorola As String
Dim goku As Excel.Workbook
Dim vegueta As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String
Workbooks.Add
bills = ActiveWorkbook.Name
tt = bills
Windows(tt).Activate
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("A1").Select
cool = Selection.End(xlDown).Row
Sheets("Base Fijos").Range("A1:BB" & cool).Copy
Windows(tt).Activate
Range("A1").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Columns(1).EntireColumn.Insert
Range("A1") = "ID"
Range("A1").Select
MsgBox "LISTO EL PEGUE", vbInformation
Windows(tt).Activate
Range("B1").Select
coool = Selection.End(xlDown).Row
Sheets(1).Range("G2:G" & coool).Copy
Range("A2").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Range("A2:BB" & cool).Copy
Windows(tt).Activate
Range("B1").Select
cloool = Selection.End(xlDown).Row
Sheets(1).Range("B" & cloool + 1).Select
ActiveSheet.Paste
Application.CutCopyMode = False
Range("A1").Select
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Range("AC2:AC" & cool).Copy
Windows(tt).Activate
Range("A1").Select
Sheets(1).Range("A" & cloool + 1).Select
Selection.PasteSpecial xlPasteValues
Application.CutCopyMode = False
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Range("A2:BB" & cool).Copy
Windows(tt).Activate
Range("B1").Select
clloool = Selection.End(xlDown).Row
Sheets(1).Range("B" & clloool + 1).Select
ActiveSheet.Paste
Application.CutCopyMode = False
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Range("AE2:AE" & cool).Copy
Windows(tt).Activate
Range("A1").Select
Sheets(1).Range("A" & clloool + 1).Select
Selection.PasteSpecial xlPasteValues
Application.CutCopyMode = False
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Range("A2:BB" & cool).Copy
Windows(tt).Activate
Range("B1").Select
cllloool = Selection.End(xlDown).Row
Sheets(1).Range("B" & cllloool + 1).Select
ActiveSheet.Paste
Application.CutCopyMode = False
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Range("AP2:AP" & cool).Copy
Windows(tt).Activate
Range("A1").Select
Sheets(1).Range("A" & cllloool + 1).Select
Selection.PasteSpecial xlPasteValues
Application.CutCopyMode = False
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Range("A2:BB" & cool).Copy
Windows(tt).Activate
Range("B1").Select
clllloool = Selection.End(xlDown).Row
Sheets(1).Range("B" & clllloool + 1).Select
ActiveSheet.Paste
Application.CutCopyMode = False
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Range("AS2:AS" & cool).Copy
Windows(tt).Activate
Range("A1").Select
Sheets(1).Range("A" & clllloool + 1).Select
Selection.PasteSpecial xlPasteValues
Application.CutCopyMode = False
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Range("A2:BB" & cool).Copy
Windows(tt).Activate
Range("B1").Select
cllllooool = Selection.End(xlDown).Row
Sheets(1).Range("B" & cllllooool + 1).Select
ActiveSheet.Paste
Application.CutCopyMode = False
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Range("AV2:AV" & cool).Copy
Windows(tt).Activate
Range("A1").Select
Sheets(1).Range("A" & cllllooool + 1).Select
Selection.PasteSpecial xlPasteValues
Application.CutCopyMode = False
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Range("A2:BB" & cool).Copy
Windows(tt).Activate
Range("B1").Select
clllllooool = Selection.End(xlDown).Row
Sheets(1).Range("B" & clllllooool + 1).Select
ActiveSheet.Paste
Application.CutCopyMode = False
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Range("AX2:AX" & cool).Copy
Windows(tt).Activate
Range("A1").Select
Sheets(1).Range("A" & clllllooool + 1).Select
Selection.PasteSpecial xlPasteValues
Application.CutCopyMode = False
Columns("V:V").Delete
Range("A1").Select
With Selection.Interior
.Pattern = xlSolid
.PatternColorIndex = xlAutomatic
.ThemeColor = xlThemeColorLight1
.TintAndShade = 0
.PatternTintAndShade = 0
End With
With Selection.Font
.ThemeColor = xlThemeColorDark1
.TintAndShade = 0
End With
Selection.Font.Bold = True
With Selection
.HorizontalAlignment = xlGeneral
.VerticalAlignment = xlCenter
.WrapText = False
.Orientation = 0
.AddIndent = False
.IndentLevel = 0
.ShrinkToFit = False
.ReadingOrder = xlContext
.MergeCells = False
End With
MsgBox "LISTO CÁLCULOS", vbInformation
Cells.Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False
Sheets(1).Select
Range("A1").Select
Range("A1").Sort Key1:=Range("A2"), Order1:=xlAscending, Header:=xlYes
MsgBox "LISTO ORDEN", vbInformation
Sheets(1).Select
Set rangodatos = Sheets(1).UsedRange
rangodatos.AutoFilter Field:=1, Criteria1:="<>0", Operator:=xlAnd, Criteria2:="<>"
Worksheets.Add
Sheets("Hoja1").Select
Range("B1").Select
medi = Selection.End(xlDown).Row
Sheets("Hoja1").Range("A1:BC" & medi).Copy
Sheets("Hoja2").Select
Range("A1").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Worksheets.Add
Sheets("Hoja2").Select
Set rangodatos = Sheets("Hoja2").UsedRange
rangodatos.AutoFilter Field:=1, Criteria1:="<>11", Operator:=xlAnd
Sheets("Hoja2").Select
Range("B1").Select
coooolp = Selection.End(xlDown).Row
Sheets("Hoja2").Range("A1:BC" & coooolp).Copy
Sheets("Hoja3").Select
Range("A1").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Range("A1").Select
Sheets("Hoja2").Delete
Sheets("Hoja1").Delete
MsgBox "LISTO SEGUNDA FASE", vbInformation
chang = "D:\AUTOMA_FULL_NEGOCIOS\PARA DETALLES\"
magia = chang & "DETALLES_FIJA" & ".xlsx"
ActiveWorkbook.SaveAs Filename:= _
magia, FileFormat:= _
xlOpenXMLWorkbook, CreateBackup:=False
simpson = ActiveWorkbook.Name
hh = simpson
Windows(hh).Activate
ActiveWorkbook.Close SaveChanges:=True
MsgBox "proceso FINAL FINAL completado CAPO, puede continuar con el siguiente paso", vbInformation
MsgBox "Por favor espere 1 minuto mientras se enfría el procesador, de lo contrario la memoria se desboradará", vbInformation
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub rumsey()
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim archivos, motorola As String
Dim goku As Excel.Workbook
Dim vegueta As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String
Workbooks.Add
bills = ActiveWorkbook.Name
tt = bills
Windows(tt).Activate
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Móvil").Select
Range("A1").Select
cool = Selection.End(xlDown).Row
Sheets("Base Móvil").Range("A1:AR" & cool).Copy
Windows(tt).Activate
Range("A1").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Columns(1).EntireColumn.Insert
Range("A1") = "ID"
Range("A1").Select
MsgBox "LISTO EL PEGUE", vbInformation
Windows(tt).Activate
Range("B1").Select
coool = Selection.End(xlDown).Row
Sheets(1).Range("N2:N" & coool).Copy
Range("A2").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Móvil").Range("A2:AR" & cool).Copy
Windows(tt).Activate
Range("B1").Select
cloool = Selection.End(xlDown).Row
Sheets(1).Range("B" & cloool + 1).Select
ActiveSheet.Paste
Application.CutCopyMode = False
Range("A1").Select
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Móvil").Range("O2:O" & cool).Copy
Windows(tt).Activate
Range("A1").Select
Sheets(1).Range("A" & cloool + 1).Select
Selection.PasteSpecial xlPasteValues
Application.CutCopyMode = False
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Móvil").Range("A2:AR" & cool).Copy
Windows(tt).Activate
Range("B1").Select
clloool = Selection.End(xlDown).Row
Sheets(1).Range("B" & clloool + 1).Select
ActiveSheet.Paste
Application.CutCopyMode = False
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Móvil").Range("Q2:Q" & cool).Copy
Windows(tt).Activate
Range("A1").Select
Sheets(1).Range("A" & clloool + 1).Select
Selection.PasteSpecial xlPasteValues
Application.CutCopyMode = False
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Móvil").Range("A2:AR" & cool).Copy
Windows(tt).Activate
Range("B1").Select
cllloool = Selection.End(xlDown).Row
Sheets(1).Range("B" & cllloool + 1).Select
ActiveSheet.Paste
Application.CutCopyMode = False
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Móvil").Range("AF2:AF" & cool).Copy
Windows(tt).Activate
Range("A1").Select
Sheets(1).Range("A" & cllloool + 1).Select
Selection.PasteSpecial xlPasteValues
Application.CutCopyMode = False
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Móvil").Range("A2:AR" & cool).Copy
Windows(tt).Activate
Range("B1").Select
clllooool = Selection.End(xlDown).Row
Sheets(1).Range("B" & clllooool + 1).Select
ActiveSheet.Paste
Application.CutCopyMode = False
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Móvil").Range("AH2:AH" & cool).Copy
Windows(tt).Activate
Range("A1").Select
Sheets(1).Range("A" & clllooool + 1).Select
Selection.PasteSpecial xlPasteValues
Application.CutCopyMode = False
Cells.Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False
Worksheets.Add
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Móvil").Range("A1:AS" & cool).Copy
Windows(tt).Activate
Sheets("Hoja2").Select
Range("A1").PasteSpecial xlPasteValues
Sheets("Hoja2").Select
Set rangodatos = Sheets("Hoja2").UsedRange
rangodatos.AutoFilter Field:=19, Criteria1:=""
Range("S1").Select
ActiveCell.FormulaR1C1 = "=RC[22]"
Selection.Copy
Range(Selection, Selection.End(xlDown)).Select
ActiveSheet.Paste
Application.CutCopyMode = False
Range("S1").Select
ActiveSheet.ShowAllData
Sheets("Hoja2").Select
Range("A1").Select
clif = Selection.End(xlDown).Row
Sheets("Hoja2").Range("A2:AS" & clif).Copy
Sheets("Hoja1").Select
Range("B1").Select
porci = Selection.End(xlDown).Row
Sheets("Hoja1").Range("B" & porci + 1).Select
ActiveSheet.Paste
Application.CutCopyMode = False
Sheets("Hoja2").Select
Sheets("Hoja2").Range("S2:S" & clif).Copy
Sheets("Hoja1").Select
Sheets("Hoja1").Range("A" & porci + 1).Select
Selection.PasteSpecial xlPasteValues
Application.CutCopyMode = False
MsgBox "LISTO SEGUNDA FASE", vbInformation
Range("A1").Select
With Selection.Interior
.Pattern = xlSolid
.PatternColorIndex = xlAutomatic
.ThemeColor = xlThemeColorLight1
.TintAndShade = 0
.PatternTintAndShade = 0
End With
With Selection.Font
.ThemeColor = xlThemeColorDark1
.TintAndShade = 0
End With
Selection.Font.Bold = True
With Selection
.HorizontalAlignment = xlGeneral
.VerticalAlignment = xlCenter
.WrapText = False
.Orientation = 0
.AddIndent = False
.IndentLevel = 0
.ShrinkToFit = False
.ReadingOrder = xlContext
.MergeCells = False
End With
MsgBox "LISTO CÁLCULOS", vbInformation
Cells.Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False
Sheets("Hoja1").Select
Range("A1").Select
Range("A1").Sort Key1:=Range("A2"), Order1:=xlAscending, Header:=xlYes
MsgBox "LISTO ORDEN", vbInformation
Worksheets.Add
Sheets("Hoja1").Select
Set rangodatos = Sheets("Hoja1").UsedRange
rangodatos.AutoFilter Field:=1, Criteria1:="<>0", Operator:=xlAnd, Criteria2:="<>"
Range("A1").Select
clifo = Selection.End(xlDown).Row
Sheets("Hoja1").Range("A1:AS" & clifo).Copy
Sheets("Hoja3").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Range("A1").Select
Sheets("Hoja2").Delete
Sheets("Hoja1").Delete
MsgBox "LISTO TERCERA FASE", vbInformation
chang = "D:\AUTOMA_FULL_NEGOCIOS\PARA DETALLES\"
magia = chang & "DETALLES_MÓVIL" & ".xlsx"
ActiveWorkbook.SaveAs Filename:= _
magia, FileFormat:= _
xlOpenXMLWorkbook, CreateBackup:=False
simpson = ActiveWorkbook.Name
hh = simpson
Windows(hh).Activate
ActiveWorkbook.Close SaveChanges:=True
MsgBox "proceso FINAL FINAL completado CAPO, puede continuar con el siguiente paso", vbInformation
MsgBox "FELICIDADES CAMPEONA DE CAMPEONAS, MUJER EMPODERADA Y LUCHONA...SE REALIZÓ LA LIQUIDACIÓN COMPLETA EN MENOS DE 45 MINUTOS, EL SISTEMA SE CERRARÁ SOLO POR FAVOR ESPERE", vbInformation
ActiveWorkbook.Close SaveChanges:=True
Application.Quit
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub sivakov()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("LIQUIDADOR").Select
Range("D3").Select
Selection.End(xlDown).Select
ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-3],HC!C1:C3,3,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-4],'DESARROLLO+PROYECTOS'!C3:C9,7,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=Meses!R1C8"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-6],'Meta General'!C[-7]:C[4],5,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(IF(AND(RC[44]>=5,RC[44]<14,RC[1]/RC[-1]<0.6),RC[-1]*70%,IF(AND(RC[44]>=15,RC[1]/RC[-1]<0.6),RC[-1]*50%,RC[-1])),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
    "=IF(RC[-7]=""CONSULTOR"",SUMIFS('Base Fijos'!C[7],'Base Fijos'!C[-4],LIQUIDADOR!RC[-8]),IF(RC[-7]=""ESPECIALISTA"",SUMIFS('Base Fijos'!C[7],'Base Fijos'!C[21],LIQUIDADOR!RC[-8]),IF(RC[-7]=""COORDINADOR"",SUMIFS('Base Fijos'!C[7],'Base Fijos'!C[21],LIQUIDADOR!RC[-8])+SUMIFS('Base Fijos'!C[7],'Base Fijos'!C[35],LIQUIDADOR!RC[-8]),IF(RC[-7]=""JEFE"",SUMIFS('Base Fijos'!C[7],'Base Fijos'!C[19],LIQUIDADOR!RC[-8]),IF(RC[-7]=""JEFECAN"",SUMIFS('Base Fijos'!C[7],'Base Fijos'!C[38],LIQUIDADOR!RC[-8]),IF(RC[-7]=""JEFETMK"",SUMIFS('Base Fijos'!C[7],'Base Fijos'!C[32],LIQUIDADOR!RC[-8]),IF(RC[-7]=""GERENTE"",SUMIFS('Base Fijos'!C[7],'Base Fijos'!C[34],LIQUIDADOR!RC[-8])+SUMIFS('Base Fijos'!C[7],'Base Fijos'!C[32],LIQUIDADOR!RC[-8]),IF(RC[-7]=""DIRECTOR"",SUM('Base Fijos'!C[7],),0))))))))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(RC[-1]/RC[-2],0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
    "=IF(RC[-4]=RC[-3],IF(RC[-9]=""CONSULTOR"",SUMIFS('Base Fijos'!C[15],'Base Fijos'!C[-6],LIQUIDADOR!RC[-10]),IF(RC[-9]=""ESPECIALISTA"",SUMIFS('Base Fijos'!C[15],'Base Fijos'!C[19],LIQUIDADOR!RC[-10]),IF(RC[-9]=""COORDINADOR"",SUMIFS('Base Fijos'!C[15],'Base Fijos'!C[19],LIQUIDADOR!RC[-10])+SUMIFS('Base Fijos'!C[15],'Base Fijos'!C[33],LIQUIDADOR!RC[-10]),IF(RC[-9]=""JEFE"",SUMIFS('Base Fijos'!C[15],'Base Fijos'!C[17],LIQUIDADOR!RC[-10]),IF(RC[-9]=""JEFECAN"",SUMIFS('Base Fijos'!C[15],'Base Fijos'!C[36],LIQUIDADOR!RC[-10]),IF(RC[-9]=""JEFETMK"",SUMIFS('Base Fijos'!C[15],'Base Fijos'!C[30],LIQUIDADOR!RC[-10]),IF(RC[-9]=""GERENTE"",SUMIFS('Base Fijos'!C[15],'Base Fijos'!C[32],LIQUIDADOR!RC[-10])+IF(RC[-9]=""GERENTE"",SUMIFS('Base Fijos'!C[15],'Base Fijos'!C[30],LIQUIDADOR!RC[-10]),0),IF(RC[-9]=""DIRECTOR"",SUM('Base Fijos'!C[15],),0)))))))),RC[-2])"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(SUM((IF(RC[-2]>=0.8,RC[-2]+((RC[-1]-RC[-3])/RC[-4]),RC[-3]/RC[-4])),RC[21]),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF((IF(AND(RC[-2]>=100%,RC[9]>100%),IF(RC[-1]>RC[-2],RC[-1],RC[-2]),RC[-1]))>=400%,400%,IF(AND(RC[-2]>=100%,RC[9]>100%),IF(RC[-1]>RC[-2],RC[-1],RC[-2]),RC[-1]))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-13],FX!R3C2:R11C6,2,0)*RC[-1]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]*RC[-11]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-10]=RC[-9],VLOOKUP(RC[-7],FX!R40C3:R49C5,3,TRUE)*VLOOKUP(RC[-15],FX!R2C2:R11C6,2,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-17],'Meta General'!C[-18]:C[-7],9,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(IF(AND(RC[33]>=5,RC[33]<15,RC[1]/RC[-1]<60%),RC[-1]*70%,IF(AND(RC[33]>=15,RC[1]/RC[-1]<60%),RC[-1]*50%,RC[-1])),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC[-18]=""CONSULTOR"",RC[-18]=""ESPECIALISTA""),SUMIFS('Base Móvil'!C[2],'Base Móvil'!C[-2],LIQUIDADOR!RC[-19],'Base Móvil'!C[10],""Ok Fechas"",'Base Móvil'!C[14],""a""),IF(OR(RC[-18]=""COORDINADOR"",RC[-18]=""COORCAN""),SUMIFS('Base Móvil'!C[2],'Base Móvil'!C[-4],LIQUIDADOR!RC[-19],'Base Móvil'!C[10],""Ok Fechas"",'Base Móvil'!C[14],""a""),IF(RC[-18]=""JEFE"",SUMIFS('Base Móvil'!C[2]," & _
        "'Base Móvil'!C[-6],LIQUIDADOR!RC[-19],'Base Móvil'!C[10],""Ok Fechas"",'Base Móvil'!C[14],""a""),IF(RC[-18]=""JEFECAN"",SUMIFS('Base Móvil'!C[2],'Base Móvil'!C[13],LIQUIDADOR!RC[-19],'Base Móvil'!C[10],""Ok Fechas"",'Base Móvil'!C[14],""a""),IF(RC[-18]=""GERENTE"",SUMIFS('Base Móvil'!C[2],'Base Móvil'!C[-8],LIQUIDADOR!RC[-19],'Base Móvil'!C[10],""Ok Fechas"",'Base Móvil'!C[14],""a""),IF(RC[-18]=""DIRECTOR"",SUMIFS('Base Móvil'!C[2],'Base Móvil'!C[10],""Ok Fechas"",'Base Móvil'!C[14],""a""),IF(RC[-18]=""JEFETMK"",SUMIFS('Base Móvil'!C[2],'Base Móvil'!C[-8],LIQUIDADOR!RC[-19],'Base Móvil'!C[10],""Ok Fechas"",'Base Móvil'!C[14],""a""))))))))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(RC[-1]/RC[-2],0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-4]=RC[-3],IF(OR(RC[-20]=""CONSULTOR"",RC[-20]=""ESPECIALISTA""),SUMIFS('Base Móvil'!C[4],'Base Móvil'!C[-4],LIQUIDADOR!RC[-21],'Base Móvil'!C[8],""Ok Fechas"",'Base Móvil'!C[12],""a""),IF(OR(RC[-20]=""COORDINADOR"",RC[-20]=""COORCAN""),SUMIFS('Base Móvil'!C[4],'Base Móvil'!C[-6],LIQUIDADOR!RC[-21],'Base Móvil'!C[8],""Ok Fechas"",'Base Móvil'!C[12],""a""),IF(RC[-20]=""JEFE"",SUMIFS('Base Móvil'!C[4],'Base Móvil'!C[-8],LIQUIDADOR!RC[-21],'Base Móvil'!C[8],""Ok Fechas"",'Base Móvil'!C[12],""a""),IF(RC[-20]=""JEFECAN"",SUMIFS('Base Móvil'!C[4],'Base Móvil'!C[11],LIQUIDADOR!RC[-21],'Base Móvil'!C[8],""Ok Fechas"",'Base Móvil'!C[12],""a""),IF(RC[-20]=""GERENTE"",SUMIFS('Base Móvil'!C[4],'Base Móvil'!C[-10],LIQUIDADOR!RC[-21]," & _
        "'Base Móvil'!C[8],""Ok Fechas"",'Base Móvil'!C[12],""a""),IF(RC[-20]=""DIRECTOR"",SUMIFS('Base Móvil'!C[4],'Base Móvil'!C[8],""Ok Fechas"",'Base Móvil'!C[12],""a""),IF(RC[-20]=""JEFETMK"",SUMIFS('Base Móvil'!C[4],'Base Móvil'!C[-10],LIQUIDADOR!RC[-21],'Base Móvil'!C[8],""Ok Fechas"",'Base Móvil'!C[12],""a"")))))))),RC[-2])"
    ActiveCell.Offset(0, 1).Range("A1").Select
     ActiveCell.FormulaR1C1 = "=IFERROR(RC[-1]/RC[-4],0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF((IF(AND(RC[-2]>=100%,RC[-13]>100%),IF(RC[-1]>RC[-2],RC[-1],RC[-2]),RC[-1]))>=400%,400%,(IF(AND(RC[-2]>=100%,RC[-13]>100%),IF(RC[-1]>RC[-2],RC[-1],RC[-2]),RC[-1])))+RC[39]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-24],FX!R3C2:R11C6,3,0)*RC[-1]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]*RC[-22]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-10]=RC[-9],VLOOKUP(RC[-7],FX!R40C3:R49C5,3,TRUE)*VLOOKUP(RC[-26],FX!R2C2:R11C6,3,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-28],'Meta General'!C[-29]:C[-18],4,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(IF(AND(RC[22]>=5,RC[22]<15,RC[1]/RC[-1]<60%),RC[-1]*70%,IF(AND(RC[22]>=15,RC[1]/RC[-1]<60%),RC[-1]*50%,RC[-1])),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=SUM(IF(RC[-29]=""CONSULTOR"",SUMIFS('Base Fijos'!C[-15],'Base Fijos'!C[-26],LIQUIDADOR!RC[-30],'Base Fijos'!C[-22],""RETO""),IF(RC[-29]=""ESPECIALISTA"",SUMIFS('Base Fijos'!C[-15],'Base Fijos'!C[-1],LIQUIDADOR!RC[-30],'Base Fijos'!C[-22],""RETO""),IF(RC[-29]=""COORDINADOR"",SUM(SUMIFS('Base Fijos'!C[-15],'Base Fijos'!C[-1],LIQUIDADOR!RC[-30],'Base Fijos'!C[-22],""RETO""),SUMIFS('Base Fijos'!C[-15],'Base Fijos'!C[13],LIQUIDADOR!RC[-30],'Base Fijos'!C[-22],""RETO"")),IF(RC[-29]=""JEFE"",SUMIFS('Base Fijos'!C[-15],'Base Fijos'!C[-3],LIQUIDADOR!RC[-30],'Base Fijos'!C[-22],""RETO""),IF(RC[-29]=""JEFECAN"",SUMIFS('Base Fijos'!C[-15],'Base Fijos'!C[16],LIQUIDADOR!RC[-30],'Base Fijos'!C[-22],""RETO""),IF(RC[-29]=""JEFETMK"",SUMIFS('Base Fijos'!C[-15]," & _
        "'Base Fijos'!C[10],LIQUIDADOR!RC[-30],'Base Fijos'!C[-22],""RETO""),IF(RC[-29]=""GERENTE"",SUM(SUMIFS('Base Fijos'!C[-15],'Base Fijos'!C[12],LIQUIDADOR!RC[-30],'Base Fijos'!C[-22],""RETO""),SUMIFS('Base Fijos'!C[-15],'Base Fijos'!C[10],LIQUIDADOR!RC[-30],'Base Fijos'!C[-22],""RETO"")),IF(RC[-29]=""DIRECTOR"",SUMIFS('Base Fijos'!C[-15],'Base Fijos'!C[-22],""RETO""),0)))))))),IF(OR(RC[-29]=""CONSULTOR"",RC[-29]=""ESPECIALISTA""),SUMIFS('Base Móvil'!C[-9],'Base Móvil'!C[-13],LIQUIDADOR!RC[-30],'Base Móvil'!C[-1],""Ok Fechas"",'Base Móvil'!C[3],""a"",'Base Móvil'!C[5],""RETO ESTRATEGICO""),IF(OR(RC[-29]=""COORDINADOR"",RC[-29]=""COORCAN""),SUMIFS('Base Móvil'!C[-9],'Base Móvil'!C[-15],LIQUIDADOR!RC[-30],'Base Móvil'!C[-1],""Ok Fechas""," & _
        "'Base Móvil'!C[3],""a"",'Base Móvil'!C[5],""RETO ESTRATEGICO""),IF(OR(RC[-29]=""JEFE"",RC[-29]=""JEFETMK""),SUMIFS('Base Móvil'!C[-9],'Base Móvil'!C[-17],LIQUIDADOR!RC[-30],'Base Móvil'!C[-1],""Ok Fechas"",'Base Móvil'!C[3],""a"",'Base Móvil'!C[5],""RETO ESTRATEGICO""),IF(RC[-29]=""JEFECAN"",SUMIFS('Base Móvil'!C[-9],'Base Móvil'!C[2],LIQUIDADOR!RC[-30],'Base Móvil'!C[-1],""Ok Fechas"",'Base Móvil'!C[3],""a"",'Base Móvil'!C[5],""RETO ESTRATEGICO""),IF(RC[-29]=""GERENTE"",SUMIFS('Base Móvil'!C[-9],'Base Móvil'!C[-19],LIQUIDADOR!RC[-30],'Base Móvil'!C[-1],""Ok Fechas"",'Base Móvil'!C[3],""a"",'Base Móvil'!C[5],""RETO ESTRATEGICO""),IF(RC[-29]=""DIRECTOR"",SUMIFS('Base Móvil'!C[-9],'Base Móvil'!C[-1],""Ok Fechas"",'Base Móvil'!C[3],""a""," & _
        "'Base Móvil'!C[5],""RETO ESTRATEGICO""),0)))))))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(RC[-1]/RC[-2],0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(IF(AND(RC[-1]>=1,RC[-31]=""DIRECTOR""),VLOOKUP(RC[-23],FX!R30C3:R36C5,3,1),0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-2]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-33],FX!R3C2:R11C6,4,0)*RC[-1]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]*RC[-31]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-8]=RC[-7],VLOOKUP(RC[-5],FX!R40C3:R49C5,3,TRUE)*VLOOKUP(RC[-35],FX!R2C2:R11C6,4,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-37],'Meta General'!C[-38]:C[-27],12,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(IF(AND(RC[13]>=5,RC[13]<15,RC[1]/RC[-1]<60%),RC[-1]*70%,IF(AND(RC[13]>=15,RC[1]/RC[-1]<60%),RC[-1]*50%,RC[-1])),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-39],'Neto Fijo'!C1:C7,7,0),0)+IFERROR(VLOOKUP(RC[-39],'Neto Movil'!C1:C7,7,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(RC[-1]/RC[-2],0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-1]<5%,0,IF(VLOOKUP(RC[-1],FX!R64C3:R77C5,3,TRUE)=""Lineal"",RC[-1],VLOOKUP(RC[-1],FX!R64C3:R77C5,3,TRUE)))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-41],FX!R3C2:R11C6,5,0)*RC[-1]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]*RC[-39]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-7]=RC[-6],IFERROR(VLOOKUP(RC[-4],FX!R40C3:R49C5,3,TRUE)*VLOOKUP(RC[-43],FX!R2C2:R11C6,5,0),0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-31],RC[-20],RC[-11],RC[-3])"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]*RC[-42]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-31],RC[-20],RC[-11],RC[-3])"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=SUM(RC[-2]:RC[-1])+RC[8]+VLOOKUP(RC[-48],'DESARROLLO+PROYECTOS'!C3:C34,32,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(AND(RC[5]>0,RC[-1]>0,RC[-1]<RC[5]),0,IF(AND(RC[-1]>0,RC[-1]>0,RC[-1]>RC[5]),RC[-1]-RC[5],RC[-1]))"
    ActiveCell.Offset(0, 2).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-50]=""CONSULTOR"",IFERROR(VLOOKUP(RC[-51],Vacaciones!C[-52]:C[-46],6,0),0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = ""
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = ""
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-54],Garantizados!C[-53]:C[-50],3,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(AVERAGE(RC[-46],RC[-35]),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(IF(AND(RC[-1]>=60%,RC[-1]<=69.99%),(SUMIFS('Base Móvil'!C[-18],'Base Móvil'!C[-17],RC[-56]))*5%,IF(AND(RC[-1]>=70%,RC[-1]<=89.99%),(SUMIFS('Base Móvil'!C[-18],'Base Móvil'!C[-17],RC[-56]))*10%,IF(RC[-1]>=90%,(SUMIFS('Base Móvil'!C[-18],'Base Móvil'!C[-17],RC[-56]))*20%,0))),0)"
    ActiveCell.Offset(0, 2).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=AVERAGE(RC[-49],RC[-38],RC[-27],RC[-18])"
    ActiveCell.Offset(0, 2).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-60],'Meta General'!C1:C13,13,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC[-60]=""CONSULTOR"",RC[-60]=""ESPECIALISTA""),SUMIFS('Base Móvil'!C[-40],'Base Móvil'!C[-44],LIQUIDADOR!RC[-61],'Base Móvil'!C[-32],""Ok Fechas"",'Base Móvil'!C[-28],""a"",'Base Móvil'!C[-21],""CONVERGENCIA""),IF(OR(RC[-60]=""COORDINADOR"",RC[-60]=""COORCAN""),SUMIFS('Base Móvil'!C[-40],'Base Móvil'!C[-46],LIQUIDADOR!RC[-61],'Base Móvil'!C[-32],""Ok Fechas"",'Base Móvil'!C[-28],""a"",'Base Móvil'!C[-21],""CONVERGENCIA""),IF(RC[-60]=""JEFE"",SUMIFS('Base Móvil'!C[-40],'Base Móvil'!C[-48],LIQUIDADOR!RC[-61],'Base Móvil'!C[-32],""Ok Fechas"",'Base Móvil'!C[-28],""a"",'Base Móvil'!C[-21],""CONVERGENCIA""),IF(RC[-60]=""JEFECAN"",SUMIFS('Base Móvil'!C[-40],'Base Móvil'!C[-29],LIQUIDADOR!RC[-61],'Base Móvil'!C[-32],""Ok Fechas""," & _
        "'Base Móvil'!C[-28],""a"",'Base Móvil'!C[-21],""CONVERGENCIA""),IF(RC[-60]=""GERENTE"",SUMIFS('Base Móvil'!C[-40],'Base Móvil'!C[-50],LIQUIDADOR!RC[-61],'Base Móvil'!C[-32],""Ok Fechas"",'Base Móvil'!C[-28],""a"",'Base Móvil'!C[-21],""CONVERGENCIA""),IF(RC[-60]=""DIRECTOR"",SUMIFS('Base Móvil'!C[-40],'Base Móvil'!C[-32],""Ok Fechas"",'Base Móvil'!C[-28],""a"",'Base Móvil'!C[-21],""CONVERGENCIA""),IF(RC[-60]=""JEFETMK"",SUMIFS('Base Móvil'!C[-40],'Base Móvil'!C[-50],LIQUIDADOR!RC[-61],'Base Móvil'!C[-32],""Ok Fechas"",'Base Móvil'!C[-28],""a"",'Base Móvil'!C[-21],""CONVERGENCIA""))))))))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(RC[-1]/RC[-2],0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-43]>=60%,IF(AND(RC[-1]>=80%,RC[-1]<100%),5%,IF(RC[-1]>=100%,10%,0)),0%)"
Range("A1").Select
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic
MsgBox "LISTO SEGUNDA FASE", vbInformation
End Sub
Sub marado()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Meses").Select
Range("H1").Copy
Sheets("DESARROLLO+PROYECTOS").Select
Range("G1").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
:=False, Transpose:=False
Application.CutCopyMode = False
Range("I1").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(R1C7,Meses!R1C10:R13C11,2,FALSE)"
MsgBox "listo primera fase", vbInformation
Sheets("LIQUIDADOR").Select
Range("B4").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Sheets("DESARROLLO+PROYECTOS").Select
Range("C3").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
:=False, Transpose:=False
Application.CutCopyMode = False
Sheets("LIQUIDADOR").Select
Range("A1").Select
MsgBox "listo segunda fase", vbInformation
Sheets("DESARROLLO+PROYECTOS").Select
Range("C2").Select
fav = Selection.End(xlDown).Row - 2
Range("C3").Select
For h = 1 To fav
 ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],HC!C1:C18,3,FALSE)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-2],HC!C1:C18,5,FALSE)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-3],HC!C1:C18,18,FALSE)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=R1C7"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-5],HC!C1:C7,7,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC3,HC!C1:C7,7,0)*RC[1]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=100%-RC[3]-RC[14]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "PROYECTOS"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-9],'Proyectos-Desarrollo'!R2C2:R131C4,3,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IF(AND(RC[-1]=100%,RC[10]>0),RC[-1]-RC[10],RC[-1])"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = ""
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = ""
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = ""
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-14],'Proyectos-Desarrollo'!R2C2:R131C5,4,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]<80%,RC[-1],IF(RC[-1]>=80%,RC[-1]+10%,0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]*RC[-6]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(AND(RC[-3]>=90%,RC[13]>=100%),(RC[13]-100%+RC[-2])*RC[-7],RC[-1])"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]*RC[-13]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "DESARROLLO"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R1C9=""SI"",IF(ISNUMBER(SEARCH(""Gerente"",RC[-18])),20%,IF(ISNUMBER(SEARCH(""Director"",RC[-18])),20%,10%)),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = ""
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = ""
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = ""
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-25],'Proyectos-Desarrollo'!R2C8:R8080C13,6,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(LOOKUP(RC[-1],FX!R2C31:R8C33,FX!R2C33:R8C33),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]*RC[-6]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(AND(COUNT(RC[6]:RC[8])>=2,RC[-3]>=90%,RC[9]>=100%),(RC[9]-100%+RC[-2])*RC[-7],RC[-1])"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]*RC[-24]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(VLOOKUP(RC[-30],LIQUIDADOR!C2:C60,59,0)>=160%,160%,VLOOKUP(RC[-30],LIQUIDADOR!C2:C60,59,0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-2]+RC[-13]"
    ActiveCell.Offset(0, 3).Range("A1").Select
    ActiveCell.FormulaR1C1 = ""
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = ""
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = ""
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF((IFERROR(AVERAGE(RC[-3]:RC[-1]),0))>=160%,160%,IFERROR(AVERAGE(RC[-3]:RC[-1]),0))"
    ActiveCell.Offset(1, 0).Range("A1").Select
    Selection.End(xlToLeft).Select
Next h
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic
MsgBox "proceso FINAL FINAL completado CAPO, puede continuar con el siguiente paso", vbInformation
MsgBox "Por favor espere DOS minutos mientras se enfría el procesador, de lo contrario la memoria se desboradará", vbInformation
End Sub
Sub copola()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim archivos
Dim goku As Excel.Workbook
Dim vegueta As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String
archivos = Dir("D:\AUTOMA_FULL_NEGOCIOS\ESTADOS\*.csv")
MsgBox archivos
If archivos <> "ESTADOS.csv" Then
MsgBox "NO HAY NINGÚN ARCHIVO DE ESTADOS CARGADO, O EL NOMBRE DEL ARCHIVO NO ES EL CORRECTO, RECUERDE QUE EL ARCHIVO SE DEBE LLAMAR: ESTADOS, debe estar escrito en mayúscula y la extensión debe ser .csv", vbCritical
Else
Do While archivos <> ""
Workbooks.Open "D:\AUTOMA_FULL_NEGOCIOS\ESTADOS\" & archivos
archivos = Dir
Loop
bills = ActiveWorkbook.Name
tt = bills
Windows(tt).Activate
Sheets(1).Select
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Móvil").Select
Range("A1").Select
dragon = Selection.End(xlDown).Row
Range("AI2").Select
ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-32],ESTADOS.csv!C1:C2,2,FALSE),"""")"
ActiveCell.Copy
Range("$AI$3:AI" & dragon).Select
ActiveSheet.Paste
Application.CutCopyMode = False
Range("AI2").Select
Windows(tt).Activate
ActiveWorkbook.Close SaveChanges:=False
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
MsgBox "EL SISTEMA SE CERRARÁ, POR FAVOR NO MANIPULE EL COMPUTADOR HASTA QUE EL EJECUTABLE DE EXCEL SE CIERRE, VUELVA ABRIR EL ARCHIVO Y EJECUTE EL BOTÓN ESTADOS", vbInformation
ActiveWorkbook.Close SaveChanges:=True
End If
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub dita()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim archivos
Dim goku As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String
archivos = Dir("D:\AUTOMA_FULL_NEGOCIOS\PROYECTOS\*.xlsx")
Do While archivos <> ""
Workbooks.Open "D:\AUTOMA_FULL_NEGOCIOS\PROYECTOS\" & archivos
archivos = Dir
Loop
bills = ActiveWorkbook.Name
tt = bills
Windows(tt).Activate
RATA = Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
Sheets(1).Range("A4:F" & RATA).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Proyectos-Desarrollo").Select
Range("A3").Select
ActiveSheet.Paste
Range("A3").Select
Application.CutCopyMode = False
Windows(tt).Activate
ActiveWorkbook.Close
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic
MsgBox "PRIMERA FASE", vbInformation
End Sub
Sub BRUNI()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim archivos
Dim goku As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String
archivos = Dir("D:\AUTOMA_FULL_NEGOCIOS\DESEMPEÑO\*.xlsx")
If archivos = "" Then
MsgBox "NO HAY NINGÚN ARCHIVO DE DESARROLLO CARGADO", vbCritical
Else
Do While archivos <> ""
Workbooks.Open "D:\AUTOMA_FULL_NEGOCIOS\DESEMPEÑO\" & archivos
archivos = Dir
Loop
bills = ActiveWorkbook.Name
tt = bills
Windows(tt).Activate
RATA = Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
Sheets(1).Range("A2:F" & RATA).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Proyectos-Desarrollo").Select
Range("H3").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Windows(tt).Activate
ActiveWorkbook.Close
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Range("A3").Select
End If
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic
MsgBox "LISTA SEGUNDA FASE", vbInformation
End Sub
Sub riddle()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Range("A1").Select
fav = Selection.End(xlDown).Row
Columns("BB:BB").Delete
Set rangodatos = Sheets("Base Fijos").Range("A1:BB" & fav)
rangodatos.AutoFilter Field:=24, Criteria1:=Array("A", "CROSS SELLING", "NEW"), Operator:=xlFilterValues
Range("BB1").Select
ActiveCell.FormulaR1C1 = "=IF((IFERROR(VLOOKUP(RC[-31],SIN_TURNOS!C1,1,FALSE),""""))=RC[-31],""VENTA SIN TURNO CAV"","""")"
ActiveCell.Copy
Sheets("Base Fijos").Range("BB1:BB" & fav).SpecialCells(xlCellTypeVisible).Select
Selection.PasteSpecial xlPasteFormulas
Application.CutCopyMode = False
ActiveSheet.ShowAllData
Range("BB1") = "MARCA NO PAGO TURNO CAV"
Set rangodatos = Sheets("Base Fijos").Range("A1:BB" & fav)
rangodatos.AutoFilter Field:=54, Criteria1:="<>"
Range("N1").Select
ActiveCell.FormulaR1C1 = "=RC[3]"
ActiveCell.Copy
Sheets("Base Fijos").Range("N1:N" & fav).SpecialCells(xlCellTypeVisible).Select
Selection.PasteSpecial xlPasteFormulas
Application.CutCopyMode = False
Range("A1").Select
Range("N1") = "CARGO MENSUAL VENTA SIN TURNO"
ActiveSheet.ShowAllData
ActiveCell.Columns("N:N").EntireColumn.Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False
Range("A1").Select
Set rangodatos = Sheets("Base Fijos").Range("A1:BB" & fav)
rangodatos.AutoFilter Field:=54, Criteria1:="<>"
Range("Q1").Select
Range("Q1") = "0"
ActiveCell.Copy
Sheets("Base Fijos").Range("Q1:Q" & fav).SpecialCells(xlCellTypeVisible).Select
Selection.PasteSpecial xlPasteValues
Application.CutCopyMode = False
Range("A1").Select
Range("Q1") = "VALOR MENSUALIDAD"
ActiveSheet.ShowAllData
MsgBox "LISTO PRIMERA FASE", vbInformation
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub tombola()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Sheets("Base Móvil").Select
Columns("AI:AI").Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
:=False, Transpose:=False
Application.CutCopyMode = False
Range("AI2").Select
ActiveWorkbook.Save
MsgBox "proceso FINAL FINAL completado CAPO, puede continuar con el siguiente paso", vbInformation
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub brie()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Móvil").Select
Range("A1").Select
dragon = Selection.End(xlDown).Row
Range("S2").Select
ActiveCell.FormulaR1C1 = _
"=IF(ISNUMBER(SEARCH(""Enl.Direc"",RC[9])),IF(IFERROR(VLOOKUP(RC[22],LIQUIDADOR!C2:C3,2,0),0)=""CONSULTOR"","""",RC[22]),IF(ISNUMBER(SEARCH(""Enlacedirecto"",RC[9])),IF(IFERROR(VLOOKUP(RC[22],LIQUIDADOR!C2:C3,2,0),0)=""CONSULTOR"","""",RC[22]),RC[22]))"
ActiveCell.Copy
Range("S3:S" & dragon).Select
ActiveSheet.Paste
Application.CutCopyMode = False
Range("S2").Select
MsgBox "proceso FINAL FINAL completado CAPO, puede continuar con el siguiente paso", vbInformation
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub kross()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("Base Fijos").Select
Columns("J:J").Select
Range("J2").Activate
Selection.Replace What:="CLOUD", Replacement:="_CLOUD", LookAt:=xlPart, _
SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
ReplaceFormat:=False
Range("D1").Select
MsgBox "Listo primera fase", vbInformation
Columns("J:J").Select
Range("J2").Activate
Selection.Replace What:="_CLOUDb", Replacement:="RETO", LookAt:=xlPart, _
SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
ReplaceFormat:=False
Range("D1").Select
MsgBox "Listo segunda fase", vbInformation
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub scatter()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim archivos
Dim goku As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String
archivos = Dir("D:\AUTOMA_FULL_NEGOCIOS\TURNOS\*.xlsx")
If archivos = "" Then
MsgBox "NO HAY NINGÚN ARCHIVO DE TURNOS CARGADO", vbCritical
Else
Do While archivos <> ""
Workbooks.Open "D:\AUTOMA_FULL_NEGOCIOS\TURNOS\" & archivos
archivos = Dir
Loop
bills = ActiveWorkbook.Name
tt = bills
Windows(tt).Activate
Sheets("servicios fijos").Select
tour = Sheets("servicios fijos").Range("A" & Rows.Count).End(xlUp).Row
Sheets("servicios fijos").Range("A2:B" & tour).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("SIN_TURNOS").Select
Range("A2").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Windows(tt).Activate
Sheets("pospago - prepago - power").Select
tour = Sheets("pospago - prepago - power").Range("A" & Rows.Count).End(xlUp).Row
Sheets("pospago - prepago - power").Range("A2:B" & tour).Copy
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Sheets("SIN_TURNOS").Select
Range("F2").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Windows(tt).Activate
ActiveWorkbook.Close
Windows("PLANTILLA_TOP_NEGOCIOS.xlsm").Activate
Range("A3").Select
End If
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic
MsgBox "proceso FINAL FINAL completado CAPO, puede continuar con el siguiente paso", vbInformation
End Sub

