Attribute VB_Name = "MODGB_OFFERTA"
Option Explicit

Public Function RTrimN(ByVal fvar_value As Variant) As Variant
    If IsNull(fvar_value) Or IsEmpty(fvar_value) Or (fvar_value = "") Then
        RTrimN = ""
    Else
        RTrimN = RTrim(fvar_value)
    End If
End Function

Public Function CDecN(ByVal fvar_value As Variant) As Variant
    If IsNull(fvar_value) Or IsEmpty(fvar_value) Or (fvar_value = "") Then
        CDecN = 0
    Else
        CDecN = CDec(fvar_value)
    End If
End Function

Public Function NVL(Valore As Variant, ValIfNull As Variant) As Variant
    On Error Resume Next

    If IsEmpty(Valore) Or IsNull(Valore) Then
        NVL = ValIfNull
    Else
        If Trim(CStr(Valore)) = "" Then
            NVL = ValIfNull
        Else
            NVL = Trim(Valore)
        End If
    End If
    
    Err.Clear
End Function


Public Function FormattaAlfanumerico(Optional ByVal Valore As String, Optional Lunghezza As Integer) As String
  If NVL(Valore, "") <> "" Then
    'Elimina caratteri a capo
    Valore = Replace(Valore, vbCrLf, "")
    
    If Len(Valore) > Lunghezza Then
      FormattaAlfanumerico = Mid(Valore, 1, Lunghezza)
    Else
      FormattaAlfanumerico = Valore
    End If
  Else
    FormattaAlfanumerico = ""
  End If
End Function

'si aspetta un formato data GG/MM/AAAA
' e restituisce il formato  AAAAMMGG
Public Function FormattaDataNum(Valore) As String
  On Error GoTo ErrTrap
  
  Dim ValoreData As Date
  If NVL(Valore, "") <> "" Then
    ValoreData = CDate(Valore)
    
    'Formato CCYYMMDD
    FormattaDataNum = Year(Valore)
    FormattaDataNum = FormattaDataNum & String(2 - Len(Month(Valore)), "0") & Month(Valore)
    FormattaDataNum = FormattaDataNum & String(2 - Len(Day(Valore)), "0") & Day(Valore)
  Else
    FormattaDataNum = ""
  End If
  
  Exit Function
ErrTrap:
  FormattaDataNum = ""
  
End Function

'si aspetta un formato data GG/MM/AAAA
' e restituisce il formato  GGMMAAAA
Public Function FormattaDataNumGMA(Valore) As String
  On Error GoTo ErrTrap
  
  Dim ValoreData As Date
  If NVL(Valore, "") <> "" Then
    ValoreData = CDate(Valore)
    
    'Formato DDMMYYYY
    FormattaDataNumGMA = String(2 - Len(Day(Valore)), "0") & Day(Valore)
    FormattaDataNumGMA = FormattaDataNumGMA & String(2 - Len(Month(Valore)), "0") & Month(Valore)
    FormattaDataNumGMA = FormattaDataNumGMA & Year(Valore)
  Else
    FormattaDataNumGMA = ""
  End If
  
  Exit Function
ErrTrap:
  FormattaDataNumGMA = ""
  
End Function
'si aspetta un formato data AAAAMMGG
' e restituisce il formato GG/MM/AAAA
Public Function FormattaDataIta(Valore) As String
  On Error GoTo ErrTrap
  
  If NVL(Valore, "") <> "" Then
    FormattaDataIta = Mid(Valore, 7, 2) & "/" & Mid(Valore, 5, 2) & "/" & Mid(Valore, 1, 4)
  Else
    FormattaDataIta = ""
  End If
  
  Exit Function
ErrTrap:
  FormattaDataIta = ""
  
End Function

'si aspetta un formato data AAAAMMGG
' e restituisce il formato MM/GG/AAAA
Public Function FormattaDataIng(Valore) As String
  On Error GoTo ErrTrap
  
  If NVL(Valore, "") <> "" Then
    FormattaDataIng = Mid(Valore, 5, 2) & "/" & Mid(Valore, 7, 2) & "/" & Mid(Valore, 1, 4)
  Else
    FormattaDataIng = ""
  End If
  
  Exit Function
ErrTrap:
  FormattaDataIng = ""
  
End Function


Public Function FormattaNumerico(Valore As Double, Lunghezza As Integer, GestSegno As Boolean, NumInteri As Integer, NumDecimali As Integer) As String
  
  Dim ValIntero     As Double
  Dim ParteIntero   As String
  
  Dim ValDecimale   As Double
  Dim ParteDecimale As String
  
  'Prendo il valore intero
  ValIntero = Int(Abs(Valore))
  ParteIntero = Format(ValIntero, String(NumInteri, "0"))
  
  'Recupero il valore decimale
'  ValDecimale = Replace(CStr(Round(Abs(Valore) - ValIntero, 6)), "0,", "")
'  ParteDecimale = Mid(ValDecimale, 1, NumDecimali)
  ValDecimale = CStr(Round(Abs(Valore) - ValIntero, 6))
  ParteDecimale = Mid(ValDecimale, 3, NumDecimali)
  If NumDecimali > Len(ParteDecimale) Then
    ParteDecimale = ParteDecimale & String(NumDecimali - Len(ParteDecimale), "0")
  Else
    ParteDecimale = Mid(ParteDecimale, 1, NumDecimali)
  End If
  
  If GestSegno Then
    If Sgn(Valore) >= 0 Then
      FormattaNumerico = "+"
    Else
      FormattaNumerico = "-"
    End If
  Else
    FormattaNumerico = ""
  End If
  
  FormattaNumerico = FormattaNumerico & ParteIntero & ParteDecimale
  
End Function


Public Function DataDB(ByRef Data As Variant) As Variant
    Dim PclsData                As CLSFW_Date                      'data
    
    On Error GoTo ErrDataDB
    Set PclsData = New CLSFW_Date
    
    PclsData.DBformat CStr("" & Data)
    
    If PclsData.RetStatus = 0 Then
       Data = PclsData.ValDBdate
    Else
       Data = Data
    End If
    
    Set PclsData = Nothing
    
    Exit Function
ErrDataDB:
    DataDB = "null"
    Err.Clear
    Exit Function
End Function


Public Function SQLDate(ByVal vdatData As Date) As String
  SQLDate = "CONVERT(DATETIME, '" & Year(vdatData) & "-" & _
                             Format(Month(vdatData), "00") & "-" & _
                             Format(Day(vdatData), "00") & _
                             " " & Format(Hour(vdatData), "00") & ":" & _
                             Format(Minute(vdatData), "00") & ":" & _
                             Format(Second(vdatData), "00") & "', 102)"
End Function

Public Function SQLDouble(ByVal vdblValue As Double) As String
  SQLDouble = Replace(vdblValue, ",", ".")
End Function

Public Function SQLString(ByVal vStr As String) As String
  SQLString = Replace(vStr, "'", "''")
End Function


