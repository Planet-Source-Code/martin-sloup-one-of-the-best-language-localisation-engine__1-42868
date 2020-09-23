Attribute VB_Name = "modLang"
Option Explicit

Private Type TVarible
    Name As String
    Value As String
End Type

Private Type TForms
    Name As String
    Value As String
End Type

Public Type TInfo
    Name As String
    Value As String
End Type

Public LR As New LoadRes
Public Varible() As TVarible
Public FormsValues() As TVarible

Public Function LoadLang(sLangFile As String) As Boolean
    Dim tmp As String, tmpArr() As String, tmpArr2() As String, i As Long
    LR.DllName = sLangFile
    
    Erase FormsValues
    Erase Varible
    
    tmp = LR.LoadHtmlFromDLL("varibles.lang", 10&)
    tmpArr = Split(tmp, vbCrLf)
    ReDim Varible(0 To UBound(tmpArr))
    For i = LBound(tmpArr) To UBound(tmpArr)
        If Trim(tmpArr(i)) <> "" And Left(Trim(tmpArr(i)), 1) <> "#" Then
            If LCase(Left(tmpArr(i), 4)) = "var " Then
                tmpArr(i) = Mid(tmpArr(i), 5)
                tmpArr2 = Split(tmpArr(i), "=")
                Varible(i).Name = tmpArr2(0)
                If Left(tmpArr2(1), 1) = """" Then tmpArr2(1) = Mid(tmpArr2(1), 2)
                If Right(tmpArr2(1), 1) = """" Then tmpArr2(1) = Left(tmpArr2(1), Len(tmpArr2(1)) - 1)
                tmpArr2(1) = Replace(tmpArr2(1), "\r\n", vbCrLf)
                Varible(i).Value = tmpArr2(1)
            End If
        End If
    Next i
    
    tmp = LR.LoadHtmlFromDLL("forms.lang", 10&)
    tmpArr = Split(tmp, vbCrLf)
    ReDim FormsValues(0 To UBound(tmpArr))
    For i = LBound(tmpArr) To UBound(tmpArr)
        If Trim(tmpArr(i)) <> "" And Left(Trim(tmpArr(i)), 1) <> "#" Then
            tmpArr2 = Split(tmpArr(i), "=")
            tmpArr2(0) = Trim(tmpArr2(0))
            tmpArr2(1) = Trim(tmpArr2(1))
            FormsValues(i).Name = tmpArr2(0)
            If Right(tmpArr2(1), 2) = vbCrLf Then tmpArr2(1) = Left(tmpArr2(1), Len(tmpArr2(1)) - 2)
            tmpArr2(1) = Trim(tmpArr2(1))
            If Left(tmpArr2(1), 1) = """" Then tmpArr2(1) = Mid(tmpArr2(1), 2)
            If Right(tmpArr2(1), 1) = """" Then tmpArr2(1) = Left(tmpArr2(1), Len(tmpArr2(1)) - 1)
            tmpArr2(1) = Replace(tmpArr2(1), "\r\n", vbCrLf)
            FormsValues(i).Value = tmpArr2(1)
        Else
            FormsValues(i).Name = ""
            FormsValues(i).Value = ""
        End If
    Next i

End Function

Public Function GetLangVar(fForm As Form, sVar As String, ParamArray arrParams()) As String
    On Error Resume Next
    Dim i As Long, y As Long
    Dim tmp As String
    For i = LBound(Varible) To UBound(Varible)
        If LCase(fForm.Name & "." & sVar) = LCase(Varible(i).Name) Then
            tmp = Varible(i).Value
            
            For y = LBound(arrParams) To UBound(arrParams)
                tmp = Replace(tmp, "%" & CStr(y + 1), arrParams(y))
            Next y
            GetLangVar = tmp
            Exit Function
        End If
    Next i
End Function

Public Sub SetFormLang(fForm As Form)
    Dim FormName As String, FormValueName As String, FormValue As String, tmp As String
    Dim arrSep() As String
    Dim i As Long, y As Long, Index As Integer
    Dim obj As Object, Control As Control, Control2 As Control
    
    On Error Resume Next
    For i = 0 To UBound(FormsValues)
        If FormsValues(i).Name <> "" Then
            FormName = Left(FormsValues(i).Name, InStr(FormsValues(i).Name, ".") - 1)
            FormValueName = Mid(FormsValues(i).Name, InStr(FormsValues(i).Name, ".") + 1)
            FormValue = FormsValues(i).Value
            
            On Error GoTo FormError
            If LCase(fForm.Name) = LCase(FormName) Then
                Set obj = fForm
                If InStr(FormValueName, ".") = 0 Then
                    CallByName fForm, FormValueName, VbLet, FormValue
                Else
                    arrSep = Split(FormValueName, ".")
                    For y = 0 To UBound(arrSep)
                        If Right(arrSep(y), 1) = ")" Then
                            tmp = Mid(arrSep(y), InStr(arrSep(y), "(") + 1)
                            Index = CInt(Left(tmp, Len(tmp) - 1))
                            If UBound(arrSep) <> y Then
                                Set obj = CallByName(obj, Left(arrSep(y), InStr(arrSep(y), "(") - 1), VbGet, Index)
                            Else
                                CallByName obj, Left(arrSep(y), InStr(arrSep(y), "(") - 1), VbLet, Index, FormValue
                            End If
                        Else
                            If UBound(arrSep) <> y Then
                                If y = 0 Then
                                    For Each Control In obj.Controls
                                       If LCase(arrSep(y)) = LCase(Control.Name) Then Set obj = Control
                                    Next Control
                                Else
                                    Set obj = CallByName(obj, arrSep(y), VbGet)
                                End If
                            Else
                                CallByName obj, arrSep(y), VbLet, FormValue
                            End If
                        End If
                    Next y
                End If
            End If
            GoTo NextFor
FormError:
            MsgBox "Error on line " & CStr(i) & " in " & GetFileFromFilePath(LR.DllName) & "/forms.lang language file :" & vbCrLf & " " & vbCrLf & Err.Description, vbCritical + vbOKOnly
NextFor:
        End If
    Next i
End Sub

Public Function LoadLangPicture(ID, PicType As PicType) As IPictureDisp
    Set LoadLangPicture = LR.LoadPictureFromDLL(ID, PicType)
End Function

Public Function LoadAllLangInfo(sPath As String, lst As ListBox) As String()
    Dim sFiles() As String
    Dim sFilename As String
    Dim arrInfo() As TInfo
    Dim MaxCount As Long
    Dim sTmp As String
    Dim i As Long
    
    MaxCount = 0
    lst.Clear
    sTmp = App.Path & "\language\"
    sFilename = Dir(sTmp)
    Do While sFilename <> ""
        DoEvents
        If LCase(Right(sFilename, 8)) = "lang.lng" Then
            ReDim Preserve sFiles(0 To MaxCount)
            sFiles(MaxCount) = sPath & sFilename
            MaxCount = MaxCount + 1
        End If
        sFilename = Dir
    Loop
    For i = LBound(sFiles) To UBound(sFiles)
        arrInfo = LoadLangInfo(sFiles(i))
        lst.AddItem GetLangInfoVal("Language_Name", arrInfo)
    Next i
    LoadAllLangInfo = sFiles
End Function

Public Function LoadLangInfo(sFile As String) As TInfo()
    Dim tmp As String, tmpArr() As String, tmpArr2() As String, i As Long
    Dim MyLR As New LoadRes, Infos() As TInfo
    MyLR.DllName = sFile
    
    tmp = MyLR.LoadHtmlFromDLL("info.lang", 10&)
    tmpArr = Split(tmp, vbCrLf)
    ReDim Infos(0 To UBound(tmpArr))
    For i = LBound(tmpArr) To UBound(tmpArr)
        If Trim(tmpArr(i)) <> "" And Left(Trim(tmpArr(i)), 1) <> "#" Then
            If LCase(Left(tmpArr(i), 4)) = "var " Then
                tmpArr(i) = Mid(tmpArr(i), 5)
                tmpArr2 = Split(tmpArr(i), "=")
                Infos(i).Name = tmpArr2(0)
                If Left(tmpArr2(1), 1) = """" Then tmpArr2(1) = Mid(tmpArr2(1), 2)
                If Right(tmpArr2(1), 1) = """" Then tmpArr2(1) = Left(tmpArr2(1), Len(tmpArr2(1)) - 1)
                tmpArr2(1) = Replace(tmpArr2(1), "\r\n", vbCrLf)
                Infos(i).Value = tmpArr2(1)
            End If
        End If
    Next i
    LoadLangInfo = Infos
End Function

Public Function GetLangInfoVal(sName As String, arrInfo() As TInfo, ParamArray arrParams()) As String
    On Error Resume Next
    Dim i As Long, y As Long
    Dim tmp As String
    For i = LBound(arrInfo) To UBound(arrInfo)
        If LCase(sName) = LCase(arrInfo(i).Name) Then
            tmp = arrInfo(i).Value
            
            For y = LBound(arrParams) To UBound(arrParams)
                tmp = Replace(tmp, "%" & CStr(i + 1), arrParams(i))
            Next y
            GetLangInfoVal = tmp
            Exit Function
        End If
    Next i
End Function

Public Function GetPathFromString(sPathIn As String) As String
    Dim i As Integer
   For i = Len(sPathIn) To 1 Step -1
      If InStr(":\", Mid$(sPathIn, i, 1)) Then Exit For
   Next
   GetPathFromString = Left$(sPathIn, i)
End Function

Public Function GetFileFromFilePath(FilePath As String) As String
    GetFileFromFilePath = Right(FilePath, Len(FilePath) - Len(GetPathFromString(FilePath)))
End Function

