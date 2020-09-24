Attribute VB_Name = "modScript"
Type typAlias
    Name As String
    Code As String
End Type

Public Aliases()    As typAlias
Public intAliasCnt  As Integer

Public parameters() As String
Public intParamCnt  As Integer
Public Sub AddAlias(strName As String, strCode As String)
    intAliasCnt = intAliasCnt + 1
    ReDim Preserve Aliases(intAliasCnt) As typAlias
    
    Aliases(intAliasCnt).Name = strName
    Aliases(intAliasCnt).Code = strCode
End Sub


Function AliasExists(strName As String) As Boolean
    Dim i As Integer
    For i = 1 To intAliasCnt
        If LCase(strName) = LCase(Aliases(i).Name) Then
            AliasExists = True
            Exit Function
        End If
    Next i
    AliasExists = False

End Function

Sub CopyParameters(strArray() As String)
    ReDim parameters(UBound(strArray) + 1) As String
    intParamCnt = UBound(strArray) + 1
    
    For i = 1 To intParamCnt
        parameters(i) = strArray(i - 1)
    Next i
End Sub

Public Sub EditAlias(strOldName As String, strNewName As String, Optional strNewCode As String = "")
    Dim i As Integer
    For i = 1 To intAliasCnt
        If LCase(strOldName) = LCase(Aliases(i).Name) Then
            Aliases(i).Name = strNewName
            If strNewCode <> "" Then Aliases(i).Code = strNewCode
            Exit Sub
        End If
    Next i
End Sub


Public Function GetAliasCode(strName As String) As String
    Dim i As Integer
    For i = 1 To intAliasCnt
        If LCase(strName) = LCase(Aliases(i).Name) Then
            GetAliasCode = Aliases(i).Code
            Exit Function
        End If
    Next i
    GetAliasCode = Chr(8)   'doesnt exist
End Function
Sub RemoveAlias(strAlias As String)
    Dim i As Integer, j As Integer
    
    For i = 1 To intAliasCnt
        If LCase(strAlias) = LCase(Aliases(i).Name) Then
            Aliases(i).Name = ""
            Aliases(i).Code = ""
            For j = i + 1 To intAliasCnt
                Aliases(j - 1).Name = Aliases(j).Name
                Aliases(j - 1).Code = Aliases(j).Code
            Next j
            intAliasCnt = intAliasCnt - 1
            Exit Sub
        End If
    Next i
    
End Sub


