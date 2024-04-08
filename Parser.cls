VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Parser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Phrases As New Dictionary
Public Event PhraseResult(ResultName As String, Operator As String, _
                          Value As String, ByRef Result As String)

Private Const AndOperator = "AND"
Private Const OrOperator = "OR"
Private Const NotOperator = "NOT"
Private Const BracesOpen = "("
Private Const BracesClose = ")"
Private Const NotEqualOperator = "!="
Private Const EqualOperator = "="
Private Const ContainOperator = "CONTAIN"

Private SCodes As String

Private Function CalculatePhrase(Phrase As String, SnomedCode As String) As String
    Dim start As Integer
    Dim ReplacePhrase As String
    CalculatePhrase = ""
    start = 1
    Phrase = UCase(Trim(Phrase))
    Phrase = Replace(Phrase, " " & AndOperator & " ", "#" & AndOperator & "#")
    Phrase = Replace(Phrase, " " & OrOperator & " ", "#" & OrOperator & "#")
    Phrase = Replace(Phrase, " " & NotOperator & " ", "#" & NotOperator & "#")
    Phrase = Replace(Phrase, " " & BracesOpen & " ", "#" & BracesOpen & "#")
    Phrase = Replace(Phrase, " " & BracesClose & " ", "#" & BracesClose & "#")
    ReplacePhrase = Phrase
    TranslatePhrase Phrase, ReplacePhrase, start
    ReplacePhrase = Replace(ReplacePhrase, "#", " ")
    If ConvBool(ReplacePhrase) Then
        CalculatePhrase = SnomedCode
    End If
End Function

Private Sub TranslatePhrase(Phrase As String, ByRef ReplacePhrase As String, _
                            ByRef start As Integer)
    Dim SubPhrase As String
    Dim idx As Integer
    idx = InStr(start, Phrase, "#")
    If idx = 0 Then idx = Len(Phrase) + 1
    
    If idx <> start Then
        SubPhrase = Mid(Phrase, start, idx - start)
        If Trim(SubPhrase) <> "" Then
            ReplacePhrase = Replace(ReplacePhrase, SubPhrase, TranslateSubPhrase(SubPhrase))
        End If
    End If
    
    start = InStr(idx + 1, Phrase, "#") + 1
    If start <> 1 Then
        Call TranslatePhrase(Phrase, ReplacePhrase, start)
    End If

End Sub


Private Function TranslateSubPhrase(SubPhrase As String) As String
    Dim ResultName As String
    Dim Operator As String
    Dim Value As String
    Dim Result As String
    Dim idx As Integer
    Dim start As Integer
    
    TranslateSubPhrase = False
    If InStr(1, SubPhrase, EqualOperator) Then
        Operator = EqualOperator
        idx = InStr(1, SubPhrase, EqualOperator)
    ElseIf InStr(1, SubPhrase, ContainOperator) Then
        Operator = ContainOperator
        idx = InStr(1, SubPhrase, ContainOperator)
    ElseIf InStr(1, SubPhrase, NotEqualOperator) Then
        Operator = NotEqualOperator
        idx = InStr(1, SubPhrase, NotEqualOperator)
    End If
    
    If idx = 0 Then
        Exit Function
    End If
    ResultName = Trim(Left(SubPhrase, idx - 1))
    start = idx + Len(Operator)
    Value = Trim(Mid(SubPhrase, start, Len(SubPhrase) - start + 1))
    RaiseEvent PhraseResult(ResultName, Operator, Value, Result)
    TranslateSubPhrase = Result
 End Function



Private Function ConvBool(s As String) As Boolean
Dim a As Integer
Dim b As Integer
Dim c As String
Dim d As String
Dim e As String
    s = Trim(UCase(s))
    If s = "TRUE" Then
        ConvBool = True
    ElseIf s = "FALSE" Then
        ConvBool = False
    Else
        If InStr(s, BracesOpen) > 0 Then
            a = InStr(s, BracesOpen)
            b = FindCloser(Mid(s, InStr(s, BracesOpen) + 1))
            c = Trim(Mid(s, 1, a - 1))
            d = Trim(Mid(s, a + 1, b - 1))
            e = Trim(Mid(s, a + b + 1))
            ConvBool = ConvBool(c & " " & CStr(ConvBool(d)) & " " & e)
        ElseIf InStr(s, OrOperator) > 0 Then
            ConvBool = ConvBool(Mid(s, 1, InStr(s, OrOperator) - 1)) Or ConvBool(Mid(s, InStr(s, OrOperator) + 2))
        ElseIf InStr(s, AndOperator) > 0 Then
            ConvBool = ConvBool(Mid(s, 1, InStr(s, AndOperator) - 1)) And ConvBool(Mid(s, InStr(s, AndOperator) + 3))
        ElseIf Left(s, 3) = NotOperator Then
            ConvBool = Not ConvBool(Mid(s, 4))
        End If
    End If
End Function

Private Function FindCloser(s As String) As Integer
    Dim Counter As Integer
    Dim i As Integer
    Counter = 1
    i = 1
    While Counter > 0
        If Mid(s, i, 1) = BracesOpen Then
            Counter = Counter + 1
        ElseIf Mid(s, i, 1) = BracesClose Then
            Counter = Counter - 1
        End If
        i = i + 1
    Wend
    FindCloser = i - 1
End Function

Public Function addPhrase(Phrase As String, SnomedCode As String)
    If Not Phrases.Exists(Phrase) Then
        Phrases.Add Phrase, SnomedCode
    End If
End Function

Public Sub CalculateSnomed()
    Dim i As Integer
    Dim CalPhrase As String
    Dim a
    Dim key As String
    
    SnomedCodes = ""
    a = Phrases.Keys
    For i = 0 To Phrases.Count - 1
        key = CStr(a(i))
        CalPhrase = CalculatePhrase(key, Phrases.Item(key))
        If CalPhrase <> "" And Not InStr(1, SnomedCodes, CalPhrase) Then
            If SnomedCodes = "" Then
                SnomedCodes = CalPhrase
            Else
                SnomedCodes = SnomedCodes & "," & CalPhrase
            End If
        End If
    Next i
    
    Phrases.RemoveAll
End Sub

Public Property Get EqOperator() As String
    EqOperator = EqualOperator
End Property

Public Property Get NotEqOperator() As String
    NotEqOperator = NotEqualOperator
End Property

Public Property Get ConOperator() As String
    ConOperator = ContainOperator
End Property

Private Sub Class_Terminate()
    Phrases.RemoveAll
    Set Phrases = Nothing
End Sub


Public Property Get SnomedCodes() As String
    SnomedCodes = SCodes
End Property

Public Property Let SnomedCodes(ByVal vNewValue As String)
    SCodes = vNewValue
End Property
