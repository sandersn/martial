VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

  'MultiUse = -1  'True
  'Persistable = 0  'NotPersistable
  'DataBindingBehavior = 0  'vbNone
  'DataSourceBehavior = 0   'vbNone
  'MTSTransactionMode = 0   'NotAnMTSObject
'End

Option Explicit
Option Compare Text

Private aKey() As String
Private aVal() As Variant
Private Upper As Long
Private Last As Long

Public Sub Add(Key As String, Value As Variant)
    Dim i As Long, thisKey As String, found As Boolean

    ReDim Preserve aKey(Upper)
    ReDim Preserve aVal(Upper)

    For i = 0 To Upper
        thisKey = aKey(i)
        If StrComp(thisKey, Key) = 0 Then
            found = True
            Exit For
        End If
    Next

    If found Then
        aKey(i) = Key
        aVal(i) = Value
        Last = i
    Else
        aKey(Upper) = Key
        aVal(Upper) = Value
        Last = Upper
        Upper = Upper + 1
    End If
End Sub

Public Function Retrieve(Key As String, Optional Default As Variant = 0) As Variant
    Dim i As Long, thisKey As String, found As Boolean

    If Upper > 0 Then
        For i = 0 To Upper - 1
            thisKey = aKey(i)
            If StrComp(thisKey, Key) = 0 Then
                found = True
                Exit For
            End If
        Next

        If found Then
            Retrieve = aVal(i)
            Last = i
            Exit Function
        End If
    End If

    If Default = 0 Then
        Retrieve = -1
    Else
        Retrieve = Default
    End If
End Function

Public Sub Delete(Key As String)
    Dim i As Long, i2 As Long, thisKey As String, found As Boolean

    If Upper = 0 Then Exit Sub
    For i = 0 To Upper - 1
        thisKey = aKey(i)
        If StrComp(thisKey, Key) = 0 Then
            found = True
            Exit For
        End If
    Next

    If found = False Then Exit Sub

    Last = i

    i = i + 1
    For i2 = i To Upper - 1
        aKey(i2 - 1) = aKey(i2)
        aVal(i2 - 1) = aVal(i2)
    Next

    Upper = Upper - 1
    ReDim Preserve aKey(Upper)
    ReDim Preserve aVal(Upper)
End Sub
'this function is currently not functional due to VB5's inability to return array from a function
Public Function Keys() As Variant  'return variant containing arrays of strings

    Dim i As Long, Key() As String
    ReDim Key(0 To Upper - 1)

    For i = 0 To Upper - 1
        Key(i) = aKey(i)
    Next

    Keys = Key
End Function
'this function is currently not functional due to VB5's inability to return array from a function
Public Function Values() As Variant 'return variant containing array of variants
    Dim i As Long, Val() As Variant
    ReDim Val(0 To Upper - 1)

    For i = 0 To Upper - 1
        Val(i) = aVal(i)
    Next

    Values = Val
End Function

Public Function Every() As Variant 'return variant containing array of variants
Dim KeyVal As Variant
'***original code***
'    Dim KeyVal() As Variant
'    ReDim KeyVal(0 To 1)
    Last = Last + 1
    If Last > Upper - 1 Then Last = 0
    
    KeyVal = Array(aKey(Last), aVal(Last))
'***original code***
'    KeyVal(0) = aKey(Last)
'    KeyVal(1) = aVal(Last)
    Every = KeyVal
End Function

Public Function getCount() As Long  'just a get func that should be a Property except that I'm too lazy to look up how to implement properties..
    getCount = Upper
End Function
