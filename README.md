# hello-world
hello-world
Hi， This is my first Code!

Public Function AImport(FilePath As String) As Variant

End Function


Implements iImport

Private Function iImport_AImport(FilePath As String) As Variant

    Dim PriceTable As Variant
    Dim i As Long
    Dim Temp As Variant
 
    Workbooks.Open FileName:=FilePath, local:=True
        PriceTable = Range(Range("A2"), Range("A1").End(xlDown).Offset(0, 4))
    ActiveWorkbook.Close SaveChanges:=False
    
    ReDim Temp(LBound(PriceTable, 1) To UBound(PriceTable, 1), 1 To 3)
  
    For i = LBound(PriceTable, 1) To UBound(PriceTable, 1)
        Temp(i, 1) = PriceTable(i, 2)
        Temp(i, 2) = PriceTable(i, 1)
        Temp(i, 3) = PriceTable(i, 5)
    Next i
    iImport_AImport = Temp
    
End Function

Function Dictionarize() As Dictionary
    Dim DName As Dictionary
    Set DName = New Dictionary
    Set Dictionarize = DName
End Function



Function RRCALC(Amt As Double, LossDate As Date, EndDate As Date, FundName As String, Optional Res As Long) As Double

    Dim i As Long
    Dim m As Long
    Dim Arr() As String
    Dim UAmt As Double
    DisArr = Empty
    
    Initialising
       
    UAmt = Amt / Coll(FundName)(LossDate)
    
    If IsInArray(FundName, KeyArr(9)) Then
    m = 1
    ReDim Preserve Arr(1 To 6, 1 To m)
    
    Arr(1, m) = FundName
    Arr(2, m) = CDbl(LossDate)
    Arr(3, m) = 0
    Arr(4, m) = 0
    Arr(5, m) = 0
    Arr(6, m) = UAmt
    
        For i = 0 To DisDict(FundName).Count - 1
                If DisDict(FundName).Keys()(i) > LossDate And DisDict(FundName).Keys()(i) < EndDate And Res = 1 Then
                    m = m + 1
                    ReDim Preserve Arr(1 To 6, 1 To m)
                    
                    Arr(1, m) = FundName
                    Arr(2, m) = CDbl(DisDict(FundName).Keys()(i))
                    Arr(3, m) = DisDict(FundName).Items()(i).CPU
                    Arr(4, m) = DisDict(FundName).Items()(i).RPrice
                    Arr(5, m) = DisDict(FundName).Items()(i).CPU * UAmt / 100
                    Arr(6, m) = UAmt
        
                ElseIf DisDict(FundName).Keys()(i) > LossDate And DisDict(FundName).Keys()(i) < EndDate And Res = 0 Then
                    UAmt = UAmt + UAmt * DisDict(FundName).Items()(i).CPU / (100 * DisDict(FundName).Items()(i).RPrice)
                    m = m + 1
                    ReDim Preserve Arr(1 To 6, 1 To m)
                    
                    Arr(1, m) = FundName
                    Arr(2, m) = CDbl(DisDict(FundName).Keys()(i))
                    Arr(3, m) = DisDict(FundName).Items()(i).CPU
                    Arr(4, m) = DisDict(FundName).Items()(i).RPrice
                    Arr(5, m) = DisDict(FundName).Items()(i).CPU * UAmt / 100
                    Arr(6, m) = UAmt
                
                End If
        Next i
        
    ReDim Preserve Arr(1 To 6, 1 To m + 1)
    
    Arr(1, m + 1) = FundName
    Arr(2, m + 1) = CDbl(EndDate)
    Arr(3, m + 1) = 0
    Arr(4, m + 1) = 0
    Arr(5, m + 1) = 0
    Arr(6, m + 1) = UAmt
           
    End If
    
    RRCALC = UAmt
    DisArr = Arr

End Function


Option Explicit

Public Coll As Collection

Public DisDict As Dictionary

Sub PriceCollDict()

    Dim X As Variant
    Dim i As Long
 
    Set Coll = New Collection

    For Each X In KeyArr(15)
        Coll.Add Dictionarize(), X
        For i = 2 To Sheet2.Cells(Rows.Count, 15).End(xlUp).Row
            If Sheet2.Cells(i, 15).Value = X Then
                Coll(X).Add Sheet2.Cells(i, 16).Value, Sheet2.Cells(i, 17).Value
            End If
        Next i
    Next

End Sub

Sub DisDictDict()

    Dim X As Variant
    Dim i As Long
  
    Set DisDict = New Dictionary
    
    For Each X In KeyArr(9)
        DisDict.Add X, Dictionarize()
        For i = 2 To Sheet2.Cells(Rows.Count, 9).End(xlUp).Row
            If Sheet2.Cells(i, 9).Value = X Then
                DisDict(X).Add Sheet2.Cells(i, 10).Value, DisRecord(Sheet2.Cells(i, 11).Value, Sheet2.Cells(i, 12).Value)
            End If
        Next i
    Next

End Sub
