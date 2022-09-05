Attribute VB_Name = "Module1"
Public Type FeedBack
    itemName As String
    ItemPrice As String
    ItemDate As String
    itemNumber As String
End Type

Public Type eBayItem
    Year As String
    Make As String
    Model As String
    Bid As String
    ItemNo As String
    Damaged As String
    Mileage As String
    VIN As String
    VehitcleTitle As String
    Transmission As String
    Engine As String
    ExtColor As String
    IntColor As String
End Type


Public Function ReadRawData(ByVal XLSFileName As String) As String()
    Dim oExcel As New Excel.Application
    Dim oWB As New Excel.Workbook
    Dim oSheet As New Excel.Worksheet
    Dim itemName As String
    Dim ItemNo As String
    Dim retVal() As String
    Dim i As Long
    
    Set oWB = oExcel.Workbooks.Open(XLSFileName)
    Set oSheet = oWB.Sheets(1)
    
    i = 0
    
    For j = 2 To oSheet.Rows.Count
        itemName = oSheet.Cells(j, 1).Value
        If itemName = "" Then Exit For
        ItemNo = Split(itemName, "(")(1)
        ReDim Preserve retVal(i)
        retVal(i) = Replace(ItemNo, ")", "")
        i = i + 1
    Next
    oExcel.Quit
    ReadRawData = retVal
End Function



Public Sub WriteRawData(MyArray() As eBayItem, ByVal XLSFileName As String)
    Dim oExcel As New Excel.Application
    Dim oWB As New Excel.Workbook
    Dim oSheet As New Excel.Worksheet
    Dim itemName As String
    Dim ItemNo As String
    Dim retVal() As String
    Dim i As Long
    
    Print #1, "Opening output file " + XLSFileName + vbCrLf
    Set oWB = oExcel.Workbooks.Open(XLSFileName)
    Set oSheet = oWB.Sheets(1)

    
    For i = LBound(MyArray) To UBound(MyArray)
        Print #1, "Writing Ietm  " + MyArray(i).ItemNo + vbCrLf
        oSheet.Cells(i + 2, 1).Value = MyArray(i).Year
        oSheet.Cells(i + 2, 2).Value = MyArray(i).Make
        oSheet.Cells(i + 2, 3).Value = MyArray(i).Model
        oSheet.Cells(i + 2, 4).Value = MyArray(i).Bid
        oSheet.Cells(i + 2, 5).Value = MyArray(i).ItemNo
        oSheet.Cells(i + 2, 6).Value = MyArray(i).Damaged
        oSheet.Cells(i + 2, 7).Value = MyArray(i).Mileage
        oSheet.Cells(i + 2, 8).Value = MyArray(i).VIN
        oSheet.Cells(i + 2, 9).Value = MyArray(i).VehitcleTitle
        oSheet.Cells(i + 2, 10).Value = MyArray(i).Transmission
        oSheet.Cells(i + 2, 11).Value = MyArray(i).Engine
        oSheet.Cells(i + 2, 12).Value = MyArray(i).ExtColor
        oSheet.Cells(i + 2, 13).Value = MyArray(i).IntColor
    Next
    
    oWB.Save
    oExcel.Quit
    
End Sub

