VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "eBay Items Collector v1.0"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11625
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   11625
   StartUpPosition =   2  'CenterScreen
   Begin SHDocVwCtl.WebBrowser WebBrowser 
      Height          =   2295
      Left            =   1920
      TabIndex        =   4
      Top             =   -9000
      Width           =   7575
      ExtentX         =   13361
      ExtentY         =   4048
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   495
      Left            =   10200
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "&Start"
      Height          =   495
      Left            =   8880
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtPath 
      Height          =   375
      Left            =   120
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   360
      Width           =   11295
   End
   Begin VB.Label lblSubStatus 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   45
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   45
   End
   Begin VB.Label Label1 
      Caption         =   "Enter the folder location of XLS files (Eg: C:\data)"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'http://feedback.ebay.com/ws/eBayISAPI.dll?ViewFeedback2&ftab=FeedbackAsSeller&userid=allbmwparts&iid=-1&de=off&items=200

Dim colFeedBacks() As FeedBack

Private Sub cmdClose_Click()
    End
End Sub

Private Sub cmdProcess_Click()

    Dim SourcePath As String
    Dim XLSFileName As String
    
    Dim arItems() As String
    Dim itemNumber As String
    Dim i As Integer
    Dim CollectedData As eBayItem
    Dim outFile As String
    Dim colArray() As eBayItem
    Dim index As Integer
    
    If Trim(txtPath.Text) = "" Then
        MsgBox "Please enter the path of the XLS files in the text box", vbInformation
        Exit Sub
    End If
    
    cmdProcess.Enabled = False
        
    If Right(Trim(txtPath.Text), 1) = "\" Then
        SourcePath = txtPath.Text
    Else
        SourcePath = txtPath.Text + "\"
    End If
    
    XLSFileName = Dir$(SourcePath + "*.xls", vbNormal)
    
    
    Open App.Path + "\log.txt" For Output As #1
    
    While XLSFileName <> ""
        lblStatus.Caption = "Reading file " + XLSFileName
        lblStatus.Refresh
        Print #1, "v1.0  Reading from " + XLSFileName + vbCrLf
        arItems = ReadRawData(SourcePath + XLSFileName)
        
        outFile = App.Path + "\output\" + XLSFileName
        FileCopy App.Path + "\template.xls", outFile
        
        index = 0
        ReDim colArray(0)
    
        On Error Resume Next
        i = UBound(arItems)
        
        If Err Then
            Print #1, "No records found in" + XLSFileName + vbCrLf
        Else
            On Error GoTo 0
            Print #1, "Records found : " + Str(UBound(arItems)) + vbCrLf
            For i = 0 To UBound(arItems)
            
                itemNumber = arItems(i)
                
                lblStatus.Caption = "Processing file " + XLSFileName + "........ Item " + itemNumber
                lblStatus.Refresh
                
                Print #1, "Collecting data"
                
                CollectedData = GetItems(itemNumber)
                
                Print #1, "Item Name collected : " + CollectedData.Make + vbCrLf
                Print #1, "Item No collected : " + CollectedData.ItemNo + vbCrLf
                
                If CollectedData.ItemNo <> 0 Then
                    ReDim Preserve colArray(index)
                    colArray(index) = CollectedData
                    index = index + 1
                End If
                
            Next
            
            On Error Resume Next
            Print #1, "Total items collected : " + Str(UBound(colArray)) + vbCrLf
            On Error GoTo 0
            
            WriteRawData colArray, outFile
            
        End If
        XLSFileName = Dir$
        On Error GoTo 0
    Wend
    
    Close #1
    MsgBox "Files are copied to " + App.Path + "\output folder", vbInformation
    cmdProcess.Enabled = True
End Sub


Private Function GetItems(sItemNo As String) As eBayItem
    Dim Content As String
    Dim i As Integer
    Dim j As Integer
    
    Dim EHTML As Variant, TableElem As Variant, colElem As Variant, itemElems As Variant
    
    Dim Title As String
    Dim TitleParts() As String
    Dim CollectedData As eBayItem
    Dim RowsCollected As Integer
    Dim dt1 As Date
    Dim dt2 As Date

    
    
     
    
    RowsCollected = 0
    
    
    
    
    dt1 = Now()
    WebBrowser.Navigate "http://cgi.ebay.com/ebaymotors/ws/eBayISAPI.dll?ViewItem&_rdc=1&item=" + sItemNo
    
    Do While WebBrowser.ReadyState <> READYSTATE_COMPLETE
        dt2 = Now()
        If DateDiff("n", dt1, dt2) > 2 Then
            'Err.Raise 100, "Timeout loading " + sItemNo
            Exit Do
        End If
        DoEvents
    Loop
    
    CollectedData.ItemNo = sItemNo
    
    
    Set itemElems = WebBrowser.Document.getElementsByTagName("SPAN")
    For i = 0 To itemElems.Length
        Set EHTML = itemElems(i)
        If EHTML Is Nothing Then
        Else
            If Left(EHTML.innerText, 4) = "US $" Then
                CollectedData.Bid = Trim(Replace(EHTML.innerText, "US ", ""))
            End If
        End If
    Next

    If CollectedData.Bid = "" Then
        'Item not found
        Print #1, "Item not found  " + vbCrLf
        Print #1, WebBrowser.Document.body.innerHTML + vbCrLf
        
        CollectedData.ItemNo = 0
        GetItems = CollectedData
        Exit Function
    End If
    
    

    
    
    
    Set itemElems = WebBrowser.Document.getElementsByTagName("TH")
    For i = 0 To itemElems.Length
        Set EHTML = itemElems(i)
        If (Not EHTML Is Nothing) Then
            If InStr(1, EHTML.outerhtml, "class=""vi-ia-hdAl vi-ia-attrLabel vi-ia-attrColPadding") > 1 Then
                If Trim(EHTML.innerText) = "Mileage:" Then
                    CollectedData.Mileage = Replace(EHTML.nextsibling.innerText, " miles", "")
                End If
                If Trim(EHTML.innerText) = "VIN:" Then
                    CollectedData.VIN = Split(EHTML.nextsibling.innerText, "|")(0)
                End If
                If Trim(EHTML.innerText) = "Vehicle title:" Then
                    CollectedData.VehitcleTitle = EHTML.nextsibling.innerText
                End If
                If Trim(EHTML.innerText) = "Engine:" Then
                    CollectedData.Engine = EHTML.nextsibling.innerText
                End If
                If Trim(EHTML.innerText) = "Exterior color:" Then
                    CollectedData.ExtColor = EHTML.nextsibling.innerText
                End If
                If Trim(EHTML.innerText) = "Interior color:" Then
                    CollectedData.IntColor = EHTML.nextsibling.innerText
                    Exit For
                End If
                If Trim(EHTML.innerText) = "Transmission:" Then
                    CollectedData.Transmission = EHTML.nextsibling.innerText
                End If
            End If
        End If
    Next
    
    If InStr(1, WebBrowser.Document.body.innerHTML, "<TD class=""vi-vs-attr vi-vs-sI vi-vs-iA vi-vs-ti vi-vs-sF"">Accident or damage reported") > 1 Then
        CollectedData.Damaged = "Yes"
    Else
        CollectedData.Damaged = "No"
    End If
    
    Set itemElems = WebBrowser.Document.getElementsByTagName("H1") 'Was DIV
    For i = 0 To itemElems.Length
        Set EHTML = itemElems(i)
        If EHTML Is Nothing Then
        Else
            'If Left(EHTML.outerhtml, 24) = vbCrLf + "<DIV class=vi-it-itHd>" Then
            If Left(EHTML.outerhtml, 23) = vbCrLf + "<H1 class=vi-it-itHd>" Then
                Title = EHTML.innerText
                TitleParts = Split(Title, " ")
                CollectedData.Year = TitleParts(0)
                CollectedData.Make = TitleParts(1)
                Title = Replace(Title, TitleParts(0) + " " + TitleParts(1), "")
                CollectedData.Model = Trim(Title)
                Exit For
            End If
        End If
    Next

    GetItems = CollectedData
        
End Function



