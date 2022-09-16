Sub auto_open()

 
    With Application
        .EnableEvents = False
        .ScreenUpdating = False


    Workbooks.Open "xxxxxxxxxxxxxxxxxxxxxxxxx\file_1.csv"
    Windows("file_1.csv").Activate
    Range("A2:AF20000").Select
    Range("A2").Activate
    Selection.Copy
    Windows("example.xlsm").Activate
    Sheet1.Select
    Range("D2").Select
    Sheet1.Paste
    Windows("file_1.csv").Application.CutCopyMode = False
    Windows("file_1.csv").Close
    
    
    Workbooks.Open "xxxxxxxxxxxxxxxxxxxxxxx\file_2.csv"
    Windows("file_2.csv").Activate
    Range("A2:D20000").Select
    Range("A2").Activate
    Selection.Copy
    Windows("example.xlsm").Activate
    Sheet1.Select
    Range("AM2").Select
    Sheet1.Paste
    Windows("file_2.csv").Application.CutCopyMode = False
    Windows("file_2.csv").Close
    

    Workbooks.Open "xxxxxxxxxxxxxxxxxxxxx\file_a.csv"
    Windows("file_a.csv").Activate
    Range("A2:D20000").Select
    Range("A2").Activate
    Selection.Copy
    Windows("example.xlsm").Activate
    Sheet1.Select
    Range("AS2").Select
    Sheet1.Paste
    Windows("file_a.csv").Application.CutCopyMode = False
    Windows("file_a.csv").Close
    
    
    Workbooks.Open "xxxxxxxxxxxxxxxxxx\file_b.csv"
    Windows("file_b.csv").Activate
    Range("A2:D20000").Select
    Range("A2").Activate
    Selection.Copy
    Windows("example.xlsm").Activate
    Sheet1.Select
    Range("AY2").Select
    Sheet1.Paste
    Windows("file_b.csv").Application.CutCopyMode = False
    Windows("file_b.csv").Close
    
    
 
    
    Workbooks.Open "xxxxxxxxxxxxxxxxxxx\file_4.csv"
    Windows("file_4.csv").Activate
    Range("A2:K20000").Select
    Range("A2").Activate
    Selection.Copy
    Windows("example.xlsm").Activate
    Sheet1.Select
    Range("BR2").Select
    Sheet1.Paste
    Windows("file_4.csv").Application.CutCopyMode = False
    Windows("file_4.csv").Close

    ActiveWorkbook.RefreshAll 'make sure the refresh in bg property is false for all connections

    For Each ws In ActiveWorkbook.Worksheets
    For Each pt In ws.PivotTables
            pt.RefreshTable
        Next pt
    Next ws
    
    Dim newDate: newDate = Format(DateAdd("M", -1, Now), "MMMM")
    
    Application.AlertBeforeOverwriting = False
    Application.DisplayAlerts = False

    ActiveWorkbook.SaveAs "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx " & newDate, 51



    
    ActiveWorkbook.RefreshAll 

    For Each ws In ActiveWorkbook.Worksheets
    For Each pt In ws.PivotTables
            pt.RefreshTable
        Next pt
    Next ws
    
      
    Application.AlertBeforeOverwriting = False
    Application.DisplayAlerts = False

    
    strbody1 = "Hello Team,<br><br>" & _
                "Hope this email finds you well.<br><br>" & _
            "Please find attached xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx </b><br>" & _
            "Let me know if you will have any questions or concerns<br><br>" & _
            "Kind regards,<br>" & _
            "Lukasz Siadkowski"
            
   
    
  
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    On Error Resume Next
            
    W = Application.WorksheetFunction.WeekNum(Now())
    Y = Year(Now())
        
    'email body:
    Body = strbody1
    
    Attached_File = ("xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx " & newDate & ".xlsx")

    With OutMail
        .To = "xxxxxxxx"
        .CC = ""
        .BCC = ""
        .Subject = "xxxxxxxxxxxxxxxx " & newDate
        .HTMLBody = Body
        .SentOnBehalfOfName = "xxxxxxxxx"
        '.Display
        .Attachments.Add Attached_File
        .Send
    End With
    
    On Error GoTo 0


    Set OutMail = Nothing
    Set OutApp = Nothing
    
    End With
    
    Application.Quit
    
End Sub


Function RangetoHTML(rng As Range)

    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With

    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         Filename:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.readall
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    TempWB.Close savechanges:=False

    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
End Function
