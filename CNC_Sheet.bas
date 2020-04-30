Attribute VB_Name = "CNC_Sheet"
'Created by Jack Wilson, W '20

'DO NOT click into excel while generating if it is displayed
'Note: all comments beginning with triple ''' are unimplemented methods

'Variables for accessing Microsoft Excel
Dim Cell As Variant
Dim Excel As Object
Dim Sheet As Object
Dim Sheets As Object
Dim Workbook As Object
Dim Workbooks As Object

'General Variables
Dim UserName As String
Dim DownloadPath As String
Dim ParamHeading As Boolean
Dim CurrCol, CurrRow As Integer
Dim WindowWidth, WindowHeight As Integer
Dim ImgPath As String

'User Form Variables
Public UserFullName As String
Public UserClass As String
Public UserLab As String
Public UserProductName As String


' Initialize variables and excel file
Sub CATMain()

UserName = CATIA.SystemService.Environ("UserName")
DownloadPath = "C:\Users\" & UserName & "\Downloads\"

Dim Response As Integer
Response = MsgBox("Make sure the stock is visible. It will not be accessible" & _
        " throughout this program. Continue?", vbYesNo, "Warning")
If (Response = vbNo) Then End

'Setup Excel
Err.Clear
On Error Resume Next
Set Excel = GetObject(, "EXCEL.Application")
If Err.Number <> 0 Then
Err.Clear
Set Excel = CreateObject("EXCEL.Application")
End If
Excel.Application.Visible = False    'set to FALSE unless testing
Excel.Application.DisplayAlerts = False
Set Workbooks = Excel.Application.Workbooks
Set Workbook = Workbooks.Add
Set Sheets = Workbook.Worksheets(1)
Set Sheet = Workbook.Sheets(1)

MachiningDocument
Excel.Sheets("Sheet1").Delete
Excel.Application.Visible = True
Excel.Application.DisplayAlerts = True
Set Workbooks = Nothing
Set Workbook = Nothing
Set Sheets = Nothing
Set Excel = Nothing

End Sub

' Takes screen shot of active window in current orientation
Sub AddPhoto(MfgActivity)

Dim myWindow As Window
Dim myViewer As Viewer
Dim myOldLayout As CatSpecsAndGeomWindowLayout
Dim WindowW, WindowH As Integer
Dim Highlighting As Boolean
Dim VisualSettings As Object
Dim RenderMode

' Disable object highlighting
Set VisualSettings = CATIA.SettingControllers.Item("CATVizVisualizationSettingCtrl")
Highlighting = VisualSettings.PreSelectionMode
VisualSettings.PreSelectionMode = False

' Have user select the setup
SelectView.Label1.Caption = "Select " & MfgActivity & " and orient to the desired view. " & _
            "Make sure nothing is highlighted."
SelectView.Show
Do Until SelectView.Visible = False
DoEvents
Loop

On Error Resume Next

Err.Clear
Set myWindow = CATIA.ActiveWindow

Dim WinErr As Integer
WinErr = Err.Number

If (WinErr = 0) Then

myOldLayout = myWindow.Layout
WindowW = myWindow.Width
WindowH = myWindow.Height
WindowWidth = WindowW
WindowHeight = WindowH

myWindow.Width = 400
myWindow.Height = 300
myWindow.Layout = 1

Dim myViewer2
Set myViewer = myWindow.ActiveViewer
Set myViewer2 = myViewer

Dim BGcolor(3)
myViewer2.GetBackgroundColor BGcolor
RenderMode = myViewer2.RenderingMode
myViewer2.RenderingMode = catRenderQuickHiddenLinesRemovalWithHalfSmoothEdgeWithoutVertices

myViewer2.PutBackgroundColor Array(1, 1, 1)
'myViewer2.Reframe 'reframes the view to see everything

Dim PicName As String
PicName = "DocFabView.jpg"
ImgPath = DownloadPath & PicName
myViewer2.CaptureToFile 5, ImgPath

myViewer2.PutBackgroundColor BGcolor
myViewer2.RenderingMode = RenderMode

Dim ViewerW, ViewerH As Integer
ViewerW = myViewer.Width
ViewerH = myViewer.Height

myWindow.Width = WindowW
myWindow.Height = WindowH
myWindow.Layout = myOldLayout

Set myViewer = Nothing
Set myViewer2 = Nothing
Set myWindow = Nothing

End If

'Excel.ActiveSheet.Pictures.Insert DownloadPath & PicName 'adds picture by reference
'Excel.ActiveSheet.Shapes.AddPicture ImgPath, False, True, 0, 0, WindowW / 3, WindowH / 3
VisualSettings.PreSelectionMode = Highlighting

'Kill ImgPath

End Sub

' Adds a new excel named sheet
Sub AddNewSheet(Name)
Workbook.Sheets.Add(After:=Workbook.Sheets(Workbook.Sheets.Count)).Name = Name
Excel.ActiveSheet.PageSetup.Orientation = 2 ' landscape (significantly slows run time)
End Sub

' Formats the infomation sheets
Sub FormatInfoSheet(CurrRow, CurrCol, ProgNb)
Dim Col As Integer
Dim Row As Integer

Row = 2
Col = 2

' Insert header
With Excel.ActiveSheet.PageSetup
    .DifferentFirstPageHeaderFooter = True
    .RightFooter = "&R&A"
    
    With .FirstPage
        .LeftHeader.Text = "Generated &D &T"
        .CenterHeader.Text = "&16" & UserProductName
        .RightHeader.Text = "&R" & UserFullName & vbCr _
                        & UserClass & vbCr _
                        & UserLab
        .RightFooter.Text = "&R&A"
    End With
End With

'Excel.Range("B1:K1").Merge
'Excel.Range("B1:K1").HorizontalAlignment = -4108
With Excel
    .Range("B:L").HorizontalAlignment = -4108
    .Cells(Col, Row) = "Op #"
    .Cells(Col, Row + 1) = "Activity"
    .Cells(Col, Row + 2) = "Tool #"
    .Cells(Col, Row + 3) = "Feedrate (IPM)"
    .Cells(Col, Row + 4) = "Spindle Speed (RPM)"
    .Cells(Col, Row + 5) = "Approach Feed (IPM)"
    .Cells(Col, Row + 6) = "Retract Feed (IPM)"
    .Cells(Col, Row + 7) = "Finishing Feed (IPM)"
    .Cells(Col, Row + 8) = "Machining Time"
    .Cells(Col, Row + 9) = "Total Time"
    
    With .ActiveSheet.Rows(2)
        .RowHeight = .RowHeight * 2
        .WrapText = True
    End With
    
    .ActiveSheet.ListObjects.Add(1, Excel.Range("$B$2:$K$" & CurrRow), _
                    , 1, , "TableStyleLight18").Name = "Table" & ProgNb
    .Range("Table" & ProgNb & "[#All]").AutoFilter
    .ActiveSheet.ListObjects("Table" & ProgNb).Unlist
    .Range("C3:C500").HorizontalAlignment = -4131
    
    For i = 3 To CurrRow
        If (.Cells(i, 2) = "") Then
            .Range("E" & i & ":H" & i).Merge
            .Cells(i, 9) = "Tool Stickout: "
            .Cells(i, 9).Font.Bold = True
            .Range("I" & i & ":K" & i).Merge
            .Range("E" & i & ":K" & i).HorizontalAlignment = -4131
        End If
    Next
    
    .Cells.ColumnWidth = 10
    .Cells.EntireColumn.AutoFit 'autofit all columns
    .Range("A:A").Delete
End With

' Add border to machining and total times
Dim TimeRange As Object
Set TimeRange = Excel.Range("I" & CurrRow + 1 & ":J" & CurrRow + 1)
With TimeRange.Borders(7)
  .LineStyle = 1
  .ColorIndex = 0
  .ThemeColor = 1
  .TintAndShade = -0.349986266670736
  .Weight = 2
End With
With TimeRange.Borders(8)
  .LineStyle = 1
  .ColorIndex = 0
  .ThemeColor = 1
  .TintAndShade = -0.349986266670736
  .Weight = 2
End With
With TimeRange.Borders(9)
  .LineStyle = 1
  .ColorIndex = 0
  .ThemeColor = 1
  .TintAndShade = -0.349986266670736
  .Weight = 2
End With
With TimeRange.Borders(10)
  .LineStyle = 1
  .ColorIndex = 0
  .ThemeColor = 1
  .TintAndShade = -0.349986266670736
  .Weight = 2
End With
With TimeRange.Borders(11)
  .LineStyle = 1
  .ColorIndex = 0
  .ThemeColor = 1
  .TintAndShade = -0.349986266670736
  .Weight = 2
End With
With TimeRange.Borders(12)
  .LineStyle = 1
  .ColorIndex = 0
  .ThemeColor = 1
  .TintAndShade = -0.349986266670736
  .Weight = 2
End With

' add the image
Dim ImgWidth, ImgHeight, ImgFromTop, ImgFromLeft As Integer

' img size -> size modifier
ImgWidth = WindowWidth * 0.2
ImgHeight = WindowHeight * 0.2

ImgFromTop = (CurrRow + 2.5) * Excel.Rows(CurrRow).RowHeight
If ((ImgFromTop Mod 510) > (510 - ImgHeight)) Then
    ImgFromTop = Round(ImgFromTop / 510) * 510
End If

ImgFromLeft = (Excel.Range("A1:K1").Width - ImgWidth) / 2

Excel.ActiveSheet.Shapes.AddPicture ImgPath, False, True, ImgFromLeft, ImgFromTop, _
        ImgWidth, ImgHeight
End Sub

' Start creating the document
Sub MachiningDocument()

Dim MfgDoc1 As Document
Set MfgDoc1 = CATIA.ActiveDocument

Dim ProgramList As MfgActivities
Dim ActivityList As MfgActivities
Dim NumberOfProgram As Integer
Dim NumberOfActivity As Integer
Dim i As Integer
Dim J As Integer
Dim K As Integer
Dim SetupName As String
Dim ProgramName As String
Dim ActivityName As String
Dim CurrentSetup As ManufacturingActivity
Dim CurrentProgram As ManufacturingActivity
Dim CurrentActivity As ManufacturingActivity
Dim CurrentTool As ManufacturingTool
Dim ActivityType As String

Dim CurrentAssembly As ManufacturingToolAssembly
Dim AssemblyNumber As Long

Dim childs As Activities
Dim child As Activity
Dim quantity As Integer
Dim aProcess As AnyObject

Dim ToolNumber As Long
Dim ToolName As Variant

Dim CurrAct As Integer

Dim error As Integer

Dim SetupNb As Integer
Dim ProgramNb As Integer
Dim OpNb As Integer

Set aProcess = MfgDoc1.GetItem("Process") 'was RootActivityName
SetupNb = 1
ProgramNb = 1
OpNb = 1

UserInfo.Show 'Gather user information

' Scan the process
quantity = 0

If (aProcess.IsSubTypeOf("PhysicalActivity")) Then
    Set childs = aProcess.ChildrenActivities
    quantity = childs.Count
    
    If quantity <= 0 Then
        Exit Sub
    End If
    
    For i = 1 To quantity
      Set child = childs.Item(i)
      
      ' Gather and add info for each manufacturing setup
      If (child.IsSubTypeOf("ManufacturingSetup")) Then
          
          'AddNewSheet "Setup " & SetupNb
          AddPhoto child.Name
          SetupNb = SetupNb + 1

          For toolnb = 1 To MaxToolNb
            '''TabToolName(toolnb) = ""
          Next

          CurrPONb = CurrPONb + 1 'Current part op number
          CurrAct = 0

          Set CurrentSetup = child

          '''CreatePartOperationSheet CurrentSetup, CurrPONb, MfgDoc1.Name

          SetupName = CurrentSetup.Name

          ' Read the Programs of the current Setup
          Set ProgramList = CurrentSetup.Programs
          NumberOfProgram = ProgramList.Count

          ' Gather and add info for each manufacturing program
          For ProgNb = 1 To NumberOfProgram
          
            AddNewSheet ProgramList.GetElement(ProgNb).Name
            'Excel.Range("B1:K1") = UserName & " Setup " & ProgramNb
            CurrCol = 1
            CurrRow = 2
            ParamHeading = False

            ProgCuttingtime = 0
            ProgTotalTime = 0
          
            Set CurrentProgram = ProgramList.GetElement(ProgNb)

            '''WriteProgFileHeader CurrentSetup, CurrPONb, CurrentProgram, aProgStream, MfgDoc1.Name
            
            ProgramName = CurrentProgram.Name
            
            ' Read the Activities of the current Program
            Set ActivityList = CurrentProgram.Activities
            NumberOfActivity = ActivityList.Count

            On Error Resume Next

            ' Gather and add info for each manufacturing activity
            For ActNb = 1 To NumberOfActivity
              CurrAct = CurrAct + 1
              Set CurrentActivity = ActivityList.GetElement(ActNb)
              ActivityName = CurrentActivity.Name
              ActivityType = CurrentActivity.Type

              ' This block may not be necessary (tool info)
              If (ActivityType <> "ToolChange" And ActivityType <> "ToolChangeLathe") Then
                Dim currcor, ncorr, ncorrl, diam As Variant
                Dim corrtype As String
                Dim pAtt As Parameter
                Err.Clear
                Set pAtt = CurrentActivity.GetAttribute("CompNumber")
                error = Err.Number

                If (error = 0) Then
                  ncorr = pAtt.Value
                  error = Err.Number

                  If (error = 0) Then
                    Set CurrentTool = CurrentActivity.Tool
                    ToolNumber = CurrentTool.ToolNumber

                    Dim NbCorr As Integer
                    NbCorr = CurrentTool.CorrectorCount

                    If (NbCorr > 0) Then
 
                      Dim aCorr As ManufacturingToolCorrector
                      Dim CorrNumber, CorrLengthNumber As Integer
                      Dim CorrDiameter As Variant

                      currcor = 0
                      Do
                        currcor = currcor + 1
                        Set aCorr = CurrentTool.GetCorrector(currcor)
                        CorrNumber = aCorr.Number
                        If (CorrNumber = ncorr) Then
                          corrtype = aCorr.Point
                          ncorrl = aCorr.LengthNumber
                          diam = aCorr.Diameter
                          ToolType = CurrentTool.ToolType
                          ToolName = CurrentTool.Name
                          '''MajTabOutils ToolNumber, ncorr, ncorrl, diam, corrtype, ToolType, ToolName
                          Exit Do
                        End If
                        If (currcor = NbCorr) Then Exit Do
                      Loop
                    End If
                  End If
                End If
              End If
              
              CurrRow = CurrRow + 1
              CurrCol = 2
              
              With Excel
                If (CurrentActivity.Type = "ToolChange") Then
                    .Cells(CurrRow, 3) = CurrentActivity.Name
                    .Cells(CurrRow, 3).Font.Bold = True
                Else
                    .Cells(CurrRow, 3) = "  " & CurrentActivity.Name
                    .Cells(CurrRow, 2) = OpNb
                    OpNb = OpNb + 1
                End If
                
                .Cells(CurrRow, 4) = CurrentActivity.Tool.ToolNumber
                
                If (CurrentActivity.Type = "ToolChange") Then
                    .Cells(CurrRow, 5) = "Tool Desc: " & CurrentActivity.Tool.Name
                    .Cells(CurrRow, 5).Characters(Start:=1, Length:=10).Font.Bold = True
                End If
              End With
              
              WriteOperationParameters CurrentActivity

              '''CreateFicheOpe CurrentActivity, CurrPONb, CurrAct, MfgDoc1.Name, SetupName, ProgramName
              '''CreateOpeSum CurrentActivity, CurrPONb, CurrAct
              '''AddOpeSum CurrPONb, CurrAct, aProgStream

              Err.Clear
              aTime = CurrentActivity.MachiningTime
              If (Err.Number = 0) Then ProgCuttingtime = ProgCuttingtime + aTime
              Err.Clear
              aTime = CurrentActivity.TotalTime
              If (Err.Number = 0) Then ProgTotalTime = ProgTotalTime + aTime

            Next

            '''CreateToolList MfgDoc1.Name, CurrentSetup, CurrPONb

            ' Adding the machining time and the total time
            'Excel.Cells(1, 1) = "Program cutting time"
            Excel.Cells(CurrRow + 1, 10) = ToHMS(ProgCuttingtime)
            'Excel.Cells(2, 1) = "Program total time"
            Excel.Cells(CurrRow + 1, 11) = ToHMS(ProgTotalTime)
            
            FormatInfoSheet CurrRow, CurrCol, ProgramNb
            ProgramNb = ProgramNb + 1
          Next
          
        
          '''CompleteToolCompSheets CurrPONb, TabOutils, MaxTabOutils
          Erase TabOutils
          MaxTabOutils = -1

          Kill ImgPath
      End If
    
    Next

    '''CreateToolAssocOpe childs

    '''CreateAssemblyAssocOpe childs

  End If

End Sub


' Transforming decimal minutes into hours minutes seconds.
Function ToHMS(aTime)

  Dim Result As String
  Result = ""

  Err.Clear

  On Error Resume Next
  wtime = CDbl(aTime)
  If (Err.Number <> 0) Then
    ToHMS = "No computed time"
    Exit Function
  End If

  If (wtime >= 0) Then

    nbhours = Int(wtime / 60)
    If (nbhours > 0) Then
      Result = nbhours & "h"
      wtime = wtime - nbhours * 60
    End If

    nbminutes = Int(wtime)
    If (nbminutes > 0) Then
      If (Result <> "") Then Result = Result & " "
      Result = Result & nbminutes & "'"
      wtime = wtime - nbminutes
    End If

    If (wtime > 0) Then
      If (Result <> "") Then Result = Result & " "
      Result = Result & Round(wtime * 60) & "''"
    End If

  End If

  If (Result = "") Then Result = "No computed time"

  ToHMS = Result

End Function

' Writes the relevant parameters for a pp instruction
Sub WritePPInstruction(aPPInstr)

  Dim ligne As Variant
  Dim i, nbcar As Integer
  nbcar = Len(aPPInstr)

  For i = 1 To nbcar
    carac = Mid(aPPInstr, i, 1)
    If (carac = Chr(13) Or carac = Chr(10)) Then
      If (ligne <> "") Then
        'aStream.Write ligne & "<br>" & EOL
        MsgBox ligne
        ligne = ""
      End If
    Else
      ligne = ligne & carac
    End If
  Next

  'If (ligne <> "") Then aStream.Write ligne & "<br>" & EOL

End Sub

' Write the operation parameters
Sub WriteOperationParameters(anOpe)

  OpeType = anOpe.Type
  Select Case OpeType
  Case "ToolChange"
    WriteParameters anOpe
  Case "ToolChangeLathe"
    WriteParameters anOpe
  Case "TableHeadRotation"
    WriteParameters anOpe
  Case "CoordinateSystem"
    WriteParameters anOpe
  Case "PPInstruction"
    WriteParameters anOpe
  Case Else
    WriteCycleParameters anOpe
  End Select
End Sub

' Write the cycle parameters
Sub WriteCycleParameters(anOpe)

  '''AddCycleStrategyParameters anOpe

  'write:
  AddCycleFeedrateParameters anOpe

  'write:
  AddCycleMachiningTime anOpe

End Sub

' Gather and write operation parameters
Sub WriteParameters(anOpe)
  Dim error As Integer
  Dim pPPInstr As Parameter
  Dim PPInstr As String
  PPInstr = ""

  On Error Resume Next

  Err.Clear
  Set pPPInstr = anOpe.GetAttribute("MFG_PPWORDS")
  error = Err.Number

  If (error = 0) Then
    PPInstr = pPPInstr.Value
  Else
    Err.Clear
  End If

    If (Len(PPInstr) > 0) Then
    'write: PP Instruction:
    WritePPInstruction PPInstr
  Else
    'write: PP Instruction:
    'write: Initialize from PP words table
  End If
End Sub

' Add cycle feedrate parameters
Sub AddCycleFeedrateParameters(aCycle)

  Dim error As Integer
  Dim TabAtt()
  Dim att As Integer
  Dim nbatt As Integer

  'write: Feedrate

  On Error Resume Next
  Err.Clear

  nbatt = aCycle.NumberOfFeedrateAttributes
  error = Err.Number
  If (error <> 0) Then nbatt = 0

  If (nbatt > 0) Then


    ReDim TabAtt(nbatt)
        aCycle.GetListOfFeedrateAttributes (TabAtt)

        For att = 0 To nbatt - 1

          AddParameterToFeedrateTable aCycle, TabAtt(att), False
          
        Next

  Else
    'write No Parameter
  End If

End Sub

' Adds a feedrate parameter
Sub AddParameterToFeedrateTable(anObj, aParam, AcceptComment)

  Dim error As Integer

  If (Not AcceptComment And aParam = "MFG_COMMENT") Then Exit Sub

  Dim anAttribute As AnyObject
  Dim AttrVal As String
  Dim anObjName As String
  Dim ColNb As Integer
  
  ColNb = 0

  On Error Resume Next
  Err.Clear
  Set anAttribute = anObj.GetAttribute(aParam)
  error = Err.Number

  If (error = 0) Then
    AttrVal = anAttribute.ValueAsString
    error = Err.Number
    If (error = 0 And AttrVal <> "") Then
        anObjName = ToNLS(anObj, aParam)
        
        'Case to test for wanted parameters
        Select Case anObjName
            Case "Machining feedrate"
                ColNb = 5
            Case "Machining spindle"
                ColNb = 6
            Case "Approach feedrate"
                ColNb = 7
            Case "Retract feedrate"
                ColNb = 8
            Case "Finishing feedrate"
                ColNb = 9
        End Select
        
        If (ColNb <> 0) Then
            Excel.Cells(CurrRow, ColNb) = Extract_Number_from_Text(ToNLS(anObj, AttrVal))
        End If
      CurrCol = CurrCol + 1 'useless now
    End If
  End If

End Sub

' Add cycle machining time parameters
Sub AddCycleMachiningTime(aCycle)

  'write: Machining Time

  On Error Resume Next
  
  '"Cutting time"
  Excel.Cells(CurrRow, 10) = ToHMS(aCycle.MachiningTime)
  '"Total time"
  Excel.Cells(CurrRow, 11) = ToHMS(aCycle.TotalTime)

End Sub

' returns attribute name or value
Function ToNLS(anObj, aParameterName)
  Dim error As Integer
  Dim NLSresult As String
  On Error Resume Next
  Err.Clear
  NLSresult = anObj.GetAttributeNLSName(aParameterName)
  error = Err.Number
  If (error <> 0 Or NLSresult = "") Then NLSresult = aParameterName
  ToNLS = NLSresult
End Function

Function Extract_Number_from_Text(Phrase As String) As Double

Dim Length_of_String As Integer
Dim Current_Pos As Integer
Dim Temp As String

Length_of_String = Len(Phrase)
Temp = ""

For Current_Pos = 1 To Length_of_String
    If (Mid(Phrase, Current_Pos, 1) = "-") Then
      Temp = Temp & Mid(Phrase, Current_Pos, 1)
    End If
    If (Mid(Phrase, Current_Pos, 1) = ".") Then
     Temp = Temp & Mid(Phrase, Current_Pos, 1)
    End If
    If (IsNumeric(Mid(Phrase, Current_Pos, 1))) = True Then
        Temp = Temp & Mid(Phrase, Current_Pos, 1)
     End If
Next Current_Pos

If Len(Temp) = 0 Then
    Extract_Number_from_Text = 0
Else
    Extract_Number_from_Text = CDbl(Temp)
End If

End Function


