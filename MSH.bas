Attribute VB_Name = "Module1"
Option Explicit
Option Compare Binary

Global Const MS_Total As Long = 81

Global Const C_Button_Normal        As Long = &H80
Global Const C_Button_NormalText    As Long = vbWhite
Global Const C_Button_Select        As Long = vbRed
Global Const C_Button_SelectText    As Long = vbBlack
Global Const C_Button_Clicked       As Long = vbRed

Global ProjectFile As String
Global Filename As String
Global MaliciousCount As Long
Global InUse As Boolean
Global EndScan As Boolean
Global ProgressLight As Long

Global DirDialogFlag As String

Type ErrorData
    Def As String
    File As String
End Type

Type Options_Settings
    Report_LogCreateType As Byte 'to tell report generator what type of
                      'log it should make
    Report_AutoCreateLog As Boolean 'instead of clicking 'save scan report'
                         'the program will automatically create one by
                         'using the default filename
    Report_OverWriteLog As Boolean 'If it is allowed to overwrite an existing log
    Report_DefaultLogFilename As String 'Obvious
    Scan_ExcludePOTENTIALClass As Boolean 'Excludes all POTENTIAL Strings
    Scan_ExcludeSUSPICIOUSClass As Boolean 'Excludes all Suspicious Strings
'Everything won't get those classes information if their set true
    ShowSplash As Byte

    Advanced_DisableChecking As Boolean 'Bypass the context checking
    Advanced_CheckComments As Boolean   'Allow comment checking
    Advanced_EraseDestructive As Boolean 'NOT AVAILABLE YET - will delete Destructive Code

    Scan_CommentOut As Boolean
    GUI_AllowSorting As Boolean

    AutoSearch_DefaultDir As String
    AutoSearch_TypeOfSearch As Byte
    AutoSearch_IncludeSubs As Byte
    AutoSearch_AutoScanAfter As Byte
End Type

Global MSH_Settings As Options_Settings

Global GetErrors() As ErrorData 'In case my scan fails to work on something
'so the report can view it

Type Report_Info
    TotalMaliciousCount As Long
    TOTALLines As Long
    TOTALBYTES As Single 'for debugging purposes of checking speed
    POTENTIAL As Long
    SUSPICIOUS As Long
    CAUTION As Long
    WARNING As Long
    DANGER As Long
    DESTRUCTIVE As Long
    MSOccur(MS_Total) As Long
End Type

Public ReportI As Report_Info

Public Type Malicious_Info
    MStr As String
    Class As String
End Type

Type FileToScan_Info
    Pathname As String
    Filename As String
End Type

Public FilesToScan() As FileToScan_Info
Public SubFiles() As String 'Files which are found inside a project
'not programmed into the file to scan at start
Public MS(MS_Total) As Malicious_Info
Public CodeLine() As String

Public Sub LoadFile(ByVal Path As String, ByVal File As String)
Dim MaxLine As Long 'To Manage the CodeLine Ubound easier
Dim MaxFile As Long 'To Manage the Files Ubound easier
Dim c As Long
Dim TCL As String 'Ucase Temp for CodeLine
Dim FRXLoaded As Boolean 'If this file is a form and
'it's FRX binary file is loaded

On Error GoTo Err_Check
MaliciousCount = 0
ReDim CodeLine(0)
Close #1
Open Path & File For Input As #1
Form1.File_Label = File
Form1.FileType_Label = GetFileType(File)
Do Until EOF(1) Or EndScan = True

ReDim Preserve CodeLine(UBound(CodeLine) + 1)
Line Input #1, CodeLine(UBound(CodeLine))
MaxLine = UBound(CodeLine)
TCL = UCase(CodeLine(MaxLine))

If Right$(File, 4) = ".FRM" And FRXLoaded = False Then
c = InStr(1, TCL, Left$(File, Len(File) - 4) & ".FRX") 'Searchs for
'Form Binary File Extensions
    If c > 0 Then
        ReDim Preserve SubFiles(UBound(SubFiles) + 1)
        SubFiles(UBound(SubFiles)) = Left$(File, Len(File) - 4) & ".FRX"
        'Adds the file to the list
        FRXLoaded = True
    End If
End If

ScanCodeLine TCL, UBound(CodeLine) ' this is the line which does the scan
Loop
Close #1

Exit Sub
Err_Check:
AddErrorMsg Err.Description, "LoadFile: " & File
Exit Sub
End Sub

Public Sub GetVBPFiles(ByVal ProjectFile As String)
Dim i As Long
On Error GoTo Err_Check
'this allows you to just give the project file and my program will handle the rest
'not bad eh?
'However if you do select a form file to scan it will bypass this and that form file
'only is checked and this applies to all other VB files, hell you can even use it
'on documents and vbscript files if you wanted to.
'Next Version will extend to cover VB Group Project Files as well

'the data can be both lower or upper case as it converted to upper case anyway

Dim MaxLine As Long 'To Manage the CodeLine Ubound easier
Dim MaxFile As Long 'To Manage the Files Ubound easier
Dim TCL As String 'Temp for CodeLine Upper-cased
Dim TSF As String 'Temp for Sub File

Dim x As Long 'type of VB file to add
Dim c As Long 'colon ; location if field is Class or Module, because they
'store their object name first then their file name
'or it's used as a instr locater too.

ReDim SubFiles(0)
ReDim CodeLine(0)

Open ProjectFile For Input As #1
    Do Until EOF(1)
    
    ReDim Preserve CodeLine(UBound(CodeLine) + 1)
    
    MaxLine = UBound(CodeLine)
    
    Line Input #1, CodeLine(MaxLine)
    TCL = UCase(CodeLine(MaxLine))
    'This is a optimization trick, don't ever use ucase more than once
    'instead store it in a non-array string so the code that needs this
    'info will be able to load it faster
    '(Using arrays in InStr slows it down a bit)

    x = 0
    If InStr(1, TCL, "FORM=") = 1 Then x = 1
    If InStr(1, TCL, "MODULE=") = 1 Then x = 2
    If InStr(1, TCL, "CLASS=") = 1 Then x = 3
    If InStr(1, TCL, "RESFILE32=") = 1 Then x = 4

        If x > 0 Then

        Do While c > 0
            c = InStr(1, TCL, Chr(34))
            If c = 0 Then Exit Do
            Mid(TCL, c, 1) = " "
        Loop

    ReDim Preserve SubFiles(UBound(SubFiles) + 1)
    MaxFile = UBound(SubFiles)
    If x = 1 Then TSF = Mid$(TCL, 6, Len(TCL) - 5)
    If x = 4 Then TSF = Mid$(TCL, 11, Len(TCL) - 10)
        If x = 2 Or x = 3 Then
            c = InStr(1, TCL, "; ")
                If c > 0 Then
                TSF = Mid$(TCL, (c + 2), Len(TCL) - (c + 1))
                End If
        End If
    End If
    SubFiles(MaxFile) = Trim$(TSF)
    Loop
    Close #1
   
Exit Sub
Err_Check:
AddErrorMsg Err.Description, ProjectFile
Exit Sub
End Sub


Private Function AddToList(ByVal MSNum As Long, ByVal CLNum As Long)
Dim x As Long
Dim IconNum As Long

If MSH_Settings.Scan_ExcludePOTENTIALClass = True And MS(MSNum).Class = "POTENTIAL" Or _
   MSH_Settings.Scan_ExcludeSUSPICIOUSClass = True And MS(MSNum).Class = "SUSPICIOUS" Then Exit Function
'To make it more efficient and not look for possible malicious strings
'you can exclude two of the six classes so only caution or higher classed
'are reported.

'IF...Elseif...End if are said to be 3% faster than select case
'I think because it streams for the right answer rather than
'look for it straight away.
'Use Else statement for the equilivent of Case Else
If MS(MSNum).Class = "POTENTIAL" Then
x = vbBlack: IconNum = 7
ElseIf MS(MSNum).Class = "SUSPICIOUS" Then
x = vbBlack: IconNum = 7
ElseIf MS(MSNum).Class = "CAUTION" Then
x = 64: IconNum = 2
ElseIf MS(MSNum).Class = "WARNING" Then
x = 128: IconNum = 3
ElseIf MS(MSNum).Class = "DANGER" Then
x = 192: IconNum = 4
ElseIf MS(MSNum).Class = "DESTRUCTIVE" Then
x = 255: IconNum = 4
End If

ReportI.MSOccur(MSNum) = ReportI.MSOccur(MSNum) + 1

With Form1
    .MSL.ListItems.Add , , Filename
    .MSL.ListItems.Item(.MSL.ListItems.Count).SmallIcon = IconNum
    .MSL.ListItems.Item(.MSL.ListItems.Count).SubItems(1) = MS(MSNum).MStr
    .MSL.ListItems.Item(.MSL.ListItems.Count).SubItems(2) = CLNum & " " & CodeLine(CLNum)
    .MSL.ListItems.Item(.MSL.ListItems.Count).SubItems(3) = MS(MSNum).Class
    .MSL.ListItems.Item(.MSL.ListItems.Count).ListSubItems. _
    Item(1).ForeColor = x
    .MSL.ListItems.Item(.MSL.ListItems.Count).ListSubItems. _
    Item(2).ForeColor = x
    .MSL.ListItems.Item(.MSL.ListItems.Count).ListSubItems. _
    Item(3).ForeColor = x
End With

        MaliciousCount = MaliciousCount + 1
        ReportI.TotalMaliciousCount = ReportI.TotalMaliciousCount + 1
        
        'I presume a select case with a if...elseif...end if
        'would optimize speed of declaring what level it is
        If MS(MSNum).Class = "POTENTIAL" Then
            ReportI.POTENTIAL = ReportI.POTENTIAL + 1
        ElseIf MS(MSNum).Class = "SUSPICIOUS" Then
            ReportI.SUSPICIOUS = ReportI.SUSPICIOUS + 1
        ElseIf MS(MSNum).Class = "CAUTION" Then
            ReportI.CAUTION = ReportI.CAUTION + 1
        ElseIf MS(MSNum).Class = "WARNING" Then
            ReportI.WARNING = ReportI.WARNING + 1
        ElseIf MS(MSNum).Class = "DANGER" Then
            ReportI.DANGER = ReportI.DANGER + 1
        ElseIf MS(MSNum).Class = "DESTRUCTIVE" Then
        ReportI.DESTRUCTIVE = ReportI.DESTRUCTIVE + 1
        End If

End Function

Public Sub ClearList()
Form1.MSL.ListItems.Clear
Form1.ProjectFiles.ListItems.Clear
ReportWindow.Report_MSL.ListItems.Clear

Form1.MSL.Refresh
Form1.ProjectFiles.Refresh
ReportWindow.Report_MSL.Refresh
Form1.CodeLine_Label = Empty
End Sub

Private Sub ScanCodeLine(ByVal TCL As String, ByVal Line_Number As Long)
On Error GoTo Err_Check
Dim i As Long

ReportI.TOTALLines = ReportI.TOTALLines + 1
For i = 0 To MS_Total&
'Check Load_MSData to see why UMS = Ucase(Ms(i).mstr) isn't here anymore
'by making that function ucase every string in the MS list there's no
'need to re-UCase them
        
        If TestContext(TCL, MS(i).MStr, _
        MSH_Settings.Advanced_CheckComments, _
        MSH_Settings.Advanced_DisableChecking) = True Then
            AddToList i, Line_Number
        End If
        
Next i

'String 52 Follows a different Context, so it's added here
If Line_Number = 1 And InStr(1, TCL, MS(52).MStr$) = 1 Then _
   AddToList i, Line_Number
   'Since all EXE's start with MZ (Well win32 Applications do)
   'The header must only be located at the first byte for it
   'to be malicious, or it could be found inside a string from above.

DoEvents

Exit Sub
Err_Check:
AddErrorMsg Err.Description, ProjectFile & " on ScanCodeLine"
Exit Sub
End Sub

Public Function Load_MSData()
Dim i As Long
'NOTE: I did this intentionally, so that when compiled it can't be edited by
'users as easily, there's 61 malicious/accessory strings listed below
'Danger: if modified or hacked some of these can be triggered by using
'Shell with the array variable data!!!
'.msh' is included in this list due to it's libraries containing
'dangerous code, so if other projects than this one use it, it could
'be a disguised attack on your PC if you have those libraries. :(
'This version (From the last version) I promised an Encryption method
'to make it so no other program can use these strings for malicious
'purposes without first decrypting them.

'Designed for Libary with Now 75 Strings! (62 Originally, 13 More)
Open App.Path & "\Reference1.MSH" For Binary Access Read Lock Read As #2
Get #2, , MS
Close #2

CryptReference

For i = 0 To MS_Total
MS(i).MStr = UCase(MS(i).MStr) 'Here's why the code in scancodeline
'is not there by ucasing all the strings beforehand then it's already
'made to compare
Next i
End Function

Private Sub CryptReference()
Dim i As Long
Dim j As Long
'It's a two way method the same function decrypts and encrypts depending
'if the data is encrypted or normal.
'Like the trick with app.CompanyName?
'Makes it easy to track down other programs using it.
'This doubles as a security feature, when compiled this would be impossible
'to access the library 100% without my company name in the exe.
'Feel free to use this encryption idea I don't care,
'it took me only 5 minutes to make (10 minutes to Debug however).
'Just add a few other things and it would make a good encryption method.

Dim DS As String
Dim MSC As String
Dim CSC As String

For i = 0 To UBound(MS)
DS = Empty
    For j = 1 To Len(MS(i).MStr)
    MSC = Mid(MS(i).MStr, j, 1)
    CSC = Mid(App.CompanyName, (i Mod Len(App.CompanyName) + 1), 1)
    DS = DS & Chr(Asc(MSC) Xor Asc(CSC))
    Next j
MS(i).MStr = DS
Next i
End Sub


Public Sub MakeLibrary()
'this function I only use, the code for this is on my machine I paste it here
'and use debug to execute it so a new resource file is created and then
'remove it leaving a easy place to paste it again otherwise my program
'integrity is unreliable with all those malicious/Accessory strings.
End Sub


Private Function GetFileType(ByVal File As String) As String

Dim UcaseFileExt As String
UcaseFileExt = UCase$(Right$(File, 4))

If UcaseFileExt = ".VBP" Then
GetFileType = "Visual Basic 6.0 Project"
ElseIf UcaseFileExt = ".FRM" Then
GetFileType = "Visual Basic 6.0 Form"
ElseIf UcaseFileExt = ".BAS" Then
GetFileType = "Visual Basic 6.0 Module"
ElseIf UcaseFileExt = ".CLS" Then
GetFileType = "Visual Basic 6.0 Class Module"
ElseIf UcaseFileExt = ".RES" Then
GetFileType = "Visual Basic 6.0 Resource File"
ElseIf UcaseFileExt = ".FRX" Then
GetFileType = "Visual Basic 6.0 Form Binary File"
ElseIf UcaseFileExt = ".TXT" Then
GetFileType = "Text Document"     'In case it's hidden in a text file
ElseIf UcaseFileExt = ".VBS" Then 'and is renamed into a script.
GetFileType = "VB Script" 'Real virus are made using this language variant
End If                    'of Visual Basic e.g 'Love Bug' so this also
                          'covers basic-level viruses too!
End Function

Public Sub TerminateApp()
'Thanks to Coding Genius article, I Did use a proper way to close this app
'(a more efficient way). (If you don't, I think unallocated memory wastes
'your RAM therefore slowing your machine down, but only a huge App could
'do that like a 3-D game, but it's highly reccommended for all apps)

'Deallocates and erases the dynamic array memory
Erase MS
Erase CodeLine
Erase FilesToScan
Erase SubFiles
Erase GetErrors
'Unloads the forms nuff said
Unload Form1
Unload ReportWindow
Unload Splash
Unload Options 'Just case of an error and it somehow forgot to unload
Unload Settings_Report 'Same as above
Unload Settings_Advanced 'Same as above too

'Finishs the app
'The mistake is that the 'END' Statement does NOT do the above for you
'like i though, those above are manually required to unload properly.
End
End Sub

Private Function TestContext(ByVal TestLine As String, ByVal TestWord As String, ByVal CheckComments As Boolean, ByVal DisableChecking As Boolean) As Boolean
'Copyright 2002 Roger Gilchrist
'This code was provided by this author on PSC, so you must obtain his
'permissions and copyrights of this code if you want to use it.
'Roger I modified it slightly to suit my program and for
'speed optimization, e.g ByVal makes
'it recieve strings by values not by reference making it faster and
'change select case statement to IF...THEN...ELSEIF...ENDIF Statement also
'for fractional speed.

'Roger's Notes (I thought I'll leave them in to help other users and
'understand what he's fixed/done):
'I liked 'Malicious String Hunter' a lot but found that it made some unnecessary
'warnings when
'I ran it against a very large project I have.
'EXAMPLES
'MZ was found in 'Function LenTrimZero',
'output in  'Function OutPutToList',
'.sys in 'SysForm.sysButton(4).Enabled'
'.dll in Private Declare Sub CoTaskMemFree Lib "ole32.dll" (byVal.... etc
'and several over-dramatically named routines with names like KillRow(it clears
'and removes a row in a MSGrid control)
'SO
'I created this routine and sent it to you
'I also rewrote Function ScanCodeLine as Sub ScanCodeLine (It doesn't return
'anything) and
'simplified it to use this Function.

  Dim TestChars As String 'String of characters which legitimately could delimit
'the TestWord
  Dim Pos As Long         'Position of TestWord in TestLine
  Dim TPos As Long        'File extention special case test
  Dim CommentTest As Long
  
  TestChars = " .*;':!?<>+-_=\()" & Chr(34)
  'Characters "\()A" were added by me James Beer, since strings
  'like format c:\ were not getting listed :(
  'Sorry for the inconvience this might of caused

    'The first two tests were originally in ScanCodeLine
    'I just moved them here for better encapsulation of tests
    '-----------------------------------------------------------
    'Test 1 TestWord is in TestLine
    Pos = InStr(1, TestLine, TestWord)
        If Pos = 0 Then
        Exit Function ' its not there so don't list it
    End If
    
If Not CheckComments = True Then
'I developed settings to disable test 2 or both 3 and 4 by going
'to General Settings then Advanced - James Beer
    '-----------------------------------------------------------
    'Test 2 is TestWord in a comment so not dangerous
    CommentTest = InStr(1, TestLine, "'")
    If CommentTest < Pos And CommentTest > 0 Then
        Exit Function ' its in a comment so don't list it
    End If

''Un-comment the next two lines and you have the original tests
End If

If DisableChecking = True Then
    TestContext = True
    Exit Function
End If

'-----------------------------------------------------------
''Below this line are my additions to the program
    'Test 3 Just in case the whole line is a dangerous word
    '(This would fail Test 4 as it has no Before/After characters)
    If TestLine = TestWord Then
        TestContext = True
        Exit Function ' the whole line is a suspicious word so include it
    End If
'-----------------------------------------------------------
    'sub-test: File extention is a special case
    TPos = Pos
    If Left$(TestWord, 1) = "." Then 'If TestWord is a file extention only check
    'that it
        TPos = 1                    'not embedded at the end of the word
    End If
    
    If InStr(1, TestLine, "EX") > 0 Then
        TestContext = True
        Exit Function
    End If
    
    If InStr(1, TestLine, "SHELL") > 0 Then
        TestContext = True
        Exit Function
    End If

'-----------------------------------------------------------
'Test 4 check surrounding characters (where they exist)
    If TPos = 1 Then 'Only test After TestWord
        TestContext = InStr(TestChars, Mid(TestLine, Pos + Len(TestWord), 1)) > 0
    ElseIf TPos = (Len(TestLine) - Len(TestWord)) Then 'Only test Before TestWord
        TestContext = InStr(TestChars, Mid(TestLine, Pos - 1, 1)) > 0
    Else 'Test Before and After TestWord
        TestContext = InStr(TestChars, Mid(TestLine, Pos - 1, 1)) > 0
        If TestContext Then 'second half of test only if first half is True
        TestContext = TestContext And InStr(TestChars, Mid(TestLine, Pos + Len(TestWord), 1)) > 0
        End If
    End If
End Function

Public Sub AddErrorMsg(ByVal ErrDesc As String, ByVal FileMsg As String)
'another function from Roger Gilchrist to manage the errors
    ReDim Preserve GetErrors(UBound(GetErrors) + 1)
    GetErrors(UBound(GetErrors)).Def = ErrDesc & " "
    GetErrors(UBound(GetErrors)).File = FileMsg
End Sub


Public Sub LoadOptions(ByVal Id As Long)
'I know, where's the Open function you ask?
'At Form_Load
'A file doesn't need to be opened more than once, when all it's
'data is stored in a UDT Variable, using this techique is useful.

'The Id is useful for so every time a certain menu requests them
'only the one they have are loaded.
'1 = Report Settings
'2 = General Settings
'3 = Advanced Settings
'4 = Auto-Searching Settings

If Id = 1 Then

With Settings_Report
    .IconBox.Picture = Form1.ImageList2.ListImages.Item(13).Picture
'This adds the data if stored from previous changes
    .Report_AutoSave.Value = (Not MSH_Settings.Report_AutoCreateLog) + 1
 
    If .Report_AutoSave.Value = 1 Then
        .Frame1(1).Enabled = True
        .Report_Filename.BackColor = vbWhite
        .ReportFileOps(0).ForeColor = vbBlack
        .ReportFileOps(1).ForeColor = vbBlack
    Else
        .Frame1(1).Enabled = False
        .Report_Filename.BackColor = &HC0C0C0
        .ReportFileOps(0).ForeColor = &HC0C0C0
        .ReportFileOps(1).ForeColor = &HC0C0C0
    End If

    If MSH_Settings.Report_LogCreateType = 2 Then
        .ReportType(0).Value = True
        ElseIf MSH_Settings.Report_LogCreateType = 1 Then
            .ReportType(1).Value = True
        ElseIf MSH_Settings.Report_LogCreateType = 0 Then
            .ReportType(2).Value = True
        End If

'To store the default filename in memory for easy access
        .Report_Filename = MSH_Settings.Report_DefaultLogFilename

'If it's going to overwrite an existing file if found,
'otherwise increments are added.

        If MSH_Settings.Report_OverWriteLog = True Then
            .ReportFileOps(0).Value = True
        ElseIf MSH_Settings.Report_OverWriteLog = False Then
            .ReportFileOps(1).Value = True
        End If
    End With

ElseIf Id = 2 Then
With Options
    .Image1.Picture = Form1.ImageList2.ListImages.Item(5).Picture
    .Image2.Picture = Form1.ImageList2.ListImages.Item(13).Picture
    .Image3.Picture = Form1.ImageList2.ListImages.Item(9).Picture
    .Image4.Picture = Form1.ImageList2.ListImages.Item(10).Picture
'Exclude classes, this is if you don't want to waste your time
'scrolling through a list of these classes if found.
    .Option_ExcludeClass(0).Value = (Not MSH_Settings.Scan_ExcludePOTENTIALClass) + 1
    .Option_ExcludeClass(1).Value = (Not MSH_Settings.Scan_ExcludeSUSPICIOUSClass) + 1
End With

ElseIf Id = 3 Then
With Settings_Advanced
    .IconBox.Picture = Form1.ImageList2.ListImages.Item(10).Picture
    
    .Check1 = (Not MSH_Settings.Advanced_DisableChecking) + 1
    .Check2 = (Not MSH_Settings.Advanced_CheckComments) + 1
    .Check3 = (Not MSH_Settings.Advanced_EraseDestructive) + 1
    
End With

ElseIf Id = 4 Then
    With Settings_AutoSearch
    .IconBox.Picture = Form1.ImageList2.ListImages.Item(9).Picture
        .Check1 = MSH_Settings.AutoSearch_IncludeSubs
        .Check2 = MSH_Settings.AutoSearch_AutoScanAfter
        If MSH_Settings.AutoSearch_TypeOfSearch = 0 Then
            .Option1(0).Value = True
        ElseIf MSH_Settings.AutoSearch_TypeOfSearch = 1 Then
            .Option1(1).Value = True
        ElseIf MSH_Settings.AutoSearch_TypeOfSearch = 2 Then
            .Option1(2).Value = True
        End If
        
    If .Option1(1).Value = True Then
        .ASDirectory.BackColor = vbWhite
        .ASDirectory.Enabled = True
    Else
        .ASDirectory.BackColor = &HC0C0C0
        .ASDirectory.Enabled = False
    End If
        
        
        .ASDirectory = MSH_Settings.AutoSearch_DefaultDir
    End With
End If

End Sub

Public Sub SaveOptions(ByVal Id As Long)
'Sets the log creation type
'0 = Basic only
'1 = Advanced only
'2 = Both

'Like in load options, The Id is useful for so every time a certain
'menu wants to save them to disk 'only the one they have are saved.
'1 = Report Settings
'2 = General Settings
'3 = Advanced Settings
'4 = Auto-Searching Settings

If Id = 1 Then
    With Settings_Report
        If .Report_Filename = Empty And .Report_AutoSave.Value = 1 Then
        MsgBox "Unacceptable Filename, Please change or disable Automatic Report Logging", vbExclamation
        Exit Sub
        End If

        MSH_Settings.Report_AutoCreateLog = Not (.Report_AutoSave.Value - 1)
        If .ReportType(0).Value = True Then
                MSH_Settings.Report_LogCreateType = 2
        ElseIf .ReportType(1).Value = True Then
                MSH_Settings.Report_LogCreateType = 1
        ElseIf .ReportType(2).Value = True Then
                MSH_Settings.Report_LogCreateType = 0
        End If

'To store the default filename in memory for easy access
        MSH_Settings.Report_DefaultLogFilename = .Report_Filename

'If it's going to overwrite an existing file if found,
'otherwise increments are added.
        If .ReportFileOps(0).Value = True Then
                MSH_Settings.Report_OverWriteLog = True
        ElseIf .ReportFileOps(1).Value = True Then
                MSH_Settings.Report_OverWriteLog = False
        End If
End With
ElseIf Id = 2 Then
    With Options
'Exclude classes, this is if you don't want to waste your time
'scrolling through a list of these classes if found.
        MSH_Settings.Scan_ExcludePOTENTIALClass = Not (.Option_ExcludeClass(0).Value) - 1
        MSH_Settings.Scan_ExcludeSUSPICIOUSClass = Not (.Option_ExcludeClass(1).Value) - 1
    End With
ElseIf Id = 3 Then
    With Settings_Advanced
        MSH_Settings.Advanced_DisableChecking = Not (.Check1.Value) - 1
        MSH_Settings.Advanced_CheckComments = Not (.Check2.Value) - 1
        MSH_Settings.Advanced_EraseDestructive = Not (.Check3.Value) - 1
    End With
ElseIf Id = 4 Then
    With Settings_AutoSearch
        MSH_Settings.AutoSearch_IncludeSubs = .Check1.Value
        MSH_Settings.AutoSearch_AutoScanAfter = .Check2.Value
        If .Option1(0).Value = True Then
                MSH_Settings.AutoSearch_TypeOfSearch = 0
        ElseIf .Option1(1).Value = True Then
                MSH_Settings.AutoSearch_TypeOfSearch = 1
        ElseIf .Option1(2).Value = True Then
                MSH_Settings.AutoSearch_TypeOfSearch = 2
        End If
        MSH_Settings.AutoSearch_DefaultDir = .ASDirectory
    End With
End If


Kill App.Path & "\Settings.MPC" 'To make sure it doesn't just
'reallocate data into the file instead create a new one.
'Another way would be to buffer the data but that's pointless
'for this kind of file
    Open App.Path & "\Settings.MPC" For Binary Access Write Lock Write As #3
        Put #3, , MSH_Settings
    Close #3
End Sub
