VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scrounger"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5490
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   5490
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList imlHot 
      Left            =   4920
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":030A
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":080E
            Key             =   "Clear"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0D12
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1216
            Key             =   "Save"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlPlain 
      Left            =   4920
      Top             =   5040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":171A
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1C1E
            Key             =   "Clear"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2122
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2626
            Key             =   "Save"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog commDlg 
      Left            =   4920
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   " Dependencies "
      Height          =   4215
      Left            =   60
      TabIndex        =   2
      Top             =   1440
      Width           =   5355
      Begin MSComctlLib.ListView lstView 
         Height          =   3855
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   5115
         _ExtentX        =   9022
         _ExtentY        =   6800
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Dependancy Information"
            Object.Width           =   5080
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Occurrence"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Executable "
      Height          =   675
      Left            =   60
      TabIndex        =   1
      Top             =   720
      Width           =   5355
      Begin VB.TextBox txtFilename 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   5115
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5490
      _ExtentX        =   9684
      _ExtentY        =   1111
      ButtonWidth     =   1138
      ButtonHeight    =   1058
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlPlain"
      HotImageList    =   "imlHot"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "111111"
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Key             =   "Exit"
            ImageKey        =   "Exit"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Clear"
            Key             =   "Clear"
            ImageKey        =   "Clear"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Open"
            Key             =   "Open"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Key             =   "Save"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==========================================================================================================='
' Author                    : Andreas Laubscher                                                             '
' Contact                   : andreaslaubscher@hotmail.com                                                  '
' Date                      : 28 January 2001                                                               '
' Description               : Extracts basic dependancy information from a selected executable              '
'==========================================================================================================='
Option Explicit
'==========================================================================================================='
' Global variable declarations                                                                              '
'==========================================================================================================='
Dim Filen                   As Integer              ' File place holder
Dim varC                    As Integer              ' Generic counter

Dim CompCol                 As New Collection       ' Collection returned from info function

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'==========================================================================================================='
' Controls event triggering                                                                                 '
'==========================================================================================================='
    Select Case Button.Key
    Case "Exit"
        End
    Case "Clear"
        txtFilename.Text = ""
        lstView.ListItems.Clear
    Case "Open"
        DoOpen
    Case "Save"
        DoSave
    End Select
    
End Sub

Private Sub DoOpen()
'==========================================================================================================='
' Opens file, and extracts dependency information                                                           '
'==========================================================================================================='
    On Error GoTo OpenErrorHandler:
    
Dim tmpDependant            As String               ' Holds Collection entry, for readability
Dim tmpOccurrence           As Integer              ' Holds Collection entry, for readability

'-----------------------------------------------------------------------------------------------------------'
' Show the open Dialog                                                                                      '
'-----------------------------------------------------------------------------------------------------------'
    With commDlg
        .DialogTitle = "Open file"
        .Flags = cdlOFNHideReadOnly
        .CancelError = True
        .Filter = "Applications (*.exe)|*.exe"
        .ShowOpen
        txtFilename.Text = .FileName
    End With
'-----------------------------------------------------------------------------------------------------------'
' Reset the display                                                                                         '
'-----------------------------------------------------------------------------------------------------------'
    Screen.MousePointer = vbHourglass
    lstView.ListItems.Clear
'-----------------------------------------------------------------------------------------------------------'
' Add the Dll's                                                                                             '
'-----------------------------------------------------------------------------------------------------------'
    Set CompCol = InfoCollection(txtFilename.Text, ".dll")
    lstView.ListItems.Add , , "DLL's"
    If CompCol.Count = 0 Then
        lstView.ListItems.Add , , "None"
    Else
        For varC = 1 To CompCol.Count
            tmpDependant = Mid(CompCol(varC), 1, InStr(1, CompCol(varC), ";") - 1)
            tmpOccurrence = Mid(CompCol(varC), InStr(1, CompCol(varC), ";") + 1)
            With lstView.ListItems.Add(, , tmpDependant)
                                  .SubItems(1) = tmpOccurrence
            End With
        Next
    End If
'-----------------------------------------------------------------------------------------------------------'
' Add a separator to the listview                                                                           '
'-----------------------------------------------------------------------------------------------------------'
    lstView.ListItems.Add , , ""
'-----------------------------------------------------------------------------------------------------------'
' Now, add the Ocx's                                                                                        '
'-----------------------------------------------------------------------------------------------------------'
    Set CompCol = InfoCollection(txtFilename.Text, ".ocx")
    lstView.ListItems.Add , , "OCX's"
    If CompCol.Count = 0 Then
        lstView.ListItems.Add , , "None"
    Else
        For varC = 1 To CompCol.Count
            tmpDependant = Mid(CompCol(varC), 1, InStr(1, CompCol(varC), ";") - 1)
            tmpOccurrence = Mid(CompCol(varC), InStr(1, CompCol(varC), ";") + 1)
            With lstView.ListItems.Add(, , tmpDependant)
                                  .SubItems(1) = tmpOccurrence
            End With
        Next
    End If
'-----------------------------------------------------------------------------------------------------------'
' Reset the display                                                                                         '
'-----------------------------------------------------------------------------------------------------------'
    Screen.MousePointer = vbNormal
    
    Exit Sub
    
OpenErrorHandler:
    If Err.Number <> cdlCancel Then
        MsgBox "A critical error has occurred and this application will now exit" & Chr(13) & _
               "Error " & Err.Number & ": " & Err.Description, _
               vbCritical, _
               "Error!"
        End
    End If

End Sub

Private Function InfoCollection(FileName As String, ToExtract As String) As Collection
'==========================================================================================================='
' Returns a collection of the names and occurrence of a specified search string in a file                   '
'==========================================================================================================='
    On Error GoTo InfoColErrorHandler

Dim CounterA                As Integer              ' Generic Counter
Dim CounterB                As Integer              ' Generic Counter
Dim tmpInt                  As Integer              ' Holds Collection entry, for readability
    
Dim Counter                 As Long                 ' Position holder in data string

Dim Contents                As String               ' Selected file data
Dim tmpWord                 As String               ' Current found item
Dim tmpChar                 As String * 1           ' Character in tmpWord

Dim tmpCollection           As New Collection       ' Collection holding found items

    Counter = 1
    '-------------------------------------------------------------------------------------------------------'
    ' Retrieve file content to variable                                                                     '
    '-------------------------------------------------------------------------------------------------------'
    Filen = FreeFile
    Open txtFilename.Text For Binary As Filen
    Contents = Space(LOF(Filen))
    Get Filen, , Contents
    Close Filen
    
    Contents = LCase(Contents)
    '-------------------------------------------------------------------------------------------------------'
    ' Find and add all instances of the sought item                                                         '
    '-------------------------------------------------------------------------------------------------------'
    Do Until InStr(Counter, Contents, ToExtract) = 0
        '---------------------------------------------------------------------------------------------------'
        ' Find and retrieve the next instance                                                               '
        '---------------------------------------------------------------------------------------------------'
        Counter = InStr(Counter, Contents, ToExtract) + 1
        tmpWord = Mid(Contents, Counter - 20, 23)
        '---------------------------------------------------------------------------------------------------'
        ' Now we have the item, we edit it into a usable string                                             '
        '---------------------------------------------------------------------------------------------------'
        CounterB = 0
        For CounterA = 1 To Len(tmpWord)
            '-----------------------------------------------------------------------------------------------'
            ' We can only use String, Numeric, and full stops in our entry                                  '
            '-----------------------------------------------------------------------------------------------'
            tmpChar = Mid(tmpWord, CounterA, 1)
            If UCase(tmpChar) = LCase(tmpChar) Then
                If Not IsNumeric(tmpChar) Then
                    If tmpChar <> "." Then
                        CounterB = CounterA
                    End If
                End If
            End If
        Next
        tmpWord = Mid(tmpWord, CounterB + 1)
        '---------------------------------------------------------------------------------------------------'
        ' OK, now we have a usable string, let's add it, or add to its tally                                '
        '---------------------------------------------------------------------------------------------------'
        tmpCollection.Add tmpWord & ";" & "1", tmpWord
    Loop
'-----------------------------------------------------------------------------------------------------------'
' Show the open Dialog                                                                                      '
'-----------------------------------------------------------------------------------------------------------'
    Set InfoCollection = tmpCollection
    
    Exit Function
    
InfoColErrorHandler:

    If Err.Number = 457 Then
        '---------------------------------------------------------------------------------------------------'
        ' We've tried to add a duplicate item to our collection                                             '
        '---------------------------------------------------------------------------------------------------'
        tmpInt = Mid(tmpCollection(tmpWord), InStr(1, tmpCollection(tmpWord), ";") + 1) + 1
        '---------------------------------------------------------------------------------------------------'
        ' Remove the item, then add it with the new value                                                   '
        '---------------------------------------------------------------------------------------------------'
        tmpCollection.Remove tmpWord
        tmpCollection.Add tmpWord & ";" & tmpInt, tmpWord
        Resume Next
    Else
        MsgBox "A critical error has occurred and this application will now exit" & Chr(13) & _
               "Error " & Err.Number & ": " & Err.Description, _
               vbCritical, _
               "Error!"
        End
    End If
    
End Function

Private Sub DoSave()
'==========================================================================================================='
' Saves extracted information to file                                                                       '
'==========================================================================================================='
    On Error GoTo SaveErrorHandler
'-----------------------------------------------------------------------------------------------------------'
' Show the Save Dialog                                                                                      '
'-----------------------------------------------------------------------------------------------------------'
    With commDlg
        .FileName = ""
        .DialogTitle = "Save file"
        .CancelError = True
        .Flags = cdlOFNOverwritePrompt + cdlOFNHideReadOnly
        .Filter = "Text File (*.txt)|*.txt"
        .ShowSave
        txtFilename.Text = .FileName
    End With
'-----------------------------------------------------------------------------------------------------------'
' Open the selected file for output                                                                         '
'-----------------------------------------------------------------------------------------------------------'
    Filen = FreeFile
    Open txtFilename.Text For Output As Filen
'-----------------------------------------------------------------------------------------------------------'
' Write all the non-cosmetic entries to file                                                                '
'-----------------------------------------------------------------------------------------------------------'
    For varC = 1 To lstView.ListItems.Count
        If (lstView.ListItems(varC) <> "DLL's") And _
           (lstView.ListItems(varC) <> "OCX's") And _
           (lstView.ListItems(varC) <> "None") And _
           (Len(lstView.ListItems(varC)) > 0) Then
           Print #Filen, lstView.ListItems(varC).Text & _
                         ", " & _
                         lstView.ListItems(varC).SubItems(1)
        End If
    Next
    
    Close #Filen
    
    Exit Sub
    
SaveErrorHandler:
    If Err.Number <> cdlCancel Then
        MsgBox "A critical error has occurred and this application will now exit" & Chr(13) & _
               "Error " & Err.Number & ": " & Err.Description, _
               vbCritical, _
               "Error!"
        End
    End If
    
End Sub
