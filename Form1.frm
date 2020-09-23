VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00EEE6DE&
   Caption         =   "Capture all IE windows' html controls"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5640
   ScaleWidth      =   8655
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   7080
      TabIndex        =   3
      Top             =   60
      Width           =   1335
   End
   Begin VB.TextBox txtData 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   540
      Width           =   8355
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Open new IE window"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   60
      Width           =   1875
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Find all open IE windows"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub Command2_Click()
    Dim oApp As Object
    Dim oWin As Object, i%
    Dim iE As New InternetExplorer
    Dim oDoc As Object
    Dim oForm As Object
    Dim oElem As Object
    Dim oTxt As Object
    
    txtData.Text = ""
    i = -1
    
    Set oApp = CreateObject("Shell.Application") 'All Explorer instances (Windows Explorer and Internet Explorer)
    For Each oWin In oApp.windows 'Loop through the windows
        If TypeName(oWin.Document) = "HTMLDocument" Then  'only take IE instances
            Set iE = oWin
            'Shows basic information for the window, so that you can choose to include then
            If MsgBox("APPLICATION: " & oWin & vbCrLf & _
                   "ie.FULL NAME: " & iE.FullName & vbCrLf & vbCrLf & _
                   "ie.LOCATION URL: " & iE.LocationURL & vbCrLf & _
                   "ie.LOCATION NAME : " & iE.LocationName & vbCrLf & _
                   "ie.STATUS TEXT : " & iE.StatusText, 33) = vbOK Then
                
                Set oDoc = iE.Document 'Document object
                
                On Error Resume Next
                For Each oForm In oDoc.Forms
                'If Not oDoc.Forms(0).Elements Is Nothing Then
                If Not oForm.Elements Is Nothing Then
                    'Set oElem = oDoc.Forms(0).Elements
                    Set oElem = oForm.Elements
                    i = 0
                    'print the window's Url and the Form's name
                    txtData.Text = txtData.Text & String(150, "=") & vbCrLf
                    txtData.Text = txtData.Text & "URL : " & iE.LocationURL & vbCrLf
                    txtData.Text = txtData.Text & "FORM: " & oForm.Name & vbCrLf & String(150, "=") & vbCrLf & vbCrLf
                    
                    'For i = 0 To oDoc.Forms(0).Elements.Count - 1
                    For Each oTxt In oElem 'Loop through every control and show their name and value
                    
                        'MsgBox "NAME: " & oTxt.Name & vbCrLf & "VALUE: " & oTxt.Value
                        txtData.Text = txtData.Text & Format$(i, "000") & _
                            " - NAME : " & oTxt.Name & vbCrLf & "      VALUE: " & _
                            oTxt.Value & vbCrLf & "---" & vbCrLf
                        
                        'You can edit it's value too!!
                        If oTxt.Name = "txtEmail" Then
                            oTxt.Value = "*******" & oTxt.Value
                            iE.Document.Forms(0).Elements(i).Value = "EXITOSO!!!!!!"
                        End If
                        i = i + 1
                        
                    Next
                    Set oElem = Nothing
                End If
                Next
                Set oDoc = Nothing
                On Error GoTo 0
            End If
            Set iE = Nothing
        End If
    Next
    Set oApp = Nothing
    If i = -1 Then MsgBox "No Open Internet Explorer Windows!", 16
End Sub
Private Sub Command3_Click()
    Dim iE As New InternetExplorer
    
    iE.Visible = True
    iE.MenuBar = False
    iE.ToolBar = 0
    iE.Navigate "www.planetsourcecode.com"
    'ie.FullScreen = True
    Set iE = Nothing
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    txtData.Width = Me.ScaleWidth - (txtData.Left * 2)
    txtData.Height = Me.ScaleHeight - txtData.Top - txtData.Left
End Sub
