<div align="center">

## \*\*\*UPDATED\*\*\* Window SPY


</div>

### Description

***UPDATED*** Gets tons of information on the window your mouse is over.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Matt Evans](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/matt-evans.md)
**Level**          |Unknown
**User Rating**    |3.9 (59 globes from 15 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Complete Applications](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/complete-applications__1-27.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/matt-evans-updated-window-spy__1-1424/archive/master.zip)

### API Declarations

```
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Public Declare Function WindowFromPointXY Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Public Declare Function GetWindowWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long) As Integer
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
```


### Source Code

```
**** Put this in a module ****
Function WindowSPY(WinHdl As TextBox, WinClass As TextBox, WinTxt As TextBox, WinStyle As TextBox, WinIDNum As TextBox, WinPHandle As TextBox, WinPText As TextBox, WinPClass As TextBox, WinModule As TextBox)
'Call This In A Timer
Dim pt32 As POINTAPI, ptx As Long, pty As Long, sWindowText As String * 100
Dim sClassName As String * 100, hWndOver As Long, hWndParent As Long
Dim sParentClassName As String * 100, wID As Long, lWindowStyle As Long
Dim hInstance As Long, sParentWindowText As String * 100
Dim sModuleFileName As String * 100, r As Long
Static hWndLast As Long
  Call GetCursorPos(pt32)
  ptx = pt32.X
  pty = pt32.Y
  hWndOver = WindowFromPointXY(ptx, pty)
  If hWndOver <> hWndLast Then
    hWndLast = hWndOver
    WinHdl.Text = "Window Handle: " & hWndOver
    r = GetWindowText(hWndOver, sWindowText, 100)
    WinTxt.Text = "Window Text: " & Left(sWindowText, r)
    r = GetClassName(hWndOver, sClassName, 100)
    WinClass.Text = "Window Class Name: " & Left(sClassName, r)
    lWindowStyle = GetWindowLong(hWndOver, GWL_STYLE)
    WinStyle.Text = "Window Style: " & lWindowStyle
    hWndParent = GetParent(hWndOver)
      If hWndParent <> 0 Then
        wID = GetWindowWord(hWndOver, GWW_ID)
        WinIDNum.Text = "Window ID Number: " & wID
        WinPHandle.Text = "Parent Window Handle: " & hWndParent
        r = GetWindowText(hWndParent, sParentWindowText, 100)
        WinPText.Text = "Parent Window Text: " & Left(sParentWindowText, r)
        r = GetClassName(hWndParent, sParentClassName, 100)
        WinPClass.Text = "Parent Window Class Name: " & Left(sParentClassName, r)
      Else
        WinIDNum.Text = "Window ID Number: N/A"
        WinPHandle.Text = "Parent Window Handle: N/A"
        WinPText.Text = "Parent Window Text : N/A"
        WinPClass.Text = "Parent Window Class Name: N/A"
      End If
        hInstance = GetWindowWord(hWndOver, GWW_HINSTANCE)
        r = GetModuleFileName(hInstance, sModuleFileName, 100)
    WinModule.Text = "Module: " & Left(sModuleFileName, r)
  End If
End Function
****** END OF MODULE ******
'Put this is notepad and rename is winspy.frm
VERSION 5.00
Begin VB.Form Form1
BackColor=&H00000000&
Caption ="Window SPY"
ClientHeight=3480
ClientLeft =2280
ClientTop=1590
ClientWidth =4440
LinkTopic="Form1"
ScaleHeight =3480
ScaleWidth =4440
Begin VB.Timer Timer1
Interval=10
Left=1080
Top =1560
End
Begin VB.TextBox Text9
Appearance =0 'Flat
BackColor=&H00000000&
BeginProperty Font
Name="Arial"
Size=8.25
Charset =0
Weight =700
Underline=0'False
Italic =0'False
Strikethrough=0'False
EndProperty
ForeColor=&H00FFFFFF&
  Height =285
  Left=120
  TabIndex=8
  Text="Text9"
  Top =3000
  Width=4215
  End
  Begin VB.TextBox Text8
  Appearance =0 'Flat
  BackColor=&H00000000&
  BeginProperty Font
  Name="Arial"
  Size=8.25
  Charset =0
  Weight =700
  Underline=0'False
  Italic =0'False
  Strikethrough=0'False
  EndProperty
  ForeColor=&H00FFFFFF&
    Height =285
    Left=120
    TabIndex=7
    Text="Text8"
    Top =2640
    Width=4215
    End
    Begin VB.TextBox Text7
    Appearance =0 'Flat
    BackColor=&H00000000&
    BeginProperty Font
    Name="Arial"
    Size=8.25
    Charset =0
    Weight =700
    Underline=0'False
    Italic =0'False
    Strikethrough=0'False
    EndProperty
    ForeColor=&H00FFFFFF&
      Height =285
      Left=120
      TabIndex=6
      Text="Text7"
      Top =2280
      Width=4215
      End
      Begin VB.TextBox Text6
      Appearance =0 'Flat
      BackColor=&H00000000&
      BeginProperty Font
      Name="Arial"
      Size=8.25
      Charset =0
      Weight =700
      Underline=0'False
      Italic =0'False
      Strikethrough=0'False
      EndProperty
      ForeColor=&H00FFFFFF&
        Height =285
        Left=120
        TabIndex=5
        Text="Text6"
        Top =1920
        Width=4215
        End
        Begin VB.TextBox Text5
        Appearance =0 'Flat
        BackColor=&H00000000&
        BeginProperty Font
        Name="Arial"
        Size=8.25
        Charset =0
        Weight =700
        Underline=0'False
        Italic =0'False
        Strikethrough=0'False
        EndProperty
        ForeColor=&H00FFFFFF&
          Height =285
          Left=120
          TabIndex=4
          Text="Text5"
          Top =1560
          Width=4215
          End
          Begin VB.TextBox Text4
          Appearance =0 'Flat
          BackColor=&H00000000&
          BeginProperty Font
          Name="Arial"
          Size=8.25
          Charset =0
          Weight =700
          Underline=0'False
          Italic =0'False
          Strikethrough=0'False
          EndProperty
          ForeColor=&H00FFFFFF&
            Height =285
            Left=120
            TabIndex=3
            Text="Text4"
            Top =1200
            Width=4215
            End
            Begin VB.TextBox Text3
            Appearance =0 'Flat
            BackColor=&H00000000&
            BeginProperty Font
            Name="Arial"
            Size=8.25
            Charset =0
            Weight =700
            Underline=0'False
            Italic =0'False
            Strikethrough=0'False
            EndProperty
            ForeColor=&H00FFFFFF&
              Height =285
              Left=120
              TabIndex=2
              Text="Text3"
              Top =840
              Width=4215
              End
              Begin VB.TextBox Text2
              Appearance =0 'Flat
              BackColor=&H00000000&
              BeginProperty Font
              Name="Arial"
              Size=8.25
              Charset =0
              Weight =700
              Underline=0'False
              Italic =0'False
              Strikethrough=0'False
              EndProperty
              ForeColor=&H00FFFFFF&
                Height =285
                Left=120
                TabIndex=1
                Text="Text2"
                Top =480
                Width=4215
                End
                Begin VB.TextBox Text1
                Appearance =0 'Flat
                BackColor=&H00000000&
                BeginProperty Font
                Name="Arial"
                Size=8.25
                Charset =0
                Weight =700
                Underline=0'False
                Italic =0'False
                Strikethrough=0'False
                EndProperty
                ForeColor=&H00FFFFFF&
                  Height =285
                  Left=120
                  TabIndex=0
                  Text="Text1"
                  Top =120
                  Width=4215
                  End
                  End
                  Attribute VB_Name = "Form1"
                  Attribute VB_GlobalNameSpace = False
                  Attribute VB_Creatable = False
                  Attribute VB_PredeclaredId = True
                  Attribute VB_Exposed = False
Private Sub Timer1_Timer()
  WindowSPY Text1, Text2, Text3, Text4, Text5, Text6, Text7, Text8, Text9
End Sub
```

