VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmPrincipalMenu 
   BackColor       =   &H00C0C0C0&
   Caption         =   "AdminFileProgram"
   ClientHeight    =   9825
   ClientLeft      =   1620
   ClientTop       =   2820
   ClientWidth     =   15675
   Icon            =   "MenuPrincipaldelSistema.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   9825
   ScaleWidth      =   15675
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtPivot 
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   9360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   600
      TabIndex        =   10
      Top             =   1200
      Width           =   3975
   End
   Begin VB.Timer Timer1 
      Interval        =   4000
      Left            =   120
      Top             =   120
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Directorio"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9960
      Picture         =   "MenuPrincipaldelSistema.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Visualizar Contenido de Directorio"
      Top             =   7560
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8160
      Picture         =   "MenuPrincipaldelSistema.frx":1B3C
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Salir del Administrador"
      Top             =   7560
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Imprimir"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3360
      Picture         =   "MenuPrincipaldelSistema.frx":2406
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Enviar a Impresión"
      Top             =   7560
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Borrar"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2040
      Picture         =   "MenuPrincipaldelSistema.frx":2CD0
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Borrar Archivo"
      Top             =   7560
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Actualizar"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      Picture         =   "MenuPrincipaldelSistema.frx":325A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Actualiza Directorio"
      Top             =   7560
      Width           =   1095
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   6165
      Left            =   4680
      TabIndex        =   0
      Top             =   1200
      Width           =   10095
      ExtentX         =   17806
      ExtentY         =   10874
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   6165
      Left            =   4680
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   10095
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Sin Archivos en Directorio."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   1680
         TabIndex        =   2
         Top             =   1800
         Width           =   4935
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Administrador de Archivos de Impresión"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   9
      Top             =   8640
      Width           =   13455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Administrador de Archivos de Impresión"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   8
      Top             =   600
      Width           =   13455
   End
   Begin VB.Shape Shape3 
      Height          =   975
      Left            =   4680
      Top             =   7440
      Width           =   10095
   End
   Begin VB.Shape Shape2 
      Height          =   975
      Left            =   600
      Top             =   7440
      Width           =   3975
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   8880
      Left            =   225
      Shape           =   4  'Rounded Rectangle
      Top             =   495
      Width           =   15255
   End
End
Attribute VB_Name = "frmPrincipalMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo Err_Command1_Click

AutoRefresh = 0

DirListFile List1, DirectoryLocate, "*.txt|*.pdf|*.html"

If List1.ListCount = 0 Then
  Frame1.Visible = True
  WebBrowser1.Visible = False
  List1.Enabled = False
  Command2.Enabled = False
  Command3.Enabled = False
  Command5.Enabled = False
Else
  Command2.Enabled = True
  Command3.Enabled = True
  Command5.Enabled = True
  
  Frame1.Visible = False
  WebBrowser1.Visible = True
  List1.Enabled = True
  List1.Selected(0) = True
  List1.SetFocus
  SetFile = DirectoryLocate & "\" & List1.Text
  WebBrowser1.Navigate DirectoryLocate & "\" & List1.Text

End If


Exit_Err_Command1_Click:
Exit Sub

Err_Command1_Click:
MsgBox Err.Description, vbInformation, "Error de Ejecución"
Resume Exit_Err_Command1_Click

End Sub

Private Sub Command2_Click()
On Error GoTo Err_Command2_Click

AutoRefresh = 0
If MsgBox("Esta Seguro de querer Borrar el Archivo: " & Chr(13) & _
          SetFile, vbOKCancel, "Confirmación") = vbOK Then

If DeleteObjectSystem(SetFile, craFile) = True Then
   Command1_Click
   MsgBox "Archivo fue borrado", vbInformation, "Ejecución Completa"
Else
   MsgBox "Archivo no pudo ser borrado.", vbInformation, "Error Borrando Archivo"
   Command1_Click
End If
Else
  Command1_Click
End If

Exit_Err_Command2_Click:
Exit Sub

Err_Command2_Click:
MsgBox Err.Description, vbInformation, "Error de Ejecución"
Resume Exit_Err_Command2_Click
End Sub

Private Sub Command3_Click()
On Error GoTo Err_Command3_Click

Dim FilePrint As Integer
Dim LineRead As String
Dim vIn As Variant
Dim vOut As Variant

AutoRefresh = 0


If UCase(Mid(SetFile, Len(SetFile) - 3, 4)) = ".TXT" Then


        FilePrint = FreeFile
        
        Open SetFile For Input As #FilePrint
        Printer.Copies = 1
        Printer.ScaleMode = vbPixels
        
        If FileExists(App.Path + "\" & Replace(AppInvoqued, ".exe", ".ini")) = True Then
          Printer.FontName = ReadStringINI(App.Path + "\" & Replace(AppInvoqued, ".exe", ".ini"), "LPD Engage", "Font", "Curier New")
          Printer.FontSize = ReadStringINI(App.Path + "\" & Replace(AppInvoqued, ".exe", ".ini"), "LPD Engage", "Pitch", "7.5")
        Else
          Printer.FontName = "Courier New"
          Printer.FontSize = "7.5"
        End If
        
        
        Do While Not EOF(FilePrint)
        Line Input #FilePrint, LineRead
        
        If Mid(LineRead, 1, 1) = Chr(12) Then
         Printer.NewPage
         Printer.Print " " & Mid(LineRead, 2, Len(LineRead) - 1)
        Else
         Printer.Print LineRead
        End If
        
        
        Loop
        
        Printer.EndDoc
        Close #FilePrint

Else
  If UCase(Mid(SetFile, Len(SetFile) - 4, 5)) = ".HTML" Then
    WebBrowser1.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DONTPROMPTUSER, vIn, vOut
  End If
End If




Exit_Err_Command3_Click:
Exit Sub

Err_Command3_Click:
MsgBox Err.Description, vbInformation, "Error de Ejecución"
Resume Exit_Err_Command3_Click

End Sub

Private Sub Command4_Click()
On Error Resume Next

AutoRefresh = 0

End

End Sub

Private Sub Command5_Click()
On Error GoTo Err_Command5_Click

AutoRefresh = 0

ConsultingDirectoryFull = True
WebBrowser1.Navigate DirectoryLocate
WebBrowser1.SetFocus

Exit_Err_Command5_Click:
Exit Sub

Err_Command5_Click:
MsgBox Err.Description, vbInformation, "Error de ejecución"
Resume Exit_Err_Command5_Click

End Sub

Private Sub Form_Activate()
On Error GoTo Err_Form_Activate

AutoRefresh = 0

If OpenFirst = True Then
   OpenFirst = False
   
  If List1.ListCount <> 0 Then
       Command2.Enabled = True
       Command3.Enabled = True
       Command5.Enabled = True
       
       List1.Enabled = True
       List1.Selected(0) = True
       List1.SetFocus
       
       SetFile = DirectoryLocate & "\" & List1.Text
       WebBrowser1.Navigate DirectoryLocate & "\" & List1.Text
       
  Else
  
      Command2.Enabled = False
      Command3.Enabled = False
      Command5.Enabled = False
      
      List1.Enabled = False
  
  End If
End If

Exit_Err_Form_Activate:
Exit Sub

Err_Form_Activate:
  MsgBox Err.Description, vbInformation, "Error de Ejecución"
Resume Exit_Err_Form_Activate
End Sub

Private Sub Form_Load()
On Error GoTo Err_Form_Load
AutoRefresh = 0


DirListFile List1, DirectoryLocate, "*.txt|*.pdf|*.html"

Label2.Caption = Label2.Caption & " - " & AppInvoqued
Label3.Caption = "Versión: " & App.Major & "." & App.Minor & Chr(13) & _
                 "Company: " & App.CompanyName

If List1.ListCount = 0 Then
  Frame1.Visible = True
  WebBrowser1.Visible = False
Else
  Frame1.Visible = False
  WebBrowser1.Visible = True
End If


Exit_Err_Form_Load:
Exit Sub

Err_Form_Load:
  MsgBox "Apertura de Componente Fallo " & Chr(13) & Err.Description, vbCritical, "Error de Ejecución"
  End
Resume Exit_Err_Form_Load

End Sub

Private Sub Form_Resize()
On Error Resume Next
'//Procedimiento para poder maximizar el formato de visualización
AutoRefresh = 0

If frmPrincipalMenu.Width < 15795 Or _
   frmPrincipalMenu.Height < 10245 Then
   
   If frmPrincipalMenu.Width < 15795 Then
      frmPrincipalMenu.Width = 15795
   ElseIf frmPrincipalMenu.Height < 10245 Then
       frmPrincipalMenu.Height = 10245
   End If
   
Else
    Shape1.Width = frmPrincipalMenu.Width - 540
    Shape1.Height = frmPrincipalMenu.Height - 1365
    Label2.Width = Shape1.Width - 2340
    
    Label3.Width = Shape1.Width - 2340
    Label3.Top = frmPrincipalMenu.Height - 1605
    
    Shape2.Top = frmPrincipalMenu.Height - 2805
    Shape3.Top = frmPrincipalMenu.Height - 2805
    
    Command1.Top = frmPrincipalMenu.Height - 2685
    Command2.Top = frmPrincipalMenu.Height - 2685
    Command3.Top = frmPrincipalMenu.Height - 2685
    Command4.Top = frmPrincipalMenu.Height - 2685
    Command5.Top = frmPrincipalMenu.Height - 2685
    
    List1.Height = frmPrincipalMenu.Height - 4080
    WebBrowser1.Height = frmPrincipalMenu.Height - 4080
    Frame1.Height = frmPrincipalMenu.Height - 4080
    Frame1.Width = frmPrincipalMenu.Width - 5700
    WebBrowser1.Width = frmPrincipalMenu.Width - 5700
    Shape3.Width = frmPrincipalMenu.Width - 5700
    
    
    Command4.Left = Shape3.Left + (Shape3.Width \ 2) - 1567
    Command5.Left = Shape3.Left + (Shape3.Width \ 2) + (1567 - Command5.Width)
End If

End Sub

Private Sub List1_Click()
On Error GoTo Err_List1_Click

AutoRefresh = 0

List1.Selected(List1.ListIndex) = True
List1.ToolTipText = "Archivo Creado: " & DateToFile(DirectoryLocate & "\" & List1.Text, TxtPivot, FechaCreacion) & _
                    " Tamaño: " & Round(Val(DateToFile(DirectoryLocate & "\" & List1.Text, TxtPivot, SizeFile)), 4) & " KB"
SetFile = DirectoryLocate & "\" & List1.Text
If UCase(Mid(SetFile, Len(SetFile) - 3, 4)) = ".TXT" Or _
   UCase(Mid(SetFile, Len(SetFile) - 4, 5)) = ".HTML" Then
   Command3.Enabled = True
Else
   Command3.Enabled = False
End If
WebBrowser1.Navigate DirectoryLocate & "\" & List1.Text

Exit_Err_List1_Click:
Exit Sub

Err_List1_Click:
MsgBox Err.Description, vbInformation, "Error de Ejecución"
Resume Exit_Err_List1_Click
End Sub

Private Sub List1_DblClick()
On Error GoTo Err_List1_DblClick

AutoRefresh = 0

DisplayConsultaHTML SetFile, SetFile, "", 12000, 18000
List1.Selected(List1.ListIndex) = True
SetFile = DirectoryLocate & "\" & List1.Text
WebBrowser1.Navigate DirectoryLocate & "\" & List1.Text

Exit_Err_List1_DblClick:
Exit Sub

Err_List1_DblClick:
MsgBox Err.Description, vbInformation, "Error de Ejecución"
Resume Exit_Err_List1_DblClick
End Sub

Private Sub Timer1_Timer()
On Error GoTo Err_Timer1_Timer
Dim PostFile As Integer

     DoEvents
     If RecuperaProcesoSistema(AppInvoqued) = False Then
       End
     End If
     DoEvents
     If Frame1.Visible = True Then
       Command1_Click
       DoEvents
       Command1_Click
       DoEvents
       Command1_Click
     End If

If ConsultingDirectoryFull = False Then
     AutoRefresh = AutoRefresh + 1
End If

     If AutoRefresh = 30 Then
        If List1.ListIndex >= 0 Then
             PostFile = List1.ListIndex
        End If
        
        Command1_Click
        If (List1.ListCount - 1 >= PostFile) And (List1.ListCount > 0) Then
          List1.Selected(PostFile) = True
          List1.SetFocus
        Else
          If List1.ListCount > 0 Then
            List1.Selected(0) = True
            List1.SetFocus
          End If
          
        End If
        AutoRefresh = 0
     End If
     
     DoEvents
Exit_Err_Timer1_Timer:
Exit Sub

Err_Timer1_Timer:
AutoRefresh = 0
PostFile = 0
Resume Exit_Err_Timer1_Timer

End Sub

Private Sub WebBrowser1_LostFocus()
On Error Resume Next
   If ConsultingDirectoryFull = True Then
      
      Command1_Click
      ConsultingDirectoryFull = False
      
   End If
   
End Sub

