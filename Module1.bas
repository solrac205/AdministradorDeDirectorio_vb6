Attribute VB_Name = "Module1"
Global AppInvoqued As String
Global DirectoryLocate As String
Global CaptionRuning As String
Global SetFile As String
Global OpenFirst As Boolean
Global ConsultingDirectoryFull As Boolean
Global AutoRefresh As Integer

Public Enum OLECMDID
  OLECMDID_OPEN = 1
  OLECMDID_NEW = 2
  OLECMDID_SAVE = 3
  OLECMDID_SAVEAS = 4
  OLECMDID_SAVECOPYAS = 5
  OLECMDID_PRINT = 6
  OLECMDID_PRINTPREVIEW = 7
  OLECMDID_PAGESETUP = 8
  OLECMDID_SPELL = 9
  OLECMDID_PROPERTIES = 10
  OLECMDID_CUT = 11
  OLECMDID_COPY = 12
  OLECMDID_PASTE = 13
  OLECMDID_PASTESPECIAL = 14
  OLECMDID_UNDO = 15
  OLECMDID_REDO = 16
  OLECMDID_SELECTALL = 17
  OLECMDID_CLEARSELECTION = 18
  OLECMDID_ZOOM = 19
  OLECMDID_GETZOOMRANGE = 20
  OLECMDID_UPDATECOMMANDS = 21
  OLECMDID_REFRESH = 22
  OLECMDID_STOP = 23
  OLECMDID_HIDETOOLBARS = 24
  OLECMDID_SETPROGRESSMAX = 25
  OLECMDID_SETPROGRESSPOS = 26
  OLECMDID_SETPROGRESSTEXT = 27
  OLECMDID_SETTITLE = 28
  OLECMDID_SETDOWNLOADSTATE = 29
  OLECMDID_STOPDOWNLOAD = 30
  OLECMDID_ONTOOLBARACTIVATED = 31
  OLECMDID_FIND = 32
  OLECMDID_DELETE = 33
  OLECMDID_HTTPEQUIV = 34
  OLECMDID_HTTPEQUIV_DONE = 35
  OLECMDID_ENABLE_INTERACTION = 36
  OLECMDID_ONUNLOAD = 37
  OLECMDID_PROPERTYBAG2 = 38
  OLECMDID_PREREFRESH = 39
  OLECMDID_SHOWSCRIPTERROR = 40
  OLECMDID_SHOWMESSAGE = 41
  OLECMDID_SHOWFIND = 42
  OLECMDID_SHOWPAGESETUP = 43
  OLECMDID_SHOWPRINT = 44
  OLECMDID_CLOSE = 45
  OLECMDID_ALLOWUILESSSAVEAS = 46
  OLECMDID_DONTDOWNLOADCSS = 47
  OLECMDID_UPDATEPAGESTATUS = 48
  OLECMDID_PRINT2 = 49
  OLECMDID_PRINTPREVIEW2 = 50
  OLECMDID_SETPRINTTEMPLATE = 51
  OLECMDID_GETPRINTTEMPLATE = 52
  OLECMDID_PAGEACTIONBLOCKED = 55
  OLECMDID_PAGEACTIONUIQUERY = 56
  OLECMDID_FOCUSVIEWCONTROLS = 57
  OLECMDID_FOCUSVIEWCONTROLSQUERY = 58
  OLECMDID_SHOWPAGEACTIONMENU = 59
  OLECMDID_ADDTRAVELENTRY = 60
  OLECMDID_UPDATETRAVELENTRY = 61
  OLECMDID_UPDATEBACKFORWARDSTATE = 62
  OLECMDID_OPTICAL_ZOOM = 63
  OLECMDID_OPTICAL_GETZOOMRANGE = 64
  OLECMDID_WINDOWSTATECHANGED = 65
  OLECMDID_ACTIVEXINSTALLSCOPE = 66
  OLECMDID_UPDATETRAVELENTRY_DATARECOVERY = 67
End Enum

''Constants
''OLECMDID_OPEN
''File menu, Open command
''OLECMDID_NEW
''File Menu, New Command
''OLECMDID_SAVE
''File menu, Save command
''OLECMDID_SAVEAS
''File menu, Save As command
''OLECMDID_SAVECOPYAS
''File menu, Save Copy As command
''OLECMDID_PRINT
''File menu, Print command
''OLECMDID_PRINTPREVIEW
''File menu, Print Preview command
''OLECMDID_PAGESETUP
''File menu, Page Setup command
''OLECMDID_SPELL
''Tools menu, Spelling command
''OLECMDID_PROPERTIES
''File menu, Properties command
''OLECMDID_CUT
''Edit menu, Cut command
''OLECMDID_COPY
''Edit menu, Copy command
''OLECMDID_PASTE
''Edit menu, Paste command
''OLECMDID_PASTESPECIAL
''Edit menu, Paste Special command
''OLECMDID_UNDO
''Edit menu, Undo command
''OLECMDID_REDO
''Edit menu, Redo command
''OLECMDID_SELECTALL
''Edit menu, Select All command
''OLECMDID_CLEARSELECTION
''Edit menu, Clear command
''OLECMDID_ZOOM
''View menu, Zoom command (see below for details.)
''OLECMDID_GETZOOMRANGE
''Retrieves zoom range applicable to View Zoom (see below for details.)
''OLECMDID_UPDATECOMMANDS
''Informs the receiver, usually a frame, of state changes. The receiver can then query the status of the commands whenever convenient.
''OLECMDID_REFRESH
''Asks the receiver to refresh its display. Implemented by the document/object.
''OLECMDID_STOP
''Stops all current processing. Implemented by the document/object.
''OLECMDID_HIDETOOLBARS
''View menu, Toolbars command. Implemented by the document/object to hide its toolbars.
''OLECMDID_SETPROGRESSMAX
''Sets the maximum value of a progress indicator if one is owned by the receiving object, usually a frame. The minimum value is always zero.
''OLECMDID_SETPROGRESSPOS
''Sets the current value of a progress indicator if one is owned by the receiving object, usually a frame.
''OLECMDID_SETPROGRESSTEXT
''Sets the text contained in a progress indicator if one is owned by the receiving object, usually a frame. If the receiver currently has no progress indicator, this text should be displayed in the status bar (if one exists) as withIOleInPlaceFrame::SetStatusText.
''OLECMDID_SETTITLE
''Sets the title bar text of the receiving object, usually a frame.
''OLECMDID_SETDOWNLOADSTATE
''Called by the object when downloading state changes. Takes a VT_BOOL parameter, which is TRUE if the object is downloading data and FALSE if it not. Primarily implemented by the frame.
''OLECMDID_STOPDOWNLOAD
''Stops the download when executed. Typically, this command is propagated to all contained objects. When queried, sets MSOCMDF_ENABLED. Implemented by the document/object.
''OLECMDID_ONTOOLBARACTIVATED
''OLECMDID_FIND
''Edit menu, Find command
''OLECMDID_DELETE
''Edit menu, Delete command
''OLECMDID_HTTPEQUIV
''Issued in response to HTTP-EQUIV metatag and results in a call to the deprecated OnHttpEquiv method with thefDone parameter set to false. This command takes a VT_BSTR parameter which is passed to OnHttpEquiv.
''OLECMDID_HTTPEQUIV_DONE
''Issued in response to HTTP-EQUIV metatag and results in a call to the deprecated OnHttpEquiv method with thefDone parameter set to true. This command takes a VT_BSTR parameter which is passed to OnHttpEquiv.
''OLECMDID_ENABLE_INTERACTION
''Pauses or resumes receiver interaction. This command takes a VT_BOOL parameter that pauses interaction when set to FALSE and resumes interaction when set to TRUE.
''OLECMDID_ONUNLOAD
''Notifies the receiver of an intent to close the window imminently. This command takes a VT_BOOL output parameter that returns TRUE if the receiver can close and FALSE if it can't.
''OLECMDID_PROPERTYBAG2
''This command has no effect.
''OLECMDID_PREREFRESH
''Notifies the receiver that a refresh is about to start.
''OLECMDID_SHOWSCRIPTERROR
''Tells the receiver to display the script error message.
''OLECMDID_SHOWMESSAGE
''This command takes an IHTMLEventObj input parameter that contains a message that the receiver shows.
''OLECMDID_SHOWFIND
''Tells the receiver to show the Find dialog box. It takes a VT_DISPATCH input param.
''OLECMDID_SHOWPAGESETUP
''Tells the receiver to show the Page Setup dialog box. It takes an IHTMLEventObj2 input parameter.
''OLECMDID_SHOWPRINT
''Tells the receiver to show the Print dialog box. It takes an IHTMLEventObj2 input parameter.
''OLECMDID_CLOSE
''The exit command for the File menu.
''OLECMDID_ALLOWUILESSSAVEAS
''Supports the QueryStatus method.
''OLECMDID_DONTDOWNLOADCSS
''Notifies the receiver that CSS files should not be downloaded when in DesignMode.
''OLECMDID_UPDATEPAGESTATUS
''This command has no effect.
''OLECMDID_PRINT2
''File menu, updated Print command
''OLECMDID_PRINTPREVIEW2
''File menu, updated Print Preview command
''OLECMDID_SETPRINTTEMPLATE
''Sets an explicit Print Template value of TRUE or FALSE, based on a VT_BOOL input parameter.
''OLECMDID_GETPRINTTEMPLATE
''Gets a VT_BOOL output parameter indicating whether the Print Template value is TRUE or FALSE.
''OLECMDID_PAGEACTIONBLOCKED
''Indicates that a page action has been blocked. PAGEACTIONBLOCKED is designed for use with applications that host the Internet Explorer WebBrowser control to implement their own UI.
''OLECMDID_PAGEACTIONUIQUERY
''Specifies which actions are displayed in the Internet Explorer notification band.
''OLECMDID_FOCUSVIEWCONTROLS
''Causes the Internet Explorer WebBrowser control to focus its default notification band. Hosts can send this command at any time. The return value is S_OK if the band is present and is in focus, or S_FALSE otherwise.
''OLECMDID_FOCUSVIEWCONTROLSQUERY
''This notification event is provided for applications that display Internet Explorers default notification band implementation. By default, when the user presses the ALT-N key combination, Internet Explorer treats it as a request to focus the notification band.
''OLECMDID_SHOWPAGEACTIONMENU
''Causes the Internet Explorer WebBrowser control to show the Information Bar menu.
''OLECMDID_ADDTRAVELENTRY
''Causes the Internet Explorer WebBrowser control to create an entry at the current Travel Log offset. The Docobject should implement ITravelLogClient and IPersist interfaces, which are used by the Travel Log as it processes this command with calls to GetWindowData and GetPersistID, respectively.
''OLECMDID_UPDATETRAVELENTRY
''Called when LoadHistory is processed to update the previous Docobject state. For synchronous handling, this command can be called before returning from the LoadHistory call. For asynchronous handling, it can be called later.
''OLECMDID_UPDATEBACKFORWARDSTATE
''Updates the state of the browser's Back and Forward buttons.
''OLECMDID_OPTICAL_ZOOM
''Windows Internet Explorer 7 and later. Sets the zoom factor of the browser. Takes a VT_I4 parameter in the range of 10 to 1000 (percent).
''OLECMDID_OPTICAL_GETZOOMRANGE
''Windows Internet Explorer 7 and later. Retrieves the minimum and maximum browser zoom factor limits. Returns a VT_I4 parameter; the LOWORD is the minimum zoom factor, the HIWORD is the maximum.
''OLECMDID_WINDOWSTATECHANGED
''Windows Internet Explorer 7 and later. Notifies the Internet Explorer WebBrowser control of changes in window states, such as losing focus, or becoming hidden or minimized. The host indicates what has changed by setting OLECMDID_WINDOWSTATE_FLAG option flags in nCmdExecOpt.
''OLECMDID_ACTIVEXINSTALLSCOPE
''Windows Internet Explorer 8 with Windows Vista. Has no effect with Windows Internet Explorer 8 with Windows XP. Notifies Trident to use the indicated Install Scope to install the ActiveX Control specified by the indicated Class ID. For more information, see the Remarks section.
''OLECMDID_UPDATETRAVELENTRY_DATARECOVERY
''Internet Explorer 8. Unlike OLECMDID_UPDATETRAVELENTRY, this updates a Travel Log entry that is not initialized from a previous Docobject state. While this command is not called from IPersistHistory::LoadHistory, it can be called separately to save browser state that can be used later to recover from a crash.



Public Enum OLECMDEXECOPT
  OLECMDEXECOPT_DODEFAULT = 0
  OLECMDEXECOPT_PROMPTUSER = 1
  OLECMDEXECOPT_DONTPROMPTUSER = 2
  OLECMDEXECOPT_SHOWHELP = 3
End Enum

''Constants
''OLECMDEXECOPT_DODEFAULT
''Prompt the user for input or not, whichever is the default behavior.
''OLECMDEXECOPT_PROMPTUSER
''Execute the command after obtaining user input.
''OLECMDEXECOPT_DONTPROMPTUSER
''Execute the command without prompting the user. For example, clicking the Print toolbar button causes a document to be immediately printed without user input.
''OLECMDEXECOPT_SHOWHELP
''Show help for the corresponding command, but do not execute.



Public Enum DatosDeArchivo
 '// Colección para identificación del dato de Fecha de Archivo que queremos consultar
 '// del archivo consultado en funcion DateToFile
     FechaCreacion = 1
     FechaModificacion = 2
     FechaUltimoAcceso = 3
     SizeFile = 4
End Enum

Public Enum SyntaxError
  NoSintaxError = 0
  ErrorSintax = 1
End Enum

Public Enum TypeObjectSystem
   craDirectory = 1
   craFile = 2
End Enum

Private Declare Function PathFileExistsW Lib "shlwapi.dll" ( _
    ByVal pszPath As Long) As Boolean
'// Declaración de API para verificación de existencia de Directorios.

Private Declare Function SHCreateDirectory Lib "shell32" ( _
    ByVal hwnd As Long, _
    ByVal pszPath As Long) As Long
'// Declaración de API para creación de Estructuras Completas de Directorios


Const TH32CS_SNAPHEAPLIST = &H1
Const TH32CS_SNAPPROCESS = &H2
Const TH32CS_SNAPTHREAD = &H4
Const TH32CS_SNAPMODULE = &H8
Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or _
                            TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Const TH32CS_INHERIT = &H80000000
Const MAX_PATH As Integer = 260
  
'Estructura para los procesos
'-----------------------------------
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long ' Flags
    szExeFile As String * MAX_PATH ' nombre del ejecutable
End Type
  
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" ( _
    ByVal lFlags As Long, _
    ByVal lProcessID As Long) As Long
  
Private Declare Function Process32First Lib "kernel32" ( _
    ByVal hSnapShot As Long, _
    uProcess As PROCESSENTRY32) As Long
  
Private Declare Function Process32Next Lib "kernel32" ( _
    ByVal hSnapShot As Long, _
    uProcess As PROCESSENTRY32) As Long
  
Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)
  
Public Function RecuperaProcesoSistema(ProcesoBuscado As String) As Boolean
On Error GoTo Err_RecuperaProcesoSistema
Dim EncontroProceso As Boolean

Dim hSnapShot As Long, uProcess As PROCESSENTRY32

EncontroProceso = False

If Nulidad(ProcesoBuscado) = True Then

  MsgBox "Error, Ningún aplicativo esta llamando al componente", vbCritical, "Error en Ejecución"
  RecuperaProcesoSistema = EncontroProceso
  Exit Function
Else

hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0&)
uProcess.dwSize = Len(uProcess)

R = Process32First(hSnapShot, uProcess)
    
Do While R
    DoEvents
    R = Process32Next(hSnapShot, uProcess)
  
    If Mid(uProcess.szExeFile, 1, Len(ProcesoBuscado)) = ProcesoBuscado Then
       EncontroProceso = True
    End If

Loop

Call CloseHandle(hSnapShot)
    
RecuperaProcesoSistema = EncontroProceso
End If

Exit_Err_RecuperaProcesoSistema:
Exit Function

Err_RecuperaProcesoSistema:
RecuperaProcesoSistema = False
Resume Exit_Err_RecuperaProcesoSistema
End Function

Function FileExists(ByVal sFIle)
'// Función de Comprobación de Existencia de Archivos.
On Error Resume Next
Dim objFSO

Set objFSO = CreateObject("Scripting.FileSystemObject")

If objFSO.FileExists(sFIle) Then
     FileExists = True
Else
     FileExists = False
End If

Set objFSO = Nothing

End Function


Public Function DeleteObjectSystem(PathAndObjectDelete As String, TypeDelete As TypeObjectSystem) As Boolean
On Error GoTo Err_DeleteObjectSystem

Dim fs As Object

If PathAndObjectDelete = "" Or IsNull(PathAndObjectDelete) Then
  MsgBox "Valor de Path / Objeto es requerido", vbCritical, "Error en función DeleteObjectSystem"
  DeleteObjectSystem = False
  Exit Function
End If


Set fs = CreateObject("Scripting.FileSystemObject")

Select Case TypeDelete
Case 1
   If PathFileExistsW(StrPtr(PathAndObjectDelete)) Then
      fs.DeleteFolder (PathAndObjectDelete)
      DeleteObjectSystem = True
      
   Else
      MsgBox "Directorio no existe, no pudo ser eliminado", vbInformation, "Error en función DeleteObjectSystem"
      DeleteObjectSystem = False
      Set fs = Nothing
      Exit Function
   End If
Case 2
   If fs.FileExists(PathAndObjectDelete) Then
   
    fs.DeleteFile PathAndObjectDelete, True
    DeleteObjectSystem = True
    
    
   Else
   
     MsgBox "Archivo no existe, no pudo ser eliminado", vbInformation, "Error en función DeleteObjectSystem"
     DeleteObjectSystem = False
     Set fs = Nothing
     Exit Function
   
   End If

Case Else

   MsgBox "Opción no Valida, Verificar TypeObjectSystem seleccionado", vbInformation, "Error en función DeleteObjectSystem"
   DeleteObjectSystem = False
   Set fs = Nothing
   Exit Function
End Select

Set fs = Nothing

Exit_Err_DeleteObjectSystem:
Exit Function

Err_DeleteObjectSystem:
MsgBox Err.Description, vbInformation, "Error en función DeleteObjectSystem"
DeleteObjectSystem = False
Set fs = Nothing
Resume Exit_Err_DeleteObjectSystem
End Function


Public Function CreateTreeFolders2(PathToCreate As String) As Boolean
'// Función de Creado de Arbol de Directorios con ayuda de API declarada.
On Error GoTo Err_CreateTreeFolders2

If PathFileExistsW(StrPtr(PathToCreate)) Then
        CreateTreeFolders2 = True
Else
        Call SHCreateDirectory(ByVal 0&, StrPtr(PathToCreate))
        If PathFileExistsW(StrPtr(PathToCreate)) Then
        'MsgBox "Directorio: " & PathToCreate & Chr(13) & "ha sido creado por primera vez", vbInformation, "Complemento de configuración"
        CreateTreeFolders2 = True
        Else
        MsgBox "Directorio: " & PathToCreate & Chr(13) & "no ha sido creado, verifique su sistema", vbExclamation, "Complemento de configuración"
        CreateTreeFolders2 = False
        End If
End If

Exit_Err_CreateTreeFolders2:
Exit Function
Err_CreateTreeFolders2:
CreateTreeFolders2 = False
Resume Exit_Err_CreateTreeFolders2
End Function



Public Function ValidateInputParameterExecution(CommandLineInput As String) As Boolean
On Error GoTo Err_ValidateInputParameterExecution
Dim i As Integer


AppInvoqued = ""        'Identificación de Entrada = 'App:'
DirectoryLocate = ""    'Identificación de Entrada = 'Dir:'
CaptionRuning = ""      'Identificación de Entrada = 'Cap:'

If InStr(1, CommandLineInput, "App:") = 0 Or _
   InStr(1, CommandLineInput, "Dir:") = 0 Or _
   InStr(1, CommandLineInput, "Cap:") = 0 Then
   
  If InStr(1, CommandLineInput, "App:") = 0 Then
     MsgBox "Aplicación no Definida", vbCritical, "Error de Ejecución"
     ValidateInputParameterExecution = False
     Exit Function
  Else
     If InStr(1, CommandLineInput, "Dir:") = 0 Then
        MsgBox "Directorio no definido", vbCritical, "Error de Ejecución"
        ValidateInputParameterExecution = False
        Exit Function
     Else
           MsgBox "Caption no definido", vbCritical, "Error de Ejecución"
           ValidateInputParameterExecution = False
           Exit Function
     End If
  End If
Else

  If Not (InStr(1, CommandLineInput, "App:") < InStr(1, CommandLineInput, "Dir:") And _
     InStr(1, CommandLineInput, "Dir:") < InStr(1, CommandLineInput, "Cap:")) Then
     
     MsgBox "Sintaxis de parametros Incorrecta" & Chr(13) & _
            "App:[String] Dir:[String] Cap:[String] ", vbCritical, "Error de Ejecución"
     ValidateInputParameterExecution = False
     Exit Function
  End If

End If

AppInvoqued = Trim(Mid(CommandLineInput, InStr(1, CommandLineInput, "App:") + 4, _
              InStr(1, CommandLineInput, "Dir:") - (InStr(1, CommandLineInput, "App:") + 4)))
              
DirectoryLocate = Trim(Mid(CommandLineInput, InStr(1, CommandLineInput, "Dir:") + 4, _
                  InStr(1, CommandLineInput, "Cap:") - (InStr(1, CommandLineInput, "Dir:") + 4)))
              
CaptionRuning = Trim(Mid(CommandLineInput, InStr(1, CommandLineInput, "Cap:") + 4, _
                  Len(CommandLineInput) - (InStr(1, CommandLineInput, "Cap:") + 4)))

If Nulidad(AppInvoqued) = True Then
   MsgBox "Invocación Incorrecta, Ejecutor no definido", vbInformation, "Error de Ejecución"
   ValidateInputParameterExecution = False
   Exit Function
End If

If Nulidad(DirectoryLocate) = True Then
   MsgBox "No ha sido ingresado directorio", vbCritical, "Error de Ejecución"
   ValidateInputParameterExecution = False
   Exit Function
Else
   If CreateTreeFolders2(DirectoryLocate) = False Then
      MsgBox "Error en directorio definido", vbCritical, "Error de Ejecución"
      ValidateInputParameterExecution = False
      Exit Function
   End If
End If


If Nulidad(CaptionRuning) = True Then
   MsgBox "Invocación Incorrecta, Caption no definido", vbInformation, "Error de Ejecución"
   ValidateInputParameterExecution = False
   Exit Function
End If

ValidateInputParameterExecution = True

Exit_Err_ValidateInputParameterExecution:
Exit Function


Err_ValidateInputParameterExecution:
MsgBox Err.Description, vbInformation, "Error de Parametros"
ValidateInputParameterExecution = False
Resume Exit_Err_ValidateInputParameterExecution
End Function

Public Sub DisplayConsultaHTML(TitleWindow As String, FileOutHTML As String, IconWindows As String, _
                               Optional SizeHeightWindow As Integer, Optional SizeWidthWindow As Integer)
On Error GoTo Err_DisplayConsultaHTML
Dim HTMLDisplay As New Dialog12

DoEvents
Unload HTMLDisplay

With HTMLDisplay

        If IsNull(SizeHeightWindow) = False And SizeHeightWindow >= 1680 Then
           .Height = SizeHeightWindow
           .WebBrowser1.Height = SizeHeightWindow - 1680
           .Command1.Top = SizeHeightWindow - 1215
           .Top = Int((Screen.Height - .Height) \ 2)
        End If
        If IsNull(SizeWidthWindow) = False And SizeWidthWindow >= 2000 Then
           .Width = SizeWidthWindow
           .WebBrowser1.Width = SizeWidthWindow - 855
           .Command1.Left = (SizeWidthWindow / 2) - 1000
           
        End If
        
        .Left = Int((Screen.Width - .Width) \ 2)
        .Caption = TitleWindow
        .Icon = LoadPicture()
        .Icon = LoadPicture(IconWindows)
        .WebBrowser1.MenuBar = True
        .WebBrowser1.ToolBar = True
        .WebBrowser1.Navigate FileOutHTML
        .Show 1

End With
Unload HTMLDisplay
DoEvents

Exit_Err_DisplayConsultaHTML:
Exit Sub

Err_DisplayConsultaHTML:
MsgBox Err.Description, vbInformation, "Error en Vista HTML"
Resume Exit_Err_DisplayConsultaHTML
End Sub

Public Function Nulidad(InputString As String) As Boolean
On Error GoTo Err_Nulidad
Dim i As Integer

If Len(InputString) = 0 Or _
   InputString = "" Or _
   IsNull(InputString) Then
   
  Nulidad = True
Else
   For i = 1 To Len(InputString)
       If Mid(InputString, i, 1) <> " " Then
           Nulidad = False
           Exit For
        Else
           Nulidad = True
       End If
   Next i
End If
Exit_Err_Nulidad:
Exit Function
Err_Nulidad:
MsgBox Err.Description, vbExclamation, "Error ejecutando Componente"
Resume Exit_Err_Nulidad
End Function

Public Sub Main()
On Error Resume Next
Dim CommandExecute As String

CommandExecute = Command()
'CommandExecute = "App:cwLPD.exe Dir:C:\cwPRINT Cap:Prueba"
OpenFirst = True

If Nulidad(CommandExecute) = True Then
   MsgBox "Aplicativo requiere parametros de ejecución", vbCritical, "Ejecución Incorrecta"
   End
Else
  If ValidateInputParameterExecution(CommandExecute) = True Then
      SetFile = ""
      frmPrincipalMenu.Caption = AppInvoqued & " - " & CaptionRuning
      frmPrincipalMenu.Show
  Else
    MsgBox "Error de ejecución del componente", vbCritical, "Ejecución Incorrecta"
    End
  End If
End If

End Sub

Public Sub DirListFile(ByRef ListObject As ListBox, ByRef PathShow As String, ByRef TypeFile As String)
On Error GoTo Err_DirListFile
Dim FileObject As Object
Dim FolderObject As Object
Dim FileList As Object
Dim Listing
Dim MultiTypeFile As Variant
Dim NoTypes As Integer
Dim i As Integer

Dim ObjRS  As Object

NoTypes = 0

ListObject.Clear

Set ObjRS = CreateObject("ADODB.Recordset")
ObjRS.CursorType = 2
ObjRS.LockType = 3
ObjRS.Fields.Append "File", 200, 200
ObjRS.Fields.Append "DateLastAccess", 7
ObjRS.Fields.Append "CreateDate", 7
ObjRS.Open



Set FileObject = CreateObject("Scripting.FileSystemObject")
Set FolderObject = FileObject.GetFolder(PathShow)
Set FileList = FolderObject.Files

MultiTypeFile = Split(TypeFile, "|")
NoTypes = UBound(MultiTypeFile)

If NoTypes = 0 Then

        For Each Listing In FileList
           If UCase(Mid(Listing.Name, Len(Listing.Name) - 3, 4)) = UCase(Replace(TypeFile, "*", "")) Then
           ObjRS.AddNew
           ObjRS.Fields("File") = Listing.Name
           ObjRS.Fields("DateLastAccess") = Listing.DateLastAccessed
           ObjRS.Fields("CreateDate") = Listing.DateCreated
           ObjRS.Update
           End If
        Next
Else

For i = 0 To NoTypes
    MultiTypeFile(i) = Replace(CStr(MultiTypeFile(i)), "*", "")
    For Each Listing In FileList
       If UCase(Mid(Listing.Name, Len(Listing.Name) - Len(CStr(MultiTypeFile(i))) + 1, Len(CStr(MultiTypeFile(i))))) = UCase(CStr(MultiTypeFile(i))) Then
       ObjRS.AddNew
       ObjRS.Fields("File") = Listing.Name
       ObjRS.Fields("DateLastAccess") = Listing.DateLastAccessed
       ObjRS.Fields("CreateDate") = Listing.DateCreated
       ObjRS.Update
       End If
    Next
Next i


End If


If ObjRS.BOF = True And ObjRS.EOF = True Then
   Exit Sub
Else

ObjRS.Sort = "CreateDate DESC"
ObjRS.MoveFirst

Do While Not ObjRS.EOF
 ListObject.AddItem ObjRS.Fields("File")
 ObjRS.MoveNext
Loop

End If

ObjRS.Close

Set ObjRS = Nothing


Exit_Err_DirListFile:
Exit Sub


Err_DirListFile:
MsgBox Err.Description, vbInformation, "Error cargando ListFile"
Resume Exit_Err_DirListFile
End Sub


Public Function DateToFile(ByRef FileToEvaluate As String, ByRef TextObject As TextBox, ByRef DateToConsult As DatosDeArchivo) As String
'// Extracción de datos de fechas de un archivo cualquiera que este sea, utilizando tres parametros de entrada y dandonos como resultado un string.
'// Los Parametros de entrada son: Archivo a Evaluar, Objeto Tipo TextBox y tipo de Fecha de consulta
On Error GoTo Err_DateToFile
Dim FileToAcces
Dim File
Dim Drive
Dim Folder
Dim SubFol As TextBox

Dim i As Integer

TextObject.Text = ""
Set SubFol = TextObject

Set FileToAcces = CreateObject("Scripting.FileSystemObject")
Set Drive = FileToAcces.Drives(Mid(FileToEvaluate, 1, 1))
Set Folder = Drive.RootFolder

For i = 1 To Len(Mid(FileToEvaluate, 4, InStr(1, FileToEvaluate, FileToAcces.GetBaseName(FileToEvaluate) & "." & FileToAcces.GetExtensionName(FileToEvaluate)) - 5))
  
  If Mid(Mid(FileToEvaluate, 4, InStr(1, FileToEvaluate, FileToAcces.GetBaseName(FileToEvaluate) & "." & FileToAcces.GetExtensionName(FileToEvaluate)) - 5), i, 1) = "\" Then
    If Dir(Folder & "\" & SubFol.Text, vbDirectory) = "" Then
     SubFol.Text = SubFol.Text & "."
    End If
    Set Folder = Folder.SubFolders(SubFol.Text)
    SubFol.Text = ""
  Else
    SubFol.Text = SubFol.Text & Mid(Mid(FileToEvaluate, 4, InStr(1, FileToEvaluate, FileToAcces.GetBaseName(FileToEvaluate) & "." & FileToAcces.GetExtensionName(FileToEvaluate)) - 5), i, 1)
  End If
Next i

If Dir(Folder & "\" & SubFol.Text, vbDirectory) = "" Then
 SubFol.Text = SubFol.Text & "."
End If

Set Folder = Folder.SubFolders(SubFol.Text)
Set File = Folder.Files(FileToAcces.GetBaseName(FileToEvaluate) & "." & FileToAcces.GetExtensionName(FileToEvaluate))

Select Case DateToConsult
        Case 1
             DateToFile = File.DateCreated
        Case 2
             DateToFile = File.DateLastModified
        Case 3
             DateToFile = File.DateLastAccessed
        Case 4
             DateToFile = Str(File.Size / 1024)
End Select

Set FileToAcces = Nothing
Set Drive = Nothing
Set Folder = Nothing
Set File = Nothing

Exit_Err_DateToFile:
Exit Function

Err_DateToFile:
DateToFile = "File Or Path Not Found"
Resume Exit_Err_DateToFile
End Function

