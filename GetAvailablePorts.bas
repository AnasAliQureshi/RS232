Attribute VB_Name = "GetAvailablePorts"
Option Explicit
 
Private Declare Function CreateFile _
 Lib "kernel32.dll" Alias "CreateFileA" ( _
 ByVal lpFileName As String, _
 ByVal dwDesiredAccess As Long, _
 ByVal dwShareMode As Long, _
 lpSecurityAttributes As SECURITY_ATTRIBUTES, _
 ByVal dwCreationDisposition As Long, _
 ByVal dwFlagsAndAttributes As Long, _
 ByVal hTemplateFile As Long) As Long
 
Private Declare Function CloseHandle _
 Lib "kernel32.dll" ( _
 ByVal hObject As Long) As Long
 
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
 
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const OPEN_EXISTING = 3
Private Const FILE_ATTRIBUTE_NORMAL = &H80
 
Private Function COMAvailable(iPortNum As Integer) As Boolean
    Dim hCOM As Long
    Dim ret As Long
    Dim sec As SECURITY_ATTRIBUTES
 
    'try to open the COM port
    hCOM = CreateFile("COM" & iPortNum & "", 0, FILE_SHARE_READ + FILE_SHARE_WRITE, _
     sec, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    If hCOM = -1 Then
        COMAvailable = False
    Else
        COMAvailable = True
        'close the COM port
        ret = CloseHandle(hCOM)
    End If
End Function
Public Function AddPortstoCombo(CmbBox As ComboBox, ComCTRL As MSComm)
On Error GoTo Port_Error
CmbBox.Clear
Dim i As Integer
Dim pass As Integer
For i = 1 To 16
    pass = 1
    ComCTRL.CommPort = i
    ComCTRL.PortOpen = True
    If pass = 1 Then
        CmbBox.AddItem "COM" & i
        ComCTRL.PortOpen = False
    End If
        ComCTRL.PortOpen = False
Next i
Port_Error:
pass = 0
Resume Next
End Function

