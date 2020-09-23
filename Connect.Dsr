VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   8880
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   9945
   _ExtentX        =   17542
   _ExtentY        =   15663
   _Version        =   393216
   Description     =   "Transparent Form Maker Creates irregular shaped forms and adds them to your Visual Basic project."
   DisplayName     =   "SSE TransForm Maker"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "None"
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public FormDisplayed          As Boolean
Public VBInstance              As VBIDE.VBE
Attribute VBInstance.VB_VarHelpID = -1
Dim mcbMenuCommandBar         As Office.CommandBarControl
Dim mfrmAddIn                 As New frmMain
Public WithEvents MenuHandler As CommandBarEvents          'command bar event handler
Attribute MenuHandler.VB_VarHelpID = -1
'Public WithEvents ToolbarHandler As CommandBarEvents
'Public WithEvents CustomHandler As CommandBarEvents

Sub Hide()
    Dim frm As Form
    
    On Error Resume Next
    
    FormDisplayed = False
    For Each frm In Forms
        frm.Hide
    Next frm
   
End Sub

Sub Show()
  
    On Error Resume Next
    
    
    If mfrmAddIn Is Nothing Then
        Set mfrmAddIn = New frmMain
    End If
    
    
    'Set mfrmAddIn.Connect = Me
    FormDisplayed = True
    mfrmAddIn.Show
   
End Sub

'------------------------------------------------------
'this method adds the Add-In to VB
'------------------------------------------------------
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    'On Error GoTo error_handler
        
    'save the vb instance
    Set VBInstance = Application
    Set objVBE = VBInstance
    
    'this is a good place to set a breakpoint and
    'test various addin objects, properties and methods
    Debug.Print VBInstance.FullName

    If ConnectMode = ext_cm_External Then
        'Used by the wizard toolbar to start this wizard
        Me.Show
    Else
        Set mcbMenuCommandBar = AddToAddInCommandBar("SSE Transform")
        'sink the event
        Set Me.MenuHandler = VBInstance.Events.CommandBarEvents(mcbMenuCommandBar)
            
                
    End If
    
        
    If ConnectMode = ext_cm_AfterStartup Then
       
    End If
  
    Exit Sub
    
error_handler:
    
    MsgBox Err.Description
    
End Sub

'------------------------------------------------------
'this method removes the Add-In from VB
'------------------------------------------------------
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    On Error Resume Next
    Dim frm As Form
    
    
    'delete the command bar entry
    mcbMenuCommandBar.Delete
   
    'shut down the Add-In
    If FormDisplayed Then
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "1"
        FormDisplayed = False
    Else
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "0"
    End If
    
    Unload mfrmAddIn
    For Each frm In Forms
        Unload frm
    Next frm
    
    Set mfrmAddIn = Nothing

End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
    'If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
        'set this to display the form on connect
        'Me.Show
    'End If
End Sub

'Private Sub CustomHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
'CheckToolbars
'
'End Sub

'this event fires when the menu is clicked in the IDE
Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Me.Show
End Sub

Function AddToAddInCommandBar(sCaption As String) As Office.CommandBarControl
    Dim cbMenuCommandBar As Office.CommandBarControl  'command bar object
    Dim cbMenu As Object
    Dim picButton As StdPicture
    
    On Error GoTo AddToAddInCommandBarErr
    
    'see if we can find the Add-Ins menu
    Set cbMenu = VBInstance.CommandBars("Add-Ins")
    If cbMenu Is Nothing Then
        'not available so we fail
        Exit Function
    End If
    
    'add it to the command bar
    Set cbMenuCommandBar = cbMenu.Controls.Add(1)
    'set the caption
    cbMenuCommandBar.Caption = sCaption
    
    Set AddToAddInCommandBar = cbMenuCommandBar
    
    Exit Function
    
AddToAddInCommandBarErr:

End Function


