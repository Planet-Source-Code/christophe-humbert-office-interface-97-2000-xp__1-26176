VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDI Aspect Test"
   ClientHeight    =   7440
   ClientLeft      =   2160
   ClientTop       =   2310
   ClientWidth     =   10260
   LinkTopic       =   "MDIForm1"
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents SysMenu As SystemMenu
Attribute SysMenu.VB_VarHelpID = -1
Private WithEvents Menubar As Menubar
Attribute Menubar.VB_VarHelpID = -1

Private Children As Collection

Private Sub MDIForm_Load()
    Dim lpItem As MenuItem

    Set Children = New Collection

    Set SysMenu = New SystemMenu
    Set Menubar = New Menubar
    
    Menubar.SetWindowHandle hWnd
    Menus.MenuDrawStyle = mdsOfficeXP
    
    SysMenu.Subclass hWnd
    
    Menubar.Add "&Window", , "WINDOW"
    
    Menubar.Submenu("WINDOW").Add "&New", , "NEW"
    Set lpItem = Menubar.Submenu("WINDOW").Add("Test Item", , "TEST")
    lpItem.Checkmark = True
    lpItem.Checked = True
    
    Set lpItem = Menubar.Submenu("WINDOW").Add("Test Item", , "TEST2")
    lpItem.Checkmark = True
    lpItem.Checked = False
    
    Set lpItem = Menubar.Submenu("WINDOW").Add("Test Item", , "TEST3")
    lpItem.Checkmark = True
    lpItem.Checked = True
    lpItem.Visual.SelectionStyle = mssNoCheckBevel
    
    Set lpItem = Menubar.Submenu("WINDOW").Add("Combo MultiLine and" + vbCrLf + "Right Test Item", , "TEST3B")
    
    lpItem.Checkmark = True
    lpItem.Checked = True
    lpItem.Visual.TextAlign = taRight
    
    Set lpItem = Menubar.Submenu("WINDOW").Add("Test Item", , "TEST4")
    Set lpItem.Picture = Me.Icon
    
    Set lpItem = Menubar.Submenu("WINDOW").Add("-")
    
    Set lpItem = Menubar.Submenu("WINDOW").Add("Switch Office Styles", , "SWITCH")
    lpItem.Visual.SelectionStyle = mssFlat
    
    
End Sub

Public Function AddChild() As Form2
    Dim varNew As New Form2
    
    Set AddChild = varNew
    Children.Add varNew, "_H" + Hex(varNew.hWnd)
    
    varNew.Show
    varNew.Menu.SetWindowHandle hWnd
    
End Function

Public Function RemoveChild(varForm As Form2)
    On Error Resume Next
    
    Children.Remove "_H" + Hex(varForm.hWnd)
    Unload varForm
        
End Function

Private Sub UnloadChildren()
    Dim varObj As Object
        
    For Each varObj In Children
        Unload varObj
    Next varObj
    
    Set Children = Nothing
        
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    UnloadChildren

End Sub


Private Sub Menubar_UserCommand(ByVal Item As StdMenuAPI.MenuItem)
    
    Select Case Item.Key
    
        Case "NEW"
        
            AddChild
            
        Case "SWITCH"
            Menus.MenuDrawStyle = (Menus.MenuDrawStyle Xor mdsOfficeXP)
            
    End Select
    
End Sub
