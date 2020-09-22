VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   5010
   ClientLeft      =   4035
   ClientTop       =   3405
   ClientWidth     =   7245
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   5010
   ScaleWidth      =   7245
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents SysMenu As SystemMenu
Attribute SysMenu.VB_VarHelpID = -1

Private WithEvents Menubar As Menubar
Attribute Menubar.VB_VarHelpID = -1

Private Sub Form_Load()
    Set SysMenu = New SystemMenu
    
    SysMenu.Subclass hWnd
    Set Menubar = New Menubar
    
    Menubar.Add "&Window", , "WINDOW"
    
    Menubar.Submenu("WINDOW").Add "&New", , "NEW"
    
End Sub

Public Property Get Menu() As Menubar
    Set Menu = Menubar
End Property

Private Sub Menubar_UserCommand(ByVal Item As StdMenuAPI.MenuItem)
    
    If (Item.Key = "NEW") Then
        
        MDIForm1.AddChild
    End If
    
End Sub

