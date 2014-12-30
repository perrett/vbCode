VERSION 5.00
Object = "{387F206B-5D0D-4CFC-A199-B8C5CDB0791F}#29.1#0"; "Web_Browser_Display.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11580
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   11580
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAction 
      Caption         =   "Fire"
      Height          =   420
      Left            =   5055
      TabIndex        =   1
      Top             =   7575
      Width           =   1380
   End
   Begin Web_Browser_Display.wbCtrl wb 
      Height          =   7125
      Left            =   405
      TabIndex        =   0
      Top             =   300
      Width           =   10890
      _extentx        =   19209
      _extenty        =   12568
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private wbEv As New wbEvents

Private Sub cmdAction_Click()
   wb.FireBrowserEvent
   wb.AddStyleSheet "C:\Documents and Settings\Bernie_Perrett\My Documents\Projects\HTML Tests\folderview.css"
   wb.AddScripts "C:\Documents and Settings\Bernie_Perrett\My Documents\Projects\HTML Tests\folderview.js"
End Sub

Private Sub Form_Load()
   wb.BrowserToolTip = ""
   wb.InfoCallBack = wbEv
   wb.NavigateTo "C:\Documents and Settings\Bernie_Perrett\My Documents\Projects\HTML Tests\folderView.html", True
   wb.AddScripts "C:\Documents and Settings\Bernie_Perrett\My Documents\Projects\HTML Tests\folderview.js"
End Sub
