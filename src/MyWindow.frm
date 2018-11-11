VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MyWindow 
   Caption         =   "Добавить строки в спецификацию"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5850
   OleObjectBlob   =   "MyWindow.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MyWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Written in 2015-2018 by Eduard E. Tikhenko <aquaried@gmail.com>
'
'To the extent possible under law, the author(s) have dedicated all copyright
'and related and neighboring rights to this software to the public domain
'worldwide. This software is distributed without any warranty.
'You should have received a copy of the CC0 Public Domain Dedication along
'with this software.
'If not, see <http://creativecommons.org/publicdomain/zero/1.0/>

Option Explicit

Private Sub boxCenter_Click()
    SaveSettingFormatStrings idCenter
End Sub

Private Sub boxLeft_Click()
    SaveSettingFormatStrings idLeft
End Sub

Private Sub btnRows_Click()
    EditConfigFile
    ExitApp
End Sub

Private Sub empBox_Click()
    SaveSettingEmptyStrings empBox.value
End Sub

Private Sub exitBut_Click()
    ExitApp
End Sub

Private Sub okBut_Click()
    Run
    ExitApp
End Sub

Private Sub underlineBox_Click()
    SaveSettingFormatTitles idUnderline
End Sub

Private Sub upperBox_Click()
    SaveSettingFormatTitles idUpper
End Sub

Private Sub UserForm_Initialize()
    'me.lstTitles.Width =

    Select Case GetSettingFormatTitles
        Case idUnderline
            underlineBox.value = True
        Case idUpper
            upperBox.value = True
    End Select
    empBox.value = GetSettingEmptyStrings()
    Select Case GetSettingFormatStrings
        Case idLeft
            Me.boxLeft.value = True
        Case idCenter
            Me.boxCenter.value = True
    End Select
End Sub
