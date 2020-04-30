VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserInfo 
   Caption         =   "User Information"
   ClientHeight    =   2775
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5055
   OleObjectBlob   =   "UserInfo.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ContinueButton_Click()

If (InName.Value <> "") Then
UserFullName = InName.Value
Else: UserFullName = UserName
End If

UserClass = InClass.Value
UserLab = InLab.Value

If (InProduct.Value <> "") Then
UserProductName = InProduct.Value
Else: UserProductName = "Product1"
End If

Unload Me
End Sub
