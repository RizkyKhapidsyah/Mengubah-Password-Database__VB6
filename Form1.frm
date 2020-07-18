VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Mengubah Password Database"
   ClientHeight    =   3870
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   ScaleHeight     =   3870
   ScaleWidth      =   6390
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Jika Anda mendapat pesan "Unrecognized Database 'format", kemungkinan Anda menggunakan Access 2000 dan 'Anda tidak mempunyai Microsoft DAO 3.6 Object Library
'pilih file C:\Program Files\Common Files\Microsoft 'Shared\Dao\dao360.dll
'Jika di komputer Anda terinstall Access 2000, Anda 'mempunyai file ini.

Private Sub ChangeAccessPassword(OldPass As String, NewPass As String)
    Dim Db As Database
    'Buka dataase, menggunakan password yang lama.
    'Ganti "C:\MyDir\Mydb1.mdb" dengan nama file
    'database Anda
    Set Db = OpenDatabase(App.Path + "\Mydb1.mdb", True, False, ";pwd=" & OldPass)
    'Ganti menjadi password baru
    Db.NewPassword OldPass, NewPass
    'Tutup database
    Db.Close
End Sub

Private Sub Command1_Click()
'Ganti "oldPassword" dengan password database, dan
'"newPassword" dengan password baru yang Anda inginkan.
    Call ChangeAccessPassword("oldPassword", "newPassword")
End Sub


