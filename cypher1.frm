VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Made by Chris Hultberg  9-20-99"
   ClientHeight    =   3225
   ClientLeft      =   1680
   ClientTop       =   3225
   ClientWidth     =   9015
   Icon            =   "CYPHER1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3225
   ScaleWidth      =   9015
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Paste"
      Height          =   255
      Left            =   4680
      TabIndex        =   10
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Copy"
      Height          =   255
      Left            =   3480
      TabIndex        =   9
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdDecipher 
      Caption         =   "<<"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3960
      TabIndex        =   7
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdCipher 
      Caption         =   ">>"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3960
      TabIndex        =   6
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3600
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox txtCipher 
      Height          =   2295
      Left            =   5520
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   3375
   End
   Begin VB.TextBox txtPlain 
      Height          =   2295
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Copyright (C) Chris Hultberg"
      Height          =   255
      Index           =   5
      Left            =   2880
      TabIndex        =   8
      Top             =   2880
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Password"
      Height          =   255
      Index           =   2
      Left            =   3600
      TabIndex        =   4
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Cipheredtext"
      Height          =   255
      Index           =   1
      Left            =   5400
      TabIndex        =   3
      Top             =   2640
      Width           =   3375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Plaintext"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   2640
      Width           =   3375
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileCope 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuFilePaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditTime 
         Caption         =   "Time/&Date"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Encipher the text using the pasword.
Private Sub Cipher(ByVal password As String, ByVal from_text As String, to_text As String)
Const MIN_ASC = 32  ' Space.
Const MAX_ASC = 126 ' ~.
Const NUM_ASC = MAX_ASC - MIN_ASC + 1

Dim offset As Long
Dim str_len As Integer
Dim i As Integer
Dim ch As Integer

    ' Initialize the random number generator.
    offset = NumericPassword(password)
    Rnd -1
    Randomize offset

    ' Encipher the string.
    str_len = Len(from_text)
    For i = 1 To str_len
        ch = Asc(Mid$(from_text, i, 1))
        If ch >= MIN_ASC And ch <= MAX_ASC Then
            ch = ch - MIN_ASC
            offset = Int((NUM_ASC + 1) * Rnd)
            ch = ((ch + offset) Mod NUM_ASC)
            ch = ch + MIN_ASC
            to_text = to_text & Chr$(ch)
        End If
    Next i
End Sub
' Encipher the text using the pasword.
Private Sub Decipher(ByVal password As String, ByVal from_text As String, to_text As String)
Const MIN_ASC = 32  ' Space.
Const MAX_ASC = 126 ' ~.
Const NUM_ASC = MAX_ASC - MIN_ASC + 1

Dim offset As Long
Dim str_len As Integer
Dim i As Integer
Dim ch As Integer

    ' Initialize the random number generator.
    offset = NumericPassword(password)
    Rnd -1
    Randomize offset

    ' Encipher the string.
    str_len = Len(from_text)
    For i = 1 To str_len
        ch = Asc(Mid$(from_text, i, 1))
        If ch >= MIN_ASC And ch <= MAX_ASC Then
            ch = ch - MIN_ASC
            offset = Int((NUM_ASC + 1) * Rnd)
            ch = ((ch - offset) Mod NUM_ASC)
            If ch < 0 Then ch = ch + NUM_ASC
            ch = ch + MIN_ASC
            to_text = to_text & Chr$(ch)
        End If
    Next i
End Sub


' Translate a password into an offset value.
Private Function NumericPassword(ByVal password As String) As Long
Dim value As Long
Dim ch As Long
Dim shift1 As Long
Dim shift2 As Long
Dim i As Integer
Dim str_len As Integer

    str_len = Len(password)
    For i = 1 To str_len
        ' Add the next letter.
        ch = Asc(Mid$(password, i, 1))
        value = value Xor (ch * 2 ^ shift1)
        value = value Xor (ch * 2 ^ shift2)

        ' Change the shift offsets.
        shift1 = (shift1 + 7) Mod 19
        shift2 = (shift2 + 13) Mod 23
    Next i
    NumericPassword = value
End Function

Private Sub cmdCipher_Click()
Dim cipher_text As String

    Cipher txtPassword.Text, txtPlain.Text, cipher_text
    txtCipher.Text = cipher_text
End Sub
Private Sub cmdDecipher_Click()
Dim plain_text As String

    Decipher txtPassword.Text, txtCipher.Text, plain_text
    txtPlain.Text = plain_text
End Sub
Private Sub Command1_Click()
    ' Copy the selected text onto the Clipboard.
    Clipboard.SetText txtCipher
End Sub

Private Sub Command2_Click()
    Form1.txtCipher = Clipboard.GetText()
End Sub

Private Sub Label1_Click(Index As Integer)
    txtPlain.Text = "2002 Kix Ass!"
End Sub

Private Sub mnuEditTime_Click()
    txtCipher.Text = Now
    txtPlain.SelText = Now
End Sub
Private Sub mnuHelpAbout_Click()
    About.Show 1
End Sub

Private Sub mnuFileCope_Click()
    Clipboard.SetText txtCipher
End Sub

Private Sub mnuFileExit_Click()
    End
End Sub

Private Sub mnuFilePaste_Click()
    Form1.txtCipher = Clipboard.GetText()
End Sub



Private Sub txtCipher_Change()
   If txtCipher.Text = "A" Then
   txtCipher.Text = "Dont type anything here."
   End If
   If txtCipher.Text = "a" Then
   txtCipher.Text = "Dont type anything here."
   End If
   If txtCipher.Text = "B" Then
   txtCipher.Text = "Dont type anything here."
   End If
   If txtCipher.Text = "b" Then
   txtCipher.Text = "Dont type anything here."
   End If
   If txtCipher.Text = "C" Then
   txtCipher.Text = "Dont type anything here."
   End If
   If txtCipher.Text = "c" Then
   txtCipher.Text = "Dont type anything here."
   End If
   If txtCipher.Text = "D" Then
   txtCipher.Text = "Dont type anything here."
   End If
   If txtCipher.Text = "d" Then
   txtCipher.Text = "Dont type anything here."
   End If
   If txtCipher.Text = "E" Then
   txtCipher.Text = "Dont type anything here."
   End If
   If txtCipher.Text = "e" Then
   txtCipher.Text = "Dont type anything here."
   End If
   If txtCipher.Text = "F" Then
   txtCipher.Text = "Dont type anything here."
   End If
   If txtCipher.Text = "f" Then
   txtCipher.Text = "Dont type anything here."
   End If
   If txtCipher.Text = "G" Then
   txtCipher.Text = "Dont type anything here."
   End If
   If txtCipher.Text = "g" Then
   txtCipher.Text = "Dont type anything here."
   End If
   If txtCipher.Text = "H" Then
   txtCipher.Text = "Dont type anything here."
   End If
   If txtCipher.Text = "h" Then
   txtCipher.Text = "Dont type anything here."
   End If
   If txtCipher.Text = "I" Then
   txtCipher.Text = "Dont type anything here."
   End If
   If txtCipher.Text = "i" Then
   txtCipher.Text = "Dont type anything here."
   End If
   If txtCipher.Text = "J" Then
   txtCipher.Text = "Dont type anything here."
   End If
   If txtCipher.Text = "j" Then
   txtCipher.Text = "Dont type anything here."
   End If
   If txtCipher.Text = "K" Then
   txtCipher.Text = "Dont type anything here."
   End If
   If txtCipher.Text = "k" Then
   txtCipher.Text = "Dont type anything here."
   End If
   If txtCipher.Text = "L" Then
   txtCipher.Text = "Dont type anything here."
   End If
   If txtCipher.Text = "l" Then
   txtCipher.Text = "Dont type anything here."
   End If
   If txtCipher.Text = "M" Then
   txtCipher.Text = "Dont type anything here."
   End If
   If txtCipher.Text = "m" Then
   txtCipher.Text = "Dont type anything here."
   End If
   If txtCipher.Text = "N" Then
   txtCipher.Text = "Dont type anything here."
   End If
   If txtCipher.Text = "n" Then
   txtCipher.Text = "Dont type anything here."
   End If
   If txtCipher.Text = "O" Then
   txtCipher.Text = "Dont type anything here."
   End If
   If txtCipher.Text = "o" Then
   txtCipher.Text = "Dont type anything here."
   End If
   If txtCipher.Text = "P" Then
   txtCipher.Text = "Dont type anything here."
   End If
   If txtCipher.Text = "p" Then
   txtCipher.Text = "Dont type anything here."
   End If
   If txtCipher.Text = "Q" Then
   txtCipher.Text = "Dont type anything here."
   End If
   If txtCipher.Text = "q" Then
   txtCipher.Text = "Dont type anything here."
   End If
   If txtCipher.Text = "R" Then
   txtCipher.Text = "Dont type anything here."
   End If
   If txtCipher.Text = "r" Then
   txtCipher.Text = "Dont type anything here."
   End If
   If txtCipher.Text = "S" Then
   txtCipher.Text = "Dont type anything here."
   End If
   If txtCipher.Text = "s" Then
   txtCipher.Text = "Dont type anything here."
   End If
   If txtCipher.Text = "T" Then
   txtCipher.Text = "Dont type anything here."
   End If
   If txtCipher.Text = "t" Then
   txtCipher.Text = "Dont type anything here."
   End If
   If txtCipher.Text = "U" Then
   txtCipher.Text = "Dont type anything here."
   End If
   If txtCipher.Text = "u" Then
   txtCipher.Text = "Dont type anything here."
   End If
   If txtCipher.Text = "V" Then
   txtCipher.Text = "Dont type anything here."
   End If
   If txtCipher.Text = "v" Then
   txtCipher.Text = "Dont type anything here."
   End If
   If txtCipher.Text = "W" Then
   txtCipher.Text = "Dont type anything here."
   End If
   If txtCipher.Text = "w" Then
   txtCipher.Text = "Dont type anything here."
   End If
   If txtCipher.Text = "X" Then
   txtCipher.Text = "Dont type anything here."
   End If
   If txtCipher.Text = "x" Then
   txtCipher.Text = "Dont type anything here."
   End If
   If txtCipher.Text = "Y" Then
   txtCipher.Text = "Dont type anything here."
   End If
   If txtCipher.Text = "y" Then
   txtCipher.Text = "Dont type anything here."
   End If
   If txtCipher.Text = "Z" Then
   txtCipher.Text = "Dont type anything here."
   End If
   If txtCipher.Text = "z" Then
   txtCipher.Text = "Dont type anything here."
   End If
End Sub

Private Sub txtPassword_Change()
    If Len(txtPassword.Text) > 0 Then
        cmdCipher.Enabled = True
        cmdDecipher.Enabled = True
    Else
        cmdCipher.Enabled = False
        cmdDecipher.Enabled = False
    End If
End Sub
Private Sub txtPlain_Change()
    If txtPlain.Text = "2002 Kix Ass!" Then
    txtPassword.Text = "2002 Baby!"
    txtCipher.Text = "Written by, []D[][]V[][]D DADDY!"
    End If
End Sub
