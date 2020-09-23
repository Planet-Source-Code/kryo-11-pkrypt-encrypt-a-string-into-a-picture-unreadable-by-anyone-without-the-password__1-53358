VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "pKrypt"
   ClientHeight    =   6075
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5175
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   405
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   345
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPw 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   840
      PasswordChar    =   "*"
      TabIndex        =   13
      Top             =   4260
      Width           =   4275
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   4020
      Top             =   3540
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdDecrypt 
      Caption         =   "Decrypt"
      Height          =   315
      Left            =   1320
      TabIndex        =   6
      Top             =   3900
      Width           =   1215
   End
   Begin VB.CommandButton cmdEncrypt 
      Caption         =   "Encrypt"
      Height          =   315
      Left            =   60
      TabIndex        =   5
      Top             =   3900
      Width           =   1215
   End
   Begin VB.TextBox txtData 
      Height          =   1155
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   4560
      Width           =   5055
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   315
      Left            =   2580
      TabIndex        =   3
      Top             =   3540
      Width           =   1395
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   315
      Left            =   60
      TabIndex        =   2
      Top             =   3540
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Top             =   3540
      Width           =   1215
   End
   Begin VB.PictureBox picCont 
      BackColor       =   &H00FFFFFF&
      Height          =   3435
      Left            =   60
      ScaleHeight     =   225
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   333
      TabIndex        =   0
      Top             =   60
      Width           =   5055
      Begin VB.VScrollBar VScroll1 
         Enabled         =   0   'False
         Height          =   3135
         Left            =   4740
         TabIndex        =   10
         Top             =   0
         Width           =   255
      End
      Begin VB.HScrollBar HScroll1 
         Enabled         =   0   'False
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   3120
         Width           =   4755
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   4740
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   11
         Top             =   3120
         Width           =   315
      End
      Begin VB.PictureBox picData 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000001&
         BorderStyle     =   0  'None
         Height          =   15
         Left            =   0
         ScaleHeight     =   1
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   1
         TabIndex        =   8
         Top             =   0
         Visible         =   0   'False
         Width           =   15
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Password:"
      Height          =   255
      Left            =   60
      TabIndex        =   12
      Top             =   4260
      Width           =   1035
   End
   Begin VB.Label Status 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Idle"
      Height          =   255
      Left            =   60
      TabIndex        =   7
      Top             =   5760
      Width           =   5055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Private Sub cmdClear_Click()
    Set picData.Picture = Nothing 'clear picture
    picData.Cls 'clear image
    picData.Visible = False 'hide picturebox
    Form_Resize 'call form resize event
End Sub

Private Sub cmdDecrypt_Click()
    Dim DSTR As String
    
    If picData.Visible Then 'if picture is hidden there is nothing to decrypt
    
        Status = "Verifying Password..." 'status message
        
        'decrypt password
        For i = 0 To picData.ScaleWidth 'each pixel in picturebox (horizontal)
        
            For j = 0 To 9 'check first 10 pixels (vertical)
                
                'the backcolor of the picture is almost black.
                'the difference cannot be seen with the human eye
                'but the program will detect it.
                If GetPixel(picData.hdc, i, j) = 0 Then 'if the pixel is black record it
                    DSTR = DSTR & j
                End If
                
                DoEvents 'to prevent freezing
                
            Next j
            
        Next i
        
        'the above function turns those black pixels into a series of numbers
        'example: 303446210211346541
        'the changes it back into letters
        DSTR = AscDecrypt(DSTR)
        
        If DSTR = txtPw Then 'if it is the same as the password then continue
        
            DSTR = "" 'clear the variable
            Status = "Decrypting Picture..." 'status message
            
            For i = 0 To picData.ScaleWidth 'each pixel (horizontal)
            
                For j = 10 To 19 'check the next 10 pixels
                
                'the backcolor of the picture is almost black.
                'the difference cannot be seen with the human eye
                'but the program will detect it.
                    If GetPixel(picData.hdc, i, j) = 0 Then 'if pixel is black record it
                        DSTR = DSTR & j - 10
                    End If
                    
                    DoEvents 'to prevent freezing
                    
                Next j
                
            Next i
            
            DSTR = AscDecrypt(DSTR) 'change number string back into letters
            txtData = DSTR 'show the string in the textbox
            
        Else
        
            txtData = "Invalid Password." 'wrong password
            
        End If
            
    End If
    
    Status = "Idle" 'status message
End Sub

Private Sub cmdEncrypt_Click()
    Dim ESTR As String, TChr As Integer, FH As Integer, RP As Integer
    
    If Len(txtData) = 0 Then Exit Sub 'if there is nothing to encrypt exit sub
    
    If Len(txtPw) = 0 Then 'if no password then exit sub
        MsgBox "You must give your picture a password", vbInformation, "No Password Detected"
        txtPw.SetFocus
        Exit Sub
    End If
    
    picData.Visible = True 'unhide the picbox
    RP = 0 'function need to be done 2 times to work correctly, this variable shows how many repititions
    Status = "Encrypting String..." 'status message
    DoEvents 'make sure status is shown
    
RepProc: 'beginning of procedure
    RP = RP + 1 'incriment count
    Set picData.Picture = Nothing 'clear picture
    picData.Cls 'clear image
    ESTR = AscEncrypt(txtPw) 'turn each letter of password into number 0-9
    
    If Len(ESTR) > 0 Then 'if there is a string to encrypt continue
    
        For i = 1 To Len(ESTR) 'each number in the string
        
            TChr = Mid(ESTR, i, 1) 'get single number
            If TChr >= FH Then FH = TChr + 1 'find true height
            picData.Left = 0 - (i - 1) 'make sure pixel is visible
            picData.Top = 0 - (TChr - 1) 'make sure pixel is visible
            SetPixel picData.hdc, i - 1, CLng(TChr), &H0& 'change the pixel to black
            
        Next i
        
        picData.Left = 0: picData.Top = 0 'reset picbox position
        picData.Width = IIf(Len(ESTR) > Len(AscEncrypt(txtData)), Len(ESTR) + 1, Len(AscEncrypt(txtData)) + 1) 'set apporpriate width
        picData.Height = FH 'set appropriate height
        
    End If
    
    ESTR = AscEncrypt(txtData) 'get encryption for the data to encrypt
    
    If Len(ESTR) > 0 Then 'if there is a string continue
    
        For i = 1 To Len(ESTR) 'each number
        
            TChr = Mid(ESTR, i, 1) + 10 'get singe number and add 10
            If TChr + 10 >= FH Then FH = TChr + 11 'find true height
            picData.Left = 0 - (i - 1) 'make sure pixel is visible
            picData.Top = 0 - (TChr - 1) 'make sure pixel is visible
            SetPixel picData.hdc, i - 1, CLng(TChr), &H0& 'change pixel to black
            
        Next i
        
            picData.Left = 0: picData.Top = 0 'reset pic position
            picData.Height = FH 'set appropriate height
            
            'now that height and width have been set the procedure needs to be repeated
            'to ensure every pixel is visible
            If RP = 1 Then
                GoTo RepProc
            End If
            
    End If
    
    txtPw = "" 'clear password box
    txtData = "" 'clear data box
    Status = "Idle" 'status message
    Form_Resize 'call resize
End Sub

Private Sub cmdLoad_Click()
    On Error GoTo 1
    
    Dim TPic As IPictureDisp
    
    With CD1
        .CancelError = True
        .DialogTitle = "Load Picture"
        .Filter = "Bitmaps|*.bmp"
        .ShowOpen
        
        If Dir(.FileName) <> "" Then
            Set TPic = LoadPicture(.FileName)
            picData.Visible = True
            picData.AutoSize = True
            picData.Picture = TPic
            picData.AutoSize = False
        End If
        
    End With
    
1:  Form_Resize
End Sub

Private Sub cmdSave_Click()
    On Error GoTo 1
    
    If picData.Visible = False Then Exit Sub
    
    Dim TPic As IPictureDisp
    
    Set TPic = picData.Image
    
    With CD1
        .CancelError = True
        .DialogTitle = "Save Encrypted Picture"
        .Filter = "Bitmap|*.bmp"
        .ShowSave
        
        If Dir(.FileName) <> "" Then
            Dim answ
            answ = MsgBox("The file " & .FileTitle & " already exists." & vbCrLf & "Are you sure you want to overwrite it?", vbYesNo, "Confirm Overwrite")
            
            If answ = vbYes Then
                SavePicture TPic, .FileName
            End If
            
        Else
        
            SavePicture TPic, .FileName
            
        End If
        
    End With
    
1:  Form_Resize
End Sub

Private Sub Form_Resize()
    'resize event: aligns each control and fixes their width and height
    On Error Resume Next
    Dim cmdWidth(1) As Long
    Dim A As Long, B As Long, C As Long, D As Long, E As Long
    
    txtPw.Height = 19
    picCont.Width = ScaleWidth - 8
    picCont.Height = Me.ScaleHeight - 176
    
    A = picCont.Top + picCont.Height + 4
    B = A + 25: C = B + 25: D = C + 23: E = D + 77
    cmdWidth(0) = (Me.ScaleWidth - 16) / 3: cmdWidth(1) = (Me.ScaleWidth - 12) / 2
    
    cmdLoad.Width = cmdWidth(0): cmdSave.Width = cmdWidth(0): cmdClear.Width = cmdWidth(0)
    cmdEncrypt.Width = cmdWidth(1): cmdDecrypt.Width = cmdWidth(1)
    
    cmdLoad.Left = 4: cmdSave.Left = cmdLoad.Left + cmdLoad.Width + 4: cmdClear.Left = cmdSave.Left + cmdSave.Width + 4
    cmdEncrypt.Left = 4: cmdDecrypt.Left = cmdEncrypt.Left + cmdEncrypt.Width + 4
    cmdLoad.Top = A: cmdSave.Top = A: cmdClear.Top = A: cmdEncrypt.Top = B: cmdDecrypt.Top = B
    Label1.Top = C: txtPw.Top = C: txtPw.Width = ScaleWidth - 8 - txtPw.Left
    txtData.Top = D: txtData.Width = ScaleWidth - 8
    Status.Top = E: Status.Width = ScaleWidth - 8
    
    HScroll1.Top = picCont.ScaleHeight - HScroll1.Height
    HScroll1.Width = picCont.ScaleWidth - VScroll1.Width
    VScroll1.Left = picCont.ScaleWidth - VScroll1.Width
    VScroll1.Height = picCont.ScaleHeight - HScroll1.Height
    Shape1.Top = VScroll1.Height: Shape1.Left = HScroll1.Width
    
    If picData.Width > picCont.ScaleWidth - VScroll1.Width And picData.Visible = True Then
        HScroll1.Max = picData.Width - picCont.ScaleWidth - HScroll1.Height
        HScroll1.SmallChange = HScroll1.Max / 100
        HScroll1.LargeChange = HScroll1.Max / 10
        HScroll1.Enabled = True
    Else
        HScroll1.Enabled = False
    End If
    
    If picData.Height > picCont.ScaleHeight - HScroll1.Height And picData.Visible = True Then
        VScroll1.Max = picData.Height - picCont.ScaleHeight - VScroll1.Width
        VScroll1.SmallChange = VScroll1.Max / 100
        VScroll1.LargeChange = VScroll1.Max / 10
        VScroll1.Enabled = True
    Else
        VScroll1.Enabled = False
    End If
End Sub

Private Sub HScroll1_Scroll()
    picData.Left = 0 - HScroll1.Value
End Sub

Private Sub VScroll1_Scroll()
    picData.Top = 0 - VScroll1.Value
End Sub

Public Function AscEncrypt(StringToEncrypt As String) As String
    On Error GoTo ErrorHandler
    Dim Char As String
    AscEncrypt = ""
    
    For i = 1 To Len(StringToEncrypt)
        Char = Asc(Mid(StringToEncrypt, i, 1))
        AscEncrypt = AscEncrypt & Len(Char) & Char
    Next i
    
    Exit Function
    
ErrorHandler:
    AscEncrypt = "Error encrypting string"
End Function

Public Function AscDecrypt(StringToDecrypt As String) As String
    On Error GoTo ErrorHandler
    Dim CharCode As String
    Dim CharPos As Integer
    Dim Char As String
    
    AscDecrypt = ""
    
    Do
    
        CharPos = Left(StringToDecrypt, 1)
        StringToDecrypt = Mid(StringToDecrypt, 2)
        CharCode = Left(StringToDecrypt, CharPos)
        StringToDecrypt = Mid(StringToDecrypt, Len(CharCode) + 1)
        AscDecrypt = AscDecrypt & Chr(CharCode)
        
    Loop Until StringToDecrypt = ""
    
    Exit Function
    
ErrorHandler:
    AscDecrypt = "Error decrypting string"
End Function
