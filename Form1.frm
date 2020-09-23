VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "HTML Encrypt"
   ClientHeight    =   930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3360
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   930
   ScaleWidth      =   3360
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar ThiagoBar 
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   840
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   1320
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Open"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   0
      TabIndex        =   4
      Text            =   "; var s= u(md);document.write (s);</script>"
      Top             =   2280
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   -120
      TabIndex        =   3
      Text            =   """"
      Top             =   1920
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   -120
      TabIndex        =   2
      Text            =   "md = """
      Top             =   1560
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   -6120
      TabIndex        =   1
      Text            =   $"Form1.frx":27A2
      Top             =   1200
      Visible         =   0   'False
      Width           =   12120
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Encrypt"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":283A
      ForeColor       =   &H8000000F&
      Height          =   135
      Left            =   0
      TabIndex        =   9
      Top             =   8640
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":2919
      Enabled         =   0   'False
      Height          =   135
      Left            =   11760
      TabIndex        =   7
      Top             =   8400
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' CRIADO POR THIAGO SANTOS SILVA RIBEIRO DE SOUZA
'                 THIAGO@XMAIL.NET
' TODOS OS DIREITOS RESERVADOS - ALL RIGHTS RESERVED
'                    Versão 1.1
Option Explicit

Public Thiagao As String
Public Thiagao2 As String

Private Function EncryptarDoThiago(Frase As String, lLEn As Long) As String
'----------------------------------------------------
' Encrypt the String ( The String is Text in Text1
' that was loaded with the Command3_Click() )...
'----------------------------------------------------

    Dim i As Long
    Dim NovoCaracter As Long
    Dim K
    i = 1

    Do Until i = lLEn + 1
        NovoCaracter = Asc(Mid(Frase, i, 1)) + 3
        EncryptarDoThiago = EncryptarDoThiago + Chr(NovoCaracter)
        i = i + 1
        K = (i * 100) / lLEn
        If K < 100 Then
        ThiagoBar.Value = K
        Else
        End If
    Loop
End Function


Private Sub Command2_Click()
'---------------------------------------------------
' Encrypt the String ( The String is Text in Text1
' that was loaded with the Command3_Click() )...
' and add a little more lines... to be decrypted
' by the browser. ( this lines are hide in textboxs )
' if you enlarge the form you´ll see the textboxs.
'---------------------------------------------------
 On Error GoTo erro
 
    Dim Frase As String
    Dim sEncryptarDoThiagoed As String
    Dim sDecryptarDoThiagoed As String
    Dim Thiago
    Dim T001
    Dim T002
    Dim T003
    Dim T004
        
    T001 = Text5.Text
    T002 = Text6.Text
    T003 = Text7.Text
    T004 = Text8.Text
    
    Frase = Thiagao
    sEncryptarDoThiagoed = EncryptarDoThiago(Frase, Len(Frase))
    Thiago = T001 & T002 & sEncryptarDoThiagoed & T003 & T004
 
Thiagao2 = Thiago
MsgBox "File Was Successfully Encrypted"
Command4.Enabled = True

erro:
' Do something when you got some error...
' i prefer do nothing.
  
End Sub

Private Sub Command3_Click()
'----------------------------------------------------
' Open The file that you want... ans set the
' Text3.Text = "The File Lines" ( as the text that
' was founded in the file.
'----------------------------------------------------
 Dim sArquivo As String
    Dim NomedoArquivo As String
    
    NomedoArquivo = CommonDialog2.FileName
    CommonDialog2.Filter = "Htm Files (*.htm)|*.htm|Html Files (*.html)|*.html|All Files (*.*)|*.*|"
    CommonDialog2.ShowOpen
    
    If NomedoArquivo = CommonDialog2.FileName Then
        If MsgBox("you need to choose a file to encrypt it!!!", vbOK) = vbNo Then
            Exit Sub
        End If
    End If
    
    NomedoArquivo = CommonDialog2.FileName
    
    Dim intFileNum As Integer
    Dim strText As String
    Dim strTemp As String
       If Trim(NomedoArquivo) <> "" Then

        intFileNum = FreeFile

        Open NomedoArquivo For Input As #intFileNum
            Do Until EOF(intFileNum)
            DoEvents
            Line Input #intFileNum, strTemp
            If strText = "" Then
                strText = strTemp
            Else
            strText = strText + strTemp
            Thiagao = strText
            End If
        Loop
            Thiagao = strText
        Close #intFileNum
    End If
    
' -----------------------------------
MsgBox "File Was Successfully Loaded"
Command2.Enabled = True
Command4.Enabled = False


erro:
' Do something when you got some error...
' i prefer do nothing.
End Sub

Private Sub Command4_Click()
'----------------------------------------------------
' Save the Text4 Text as the file that you want.
' Save .HTM to the work perfectly.
'----------------------------------------------------
Dim fname
ThiagoBar.Value = 0
On Error GoTo erro
CommonDialog1.Filter = "Htm Files (*.htm)|*.htm|Html Files (*.html)|*.html|All Files (*.*)|*.*|"
CommonDialog1.ShowSave
fname = CommonDialog1.FileName

Dim txt As String
txt = Thiagao2
Open fname For Output As #1
Print #1, txt
Close #1
MsgBox "File Was Successfully Saved"
Exit Sub
erro:
' Do something when you got some error...
' i prefer do nothing.
End Sub

Private Sub Form_Load()
'-----------------------------------------------------
' Agradeço ao criador do Usha's HTML Encryptor,
' que foi a minha base... com análise neste programa
' consegui desenvolver meu próprio encryptador HTML.
' E ao Jaime Muscatelli, que foi com base no programa
' de encryptação de textos que usei como base para
' conseguir desenvolver meu próprio programa de
' encryptografia de Códigos Fontes em HTML.
'------------------------------------------------------
'
'             SORRY FOR THE BAD ENGLISH
'
' I am thankful the creator of Usha's HTML Encryptor,
' that was my base... with analysis in this program
' obtained to develop my proper encryptation HTML and
' to the Jaime Muscatelli, who was on the basis of the
' program of encryptation of texts that I used as base
' to obtain to develop my proper program of
' encryptation of Codes Sources (HTML).
'------------------------------------------------------
'*********************README**************************
' When you click on open button, and select the file,
' this file is loaded and we put the text of the file
' into thiagao var. Then we encrypt the text and
' put it into thiagao2 var, so we take the text of
' the thiagao2 and save it using the save button.
' So if you take a look at the form, you´ll see a lot
' of hide textboxs ... Note: The Properties of the
' Text Boxs are Visible.False. But you i´ll see at
' visual basic.
'
' If you need something... THIAGO@XMAIL.NET
'
'*****************************************************
' If you to use this program, please give me some
' credits
'*****************************************************


End Sub

