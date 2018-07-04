VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Read BOM device value to 3070board format..1.0"
   ClientHeight    =   3450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8145
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   8145
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   7935
      Begin VB.CommandButton mcGo 
         Caption         =   "&GO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6600
         TabIndex        =   1
         Top             =   280
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Pin Linrary"
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   5280
         TabIndex        =   22
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox CheckR 
         Caption         =   "Resistor"
         Height          =   495
         Left            =   1560
         TabIndex        =   8
         Top             =   240
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox CheckD 
         Caption         =   "Diode"
         Height          =   495
         Left            =   2760
         TabIndex        =   7
         Top             =   240
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox CheckCn 
         Caption         =   "Connector"
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   3840
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox CheckC 
         Caption         =   "Catacitor"
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Value           =   1  'Checked
         Width           =   1215
      End
   End
   Begin VB.TextBox txtBomPath 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "Please open bom file!(DblClick me open file!)"
      Top             =   240
      Width           =   6735
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   960
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Height          =   1815
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   7935
      Begin VB.CommandButton Command1 
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7560
         TabIndex        =   2
         Top             =   1440
         Width           =   255
      End
      Begin VB.TextBox txtDH 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   25
         Text            =   "0.9"
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txtDL 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4080
         MaxLength       =   3
         TabIndex        =   24
         Text            =   "0.5"
         Top             =   840
         Width           =   495
      End
      Begin VB.CheckBox CheckIND 
         Caption         =   "Inductance and Fuse to Jumper test"
         Height          =   375
         Left            =   4560
         TabIndex        =   23
         Top             =   1320
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.TextBox txtJumper 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3720
         MaxLength       =   2
         TabIndex        =   20
         Text            =   "11"
         Top             =   1320
         Width           =   375
      End
      Begin VB.CheckBox CheckJumper 
         Caption         =   "Resistor value low to Jumper"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   1320
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.TextBox txtRL 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   7200
         MaxLength       =   3
         TabIndex        =   17
         Text            =   "10"
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txtRH 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   7200
         MaxLength       =   3
         TabIndex        =   16
         Text            =   "10"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtCL 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4080
         MaxLength       =   3
         TabIndex        =   13
         Text            =   "30"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtCH 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   12
         Text            =   "30"
         Top             =   360
         Width           =   495
      End
      Begin VB.Line Line4 
         X1              =   5400
         X2              =   5520
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label2 
         Caption         =   "Resistor"
         Height          =   255
         Left            =   4800
         TabIndex        =   28
         Top             =   720
         Width           =   615
      End
      Begin VB.Line Line3 
         X1              =   5520
         X2              =   5640
         Y1              =   600
         Y2              =   480
      End
      Begin VB.Line Line2 
         X1              =   5520
         X2              =   5640
         Y1              =   960
         Y2              =   1080
      End
      Begin VB.Line Line1 
         X1              =   5520
         X2              =   5520
         Y1              =   960
         Y2              =   600
      End
      Begin VB.Label Label10 
         Caption         =   "Diode High Limit:"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "Diode Low Limit:"
         Height          =   255
         Left            =   2520
         TabIndex        =   26
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "o"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4200
         TabIndex        =   21
         Top             =   1440
         Width           =   135
      End
      Begin VB.Label Label7 
         Caption         =   "Value low >"
         Height          =   255
         Left            =   2760
         TabIndex        =   19
         Top             =   1400
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Resistor Low Limit:"
         Height          =   255
         Left            =   5640
         TabIndex        =   15
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Resistor High Limit:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5640
         TabIndex        =   14
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Catacitor Low Limit:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2520
         TabIndex        =   11
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Catacitor High Limit:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Label Label1 
      Caption         =   "BomPath:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim BomPath As String
Dim PrmPath As String
Dim bListCatacitor As Boolean
Dim bListResistor As Boolean
Dim bListDiode As Boolean
Dim strCH As String
Dim strCL As String
Dim strRH As String
Dim strRL As String
Dim strDH As String
Dim strDL As String

Private Sub CheckJumper_Click()
If CheckJumper.Value = 1 Then
  txtJumper.Enabled = True
  Else
   
  txtJumper.Enabled = False
End If
End Sub

Private Sub CheckR_Click()
If CheckR.Value = 1 Then
  CheckJumper.Enabled = True
  'txtJumper.Enabled = True
  Else
 CheckJumper.Enabled = False
  txtJumper.Enabled = False
   CheckJumper.Value = 0
  
End If
End Sub

Private Sub Command1_Click()
Unload Me
End
End Sub

Private Sub Form_Load()
On Error Resume Next
 
 
PrmPath = App.Path
If Right(PrmPath, 1) <> "\" Then PrmPath = PrmPath & "\"
MkDir PrmPath & "ReadBomValue"

End Sub

Private Sub Read_BomFile()
Dim Mystr As String
Dim intI As String
Dim strDeviceName As String
Dim TmpStr() As String
Dim tmpSTR1 As String
Dim DeviceType_ As String
Dim DeviceType_A As String
Dim DeviceType_1 As String
Dim CValue As String
Dim RValue As String
Dim strCAP() As String
Dim strRES() As String
Dim strReadText As String
Dim LowToJumper

strCH = txtCH.Text
strCL = txtCL.Text
strRH = txtRH.Text
strRL = txtRH.Text
strDH = txtDH.Text
strDL = txtDH.Text


If CheckC.Value = 1 Then
   bListCatacitor = True
   Else
   bListCatacitor = False
End If

 
If CheckR.Value = 1 Then
   bListResistor = True
   Else
   bListResistor = False
End If
If CheckD.Value = 1 Then
   bListDiode = True
   Else
   bListDiode = False
End If

 


On Error GoTo EX
  ' Open PrmPath & "ReadBomValue\WaitCheck.txt" For Output As #7
  '   Print #7, Now
   '   Print #7,
   Open PrmPath & "ReadBomValue\Jumper.txt" For Output As #6
     Print #6, Now
      Print #6,
   If bListCatacitor = True Then
      Open PrmPath & "ReadBomValue\Catacitor.txt" For Output As #2
      Print #2, Now
      Print #2,
   End If
   If bListResistor = True Then
      Open PrmPath & "ReadBomValue\Resistor.txt" For Output As #4
      Print #4, Now
      Print #4,
   End If
   If bListDiode = True Then
      Open PrmPath & "ReadBomValue\Diode.txt" For Output As #8
      Print #8, Now
      Print #8,
   
   End If
   
   
   
   
   
      Open PrmPath & "ReadBomValue\Unknow.txt" For Output As #5
         Print #5, Now
         Print #5,
      Close #5
   Open Trim(txtBomPath.Text) For Input As #1
      Do Until EOF(1)
        Line Input #1, Mystr
         strReadText = Mystr
           Mystr = Trim(UCase(Mystr))
           
             If Mystr <> "" Then
                If Left(Mystr, 1) <> "-" Then
                  TmpStr = Split(Mystr, " ")
                  strDeviceName = TmpStr(UBound(TmpStr))
                  tmpSTR1 = Trim(tmpSTR1)
                  tmpSTR1 = Trim(Replace(Mystr, TmpStr(0), ""))
                  TmpStr = Split(tmpSTR1, " ")
                  DeviceType_ = Trim(TmpStr(0))
                  Select Case DeviceType_
                    
                     Case "CONN"
                     Case "SKT"
                     Case "HEAD"
                     Case "EMI"
                     Case "BOSS"
                     Case "2HIP"
                     Case "SKT"
                     Case "CHIP"
                         tmpSTR1 = Trim(Replace(tmpSTR1, TmpStr(0), ""))
                         tmpSTR1 = Trim(tmpSTR1)
                         TmpStr = Split(tmpSTR1, " ")
                         DeviceType_A = Trim(TmpStr(0))
                          If Len(DeviceType_A) > 1 Then
                             Select Case Left(DeviceType_A, 3)
                                 
                               Case "CAP"
                                 If bListCatacitor = True Then
                                   tmpSTR1 = Trim(Replace(tmpSTR1, strDeviceName, ""))
                                   tmpSTR1 = Trim(tmpSTR1)
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "CAP", ""))
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "T", ""))
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "F", ""))
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "NEO", ""))
                                   tmpSTR1 = Trim(tmpSTR1)
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "POS", ""))
                                   tmpSTR1 = Trim(tmpSTR1)
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "C ", ""))
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "C", ""))
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "EL", ""))
                                   tmpSTR1 = Trim(tmpSTR1)
                                   strCAP = Split(tmpSTR1, " ")

                                     If InStr(strCAP(0), "U") <> 0 And InStr(strCAP(0), "V") <> 0 Then
                                         CValue = Left(strCAP(0), InStr(strCAP(0), "U"))
                                               Print #2, strDeviceName; Tab(25); CValue; Tab(35); strCH; Tab(40); strCL; Tab(45); Tab(50); "f    PN""" & strDeviceName & """    ;"
                                           strReadText = "OK"
                                        Else
                                         If InStr(strCAP(0), "N") <> 0 And InStr(strCAP(0), "V") <> 0 Then
                                             CValue = Left(strCAP(0), InStr(strCAP(0), "N"))
                                               Print #2, strDeviceName; Tab(25); CValue; Tab(35); strCH; Tab(40); strCL; Tab(45); Tab(50); "f    PN""" & strDeviceName & """    ;"
                                               strReadText = "OK"
                                            Else
                                             If InStr(strCAP(0), "P") <> 0 And InStr(strCAP(0), "V") <> 0 Then
                                               CValue = Left(strCAP(0), InStr(strCAP(0), "P"))
                                               Print #2, strDeviceName; Tab(25); CValue; Tab(35); strCH; Tab(40); strCL; Tab(45); Tab(50); "f    PN""" & strDeviceName & """    ;"
                                                strReadText = "OK"
                                               Else
                                                 CValue = strCAP(0)
                                               
                                               Print #2, strDeviceName; Tab(25); CValue; Tab(35); strCH; Tab(40); strCL; Tab(45); Tab(50); "f    PN""" & strDeviceName & """    ;"
                                                strReadText = "OK"
                                             End If 'P,V
                                         End If 'N,V

                                     End If 'U,V
                                     
                                 End If 'bListCatacitor=true
                                 
                               Case "RES"
                                 If bListResistor = True Then
                                   tmpSTR1 = Trim(Replace(tmpSTR1, strDeviceName, ""))
                                   tmpSTR1 = Trim(tmpSTR1)
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "RES", ""))
                                   tmpSTR1 = Trim(tmpSTR1)
                                   strRES = Split(tmpSTR1, " ")
                                   RValue = strRES(0)
                                       RValue1 = Val(RValue)
                                    If Right(RValue, 1) <> "K" And Right(RValue, 1) <> "M" And InStr(RValue, "K") = 0 And InStr(RValue, "M") = 0 Then
                                       If CheckJumper.Value = 1 Then
                                           
                                           LowToJumper = Val(txtJumper.Text)
                                        If RValue1 < LowToJumper Then
                                           Print #6, strDeviceName; Tab(25); "CLOSED"; Tab(35); "PN""" & strDeviceName & """  ;      !BOM Value: " & RValue
                                           strReadText = "OK"
                                           Else
                                             Print #4, strDeviceName; Tab(25); RValue; Tab(35); strRH; Tab(40); strRL; Tab(45); Tab(50); "f    PN""" & strDeviceName & """    ;"
                                            strReadText = "OK"
                                        End If
                                      End If
                                      Else
                                         Print #4, strDeviceName; Tab(25); RValue; Tab(35); strRH; Tab(40); strRL; Tab(45); Tab(50); "f    PN""" & strDeviceName & """    ;"
                                          strReadText = "OK"
                                   End If
                                   
'                                   Print #4, strDeviceName; Tab(25); RValue; Tab(35); strRH; Tab(40); strRL; Tab(45); Tab(50); "f    PN""" & strDeviceName & """    ;"
'                                    strReadText = "OK"
                                 End If
                                  
                                  
                               Case "FUS"
                                If CheckIND.Value = 1 Then
                                   Print #6, strDeviceName; Tab(25); "CLOSED"; Tab(35); "PN""" & strDeviceName & """  ;"
                                   strReadText = "OK"
                                End If
                               Case "NTW"
                               Case "BEA"
                                   If CheckIND.Value = 1 Then
                                       Print #6, strDeviceName; Tab(25); "CLOSED"; Tab(35); "PN""" & strDeviceName & """  ;"
                                       strReadText = "OK"
                                    End If
                               
                               Case "CHO"
                                   If CheckIND.Value = 1 Then
                                       Print #6, strDeviceName; Tab(25); "CLOSED"; Tab(35); "PN""" & strDeviceName & """  ;"
                                       strReadText = "OK"
                                    End If
                               
                               Case "IND"
                                If CheckIND.Value = 1 Then
                                   Print #6, strDeviceName; Tab(25); "CLOSED"; Tab(35); "PN""" & strDeviceName & """  ;"
                                   strReadText = "OK"
                                End If
                                
                               Case "VAR"
                              Case "0.0"
                                 'If bListCatacitor = True Then
                                   tmpSTR1 = Trim(Replace(tmpSTR1, strDeviceName, ""))
                                   tmpSTR1 = Trim(tmpSTR1)
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "CAP", ""))
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "NEO", ""))
                                   tmpSTR1 = Trim(tmpSTR1)
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "POS", ""))
                                   tmpSTR1 = Trim(tmpSTR1)
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "C ", ""))
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "C", ""))
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "EL", ""))
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "F", ""))
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "T", ""))
                                   tmpSTR1 = Trim(tmpSTR1)
                                   strCAP = Split(tmpSTR1, " ")

                                     If InStr(strCAP(0), "U") <> 0 And InStr(strCAP(0), "V") <> 0 Then
                                         CValue = Left(strCAP(0), InStr(strCAP(0), "U"))
                                               Print #2, strDeviceName; Tab(25); CValue; Tab(35); strCH; Tab(40); strCL; Tab(45); Tab(50); "f    PN""" & strDeviceName & """    ;"
                                           strReadText = "OK"
                                        Else
                                         If InStr(strCAP(0), "N") <> 0 And InStr(strCAP(0), "V") <> 0 Then
                                             CValue = Left(strCAP(0), InStr(strCAP(0), "N"))
                                               Print #2, strDeviceName; Tab(25); CValue; Tab(35); strCH; Tab(40); strCL; Tab(45); Tab(50); "f    PN""" & strDeviceName & """    ;"
                                               strReadText = "OK"
                                            Else
                                             If InStr(strCAP(0), "P") <> 0 And InStr(strCAP(0), "V") <> 0 Then
                                               CValue = Left(strCAP(0), InStr(strCAP(0), "P"))
                                               Print #2, strDeviceName; Tab(25); CValue; Tab(35); strCH; Tab(40); strCL; Tab(45); Tab(50); "f    PN""" & strDeviceName & """    ;"
                                                strReadText = "OK"
                                               Else
                                                 CValue = strCAP(0)
                                               
                                               Print #2, strDeviceName; Tab(25); CValue; Tab(35); strCH; Tab(40); strCL; Tab(45); Tab(50); "f    PN""" & strDeviceName & """    ;"
                                                strReadText = "OK"
                                             End If 'P,V
                                         End If 'N,V

                                     End If 'U,V
                                     
                                ' End If 'bListCatacitor=true
                               
                               
                               
                               
                             End Select
                          Else
                             If Left(DeviceType_A, 1) = "C" Then
                                 If bListCatacitor = True Then
                                   tmpSTR1 = Trim(Replace(tmpSTR1, strDeviceName, ""))
                                   tmpSTR1 = Trim(tmpSTR1)
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "CAP", ""))
                                   tmpSTR1 = Trim(tmpSTR1)
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "POS", ""))
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "NEO", ""))
                                   tmpSTR1 = Trim(tmpSTR1)
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "C ", ""))
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "C", ""))
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "T", ""))
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "F", ""))
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "EL", ""))
                                   tmpSTR1 = Trim(tmpSTR1)
                                   strCAP = Split(tmpSTR1, " ")

                                     If InStr(strCAP(0), "U") <> 0 And InStr(strCAP(0), "V") <> 0 Then
                                         CValue = Left(strCAP(0), InStr(strCAP(0), "U"))
                                               Print #2, strDeviceName; Tab(25); CValue; Tab(35); strCH; Tab(40); strCL; Tab(45); Tab(50); "f    PN""" & strDeviceName & """    ;"
                                           strReadText = "OK"
                                        Else
                                         If InStr(strCAP(0), "N") <> 0 And InStr(strCAP(0), "V") <> 0 Then
                                             CValue = Left(strCAP(0), InStr(strCAP(0), "N"))
                                               Print #2, strDeviceName; Tab(25); CValue; Tab(35); strCH; Tab(40); strCL; Tab(45); Tab(50); "f    PN""" & strDeviceName & """    ;"
                                               strReadText = "OK"
                                            
                                            Else
                                             If InStr(strCAP(0), "P") <> 0 And InStr(strCAP(0), "V") <> 0 Then
                                               CValue = Left(strCAP(0), InStr(strCAP(0), "P"))
                                               Print #2, strDeviceName; Tab(25); CValue; Tab(35); strCH; Tab(40); strCL; Tab(45); Tab(50); "f    PN""" & strDeviceName & """    ;"
                                               strReadText = "OK"
                                               
                                               Else
                                                 CValue = strCAP(0)
                                               Print #2, strDeviceName; Tab(25); CValue; Tab(35); strCH; Tab(40); strCL; Tab(45); Tab(50); "f    PN""" & strDeviceName & """    ;"
                                                strReadText = "OK"
                                             End If 'P,V
                                         End If 'N,V

                                     End If 'U,V
                                 End If 'bListCatacitor=true
                    
                    
                    
                    
                           End If 'Left(DeviceType_A, 1) = "C"
                             
                             
                          End If 'Len(DeviceType_A) > 1
                     Case "IC"
                     Case "XFORM"
                     Case "THERM"
                     Case "IR"
                     Case "RESO"
                     Case "XTAL"
                     Case "DIODE"
                       If bListDiode = True Then
                           Print #8, strDeviceName; Tab(25); strDH; Tab(35); strDL; Tab(45); "PN""" & strDeviceName & """    ;"
                           strReadText = "OK"
                       End If
                     
                     Case "DIODES"
                       If bListDiode = True Then
                           Print #8, strDeviceName; Tab(25); strDH; Tab(35); strDL; Tab(45); "PN""" & strDeviceName & """    ;"
                           strReadText = "OK"
                       End If
                     
                     Case "LED"
                     Case "XTOR"
                     Case "FET"
                     Case "STANDOFF"

                     
                  End Select


                    
                End If 'Left(Mystr, 1) <> "-"
             End If 'Mystr <> ""
            If strReadText <> "OK" Then
                Open PrmPath & "ReadBomValue\Unknow.txt" For Append As #3
                   Print #3, strReadText
                Close #3
                
            End If
            strReadText = ""
            DoEvents
            tmpSTR1 = ""
            Mystr = ""
            strDeviceName = ""
            CValue = ""
            RValue = ""
             RValue1 = ""
      Loop
   Close #1
   If bListCatacitor = True Then
     Close #2
   End If
  
 ' Resistor
  
   If bListResistor = True Then
     Close #4
   End If
     Close #6
   '  Close #7
   If bListDiode = True Then
      Close #8
   End If
  
   MsgBox "OK" & Chr(13) & Chr(10) & "File save path:" & PrmPath & "ReadBomValue\", vbInformation
  
Exit Sub
EX:
MsgBox Err.Description, vbCritical

End Sub

Private Sub mcGo_Click()
On Error GoTo EX
  
   If Trim(txtBomPath.Text) = "" Then txtBomPath.Text = " Please open bom file!(DblClick me open file!)"
    If Dir(txtBomPath.Text) = "" Then
        txtBomPath.Text = " Please open bom file!(DblClick me open file!)"
        MsgBox "File not find!", vbCritical
        txtBomPath.SetFocus
      Exit Sub
    End If
    If FileLen(txtBomPath.Text) = 0 Then
        txtBomPath.Text = " Please open bom file!(DblClick me open file!)"
        MsgBox "The file text is null ,please check!", vbCritical
        txtBomPath.SetFocus
        Exit Sub
    End If
    mcGo.Enabled = False
    'start
       If CheckJumper.Value = 1 Then
          If Trim(txtJumper.Text) = "" Then
             txtJumper.Text = 11
          End If
          a = Val(txtJumper.Text)
          txtJumper.Text = a
       End If
       
      Call Read_BomFile
    'end
    mcGo.Enabled = True
    mcGo.SetFocus
    Exit Sub
EX:
mcGo.Enabled = True
 mcGo.SetFocus
 MsgBox Err.Description, vbCritical
End Sub

Private Sub txtBomPath_DblClick()
On Error GoTo errh
With Me.CommonDialog1
    .CancelError = True
    '.Filter = "*.txt|*.txt|*.log|*.log|*.*|*.*"
    
    .Filter = "bom file *.txt|*.txt|*.*|*.*"
    .ShowOpen

End With
    txtBomPath.Text = Me.CommonDialog1.FileName
    If Dir(txtBomPath.Text) = "" Then
        txtBomPath.Text = " Please open bom file!(DblClick me open file!)"
        MsgBox "File not find!", vbCritical
        txtBomPath.SetFocus
      Exit Sub
    End If
    If FileLen(txtBomPath.Text) = 0 Then
        txtBomPath.Text = " Please open bom file!(DblClick me open file!)"
        MsgBox "The file text is null ,please check!", vbCritical
        txtBomPath.SetFocus
        Exit Sub
    End If
 
Exit Sub

errh:
MsgBox Err.Description, vbCritical
    txtBomPath.Text = "Please open bom file!(DblClick me open file!)"
    txtBomPath.SetFocus


End Sub
