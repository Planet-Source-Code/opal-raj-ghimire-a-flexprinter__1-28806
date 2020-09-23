VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form1"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10935
   FillColor       =   &H00E0E0E0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   10935
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command15 
      Caption         =   "Remove"
      Height          =   285
      Left            =   9900
      TabIndex        =   55
      ToolTipText     =   "   Removes some rows and cols      "
      Top             =   90
      Width           =   870
   End
   Begin VB.CommandButton Command3 
      Caption         =   "About"
      Height          =   420
      Left            =   9225
      TabIndex        =   54
      Top             =   7290
      Width           =   1500
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Border Print"
      Height          =   240
      Left            =   1950
      TabIndex        =   46
      Top             =   2610
      Value           =   1  'Checked
      Width           =   1500
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "WordWrap"
      Height          =   195
      Left            =   4140
      TabIndex        =   45
      Top             =   2610
      Width           =   1545
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Alignments"
      Height          =   330
      Left            =   180
      TabIndex        =   44
      ToolTipText     =   "   other alignmets   "
      Top             =   2520
      Width           =   1545
   End
   Begin VB.PictureBox Picture3 
      Height          =   4650
      Left            =   135
      ScaleHeight     =   4590
      ScaleWidth      =   8775
      TabIndex        =   42
      Top             =   2925
      Width           =   8835
      Begin VB.PictureBox PIC 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   9045
         Left            =   135
         ScaleHeight     =   601
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   803
         TabIndex        =   43
         Top             =   135
         Width           =   12075
      End
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Inch"
      Height          =   195
      Left            =   9090
      TabIndex        =   40
      Top             =   6075
      Width           =   870
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cm"
      Height          =   195
      Left            =   10035
      TabIndex        =   39
      Top             =   6075
      Value           =   -1  'True
      Width           =   825
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   9135
      TabIndex        =   38
      Text            =   "8000"
      Top             =   5220
      Width           =   1230
   End
   Begin VB.CheckBox Hor 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Horizontal Ruler"
      Height          =   285
      Left            =   9135
      TabIndex        =   36
      Top             =   5760
      Width           =   2175
   End
   Begin VB.CheckBox Ver 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Verticle Ruler"
      Height          =   285
      Left            =   9135
      TabIndex        =   35
      Top             =   5535
      Width           =   2130
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Rows"
      Height          =   825
      Left            =   9000
      TabIndex        =   30
      Top             =   4050
      Width           =   1770
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1035
         TabIndex        =   34
         ToolTipText     =   " No. of final  row "
         Top             =   450
         Width           =   645
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1035
         TabIndex        =   33
         ToolTipText     =   " No. of initial  row "
         Top             =   135
         Width           =   645
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "To"
         Height          =   240
         Left            =   45
         TabIndex        =   32
         Top             =   540
         Width           =   915
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "From"
         Height          =   195
         Left            =   45
         TabIndex        =   31
         Top             =   225
         Width           =   870
      End
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Grid Print"
      Height          =   240
      Left            =   6615
      TabIndex        =   29
      Top             =   2610
      Value           =   1  'Checked
      Width           =   1275
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   9
      Left            =   10035
      TabIndex        =   27
      Top             =   3780
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   8
      Left            =   10035
      TabIndex        =   26
      ToolTipText     =   " Value to round the cornor of rectangle"
      Top             =   3465
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   7
      Left            =   10035
      TabIndex        =   25
      ToolTipText     =   " Value to round the cornor of rectangle"
      Top             =   3195
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   6
      Left            =   10035
      TabIndex        =   24
      ToolTipText     =   "Horizontle space between Rows"
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   10035
      TabIndex        =   23
      ToolTipText     =   "Verticle Space between columns"
      Top             =   2610
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   10035
      TabIndex        =   22
      Top             =   2295
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   10035
      TabIndex        =   21
      Top             =   2025
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   10200
      TabIndex        =   20
      Top             =   1260
      Width           =   600
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   10200
      TabIndex        =   19
      Top             =   945
      Width           =   600
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Border"
      Height          =   1545
      Left            =   9300
      TabIndex        =   14
      Top             =   450
      Width           =   1590
      Begin VB.PictureBox Picture2 
         Height          =   285
         Left            =   900
         ScaleHeight     =   225
         ScaleWidth      =   540
         TabIndex        =   28
         ToolTipText     =   "      Drag and drop color here      "
         Top             =   1125
         Width           =   600
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   900
         TabIndex        =   18
         Top             =   180
         Width           =   600
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Color"
         Height          =   240
         Left            =   90
         TabIndex        =   41
         Top             =   1215
         Width           =   690
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Width"
         Height          =   285
         Left            =   45
         TabIndex        =   17
         Top             =   900
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Style"
         Height          =   195
         Left            =   90
         TabIndex        =   16
         Top             =   600
         Width           =   690
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Distance"
         Height          =   195
         Left            =   0
         TabIndex        =   15
         Top             =   270
         Width           =   780
      End
   End
   Begin VB.PictureBox Picture1 
      DragIcon        =   "PrintGrid.frx":0000
      Height          =   1680
      Left            =   7965
      Picture         =   "PrintGrid.frx":0442
      ScaleHeight     =   1620
      ScaleWidth      =   1215
      TabIndex        =   13
      ToolTipText     =   "  Click to change Text color Drag to change Boarder color   "
      Top             =   405
      Width           =   1275
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Print"
      Height          =   420
      Left            =   9225
      TabIndex        =   12
      Top             =   6840
      Width           =   1500
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Preview/Refresh"
      Height          =   465
      Left            =   9225
      TabIndex        =   11
      Top             =   6345
      Width           =   1500
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   8235
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   45
      Width           =   1455
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Right>>"
      Height          =   285
      Left            =   7470
      TabIndex        =   9
      Top             =   45
      Width           =   735
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Center"
      Height          =   285
      Left            =   6660
      TabIndex        =   8
      Top             =   45
      Width           =   780
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0C0&
      Caption         =   "<<Left"
      Height          =   285
      Left            =   5895
      TabIndex        =   7
      Top             =   45
      Width           =   780
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0C0&
      Caption         =   "StrikeThorugh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   -1  'True
      EndProperty
      Height          =   285
      Left            =   4545
      TabIndex        =   6
      Top             =   45
      Width           =   1365
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Underline"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3645
      TabIndex        =   5
      Top             =   45
      Width           =   915
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Italics"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2790
      TabIndex        =   4
      Top             =   45
      Width           =   870
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Bold"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2070
      TabIndex        =   3
      Top             =   45
      Width           =   780
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "FontSize -"
      Height          =   285
      Left            =   1125
      TabIndex        =   2
      Top             =   45
      Width           =   960
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "FontSize +"
      Height          =   285
      Left            =   180
      TabIndex        =   1
      Top             =   45
      Width           =   960
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flex 
      Height          =   2040
      Left            =   135
      TabIndex        =   0
      ToolTipText     =   $"PrintGrid.frx":2AE4
      Top             =   405
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   3598
      _Version        =   393216
      Cols            =   8
      RowHeightMin    =   250
      GridColor       =   0
      AllowBigSelection=   0   'False
      FocusRect       =   2
      FillStyle       =   1
      AllowUserResizing=   3
      GridLineWidthFixed=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
      _Band(0).GridLineWidthBand=   1
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "H Space"
      Height          =   240
      Left            =   8910
      TabIndex        =   53
      Top             =   2970
      Width           =   1050
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "V Space"
      Height          =   240
      Left            =   8865
      TabIndex        =   52
      Top             =   2655
      Width           =   1095
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Round X"
      Height          =   240
      Left            =   8865
      TabIndex        =   51
      Top             =   3285
      Width           =   1095
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Round Y"
      Height          =   240
      Left            =   8820
      TabIndex        =   50
      Top             =   3555
      Width           =   1140
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Left"
      Height          =   240
      Left            =   8865
      TabIndex        =   49
      Top             =   2115
      Width           =   1095
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Top"
      Height          =   240
      Left            =   8865
      TabIndex        =   48
      Top             =   2385
      Width           =   1095
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Style"
      Height          =   240
      Left            =   8775
      TabIndex        =   47
      Top             =   3825
      Width           =   1185
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ruler Length"
      Height          =   240
      Left            =   9135
      TabIndex        =   37
      Top             =   4950
      Width           =   1500
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'A sample to show use of FlexPrinter class


Dim CCC As Long

Dim M As FlexPrinter
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


Private Sub Check1_Click()
Command12_Click
End Sub

Private Sub Check2_Click()
Command12_Click
End Sub

Private Sub Check3_Click()
Flex.WordWrap = Check3.Value
End Sub

Private Sub Combo1_Click()
Flex.CellFontName = Combo1.Text
Command12_Click
End Sub



Private Sub Command1_Click()
Static X As Integer
X = X + 1
Flex.CellAlignment = X
If X > 8 Then X = 0
End Sub

Private Sub Command10_Click()
Flex.CellAlignment = 4

End Sub

Private Sub Command11_Click()
Flex.CellAlignment = 7
End Sub






Private Sub Command12_Click()
PIC.Cls

Dim MEs As String
If Option1.Value = True Then MEs = "CM" Else MEs = "INCH"
With M
.RowsFrom = Val(Text2)
.RowsTo = Val(Text3)

'to center
'.PosTop = (PIC.ScaleHeight - .GetHeight(PIC)) / 2 '
'.PosLeft = (PIC.ScaleWidth - .GetWidth(PIC)) / 2 '

.PosTop = Val(Text1(4).Text)
.PosLeft = Val(Text1(3).Text)

.HSpace = Val(Text1(6).Text)
.VSpace = Val(Text1(5).Text)

.RoundCorX = Val(Text1(7).Text)
.RoundCorY = Val(Text1(8).Text)

.GridPenStyle = Val(Text1(9).Text)
.GridPrint = Check2.Value
.DrawBorder = Check1.Value
.BorderColor = Picture2.BackColor
.BorderStyle = Val(Text1(1).Text)
.BorderWidth = Val(Text1(2).Text)
.BorderDistance = Val(Text1(0).Text)
.RowsFrom = Val(Text2.Text)
.RowsTo = Val(Text3.Text)

.PrintOut PIC

If Ver.Value = vbChecked Then
.DrawRulerV PIC, Val(Text1(3)), Val(Text1(4)), Val(Text4.Text), MEs
End If
If Hor.Value = vbChecked Then
.DrawRulerH PIC, Val(Text1(3)), Val(Text1(4)), Val(Text4.Text), MEs

End If

End With

'these two lines to check height and width of picture in picturebox
'PIC.Line (M.PosLeft, M.PosTop - 5)-(M.GetWidth(PIC) + M.PosLeft, M.PosTop - 5)
'PIC.Line (M.PosLeft - 5, M.PosTop)-(M.PosLeft - 5, M.GetHeight(PIC) + M.PosTop)

End Sub



Private Sub Command14_Click()
Printer.PaperSize = 9
Printer.Orientation = 1
Printer.ScaleMode = 3
M.RowsFrom = Val(Text2)
M.RowsTo = Val(Text3)

M.PosTop = (Printer.ScaleHeight - M.GetHeight(Printer)) / 2
M.PosLeft = (Printer.ScaleWidth - M.GetWidth(Printer)) / 2
M.PrintOut Printer

Printer.EndDoc
End Sub


Private Sub Command15_Click()
Flex.RowHeight(5) = 0
Flex.RowHeight(4) = 0
Flex.RowHeight(2) = 0
Flex.ColWidth(3) = 0


End Sub







Private Sub Command2_Click()

Flex.CellFontSize = Flex.CellFontSize + 2
End Sub





Private Sub Command3_Click()
MsgBox "An Ordinary MSHFlex Grid Printer" + vbCrLf + "Opal Raj Ghimire" + vbCrLf + "Kathmandu, Nepal" + vbCrLf + vbCrLf + "buna48@hotmail.com" + vbCrLf + "http://geocities.com/opalraj/vb", , "Flex Printer"
'M.ColSetupCode


End Sub

Private Sub Command4_Click()
Flex.CellFontSize = Flex.CellFontSize - 2
If Flex.CellFontSize < 8 Then Flex.CellFontSize = 8
End Sub

Private Sub Command5_Click()
Flex.CellFontBold = Not Flex.CellFontBold
End Sub

Private Sub Command6_Click()
Flex.CellFontItalic = Not Flex.CellFontItalic
End Sub

Private Sub Command7_Click()
Flex.CellFontUnderline = Not Flex.CellFontUnderline
End Sub



Private Sub Command8_Click()
Flex.CellFontStrikeThrough = Not Flex.CellFontStrikeThrough
End Sub

Private Sub Command9_Click()
Flex.CellAlignment = 1

End Sub

Private Sub Form_Load()
Dim I As Integer
Dim CC As Control
For Each CC In Form1
CC.FontName = "MS Sans Serif"
CC.FontSize = 8
Next

With Flex

.ColWidth(0) = 300
.ColWidth(1) = 1920
.ColWidth(2) = 900
.ColWidth(3) = 1095
.ColWidth(4) = 735
.ColWidth(5) = 600
.ColWidth(6) = 540
.ColWidth(7) = 915
.TextMatrix(0, 0) = "Sr"
.TextMatrix(0, 1) = "Clients"
.TextMatrix(0, 2) = "Time"
.TextMatrix(0, 3) = "Date"
.TextMatrix(0, 4) = "Code"
.TextMatrix(0, 5) = "Cost"
.TextMatrix(0, 6) = "Unit"
.TextMatrix(0, 7) = "Amount"
.AddItem "12" + Chr(9) + "Inter Continental Potato Traders" + Chr(9) + "10:40 AM" + Chr(9) + "2001/04/05" + Chr(9) + "M88" + Chr(9) + "10" + Chr(9) + "5" + Chr(9) + "50.00", 1
.AddItem "11" + Chr(9) + "James Bond Private Eye Service" + Chr(9) + "11:00 AM" + Chr(9) + "2001/04/05" + Chr(9) + "IMX32" + Chr(9) + "100" + Chr(9) + "25" + Chr(9) + "2500.00", 1
.AddItem "10" + Chr(9) + "Jhon & Jony Jewllery Suppliers" + Chr(9) + "12:10 PM" + Chr(9) + "2001/04/05" + Chr(9) + "MX8" + Chr(9) + "8" + Chr(9) + "5" + Chr(9) + "40.00", 1
.AddItem "9" + Chr(9) + "Soft and Smooth Business Software" + Chr(9) + "12:40 PM" + Chr(9) + "2001/04/05" + Chr(9) + "A4G" + Chr(9) + "111" + Chr(9) + "5" + Chr(9) + "555.00", 1
.AddItem "8" + Chr(9) + "Petro, The Odd Job Man" + Chr(9) + "01:40 PM" + Chr(9) + "2001/04/05" + Chr(9) + "GM9" + Chr(9) + "10" + Chr(9) + "8" + Chr(9) + "80.00", 1
.AddItem "7" + Chr(9) + "Carpet World" + Chr(9) + "02:00 PM" + Chr(9) + "2001/04/05" + Chr(9) + "A6L" + Chr(9) + "66.33" + Chr(9) + "5" + Chr(9) + "331.65", 1
.AddItem "6" + Chr(9) + "Anita Publishers" + Chr(9) + "03:20 PM" + Chr(9) + "2001/04/05" + Chr(9) + "XL2" + Chr(9) + "86.49" + Chr(9) + "18" + Chr(9) + "1556.82", 1
.AddItem "5" + Chr(9) + "Doom Nurshing Home Pvt Ltd" + Chr(9) + "03:25 PM" + Chr(9) + "2001/04/05" + Chr(9) + "CBZ33" + Chr(9) + "59.99" + Chr(9) + "7" + Chr(9) + "419.93", 1
.AddItem "4" + Chr(9) + "Solid Gold Furnitures" + Chr(9) + "03:20 PM" + Chr(9) + "2001/04/05" + Chr(9) + "XL2" + Chr(9) + "86.49" + Chr(9) + "18" + Chr(9) + "1556.82", 1
.AddItem "3" + Chr(9) + "K Computer Parts Dealer" + Chr(9) + "03:25 PM" + Chr(9) + "2001/04/05" + Chr(9) + "CBZ33" + Chr(9) + "59.99" + Chr(9) + "7" + Chr(9) + "419.93", 1
.AddItem "2" + Chr(9) + "Solid Gold Furnitures" + Chr(9) + "03:20 PM" + Chr(9) + "2001/04/05" + Chr(9) + "XL2" + Chr(9) + "86.49" + Chr(9) + "18" + Chr(9) + "1556.82", 1
.AddItem "1" + Chr(9) + "J K International Man Power Suppliers Pvt Ltd" + Chr(9) + "03:25 PM" + Chr(9) + "2001/04/05" + Chr(9) + "CBZ33" + Chr(9) + "59.99" + Chr(9) + "7" + Chr(9) + "419.93", 1

End With
Check3.Value = IIf(Flex.WordWrap, vbChecked, vbUnchecked)
 For I = 0 To Printer.FontCount - 1  ' Determine number of fonts.
     Combo1.AddItem Printer.Fonts(I)  ' Put each font into list box.
    Next I
'set up
Text1(0) = "5"
Text1(1) = "0"
Text1(2) = "1"
Text1(3) = "50"
Text1(4) = "50"
Text1(5) = "0"
Text1(6) = "0"
Text1(7) = "10"
Text1(8) = "10"
Text1(9) = "0"
'Text1(10) = "0"
'Text1(11) = "1"

Set M = New FlexPrinter
Set M.FlexName = Flex

'Check1.Value = vbChecked
'Check2.Value = vbChecked
Picture2.BackColor = vbBlack
 Text2 = 0: M.RowsFrom = Val(Text2)
 Text3 = 12: M.RowsTo = Val(Text3)

End Sub


Private Sub Flex_Click()
Combo1 = Flex.CellFontName

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set M = Nothing
End Sub



Private Sub Hor_Click()
Command12_Click
End Sub

Private Sub Option1_Click()
Command12_Click
End Sub

Private Sub Option2_Click()
Command12_Click
End Sub



Private Sub PIC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'MsgBox Str(X) + " " + Str(Y)
End Sub

Private Sub PIC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
        ReleaseCapture
        Call SendMessage(PIC.hwnd, &H112, &HF012, 0)
End If

End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
CCC = Picture1.Point(X, Y)
Flex.CellForeColor = CCC
If Button = 1 Then Picture1.Drag
End Sub

Private Sub Picture2_DragDrop(Source As Control, X As Single, Y As Single)
If Source.Name = "Picture1" Then
Picture2.BackColor = CCC
Command12_Click
End If
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then Command12_Click
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
If IsNumeric(Text1(Index).Text) = False Then
MsgBox "Numeric Expected"
Cancel = True
End If

End Sub


Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command12_Click

End Sub

Private Sub Text2_Validate(Cancel As Boolean)
If IsNumeric(Text2.Text) = False Then
MsgBox "Numeric Expected"
Cancel = True
End If

End Sub


Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command12_Click
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
If IsNumeric(Text3.Text) = False Then
MsgBox "Numeric Expected"
Cancel = True
End If

End Sub


Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command12_Click
End Sub

Private Sub Text4_Validate(Cancel As Boolean)
If IsNumeric(Text4.Text) = False Then
MsgBox "Numeric Expected"
Cancel = True
End If

End Sub

Private Sub Ver_Click()
Command12_Click
End Sub
