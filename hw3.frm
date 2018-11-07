VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8010
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   ScaleHeight     =   8010
   ScaleWidth      =   10680
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton read 
      Caption         =   "read"
      Height          =   495
      Left            =   7800
      TabIndex        =   7
      Top             =   1800
      Width           =   2175
   End
   Begin VB.CommandButton cross_validation 
      Caption         =   "5-fold cross validation"
      Height          =   375
      Left            =   7200
      TabIndex        =   6
      Top             =   3000
      Width           =   3375
   End
   Begin VB.CommandButton backward 
      Caption         =   "backward"
      Height          =   615
      Left            =   8880
      TabIndex        =   5
      Top             =   4200
      Width           =   1695
   End
   Begin VB.CommandButton forward 
      Caption         =   "forward"
      Height          =   615
      Left            =   7200
      TabIndex        =   4
      Top             =   4200
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   6900
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   6735
   End
   Begin VB.TextBox nerghbors_num 
      Height          =   375
      Left            =   7800
      TabIndex        =   2
      Text            =   "3"
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox datanumber 
      Height          =   375
      Left            =   7800
      TabIndex        =   1
      Text            =   "1484"
      Top             =   240
      Width           =   2175
   End
   Begin VB.TextBox datatxt 
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Text            =   "yeast.txt"
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim file As String
Dim datanum As Integer
Dim nei_number As Integer
Dim data2darray(9, 1483) As String

Static Function distance(ByRef xarray() As Double, ByRef yarray() As Double)

Dim attr_number As Double
Dim final_class As String


attr_number = UBound(xarray) + 1


distance = CStr(attrmax) + "," + CStr(tempmax)
End Function

Static Function predclass(ByVal testoneindex As Double, ByRef ttemptrainarray() As Double)
Dim tttemptrainarray() As Double
Dim final_class As String
Dim distarray() As Double
Dim distarrayindex() As Double
Dim topndist() As Double
Dim topndistindex() As Double
Dim tempindexattr() As String

tttemptrainarray() = ttemptrainarray()
ReDim distarray(UBound(tttemptrainarray))
ReDim distarrayindex(UBound(tttemptrainarray))
Dim tempgmax As String
ReDim topndist(nei_number - 1) '存前n近的值
ReDim topndistindex(nei_number - 1) '存前n近的index

For i = 0 To UBound(tttemptrainarray)
'distarray(i)=distance(testoneindex,tttemptrainarray(i))
'distarrayindex(i)=tttemptrainarray(i)
Next i

For i = 0 To (nei_number - 1)
'tempgmax=gmax(distarray,distarrayindex)
tempindexattr = Split(tempgmax, ",")
topndistindex(i) = tempindexattr(0)

'把top1設定成-1
For j = 0 To UBound(tttemptrainarray)
If (tttemptrainarray(j)) = tempindexattr(0) Then
    tttemptrainarray(j) = -1
    distarray(j) = -1
End If
Next j

Next i

'從這裡開始!!!!
'用topndistindex()找到對應的class

distance = CStr(attrmax) + "," + CStr(tempmax)
End Function

Static Function correctrate(ByRef testarray() As Double, ByRef trainarray() As Double)
Dim temptestarray() As Double
Dim temptrainarray() As Double
Dim correctnum As Double
Dim predictclass As String
temptestarray() = testarray()
temptrainarray() = trainarray()
correctnum = 0

For i = 0 To uboubd(testarray)
'predictclass=predclass(testarray(i),temptrainarray)
If data2darray(9, testarray(i)) = predictclass Then
correctnum = correctnum + 1
End If
Next i


correctrate = correctnum
End Function

Static Function fivefoldindex()
Dim attr_number(4) As String
'For i = 0 To 4
'For j = 0 To 300
'attr_number(i, j) = i
'Next j
'Next i


fivefoldindex = attr_number() '回傳分好5組的index的字串
End Function

Private Sub cross_validation_Click()
List1.Clear
Dim tempfivefoldindex() As String '存5組字串
Dim trainindexstring As String
Dim trainnumber As Double

'call fivefoldindex()
tempfivefoldindex = fivefoldindex()
'選一個當測試資料,其他四個合成訓練資料
'測試資料string變成array
'訓練資料string變成array
'call correctrate()


GoTo endd
endd:
End Sub

Private Sub datanumber_Change()
datanum = CInt(datanumber.Text)
End Sub

Private Sub datatxt_Change()
file = datatxt.Text
End Sub



Private Sub nerghbors_num_Change()
nei_number = CInt(nerghbors_num.Text)
End Sub

Private Sub read_Click()
List1.Clear
Dim datacounter As Integer
Dim temp() As String
List1.AddItem nei_number
List1.AddItem ""

datacounter = 0

Open App.Path & "\" + file For Input As #1
Do While Not EOF(1) And datacounter < datanum

Line Input #1, tmpline

tmpline = Replace(tmpline, "  ", " ")
tmpline = Replace(tmpline, "  ", " ")
List1.AddItem tmpline
temp = Split(tmpline, " ")
'List1.AddItem temp

For i = 0 To 9
    data2darray(i, datacounter) = Trim(temp(i))
Next i
datacounter = datacounter + 1
Loop
Close #1
End Sub
