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
Dim totalfivefoldindex() As String
Dim allattribute(7) As Double



Static Function distance(ByVal xindex As Double, ByVal yindex As Double, ByRef attrarray() As Double)
Dim dimnumber As Double
Dim xydistance As Double
Dim totalsum As Double
Dim tempattrarray() As Double
Dim xarray() As Double
Dim yarray() As Double
xydistance = 0
totalsum = 0
tempattrarray() = attrarray()
dimnumber = UBound(tempattrarray) + 1
ReDim xarray(UBound(tempattrarray))
ReDim yarray(UBound(tempattrarray))

For i = 0 To UBound(tempattrarray)
xarray(i) = CDbl(data2darray(tempattrarray(i), xindex))
yarray(i) = CDbl(data2darray(tempattrarray(i), yindex))
Next i

'For i = 0 To UBound(tempattrarray)
'List1.AddItem xarray(i)
'Next i
'List1.AddItem ""
'For i = 0 To UBound(tempattrarray)
'List1.AddItem yarray(i)
'Next i


For i = 0 To UBound(tempattrarray)
totalsum = totalsum + ((xarray(i) - yarray(i)) ^ 2)
Next i

xydistance = (totalsum ^ (1 / 2))

distance = xydistance
End Function

Static Function gmax(ByRef distarray() As Double, ByRef distindexarray() As Double)
Dim tempdistarray() As Double
Dim tempdistindexarray() As Double
Dim distmax As Double
Dim indexmax As Double
distmax = -100
indexmax = -100
tempdistarray() = distarray()
tempdistindexarray() = distindexarray()

For i = 0 To UBound(tempdistarray)
If tempmax < tempdistarray(i) Then
distmax = tempdistarray(i)
indexmax = tempdistindexarray(i)

End If
Next i

gmax = CStr(indexmax) + "," + CStr(distmax)
End Function

Static Function sortrnd(ByRef tempdataindex() As Double, ByRef temprndarray() As Double)

Dim tmp As Double
Dim tmpindex As Double

For i = 0 To UBound(tempdataindex)
    For j = i To UBound(tempdataindex)
        If temprndarray(i) > temprndarray(j) Then
            tmp = temprndarray(i)
            temprndarray(i) = temprndarray(j)
            temprndarray(j) = tmp
            
            tmpindex = tempdataindex(i)
            tempdataindex(i) = tempdataindex(j)
            tempdataindex(j) = tmpindex
        End If
    Next j
Next i


sortrnd = "sortrnd"
End Function

Static Function predclass(ByVal testoneindex As Double, ByRef ttemptrainarray() As Double)
Dim tttemptrainarray() As Double
Dim final_class As String
Dim distarray() As Double
Dim distarrayindex() As Double
Dim topndist() As Double
Dim topndistindex() As Double
Dim tempindexattr() As String
Dim attributesubset() As Double
Dim tempgmax As String

For i = 0 To 7
allattribute(i) = i + 1
Next i
attributesubset() = allattribute()
tttemptrainarray() = ttemptrainarray()

ReDim distarray(UBound(tttemptrainarray))
ReDim distarrayindex(UBound(tttemptrainarray))
ReDim topndist(nei_number - 1) '存前n近的值
ReDim topndistindex(nei_number - 1) '存前n近的index

For i = 0 To UBound(tttemptrainarray)
distarray(i) = distance(testoneindex, tttemptrainarray(i), attributesubset)
distarrayindex(i) = tttemptrainarray(i)
Next i



For i = 0 To (nei_number - 1)
tempgmax = gmax(distarray, distarrayindex)
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

For i = 0 To UBound(topndistindex)
    List1.AddItem topndistindex(i)
Next i
List1.AddItem ""
'從這裡開始!!!!
'用topndistindex()找到對應的class
'比較class然後投票決定predclass是什麼

predclass = "class"
End Function

Static Function correctrate(ByRef testarray() As Double, ByRef trainarray() As Double)
Dim temptestarray() As Double
Dim temptrainarray() As Double
Dim correctnum As Double
Dim correctrateresult As Double
Dim predictclass As String
temptestarray() = testarray()
temptrainarray() = trainarray()
correctnum = 0

predictclass = predclass(temptestarray(i), temptrainarray)
'For i = 0 To UBound(temptestarray)
'predictclass = predclass(temptestarray(i), temptrainarray)
'If data2darray(9, testarray(i)) = predictclass Then
'correctnum = correctnum + 1
'End If
'Next i

correctrateresult = (correctnum) / (UBound(temptestarray) + 1)

correctrate = correctrateresult
End Function

Static Function fivefoldindex()
Dim foldindex(4) As String
Dim eachfoldindex As String
Dim rndarray(1483) As Double
Dim dataindex(1483) As Double
Dim tempsortrnd As String
Dim counter As Double
counter = 0

For i = 0 To 1483
    dataindex(i) = i
Next i

For i = 0 To 1483
    'Randomize (Timer)
    rndarray(i) = Rnd()
    'List1.AddItem rndarray(i)
    If i = 0 Then
        GoTo nozero
    End If
    For j = 0 To i - 1 '讓第i個產生的號碼,跟之前已經產生過的號碼比較, 如果重複，便重新選取,因此要讓i=i-1(倒退一位)
        If rndarray(i) = rndarray(j) Then
            i = i - 1
            Exit For
        End If
    Next j
nozero:
Next i


tempsortrnd = sortrnd(dataindex, rndarray)

'For i = 0 To 1483
'List1.AddItem dataindex(i)
''檢驗有沒有重複
''If i = 0 Then
''    GoTo ggg
''End If
''    For j = 0 To i - 1 '讓第i個產生的號碼,跟之前已經產生過的號碼比較, 如果重複，便重新選取,因此要讓i=i-1(倒退一位)
''        If dataindex(i) = dataindex(j) Then
''            List1.AddItem "----------------------------------"
''        End If
''    Next j
''ggg:
'Next i

For i = 0 To 4
    eachfoldindex = ""
    For j = 0 To 296
    If i = 4 And j = 296 Then '第5個fold是296個
        Exit For
    End If
    eachfoldindex = eachfoldindex + CStr(dataindex(counter)) + " "
    counter = counter + 1
    Next j
    foldindex(i) = eachfoldindex
Next i
fivefoldindex = foldindex() '回傳分好5組的index的字串
End Function

Private Sub cross_validation_Click()
List1.Clear
'Dim tempfivefoldindex() As String '存5組字串
Dim trainindexstring As String
Dim trainnumber As Double
Dim testrnd As Double
Dim subsetcounter As Double
Dim eachtrainsubset(3) As Double
Dim correctratearray(4) As Double
Dim eachtrainDouble() As Double
Dim eachtestDouble() As Double

'測試distance
'Dim distanceDouble As Double
'Dim xi As Double
'Dim yi As Double
'Dim dimnum(7) As Double
'xi = 0
'yi = 1
'For i = 0 To UBound(dimnum)
'dimnum(i) = (i + 1)
'Next i
'distanceDouble = distance(xi, yi, dimnum)
'List1.AddItem distanceDouble
'目前進度完成測試distance

'測試rnd
'Randomize (Timer)
'testrnd = Rnd()
'List1.AddItem testrnd

'測試sortrnd
'Dim testindex(2) As Double
'Dim testvalue(2) As Double
'Dim sort As String
'For i = 0 To 2
'    testindex(i) = i
'Next i
'testvalue(0) = 0.4367
'testvalue(1) = 0.7425
'testvalue(2) = 0.2549
'sort = sortrnd(testindex, testvalue)
'For i = 0 To 2
'    List1.AddItem testvalue(i)
'    List1.AddItem testindex(i)
'Next i

'回傳分好五份的資料index
totalfivefoldindex() = fivefoldindex()

'選一個當測試資料,其他四個合成訓練資料
For i = 0 To 4
Dim eachtrain() As String
Dim eachtrainstring As String
Dim eachtest() As String
Dim eachteststring As String
subsetcounter = 0
eachtrainstring = ""
eachteststring = ""
'eachfoldarray() = Split(totalfivefoldindex(i), " ")

'區分出每次的訓練以及測試資料
For j = 0 To 4
    If j <> i Then
    eachtrainsubset(subsetcounter) = j
    subsetcounter = subsetcounter + 1
    End If
Next j

For j = 0 To 3
    eachtrainstring = eachtrainstring + totalfivefoldindex(eachtrainsubset(j))
Next j

eachteststring = Trim(totalfivefoldindex(i))
eachtrainstring = Trim(eachtrainstring)

eachtest() = Split(eachteststring, " ")
eachtrain() = Split(eachtrainstring, " ")

ReDim eachtestDouble(UBound(eachtest) + 1)
ReDim eachtrainDouble(UBound(eachtrain) + 1)

For j = 0 To UBound(eachtest)
eachtestDouble(j) = CDbl(eachtest(j))
Next j

For j = 0 To UBound(eachtrain)
eachtrainDouble(j) = CDbl(eachtrain(j))
Next j
'測試每次的test和train的筆數
'List1.AddItem totalfivefoldindex(i)
'List1.AddItem UBound(eachtestDouble) + 1
'List1.AddItem eachtrainDouble(0)
'List1.AddItem UBound(eachtrainDouble) + 1
'List1.AddItem ""
correctratearray(i) = correctrate(eachtestDouble, eachtrainDouble)
Next i


'GoTo endd
'endd:
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
