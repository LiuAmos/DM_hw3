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
      Caption         =   "5-fold cross validation(total attributes)"
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
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   6975
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
Dim attr9array(9) As String
Dim allattribute() As Double



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

Static Function topNindex(ByRef distarray() As Double, ByRef distindexarray() As Double)
Dim tempdistarray() As Double
Dim tempdistindexarray() As Double
Dim distmin As Double
Dim indexmin() As Double
ReDim indexmin(nei_number - 1)
tempdistarray() = distarray()
tempdistindexarray() = distindexarray()
'這裡卡卡的
For j = 0 To (nei_number - 1)
distmin = 100
'indexmin = 100

For i = 0 To UBound(tempdistarray)
If distmin > tempdistarray(i) Then
distmin = tempdistarray(i)
indexmin(j) = tempdistindexarray(i)
End If
Next i

For k = 0 To UBound(tempdistindexarray)
If indexmin(j) = tempdistindexarray(k) Then
tempdistarray(k) = 90
End If
Next k

Next j

topNindex = indexmin()
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
Static Function vote(ByRef candidateclass() As String)
Dim tempcandidateclass() As String
Dim eachclassnum(9) As Double
Dim finalclass As String
tempcandidateclass() = candidateclass()
For i = 0 To 9
eachclassnum(i) = 0
Next i

For i = 0 To UBound(tempcandidateclass)
For j = 0 To 9
If tempcandidateclass(i) = attr9array(j) Then
eachclassnum(j) = eachclassnum(j) + 1
End If
Next j
Next i

Select Case nei_number
    Case 3
        Dim candidate(2) As String
        Dim c As Double
        Dim rndindex As Double
        c = 0
        For i = 0 To 9
        
        If eachclassnum(i) >= 2 Then
        finalclass = attr9array(i)
        GoTo voteend
        End If
        
        If eachclassnum(i) = 1 Then
        candidate(c) = attr9array(i)
        c = c + 1
        End If
        
        Next i
        rndindex = rndclass(nei_number)
        finalclass = candidate(rndindex) '不懂為啥設為0後1-8跟8-1依舊不同
        GoTo voteend
        
    Case 4
        Dim twocounter As Double
        Dim onecounter As Double
        Dim rndindex4 As Double
        Dim candidatetwo(1) As String
        Dim candidatefour(3) As String
        twocounter = 0
        onecounter = 0
        
        For i = 0 To 9
        
        If eachclassnum(i) = 1 Then
        candidatefour(onecounter) = attr9array(i)
        onecounter = onecounter + 1
        End If
        
        If eachclassnum(i) > 2 Then
        finalclass = attr9array(i)
        GoTo voteend
        End If
        
        If eachclassnum(i) = 2 Then
        finalclass = attr9array(i)
        candidatetwo(twocounter) = attr9array(i)
        twocounter = twocounter + 1
        End If
        
        If twocounter = 2 Then
        rndindex4 = rndclass(twocounter)
        finalclass = candidatetwo(rndindex4)
        GoTo voteend
        End If
        
        If onecounter = 4 Then
        rndindex4 = rndclass(onecounter)
        finalclass = candidatefour(rndindex4)
        GoTo voteend
        End If
        
        Next i
        GoTo voteend
    Case 5
        Dim twocounter5 As Double
        Dim onecounter5 As Double
        Dim rndindex5 As Double
        Dim candidatetwo5(1) As String
        Dim candidatefive5(4) As String
        twocounter5 = 0
        onecounter5 = 0
        
        For i = 0 To 9
        
        If eachclassnum(i) = 1 Then
        candidatefive5(onecounter5) = attr9array(i)
        onecounter5 = onecounter5 + 1
        End If
        
        If eachclassnum(i) >= 3 Then
        finalclass = attr9array(i)
        GoTo voteend
        End If
        
        If eachclassnum(i) = 2 Then
        finalclass = attr9array(i)
        candidatetwo5(twocounter5) = attr9array(i)
        twocounter5 = twocounter5 + 1
        End If
        
        If twocounter5 = 2 Then
        rndindex5 = rndclass(twocounter5)
        finalclass = candidatetwo5(rndindex5)
        GoTo voteend
        End If
        
        If onecounter5 = 5 Then
        rndindex5 = rndclass(onecounter5)
        finalclass = candidatefive5(rndindex5)
        GoTo voteend
        End If
        
        Next i
        GoTo voteend
    Case 6
        Dim threecounter6 As Double
        Dim twocounter6 As Double
        Dim onecounter6 As Double
        Dim rndindex6 As Double
        Dim flag As Double
        Dim candidatetwo6(2) As String
        Dim candidatethree6(1) As String
        Dim candidatesix6(5) As String
        flag = 0
        twocounter6 = 0
        onecounter6 = 0
        threecounter6 = 0
        
        For i = 0 To 9
        
        If eachclassnum(i) = 1 Then
        candidatesix6(onecounter6) = attr9array(i)
        onecounter6 = onecounter6 + 1
        End If
        
        If eachclassnum(i) > 3 Then
        finalclass = attr9array(i)
        GoTo voteend
        End If
        
        If eachclassnum(i) = 3 Then
        flag = 1
        finalclass = attr9array(i)
        candidatethree6(threecounter6) = attr9array(i)
        threecounter6 = threecounter6 + 1
        End If
        
        If threecounter6 = 2 Then
        rndindex6 = rndclass(threecounter6)
        finalclass = candidatethree6(rndindex6)
        GoTo voteend
        End If
        
        If eachclassnum(i) = 2 And flag = 0 Then
        finalclass = attr9array(i)
        candidatetwo6(twocounter6) = attr9array(i)
        twocounter6 = twocounter6 + 1
        End If
        
        If twocounter6 = 2 And onecounter6 = 2 Then
        rndindex6 = rndclass(twocounter6)
        finalclass = candidatetwo6(rndindex6)
        GoTo voteend
        End If
        
        If twocounter6 = 3 Then
        rndindex6 = rndclass(twocounter6)
        finalclass = candidatetwo6(rndindex6)
        GoTo voteend
        End If
        
        If onecounter6 = 6 Then
        rndindex6 = rndclass(onecounter6)
        finalclass = candidatesix6(rndindex6)
        GoTo voteend
        End If
        
        Next i
End Select
voteend:
vote = finalclass
End Function
Static Function rndclass(ByVal num As Double)
Dim returnnum As Double
Randomize (Timer)
returnnum = Int(Rnd() * num)
rndclass = returnnum
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
Dim classarray() As String
Dim finalclass As String
Dim sortresult As String
ReDim classarray(nei_number - 1)


attributesubset() = allattribute() '之後不同feature個數的時候
tttemptrainarray() = ttemptrainarray()
'測試
'For i = 0 To UBound(attributesubset)
'List1.AddItem attributesubset(i)
'Next i


ReDim distarray(UBound(tttemptrainarray)) '存trainarray每點的距離
ReDim distarrayindex(UBound(tttemptrainarray))
ReDim topndist(nei_number - 1) '存前n近的值
'ReDim topndistindex(nei_number - 1) '存前n近的index

'List1.AddItem UBound(tttemptrainarray) + 1

For i = 0 To UBound(tttemptrainarray)
distarray(i) = distance(testoneindex, tttemptrainarray(i), attributesubset)
distarrayindex(i) = tttemptrainarray(i)
'List1.AddItem distarrayindex(i)
'List1.AddItem distarray(i)
Next i


'用排序選topn太慢
'sortresult = sortrnd(distarrayindex, distarray)
''看排序結果
'For i = 0 To UBound(distarrayindex)
'List1.AddItem distarray(i)
'List1.AddItem distarrayindex(i)
'Next i
'List1.AddItem ""

'topn
topndistindex() = topNindex(distarray, distarrayindex)
'看topn
'For i = 0 To UBound(topndistindex)
'List1.AddItem topndistindex(i)
'Next i

'把前k小的距離index保存下來
'For i = 0 To (nei_number - 1)
'classarray(i) = distarrayindex(i)
''List1.AddItem classarray(i)
'Next i


'For i = 0 To (nei_number - 1)
'tempgmax = gmax(distarray, distarrayindex)
'tempindexattr = Split(tempgmax, ",")
'topndistindex(i) = tempindexattr(0)
'
''把top1設定成-1
'For j = 0 To UBound(tttemptrainarray)
'If (tttemptrainarray(j)) = tempindexattr(0) Then
'    tttemptrainarray(j) = -1
'    distarray(j) = -1
'End If
'Next j
'
'Next i
'
'用topndistindex()找到對應的class
For i = 0 To UBound(topndistindex)
    classarray(i) = data2darray(9, topndistindex(i))
Next i

'比較class然後投票決定predclass是什麼
finalclass = vote(classarray)
'GoTo predclassend

'predclassend:
predclass = finalclass

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

For i = 0 To UBound(temptestarray)
predictclass = predclass(temptestarray(i), temptrainarray)
'predictclass = data2darray(9, temptestarray(i)) '測試用
If data2darray(9, temptestarray(i)) = predictclass Then
correctnum = correctnum + 1
End If
Next i

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
Static Function correctratesutset(ByRef attrsebset() As Double)
Dim trainindexstring As String
Dim trainnumber As Double
Dim testrnd As Double
Dim subsetcounter As Double
Dim eachtrainsubset(3) As Double
Dim correctratearray(4) As Double
Dim eachtrainDouble() As Double
Dim eachtestDouble() As Double
Dim avgcorrectrate As Double
avgcorrectrate = 0
allattribute() = attrsebset()

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

ReDim eachtestDouble(UBound(eachtest))
ReDim eachtrainDouble(UBound(eachtrain))

For j = 0 To UBound(eachtest)
eachtestDouble(j) = CDbl(eachtest(j))
Next j

For j = 0 To UBound(eachtrain)
eachtrainDouble(j) = CDbl(eachtrain(j))
Next j

correctratearray(i) = correctrate(eachtestDouble, eachtrainDouble)

Next i

For i = 0 To 4
avgcorrectrate = avgcorrectrate + correctratearray(i)
Next i
avgcorrectrate = (avgcorrectrate / 5)

correctratesutset = avgcorrectrate
End Function

Private Sub backward_Click()
'測試correctratesutset
List1.Clear
'Dim subset(6) As Double
'Dim result As Double
'For i = 0 To UBound(subset)
'subset(i) = i + 1
'Next i
'result = correctratesutset(subset)
'List1.AddItem result
'
Dim subset1(7) As Double
Dim result1 As Double
For i = 0 To UBound(subset1)
subset1(i) = (8 - i)
Next i
result1 = correctratesutset(subset1)
List1.AddItem result1

End Sub

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
Dim avgcorrectrate As Double
avgcorrectrate = 0

'這是在8個屬性都選的情況下
ReDim allattribute(7)
For i = 0 To 7
allattribute(i) = i + 1
Next i

'測試topNindex(dist,index)
'Dim index(4) As Double
'Dim dist(4) As Double
'Dim result() As Double
'index(0) = 1
'index(1) = 2
'index(2) = 3
'index(3) = 4
'index(4) = 5
'dist(0) = 7
'dist(1) = 8
'dist(2) = 3
'dist(3) = 9
'dist(4) = 2
'result() = topNindex(dist, index)
'For i = 0 To UBound(result)
'List1.AddItem result(i)
'Next i

'測試rndclass
'Dim num As Double
'Dim result As Double
'num = 4
'result = rndclass(num)
'List1.AddItem result

'測試vote
'Dim votestring(2) As String
'Dim result As String
''CYT,NUC,MIT,VAC,POX,ERL
'votestring(0) = "ERL"
'votestring(1) = "POX"
'votestring(2) = "POX"
''votestring(3) = "POX"
''votestring(4) = "POX"
''votestring(5) = "POX"
'result = vote(votestring)
'List1.AddItem result

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
'dimnum(3) = 7
'dimnum(6) = 4
'distanceDouble = distance(xi, yi, dimnum)
'List1.AddItem distanceDouble


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
'For i = 0 To 4
'List1.AddItem totalfivefoldindex(i)
'Next i


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

'看eachtest長度
'List1.AddItem UBound(eachtest) + 1

ReDim eachtestDouble(UBound(eachtest))
ReDim eachtrainDouble(UBound(eachtrain))

For j = 0 To UBound(eachtest)
eachtestDouble(j) = CDbl(eachtest(j))
Next j

For j = 0 To UBound(eachtrain)
eachtrainDouble(j) = CDbl(eachtrain(j))
Next j
'測試每次的test和train的筆數
'If i = 0 Then
'List1.AddItem UBound(eachtestDouble) + 1
'List1.AddItem UBound(eachtrainDouble) + 1
'List1.AddItem "------------------------------"
'For j = 0 To UBound(eachtestDouble)
'List1.AddItem eachtestDouble(j)
'Next j
'List1.AddItem "----------------------------------------------------"
'For j = 0 To UBound(eachtrainDouble)
'List1.AddItem eachtrainDouble(j)
'Next j
'List1.AddItem "----------------------------------------------------"
'correctratearray(i) = correctrate(eachtestDouble, eachtrainDouble)
'List1.AddItem correctratearray(i)

'測試predictclass
'Dim testpredclass As String
'testpredclass = predclass(6, eachtrainDouble)
'List1.AddItem testpredclass
'List1.AddItem data2darray(9, 6)
'
'End If 'i=0



correctratearray(i) = correctrate(eachtestDouble, eachtrainDouble)
'List1.AddItem correctratearray(i)


Next i

For i = 0 To 4
avgcorrectrate = avgcorrectrate + correctratearray(i)
Next i
avgcorrectrate = (avgcorrectrate / 5)

List1.AddItem avgcorrectrate

'GoTo endd
'endd:
End Sub

Private Sub datanumber_Change()
datanum = CInt(datanumber.Text)
End Sub

Private Sub datatxt_Change()
file = datatxt.Text
End Sub

Static Function unpickAttr(ByRef resultattrs() As Double)
Dim tempresultattrs() As Double
Dim allattr(7) As Double
Dim unpickAttrs() As Double
Dim counter As Double
counter = 0
tempresultattrs() = resultattrs()

For i = 0 To 7
    If tempresultattrs(i) = -1 Then
        ReDim unpickAttrs(7 - i)
        Exit For
    End If
Next i

For i = 0 To 7
    allattr(i) = i + 1 '注意
Next i

For i = 0 To 7
    If tempresultattrs(i) <> -1 Then
        allattr(tempresultattrs(i) - 1) = 10 '偵錯
    End If
Next i

For i = 0 To 7
    If allattr(i) <> 10 Then
        unpickAttrs(counter) = allattr(i)
        counter = counter + 1
    End If
Next i

unpickAttr = unpickAttrs()

End Function

Private Sub forward_Click()
List1.Clear
'result = correctratesutset(subset)
Dim resultattr(7) As Double '每個set的最大值新選的那個attr
Dim resultmaxvalue(7) As Double '每個set數量的最大值
Dim setnum() As Double

Dim tempgoodness() As Double
Dim tempattr() As Double

For i = 0 To 7
resultattr(i) = -1
resultmaxvalue(i) = -1
Next i

For i = 0 To 7
'先決定要丟入subsetgoodness的陣列長度
ReDim setnum(i) '丟入subsetgoodness的陣列
Dim tempunpickAttr() As Double
ReDim tempgoodness(7 - i)
ReDim tempattr(7 - i)



'接著為該陣列塞入目前已選的attr

If i = 0 Then
    GoTo izero
End If
'把已經選好的resultattr丟給setnum
   For j = 0 To (UBound(setnum) - 1)
       setnum(j) = resultattr(j) '注意 一下
   Next j
izero:
'跑unpickAttr()回傳還未被選的attr陣列
tempunpickAttr() = unpickAttr(resultattr)

'從0-7變成1-8
'For j = 0 To UBound(tempunpickAttr)
''tempunpickAttr(j) = tempunpickAttr(j) + 1
'List1.AddItem tempunpickAttr(j)
'Next j
'List1.AddItem ""



'跑回圈幫setnum(i)塞入不同還未被選的attr
    For j = 0 To UBound(tempunpickAttr)
        setnum(i) = tempunpickAttr(j) '幫setnum挑一個還沒進來的attr
        tempattr(j) = tempunpickAttr(j) '丟進gmax用的
        tempgoodness(j) = correctratesutset(setnum)
    Next j
tempgmaxstr = gmax(tempgoodness, tempattr)
tempav = Split(tempgmaxstr, ",")
resultattr(i) = CDbl(tempav(0))
resultmaxvalue(i) = CDbl(tempav(1))


'停止條件
'If i > 0 Then
'If resultmaxvalue(i) < resultmaxvalue(i - 1) Then
'Exit For
'End If
'End If

'GoTo test
'test:

Next i

'把最終結果印出來
For i = 0 To UBound(resultattr)
List1.AddItem resultattr(i)
List1.AddItem resultmaxvalue(i)
List1.AddItem ""
Next i

'List1.AddItem "attribute:" + CStr(resultattr(0) + 1)
'List1.AddItem resultmaxvalue(0)
'
'List1.AddItem "attribute:" + CStr(resultattr(0) + 1) + CStr(resultattr(1) + 1)
'
'List1.AddItem resultmaxvalue(1)
'
'List1.AddItem "attribute:" + CStr(resultattr(0) + 1) + CStr(resultattr(1) + 1) + CStr(resultattr(2) + 1)
'List1.AddItem resultmaxvalue(2)
'
'List1.AddItem "attribute:" + CStr(resultattr(0) + 1) + CStr(resultattr(1) + 1) + CStr(resultattr(2) + 1) + CStr(resultattr(3) + 1)
'List1.AddItem resultmaxvalue(3)
'
'List1.AddItem "attribute:" + CStr(resultattr(0) + 1) + CStr(resultattr(1) + 1) + CStr(resultattr(2) + 1) + CStr(resultattr(3) + 1) + CStr(resultattr(4) + 1)
'List1.AddItem resultmaxvalue(4)
'If choicefb = 0 Then
'GoTo choicefbzero
'End If
'List1.AddItem "attribute:" + CStr(resultattr(0) + 1) + CStr(resultattr(1) + 1) + CStr(resultattr(2) + 1) + CStr(resultattr(3) + 1) + CStr(resultattr(4) + 1) + CStr(resultattr(5) + 1)
'List1.AddItem resultmaxvalue(5)
'choicefbzero:



'動態陣列測試
'Dim testnum() As Double
'For i = 0 To 3
'ReDim testnum(i)
'List1.AddItem UBound(testnum)
'Next i

End Sub

Private Sub nerghbors_num_Change()
nei_number = CInt(nerghbors_num.Text)
End Sub

Private Sub read_Click()
List1.Clear
Dim datacounter As Integer
Dim temp() As String
attr9array(0) = "CYT"
attr9array(1) = "NUC"
attr9array(2) = "MIT"
attr9array(3) = "ME3"
attr9array(4) = "ME2"
attr9array(5) = "ME1"
attr9array(6) = "EXC"
attr9array(7) = "VAC"
attr9array(8) = "POX"
attr9array(9) = "ERL"

'List1.AddItem nei_number
'List1.AddItem ""

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
