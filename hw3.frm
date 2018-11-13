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
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton generatefivefold 
      Caption         =   "Generate five fold"
      Height          =   375
      Left            =   8040
      TabIndex        =   8
      Top             =   3960
      Width           =   1695
   End
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
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton forward 
      Caption         =   "forward"
      Height          =   615
      Left            =   7200
      TabIndex        =   4
      Top             =   4800
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
'�ާ@�覡����
'txt�ɦW/��Ƶ���/k��(3,4,5,6)�ݭn��ʿ�J

'��8���ݩʪ�5-fold�b���Pk�Ȫ����T�v��,�Ъ`�N�C��@��k�ȳ��n���srun�{��
'���k����:read->5-fold cross validation(total attributes)(�O�o��k��)
'���ݮɶ��j��10-15��

'���ۦp�G�n�ݤ��Pk�Ȧbforward,backward�����T�v�]���ݭn�N�{���������srun
'���k�Ҭ�:read->Generate five fold->forward or backward(�O�o��k��)
'���ݮɶ�foward�j�h���W�L4��30��,backward�j�h���W�L3��30��

Dim file As String
Dim datanum As Integer
Dim nei_number As Integer
Dim data2darray(9, 1483) As String
Dim totalfivefoldindex() As String
'Dim attr9array(9) As String
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
If distmax < tempdistarray(i) Then
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
'�o�̥d�d��
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
Dim selectattr() As String
Dim eachclassnum() As Double
Dim finalclass As String
Dim selectattrcounter As Double
Dim flagbegin As Double

tempcandidateclass() = candidateclass()
ReDim eachclassnum(nei_number - 1)
ReDim selectattr(nei_number - 1)
selectattrcounter = 0
flagbegin = 0

For i = 0 To UBound(selectattr)
eachclassnum(i) = 0
selectattr(i) = "0"
Next i

For i = 0 To UBound(tempcandidateclass)
    flagbegin = 0
    If i = 0 Then
    selectattr(0) = tempcandidateclass(0)
    selectattrcounter = selectattrcounter + 1 '�n�A�Q�Q
    eachclassnum(0) = eachclassnum(0) + 1
    GoTo zero
    End If
    
    For j = 0 To i
        If selectattr(j) = tempcandidateclass(i) Then
        eachclassnum(j) = eachclassnum(j) + 1
        flagbegin = 1
        End If
    Next j
    
    If flagbegin = 0 Then
    selectattr(selectattrcounter) = tempcandidateclass(i)
    eachclassnum(selectattrcounter) = eachclassnum(selectattrcounter) + 1
    selectattrcounter = selectattrcounter + 1
    End If
zero:
Next i

'debug
'For i = 0 To UBound(selectattr)
'List1.AddItem selectattr(i)
'List1.AddItem eachclassnum(i)
'Next i



'�o�O�ª�
'For i = 0 To UBound(tempcandidateclass)
'For j = 0 To 9
'If tempcandidateclass(i) = attr9array(j) Then
'eachclassnum(j) = eachclassnum(j) + 1
'End If
'Next j
'Next i



Select Case nei_number
    Case 3
        For i = 0 To (nei_number - 1)
        
        If eachclassnum(i) >= 2 Then
        finalclass = selectattr(i)
        GoTo voteend
        End If
        
        Next i
 
        finalclass = selectattr(0) '������ԣ�]��0��1-8��8-1���¤��P
        GoTo voteend
        
    Case 4
        Dim twocounter As Double
        Dim onecounter As Double
        Dim candidatetwo(1) As String
        Dim candidatefour(3) As String
        twocounter = 0
        onecounter = 0
        
        For i = 0 To (nei_number - 1)
        
        If eachclassnum(i) = 1 Then
        candidatefour(onecounter) = selectattr(i)
        onecounter = onecounter + 1
        End If
        
        If eachclassnum(i) > 2 Then
        finalclass = selectattr(i)
        GoTo voteend
        End If
        
        If eachclassnum(i) = 2 Then
        candidatetwo(twocounter) = selectattr(i)
        finalclass = candidatetwo(0)
        twocounter = twocounter + 1
        End If
        
        If twocounter = 2 Then
        GoTo voteend
        End If
        
        If onecounter = 4 Then
        finalclass = candidatefour(0)
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
        
        For i = 0 To (nei_number - 1)
        
        If eachclassnum(i) = 1 Then
        candidatefive5(onecounter5) = selectattr(i)
        onecounter5 = onecounter5 + 1
        End If
        
        If eachclassnum(i) >= 3 Then
        finalclass = selectattr(i)
        GoTo voteend
        End If
        
        If eachclassnum(i) = 2 Then
        candidatetwo5(twocounter5) = selectattr(i)
        finalclass = candidatetwo5(0)
        twocounter5 = twocounter5 + 1
        End If
        
        If twocounter5 = 2 Then
        GoTo voteend
        End If
        
        If onecounter5 = 5 Then
        finalclass = candidatefive5(0)
        GoTo voteend
        End If
        
        Next i
        GoTo voteend
    Case 6
        Dim threecounter6 As Double
        Dim twocounter6 As Double
        Dim onecounter6 As Double
        Dim flag As Double
        Dim candidatetwo6(2) As String
        Dim candidatethree6(1) As String
        Dim candidatesix6(5) As String
        flag = 0
        twocounter6 = 0
        onecounter6 = 0
        threecounter6 = 0
        
        For i = 0 To (nei_number - 1)
        
        If eachclassnum(i) = 1 Then
        candidatesix6(onecounter6) = selectattr(i)
        onecounter6 = onecounter6 + 1
        End If
        
        If eachclassnum(i) > 3 Then
        finalclass = selectattr(i)
        GoTo voteend
        End If
        
        If eachclassnum(i) = 3 Then
        flag = 1 '�Ψӿ��� 3 2 1�����p
        candidatethree6(threecounter6) = selectattr(i)
        finalclass = candidatethree6(0)
        threecounter6 = threecounter6 + 1
        End If
        
        If threecounter6 = 2 Then
        GoTo voteend
        End If
        
        If eachclassnum(i) = 2 And flag = 0 Then
        candidatetwo6(twocounter6) = selectattr(i)
        finalclass = candidatetwo6(0)
        twocounter6 = twocounter6 + 1
        End If
        
'        If twocounter6 = 2 And onecounter6 = 2 Then
'        rndindex6 = rndclass(twocounter6)
'        finalclass = candidatetwo6(rndindex6)
'        GoTo voteend
'        End If
        
        If twocounter6 = 3 Then
        GoTo voteend
        End If
        
        If onecounter6 = 6 Then
        finalclass = candidatesix6(0)
        GoTo voteend
        End If
        
        Next i
End Select
voteend:
'GoTo debugend
'debugend:
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


attributesubset() = allattribute() '���ᤣ�Pfeature�Ӽƪ��ɭ�
tttemptrainarray() = ttemptrainarray()
'����
'For i = 0 To UBound(attributesubset)
'List1.AddItem attributesubset(i)
'Next i


ReDim distarray(UBound(tttemptrainarray)) '�strainarray�C�I���Z��
ReDim distarrayindex(UBound(tttemptrainarray))
ReDim topndist(nei_number - 1) '�s�en�񪺭�
'ReDim topndistindex(nei_number - 1) '�s�en��index

'List1.AddItem UBound(tttemptrainarray) + 1

For i = 0 To UBound(tttemptrainarray)
distarray(i) = distance(testoneindex, tttemptrainarray(i), attributesubset)
distarrayindex(i) = tttemptrainarray(i)
'List1.AddItem distarrayindex(i)
'List1.AddItem distarray(i)
Next i


'�αƧǿ�topn�ӺC
'sortresult = sortrnd(distarrayindex, distarray)
''�ݱƧǵ��G
'For i = 0 To UBound(distarrayindex)
'List1.AddItem distarray(i)
'List1.AddItem distarrayindex(i)
'Next i
'List1.AddItem ""

'topn
topndistindex() = topNindex(distarray, distarrayindex)
'��topn
'For i = 0 To UBound(topndistindex)
'List1.AddItem topndistindex(i)
'Next i

'��ek�p���Z��index�O�s�U��
'For i = 0 To (nei_number - 1)
'classarray(i) = distarrayindex(i)
''List1.AddItem classarray(i)
'Next i


'For i = 0 To (nei_number - 1)
'tempgmax = gmax(distarray, distarrayindex)
'tempindexattr = Split(tempgmax, ",")
'topndistindex(i) = tempindexattr(0)
'
''��top1�]�w��-1
'For j = 0 To UBound(tttemptrainarray)
'If (tttemptrainarray(j)) = tempindexattr(0) Then
'    tttemptrainarray(j) = -1
'    distarray(j) = -1
'End If
'Next j
'
'Next i
'
'��topndistindex()��������class
For i = 0 To UBound(topndistindex)
    classarray(i) = data2darray(9, topndistindex(i))
Next i

'���class�M��벼�M�wpredclass�O����
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
'predictclass = data2darray(9, temptestarray(i)) '���ե�
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
    For j = 0 To i - 1 '����i�Ӳ��ͪ����X,�򤧫e�w�g���͹L�����X���, �p�G���ơA�K���s���,�]���n��i=i-1(�˰h�@��)
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
''���禳�S������
''If i = 0 Then
''    GoTo ggg
''End If
''    For j = 0 To i - 1 '����i�Ӳ��ͪ����X,�򤧫e�w�g���͹L�����X���, �p�G���ơA�K���s���,�]���n��i=i-1(�˰h�@��)
''        If dataindex(i) = dataindex(j) Then
''            List1.AddItem "----------------------------------"
''        End If
''    Next j
''ggg:
'Next i

For i = 0 To 4
    eachfoldindex = ""
    For j = 0 To 296
    If i = 4 And j = 296 Then '��5��fold�O296��
        Exit For
    End If
    eachfoldindex = eachfoldindex + CStr(dataindex(counter)) + " "
    counter = counter + 1
    Next j
    foldindex(i) = eachfoldindex
Next i
fivefoldindex = foldindex() '�^�Ǥ��n5�ժ�index���r��
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

'totalfivefoldindex() = fivefoldindex()

'��@�ӷ���ո��,��L�|�ӦX���V�m���
For i = 0 To 4
Dim eachtrain() As String
Dim eachtrainstring As String
Dim eachtest() As String
Dim eachteststring As String
subsetcounter = 0
eachtrainstring = ""
eachteststring = ""

'�Ϥ��X�C�����V�m�H�δ��ո��
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
'����correctratesutset
List1.Clear

'Dim subset(0) As Double
'Dim result As Double
'For i = 0 To 7
'subset(0) = i + 1
'result = correctratesutset(subset)
'List1.AddItem result
'Next i

'
'Dim subset1(7) As Double
'Dim result1 As Double
'For i = 0 To UBound(subset1)
'subset1(i) = i + 1
'Next i
''subset1(0) = 8
''subset1(7) = 1
'result1 = correctratesutset(subset1)
'List1.AddItem result1

'����unpick
'Dim test(7) As Double
'Dim result() As Double
'For i = 0 To 7
'test(i) = -1
'Next i
'test(0) = 3
'test(1) = 8
'result() = unpickAttr(test)
'List1.AddItem UBound(result) + 1
'For i = 0 To UBound(result)
'List1.AddItem result(i)
'Next i

'����gmax
'gmax(ByRef distarray() As Double, ByRef distindexarray() As Double)
'Dim dist(2) As Double
'Dim index(2) As Double
'Dim result As String
'index(0) = 1
'index(1) = 2
'index(2) = 3
'dist(0) = 13
'dist(1) = 9
'dist(2) = 11
'result = gmax(dist, index)
'List1.AddItem result

Dim resultattr(7) As Double '�O���C���𱼪����@��
Dim resultmaxvalue(7) As Double '�C��set�ƶq���̤j��
Dim deleteone() As Double '�C�^�X�R���Y�Ӥ��᪺subset
Dim tempgoodness() As Double
Dim tempattr() As Double
Dim counter As Double
Dim tempgmaxstr As String
Dim tempav() As String

For i = 0 To 7
resultattr(i) = -1
resultmaxvalue(i) = -1
Next i

For i = 0 To 7
'���M�w�n��Jsubsetgoodness���}�C����

ReDim deleteone(6 - i) '�C�^�X�R���Y�Ӥ��᪺subset,�n��J�⥿�T�v
Dim setnum() As Double '�ǳƭn�����@��attr�e�����㪩
ReDim tempgoodness(7 - i)
ReDim tempattr(7 - i) '�Կ�Q��attr




'���۬��Ӱ}�C��J�ثe�w�諸attr

'��X�n�d�U�Ӫ�attr
setnum() = unpickAttr(resultattr)




'�C��attr���R�ݬ�
    For j = 0 To UBound(setnum)
        If j = 0 Then
            GoTo jzero
        End If
        setnum(j - 1) = tempattr(j - 1)
jzero:
        tempattr(j) = setnum(j)
        setnum(j) = -1 '��setnum�R���Y�@��attr
        counter = 0
        For k = 0 To UBound(setnum)
            If setnum(k) <> -1 Then
                deleteone(counter) = setnum(k)
                counter = counter + 1
            End If
        Next k
         '�s�R�������@�Ӫ�attr
        tempgoodness(j) = correctratesutset(deleteone)
    Next j
tempgmaxstr = gmax(tempgoodness, tempattr)
tempav = Split(tempgmaxstr, ",")
resultattr(i) = CDbl(tempav(0))
resultmaxvalue(i) = CDbl(tempav(1))


'�������
If i > 0 Then
If resultmaxvalue(i) < resultmaxvalue(i - 1) Then
Exit For
End If
End If

'�H���U�@�]�ܤ[
If i = 2 Then
Exit For
End If

Next i

'��̲׵��G�L�X��
'For i = 0 To UBound(resultattr)
'List1.AddItem resultattr(i)
'List1.AddItem resultmaxvalue(i)
'List1.AddItem ""
'Next i
'GoTo backend
'backend:

'�L�X���G����
'Dim ttresultattr(7) As Double
'Dim ttresultmaxvalue(7) As Double
'For i = 0 To 7
'ttresultattr(i) = -1
'ttresultmaxvalue(i) = -1
'Next i
'ttresultattr(0) = 5
'ttresultattr(1) = 8
'ttresultattr(2) = 3
'ttresultmaxvalue(0) = 11
'ttresultmaxvalue(1) = 12
'ttresultmaxvalue(2) = 13
'�L�X���G
Dim resultback As String
resultback = printback(resultattr, resultmaxvalue)
List1.AddItem resultback

'Dim inputarr(7) As Double
'For i = 0 To 5
'
'    Dim printresult() As Double
'    Dim attrstr As String
'    attrstr = ""
'    For j = 0 To 7
'    inputarr(j) = -1
'    Next j
'
'    For j = 0 To i
'    inputarr(j) = resultattr(j)
'    Next j
'    printresult() = unpickAttr(inputarr)
'    For j = 0 To UBound(printresult)
'    attrstr = attrstr + CStr(printresult(j) + 1)
'    Next j
'    List1.AddItem "attribute:" + attrstr
'    List1.AddItem resultmaxvalue(i)
'    If i = 0 Then
'    GoTo nozero
'    End If
'
'    If resultmaxvalue(i) < resultmaxvalue(i - 1) Then
'    Exit For
'    End If
'
'nozero:
'Next i

'�ʺA�}�C����
'Dim testnum() As Double
'For i = 0 To 3
'ReDim testnum(i)
'List1.AddItem UBound(testnum)
'Next i

'totalsetb = "End"


End Sub
Static Function printback(ByRef tresultattr() As Double, ByRef tresultmaxvalue() As Double)
Dim tempresultattr() As Double
Dim tempresultmaxvalue() As Double
Dim attrstring As String
Dim maxvaluetring As String
Dim inputunpick(7) As Double
Dim outputunpick() As Double
tempresultattr() = tresultattr()
tempresultmaxvalue() = tresultmaxvalue()
Dim tattrarray(7) As Double
Dim tattrresult As Double
For i = 0 To UBound(tattrarray)
tattrarray(i) = i + 1
Next i
tattrresult = correctratesutset(tattrarray)

List1.AddItem "Attribute: 1 2 3 4 5 6 7 8"
List1.AddItem "Correct rate: " + CStr(tattrresult)
List1.AddItem ""

For i = 0 To UBound(tempresultattr)
If tempresultattr(i) = -1 Then
Exit For
End If

'ReDim inputunpick(i)
For j = 0 To UBound(inputunpick)
inputunpick(j) = -1
Next j
attrstring = "Attribute:"
maxvaluetring = "Correct rate: "

For j = 0 To i
inputunpick(j) = tempresultattr(j)
Next j
outputunpick = unpickAttr(inputunpick)

For j = 0 To UBound(outputunpick)
attrstring = attrstring + " " + CStr(outputunpick(j))
Next j
maxvaluetring = maxvaluetring + " " + CStr(tempresultmaxvalue(i))

List1.AddItem attrstring
List1.AddItem maxvaluetring
List1.AddItem ""
Next i
printback = "End"
End Function
Private Sub cross_validation_Click()
List1.Clear
'Dim tempfivefoldindex() As String '�s5�զr��
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

'�o�O�b8���ݩʳ��諸���p�U
ReDim allattribute(7)
For i = 0 To 7
allattribute(i) = i + 1
Next i

'����topNindex(dist,index)
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

'����rndclass
'Dim num As Double
'Dim result As Double
'num = 4
'result = rndclass(num)
'List1.AddItem result

'����vote
'Dim votestring(5) As String
'Dim result As String
''CYT,NUC,MIT,VAC,POX,ERL
'votestring(0) = "POX"
'votestring(1) = "ERL"
'votestring(2) = "ERL"
'votestring(3) = "VAC"
'votestring(4) = "VAC"
'votestring(5) = "MIT"
'result = vote(votestring)
'List1.AddItem result

'����distance
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


'����rnd
'Randomize (Timer)
'testrnd = Rnd()
'List1.AddItem testrnd

'����sortrnd
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

'�^�Ǥ��n���������index
totalfivefoldindex() = fivefoldindex()
'For i = 0 To 4
'List1.AddItem totalfivefoldindex(i)
'Next i


'��@�ӷ���ո��,��L�|�ӦX���V�m���
For i = 0 To 4
Dim eachtrain() As String
Dim eachtrainstring As String
Dim eachtest() As String
Dim eachteststring As String
subsetcounter = 0
eachtrainstring = ""
eachteststring = ""
'eachfoldarray() = Split(totalfivefoldindex(i), " ")

'�Ϥ��X�C�����V�m�H�δ��ո��
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

'��eachtest����
'List1.AddItem UBound(eachtest) + 1

ReDim eachtestDouble(UBound(eachtest))
ReDim eachtrainDouble(UBound(eachtrain))

For j = 0 To UBound(eachtest)
eachtestDouble(j) = CDbl(eachtest(j))
Next j

For j = 0 To UBound(eachtrain)
eachtrainDouble(j) = CDbl(eachtrain(j))
Next j
'���ըC����test�Mtrain������
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

'����predictclass
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
    allattr(i) = i + 1 '�`�N
Next i

For i = 0 To 7
    If tempresultattrs(i) <> -1 Then
        allattr(tempresultattrs(i) - 1) = 10 '����
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
Dim resultattr(7) As Double '�C��set���̤j�ȷs�諸����attr
Dim resultmaxvalue(7) As Double '�C��set�ƶq���̤j��
Dim setnum() As Double
Dim tempgmaxstr As String
Dim tempav() As String


Dim tempgoodness() As Double
Dim tempattr() As Double

For i = 0 To 7
resultattr(i) = -1
resultmaxvalue(i) = -1
Next i

For i = 0 To 7
'���M�w�n��Jsubsetgoodness���}�C����
ReDim setnum(i) '��i�h�⥿�T�v
Dim tempunpickAttr() As Double
ReDim tempgoodness(7 - i)
ReDim tempattr(7 - i)



'���۬��Ӱ}�C��J�ثe�w�諸attr

If i = 0 Then
    GoTo izero
End If
'��w�g��n��resultattr�ᵹsetnum
   For j = 0 To (UBound(setnum) - 1)
       setnum(j) = resultattr(j) '�`�N �@�U
   Next j
izero:
'�]unpickAttr()�^���٥��Q�諸attr�}�C
tempunpickAttr() = unpickAttr(resultattr)

'�q0-7�ܦ�1-8
'For j = 0 To UBound(tempunpickAttr)
''tempunpickAttr(j) = tempunpickAttr(j) + 1
'List1.AddItem tempunpickAttr(j)
'Next j
'List1.AddItem ""



'�]�^����setnum(i)��J���P�٥��Q�諸attr
    For j = 0 To UBound(tempunpickAttr)
        setnum(i) = tempunpickAttr(j) '��setnum�D�@���٨S�i�Ӫ�attr
        tempattr(j) = tempunpickAttr(j) '��igmax�Ϊ�
        tempgoodness(j) = correctratesutset(setnum)
    Next j
    
tempgmaxstr = gmax(tempgoodness, tempattr)
tempav = Split(tempgmaxstr, ",")
resultattr(i) = CDbl(tempav(0))
resultmaxvalue(i) = CDbl(tempav(1))


'�������
If i > 0 Then
If resultmaxvalue(i) < resultmaxvalue(i - 1) Then
Exit For
End If
End If

'debug
'If i = 1 Then
'Exit For
'End If

'GoTo test
'test:

Next i

'��̲׵��G�L�X��
'For i = 0 To UBound(resultattr)
'List1.AddItem resultattr(i)
'List1.AddItem resultmaxvalue(i)
'List1.AddItem ""
'Next i

'GoTo forend
'forend:


'�L�X���G����
'Dim ttresultattr(7) As Double
'Dim ttresultmaxvalue(7) As Double
'For i = 0 To 7
'ttresultattr(i) = -1
'ttresultmaxvalue(i) = -1
'Next i
'ttresultattr(0) = 5
'ttresultattr(1) = 8
'ttresultattr(2) = 3
'ttresultmaxvalue(0) = 11
'ttresultmaxvalue(1) = 12
'ttresultmaxvalue(2) = 13
'�L�X���G
Dim resultback As String
resultback = printfor(resultattr, resultmaxvalue)
List1.AddItem resultback


'�ʺA�}�C����
'Dim testnum() As Double
'For i = 0 To 3
'ReDim testnum(i)
'List1.AddItem UBound(testnum)
'Next i

End Sub
Static Function printfor(ByRef tresultattr() As Double, ByRef tresultmaxvalue() As Double)
Dim tempresultattr() As Double
Dim tempresultmaxvalue() As Double
Dim attrstring As String
Dim maxvaluetring As String
Dim eachattribute() As Double
tempresultattr() = tresultattr()
tempresultmaxvalue() = tresultmaxvalue()



For i = 0 To UBound(tempresultattr)
If tempresultattr(i) = -1 Then
Exit For
End If

ReDim eachattribute(i)
attrstring = "Attribute:"
maxvaluetring = "Correct rate: "

For j = 0 To i
eachattribute(j) = tempresultattr(j)
Next j


For j = 0 To UBound(eachattribute)
attrstring = attrstring + " " + CStr(eachattribute(j))
Next j
maxvaluetring = maxvaluetring + " " + CStr(tempresultmaxvalue(i))

List1.AddItem attrstring
List1.AddItem maxvaluetring
List1.AddItem ""
Next i
printfor = "End"
End Function

Private Sub generatefivefold_Click()
totalfivefoldindex() = fivefoldindex()
End Sub

Private Sub nerghbors_num_Change()
nei_number = CInt(nerghbors_num.Text)
End Sub

Private Sub read_Click()
List1.Clear
Dim datacounter As Integer
Dim temp() As String
'attr9array(0) = "CYT"
'attr9array(1) = "NUC"
'attr9array(2) = "MIT"
'attr9array(3) = "ME3"
'attr9array(4) = "ME2"
'attr9array(5) = "ME1"
'attr9array(6) = "EXC"
'attr9array(7) = "VAC"
'attr9array(8) = "POX"
'attr9array(9) = "ERL"

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
