VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "omL k=3"
   LinkTopic       =   "Form1"
   StartUpPosition =   3  '�t�ιw�]��
   ClientLeft      =   120
   ClientTop       =   465
   ClientHeight    =   7485
   ClientWidth     =   17580
   ScaleHeight     =   7485
   ScaleWidth      =   17580
   
   Begin VB.PictureBox Picture12 
      Left            =   14760
      TabIndex        =   23
      Top             =   4440
      ScaleHeight     =   2355
      ScaleWidth      =   2715
      Height          =   2415
      Width           =   2775
   End
   Begin VB.PictureBox Picture11 
      Left            =   11880
      TabIndex        =   22
      Top             =   4440
      ScaleHeight     =   2355
      ScaleWidth      =   2715
      Height          =   2415
      Width           =   2775
   End
   Begin VB.PictureBox Picture10 
      Left            =   8880
      TabIndex        =   21
      Top             =   4440
      ScaleHeight     =   2355
      ScaleWidth      =   2715
      Height          =   2415
      Width           =   2775
   End
   Begin VB.PictureBox Picture9 
      Left            =   6000
      TabIndex        =   20
      Top             =   4440
      ScaleHeight     =   2355
      ScaleWidth      =   2715
      Height          =   2415
      Width           =   2775
   End
   Begin VB.PictureBox Picture8 
      Left            =   3000
      TabIndex        =   19
      Top             =   4440
      ScaleHeight     =   2355
      ScaleWidth      =   2715
      Height          =   2415
      Width           =   2775
   End
   Begin VB.PictureBox Picture7 
      Left            =   120
      TabIndex        =   18
      Top             =   4440
      ScaleHeight     =   2355
      ScaleWidth      =   2715
      Height          =   2415
      Width           =   2775
   End
   Begin VB.PictureBox Picture6 
      Height          =   2415
      Left            =   14760
      ScaleHeight     =   2355
      ScaleWidth      =   2715
      TabIndex        =   17
      Top             =   1200
      Width           =   2775
   End
   Begin VB.PictureBox Picture5 
      Height          =   2415
      Left            =   11880
      ScaleHeight     =   2355
      ScaleWidth      =   2715
      TabIndex        =   16
      Top             =   1200
      Width           =   2775
   End
   Begin VB.PictureBox Picture4 
      Height          =   2415
      Left            =   8880
      ScaleHeight     =   2355
      ScaleWidth      =   2715
      TabIndex        =   15
      Top             =   1200
      Width           =   2775
   End
   Begin VB.PictureBox Picture3 
      Height          =   2415
      Left            =   6000
      ScaleHeight     =   2355
      ScaleWidth      =   2715
      TabIndex        =   14
      Top             =   1200
      Width           =   2775
   End
   Begin VB.PictureBox Picture2 
      Height          =   2415
      Left            =   3120
      ScaleHeight     =   2355
      ScaleWidth      =   2715
      TabIndex        =   13
      Top             =   1200
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      Height          =   2415
      Left            =   120
      ScaleHeight     =   2355
      ScaleWidth      =   2715
      TabIndex        =   24
      Top             =   1200
      Width           =   2775
   End


   Begin VB.CommandButton Command13 
      Caption         =   "omL k=6"
      Height          =   375
      Left            =   15360
      TabIndex        =   12
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command12 
      Caption         =   "omL k=5"
      Height          =   375
      Left            =   12480
      TabIndex        =   11
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command11 
      Caption         =   "om k=6"
      Height          =   375
      Left            =   9840
      TabIndex        =   10
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command10 
      Caption         =   "omL k=4"
      Height          =   375
      Left            =   15240
      TabIndex        =   9
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Command9 
      Caption         =   "omL k=3"
      Height          =   375
      Left            =   12480
      TabIndex        =   8
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "om k=4"
      Height          =   375
      Left            =   9720
      TabIndex        =   7
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "om k=5"
      Height          =   375
      Left            =   6720
      TabIndex        =   6
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "imU k=6"
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "imU k=5"
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "om k=3"
      Height          =   375
      Left            =   6720
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "imU k=4"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "imU k=3"
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Read Data"
      Height          =   375
      Left            =   8160
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End

Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1

Dim Data(336, 9) As String
Dim distance(336, 335) As Double
Dim Index(336, 335) As Double

Dim total(336) As Double

Dim weight(336, 3) As Double
Dim weight4(336, 4) As Double
Dim weight5(336, 5) As Double
Dim weight6(336, 6) As Double

Dim probab(336) As Double
Dim classify(336) As String

Dim index2(336) As Double

Private Sub Command1_Click()
 Dim str As String
 Open "C:\Users\user\Desktop\hw\hw2\ecoli.txt" For Input As #1
 Do Until EOF(1)
  For i = 1 To 336
    Line Input #1, str
      For j = 1 To 9
      str_splitText = Split(str, "  ")
      Data(i, j) = str_splitText(j - 1)
       Next
  Next
  Loop
  Close #1
'----continuous attributes are 2,3,6,7,8 th
'----calculate distance between ith instanse and other instances
'----list ith instance's neighbors' index from low to high
  For i = 1 To 336
    For j = 1 To 336
        Select Case j
        
           Case Is < i
           distance(i, j) = dist(i, j)
           Index(i, j) = j
           Case Is = i
                If j <> 336 Then
                    distance(i, j) = dist(i, j + 1)
                    Index(i, j) = j + 1
                    j = j + 1
                End If
                
           Case Is > i
           distance(i, j - 1) = dist(i, j)
           Index(i, j - 1) = j
        End Select
    Next j
Next i

'----sorting distance ,index matrix according to distance value (from low to high)

For i = 1 To 336
    For j = 1 To 334
        For k = 334 To j Step -1
            Dim temp As Double
            Dim tempind As Double
            If distance(i, k + 1) < distance(i, k) Then
                temp = distance(i, k)
                distance(i, k) = distance(i, k + 1)
                distance(i, k + 1) = temp
                tempind = Index(i, k)
                Index(i, k) = Index(i, k + 1)
                Index(i, k + 1) = tempind
            End If
        Next k
    Next j
Next i
      
MsgBox "Successfully Read ecoli txt! (336 instances,7 attributes,1 class)"
   
End Sub


Private Sub Command2_Click()
'----k=3,imU (35 instances)----
'----calculate the total of inverse of distance ----
For i = 1 To 336
    total(i) = 0
    For j = 1 To 3
        If distance(i, j) = 0 Then '----if distance=0,total approximately add 100 (cause we can't divide 0).I have tested,only few instances have 0 distance with other instances
            total(i) = total(i) + 100
        Else
            total(i) = total(i) + 1 / distance(i, j)
        End If
    Next j
Next i

'----calculate weight - inverse of distance is divided by total
For i = 1 To 336
    For j = 1 To 3
        weight(i, j) = 0
        If distance(i, j) = 0 Then
           weight(i, j) = 100 / total(i) '----if distance=0,inverse of distance is set to 100 (cause we can't divide 0).I have tested,only few instances have 0 distance with other instances
        Else
           weight(i, j) = 1 / distance(i, j) / total(i)
        End If
        If Data(Index(i, j), 9) <> "imU" Then '----remove neighbor's weight if its class value isn't our target class value
            weight(i, j) = 0
        End If
    Next j
Next i

For i = 1 To 336
    probab(i) = 0
    For j = 1 To 3
        probab(i) = probab(i) + weight(i, j) '----calculate probability of having the target class value
    Next j
Next i

For i = 1 To 336
    classify(i) = 0
    If 301 / 336 * (probab(i)) >= 35 / 336 * (1 - probab(i)) Then '----class imbalance �B�z,�վ������k���W�h��---
       classify(i) = "imU"
    Else
       classify(i) = "others"
    End If
Next i

Dim accuracy As Double
Dim recall As Double
Dim precision As Double

For i = 1 To 336
    If Data(i, 9) = "imU" Then '----calculate recall
        If classify(i) = "imU" Then
            recall = recall + 1
        End If
    End If
Next i

Dim count As Integer

For i = 1 To 336
    If classify(i) = "imU" Then
        count = count + 1
        If Data(i, 9) = "imU" Then
            precision = precision + 1 '----calculate precision
        End If
    End If
Next i

recall = Int(recall / 35 * 1000 + 0.5) / 1000 '----round to 3th digit after decimal point
precision = Int(precision / count * 1000 + 0.5) / 1000
'---------------------------------------------------------------

Picture1.Line (300, 300)-(2150, 2150), , B

Picture1.CurrentX = 2300
Picture1.CurrentY = 2200
Picture1.Print "FP%"

Picture1.CurrentX = 200
Picture1.CurrentY = 150
Picture1.Print "TP%"

 For i = 1 To 336
    index2(i) = i
 Next i

'----sort neighbor's index and probability according to probability value (from High to Low)
 For j = 1 To 335
     For k = 335 To j Step -1
            If probab(k + 1) > probab(k) Then
                temp = index2(k)
                index2(k) = index2(k + 1)
                index2(k + 1) = temp
                temp = probab(k)
                probab(k) = probab(k + 1)
                probab(k + 1) = temp
            End If
    Next k
 Next j
    
Dim X As Double
Dim Y As Double
Dim AUC As Double

X = 300
Y = 2150
'----Draw R.O.C curve (if predict right:draw the line upward,else: draw the line from left to right)
'----Calculate AUC (each time right turn, calculate the rectangular's area value(current height times interval we go right each time))

For i = 1 To 336
    If Data(index2(i), 9) = "imU" Then
        Y = Y - 1850 / 35
        Picture1.Line (X, Y + 1850 / 35)-(X, Y)
    Else
        X = X + 1850 / 301
        Picture1.Line (X - 1850 / 301, Y)-(X, Y)
        AUC = AUC + (2150 - Y) * 1850 / 301
    End If
Next i

AUC = AUC / 1850 / 1850
AUC = Int(AUC * 1000 + 0.5) / 1000

'----Print Recall, Precision,AUC on the picture
Picture1.CurrentX = 100
Picture1.CurrentY = 2200
Picture1.Print "Recall=" & recall & " Precision=" & precision

Picture1.CurrentX = 900
Picture1.CurrentY = 1000
Picture1.Print "AUC=" & AUC & vbCrLf
End Sub


Private Sub Command3_Click()
'----k=4,imU (35 instances)--------
For i = 1 To 336
    total(i) = 0
    For j = 1 To 4
        If distance(i, j) = 0 Then
            total(i) = total(i) + 100
        Else
            total(i) = total(i) + 1 / distance(i, j)
        End If
    Next j
Next i

For i = 1 To 336
    For j = 1 To 4
        weight4(i, j) = 0
        If distance(i, j) = 0 Then
           weight4(i, j) = 100 / total(i)
        Else
           weight4(i, j) = 1 / distance(i, j) / total(i)
        End If
        If Data(Index(i, j), 9) <> "imU" Then
            weight4(i, j) = 0
        End If
    Next j
Next i

For i = 1 To 336
    probab(i) = 0
    For j = 1 To 4
        probab(i) = probab(i) + weight4(i, j)
    Next j
Next i

For i = 1 To 336
    classify(i) = 0
    If 301 / 336 * (probab(i)) >= 35 / 336 * (1 - probab(i)) Then 'class imbalance �B�z,�վ������k���W�h��---
       classify(i) = "imU"
    Else
       classify(i) = "others"
    End If
Next i

Dim accuracy As Double
Dim recall As Double
Dim precision As Double

For i = 1 To 336
    If Data(i, 9) = "imU" Then
        If classify(i) = "imU" Then
            recall = recall + 1
        End If
    End If
Next i

Dim count As Integer

For i = 1 To 336
    If classify(i) = "imU" Then
        count = count + 1
        If Data(i, 9) = "imU" Then
            precision = precision + 1
        End If
    End If
Next i

recall = Int(recall / 35 * 1000 + 0.5) / 1000
precision = Int(precision / count * 1000 + 0.5) / 1000

'---------------------------------------------------------------

Picture2.Line (300, 300)-(2150, 2150), , B

Picture2.CurrentX = 2300
Picture2.CurrentY = 2200
Picture2.Print "FP%"

Picture2.CurrentX = 200
Picture2.CurrentY = 150
Picture2.Print "TP%"

 For i = 1 To 336
    index2(i) = i
 Next i

 For j = 1 To 335
     For k = 335 To j Step -1
            If probab(k + 1) > probab(k) Then
                temp = index2(k)
                index2(k) = index2(k + 1)
                index2(k + 1) = temp
                temp = probab(k)
                probab(k) = probab(k + 1)
                probab(k + 1) = temp
            End If
    Next k
 Next j
    
Dim X As Double
Dim Y As Double
Dim AUC As Double

X = 300
Y = 2150
For i = 1 To 336
    If Data(index2(i), 9) = "imU" Then
        Y = Y - 1850 / 35
        Picture2.Line (X, Y + 1850 / 35)-(X, Y)
    Else
        X = X + 1850 / 301
        Picture2.Line (X - 1850 / 301, Y)-(X, Y)
        AUC = AUC + (2150 - Y) * 1850 / 301
    End If
Next i

AUC = AUC / 1850 / 1850
AUC = Int(AUC * 1000 + 0.5) / 1000


Picture2.CurrentX = 100
Picture2.CurrentY = 2200
Picture2.Print "Recall=" & recall & " Precision=" & precision

Picture2.CurrentX = 900
Picture2.CurrentY = 1000
Picture2.Print "AUC=" & AUC & vbCrLf
End Sub


Private Sub Command5_Click()
'----k=5,imU (35 instances)----
For i = 1 To 336
    total(i) = 0
    For j = 1 To 5
        If distance(i, j) = 0 Then
            total(i) = total(i) + 100
        Else
            total(i) = total(i) + 1 / distance(i, j)
        End If
    Next j
Next i

For i = 1 To 336
    For j = 1 To 5
        weight5(i, j) = 0
        If distance(i, j) = 0 Then
           weight5(i, j) = 100 / total(i)
        Else
           weight5(i, j) = 1 / distance(i, j) / total(i)
        End If
        If Data(Index(i, j), 9) <> "imU" Then
            weight5(i, j) = 0
        End If
    Next j
Next i

For i = 1 To 336
    probab(i) = 0
    For j = 1 To 5
        probab(i) = probab(i) + weight5(i, j)
    Next j
Next i

For i = 1 To 336
    classify(i) = 0
    If 301 / 336 * (probab(i)) >= 35 / 336 * (1 - probab(i)) Then 'class imbalance �B�z,�վ������k���W�h��---
       classify(i) = "imU"
    Else
       classify(i) = "others"
    End If
Next i

Dim accuracy As Double
Dim recall As Double
Dim precision As Double

For i = 1 To 336
    If Data(i, 9) = "imU" Then
        If classify(i) = "imU" Then
            recall = recall + 1
        End If
    End If
Next i

Dim count As Integer

For i = 1 To 336
    If classify(i) = "imU" Then
        count = count + 1
        If Data(i, 9) = "imU" Then
            precision = precision + 1
        End If
    End If
Next i

recall = Int(recall / 35 * 1000 + 0.5) / 1000
precision = Int(precision / count * 1000 + 0.5) / 1000

'---------------------------------------------------------------

Picture7.Line (300, 300)-(2150, 2150), , B

Picture7.CurrentX = 2300
Picture7.CurrentY = 2200
Picture7.Print "FP%"

Picture7.CurrentX = 200
Picture7.CurrentY = 150
Picture7.Print "TP%"

 For i = 1 To 336
    index2(i) = i
 Next i

 For j = 1 To 335
     For k = 335 To j Step -1
            If probab(k + 1) > probab(k) Then
                temp = index2(k)
                index2(k) = index2(k + 1)
                index2(k + 1) = temp
                temp = probab(k)
                probab(k) = probab(k + 1)
                probab(k + 1) = temp
            End If
    Next k
 Next j
    
Dim X As Double
Dim Y As Double
Dim AUC As Double

X = 300
Y = 2150
For i = 1 To 336
    If Data(index2(i), 9) = "imU" Then
        Y = Y - 1850 / 35
        Picture7.Line (X, Y + 1850 / 35)-(X, Y)
    Else
        X = X + 1850 / 301
        Picture7.Line (X - 1850 / 301, Y)-(X, Y)
        AUC = AUC + (2150 - Y) * 1850 / 301
    End If
Next i

AUC = AUC / 1850 / 1850
AUC = Int(AUC * 1000 + 0.5) / 1000


Picture7.CurrentX = 100
Picture7.CurrentY = 2200
Picture7.Print "Recall=" & recall & " Precision=" & precision

Picture7.CurrentX = 900
Picture7.CurrentY = 1000
Picture7.Print "AUC=" & AUC & vbCrLf

End Sub

Private Sub Command6_Click()
'----k=6,imU (35 instances)----
For i = 1 To 336
    total(i) = 0
    For j = 1 To 6
        If distance(i, j) = 0 Then
            total(i) = total(i) + 100
        Else
            total(i) = total(i) + 1 / distance(i, j)
        End If
    Next j
Next i

For i = 1 To 336
    For j = 1 To 6
        weight6(i, j) = 0
        If distance(i, j) = 0 Then
           weight6(i, j) = 100 / total(i)
        Else
           weight6(i, j) = 1 / distance(i, j) / total(i)
        End If
        If Data(Index(i, j), 9) <> "imU" Then
            weight6(i, j) = 0
        End If
    Next j
Next i

For i = 1 To 336
    probab(i) = 0
    For j = 1 To 6
        probab(i) = probab(i) + weight6(i, j)
    Next j
Next i

For i = 1 To 336
    classify(i) = 0
    If 301 / 336 * (probab(i)) >= 35 / 336 * (1 - probab(i)) Then 'class imbalance �B�z,�վ������k���W�h��---
       classify(i) = "imU"
    Else
       classify(i) = "others"
    End If
Next i

Dim accuracy As Double
Dim recall As Double
Dim precision As Double

For i = 1 To 336
    If Data(i, 9) = "imU" Then
        If classify(i) = "imU" Then
            recall = recall + 1
        End If
    End If
Next i

Dim count As Integer

For i = 1 To 336
    If classify(i) = "imU" Then
        count = count + 1
        If Data(i, 9) = "imU" Then
            precision = precision + 1
        End If
    End If
Next i

recall = Int(recall / 35 * 1000 + 0.5) / 1000
precision = Int(precision / count * 1000 + 0.5) / 1000

'---------------------------------------------------------------

Picture8.Line (300, 300)-(2150, 2150), , B

Picture8.CurrentX = 2300
Picture8.CurrentY = 2200
Picture8.Print "FP%"

Picture8.CurrentX = 200
Picture8.CurrentY = 150
Picture8.Print "TP%"

 For i = 1 To 336
    index2(i) = i
 Next i

 For j = 1 To 335
     For k = 335 To j Step -1
            If probab(k + 1) > probab(k) Then
                temp = index2(k)
                index2(k) = index2(k + 1)
                index2(k + 1) = temp
                temp = probab(k)
                probab(k) = probab(k + 1)
                probab(k + 1) = temp
            End If
    Next k
 Next j
    
Dim X As Double
Dim Y As Double
Dim AUC As Double

X = 300
Y = 2150
For i = 1 To 336
    If Data(index2(i), 9) = "imU" Then
        Y = Y - 1850 / 35
        Picture8.Line (X, Y + 1850 / 35)-(X, Y)
    Else
        X = X + 1850 / 301
        Picture8.Line (X - 1850 / 301, Y)-(X, Y)
        AUC = AUC + (2150 - Y) * 1850 / 301
    End If
Next i

AUC = AUC / 1850 / 1850
AUC = Int(AUC * 1000 + 0.5) / 1000


Picture8.CurrentX = 100
Picture8.CurrentY = 2200
Picture8.Print "Recall=" & recall & " Precision=" & precision

Picture8.CurrentX = 900
Picture8.CurrentY = 1000
Picture8.Print "AUC=" & AUC & vbCrLf

End Sub

Private Sub Command4_Click()
'----k=3, om (20 instances)----
For i = 1 To 336
    total(i) = 0
    For j = 1 To 3
        If distance(i, j) = 0 Then
            total(i) = total(i) + 100
        Else
            total(i) = total(i) + 1 / distance(i, j)
        End If
    Next j
Next i

For i = 1 To 336
    For j = 1 To 3
        weight(i, j) = 0
        If distance(i, j) = 0 Then
           weight(i, j) = 100 / total(i)
        Else
           weight(i, j) = 1 / distance(i, j) / total(i)
        End If
        If Data(Index(i, j), 9) <> " om" Then
            weight(i, j) = 0
        End If
    Next j
Next i

For i = 1 To 336
    probab(i) = 0
    For j = 1 To 3
        probab(i) = probab(i) + weight(i, j)
    Next j
Next i

For i = 1 To 336
    classify(i) = 0
    If 316 / 336 * (probab(i)) >= 20 / 336 * (1 - probab(i)) Then 'class imbalance �B�z,�վ������k���W�h��---
       classify(i) = " om"
    Else
       classify(i) = "others"
    End If
Next i

Dim accuracy As Double
Dim recall As Double
Dim precision As Double

For i = 1 To 336
    If Data(i, 9) = " om" Then
        If classify(i) = " om" Then
            recall = recall + 1
        End If
    End If
Next i

Dim count As Integer

For i = 1 To 336
    If classify(i) = " om" Then
        count = count + 1
        If Data(i, 9) = " om" Then
            precision = precision + 1
        End If
    End If
Next i

recall = Int(recall / 20 * 1000 + 0.5) / 1000
precision = Int(precision / count * 1000 + 0.5) / 1000

'---------------------------------------------------------------

Picture3.Line (300, 300)-(2150, 2150), , B

Picture3.CurrentX = 2300
Picture3.CurrentY = 2200
Picture3.Print "FP%"

Picture3.CurrentX = 200
Picture3.CurrentY = 150
Picture3.Print "TP%"

 For i = 1 To 336
    index2(i) = i
 Next i

 For j = 1 To 335
     For k = 335 To j Step -1
            If probab(k + 1) > probab(k) Then
                temp = index2(k)
                index2(k) = index2(k + 1)
                index2(k + 1) = temp
                temp = probab(k)
                probab(k) = probab(k + 1)
                probab(k + 1) = temp
            End If
    Next k
 Next j
    
Dim X As Double
Dim Y As Double
Dim AUC As Double

X = 300
Y = 2150
For i = 1 To 336
    If Data(index2(i), 9) = " om" Then
        Y = Y - 1850 / 20
        Picture3.Line (X, Y + 1850 / 20)-(X, Y)
    Else
        X = X + 1850 / 316
        Picture3.Line (X - 1850 / 316, Y)-(X, Y)
        AUC = AUC + (2150 - Y) * 1850 / 316
    End If
Next i

AUC = AUC / 1850 / 1850
AUC = Int(AUC * 1000 + 0.5) / 1000


Picture3.CurrentX = 100
Picture3.CurrentY = 2200
Picture3.Print "Recall=" & recall & " Precision=" & precision

Picture3.CurrentX = 900
Picture3.CurrentY = 1000
Picture3.Print "AUC=" & AUC & vbCrLf


End Sub

Private Sub Command8_Click()
'----k=4, om (20 instances)
For i = 1 To 336
    total(i) = 0
    For j = 1 To 4
        If distance(i, j) = 0 Then
            total(i) = total(i) + 100
        Else
            total(i) = total(i) + 1 / distance(i, j)
        End If
    Next j
Next i

For i = 1 To 336
    For j = 1 To 4
        weight4(i, j) = 0
        If distance(i, j) = 0 Then
           weight4(i, j) = 100 / total(i)
        Else
           weight4(i, j) = 1 / distance(i, j) / total(i)
        End If
        If Data(Index(i, j), 9) <> " om" Then
            weight4(i, j) = 0
        End If
    Next j
Next i

For i = 1 To 336
    probab(i) = 0
    For j = 1 To 4
        probab(i) = probab(i) + weight4(i, j)
    Next j
Next i

For i = 1 To 336
    classify(i) = 0
    If 316 / 336 * (probab(i)) >= 20 / 336 * (1 - probab(i)) Then 'class imbalance �B�z,�վ������k���W�h��---
       classify(i) = " om"
    Else
       classify(i) = "others"
    End If
Next i

Dim accuracy As Double
Dim recall As Double
Dim precision As Double

For i = 1 To 336
    If Data(i, 9) = " om" Then
        If classify(i) = " om" Then
            recall = recall + 1
        End If
    End If
Next i

Dim count As Integer

For i = 1 To 336
    If classify(i) = " om" Then
        count = count + 1
        If Data(i, 9) = " om" Then
            precision = precision + 1
        End If
    End If
Next i

recall = Int(recall / 20 * 1000 + 0.5) / 1000
precision = Int(precision / count * 1000 + 0.5) / 1000

'---------------------------------------------------------------

Picture4.Line (300, 300)-(2150, 2150), , B

Picture4.CurrentX = 2300
Picture4.CurrentY = 2200
Picture4.Print "FP%"

Picture4.CurrentX = 200
Picture4.CurrentY = 150
Picture4.Print "TP%"

 For i = 1 To 336
    index2(i) = i
 Next i

 For j = 1 To 335
     For k = 335 To j Step -1
            If probab(k + 1) > probab(k) Then
                temp = index2(k)
                index2(k) = index2(k + 1)
                index2(k + 1) = temp
                temp = probab(k)
                probab(k) = probab(k + 1)
                probab(k + 1) = temp
            End If
    Next k
 Next j
    
Dim X As Double
Dim Y As Double
Dim AUC As Double

X = 300
Y = 2150
For i = 1 To 336
    If Data(index2(i), 9) = " om" Then
        Y = Y - 1850 / 20
        Picture4.Line (X, Y + 1850 / 20)-(X, Y)
    Else
        X = X + 1850 / 316
        Picture4.Line (X - 1850 / 316, Y)-(X, Y)
        AUC = AUC + (2150 - Y) * 1850 / 316
    End If
Next i

AUC = AUC / 1850 / 1850
AUC = Int(AUC * 1000 + 0.5) / 1000


Picture4.CurrentX = 100
Picture4.CurrentY = 2200
Picture4.Print "Recall=" & recall & " Precision=" & precision

Picture4.CurrentX = 900
Picture4.CurrentY = 1000
Picture4.Print "AUC=" & AUC & vbCrLf

End Sub

Private Sub Command7_Click()
'----k=5, om (20 instances)
For i = 1 To 336
    total(i) = 0
    For j = 1 To 5
        If distance(i, j) = 0 Then
            total(i) = total(i) + 100
        Else
            total(i) = total(i) + 1 / distance(i, j)
        End If
    Next j
Next i


For i = 1 To 336
    For j = 1 To 5
        weight5(i, j) = 0
        If distance(i, j) = 0 Then
           weight5(i, j) = 100 / total(i)
        Else
           weight5(i, j) = 1 / distance(i, j) / total(i)
        End If
        If Data(Index(i, j), 9) <> " om" Then
            weight5(i, j) = 0
        End If
    Next j
Next i

For i = 1 To 336
    probab(i) = 0
    For j = 1 To 5
        probab(i) = probab(i) + weight5(i, j)
    Next j
Next i

For i = 1 To 336
    classify(i) = 0
    If 316 / 336 * (probab(i)) >= 20 / 336 * (1 - probab(i)) Then 'class imbalance �B�z,�վ������k���W�h��---
       classify(i) = " om"
    Else
       classify(i) = "others"
    End If
Next i

Dim accuracy As Double
Dim recall As Double
Dim precision As Double

For i = 1 To 336
    If Data(i, 9) = " om" Then
        If classify(i) = " om" Then
            recall = recall + 1
        End If
    End If
Next i

Dim count As Integer

For i = 1 To 336
    If classify(i) = " om" Then
        count = count + 1
        If Data(i, 9) = " om" Then
            precision = precision + 1
        End If
    End If
Next i

recall = Int(recall / 20 * 1000 + 0.5) / 1000
precision = Int(precision / count * 1000 + 0.5) / 1000

'---------------------------------------------------------------

Picture9.Line (300, 300)-(2150, 2150), , B

Picture9.CurrentX = 2300
Picture9.CurrentY = 2200
Picture9.Print "FP%"

Picture9.CurrentX = 200
Picture9.CurrentY = 150
Picture9.Print "TP%"

 For i = 1 To 336
    index2(i) = i
 Next i

 For j = 1 To 335
     For k = 335 To j Step -1
            If probab(k + 1) > probab(k) Then
                temp = index2(k)
                index2(k) = index2(k + 1)
                index2(k + 1) = temp
                temp = probab(k)
                probab(k) = probab(k + 1)
                probab(k + 1) = temp
            End If
    Next k
 Next j
    
Dim X As Double
Dim Y As Double
Dim AUC As Double

X = 300
Y = 2150
For i = 1 To 336
    If Data(index2(i), 9) = " om" Then
        Y = Y - 1850 / 20
        Picture9.Line (X, Y + 1850 / 20)-(X, Y)
    Else
        X = X + 1850 / 316
        Picture9.Line (X - 1850 / 316, Y)-(X, Y)
        AUC = AUC + (2150 - Y) * 1850 / 316
    End If
Next i

AUC = AUC / 1850 / 1850
AUC = Int(AUC * 1000 + 0.5) / 1000


Picture9.CurrentX = 100
Picture9.CurrentY = 2200
Picture9.Print "Recall=" & recall & " Precision=" & precision

Picture9.CurrentX = 900
Picture9.CurrentY = 1000
Picture9.Print "AUC=" & AUC & vbCrLf

End Sub

Private Sub Command11_Click()
'----k=6, om (20 instances)
For i = 1 To 336
    total(i) = 0
    For j = 1 To 6
        If distance(i, j) = 0 Then
            total(i) = total(i) + 100
        Else
            total(i) = total(i) + 1 / distance(i, j)
        End If
    Next j
Next i


For i = 1 To 336
    For j = 1 To 6
        weight6(i, j) = 0
        If distance(i, j) = 0 Then
           weight6(i, j) = 100 / total(i)
        Else
           weight6(i, j) = 1 / distance(i, j) / total(i)
        End If
        If Data(Index(i, j), 9) <> " om" Then
            weight6(i, j) = 0
        End If
    Next j
Next i

For i = 1 To 336
    probab(i) = 0
    For j = 1 To 6
        probab(i) = probab(i) + weight6(i, j)
    Next j
Next i

For i = 1 To 336
    classify(i) = 0
    If 316 / 336 * (probab(i)) >= 20 / 336 * (1 - probab(i)) Then 'class imbalance �B�z,�վ������k���W�h��---
       classify(i) = " om"
    Else
       classify(i) = "others"
    End If
Next i

Dim accuracy As Double
Dim recall As Double
Dim precision As Double

For i = 1 To 336
    If Data(i, 9) = " om" Then
        If classify(i) = " om" Then
            recall = recall + 1
        End If
    End If
Next i

Dim count As Integer

For i = 1 To 336
    If classify(i) = " om" Then
        count = count + 1
        If Data(i, 9) = " om" Then
            precision = precision + 1
        End If
    End If
Next i

recall = Int(recall / 20 * 1000 + 0.5) / 1000
precision = Int(precision / count * 1000 + 0.5) / 1000

'---------------------------------------------------------------

Picture10.Line (300, 300)-(2150, 2150), , B

Picture10.CurrentX = 2300
Picture10.CurrentY = 2200
Picture10.Print "FP%"

Picture10.CurrentX = 200
Picture10.CurrentY = 150
Picture10.Print "TP%"

 For i = 1 To 336
    index2(i) = i
 Next i

 For j = 1 To 335
     For k = 335 To j Step -1
            If probab(k + 1) > probab(k) Then
                temp = index2(k)
                index2(k) = index2(k + 1)
                index2(k + 1) = temp
                temp = probab(k)
                probab(k) = probab(k + 1)
                probab(k + 1) = temp
            End If
    Next k
 Next j
    
Dim X As Double
Dim Y As Double
Dim AUC As Double

X = 300
Y = 2150
For i = 1 To 336
    If Data(index2(i), 9) = " om" Then
        Y = Y - 1850 / 20
        Picture10.Line (X, Y + 1850 / 20)-(X, Y)
    Else
        X = X + 1850 / 316
        Picture10.Line (X - 1850 / 316, Y)-(X, Y)
        AUC = AUC + (2150 - Y) * 1850 / 316
    End If
Next i

AUC = AUC / 1850 / 1850
AUC = Int(AUC * 1000 + 0.5) / 1000


Picture10.CurrentX = 100
Picture10.CurrentY = 2200
Picture10.Print "Recall=" & recall & " Precision=" & precision

Picture10.CurrentX = 900
Picture10.CurrentY = 1000
Picture10.Print "AUC=" & AUC & vbCrLf

End Sub

Private Sub Command9_Click()
'----k=3, omL (5 instances)
For i = 1 To 336
    total(i) = 0
    For j = 1 To 3
        If distance(i, j) = 0 Then
            total(i) = total(i) + 100
        Else
            total(i) = total(i) + 1 / distance(i, j)
        End If
    Next j
Next i

For i = 1 To 336
    For j = 1 To 3
        weight(i, j) = 0
        If distance(i, j) = 0 Then
           weight(i, j) = 100 / total(i)
        Else
           weight(i, j) = 1 / distance(i, j) / total(i)
        End If
        If Data(Index(i, j), 9) <> "omL" Then
            weight(i, j) = 0
        End If
    Next j
Next i

For i = 1 To 336
    probab(i) = 0
    For j = 1 To 3
        probab(i) = probab(i) + weight(i, j)
    Next j
Next i

For i = 1 To 336
    classify(i) = 0
    If 331 / 336 * (probab(i)) >= 5 / 336 * (1 - probab(i)) Then 'class imbalance �B�z,�վ������k���W�h��---
       classify(i) = "omL"
    Else
       classify(i) = "others"
    End If
Next i

Dim accuracy As Double
Dim recall As Double
Dim precision As Double

For i = 1 To 336
    If Data(i, 9) = "omL" Then
        If classify(i) = "omL" Then
            recall = recall + 1
        End If
    End If
Next i

Dim count As Integer

For i = 1 To 336
    If classify(i) = "omL" Then
        count = count + 1
        If Data(i, 9) = "omL" Then
            precision = precision + 1
        End If
    End If
Next i

recall = Int(recall / 5 * 1000 + 0.5) / 1000
precision = Int(precision / count * 1000 + 0.5) / 1000

'---------------------------------------------------------------

Picture5.Line (300, 300)-(2150, 2150), , B

Picture5.CurrentX = 2300
Picture5.CurrentY = 2200
Picture5.Print "FP%"

Picture5.CurrentX = 200
Picture5.CurrentY = 150
Picture5.Print "TP%"

 For i = 1 To 336
    index2(i) = i
 Next i

 For j = 1 To 335
     For k = 335 To j Step -1
            If probab(k + 1) > probab(k) Then
                temp = index2(k)
                index2(k) = index2(k + 1)
                index2(k + 1) = temp
                temp = probab(k)
                probab(k) = probab(k + 1)
                probab(k + 1) = temp
            End If
    Next k
 Next j
    
Dim X As Double
Dim Y As Double
Dim AUC As Double

X = 300
Y = 2150
For i = 1 To 336
    If Data(index2(i), 9) = "omL" Then
        Y = Y - 1850 / 5
        Picture5.Line (X, Y + 1850 / 5)-(X, Y)
    Else
        X = X + 1850 / 331
        Picture5.Line (X - 1850 / 331, Y)-(X, Y)
        AUC = AUC + (2150 - Y) * 1850 / 331
    End If
Next i

AUC = AUC / 1850 / 1850
AUC = Int(AUC * 1000 + 0.5) / 1000


Picture5.CurrentX = 100
Picture5.CurrentY = 2200
Picture5.Print "Recall=" & recall & " Precision=" & precision

Picture5.CurrentX = 900
Picture5.CurrentY = 1000
Picture5.Print "AUC=" & AUC & vbCrLf


End Sub

Private Sub Command10_Click()
'----k=4, omL (5 instances)
For i = 1 To 336
    total(i) = 0
    For j = 1 To 4
        If distance(i, j) = 0 Then
            total(i) = total(i) + 100
        Else
            total(i) = total(i) + 1 / distance(i, j)
        End If
    Next j
Next i

For i = 1 To 336
    For j = 1 To 4
        weight4(i, j) = 0
        If distance(i, j) = 0 Then
           weight4(i, j) = 100 / total(i)
        Else
           weight4(i, j) = 1 / distance(i, j) / total(i)
        End If
        If Data(Index(i, j), 9) <> "omL" Then
            weight4(i, j) = 0
        End If
    Next j
Next i

For i = 1 To 336
    probab(i) = 0
    For j = 1 To 4
        probab(i) = probab(i) + weight4(i, j)
    Next j
Next i

For i = 1 To 336
    classify(i) = 0
    If 331 / 336 * (probab(i)) >= 5 / 336 * (1 - probab(i)) Then 'class imbalance �B�z,�վ������k���W�h��---
       classify(i) = "omL"
    Else
       classify(i) = "others"
    End If
Next i

Dim accuracy As Double
Dim recall As Double
Dim precision As Double

For i = 1 To 336
    If Data(i, 9) = "omL" Then
        If classify(i) = "omL" Then
            recall = recall + 1
        End If
    End If
Next i

Dim count As Integer

For i = 1 To 336
    If classify(i) = "omL" Then
        count = count + 1
        If Data(i, 9) = "omL" Then
            precision = precision + 1
        End If
    End If
Next i

recall = Int(recall / 5 * 1000 + 0.5) / 1000
precision = Int(precision / count * 1000 + 0.5) / 1000

'---------------------------------------------------------------

Picture6.Line (300, 300)-(2150, 2150), , B

Picture6.CurrentX = 2300
Picture6.CurrentY = 2200
Picture6.Print "FP%"

Picture6.CurrentX = 200
Picture6.CurrentY = 150
Picture6.Print "TP%"

 For i = 1 To 336
    index2(i) = i
 Next i

 For j = 1 To 335
     For k = 335 To j Step -1
            If probab(k + 1) > probab(k) Then
                temp = index2(k)
                index2(k) = index2(k + 1)
                index2(k + 1) = temp
                temp = probab(k)
                probab(k) = probab(k + 1)
                probab(k + 1) = temp
            End If
    Next k
 Next j
    
Dim X As Double
Dim Y As Double
Dim AUC As Double

X = 300
Y = 2150
For i = 1 To 336
    If Data(index2(i), 9) = "omL" Then
        Y = Y - 1850 / 5
        Picture6.Line (X, Y + 1850 / 5)-(X, Y)
    Else
        X = X + 1850 / 331
        Picture6.Line (X - 1850 / 331, Y)-(X, Y)
        AUC = AUC + (2150 - Y) * 1850 / 331
    End If
Next i

AUC = AUC / 1850 / 1850
AUC = Int(AUC * 1000 + 0.5) / 1000


Picture6.CurrentX = 100
Picture6.CurrentY = 2200
Picture6.Print "Recall=" & recall & " Precision=" & precision

Picture6.CurrentX = 900
Picture6.CurrentY = 1000
Picture6.Print "AUC=" & AUC & vbCrLf

End Sub

Private Sub Command12_Click()
'----k=5, omL (5 instances)
For i = 1 To 336
    total(i) = 0
    For j = 1 To 5
        If distance(i, j) = 0 Then
            total(i) = total(i) + 100
        Else
            total(i) = total(i) + 1 / distance(i, j)
        End If
    Next j
Next i


For i = 1 To 336
    For j = 1 To 5
        weight5(i, j) = 0
        If distance(i, j) = 0 Then
           weight5(i, j) = 100 / total(i)
        Else
           weight5(i, j) = 1 / distance(i, j) / total(i)
        End If
        If Data(Index(i, j), 9) <> "omL" Then
            weight5(i, j) = 0
        End If
    Next j
Next i

For i = 1 To 336
    probab(i) = 0
    For j = 1 To 5
        probab(i) = probab(i) + weight5(i, j)
    Next j
Next i

For i = 1 To 336
    classify(i) = 0
    If 331 / 336 * (probab(i)) >= 5 / 336 * (1 - probab(i)) Then 'class imbalance �B�z,�վ������k���W�h��---
       classify(i) = "omL"
    Else
       classify(i) = "others"
    End If
Next i

Dim accuracy As Double
Dim recall As Double
Dim precision As Double

For i = 1 To 336
    If Data(i, 9) = "omL" Then
        If classify(i) = "omL" Then
            recall = recall + 1
        End If
    End If
Next i

Dim count As Integer

For i = 1 To 336
    If classify(i) = "omL" Then
        count = count + 1
        If Data(i, 9) = "omL" Then
            precision = precision + 1
        End If
    End If
Next i

recall = Int(recall / 5 * 1000 + 0.5) / 1000
precision = Int(precision / count * 1000 + 0.5) / 1000

'---------------------------------------------------------------

Picture11.Line (300, 300)-(2150, 2150), , B

Picture11.CurrentX = 2300
Picture11.CurrentY = 2200
Picture11.Print "FP%"

Picture11.CurrentX = 200
Picture11.CurrentY = 150
Picture11.Print "TP%"

 For i = 1 To 336
    index2(i) = i
 Next i

 For j = 1 To 335
     For k = 335 To j Step -1
            If probab(k + 1) > probab(k) Then
                temp = index2(k)
                index2(k) = index2(k + 1)
                index2(k + 1) = temp
                temp = probab(k)
                probab(k) = probab(k + 1)
                probab(k + 1) = temp
            End If
    Next k
 Next j
    
Dim X As Double
Dim Y As Double
Dim AUC As Double

X = 300
Y = 2150
For i = 1 To 336
    If Data(index2(i), 9) = "omL" Then
        Y = Y - 1850 / 5
        Picture11.Line (X, Y + 1850 / 5)-(X, Y)
    Else
        X = X + 1850 / 331
        Picture11.Line (X - 1850 / 331, Y)-(X, Y)
        AUC = AUC + (2150 - Y) * 1850 / 331
    End If
Next i

AUC = AUC / 1850 / 1850
AUC = Int(AUC * 1000 + 0.5) / 1000


Picture11.CurrentX = 100
Picture11.CurrentY = 2200
Picture11.Print "Recall=" & recall & " Precision=" & precision

Picture11.CurrentX = 900
Picture11.CurrentY = 1000
Picture11.Print "AUC=" & AUC & vbCrLf

End Sub

Private Sub Command13_Click()
'----k=6, omL (5 instances)
For i = 1 To 336
    total(i) = 0
    For j = 1 To 6
        If distance(i, j) = 0 Then
            total(i) = total(i) + 100
        Else
            total(i) = total(i) + 1 / distance(i, j)
        End If
    Next j
Next i


For i = 1 To 336
    For j = 1 To 6
        weight6(i, j) = 0
        If distance(i, j) = 0 Then
           weight6(i, j) = 100 / total(i)
        Else
           weight6(i, j) = 1 / distance(i, j) / total(i)
        End If
        If Data(Index(i, j), 9) <> "omL" Then
            weight6(i, j) = 0
        End If
    Next j
Next i

For i = 1 To 336
    probab(i) = 0
    For j = 1 To 6
        probab(i) = probab(i) + weight6(i, j)
    Next j
Next i

For i = 1 To 336
    classify(i) = 0
    If 331 / 336 * (probab(i)) >= 5 / 336 * (1 - probab(i)) Then 'class imbalance �B�z,�վ������k���W�h��---
       classify(i) = "omL"
    Else
       classify(i) = "others"
    End If
Next i

Dim accuracy As Double
Dim recall As Double
Dim precision As Double

For i = 1 To 336
    If Data(i, 9) = "omL" Then
        If classify(i) = "omL" Then
            recall = recall + 1
        End If
    End If
Next i

Dim count As Integer

For i = 1 To 336
    If classify(i) = "omL" Then
        count = count + 1
        If Data(i, 9) = "omL" Then
            precision = precision + 1
        End If
    End If
Next i

recall = Int(recall / 5 * 1000 + 0.5) / 1000
precision = Int(precision / count * 1000 + 0.5) / 1000

'---------------------------------------------------------------

Picture12.Line (300, 300)-(2150, 2150), , B

Picture12.CurrentX = 2300
Picture12.CurrentY = 2200
Picture12.Print "FP%"

Picture12.CurrentX = 200
Picture12.CurrentY = 150
Picture12.Print "TP%"

 For i = 1 To 336
    index2(i) = i
 Next i

 For j = 1 To 335
     For k = 335 To j Step -1
            If probab(k + 1) > probab(k) Then
                temp = index2(k)
                index2(k) = index2(k + 1)
                index2(k + 1) = temp
                temp = probab(k)
                probab(k) = probab(k + 1)
                probab(k + 1) = temp
            End If
    Next k
 Next j
    
Dim X As Double
Dim Y As Double
Dim AUC As Double

X = 300
Y = 2150
For i = 1 To 336
    If Data(index2(i), 9) = "omL" Then
        Y = Y - 1850 / 5
        Picture12.Line (X, Y + 1850 / 5)-(X, Y)
    Else
        X = X + 1850 / 331
        Picture12.Line (X - 1850 / 331, Y)-(X, Y)
        AUC = AUC + (2150 - Y) * 1850 / 331
    End If
Next i

AUC = AUC / 1850 / 1850
AUC = Int(AUC * 1000 + 0.5) / 1000


Picture12.CurrentX = 100
Picture12.CurrentY = 2200
Picture12.Print "Recall=" & recall & " Precision=" & precision

Picture12.CurrentX = 900
Picture12.CurrentY = 1000
Picture12.Print "AUC=" & AUC & vbCrLf

End Sub

Private Function dist(i, j) As Double
    dist = (Val(Data(i, 2)) - Val(Data(j, 2))) * (Val(Data(i, 2)) - Val(Data(j, 2))) + (Val(Data(i, 3)) - Val(Data(j, 3))) * (Val(Data(i, 3)) - Val(Data(j, 3))) + (Val(Data(i, 6)) - Val(Data(j, 6))) * (Val(Data(i, 6)) - Val(Data(j, 6))) + (Val(Data(i, 7)) - Val(Data(j, 7))) * (Val(Data(i, 7)) - Val(Data(j, 7))) + (Val(Data(i, 8)) - Val(Data(j, 8))) * (Val(Data(i, 8)) - Val(Data(j, 8)))
End Function
