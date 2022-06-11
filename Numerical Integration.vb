
Option Base 0

Public Function dy1dt(y1 As Double, y2 As Double, y3 As Double, y4 As Double) As Double
    dy1dt = y2
End Function

Public Function dy2dt(y1 As Double, y2 As Double, y3 As Double, y4 As Double) As Double
    dy2dt = (-15600# * (2# * ((Abs(y1)) ^ -0.74) + 3# * ((Abs(y3 - y1)) ^ -0.74))) / 7# * y1 + _
            (3# * 15600# * ((Abs(y3 - y1)) ^ -0.74)) / 7# * y3
    
End Function

Public Function dy3dt(y1 As Double, y2 As Double, y3 As Double, y4 As Double) As Double
    dy3dt = y4
End Function

Public Function dy4dt(y1 As Double, y2 As Double, y3 As Double, y4 As Double) As Double
    dy4dt = (-15600# * (-((Abs(y1)) ^ -0.74) - 5# * ((Abs(y3 - y1)) ^ -0.74))) / 7# * y1 - _
            (5# * 15600# * ((Abs(y3 - y1)) ^ -0.74)) / 7# * y3
    
End Function

Public Sub DoEuler()
    Dim y(1 To 4) As Double
    Dim t(1 To 4) As Double
    Dim n As Integer
    Dim dt As Double
    Dim time As Double
        
    y(1) = 29.8613
    y(2) = 0
    y(3) = 30.0616
    y(4) = 0
    
    n = 5000
    dt = 5 / n
    time = 0
    
    For i = 1 To n
        t(1) = y(1) + dt * dy1dt(y(1), y(2), y(3), y(4))
        t(2) = y(2) + dt * dy2dt(y(1), y(2), y(3), y(4))
        t(3) = y(3) + dt * dy3dt(y(1), y(2), y(3), y(4))
        t(4) = y(4) + dt * dy4dt(y(1), y(2), y(3), y(4))
        
        y(1) = t(1)
        y(2) = t(2)
        y(3) = t(3)
        y(4) = t(4)
    
        time = time + dt
        ActiveSheet.Cells(i + 1, 1) = time
        ActiveSheet.Cells(i + 1, 2) = y(1)
        ActiveSheet.Cells(i + 1, 3) = y(2)
        ActiveSheet.Cells(i + 1, 4) = y(3)
        ActiveSheet.Cells(i + 1, 5) = y(4)
    Next i
End Sub

Public Sub DoRK2ndOrder()
    Dim t As Double     ' Thrust
    Dim Cd As Double    ' Drag coefficient
    Dim M As Double     ' Mass
    Dim dt As Double    ' Time step size
    Dim F As Double     ' Force
    Dim A As Double     ' Acceleration
    Dim Vn As Double    ' Velocity at time t
    Dim Vn1 As Double   ' Velocity at time t + dt
    Dim Sn As Double    ' Displacement at time t
    Dim Sn1 As Double   ' Displacement at time t + dt
    Dim time As Double  ' Total time
    Dim k1 As Double    ' Runge-Kutta k1
    Dim k2 As Double    ' Runge-Kutta k2
    Dim k3 As Double    ' Runge-Kutta k3
    Dim k4 As Double    ' Runge-Kutta k4
    Dim n As Integer    ' Counter controlling total number of time steps
    Dim C As Integer    ' Counter controlling output of results to spreadsheet
    Dim k As Integer    ' Counter controlling output row
    Dim r As Integer    ' Number of output rows
    
    ' Extract given data from the active spreadsheet:
    With ActiveSheet
        dt = .Range("dt")
        t = .Range("T")
        M = .Range("M")
        Cd = .Range("Cd")
        n = .Range("n")
        r = .Range("r_")
    End With
    
    ' Initialize variables:
    k = 1
    time = 0
    C = n / r
    Vn = 0
    Sn = 0
    
    ' Start iterations:
    For i = 1 To n
        
        ' Compute k1:
        F = (t - (Cd * Vn))
        A = F / M
        k1 = dt * A
    
        ' Compute k2:
        F = (t - (Cd * (Vn + k1 / 2)))
        A = F / M
        k2 = dt * A
    
        ' Compute k3:
        F = (t - (Cd * (Vn + k2 / 2)))
        A = F / M
        k3 = dt * A
    
        ' Compute k4:
        F = (t - (Cd * (Vn + k3)))
        A = F / M
        k4 = dt * A
    
        ' Compute velocity at t + dt:
        Vn1 = Vn + (k1 + 2 * k2 + 2 * k3 + k4) / 6
        
        ' Compute displacement at t + dt using Euler prediction:
        Sn1 = Sn + Vn1 * dt
    
        ' Update variables:
        time = time + dt
        Vn = Vn1
        Sn = Sn1
        
        ' Output results to the active spreadsheet:
        If C >= n / r Then
            ActiveSheet.Cells(k + 1, 1) = time
            ActiveSheet.Cells(k + 1, 2) = Sn
            ActiveSheet.Cells(k + 1, 3) = Vn
            k = k + 1
            C = 0
        Else
            C = C + 1
        End If
    Next i
End Sub


Public Sub DoEuler1stOrder()
    Dim yn As Double
    Dim yn1 As Double
    Dim xn As Double
    Dim dx As Double
    Dim n As Integer
    Dim C As Integer
    Dim k As Integer
                    
    yn = 0
    xn = 0
    dx = 0.0001
    n = 11000
    C = n / 10
    k = 1
    
    For i = 1 To n
        yn1 = yn + (xn + yn) * dx
        xn = xn + dx
        yn = yn1
        
        If C >= (n / 10) Then
            ActiveSheet.Cells(k + 1, 1) = xn
            ActiveSheet.Cells(k + 1, 2) = yn
            k = k + 1
            C = 0
        Else
            C = C + 1
        End If
    Next i
End Sub

Public Sub DoEuler1stOrder1()
    Dim yn As Double
    Dim yn1 As Double
    Dim xn As Double
    Dim dx As Double
    Dim n As Integer
                    
    yn = 0
    xn = 0
    dx = 0.001
    n = 1000
    
    For i = 1 To n
        yn1 = yn + (xn + yn) * dx
        xn = xn + dx
        yn = yn1
        ActiveSheet.Cells(i + 1, 1) = xn
        ActiveSheet.Cells(i + 1, 2) = yn
    Next i
End Sub




Public Sub DoImprovedEuler()
    Dim y(1 To 4) As Double
    Dim k1(1 To 4) As Double
    Dim k2(1 To 4) As Double
    Dim tmp(1 To 4) As Double
    Dim n As Integer
    Dim dt As Double
    Dim time As Double
    
    y(1) = 29.8613
    y(2) = 0
    y(3) = 30.0616
    y(4) = 0
    n = 5000
    dt = 5 / n
    time = 0
    
    For i = 1 To n
        k1(1) = dt * dy1dt(y(1), y(2), y(3), y(4))
        k1(2) = dt * dy2dt(y(1), y(2), y(3), y(4))
        k1(3) = dt * dy3dt(y(1), y(2), y(3), y(4))
        k1(4) = dt * dy4dt(y(1), y(2), y(3), y(4))
               
        k2(1) = dt * dy1dt((y(1) + k1(1)), (y(2) + k1(2)), (y(3) + k1(3)), (y(4) + k1(4)))
        k2(2) = dt * dy2dt((y(1) + k1(1)), (y(2) + k1(2)), (y(3) + k1(3)), (y(4) + k1(4)))
        k2(3) = dt * dy3dt((y(1) + k1(1)), (y(2) + k1(2)), (y(3) + k1(3)), (y(4) + k1(4)))
        k2(4) = dt * dy4dt((y(1) + k1(1)), (y(2) + k1(2)), (y(3) + k1(3)), (y(4) + k1(4)))

        y(1) = y(1) + (k1(1) + k2(1)) / 2
        y(2) = y(2) + (k1(2) + k2(2)) / 2
        y(3) = y(3) + (k1(3) + k2(3)) / 2
        y(4) = y(4) + (k1(4) + k2(4)) / 2
         
        time = time + dt
        ActiveSheet.Cells(i + 1, 1) = time
        ActiveSheet.Cells(i + 1, 2) = y(1)
        ActiveSheet.Cells(i + 1, 3) = y(2)
        ActiveSheet.Cells(i + 1, 4) = y(3)
        ActiveSheet.Cells(i + 1, 5) = y(4)
    Next i
    Beep
    
End Sub

Public Sub DoRungeKutta()
    ' Declar Local variables
    Dim y(1 To 4) As Double
    Dim k1(1 To 4) As Double
    Dim k2(1 To 4) As Double
    Dim k3(1 To 4) As Double
    Dim k4(1 To 4) As Double
    Dim n As Integer
    Dim dt As Double
    Dim time As Double
        
    Application.ScreenUpdating = False
        
    ' Initialize variables
    y(1) = 29.8613
    y(2) = 0
    y(3) = 30.0616
    y(4) = 0
    n = 5000
    dt = 5 / n:
    time = 0
    
    ' Perform iterations
    For i = 1 To n
        k1(1) = dt * dy1dt(y(1), y(2), y(3), y(4))
        k1(2) = dt * dy2dt(y(1), y(2), y(3), y(4))
        k1(3) = dt * dy3dt(y(1), y(2), y(3), y(4))
        k1(4) = dt * dy4dt(y(1), y(2), y(3), y(4))
        
        k2(1) = dt * dy1dt(y(1) + k1(1) / 2#, y(2) + k1(2) / 2#, y(3) + k1(3) / 2#, y(4) + k1(4) / 2#)
        k2(2) = dt * dy2dt(y(1) + k1(1) / 2#, y(2) + k1(2) / 2#, y(3) + k1(3) / 2#, y(4) + k1(4) / 2#)
        k2(3) = dt * dy3dt(y(1) + k1(1) / 2#, y(2) + k1(2) / 2#, y(3) + k1(3) / 2#, y(4) + k1(4) / 2#)
        k2(4) = dt * dy4dt(y(1) + k1(1) / 2#, y(2) + k1(2) / 2#, y(3) + k1(3) / 2#, y(4) + k1(4) / 2#)

        k3(1) = dt * dy1dt(y(1) + k2(1) / 2#, y(2) + k2(2) / 2#, y(3) + k2(3) / 2#, y(4) + k2(4) / 2#)
        k3(2) = dt * dy2dt(y(1) + k2(1) / 2#, y(2) + k2(2) / 2#, y(3) + k2(3) / 2#, y(4) + k2(4) / 2#)
        k3(3) = dt * dy3dt(y(1) + k2(1) / 2#, y(2) + k2(2) / 2#, y(3) + k2(3) / 2#, y(4) + k2(4) / 2#)
        k3(4) = dt * dy4dt(y(1) + k2(1) / 2#, y(2) + k2(2) / 2#, y(3) + k2(3) / 2#, y(4) + k2(4) / 2#)

        k4(1) = dt * dy1dt(y(1) + k3(1), y(2) + k3(2), y(3) + k3(3), y(4) + k3(4))
        k4(2) = dt * dy2dt(y(1) + k3(1), y(2) + k3(2), y(3) + k3(3), y(4) + k3(4))
        k4(3) = dt * dy3dt(y(1) + k3(1), y(2) + k3(2), y(3) + k3(3), y(4) + k3(4))
        k4(4) = dt * dy4dt(y(1) + k3(1), y(2) + k3(2), y(3) + k3(3), y(4) + k3(4))

        y(1) = y(1) + (k1(1) + 2 * k2(1) + 2 * k3(1) + k4(1)) / 6
        y(2) = y(2) + (k1(2) + 2 * k2(2) + 2 * k3(2) + k4(2)) / 6
        y(3) = y(3) + (k1(3) + 2 * k2(3) + 2 * k3(3) + k4(3)) / 6
        y(4) = y(4) + (k1(4) + 2 * k2(4) + 2 * k3(4) + k4(4)) / 6
         
        time = time + dt
        ActiveSheet.Cells(i + 1, 1) = time
        ActiveSheet.Cells(i + 1, 2) = y(1)
        ActiveSheet.Cells(i + 1, 3) = y(2)
        ActiveSheet.Cells(i + 1, 4) = y(3)
        ActiveSheet.Cells(i + 1, 5) = y(4)
    Next i
    Application.ScreenUpdating = True
    Application.ActiveWorkbook.Save
 
End Sub

Public Function DoDivide() As Double
    DoDivide = 1 / 2
End Function

Public Sub DoLoops()
    Dim time As Integer
    
    Dim M(4, 4) As Integer
    Dim StopNow As Boolean
    
    
    
    
    time = 0
    
    M(1, 1) = 3.456
        

    Do While (time < 32000)
        ' Statements go here
        time = time + 100
    Loop
    
    
    
    
    Do
        ' Statments go here
        time = time + 100
    Loop While (time < 32000)
    
    
    Do Until (time > 32000)
        ' Statements go here
        time = time + 100
    Loop
    
    
    Do
        ' Statments go here
        time = time + 100
    Loop Until (time > 32000)
    
    If (StopNow = True) Then
    
    End If
    
    If (time = 32000) Then
    
    End If
    
    
    
    If (time > 10000) Then
       ' statements
    Else
       ' more statements
    End If
    
    If (time > 32000) Then
       ' statements
    ElseIf (time < 10000) Then
       ' statements
    ElseIf (time < 15) Then
       ' statements
    End If
       
    
    
    
End Sub

Public Sub Testfunctions()
    Dim num1 As Integer
    Dim num2 As Double
    Dim num3 As Double
    Dim r As Double
    Dim b As Double
    Dim A As Double
    
    b = 0.25
    A = 0#
    r = 2#
    
    
    
    area = WorksheetFunction.Pi * r ^ 2
    
    A = WorksheetFunction.Acos(b)
    
    Set TestRange = Worksheets("Sheet1").Range("A1:A5")
    Sum = WorksheetFunction.Sum(TestRange)
    
    Worksheets("Sheet1").Range("A7").Value = Sum
    Worksheets("Sheet1").Range("B7").Formula = "=Sum(A1:A5)"

'For Each c In Worksheets("Sheet1").Range("A1:D10")
'    If c.Value < 0.001 Then
'        c.Value = 0
'    End If
' Next c

    

    
    
End Sub



