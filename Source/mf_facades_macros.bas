Attribute VB_Name = "MF"
Sub MF102()

' Creating search objects
Dim AP As Excel.Application
Dim WB As Excel.Workbook
Dim WS As Excel.Worksheet

Set AP = Excel.Application
Set WB = AP.Workbooks.Open("C:\Users\������\Desktop\������\2.xlsx")
Set WS = WB.Worksheets("������ �˨������ ������")


c = 19
a = Cells(c, 2)
b = Cells(c, 3)
d = Cells(c, 4)
E = 1
I = 0
On Error Resume Next

' Description of all the points that will make up rectangles
  Dim plineObj As AcadLWPolyline
  Dim points(0 To 7) As Double
points(6) = 0

Do While a > 0

' Calculate the coordinates of each point
  points(0) = points(6) + I:    points(1) = 0
  points(2) = points(6) + I:    points(3) = a
  points(4) = points(0) + b:    points(5) = a
  points(6) = points(0) + b:    points(7) = 0

' Constructing four segments and connecting the corresponding points
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
  plineObj.Closed = True
 
Dim a1(0 To 2) As Double
Dim a2(0 To 2) As Double
Dim A3(0 To 2) As Double
Dim A4(0 To 2) As Double
Dim A5(0 To 2) As Double
Dim A6(0 To 2) As Double
Dim A7(0 To 2) As Double
Dim A8(0 To 2) As Double
Dim lineObj As AcadLine
  
  a1(0) = points(0) + 60: a1(1) = 0:      a1(2) = 0
  a2(0) = points(2) + 60: a2(1) = a:      a2(2) = 0
  A3(0) = points(4) - 60: A3(1) = a:      A3(2) = 0
  A4(0) = points(6) - 60: A4(1) = 0:      A4(2) = 0
  
  A5(0) = points(0) + 60: A5(1) = 90:      A5(2) = 0
  A6(0) = points(2) + 60: A6(1) = a - 90:  A6(2) = 0
  A7(0) = points(4) - 60: A7(1) = a - 90:  A7(2) = 0
  A8(0) = points(6) - 60: A8(1) = 90:      A8(2) = 0

If a > 70 Then
If b > 70 Then

lineObj.Layer = "Ball-12"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(a1, a2)
lineObj.Layer = "Ball-12"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A6, A7)
lineObj.Layer = "Ball-12"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A3, A4)
lineObj.Layer = "Ball-12"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A8, A5)
lineObj.Layer = "Ball-12"
lineObj.Update
    
End If
End If

  ' Offset the polyline
  ' Dim offsetObj As Variant
 
                                 'plineObj.color = acByLayer
                                 'Dim layerObj As AcadLayer
                                 'Set layerObj = ThisDrawing.Layers.Add("Ball-6.1")
                                 'layerObj.color = acRed

I = 100
  
 Dim plineObj2 As AcadLWPolyline
  Dim points2(0 To 7) As Double
  
  points2(0) = points(0):    points2(1) = a + I:
  points2(2) = points(0):    points2(3) = (a * 2) + I:
  points2(4) = points(6):    points2(5) = (a * 2) + I:
  points2(6) = points(6):    points2(7) = a + I:
  
 For d = 2 To Cells(c, 4)
If d > 1 Then

Set plineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(points2)
  plineObj2.Closed = True

  
  a1(0) = points2(0) + 60: a1(1) = points2(1):     a1(2) = 0
  a2(0) = points2(2) + 60: a2(1) = points2(3):     a2(2) = 0
  A3(0) = points2(4) - 60: A3(1) = points2(5):     A3(2) = 0
  A4(0) = points2(6) - 60: A4(1) = points2(7):     A4(2) = 0
  
  A5(0) = points2(0) + 60: A5(1) = points2(1) + 90:    A5(2) = 0
  A6(0) = points2(2) + 60: A6(1) = points2(3) - 90:    A6(2) = 0
  A7(0) = points2(4) - 60: A7(1) = points2(5) - 90:    A7(2) = 0
  A8(0) = points2(6) - 60: A8(1) = points2(7) + 90:    A8(2) = 0
  
If a > 70 Then
If b > 70 Then
  
lineObj.Layer = "Ball-12"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(a1, a2)
lineObj.Layer = "Ball-12"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A6, A7)
lineObj.Layer = "Ball-12"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A3, A4)
lineObj.Layer = "Ball-12"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A5, A8)
lineObj.Layer = "Ball-12"
lineObj.Update
  
End If
End If

  points2(1) = points2(3) + I:
  points2(3) = points2(3) + a + I:
  points2(5) = points2(3):
  points2(7) = points2(1):

Else: d = 1
d = d + 1

d = Cells(c, 4)
End If

Next d
c = c + 1

a = Cells(c, 2)
b = Cells(c, 3)

I = 100

Loop
              
ZoomAll
AP.Quit
End Sub

Sub MF104()

' Creating search objects
Dim AP As Excel.Application
Dim WB As Excel.Workbook
Dim WS As Excel.Worksheet

Set AP = Excel.Application
Set WB = AP.Workbooks.Open("C:\Users\������\Desktop\������\2.xlsx")
Set WS = WB.Worksheets("������ �˨������ ������")
c = 19
a = Cells(c, 2)
b = Cells(c, 3)
d = Cells(c, 4)
E = 1
I = 0
On Error Resume Next

' Description of all the points that will make up rectangles
 Dim plineObj As AcadLWPolyline
  Dim points(0 To 7) As Double
points(6) = 0

Do While a > 0

' Calculate the coordinates of each point
  points(0) = points(6) + I:    points(1) = 0
  points(2) = points(6) + I:    points(3) = a
  points(4) = points(0) + b:    points(5) = a
  points(6) = points(0) + b:    points(7) = 0

' Constructing four segments and connecting the corresponding points (facade dimensions)
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
  plineObj.Closed = True
  
 
                                 'plineObj.color = acByLayer
                                 'Dim layerObj As AcadLayer
                                 'Set layerObj = ThisDrawing.Layers.Add("Ball-6.1")
                                 'layerObj.color = acRed
If a > 256 Then
If b > 256 Then

' Creating dots and constructing the top pattern
'===================================
Dim droppoints(0 To 7) As Double
  
  droppoints(0) = points(0) + 35.5:       droppoints(1) = points(3) - 27
  droppoints(2) = points(0) + 44.5:       droppoints(3) = points(3) - 27
  droppoints(4) = points(0) + 42:         droppoints(5) = points(3) - (a * 0.55)
  droppoints(6) = points(0) + 38:         droppoints(7) = points(3) - (a * 0.55)
  
    
  Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(droppoints)
    plineObj.Layer = "Ball-12"
    plineObj.Update

   ' Find the bulge of the third segment
    Dim currentBulge As Double
    currentBulge = plineObj.GetBulge(2)

    ' Change the bulge of the third segment
    plineObj.SetBulge 0, -1
    plineObj.Update
    plineObj.SetBulge 2, -1
    plineObj.Update
    plineObj.Layer = "Ball-12"
    plineObj.Update
  plineObj.Closed = True
  
  droppoints(0) = points(0) + 61.5:       droppoints(1) = points(3) - 53
  droppoints(2) = points(0) + 70.5:       droppoints(3) = points(3) - 53
  droppoints(4) = points(0) + 68:         droppoints(5) = points(3) - ((a * 0.55) - 26)
  droppoints(6) = points(0) + 64:         droppoints(7) = points(3) - ((a * 0.55) - 26)
  
    
  Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(droppoints)
    plineObj.Layer = "Ball-12"
    plineObj.Update
   ' Find the bulge of the third segment
    currentBulge = plineObj.GetBulge(2)
    ' Change the bulge of the third segment
    plineObj.SetBulge 0, -1
    plineObj.Update
    plineObj.SetBulge 2, -1
    plineObj.Update
    plineObj.Layer = "Ball-12"
    plineObj.Update
  plineObj.Closed = True

  droppoints(0) = points(0) + 27:           droppoints(1) = points(3) - 44.5
  droppoints(2) = points(0) + 27:           droppoints(3) = points(3) - 35.5
  droppoints(4) = points(0) + (b * 0.55):   droppoints(5) = points(3) - 38
  droppoints(6) = points(0) + (b * 0.55):   droppoints(7) = points(3) - 42
  
    
  Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(droppoints)
    plineObj.Layer = "Ball-12"
    plineObj.Update
   ' Find the bulge of the third segment
    currentBulge = plineObj.GetBulge(2)
    ' Change the bulge of the third segment
    plineObj.SetBulge 0, -1
    plineObj.Update
    plineObj.SetBulge 2, -1
    plineObj.Update
    plineObj.Layer = "Ball-12"
    plineObj.Update
  plineObj.Closed = True
 
  droppoints(0) = points(0) + 53:                    droppoints(1) = points(3) - 70.5
  droppoints(2) = points(0) + 53:                    droppoints(3) = points(3) - 61.5:
  droppoints(4) = points(0) + ((b * 0.55) - 26):     droppoints(5) = points(3) - 64
  droppoints(6) = points(0) + ((b * 0.55) - 26):     droppoints(7) = points(3) - 68

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(droppoints)
    plineObj.Layer = "Ball-12"
    plineObj.Update
   ' Find the bulge of the third segment
    currentBulge = plineObj.GetBulge(2)
    ' Change the bulge of the third segment
    plineObj.SetBulge 0, -1
    plineObj.Update
    plineObj.SetBulge 2, -1
    plineObj.Update
    plineObj.Layer = "Ball-12"
    plineObj.Update
  plineObj.Closed = True
  
' Creating dots and constructing the bottom pattern
'===================================
  droppoints(0) = points(4) - 35.5:       droppoints(1) = points(7) + 27
  droppoints(2) = points(4) - 44.5:       droppoints(3) = points(7) + 27
  droppoints(4) = points(4) - 42:         droppoints(5) = points(7) + (a * 0.55)
  droppoints(6) = points(4) - 38:         droppoints(7) = points(7) + (a * 0.55)
  
    
  Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(droppoints)
    plineObj.Layer = "Ball-12"
    plineObj.Update
   ' Find the bulge of the third segment
    currentBulge = plineObj.GetBulge(2)
    ' Change the bulge of the third segment
    plineObj.SetBulge 0, -1
    plineObj.Update
    plineObj.SetBulge 2, -1
    plineObj.Update
    plineObj.Layer = "Ball-12"
    plineObj.Update
  plineObj.Closed = True
  
  droppoints(0) = points(4) - 61.5:       droppoints(1) = points(7) + 53
  droppoints(2) = points(4) - 70.5:       droppoints(3) = points(7) + 53
  droppoints(4) = points(4) - 68:         droppoints(5) = points(7) + ((a * 0.55) - 26)
  droppoints(6) = points(4) - 64:         droppoints(7) = points(7) + ((a * 0.55) - 26)
  
    
  Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(droppoints)
    plineObj.Layer = "Ball-12"
    plineObj.Update
   ' Find the bulge of the third segment
    currentBulge = plineObj.GetBulge(2)
    ' Change the bulge of the third segment
    plineObj.SetBulge 0, -1
    plineObj.Update
    plineObj.SetBulge 2, -1
    plineObj.Update
    plineObj.Layer = "Ball-12"
    plineObj.Update
  plineObj.Closed = True

  droppoints(0) = points(4) - 27:           droppoints(1) = points(7) + 44.5
  droppoints(2) = points(4) - 27:           droppoints(3) = points(7) + 35.5
  droppoints(4) = points(4) - (b * 0.55):   droppoints(5) = points(7) + 38
  droppoints(6) = points(4) - (b * 0.55):   droppoints(7) = points(7) + 42
  
    
  Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(droppoints)
    plineObj.Layer = "Ball-12"
    plineObj.Update
   ' Find the bulge of the third segment
    currentBulge = plineObj.GetBulge(2)
    ' Change the bulge of the third segment
    plineObj.SetBulge 0, -1
    plineObj.Update
    plineObj.SetBulge 2, -1
    plineObj.Update
    plineObj.Layer = "Ball-12"
    plineObj.Update
  plineObj.Closed = True
  
  droppoints(0) = points(4) - 53:                    droppoints(1) = points(7) + 70.5
  droppoints(2) = points(4) - 53:                    droppoints(3) = points(7) + 61.5:
  droppoints(4) = points(4) - ((b * 0.55) - 26):     droppoints(5) = points(7) + 64
  droppoints(6) = points(4) - ((b * 0.55) - 26):     droppoints(7) = points(7) + 68

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(droppoints)
    plineObj.Layer = "Ball-12"
    plineObj.Update
   ' Find the bulge of the third segment
    currentBulge = plineObj.GetBulge(2)
    ' Change the bulge of the third segment
    plineObj.SetBulge 0, -1
    plineObj.Update
    plineObj.SetBulge 2, -1
    plineObj.Update
    plineObj.Layer = "Ball-12"
    plineObj.Update
  plineObj.Closed = True

End If
End If

I = 100
 
 
 ' For quantity greater than 1
  Dim plineObj2 As AcadLWPolyline
  Dim points2(0 To 7) As Double
  
  points2(0) = points(0):    points2(1) = a + I:
  points2(2) = points(0):    points2(3) = (a * 2) + I:
  points2(4) = points(6):    points2(5) = (a * 2) + I:
  points2(6) = points(6):    points2(7) = a + I:
  
 For d = 2 To Cells(c, 4)
If d > 1 Then

Set plineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(points2)
  plineObj2.Closed = True

If a > 256 Then
If b > 256 Then
    
' Creating dots and constructing the top pattern
'===================================
  
  droppoints(0) = points2(0) + 35.5:       droppoints(1) = points2(3) - 27
  droppoints(2) = points2(0) + 44.5:       droppoints(3) = points2(3) - 27
  droppoints(4) = points2(0) + 42:         droppoints(5) = points2(3) - (a * 0.55)
  droppoints(6) = points2(0) + 38:         droppoints(7) = points2(3) - (a * 0.55)
  
    
  Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(droppoints)
    plineObj.Layer = "Ball-12"
    plineObj.Update
  ' Find the bulge of the third segment
    currentBulge = plineObj.GetBulge(2)
    ' Change the bulge of the third segment
    plineObj.SetBulge 0, -1
    plineObj.Update
    plineObj.SetBulge 2, -1
    plineObj.Update
    plineObj.Layer = "Ball-12"
    plineObj.Update
  plineObj.Closed = True
  
  droppoints(0) = points2(0) + 61.5:       droppoints(1) = points2(3) - 53
  droppoints(2) = points2(0) + 70.5:       droppoints(3) = points2(3) - 53
  droppoints(4) = points2(0) + 68:         droppoints(5) = points2(3) - ((a * 0.55) - 26)
  droppoints(6) = points2(0) + 64:         droppoints(7) = points2(3) - ((a * 0.55) - 26)
  
    
  Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(droppoints)
    plineObj.Layer = "Ball-12"
    plineObj.Update
  ' Find the bulge of the third segment
    currentBulge = plineObj.GetBulge(2)
    ' Change the bulge of the third segment
    plineObj.SetBulge 0, -1
    plineObj.Update
    plineObj.SetBulge 2, -1
    plineObj.Update
    plineObj.Layer = "Ball-12"
    plineObj.Update
  plineObj.Closed = True

  droppoints(0) = points2(0) + 27:           droppoints(1) = points2(3) - 44.5
  droppoints(2) = points2(0) + 27:           droppoints(3) = points2(3) - 35.5
  droppoints(4) = points2(0) + (b * 0.55):   droppoints(5) = points2(3) - 38
  droppoints(6) = points2(0) + (b * 0.55):   droppoints(7) = points2(3) - 42
  
    
  Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(droppoints)
    plineObj.Layer = "Ball-12"
    plineObj.Update
  ' Find the bulge of the third segment
    currentBulge = plineObj.GetBulge(2)
    ' Change the bulge of the third segment
    plineObj.SetBulge 0, -1
    plineObj.Update
    plineObj.SetBulge 2, -1
    plineObj.Update
    plineObj.Layer = "Ball-12"
    plineObj.Update
  plineObj.Closed = True
 
  droppoints(0) = points2(0) + 53:                    droppoints(1) = points2(3) - 70.5
  droppoints(2) = points2(0) + 53:                    droppoints(3) = points2(3) - 61.5:
  droppoints(4) = points2(0) + ((b * 0.55) - 26):     droppoints(5) = points2(3) - 64
  droppoints(6) = points2(0) + ((b * 0.55) - 26):     droppoints(7) = points2(3) - 68

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(droppoints)
    plineObj.Layer = "Ball-12"
    plineObj.Update
  ' Find the bulge of the third segment
    currentBulge = plineObj.GetBulge(2)
    ' Change the bulge of the third segment
    plineObj.SetBulge 0, -1
    plineObj.Update
    plineObj.SetBulge 2, -1
    plineObj.Update
    plineObj.Layer = "Ball-12"
    plineObj.Update
  plineObj.Closed = True
  
' Creating dots and constructing the bottom pattern
'===================================
  droppoints(0) = points2(4) - 35.5:       droppoints(1) = points2(7) + 27
  droppoints(2) = points2(4) - 44.5:       droppoints(3) = points2(7) + 27
  droppoints(4) = points2(4) - 42:         droppoints(5) = points2(7) + (a * 0.55)
  droppoints(6) = points2(4) - 38:         droppoints(7) = points2(7) + (a * 0.55)
  
    
  Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(droppoints)
    plineObj.Layer = "Ball-12"
    plineObj.Update
  ' Find the bulge of the third segment
    currentBulge = plineObj.GetBulge(2)
    ' Change the bulge of the third segment
    plineObj.SetBulge 0, -1
    plineObj.Update
    plineObj.SetBulge 2, -1
    plineObj.Update
    plineObj.Layer = "Ball-12"
    plineObj.Update
  plineObj.Closed = True
  
  droppoints(0) = points2(4) - 61.5:       droppoints(1) = points2(7) + 53
  droppoints(2) = points2(4) - 70.5:       droppoints(3) = points2(7) + 53
  droppoints(4) = points2(4) - 68:         droppoints(5) = points2(7) + ((a * 0.55) - 26)
  droppoints(6) = points2(4) - 64:         droppoints(7) = points2(7) + ((a * 0.55) - 26)
  
    
  Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(droppoints)
    plineObj.Layer = "Ball-12"
    plineObj.Update
  ' Find the bulge of the third segment
    currentBulge = plineObj.GetBulge(2)
    ' Change the bulge of the third segment
    plineObj.SetBulge 0, -1
    plineObj.Update
    plineObj.SetBulge 2, -1
    plineObj.Update
    plineObj.Layer = "Ball-12"
    plineObj.Update
  plineObj.Closed = True

  droppoints(0) = points2(4) - 27:           droppoints(1) = points2(7) + 44.5
  droppoints(2) = points2(4) - 27:           droppoints(3) = points2(7) + 35.5
  droppoints(4) = points2(4) - (b * 0.55):   droppoints(5) = points2(7) + 38
  droppoints(6) = points2(4) - (b * 0.55):   droppoints(7) = points2(7) + 42
  
    
  Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(droppoints)
    plineObj.Layer = "Ball-12"
    plineObj.Update
  ' Find the bulge of the third segment
    currentBulge = plineObj.GetBulge(2)
    ' Change the bulge of the third segment
    plineObj.SetBulge 0, -1
    plineObj.Update
    plineObj.SetBulge 2, -1
    plineObj.Update
    plineObj.Layer = "Ball-12"
    plineObj.Update
  plineObj.Closed = True
 
  droppoints(0) = points2(4) - 53:                    droppoints(1) = points2(7) + 70.5
  droppoints(2) = points2(4) - 53:                    droppoints(3) = points2(7) + 61.5:
  droppoints(4) = points2(4) - ((b * 0.55) - 26):     droppoints(5) = points2(7) + 64
  droppoints(6) = points2(4) - ((b * 0.55) - 26):     droppoints(7) = points2(7) + 68

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(droppoints)
    plineObj.Layer = "Ball-12"
    plineObj.Update
  ' Find the bulge of the third segment
    currentBulge = plineObj.GetBulge(2)
    ' Change the bulge of the third segment
    plineObj.SetBulge 0, -1
    plineObj.Update
    plineObj.SetBulge 2, -1
    plineObj.Update
    plineObj.Layer = "Ball-12"
    plineObj.Update
  plineObj.Closed = True

End If
End If
  
  points2(1) = points2(3) + I:
  points2(3) = points2(3) + a + I:
  points2(5) = points2(3):
  points2(7) = points2(1):

    
Else: d = 1
d = d + 1

d = Cells(c, 4)
End If

Next d
c = c + 1

a = Cells(c, 2)
b = Cells(c, 3)

I = 100

Loop
              
ZoomAll
AP.Quit
End Sub

Sub MF105()
' Creating search objects
Dim AP As Excel.Application
Dim WB As Excel.Workbook
Dim WS As Excel.Worksheet

Set AP = Excel.Application
Set WB = AP.Workbooks.Open("C:\Users\������\Desktop\������\2.xlsx")
Set WS = WB.Worksheets("������ �˨������ ������")
c = 19
a = Cells(c, 2)
b = Cells(c, 3)
d = Cells(c, 4)
E = 1
I = 0
On Error Resume Next

' Description of all the points that will make up rectangles
  Dim plineObj As AcadLWPolyline
  Dim points(0 To 7) As Double
  Dim offsetObj As Variant
points(6) = 0

Do While a > 0

' Calculate the coordinates of each point
  points(0) = points(6) + I:    points(1) = 0
  points(2) = points(6) + I:    points(3) = a
  points(4) = points(0) + b:    points(5) = a
  points(6) = points(0) + b:    points(7) = 0

' Constructing four segments and connecting the corresponding points

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
  plineObj.Closed = True
 
Dim a1(0 To 2) As Double
Dim a2(0 To 2) As Double
Dim A3(0 To 2) As Double
Dim A4(0 To 2) As Double
Dim A5(0 To 2) As Double
Dim A6(0 To 2) As Double
Dim A7(0 To 2) As Double
Dim A8(0 To 2) As Double
Dim lineObj As AcadLine
  
  a1(0) = points(0) + 30: a1(1) = 0:                a1(2) = 0
  a2(0) = points(2) + 30: a2(1) = a:                a2(2) = 0
  A3(0) = points(0):      A3(1) = points(1) + 30:   A3(2) = 0
  A4(0) = points(6):      A4(1) = points(1) + 30:   A4(2) = 0

If a > 60 Then
If b > 60 Then
lineObj.Layer = "Ball-6"
lineObj.Update
  Set lineObj = ThisDrawing.ModelSpace.AddLine(a1, a2)
lineObj.Layer = "Ball-6"
lineObj.Update
  offsetObj = lineObj.Offset(-30)
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A3, A4)
lineObj.Layer = "Ball-6"
lineObj.Update
  offsetObj = lineObj.Offset(30)
lineObj.Layer = "Ball-6"
lineObj.Update
 
 End If
 End If
                                 'plineObj.color = acByLayer
                                 'Dim layerObj As AcadLayer
                                 'Set layerObj = ThisDrawing.Layers.Add("Ball-6.1")
                                 'layerObj.color = acRed

I = 100
  
 Dim plineObj2 As AcadLWPolyline
  Dim points2(0 To 7) As Double
  
  points2(0) = points(0):    points2(1) = a + I:
  points2(2) = points(0):    points2(3) = (a * 2) + I:
  points2(4) = points(6):    points2(5) = (a * 2) + I:
  points2(6) = points(6):    points2(7) = a + I:
  
For d = 2 To Cells(c, 4)
If d > 1 Then

Set plineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(points2)
  plineObj2.Closed = True
  
  a1(0) = points2(0) + 30: a1(1) = points2(1):     a1(2) = 0
  a2(0) = points2(2) + 30: a2(1) = points2(3):     a2(2) = 0
  A3(0) = points2(0):      A3(1) = points2(1) + 30:   A3(2) = 0
  A4(0) = points2(6):      A4(1) = points2(1) + 30:   A4(2) = 0
  
If a > 60 Then
If b > 60 Then
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(a1, a2)
lineObj.Layer = "Ball-6"
lineObj.Update
  offsetObj = lineObj.Offset(-30)
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A3, A4)
lineObj.Layer = "Ball-6"
lineObj.Update
  offsetObj = lineObj.Offset(30)
lineObj.Layer = "Ball-6"
lineObj.Update
End If
End If

  points2(1) = points2(3) + I:
  points2(3) = points2(3) + a + I:
  points2(5) = points2(3):
  points2(7) = points2(1):

Else: d = 1
d = d + 1

d = Cells(c, 4)
End If

Next d
c = c + 1

a = Cells(c, 2)
b = Cells(c, 3)

I = 100

Loop
              
ZoomAll
AP.Quit
End Sub

Sub MF106()
' Creating search objects
Dim AP As Excel.Application
Dim WB As Excel.Workbook
Dim WS As Excel.Worksheet

Set AP = Excel.Application
Set WB = AP.Workbooks.Open("C:\Users\������\Desktop\������\2.xlsx")
Set WS = WB.Worksheets("������ �˨������ ������")
c = 19
a = Cells(c, 2)
b = Cells(c, 3)
d = Cells(c, 4)
E = 1
I = 0
On Error Resume Next

' Description of all the points that will make up rectangles
  Dim plineObj As AcadLWPolyline
  Dim points(0 To 7) As Double
  Dim offsetObj As Variant
points(6) = 0

Do While a > 0

' Calculate the coordinates of each point
  points(0) = points(6) + I:    points(1) = 0
  points(2) = points(6) + I:    points(3) = a
  points(4) = points(0) + b:    points(5) = a
  points(6) = points(0) + b:    points(7) = 0

' Constructing four segments and connecting the corresponding points

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
  plineObj.Closed = True
 
Dim a1(0 To 2) As Double
Dim a2(0 To 2) As Double
Dim A3(0 To 2) As Double
Dim A4(0 To 2) As Double
Dim A5(0 To 2) As Double
Dim A6(0 To 2) As Double
Dim A7(0 To 2) As Double
Dim A8(0 To 2) As Double
Dim lineObj As AcadLine
  
  a1(0) = points(6) - 30: a1(1) = 0:      a1(2) = 0
  a2(0) = points(4) - 30: a2(1) = a:      a2(2) = 0
  A3(0) = points(0):      A3(1) = 30:     A3(2) = 0
  A4(0) = points(6):      A4(1) = 30:     A4(2) = 0

If a > 60 Then
If b > 60 Then
lineObj.Layer = "Ball-6"
lineObj.Update
  Set lineObj = ThisDrawing.ModelSpace.AddLine(a1, a2)
lineObj.Layer = "Ball-6"
lineObj.Update
  offsetObj = lineObj.Offset(15)
lineObj.Layer = "Ball-6"
lineObj.Update
  offsetObj = lineObj.Offset(30)
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A3, A4)
lineObj.Layer = "Ball-6"
lineObj.Update
  offsetObj = lineObj.Offset(15)
lineObj.Layer = "Ball-6"
lineObj.Update
  offsetObj = lineObj.Offset(30)
lineObj.Layer = "Ball-6"
lineObj.Update
 
 End If
 End If
                                 'plineObj.color = acByLayer
                                 'Dim layerObj As AcadLayer
                                 'Set layerObj = ThisDrawing.Layers.Add("Ball-6.1")
                                 'layerObj.color = acRed

I = 100
  
 Dim plineObj2 As AcadLWPolyline
  Dim points2(0 To 7) As Double
  
  points2(0) = points(0):    points2(1) = a + I:
  points2(2) = points(0):    points2(3) = (a * 2) + I:
  points2(4) = points(6):    points2(5) = (a * 2) + I:
  points2(6) = points(6):    points2(7) = a + I:
  
For d = 2 To Cells(c, 4)
If d > 1 Then

Set plineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(points2)
  plineObj2.Closed = True
  
  a1(0) = points2(6) - 30: a1(1) = points2(1):     a1(2) = 0
  a2(0) = points2(4) - 30: a2(1) = points2(3):     a2(2) = 0
  A3(0) = points2(0):      A3(1) = points2(1) + 30:   A3(2) = 0
  A4(0) = points2(6):      A4(1) = points2(1) + 30:   A4(2) = 0
  
If a > 60 Then
If b > 60 Then
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(a1, a2)
lineObj.Layer = "Ball-6"
lineObj.Update
offsetObj = lineObj.Offset(15)
lineObj.Layer = "Ball-6"
lineObj.Update
  offsetObj = lineObj.Offset(30)
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A3, A4)
lineObj.Layer = "Ball-6"
lineObj.Update
offsetObj = lineObj.Offset(15)
lineObj.Layer = "Ball-6"
lineObj.Update
  offsetObj = lineObj.Offset(30)
lineObj.Layer = "Ball-6"
lineObj.Update
End If
End If

  points2(1) = points2(3) + I:
  points2(3) = points2(3) + a + I:
  points2(5) = points2(3):
  points2(7) = points2(1):

Else: d = 1
d = d + 1

d = Cells(c, 4)
End If

Next d
c = c + 1

a = Cells(c, 2)
b = Cells(c, 3)

I = 100

Loop
              
ZoomAll
AP.Quit
End Sub

Sub MF107()
' Creating search objects
Dim AP As Excel.Application
Dim WB As Excel.Workbook
Dim WS As Excel.Worksheet

Set AP = Excel.Application
Set WB = AP.Workbooks.Open("C:\Users\������\Desktop\������\2.xlsx")
Set WS = WB.Worksheets("������ �˨������ ������")
c = 19
a = Cells(c, 2)
b = Cells(c, 3)
d = Cells(c, 4)
E = 1
I = 0
On Error Resume Next

' Description of all the points that will make up rectangles
  Dim plineObj As AcadLWPolyline
  Dim points(0 To 7) As Double
points(6) = 0

Do While a > 0

' Calculate the coordinates of each point
  points(0) = points(6) + I:    points(1) = 0
  points(2) = points(6) + I:    points(3) = a
  points(4) = points(0) + b:    points(5) = a
  points(6) = points(0) + b:    points(7) = 0

' Constructing four segments and connecting the corresponding points

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
  plineObj.Closed = True
 
If a > 120 Then
If b > 120 Then
 
Dim a1(0 To 2) As Double
Dim a2(0 To 2) As Double
Dim A3(0 To 2) As Double
Dim A4(0 To 2) As Double
Dim A5(0 To 2) As Double
Dim A6(0 To 2) As Double
Dim A7(0 To 2) As Double
Dim A8(0 To 2) As Double
Dim lineObj As AcadLine
  
  a1(0) = points(0) + 60: a1(1) = 0:      a1(2) = 0
  a2(0) = points(2) + 60: a2(1) = a:      a2(2) = 0
  A3(0) = points(4) - 60: A3(1) = a:      A3(2) = 0
  A4(0) = points(6) - 60: A4(1) = 0:      A4(2) = 0
  
  A5(0) = points(0):      A5(1) = 60:      A5(2) = 0
  A6(0) = points(2):      A6(1) = a - 60:  A6(2) = 0
  A7(0) = points(4):      A7(1) = a - 60:  A7(2) = 0
  A8(0) = points(6):      A8(1) = 60:      A8(2) = 0

lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(a1, a2)
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A6, A7)
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A3, A4)
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A8, A5)
lineObj.Layer = "Ball-6"
lineObj.Update
  
End If
End If
  ' Offset the polyline
  ' Dim offsetObj As Variant
 
                                 'plineObj.color = acByLayer
                                 'Dim layerObj As AcadLayer
                                 'Set layerObj = ThisDrawing.Layers.Add("Ball-6.1")
                                 'layerObj.color = acRed

I = 100
  
 Dim plineObj2 As AcadLWPolyline
  Dim points2(0 To 7) As Double

  points2(0) = points(0):    points2(1) = a + I:
  points2(2) = points(0):    points2(3) = (a * 2) + I:
  points2(4) = points(6):    points2(5) = (a * 2) + I:
  points2(6) = points(6):    points2(7) = a + I:
  
For d = 2 To Cells(c, 4)
 If d > 1 Then

Set plineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(points2)
  plineObj2.Closed = True
 
If a > 120 Then
If b > 120 Then

  a1(0) = points2(0) + 60: a1(1) = points2(1):     a1(2) = 0
  a2(0) = points2(2) + 60: a2(1) = points2(3):     a2(2) = 0
  A3(0) = points2(4) - 60: A3(1) = points2(5):     A3(2) = 0
  A4(0) = points2(6) - 60: A4(1) = points2(7):     A4(2) = 0
  
  A5(0) = points2(0):     A5(1) = points2(1) + 60:    A5(2) = 0
  A6(0) = points2(2):     A6(1) = points2(3) - 60:    A6(2) = 0
  A7(0) = points2(4):     A7(1) = points2(5) - 60:    A7(2) = 0
  A8(0) = points2(6):     A8(1) = points2(7) + 60:    A8(2) = 0
  

lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(a1, a2)
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A6, A7)
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A3, A4)
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A5, A8)
lineObj.Layer = "Ball-6"
lineObj.Update

End If
End If

  points2(1) = points2(3) + I:
  points2(3) = points2(3) + a + I:
  points2(5) = points2(3):
  points2(7) = points2(1):
   
Else: d = 1
d = d + 1

d = Cells(c, 4)
End If

Next d
c = c + 1

a = Cells(c, 2)
b = Cells(c, 3)

I = 100

Loop
              
ZoomAll
AP.Quit
End Sub

Sub MF108()
' Creating search objects
Dim AP As Excel.Application
Dim WB As Excel.Workbook
Dim WS As Excel.Worksheet

Set AP = Excel.Application
Set WB = AP.Workbooks.Open("C:\Users\������\Desktop\������\2.xlsx")
Set WS = WB.Worksheets("������ �˨������ ������")
c = 19
a = Cells(c, 2)
b = Cells(c, 3)
d = Cells(c, 4)
E = 1
I = 0
On Error Resume Next

' Description of all the points that will make up rectangles
  Dim plineObj As AcadLWPolyline
  Dim points(0 To 7) As Double
points(6) = 0

Do While a > 0

' Calculate the coordinates of each point
  points(0) = points(6) + I:    points(1) = 0
  points(2) = points(6) + I:    points(3) = a
  points(4) = points(0) + b:    points(5) = a
  points(6) = points(0) + b:    points(7) = 0

' Constructing four segments and connecting the corresponding points

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
  plineObj.Closed = True
 
If a > 120 Then
If b > 120 Then
 
Dim a1(0 To 2) As Double
Dim a2(0 To 2) As Double
Dim A3(0 To 2) As Double
Dim A4(0 To 2) As Double
Dim A5(0 To 2) As Double
Dim A6(0 To 2) As Double
Dim A7(0 To 2) As Double
Dim A8(0 To 2) As Double
Dim A9(0 To 2) As Double
Dim A10(0 To 2) As Double
Dim A11(0 To 2) As Double
Dim A12(0 To 2) As Double
Dim lineObj As AcadLine
  
  a1(0) = points(0) + 30: a1(1) = 0:      a1(2) = 0
  a2(0) = points(2) + 30: a2(1) = a:      a2(2) = 0
  A3(0) = points(4) - 30: A3(1) = a:      A3(2) = 0
  A4(0) = points(6) - 30: A4(1) = 0:      A4(2) = 0
  
  A5(0) = points(0) + 45:    A5(1) = 0:      A5(2) = 0
  A6(0) = points(2) + 45:    A6(1) = a:      A6(2) = 0
  A7(0) = points(4) - 45:    A7(1) = a:      A7(2) = 0
  A8(0) = points(6) - 45:    A8(1) = 0:      A8(2) = 0

  A9(0) = points(0) + 60:    A9(1) = 0:      A9(2) = 0
  A10(0) = points(2) + 60:   A10(1) = a:     A10(2) = 0
  A11(0) = points(4) - 60:   A11(1) = a:     A11(2) = 0
  A12(0) = points(6) - 60:   A12(1) = 0:     A12(2) = 0

lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(a1, a2)
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A3, A4)
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A5, A6)
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A7, A8)
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A9, A10)
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A11, A12)
lineObj.Layer = "Ball-6"
lineObj.Update
  
End If
End If
  ' Offset the polyline
  ' Dim offsetObj As Variant
 
                                 'plineObj.color = acByLayer
                                 'Dim layerObj As AcadLayer
                                 'Set layerObj = ThisDrawing.Layers.Add("Ball-6.1")
                                 'layerObj.color = acRed

I = 100
  
 Dim plineObj2 As AcadLWPolyline
  Dim points2(0 To 7) As Double

  points2(0) = points(0):    points2(1) = a + I:
  points2(2) = points(0):    points2(3) = (a * 2) + I:
  points2(4) = points(6):    points2(5) = (a * 2) + I:
  points2(6) = points(6):    points2(7) = a + I:
  
For d = 2 To Cells(c, 4)
 If d > 1 Then

Set plineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(points2)
  plineObj2.Closed = True
 
If a > 120 Then
If b > 120 Then

  a1(0) = points2(0) + 30: a1(1) = points2(1):        a1(2) = 0
  a2(0) = points2(2) + 30: a2(1) = points2(1) + a:    a2(2) = 0
  A3(0) = points2(4) - 30: A3(1) = points2(1) + a:    A3(2) = 0
  A4(0) = points2(6) - 30: A4(1) = points2(1):        A4(2) = 0
  
  A5(0) = points2(0) + 45:    A5(1) = points2(1):          A5(2) = 0
  A6(0) = points2(2) + 45:    A6(1) = points2(1) + a:      A6(2) = 0
  A7(0) = points2(4) - 45:    A7(1) = points2(1) + a:      A7(2) = 0
  A8(0) = points2(6) - 45:    A8(1) = points2(1):          A8(2) = 0

  A9(0) = points2(0) + 60:    A9(1) = points2(1):          A9(2) = 0
  A10(0) = points2(2) + 60:   A10(1) = points2(1) + a:     A10(2) = 0
  A11(0) = points2(4) - 60:   A11(1) = points2(1) + a:     A11(2) = 0
  A12(0) = points2(6) - 60:   A12(1) = points2(1):         A12(2) = 0

lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(a1, a2)
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A3, A4)
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A5, A6)
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A7, A8)
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A9, A10)
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A11, A12)
lineObj.Layer = "Ball-6"
lineObj.Update

End If
End If

  points2(1) = points2(3) + I:
  points2(3) = points2(3) + a + I:
  points2(5) = points2(3):
  points2(7) = points2(1):
   
Else: d = 1
d = d + 1

d = Cells(c, 4)
End If

Next d
c = c + 1

a = Cells(c, 2)
b = Cells(c, 3)

I = 100

Loop
              
ZoomAll
AP.Quit
End Sub

Sub MF109()
' Creating search objects
Dim AP As Excel.Application
Dim WB As Excel.Workbook
Dim WS As Excel.Worksheet

Set AP = Excel.Application
Set WB = AP.Workbooks.Open("C:\Users\������\Desktop\������\2.xlsx")
Set WS = WB.Worksheets("������ �˨������ ������")
c = 19
a = Cells(c, 2)
b = Cells(c, 3)
d = Cells(c, 4)
E = 1
I = 0
' If an error occurs, do not stop the program
On Error Resume Next

' Description of all the points that will make up rectangles
 Dim plineObj As AcadLWPolyline
 Dim braidplineObj As AcadLWPolyline
 Dim braidplineObj2 As AcadLWPolyline
  Dim points(0 To 7) As Double
  Dim P As Variant
points(6) = 0

Do While a > 0

' Calculate the coordinates of each point
  points(0) = points(6) + I:    points(1) = 0
  points(2) = points(6) + I:    points(3) = a
  points(4) = points(0) + b:    points(5) = a
  points(6) = points(0) + b:    points(7) = 0

' Constructing a polyline by connecting corresponding points

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
  plineObj.Closed = True
  
' Description of all ornament points
  Dim braidpoints(0 To 5) As Double
  Dim braidpoints2(0 To 3) As Double
  Dim j As Double
'============================================================================= Vector 1, left
t = a / 240
If a >= 240 Then
t1 = Int(t)
Else: t1 = Round(t, 0)
End If
P = 20
q = 0
t = t1 + 1
j = (a / t) - 20
J1 = j + 20
g = t
S = 30
s1 = 20
P = 0

If a > 139 Then
If b > 139 Then
braidpoints(4) = points(0)
braidpoints(3) = points(1)
' Limitation of the cycle of construction of the ornament
Do While braidpoints(3) < a - j - 50


  braidpoints(0) = braidpoints(4) + S:       braidpoints(1) = braidpoints(3) + S + P
  braidpoints(2) = braidpoints(0):           braidpoints(3) = braidpoints(1) + j - s1
  P = 20
  s1 = 0
' Condition of the direction of deviation along X-axis for constructing crosshairs
If braidpoints(0) = points(0) + 50 Then
q = -20
End If
If braidpoints(0) = points(0) + 30 Then
q = 20
End If
   braidpoints(4) = braidpoints(2) + q: braidpoints(5) = braidpoints(3) + P

Set braidplineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(braidpoints)
braidplineObj.Layer = "Ball-6"
braidplineObj.Update

S = 0
q = 0

P = 20

Loop
' The condition for the end of the last segment of the ornament depending on its position along X-axis
    braidpoints2(0) = braidpoints(4):         braidpoints2(1) = braidpoints(5)
    braidpoints2(2) = braidpoints2(0):        braidpoints2(3) = braidpoints2(1) + j
      If braidpoints2(2) = points(0) + 50 Then
      braidpoints2(3) = a - 50
             End If
      If braidpoints2(2) = points(0) + 30 Then
      braidpoints2(3) = a - 30
              End If
  
Set braidplineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(braidpoints2)
braidplineObj2.Layer = "Ball-6"
braidplineObj2.Update

'============================================================================= Vector 1, top

t = b / 240
If b >= 240 Then
t1 = Int(t)
Else: t1 = Round(t, 0)
End If
P = 20
q = 0
t = t1 + 1
j = (b / t) - 20
J1 = j + 20
g = t
S = 30
s1 = 0
P = 20

      
braidpoints(4) = braidpoints2(2)
braidpoints(5) = braidpoints2(3)

If braidpoints(0) = points(0) + 50 Then
s1 = 0
End If
If braidpoints(0) = points(0) + 30 Then
s1 = 20
End If


' Limitation of the cycle of construction of the ornament
Do While braidpoints(4) - points(0) < b - j - 30
  braidpoints(0) = braidpoints(4):                braidpoints(1) = braidpoints(5)
  braidpoints(2) = braidpoints(0) + j - P - s1:   braidpoints(3) = braidpoints(1)
  P = 20
  s1 = -20

' Condition of the direction of deviation along X-axis for constructing crosshairs
If braidpoints(1) = points(3) - 50 Then
q = 20
End If
If braidpoints(1) = points(3) - 30 Then
q = -20
End If
   braidpoints(4) = braidpoints(2) + P: braidpoints(5) = braidpoints(3) + q

Set braidplineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(braidpoints)
braidplineObj.Layer = "Ball-6"
braidplineObj.Update

S = 0
q = 0
P = 20
Loop

' The condition for the end of the last segment of the ornament depending on its position along X-axis
    braidpoints2(0) = braidpoints(4):           braidpoints2(1) = braidpoints(5)
    braidpoints2(2) = braidpoints2(0) + j:      braidpoints2(3) = braidpoints2(1)
      If braidpoints2(3) = points(3) - 50 Then
      braidpoints2(2) = braidpoints2(2) - 40
             End If
      If braidpoints2(3) = points(3) - 30 Then
      braidpoints2(2) = braidpoints2(2) - 20
              End If
  
Set braidplineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(braidpoints2)
braidplineObj2.Layer = "Ball-6"
braidplineObj2.Update

'============================================================================= = Vector 1, bottom


t = b / 240
If b >= 240 Then
t1 = Int(t)
Else: t1 = Round(t, 0)
End If
P = 20
q = 0
t = t1 + 1
j = (b / t) - 20
J1 = j + 20
g = t
S = 30
s1 = 20
P = 20

      
braidpoints(4) = points(0)
braidpoints(5) = points(1)


' Limitation of the cycle of construction of the ornament
Do While braidpoints(4) - points(0) < b - j - 30
  braidpoints(0) = braidpoints(4) + S:           braidpoints(1) = braidpoints(5) + S
  braidpoints(2) = braidpoints(0) + j - s1:      braidpoints(3) = braidpoints(1)
  P = 20
  s1 = 0

' Condition of the direction of deviation along X-axis for constructing crosshairs
If braidpoints(1) = points(1) + 50 Then
q = -20
End If
If braidpoints(1) = points(1) + 30 Then
q = 20
End If
   braidpoints(4) = braidpoints(2) + P: braidpoints(5) = braidpoints(3) + q

Set braidplineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(braidpoints)
braidplineObj.Layer = "Ball-6"
braidplineObj.Update

S = 0
q = 0
P = 20
Loop

' The condition for the end of the last segment of the ornament depending on its position along X-axis
    braidpoints2(0) = braidpoints(4):           braidpoints2(1) = braidpoints(5)
    braidpoints2(2) = braidpoints2(0) + j:      braidpoints2(3) = braidpoints2(1)
      If braidpoints2(3) = points(1) + 50 Then
      braidpoints2(2) = braidpoints2(2) - 40
             End If
      If braidpoints2(3) = points(1) + 30 Then
      braidpoints2(2) = braidpoints2(2) - 20
              End If
  
Set braidplineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(braidpoints2)
braidplineObj2.Layer = "Ball-6"
braidplineObj2.Update

'============================================================================= Vector 1, right

t = a / 240
If a >= 240 Then
t1 = Int(t)
Else: t1 = Round(t, 0)
End If
P = 20
q = 0
t = t1 + 1
j = (a / t) - 20
J1 = j + 20
g = t
S = 30
s1 = 40
P = 0

braidpoints(4) = braidpoints2(2)
braidpoints(3) = braidpoints2(3)


If braidpoints(1) = points(1) + 50 Then
s1 = 20
End If
If braidpoints(1) = points(1) + 30 Then
s1 = 40
End If


' Limitation of the cycle of construction of the ornament
Do While braidpoints(3) < a - j - 35

 
  braidpoints(0) = braidpoints(4):           braidpoints(1) = braidpoints(3) + P
  braidpoints(2) = braidpoints(0):           braidpoints(3) = braidpoints(1) + j - s1
  P = 20
  s1 = 0
' Condition of the direction of deviation along X-axis for constructing crosshairs
If braidpoints(2) <= points(4) - 45 Then
q = 20
End If
If braidpoints(2) >= points(4) - 35 Then
q = -20
End If
   braidpoints(4) = braidpoints(2) + q: braidpoints(5) = braidpoints(3) + P

Set braidplineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(braidpoints)
braidplineObj.Layer = "Ball-6"
braidplineObj.Update

S = 0
q = 0

P = 20

Loop
' The condition for the end of the last segment of the ornament depending on its position along X-axis
    braidpoints2(0) = braidpoints(4):         braidpoints2(1) = braidpoints(5)
    braidpoints2(2) = braidpoints2(0):        braidpoints2(3) = braidpoints2(1) + j
      If braidpoints2(2) <= points(4) - 45 Then
      braidpoints2(3) = a - 50
             End If
      If braidpoints2(2) >= points(4) - 35 Then
      braidpoints2(3) = a - 30
              End If
  
Set braidplineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(braidpoints2)
braidplineObj2.Layer = "Ball-6"
braidplineObj2.Update

'============================================================================= Vector 2, left
t = a / 240
If a >= 240 Then
t1 = Int(t)
Else: t1 = Round(t, 0)
End If
P = 20
q = 0
t = t1 + 1
j = (a / t) - 20
J1 = j + 20
g = t
S = 50
s1 = 20
P = 0

braidpoints(4) = points(0)
braidpoints(3) = points(1)
' Limitation of the cycle of construction of the ornament
Do While braidpoints(3) < a - j - 50


  braidpoints(0) = braidpoints(4) + S:       braidpoints(1) = braidpoints(3) + S + P
  braidpoints(2) = braidpoints(0):           braidpoints(3) = braidpoints(1) + j - (s1 * 2)
  P = 20
  s1 = 0
' Condition of the direction of deviation along X-axis for constructing crosshairs
If braidpoints(0) = points(0) + 50 Then
q = -20
End If
If braidpoints(0) = points(0) + 30 Then
q = 20
End If
   braidpoints(4) = braidpoints(2) + q: braidpoints(5) = braidpoints(3) + P

Set braidplineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(braidpoints)
braidplineObj.Layer = "Ball-6"
braidplineObj.Update

S = 0
q = 0

P = 20

Loop
' The condition for the end of the last segment of the ornament depending on its position along X-axis
    braidpoints2(0) = braidpoints(4):         braidpoints2(1) = braidpoints(5)
    braidpoints2(2) = braidpoints2(0):        braidpoints2(3) = braidpoints2(1) + j
      If braidpoints2(2) = points(0) + 50 Then
      braidpoints2(3) = a - 50
             End If
      If braidpoints2(2) = points(0) + 30 Then
      braidpoints2(3) = a - 30
              End If
  
Set braidplineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(braidpoints2)
braidplineObj2.Layer = "Ball-6"
braidplineObj2.Update

'============================================================================= Vector 2, top

t = b / 240
If b >= 240 Then
t1 = Int(t)
Else: t1 = Round(t, 0)
End If
P = 20
q = 0
t = t1 + 1
j = (b / t) - 20
J1 = j + 20
g = t
S = 30
s1 = 0
P = 20

      
braidpoints(4) = braidpoints2(2)
braidpoints(5) = braidpoints2(3)

If braidpoints(0) = points(0) + 50 Then
s1 = 0
End If
If braidpoints(0) = points(0) + 30 Then
s1 = 20
End If


' Limitation of the cycle of construction of the ornament
Do While braidpoints(4) - points(0) < b - j - 30
  braidpoints(0) = braidpoints(4):                braidpoints(1) = braidpoints(5)
  braidpoints(2) = braidpoints(0) + j - P - s1:   braidpoints(3) = braidpoints(1)
  P = 20
  s1 = -20

' Condition of the direction of deviation along X-axis for constructing crosshairs
If braidpoints(1) = points(3) - 50 Then
q = 20
End If
If braidpoints(1) = points(3) - 30 Then
q = -20
End If
   braidpoints(4) = braidpoints(2) + P: braidpoints(5) = braidpoints(3) + q

Set braidplineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(braidpoints)
braidplineObj.Layer = "Ball-6"
braidplineObj.Update

S = 0
q = 0
P = 20
Loop

' The condition for the end of the last segment of the ornament depending on its position along X-axis
    braidpoints2(0) = braidpoints(4):           braidpoints2(1) = braidpoints(5)
    braidpoints2(2) = braidpoints2(0) + j:      braidpoints2(3) = braidpoints2(1)
      If braidpoints2(3) = points(3) - 50 Then
      braidpoints2(2) = braidpoints2(2) - 40
             End If
      If braidpoints2(3) = points(3) - 30 Then
      braidpoints2(2) = braidpoints2(2) - 20
              End If
  
Set braidplineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(braidpoints2)
braidplineObj2.Layer = "Ball-6"
braidplineObj2.Update

'============================================================================= Vector 2, bottom


t = b / 240
If b >= 240 Then
t1 = Int(t)
Else: t1 = Round(t, 0)
End If
P = 20
q = 0
t = t1 + 1
j = (b / t) - 20
J1 = j + 20
g = t
S = 50
s1 = 40
P = 20

      
braidpoints(4) = points(0)
braidpoints(5) = points(1)


' Limitation of the cycle of construction of the ornament
Do While braidpoints(4) - points(0) < b - j - 30
  braidpoints(0) = braidpoints(4) + S:           braidpoints(1) = braidpoints(5) + S
  braidpoints(2) = braidpoints(0) + j - s1:      braidpoints(3) = braidpoints(1)
  P = 20
  s1 = 0

' Condition of the direction of deviation along X-axis for constructing crosshairs
If braidpoints(1) = points(1) + 50 Then
q = -20
End If
If braidpoints(1) = points(1) + 30 Then
q = 20
End If
   braidpoints(4) = braidpoints(2) + P: braidpoints(5) = braidpoints(3) + q

Set braidplineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(braidpoints)
braidplineObj.Layer = "Ball-6"
braidplineObj.Update

S = 0
q = 0
P = 20
Loop

' The condition for the end of the last segment of the ornament depending on its position along X-axis
    braidpoints2(0) = braidpoints(4):           braidpoints2(1) = braidpoints(5)
    braidpoints2(2) = braidpoints2(0) + j:      braidpoints2(3) = braidpoints2(1)
      If braidpoints2(3) = points(1) + 50 Then
      braidpoints2(2) = braidpoints2(2) - 40
             End If
      If braidpoints2(3) = points(1) + 30 Then
      braidpoints2(2) = braidpoints2(2) - 20
              End If
  
Set braidplineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(braidpoints2)
braidplineObj2.Layer = "Ball-6"
braidplineObj2.Update

'============================================================================= Vector 2, right

t = a / 240
If a >= 240 Then
t1 = Int(t)
Else: t1 = Round(t, 0)
End If
P = 20
q = 0
t = t1 + 1
j = (a / t) - 20
J1 = j + 20
g = t
S = 30
s1 = 40
P = 0

braidpoints(4) = braidpoints2(2)
braidpoints(3) = braidpoints2(3)


If braidpoints(1) = points(1) + 50 Then
s1 = 20
End If
If braidpoints(1) = points(1) + 30 Then
s1 = 40
End If


' Limitation of the cycle of construction of the ornament
Do While braidpoints(3) < a - j - 35

 
  braidpoints(0) = braidpoints(4):           braidpoints(1) = braidpoints(3) + P
  braidpoints(2) = braidpoints(0):           braidpoints(3) = braidpoints(1) + j - s1
  P = 20
  s1 = 0
' Condition of the direction of deviation along X-axis for constructing crosshairs
If braidpoints(2) <= points(4) - 45 Then
q = 20
End If
If braidpoints(2) >= points(4) - 35 Then
q = -20
End If
   braidpoints(4) = braidpoints(2) + q: braidpoints(5) = braidpoints(3) + P

Set braidplineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(braidpoints)
braidplineObj.Layer = "Ball-6"
braidplineObj.Update

S = 0
q = 0

P = 20

Loop
' The condition for the end of the last segment of the ornament depending on its position along X-axis
    braidpoints2(0) = braidpoints(4):         braidpoints2(1) = braidpoints(5)
    braidpoints2(2) = braidpoints2(0):        braidpoints2(3) = braidpoints2(1) + j
      If braidpoints2(2) <= points(4) - 45 Then
      braidpoints2(3) = a - 50
             End If
      If braidpoints2(2) >= points(4) - 35 Then
      braidpoints2(3) = a - 30
              End If
  
Set braidplineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(braidpoints2)
braidplineObj2.Layer = "Ball-6"
braidplineObj2.Update
End If
End If

'============================================================================= For quantity more than 1

I = 100
  
 Dim plineObj2 As AcadLWPolyline
  Dim points2(0 To 7) As Double
  
  points2(0) = points(0):    points2(1) = a + I:
  points2(2) = points(0):    points2(3) = (a * 2) + I:
  points2(4) = points(6):    points2(5) = (a * 2) + I:
  points2(6) = points(6):    points2(7) = a + I:
  
 For d = 2 To Cells(c, 4)
If d > 1 Then

Set plineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(points2)
  plineObj2.Closed = True
  
  '************************************************************
  
 '============================================================================= Vector 1, left
t = a / 240
If a >= 240 Then
t1 = Int(t)
Else: t1 = Round(t, 0)
End If
P = 20
q = 0
t = t1 + 1
j = a / t - 20
J1 = j + 20
g = t
S = 30
s1 = 20
P = 0

If a > 139 Then
If b > 139 Then
braidpoints(4) = points2(0)
braidpoints(3) = points2(1)
' Limitation of the cycle of construction of the ornament
Do While braidpoints(3) < points2(3) - j - 50


  braidpoints(0) = braidpoints(4) + S:       braidpoints(1) = braidpoints(3) + S + P
  braidpoints(2) = braidpoints(0):           braidpoints(3) = braidpoints(1) + j - s1
  P = 20
  s1 = 0
' Condition of the direction of deviation along X-axis for constructing crosshairs
If braidpoints(0) = points2(0) + 50 Then
q = -20
End If
If braidpoints(0) = points2(0) + 30 Then
q = 20
End If
   braidpoints(4) = braidpoints(2) + q: braidpoints(5) = braidpoints(3) + P

Set braidplineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(braidpoints)
braidplineObj.Layer = "Ball-6"
braidplineObj.Update

S = 0
q = 0

P = 20

Loop
' The condition for the end of the last segment of the ornament depending on its position along X-axis
    braidpoints2(0) = braidpoints(4):         braidpoints2(1) = braidpoints(5)
    braidpoints2(2) = braidpoints2(0):        braidpoints2(3) = braidpoints2(1) + j
      If braidpoints2(2) = points2(0) + 50 Then
      braidpoints2(3) = points2(3) - 50
             End If
      If braidpoints2(2) = points2(0) + 30 Then
      braidpoints2(3) = points2(3) - 30
              End If
  
Set braidplineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(braidpoints2)
braidplineObj2.Layer = "Ball-6"
braidplineObj2.Update

'============================================================================= Vector 1, top

t = b / 240
If b >= 240 Then
t1 = Int(t)
Else: t1 = Round(t, 0)
End If
P = 20
q = 0
t = t1 + 1
j = (b / t) - 20
J1 = j + 20
g = t
S = 30
s1 = 0
P = 20

      
braidpoints(4) = braidpoints2(2)
braidpoints(5) = braidpoints2(3)

If braidpoints(0) = points2(0) + 50 Then
s1 = 0
End If
If braidpoints(0) = points2(0) + 30 Then
s1 = 20
End If


' Limitation of the cycle of construction of the ornament
Do While braidpoints(4) - points2(0) < b - j - 30
  braidpoints(0) = braidpoints(4):                braidpoints(1) = braidpoints(5)
  braidpoints(2) = braidpoints(0) + j - P - s1:   braidpoints(3) = braidpoints(1)
  P = 20
  s1 = -20

' Condition of the direction of deviation along X-axis for constructing crosshairs
If braidpoints(1) = points2(3) - 50 Then
q = 20
End If
If braidpoints(1) = points2(3) - 30 Then
q = -20
End If
   braidpoints(4) = braidpoints(2) + P: braidpoints(5) = braidpoints(3) + q

Set braidplineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(braidpoints)
braidplineObj.Layer = "Ball-6"
braidplineObj.Update

S = 0
q = 0
P = 20
Loop

' The condition for the end of the last segment of the ornament depending on its position along X-axis
    braidpoints2(0) = braidpoints(4):           braidpoints2(1) = braidpoints(5)
    braidpoints2(2) = braidpoints2(0) + j:      braidpoints2(3) = braidpoints2(1)
      If braidpoints2(3) = points2(3) - 50 Then
      braidpoints2(2) = braidpoints2(2) - 40
             End If
      If braidpoints2(3) = points2(3) - 30 Then
      braidpoints2(2) = braidpoints2(2) - 20
              End If
  
Set braidplineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(braidpoints2)
braidplineObj2.Layer = "Ball-6"
braidplineObj2.Update

'============================================================================= = Vector 1, bottom


t = b / 240
If b >= 240 Then
t1 = Int(t)
Else: t1 = Round(t, 0)
End If
P = 20
q = 0
t = t1 + 1
j = (b / t) - 20
J1 = j + 20
g = t
S = 30
s1 = 20
P = 20

      
braidpoints(4) = points2(0)
braidpoints(5) = points2(1)


' Limitation of the cycle of construction of the ornament
Do While braidpoints(4) - points2(0) < b - j - 30
  braidpoints(0) = braidpoints(4) + S:           braidpoints(1) = braidpoints(5) + S
  braidpoints(2) = braidpoints(0) + j - s1:      braidpoints(3) = braidpoints(1)
  P = 20
  s1 = 0

' Condition of the direction of deviation along X-axis for constructing crosshairs
If braidpoints(1) = points2(1) + 50 Then
q = -20
End If
If braidpoints(1) = points2(1) + 30 Then
q = 20
End If
   braidpoints(4) = braidpoints(2) + P: braidpoints(5) = braidpoints(3) + q

Set braidplineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(braidpoints)
braidplineObj.Layer = "Ball-6"
braidplineObj.Update

S = 0
q = 0
P = 20
Loop

' The condition for the end of the last segment of the ornament depending on its position along X-axis
    braidpoints2(0) = braidpoints(4):           braidpoints2(1) = braidpoints(5)
    braidpoints2(2) = braidpoints2(0) + j:      braidpoints2(3) = braidpoints2(1)
      If braidpoints2(3) = points2(1) + 50 Then
      braidpoints2(2) = braidpoints2(2) - 40
             End If
      If braidpoints2(3) = points2(1) + 30 Then
      braidpoints2(2) = braidpoints2(2) - 20
              End If
  
Set braidplineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(braidpoints2)
braidplineObj2.Layer = "Ball-6"
braidplineObj2.Update

'============================================================================= Vector 1, right

t = a / 240
If a >= 240 Then
t1 = Int(t)
Else: t1 = Round(t, 0)
End If
P = 20
q = 0
t = t1 + 1
j = (a / t) - 20
J1 = j + 20
g = t
S = 30
s1 = 40
P = 0

braidpoints(4) = braidpoints2(2)
braidpoints(3) = braidpoints2(3)


If braidpoints(1) = points2(1) + 50 Then
s1 = 20
End If
If braidpoints(1) = points2(1) + 30 Then
s1 = 40
End If


' Limitation of the cycle of construction of the ornament
Do While braidpoints(3) < points2(3) - j - 35

 
  braidpoints(0) = braidpoints(4):           braidpoints(1) = braidpoints(3) + P
  braidpoints(2) = braidpoints(0):           braidpoints(3) = braidpoints(1) + j - s1
  P = 20
  s1 = 0
' Condition of the direction of deviation along X-axis for constructing crosshairs
If braidpoints(2) <= points2(4) - 45 Then
q = 20
End If
If braidpoints(2) >= points2(4) - 35 Then
q = -20
End If
   braidpoints(4) = braidpoints(2) + q: braidpoints(5) = braidpoints(3) + P

Set braidplineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(braidpoints)
braidplineObj.Layer = "Ball-6"
braidplineObj.Update

S = 0
q = 0

P = 20

Loop
' The condition for the end of the last segment of the ornament depending on its position along X-axis
    braidpoints2(0) = braidpoints(4):         braidpoints2(1) = braidpoints(5)
    braidpoints2(2) = braidpoints2(0):        braidpoints2(3) = braidpoints2(1) + j
      If braidpoints2(2) <= points2(4) - 45 Then
      braidpoints2(3) = points2(3) - 50
             End If
      If braidpoints2(2) >= points2(4) - 35 Then
      braidpoints2(3) = points2(3) - 30
              End If
  
Set braidplineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(braidpoints2)
braidplineObj2.Layer = "Ball-6"
braidplineObj2.Update

'============================================================================= Vector 2, left
t = a / 240
If a >= 240 Then
t1 = Int(t)
Else: t1 = Round(t, 0)
End If
P = 20
q = 0
t = t1 + 1
j = (a / t) - 20
J1 = j + 20
g = t
S = 50
s1 = 20
P = 0

braidpoints(4) = points2(0)
braidpoints(3) = points2(1)
' Limitation of the cycle of construction of the ornament
Do While braidpoints(3) < points2(3) - j - 50


  braidpoints(0) = braidpoints(4) + S:       braidpoints(1) = braidpoints(3) + S + P
  braidpoints(2) = braidpoints(0):           braidpoints(3) = braidpoints(1) + j - (s1 * 2)
  P = 20
  s1 = 0
' Condition of the direction of deviation along X-axis for constructing crosshairs
If braidpoints(0) = points2(0) + 50 Then
q = -20
End If
If braidpoints(0) = points2(0) + 30 Then
q = 20
End If
   braidpoints(4) = braidpoints(2) + q: braidpoints(5) = braidpoints(3) + P

Set braidplineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(braidpoints)
braidplineObj.Layer = "Ball-6"
braidplineObj.Update

S = 0
q = 0

P = 20

Loop
' The condition for the end of the last segment of the ornament depending on its position along X-axis
    braidpoints2(0) = braidpoints(4):         braidpoints2(1) = braidpoints(5)
    braidpoints2(2) = braidpoints2(0):        braidpoints2(3) = braidpoints2(1) + j
      If braidpoints2(2) = points2(0) + 50 Then
      braidpoints2(3) = points2(3) - 50
             End If
      If braidpoints2(2) = points2(0) + 30 Then
      braidpoints2(3) = points2(3) - 30
              End If
  
Set braidplineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(braidpoints2)
braidplineObj2.Layer = "Ball-6"
braidplineObj2.Update

'============================================================================= Vector 2, top

t = b / 240
If b >= 240 Then
t1 = Int(t)
Else: t1 = Round(t, 0)
End If
P = 20
q = 0
t = t1 + 1
j = (b / t) - 20
J1 = j + 20
g = t
S = 30
s1 = 0
P = 20

      
braidpoints(4) = braidpoints2(2)
braidpoints(5) = braidpoints2(3)

If braidpoints(0) = points2(0) + 50 Then
s1 = 0
End If
If braidpoints(0) = points2(0) + 30 Then
s1 = 20
End If


' Limitation of the cycle of construction of the ornament
Do While braidpoints(4) - points2(0) < b - j - 30
  braidpoints(0) = braidpoints(4):                braidpoints(1) = braidpoints(5)
  braidpoints(2) = braidpoints(0) + j - P - s1:   braidpoints(3) = braidpoints(1)
  P = 20
  s1 = -20

' Condition of the direction of deviation along X-axis for constructing crosshairs
If braidpoints(1) = points2(3) - 50 Then
q = 20
End If
If braidpoints(1) = points2(3) - 30 Then
q = -20
End If
   braidpoints(4) = braidpoints(2) + P: braidpoints(5) = braidpoints(3) + q

Set braidplineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(braidpoints)
braidplineObj.Layer = "Ball-6"
braidplineObj.Update

S = 0
q = 0
P = 20
Loop

' The condition for the end of the last segment of the ornament depending on its position along X-axis
    braidpoints2(0) = braidpoints(4):           braidpoints2(1) = braidpoints(5)
    braidpoints2(2) = braidpoints2(0) + j:      braidpoints2(3) = braidpoints2(1)
      If braidpoints2(3) = points2(3) - 50 Then
      braidpoints2(2) = braidpoints2(2) - 40
             End If
      If braidpoints2(3) = points2(3) - 30 Then
      braidpoints2(2) = braidpoints2(2) - 20
              End If
  
Set braidplineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(braidpoints2)
braidplineObj2.Layer = "Ball-6"
braidplineObj2.Update

'============================================================================= Vector 2, bottom


t = b / 240
If b >= 240 Then
t1 = Int(t)
Else: t1 = Round(t, 0)
End If
P = 20
q = 0
t = t1 + 1
j = (b / t) - 20
J1 = j + 20
g = t
S = 50
s1 = 40
P = 20

      
braidpoints(4) = points2(0)
braidpoints(5) = points2(1)


' Limitation of the cycle of construction of the ornament
Do While braidpoints(4) - points2(0) < b - j - 30
  braidpoints(0) = braidpoints(4) + S:           braidpoints(1) = braidpoints(5) + S
  braidpoints(2) = braidpoints(0) + j - s1:      braidpoints(3) = braidpoints(1)
  P = 20
  s1 = 0

' Condition of the direction of deviation along X-axis for constructing crosshairs
If braidpoints(1) = points2(1) + 50 Then
q = -20
End If
If braidpoints(1) = points2(1) + 30 Then
q = 20
End If
   braidpoints(4) = braidpoints(2) + P: braidpoints(5) = braidpoints(3) + q

Set braidplineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(braidpoints)
braidplineObj.Layer = "Ball-6"
braidplineObj.Update

S = 0
q = 0
P = 20
Loop

' The condition for the end of the last segment of the ornament depending on its position along X-axis
    braidpoints2(0) = braidpoints(4):           braidpoints2(1) = braidpoints(5)
    braidpoints2(2) = braidpoints2(0) + j:      braidpoints2(3) = braidpoints2(1)
      If braidpoints2(3) = points2(1) + 50 Then
      braidpoints2(2) = braidpoints2(2) - 40
             End If
      If braidpoints2(3) = points2(1) + 30 Then
      braidpoints2(2) = braidpoints2(2) - 20
              End If
  
Set braidplineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(braidpoints2)
braidplineObj2.Layer = "Ball-6"
braidplineObj2.Update

'============================================================================= Vector 2, right

t = a / 240
If a >= 240 Then
t1 = Int(t)
Else: t1 = Round(t, 0)
End If
P = 20
q = 0
t = t1 + 1
j = (a / t) - 20
J1 = j + 20
g = t
S = 30
s1 = 40
P = 0

braidpoints(4) = braidpoints2(2)
braidpoints(3) = braidpoints2(3)


If braidpoints(1) = points2(1) + 50 Then
s1 = 20
End If
If braidpoints(1) = points2(1) + 30 Then
s1 = 40
End If


' Limitation of the cycle of construction of the ornament
Do While braidpoints(3) < points2(3) - j - 35

 
  braidpoints(0) = braidpoints(4):           braidpoints(1) = braidpoints(3) + P
  braidpoints(2) = braidpoints(0):           braidpoints(3) = braidpoints(1) + j - s1
  P = 20
  s1 = 0
' Condition of the direction of deviation along X-axis for constructing crosshairs
If braidpoints(2) <= points2(4) - 45 Then
q = 20
End If
If braidpoints(2) >= points2(4) - 35 Then
q = -20
End If
   braidpoints(4) = braidpoints(2) + q: braidpoints(5) = braidpoints(3) + P

Set braidplineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(braidpoints)
braidplineObj.Layer = "Ball-6"
braidplineObj.Update

S = 0
q = 0

P = 20

Loop
' The condition for the end of the last segment of the ornament depending on its position along X-axis
    braidpoints2(0) = braidpoints(4):         braidpoints2(1) = braidpoints(5)
    braidpoints2(2) = braidpoints2(0):        braidpoints2(3) = braidpoints2(1) + j
      If braidpoints2(2) <= points2(4) - 45 Then
      braidpoints2(3) = points2(3) - 50
             End If
      If braidpoints2(2) >= points2(4) - 35 Then
      braidpoints2(3) = points2(3) - 30
              End If
  
Set braidplineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(braidpoints2)
braidplineObj2.Layer = "Ball-6"
braidplineObj2.Update
End If
End If

  '************************************************************
  points2(1) = points2(3) + I:
  points2(3) = points2(3) + a + I:
  points2(5) = points2(3):
  points2(7) = points2(1):
  
Else: d = 1
d = d + 1

d = Cells(c, 4)
End If

Next d
c = c + 1

a = Cells(c, 2)
b = Cells(c, 3)

I = 100

Loop
              
ZoomAll
AP.Quit
End Sub

Sub MF110()
' Creating search objects
Dim AP As Excel.Application
Dim WB As Excel.Workbook
Dim WS As Excel.Worksheet

Set AP = Excel.Application
Set WB = AP.Workbooks.Open("C:\Users\������\Desktop\������\2.xlsx")
Set WS = WB.Worksheets("������ �˨������ ������")
c = 19
a = Cells(c, 2)
b = Cells(c, 3)
d = Cells(c, 4)
E = 1
I = 0
On Error Resume Next

' Description of all the points that will make up rectangles
  Dim plineObj As AcadLWPolyline
  Dim points(0 To 7) As Double
points(6) = 0

Do While a > 0

' Calculate the coordinates of each point
  points(0) = points(6) + I:    points(1) = 0
  points(2) = points(6) + I:    points(3) = a
  points(4) = points(0) + b:    points(5) = a
  points(6) = points(0) + b:    points(7) = 0

' Constructing four segments and connecting the corresponding points

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
  plineObj.Closed = True
 
Dim a1(0 To 2) As Double
Dim a2(0 To 2) As Double
Dim A3(0 To 2) As Double
Dim A4(0 To 2) As Double
Dim A5(0 To 2) As Double
Dim A6(0 To 2) As Double
Dim A7(0 To 2) As Double
Dim A8(0 To 2) As Double
Dim A9(0 To 2) As Double
Dim A10(0 To 2) As Double
Dim A11(0 To 2) As Double
Dim A12(0 To 2) As Double
Dim A13(0 To 2) As Double
Dim A14(0 To 2) As Double
Dim A15(0 To 2) As Double
Dim A16(0 To 2) As Double
Dim lineObj As AcadLine
  
  a1(0) = points(0) + 60: a1(1) = 0:      a1(2) = 0
  a2(0) = points(2) + 60: a2(1) = a:      a2(2) = 0
  A3(0) = points(4) - 60: A3(1) = a:      A3(2) = 0
  A4(0) = points(6) - 60: A4(1) = 0:      A4(2) = 0
  
  A5(0) = points(0) + 60: A5(1) = 60:      A5(2) = 0
  A6(0) = points(2) + 60: A6(1) = a - 60:  A6(2) = 0
  A7(0) = points(4) - 60: A7(1) = a - 60:  A7(2) = 0
  A8(0) = points(6) - 60: A8(1) = 60:      A8(2) = 0
   
  A9(0) = points(0) + 60: A9(1) = 30:        A9(2) = 0
  A10(0) = points(2) + 60: A10(1) = a - 30:  A10(2) = 0
  A11(0) = points(4) - 60: A11(1) = a - 30:  A11(2) = 0
  A12(0) = points(6) - 60: A12(1) = 30:      A12(2) = 0
  
  A13(0) = points(0) + 60: A13(1) = 90:      A13(2) = 0
  A14(0) = points(2) + 60: A14(1) = a - 90:  A14(2) = 0
  A15(0) = points(4) - 60: A15(1) = a - 90:  A15(2) = 0
  A16(0) = points(6) - 60: A16(1) = 90:      A16(2) = 0
   
   
If a > 100 Then
If b > 100 Then
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(a1, a2)
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A3, A4)
lineObj.Layer = "Ball-6"
lineObj.Update

If a >= 150 Then
If b >= 100 Then
Set lineObj = ThisDrawing.ModelSpace.AddLine(A6, A7)
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A8, A5)
lineObj.Layer = "Ball-6"
lineObj.Update
End If
End If

Set lineObj = ThisDrawing.ModelSpace.AddLine(A10, A11)
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A9, A12)
lineObj.Layer = "Ball-6"
lineObj.Update

If a >= 210 Then
If b >= 100 Then
Set lineObj = ThisDrawing.ModelSpace.AddLine(A14, A15)
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A13, A16)
lineObj.Layer = "Ball-6"
lineObj.Update
End If
End If

End If
End If

I = 100
  
 Dim plineObj2 As AcadLWPolyline
  Dim points2(0 To 7) As Double
  
  points2(0) = points(0):    points2(1) = a + I:
  points2(2) = points(0):    points2(3) = (a * 2) + I:
  points2(4) = points(6):    points2(5) = (a * 2) + I:
  points2(6) = points(6):    points2(7) = a + I:
  
 For d = 2 To Cells(c, 4)
If d > 1 Then

Set plineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(points2)
  plineObj2.Closed = True
If a > 100 Then
If b > 100 Then
  
  a1(0) = points2(0) + 60: a1(1) = points2(1):     a1(2) = 0
  a2(0) = points2(2) + 60: a2(1) = points2(3):     a2(2) = 0
  A3(0) = points2(4) - 60: A3(1) = points2(5):     A3(2) = 0
  A4(0) = points2(6) - 60: A4(1) = points2(7):     A4(2) = 0
  
  A5(0) = points2(0) + 60: A5(1) = points2(1) + 60:    A5(2) = 0
  A6(0) = points2(2) + 60: A6(1) = points2(3) - 60:    A6(2) = 0
  A7(0) = points2(4) - 60: A7(1) = points2(5) - 60:    A7(2) = 0
  A8(0) = points2(6) - 60: A8(1) = points2(7) + 60:    A8(2) = 0
  
  A9(0) = points2(0) + 60:   A9(1) = points2(1) + 30:       A9(2) = 0
  A10(0) = points2(2) + 60: A10(1) = points2(3) - 30:      A10(2) = 0
  A11(0) = points2(4) - 60: A11(1) = points2(5) - 30:      A11(2) = 0
  A12(0) = points2(6) - 60: A12(1) = points2(7) + 30:      A12(2) = 0
  
  A13(0) = points2(0) + 60: A13(1) = points2(1) + 90:      A13(2) = 0
  A14(0) = points2(2) + 60: A14(1) = points2(3) - 90:      A14(2) = 0
  A15(0) = points2(4) - 60: A15(1) = points2(5) - 90:      A15(2) = 0
  A16(0) = points2(6) - 60: A16(1) = points2(7) + 90:      A16(2) = 0
  
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(a1, a2)
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A3, A4)
lineObj.Layer = "Ball-6"
lineObj.Update

If a >= 150 Then
If b >= 100 Then
Set lineObj = ThisDrawing.ModelSpace.AddLine(A5, A8)
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A6, A7)
lineObj.Layer = "Ball-6"
lineObj.Update
End If
End If

Set lineObj = ThisDrawing.ModelSpace.AddLine(A10, A11)
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A9, A12)
lineObj.Layer = "Ball-6"
lineObj.Update

If a >= 210 Then
If b >= 100 Then
Set lineObj = ThisDrawing.ModelSpace.AddLine(A14, A15)
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A13, A16)
lineObj.Layer = "Ball-6"
lineObj.Update
End If
End If

End If
End If

  points2(1) = points2(3) + I:
  points2(3) = points2(3) + a + I:
  points2(5) = points2(3):
  points2(7) = points2(1):
 
Else: d = 1
d = d + 1

d = Cells(c, 4)
End If

Next d
c = c + 1

a = Cells(c, 2)
b = Cells(c, 3)

I = 100

Loop
              
ZoomAll
AP.Quit
End Sub

Sub MF112()
' Creating search objects
Dim AP As Excel.Application
Dim WB As Excel.Workbook
Dim WS As Excel.Worksheet

Set AP = Excel.Application
Set WB = AP.Workbooks.Open("C:\Users\������\Desktop\������\2.xlsx")
Set WS = WB.Worksheets("������ �˨������ ������")
c = 19
a = Cells(c, 2)
b = Cells(c, 3)
d = Cells(c, 4)
E = 1
I = 0
On Error Resume Next

' Description of all the points that will make up rectangles
 Dim plineObj As AcadLWPolyline
  Dim points(0 To 7) As Double
  Dim pointsb12(0 To 13) As Double
  Dim pointsb122(0 To 13) As Double
points(6) = 0

Do While a > 0

' Calculate the coordinates of each point
  points(0) = points(6) + I:    points(1) = 0
  points(2) = points(6) + I:    points(3) = a
  points(4) = points(0) + b:    points(5) = a
  points(6) = points(0) + b:    points(7) = 0

' Constructing four segments and connecting the corresponding points

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
  plineObj.Closed = True
  
  ' Offset the polyline
Dim offsetObj As Variant
                                 'plineObj.color = acByLayer
                                 'Dim layerObj As AcadLayer
                                 'Set layerObj = ThisDrawing.Layers.Add("Ball-6.1")
                                 'layerObj.color = acRed
                               
pointsb12(13) = 0
Do While pointsb12(13) < a

  pointsb12(0) = points(0) + 50:    pointsb12(1) = points(1) + 0
  
  pointsb12(2) = points(0) + 50:    pointsb12(3) = points(1) + 50
  pointsb12(4) = points(0) + 75:    pointsb12(5) = points(1) + 50
  pointsb12(6) = points(0) + 75:    pointsb12(7) = points(1) + 25
  pointsb12(8) = points(0) + 100:   pointsb12(9) = points(1) + 25
  
  pointsb12(10) = points(0) + 100:  pointsb12(11) = points(1) + 75
  pointsb12(12) = points(0) + 50:   pointsb12(13) = points(1) + 75
  
points(1) = points(1) + 75
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsb12)
plineObj.Layer = "Ball-12"
plineObj.Update

Loop

I = 100
  
 Dim plineObj2 As AcadLWPolyline
  Dim points2(0 To 7) As Double
  
  points2(0) = points(0):    points2(1) = a + I:
  points2(2) = points(0):    points2(3) = (a * 2) + I:
  points2(4) = points(6):    points2(5) = (a * 2) + I:
  points2(6) = points(6):    points2(7) = a + I:
  
 For d = 2 To Cells(c, 4)
If d > 1 Then

Set plineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(points2)
  plineObj2.Closed = True
  
 pointsb122(13) = points2(1)
Do While pointsb122(13) < (a * d) + ((I * d) - I)

  pointsb122(0) = points2(0) + 50:    pointsb122(1) = points2(1) + 0
  
  pointsb122(2) = points2(0) + 50:    pointsb122(3) = points2(1) + 50
  pointsb122(4) = points2(0) + 75:    pointsb122(5) = points2(1) + 50
  pointsb122(6) = points2(0) + 75:    pointsb122(7) = points2(1) + 25
  pointsb122(8) = points2(0) + 100:   pointsb122(9) = points2(1) + 25
  
  pointsb122(10) = points2(0) + 100:  pointsb122(11) = points2(1) + 75
  pointsb122(12) = points2(0) + 50:   pointsb122(13) = points2(1) + 75
  
points2(1) = points2(1) + 75
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsb122)
plineObj.Layer = "Ball-12"
plineObj.Update

Loop
  
  
  points2(1) = points2(3) + I:
  points2(3) = points2(3) + a + I:
  points2(5) = points2(3):
  points2(7) = points2(1):

Else: d = 1
d = d + 1

d = Cells(c, 4)
End If

Next d
c = c + 1

a = Cells(c, 2)
b = Cells(c, 3)

I = 100

Loop
              
ZoomAll
AP.Quit
End Sub

Sub MF113()
' Creating search objects
Dim AP As Excel.Application
Dim WB As Excel.Workbook
Dim WS As Excel.Worksheet

Set AP = Excel.Application
Set WB = AP.Workbooks.Open("C:\Users\������\Desktop\������\2.xlsx")
Set WS = WB.Worksheets("������ �˨������ ������")
c = 19
a = Cells(c, 2)
b = Cells(c, 3)
d = Cells(c, 4)
E = 1
I = 0
On Error Resume Next

' Description of all the points that will make up rectangles
  Dim plineObj As AcadLWPolyline
  Dim points(0 To 7) As Double
  Dim pointswithin(0 To 7) As Double
  Dim pointswithin2(0 To 7) As Double
points(6) = 0

Do While a > 0

' Calculate the coordinates of each point
  points(0) = points(6) + I:    points(1) = 0
  points(2) = points(6) + I:    points(3) = a
  points(4) = points(0) + b:    points(5) = a
  points(6) = points(0) + b:    points(7) = 0

' Constructing four segments and connecting the corresponding points

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
  plineObj.Closed = True
 
  pointswithin(0) = points(0) + 66:    pointswithin(1) = 66
  pointswithin(2) = points(2) + 66:    pointswithin(3) = a - 90
  pointswithin(4) = points(4) - 66:    pointswithin(5) = a - 90
  pointswithin(6) = points(6) - 66:    pointswithin(7) = 66
    
If a > 180 Then
If b > 180 Then
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointswithin)
  ' Find the bulge of the third segment
    Dim currentBulge As Double
    currentBulge = plineObj.GetBulge(2)
   k = (pointswithin(4) - pointswithin(2)) / 2
    ' Change the bulge of the third segment
    plineObj.SetBulge 1, -(24 / k)
    plineObj.Update
    plineObj.Layer = "C-Mill"
    plineObj.Update
  plineObj.Closed = True
 End If
 End If
                                 'plineObj.color = acByLayer
                                 'Dim layerObj As AcadLayer
                                 'Set layerObj = ThisDrawing.Layers.Add("Ball-6.1")
                                 'layerObj.color = acRed
   
I = 100
  
 Dim plineObj2 As AcadLWPolyline
  Dim points2(0 To 7) As Double
  
  points2(0) = points(0):    points2(1) = a + I:
  points2(2) = points(0):    points2(3) = (a * 2) + I:
  points2(4) = points(6):    points2(5) = (a * 2) + I:
  points2(6) = points(6):    points2(7) = a + I:
  
 For d = 2 To Cells(c, 4)
If d > 1 Then

Set plineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(points2)
  plineObj2.Closed = True

  pointswithin2(0) = points2(0) + 66:    pointswithin2(1) = points2(1) + 66
  pointswithin2(2) = points2(2) + 66:    pointswithin2(3) = points2(3) - 90
  pointswithin2(4) = points2(4) - 66:    pointswithin2(5) = points2(5) - 90
  pointswithin2(6) = points2(6) - 66:    pointswithin2(7) = points2(7) + 66

If a > 180 Then
If b > 180 Then
  ' Find the bulge of the third segment
  Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointswithin2)
    currentBulge = plineObj.GetBulge(2)
   k = (pointswithin(4) - pointswithin(2)) / 2
    ' Change the bulge of the third segment
    plineObj.SetBulge 1, -(24 / k)
    plineObj.Update
    plineObj.Layer = "C-Mill"
    plineObj.Update
  plineObj.Closed = True
End If
End If
  points2(1) = points2(3) + I:
  points2(3) = points2(3) + a + I:
  points2(5) = points2(3):
  points2(7) = points2(1):
      
Else: d = 1
d = d + 1

d = Cells(c, 4)
End If

Next d
c = c + 1

a = Cells(c, 2)
b = Cells(c, 3)

I = 100

Loop
              
ZoomAll
AP.Quit
End Sub

Sub MF114()
' Creating search objects
Dim AP As Excel.Application
Dim WB As Excel.Workbook
Dim WS As Excel.Worksheet

Set AP = Excel.Application
Set WB = AP.Workbooks.Open("C:\Users\������\Desktop\������\2.xlsx")
Set WS = WB.Worksheets("������ �˨������ ������")
c = 19
a = Cells(c, 2)
b = Cells(c, 3)
d = Cells(c, 4)
E = 1
I = 0
On Error Resume Next

' Description of all the points that will make up rectangles
  Dim plineObj As AcadLWPolyline
  Dim points(0 To 7) As Double
  Dim pointswithin(0 To 7) As Double
  Dim pointswithin2(0 To 7) As Double
points(6) = 0

Do While a > 0

' Calculate the coordinates of each point
  points(0) = points(6) + I:    points(1) = 0
  points(2) = points(6) + I:    points(3) = a
  points(4) = points(0) + b:    points(5) = a
  points(6) = points(0) + b:    points(7) = 0

' Constructing four segments and connecting the corresponding points

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
  plineObj.Closed = True
 
  pointswithin(0) = points(0) + 66:    pointswithin(1) = 90
  pointswithin(2) = points(2) + 66:    pointswithin(3) = a - 90
  pointswithin(4) = points(4) - 66:    pointswithin(5) = a - 90
  pointswithin(6) = points(6) - 66:    pointswithin(7) = 90
If a > 180 Then
If b > 180 Then
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointswithin)
  ' Find the bulge of the third segment
    Dim currentBulge As Double
    currentBulge = plineObj.GetBulge(2)
   k = (pointswithin(4) - pointswithin(2)) / 2
    ' Change the bulge of the third segment
    plineObj.SetBulge 1, -(24 / k)
    plineObj.Update
    plineObj.SetBulge 3, -(24 / k)
    plineObj.Update
    plineObj.Layer = "C-Mill"
    plineObj.Update
  plineObj.Closed = True
 End If
 End If
                                 'plineObj.color = acByLayer
                                 'Dim layerObj As AcadLayer
                                 'Set layerObj = ThisDrawing.Layers.Add("Ball-6.1")
                                 'layerObj.color = acRed
   
I = 100
  
 Dim plineObj2 As AcadLWPolyline
  Dim points2(0 To 7) As Double
  
  points2(0) = points(0):    points2(1) = a + I:
  points2(2) = points(0):    points2(3) = (a * 2) + I:
  points2(4) = points(6):    points2(5) = (a * 2) + I:
  points2(6) = points(6):    points2(7) = a + I:
  
 For d = 2 To Cells(c, 4)
If d > 1 Then

Set plineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(points2)
  plineObj2.Closed = True

  pointswithin2(0) = points2(0) + 66:    pointswithin2(1) = points2(1) + 90
  pointswithin2(2) = points2(2) + 66:    pointswithin2(3) = points2(3) - 90
  pointswithin2(4) = points2(4) - 66:    pointswithin2(5) = points2(5) - 90
  pointswithin2(6) = points2(6) - 66:    pointswithin2(7) = points2(7) + 90
If a > 180 Then
If b > 180 Then
  ' Find the bulge of the third segment
  Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointswithin2)
    currentBulge = plineObj.GetBulge(2)
   k = (pointswithin(4) - pointswithin(2)) / 2
    ' Change the bulge of the third segment
    plineObj.SetBulge 1, -(24 / k)
    plineObj.Update
    plineObj.SetBulge 3, -(24 / k)
    plineObj.Update
    plineObj.Layer = "C-Mill"
    plineObj.Update
  plineObj.Closed = True
 End If
 End If
  points2(1) = points2(3) + I:
  points2(3) = points2(3) + a + I:
  points2(5) = points2(3):
  points2(7) = points2(1):
      
Else: d = 1
d = d + 1

d = Cells(c, 4)
End If

Next d
c = c + 1

a = Cells(c, 2)
b = Cells(c, 3)

I = 100

Loop
              
ZoomAll
AP.Quit
End Sub

Sub MF115()
' Creating search objects
Dim AP As Excel.Application
Dim WB As Excel.Workbook
Dim WS As Excel.Worksheet

Set AP = Excel.Application
Set WB = AP.Workbooks.Open("C:\Users\������\Desktop\������\2.xlsx")
Set WS = WB.Worksheets("������ �˨������ ������")
c = 19
a = Cells(c, 2)
b = Cells(c, 3)
d = Cells(c, 4)
E = 1
I = 0
On Error Resume Next

' Description of all the points that will make up rectangles
  Dim plineObj As AcadLWPolyline
  Dim points(0 To 7) As Double
  Dim pointswithin(0 To 13) As Double
  Dim pointswithin2(0 To 13) As Double
  Dim P As Variant
  
points(6) = 0

Do While a > 0

' Calculate the coordinates of each point
  points(0) = points(6) + I:    points(1) = 0
  points(2) = points(6) + I:    points(3) = a
  points(4) = points(0) + b:    points(5) = a
  points(6) = points(0) + b:    points(7) = 0

' Constructing four segments and connecting the corresponding points

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
  plineObj.Closed = True
  P = (b - 132) / 4
  pointswithin(0) = points(0) + 66:           pointswithin(1) = points(1) + 66
  pointswithin(2) = points(2) + 66:           pointswithin(3) = points(3) - 90
  pointswithin(4) = points(2) + 66 + P:       pointswithin(5) = points(3) - 78
  pointswithin(6) = points(2) + 66 + (2 * P): pointswithin(7) = points(3) - 66
  pointswithin(8) = points(4) - 66 - P:       pointswithin(9) = points(3) - 78
  pointswithin(10) = points(4) - 66:          pointswithin(11) = points(3) - 90
  pointswithin(12) = points(4) - 66:          pointswithin(13) = points(7) + 66
  
If a > 180 Then
If b > 180 Then
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointswithin)
  ' Find the bulge of the third segment
    Dim currentBulge As Double
    currentBulge = plineObj.GetBulge(2)
   g = Sqr((P * P) + 144)
   angle = Atn((12 / g) / Sqr((-12 / g) * (12 / g) + 1))
   radius = (g / 2) / Sin(12 / g)
   h = radius * (1 - Cos(angle))
   k = h / g
    ' Change the bulge of the third segment
    plineObj.SetBulge 1, k * 2
    plineObj.Update
    plineObj.SetBulge 2, -k * 2
    plineObj.Update
    plineObj.SetBulge 3, -k * 2
    plineObj.Update
    plineObj.SetBulge 4, k * 2
    plineObj.Update
    plineObj.Layer = "C-Mill"
    plineObj.Update
  plineObj.Closed = True
 
 End If
 End If
                                 'plineObj.color = acByLayer
                                 'Dim layerObj As AcadLayer
                                 'Set layerObj = ThisDrawing.Layers.Add("Ball-6.1")
                                 'layerObj.color = acRed
   
I = 100
  
 Dim plineObj2 As AcadLWPolyline
  Dim points2(0 To 7) As Double
  
  points2(0) = points(0):    points2(1) = a + I:
  points2(2) = points(0):    points2(3) = (a * 2) + I:
  points2(4) = points(6):    points2(5) = (a * 2) + I:
  points2(6) = points(6):    points2(7) = a + I:
  
 For d = 2 To Cells(c, 4)
If d > 1 Then

Set plineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(points2)
  plineObj2.Closed = True

  pointswithin2(0) = points2(0) + 66:             pointswithin2(1) = points2(1) + 66
  pointswithin2(2) = points2(2) + 66:             pointswithin2(3) = points2(3) - 90
  pointswithin2(4) = points2(2) + 66 + P:         pointswithin2(5) = points2(3) - 78
  pointswithin2(6) = points2(2) + 66 + (2 * P):   pointswithin2(7) = points2(3) - 66
  pointswithin2(8) = points2(4) - 66 - P:         pointswithin2(9) = points2(3) - 78
  pointswithin2(10) = points2(4) - 66:            pointswithin2(11) = points2(3) - 90
  pointswithin2(12) = points2(4) - 66:            pointswithin2(13) = points2(7) + 66
  
If a > 180 Then
If b > 180 Then
  ' Find the bulge of the third segment
  Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointswithin2)
    currentBulge = plineObj.GetBulge(2)
    ' Change the bulge of the third segment
    plineObj.SetBulge 1, k * 2
    plineObj.Update
    plineObj.SetBulge 2, -k * 2
    plineObj.Update
    plineObj.SetBulge 3, -k * 2
    plineObj.Update
    plineObj.SetBulge 4, k * 2
    plineObj.Update
    plineObj.Layer = "C-Mill"
    plineObj.Update
  plineObj.Closed = True
 
 End If
 End If
  
  points2(1) = points2(3) + I:
  points2(3) = points2(3) + a + I:
  points2(5) = points2(3):
  points2(7) = points2(1):
      
Else: d = 1
d = d + 1

d = Cells(c, 4)
End If

Next d
c = c + 1

a = Cells(c, 2)
b = Cells(c, 3)

I = 100

Loop
              
ZoomAll
AP.Quit
End Sub

Sub MF116()
' Creating search objects
Dim AP As Excel.Application
Dim WB As Excel.Workbook
Dim WS As Excel.Worksheet

Set AP = Excel.Application
Set WB = AP.Workbooks.Open("C:\Users\������\Desktop\������\2.xlsx")
Set WS = WB.Worksheets("������ �˨������ ������")
c = 19
a = Cells(c, 2)
b = Cells(c, 3)
d = Cells(c, 4)
E = 1
I = 0
On Error Resume Next

' Description of all the points that will make up rectangles
  Dim plineObj As AcadLWPolyline
  Dim points(0 To 7) As Double
  Dim pointswithin(0 To 19) As Double
  Dim pointswithin2(0 To 19) As Double
  Dim P As Variant
  
points(6) = 0

Do While a > 0

' Calculate the coordinates of each point
  points(0) = points(6) + I:    points(1) = 0
  points(2) = points(6) + I:    points(3) = a
  points(4) = points(0) + b:    points(5) = a
  points(6) = points(0) + b:    points(7) = 0

' Constructing four segments and connecting the corresponding points

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
  plineObj.Closed = True
  P = (b - 132) / 4
  pointswithin(0) = points(0) + 66:            pointswithin(1) = points(1) + 90
  pointswithin(2) = points(2) + 66:            pointswithin(3) = points(3) - 90
  pointswithin(4) = points(2) + 66 + P:        pointswithin(5) = points(3) - 78
  pointswithin(6) = points(2) + 66 + (2 * P):  pointswithin(7) = points(3) - 66
  pointswithin(8) = points(4) - 66 - P:        pointswithin(9) = points(3) - 78
  pointswithin(10) = points(4) - 66:           pointswithin(11) = points(3) - 90
  pointswithin(12) = points(4) - 66:           pointswithin(13) = points(7) + 90
  pointswithin(14) = points(4) - 66 - P:       pointswithin(15) = points(7) + 78
  pointswithin(16) = points(4) - 66 - (2 * P): pointswithin(17) = points(7) + 66
  pointswithin(18) = points(0) + 66 + P:       pointswithin(19) = points(7) + 78
  
If a > 180 Then
If b > 180 Then
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointswithin)
  ' Find the bulge of the third segment
    Dim currentBulge As Double
    currentBulge = plineObj.GetBulge(2)
   g = Sqr((P * P) + 144)
   angle = Atn((12 / g) / Sqr((-12 / g) * (12 / g) + 1))
   radius = (g / 2) / Sin(12 / g)
   h = radius * (1 - Cos(angle))
   k = h / g
    ' Change the bulge of the third segment
    plineObj.SetBulge 1, k * 2
    plineObj.Update
    plineObj.SetBulge 2, -k * 2
    plineObj.Update
    plineObj.SetBulge 3, -k * 2
    plineObj.Update
    plineObj.SetBulge 4, k * 2
    plineObj.Update
    plineObj.SetBulge 6, k * 2
    plineObj.Update
    plineObj.SetBulge 7, -k * 2
    plineObj.Update
    plineObj.SetBulge 8, -k * 2
    plineObj.Update
    plineObj.SetBulge 9, k * 2
    plineObj.Update
    plineObj.Layer = "C-Mill"
    plineObj.Update
  plineObj.Closed = True
End If
End If
                                 'plineObj.color = acByLayer
                                 'Dim layerObj As AcadLayer
                                 'Set layerObj = ThisDrawing.Layers.Add("Ball-6.1")
                                 'layerObj.color = acRed
   
I = 100
  
 Dim plineObj2 As AcadLWPolyline
  Dim points2(0 To 7) As Double
  
  points2(0) = points(0):    points2(1) = a + I:
  points2(2) = points(0):    points2(3) = (a * 2) + I:
  points2(4) = points(6):    points2(5) = (a * 2) + I:
  points2(6) = points(6):    points2(7) = a + I:
  
 For d = 2 To Cells(c, 4)
If d > 1 Then

Set plineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(points2)
  plineObj2.Closed = True

  pointswithin2(0) = points2(0) + 66:              pointswithin2(1) = points2(1) + 90
  pointswithin2(2) = points2(2) + 66:              pointswithin2(3) = points2(3) - 90
  pointswithin2(4) = points2(2) + 66 + P:          pointswithin2(5) = points2(3) - 78
  pointswithin2(6) = points2(2) + 66 + (2 * P):    pointswithin2(7) = points2(3) - 66
  pointswithin2(8) = points2(4) - 66 - P:          pointswithin2(9) = points2(3) - 78
  pointswithin2(10) = points2(4) - 66:             pointswithin2(11) = points2(3) - 90
  pointswithin2(12) = points2(4) - 66:             pointswithin2(13) = points2(7) + 90
  pointswithin2(14) = points2(4) - 66 - P:         pointswithin2(15) = points2(7) + 78
  pointswithin2(16) = points2(4) - 66 - (2 * P):   pointswithin2(17) = points2(7) + 66
  pointswithin2(18) = points2(0) + 66 + P:         pointswithin2(19) = points2(7) + 78
  
If a > 180 Then
If b > 180 Then
  ' Find the bulge of the third segment
  Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointswithin2)
    currentBulge = plineObj.GetBulge(2)
    ' Change the bulge of the third segment
    plineObj.SetBulge 1, k * 2
    plineObj.Update
    plineObj.SetBulge 2, -k * 2
    plineObj.Update
    plineObj.SetBulge 3, -k * 2
    plineObj.Update
    plineObj.SetBulge 4, k * 2
    plineObj.Update
    plineObj.SetBulge 6, k * 2
    plineObj.Update
    plineObj.SetBulge 7, -k * 2
    plineObj.Update
    plineObj.SetBulge 8, -k * 2
    plineObj.Update
    plineObj.SetBulge 9, k * 2
    plineObj.Update
    plineObj.Layer = "C-Mill"
    plineObj.Update
  plineObj.Closed = True
 End If
 End If
 
  points2(1) = points2(3) + I:
  points2(3) = points2(3) + a + I:
  points2(5) = points2(3):
  points2(7) = points2(1):
      
Else: d = 1
d = d + 1

d = Cells(c, 4)
End If

Next d
c = c + 1

a = Cells(c, 2)
b = Cells(c, 3)

I = 100

Loop
              
ZoomAll
AP.Quit
End Sub

Sub MF117()
' Creating search objects
Dim AP As Excel.Application
Dim WB As Excel.Workbook
Dim WS As Excel.Worksheet

Set AP = Excel.Application
Set WB = AP.Workbooks.Open("C:\Users\������\Desktop\������\2.xlsx")
Set WS = WB.Worksheets("������ �˨������ ������")
c = 19
a = Cells(c, 2)
b = Cells(c, 3)
d = Cells(c, 4)
E = 1
I = 0
On Error Resume Next

' Description of all the points that will make up rectangles
 Dim plineObj As AcadLWPolyline
  Dim points(0 To 7) As Double
points(6) = 0

Do While a > 0

' Calculate the coordinates of each point
  points(0) = points(6) + I:    points(1) = 0
  points(2) = points(6) + I:    points(3) = a
  points(4) = points(0) + b:    points(5) = a
  points(6) = points(0) + b:    points(7) = 0

' Constructing four segments and connecting the corresponding points

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
  plineObj.Closed = True
  
  ' Offset the polyline
  Dim offsetObj As Variant
                                 'plineObj.color = acByLayer
                                 'Dim layerObj As AcadLayer
                                 'Set layerObj = ThisDrawing.Layers.Add("Ball-6.1")
                                 'layerObj.color = acRed

plineObj.Layer = "C-Mill"
plineObj.Update
  offsetObj = plineObj.Offset(66)
plineObj.Layer = "0"
plineObj.Update

I = 100
  
 Dim plineObj2 As AcadLWPolyline
  Dim points2(0 To 7) As Double
  
  points2(0) = points(0):    points2(1) = a + I:
  points2(2) = points(0):    points2(3) = (a * 2) + I:
  points2(4) = points(6):    points2(5) = (a * 2) + I:
  points2(6) = points(6):    points2(7) = a + I:
  
 For d = 2 To Cells(c, 4)
If d > 1 Then

Set plineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(points2)
  plineObj2.Closed = True
  
  points2(1) = points2(3) + I:
  points2(3) = points2(3) + a + I:
  points2(5) = points2(3):
  points2(7) = points2(1):
  
  ' Offset the polyline
plineObj2.Layer = "C-Mill"
plineObj2.Update
  offsetObj = plineObj2.Offset(66)
plineObj2.Layer = "0"
plineObj2.Update

Else: d = 1
d = d + 1

d = Cells(c, 4)
End If

Next d
c = c + 1

a = Cells(c, 2)
b = Cells(c, 3)

I = 100

Loop
              
ZoomAll
AP.Quit
End Sub

Sub MF119()
' Creating search objects
Dim AP As Excel.Application
Dim WB As Excel.Workbook
Dim WS As Excel.Worksheet

Set AP = Excel.Application
Set WB = AP.Workbooks.Open("C:\Users\������\Desktop\������\2.xlsx")
Set WS = WB.Worksheets("������ �˨������ ������")
c = 19
a = Cells(c, 2)
b = Cells(c, 3)
d = Cells(c, 4)
E = 1
I = 0
On Error Resume Next

' Description of all the points that will make up rectangles
  Dim plineObj As AcadLWPolyline
  Dim plineObjint1 As AcadLWPolyline
  Dim plineObjint2 As AcadLWPolyline
  Dim points(0 To 7) As Double
  Dim pointsrad(0 To 3) As Double
  Dim pointsrad2(0 To 3) As Double
  Dim pointswithin(0 To 21) As Double
  Dim pointswithin2(0 To 7) As Double
  Dim pointswithin436(0 To 25) As Double
  Dim circleObj1 As AcadCircle
  Dim circleObj2 As AcadCircle
  Dim circleObj3 As AcadCircle
  Dim intPoints1
  Dim intPoints2
  Dim intPoints3(0 To 1) As Variant
  Dim intPoints4(0 To 1) As Variant
  Dim intPoints5
  Dim intPoints6(0 To 1) As Variant
  Dim radius As Double
  Dim currentBulge As Double
  Dim distbtwnrad As Double
  Dim angle1 As Double
  Dim angle2 As Double
  Dim angle3 As Double
  Dim cosangle1 As Double
  Dim angle1rad As Double
  Dim offsetObj As Variant
  Dim cntr1(0 To 2) As Double
  Dim cntr2(0 To 2) As Double
  Dim cntr3(0 To 2) As Double
  
points(6) = 0

Do While a > 0

' Calculate the coordinates of each point
  points(0) = points(6) + I:    points(1) = 0
  points(2) = points(6) + I:    points(3) = a
  points(4) = points(0) + b:    points(5) = a
  points(6) = points(0) + b:    points(7) = 0

' Constructing four segments and connecting the corresponding points

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
  plineObj.Closed = True

If a > 180 Then
If b > 180 Then

 cntr1(0) = points(0) + 76: cntr1(1) = points(3) - 100
 radius = (((b - 172) ^ 2) / 96 + 24) / 2
 cntr2(0) = points(0) + (b / 2): cntr2(1) = points(3) - 66 - radius
 katg = cntr2(0) - cntr1(0)
 katv = cntr1(1) - cntr2(1)
 
 
 Set circleObj1 = ThisDrawing.ModelSpace.AddCircle(cntr1, 10)
 Set circleObj2 = ThisDrawing.ModelSpace.AddCircle(cntr2, radius)
 x = cntr2(0) - cntr1(0)
 y = cntr2(1) - cntr1(1)
 distbtwnrad = Sqr((x * x) + (y * y))
 outerradius1 = 10 + 57
 outerradius2 = radius + 57
 
 
 cosangle1 = (((outerradius1 * outerradius1) + (distbtwnrad * distbtwnrad)) - (outerradius2 * outerradius2)) / (2 * (outerradius1 * distbtwnrad))
 angle1rad = Atn(-cosangle1 / Sqr(-cosangle1 * cosangle1 + 1)) + 2 * Atn(1)
 angle1grad = angle1rad * (180 / 3.14159265358979)

 pointsrad(0) = cntr1(0):    pointsrad(1) = cntr1(1)
 pointsrad(2) = cntr2(0):    pointsrad(3) = cntr2(1)
 Set plineObjint1 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsrad)
  plineObjint1.Rotate cntr1, angle1rad
  plineObjint1.Update
cosangle2 = (((distbtwnrad * distbtwnrad) + (outerradius2 * outerradius2)) - (outerradius1 * outerradius1)) / (2 * (distbtwnrad * outerradius2))
angle2rad = Atn(-cosangle2 / Sqr(-cosangle2 * cosangle2 + 1)) + 2 * Atn(1)

pointsrad(0) = cntr2(0):         pointsrad(1) = cntr2(1)
pointsrad(2) = cntr1(0) - katg:  pointsrad(3) = cntr1(1) + katv
 Set plineObjint2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsrad)
  plineObjint2.Rotate cntr2, -angle2rad
  plineObjint2.Update
  
intPoints = plineObjint1.IntersectWith(plineObjint2, acExtendBoth)
intPoints1 = plineObjint1.IntersectWith(circleObj1, acExtendNone)
intPoints2 = plineObjint2.IntersectWith(circleObj2, acExtendThisEntity)
If intPoints2(1) < cntr2(1) Then
intPoints2 = plineObjint2.IntersectWith(circleObj2, acExtendOtherEntity)
End If
z1 = intPoints1(0) - points(0)
z2 = intPoints2(0) - points(0)
intPoints3(0) = points(4) - z1: intPoints3(1) = intPoints1(1)
intPoints4(0) = points(4) - z2: intPoints4(1) = intPoints2(1)

  pointswithin(0) = points(0) + 66:      pointswithin(1) = points(1) + 76
  pointswithin(2) = points(0) + 66:      pointswithin(3) = points(3) - 100
  pointswithin(4) = intPoints1(0):       pointswithin(5) = intPoints1(1)
  pointswithin(6) = intPoints2(0):       pointswithin(7) = intPoints2(1)
  pointswithin(8) = points(4) - (b / 2): pointswithin(9) = points(3) - 66
  pointswithin(10) = intPoints4(0):      pointswithin(11) = intPoints4(1)
  pointswithin(12) = intPoints3(0):      pointswithin(13) = intPoints3(1)
  pointswithin(14) = points(4) - 66:     pointswithin(15) = points(3) - 100
  pointswithin(16) = points(4) - 66:     pointswithin(17) = points(1) + 76
  pointswithin(18) = points(4) - 76:     pointswithin(19) = points(1) + 66
  pointswithin(20) = points(0) + 76:     pointswithin(21) = points(1) + 66
  
plineObjint1.Delete
plineObjint2.Delete
circleObj1.Delete
circleObj2.Delete
    
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointswithin)
    
  If b >= 436 Then
  gptnz4 = radius + 57
  kttv = radius + 33
  kttg = Sqr(((gptnz4) ^ 2) - ((kttv) ^ 2))
  cntr3(0) = points(4) - (b / 2) - kttg: cntr3(1) = points(3) - 33
  Set circleObj3 = ThisDrawing.ModelSpace.AddCircle(cntr3, 57)
  pointsrad2(0) = cntr3(0):         pointsrad2(1) = cntr3(1)
  pointsrad2(2) = cntr3(0):         pointsrad2(3) = cntr3(1) - radius
  Set plineObjint3 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsrad2)
  intPoints5 = plineObjint3.IntersectWith(circleObj3, acExtendNone)
  z3 = intPoints5(0) - points(0)
  intPoints6(0) = points(4) - z3: intPoints6(1) = intPoints5(1)
  circleObj3.Delete
  plineObjint3.Delete
  
  pointswithin436(0) = points(0) + 66:                          pointswithin436(1) = points(1) + 76
  pointswithin436(2) = points(0) + 66:                          pointswithin436(3) = points(3) - 100
  pointswithin436(4) = cntr1(0):                                pointswithin436(5) = cntr1(1) + 10
  pointswithin436(6) = intPoints5(0):                           pointswithin436(7) = intPoints5(1)
  pointswithin436(8) = intPoints2(0):                           pointswithin436(9) = intPoints2(1)
  pointswithin436(10) = points(4) - (b / 2):                    pointswithin436(11) = points(3) - 66
  pointswithin436(12) = intPoints4(0):                          pointswithin436(13) = intPoints4(1)
  pointswithin436(14) = intPoints6(0):                          pointswithin436(15) = intPoints6(1)
  pointswithin436(16) = points(4) - (cntr1(0) - points(0)):     pointswithin436(17) = cntr1(1) + 10
  pointswithin436(18) = points(4) - 66:                         pointswithin436(19) = points(3) - 100
  pointswithin436(20) = points(4) - 66:                         pointswithin436(21) = points(1) + 76
  pointswithin436(22) = points(4) - 76:                         pointswithin436(23) = points(1) + 66
  pointswithin436(24) = points(0) + 76:                         pointswithin436(25) = points(1) + 66
  
  End If

gptnz1 = Sqr((pointswithin(4) - pointswithin(2)) ^ 2 + (pointswithin(5) - pointswithin(3)) ^ 2)
gptnz2 = Sqr((pointswithin(6) - pointswithin(4)) ^ 2 + (pointswithin(7) - pointswithin(5)) ^ 2)
gptnz3 = Sqr((pointswithin(8) - pointswithin(6)) ^ 2 + (pointswithin(9) - pointswithin(7)) ^ 2)

anglebulge1 = Atn((gptnz1 / 20) / Sqr(-(gptnz1 / 20) * (gptnz1 / 20) + 1))
anglebulge2 = Atn((gptnz2 / 114) / Sqr(-(gptnz2 / 114) * (gptnz2 / 114) + 1))
anglebulge3 = Atn((gptnz3 / (2 * radius)) / Sqr(-(gptnz3 / (2 * radius)) * (gptnz3 / (2 * radius)) + 1))
h = 20 * ((1 - Cos(anglebulge1)) / 2)
k = h / gptnz1
   ' Find the bulge of the third segment
    currentBulge = plineObj.GetBulge(2)
If b < 436 Then
    ' Change the bulge of the third segment
    plineObj.SetBulge 1, -k * 2
    plineObj.Update
    plineObj.SetBulge 6, -k * 2
    plineObj.Update
h = 114 * ((1 - Cos(anglebulge2)) / 2)
k = h / gptnz2
    plineObj.SetBulge 2, k * 2
    plineObj.Update
    plineObj.SetBulge 5, k * 2
    plineObj.Update
h = (2 * radius) * ((1 - Cos(anglebulge3)) / 2)
k = h / gptnz3
    plineObj.SetBulge 3, -k * 2
    plineObj.Update
    plineObj.SetBulge 4, -k * 2
    plineObj.Update
    plineObj.SetBulge 8, -0.41421356
    plineObj.Update
    plineObj.SetBulge 10, -0.41421356
    plineObj.Update
    plineObj.Layer = "C-Mill"
    plineObj.Update
    plineObj.Closed = True
    
plineObj.Layer = "Ball-6"
plineObj.Update
  offsetObj = plineObj.Offset(-30)
plineObj.Layer = "C-Mill"
plineObj.Update
    
End If
If b >= 436 Then
plineObj.Delete
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointswithin436)

gptnz1 = Sqr((pointswithin436(4) - pointswithin436(2)) ^ 2 + (pointswithin436(5) - pointswithin436(3)) ^ 2)
gptnz2 = Sqr((pointswithin436(8) - pointswithin436(6)) ^ 2 + (pointswithin436(9) - pointswithin436(7)) ^ 2)
anglebulge1 = Atn((gptnz1 / 20) / Sqr(-(gptnz1 / 20) * (gptnz1 / 20) + 1))
anglebulge2 = Atn((gptnz2 / 114) / Sqr(-(gptnz2 / 114) * (gptnz2 / 114) + 1))
h = 20 * ((1 - Cos(anglebulge1)) / 2)
k = h / gptnz1
    plineObj.SetBulge 1, -k * 2
    plineObj.Update
    plineObj.SetBulge 8, -k * 2
    plineObj.Update
h = 114 * ((1 - Cos(anglebulge2)) / 2)
k = h / gptnz2
    plineObj.SetBulge 3, k * 2
    plineObj.Update
    plineObj.SetBulge 6, k * 2
    plineObj.Update
h = (2 * radius) * ((1 - Cos(anglebulge3)) / 2)
k = h / gptnz3
    plineObj.SetBulge 4, -k * 2
    plineObj.Update
    plineObj.SetBulge 5, -k * 2
    plineObj.Update
    plineObj.SetBulge 10, -0.41421356
    plineObj.Update
    plineObj.SetBulge 12, -0.41421356
    plineObj.Update
    plineObj.Layer = "C-Mill"
    plineObj.Update
    plineObj.Closed = True
    
plineObj.Layer = "Ball-6"
plineObj.Update
  offsetObj = plineObj.Offset(-30)
plineObj.Layer = "C-Mill"
plineObj.Update
    
End If
End If
End If
                                 'plineObj.color = acByLayer
                                 'Dim layerObj As AcadLayer
                                 'Set layerObj = ThisDrawing.Layers.Add("Ball-6.1")
                                 'layerObj.color = acRed
   
I = 100
  
 Dim plineObj2 As AcadLWPolyline
  Dim points2(0 To 7) As Double
  
  points2(0) = points(0):    points2(1) = a + I:
  points2(2) = points(0):    points2(3) = (a * 2) + I:
  points2(4) = points(6):    points2(5) = (a * 2) + I:
  points2(6) = points(6):    points2(7) = a + I:
  
 For d = 2 To Cells(c, 4)
If d > 1 Then

Set plineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(points2)
  plineObj2.Closed = True

  pointswithin2(0) = points2(0) + 66:    pointswithin2(1) = points2(1) + 66
  pointswithin2(2) = points2(2) + 66:    pointswithin2(3) = points2(3) - 90
  pointswithin2(4) = points2(4) - 66:    pointswithin2(5) = points2(5) - 90
  pointswithin2(6) = points2(6) - 66:    pointswithin2(7) = points2(7) + 66

If a > 180 Then
If b > 180 Then

 cntr1(0) = points2(0) + 76: cntr1(1) = points2(3) - 100
 radius = (((b - 172) ^ 2) / 96 + 24) / 2
 cntr2(0) = points2(0) + (b / 2): cntr2(1) = points2(3) - 66 - radius
 katg = cntr2(0) - cntr1(0)
 katv = cntr1(1) - cntr2(1)
 
 
 Set circleObj1 = ThisDrawing.ModelSpace.AddCircle(cntr1, 10)
 Set circleObj2 = ThisDrawing.ModelSpace.AddCircle(cntr2, radius)
 x = cntr2(0) - cntr1(0)
 y = cntr2(1) - cntr1(1)
 distbtwnrad = Sqr((x * x) + (y * y))
 outerradius1 = 10 + 57
 outerradius2 = radius + 57
 
 
 cosangle1 = (((outerradius1 * outerradius1) + (distbtwnrad * distbtwnrad)) - (outerradius2 * outerradius2)) / (2 * (outerradius1 * distbtwnrad))
 angle1rad = Atn(-cosangle1 / Sqr(-cosangle1 * cosangle1 + 1)) + 2 * Atn(1)
 angle1grad = angle1rad * (180 / 3.14159265358979)

 pointsrad(0) = cntr1(0):    pointsrad(1) = cntr1(1)
 pointsrad(2) = cntr2(0):    pointsrad(3) = cntr2(1)
 Set plineObjint1 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsrad)
  plineObjint1.Rotate cntr1, angle1rad
  plineObjint1.Update
cosangle2 = (((distbtwnrad * distbtwnrad) + (outerradius2 * outerradius2)) - (outerradius1 * outerradius1)) / (2 * (distbtwnrad * outerradius2))
angle2rad = Atn(-cosangle2 / Sqr(-cosangle2 * cosangle2 + 1)) + 2 * Atn(1)

pointsrad(0) = cntr2(0):         pointsrad(1) = cntr2(1)
pointsrad(2) = cntr1(0) - katg:  pointsrad(3) = cntr1(1) + katv
 Set plineObjint2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsrad)
  plineObjint2.Rotate cntr2, -angle2rad
  plineObjint2.Update
  
intPoints = plineObjint1.IntersectWith(plineObjint2, acExtendBoth)
intPoints1 = plineObjint1.IntersectWith(circleObj1, acExtendNone)
intPoints2 = plineObjint2.IntersectWith(circleObj2, acExtendThisEntity)
If intPoints2(1) < cntr2(1) Then
intPoints2 = plineObjint2.IntersectWith(circleObj2, acExtendOtherEntity)
End If
z1 = intPoints1(0) - points2(0)
z2 = intPoints2(0) - points2(0)
intPoints3(0) = points2(4) - z1: intPoints3(1) = intPoints1(1)
intPoints4(0) = points2(4) - z2: intPoints4(1) = intPoints2(1)

  pointswithin(0) = points2(0) + 66:      pointswithin(1) = points2(1) + 76
  pointswithin(2) = points2(0) + 66:      pointswithin(3) = points2(3) - 100
  pointswithin(4) = intPoints1(0):       pointswithin(5) = intPoints1(1)
  pointswithin(6) = intPoints2(0):       pointswithin(7) = intPoints2(1)
  pointswithin(8) = points2(4) - (b / 2): pointswithin(9) = points2(3) - 66
  pointswithin(10) = intPoints4(0):      pointswithin(11) = intPoints4(1)
  pointswithin(12) = intPoints3(0):      pointswithin(13) = intPoints3(1)
  pointswithin(14) = points2(4) - 66:     pointswithin(15) = points2(3) - 100
  pointswithin(16) = points2(4) - 66:     pointswithin(17) = points2(1) + 76
  pointswithin(18) = points2(4) - 76:     pointswithin(19) = points2(1) + 66
  pointswithin(20) = points2(0) + 76:     pointswithin(21) = points2(1) + 66
  
plineObjint1.Delete
plineObjint2.Delete
circleObj1.Delete
circleObj2.Delete
    
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointswithin)
    
  If b >= 436 Then
  gptnz4 = radius + 57
  kttv = radius + 33
  kttg = Sqr(((gptnz4) ^ 2) - ((kttv) ^ 2))
  cntr3(0) = points2(4) - (b / 2) - kttg: cntr3(1) = points2(3) - 33
  Set circleObj3 = ThisDrawing.ModelSpace.AddCircle(cntr3, 57)
  pointsrad2(0) = cntr3(0):         pointsrad2(1) = cntr3(1)
  pointsrad2(2) = cntr3(0):         pointsrad2(3) = cntr3(1) - radius
  Set plineObjint3 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsrad2)
  intPoints5 = plineObjint3.IntersectWith(circleObj3, acExtendNone)
  z3 = intPoints5(0) - points2(0)
  intPoints6(0) = points2(4) - z3: intPoints6(1) = intPoints5(1)
  circleObj3.Delete
  plineObjint3.Delete
  
  pointswithin436(0) = points2(0) + 66:                          pointswithin436(1) = points2(1) + 76
  pointswithin436(2) = points2(0) + 66:                          pointswithin436(3) = points2(3) - 100
  pointswithin436(4) = cntr1(0):                                pointswithin436(5) = cntr1(1) + 10
  pointswithin436(6) = intPoints5(0):                           pointswithin436(7) = intPoints5(1)
  pointswithin436(8) = intPoints2(0):                           pointswithin436(9) = intPoints2(1)
  pointswithin436(10) = points2(4) - (b / 2):                    pointswithin436(11) = points2(3) - 66
  pointswithin436(12) = intPoints4(0):                          pointswithin436(13) = intPoints4(1)
  pointswithin436(14) = intPoints6(0):                          pointswithin436(15) = intPoints6(1)
  pointswithin436(16) = points2(4) - (cntr1(0) - points2(0)):     pointswithin436(17) = cntr1(1) + 10
  pointswithin436(18) = points2(4) - 66:                         pointswithin436(19) = points2(3) - 100
  pointswithin436(20) = points2(4) - 66:                         pointswithin436(21) = points2(1) + 76
  pointswithin436(22) = points2(4) - 76:                         pointswithin436(23) = points2(1) + 66
  pointswithin436(24) = points2(0) + 76:                         pointswithin436(25) = points2(1) + 66
  
  End If

gptnz1 = Sqr((pointswithin(4) - pointswithin(2)) ^ 2 + (pointswithin(5) - pointswithin(3)) ^ 2)
gptnz2 = Sqr((pointswithin(6) - pointswithin(4)) ^ 2 + (pointswithin(7) - pointswithin(5)) ^ 2)
gptnz3 = Sqr((pointswithin(8) - pointswithin(6)) ^ 2 + (pointswithin(9) - pointswithin(7)) ^ 2)

anglebulge1 = Atn((gptnz1 / 20) / Sqr(-(gptnz1 / 20) * (gptnz1 / 20) + 1))
anglebulge2 = Atn((gptnz2 / 114) / Sqr(-(gptnz2 / 114) * (gptnz2 / 114) + 1))
anglebulge3 = Atn((gptnz3 / (2 * radius)) / Sqr(-(gptnz3 / (2 * radius)) * (gptnz3 / (2 * radius)) + 1))
h = 20 * ((1 - Cos(anglebulge1)) / 2)
k = h / gptnz1
   ' Find the bulge of the third segment
    currentBulge = plineObj.GetBulge(2)
If b < 436 Then
    ' Change the bulge of the third segment
    plineObj.SetBulge 1, -k * 2
    plineObj.Update
    plineObj.SetBulge 6, -k * 2
    plineObj.Update
h = 114 * ((1 - Cos(anglebulge2)) / 2)
k = h / gptnz2
    plineObj.SetBulge 2, k * 2
    plineObj.Update
    plineObj.SetBulge 5, k * 2
    plineObj.Update
h = (2 * radius) * ((1 - Cos(anglebulge3)) / 2)
k = h / gptnz3
    plineObj.SetBulge 3, -k * 2
    plineObj.Update
    plineObj.SetBulge 4, -k * 2
    plineObj.Update
    plineObj.SetBulge 8, -0.41421356
    plineObj.Update
    plineObj.SetBulge 10, -0.41421356
    plineObj.Update
    plineObj.Layer = "C-Mill"
    plineObj.Update
    plineObj.Closed = True
    
plineObj.Layer = "Ball-6"
plineObj.Update
  offsetObj = plineObj.Offset(-30)
plineObj.Layer = "C-Mill"
plineObj.Update
    
End If
If b >= 436 Then
plineObj.Delete
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointswithin436)

gptnz1 = Sqr((pointswithin436(4) - pointswithin436(2)) ^ 2 + (pointswithin436(5) - pointswithin436(3)) ^ 2)
gptnz2 = Sqr((pointswithin436(8) - pointswithin436(6)) ^ 2 + (pointswithin436(9) - pointswithin436(7)) ^ 2)
anglebulge1 = Atn((gptnz1 / 20) / Sqr(-(gptnz1 / 20) * (gptnz1 / 20) + 1))
anglebulge2 = Atn((gptnz2 / 114) / Sqr(-(gptnz2 / 114) * (gptnz2 / 114) + 1))
h = 20 * ((1 - Cos(anglebulge1)) / 2)
k = h / gptnz1
    plineObj.SetBulge 1, -k * 2
    plineObj.Update
    plineObj.SetBulge 8, -k * 2
    plineObj.Update
h = 114 * ((1 - Cos(anglebulge2)) / 2)
k = h / gptnz2
    plineObj.SetBulge 3, k * 2
    plineObj.Update
    plineObj.SetBulge 6, k * 2
    plineObj.Update
h = (2 * radius) * ((1 - Cos(anglebulge3)) / 2)
k = h / gptnz3
    plineObj.SetBulge 4, -k * 2
    plineObj.Update
    plineObj.SetBulge 5, -k * 2
    plineObj.Update
    plineObj.SetBulge 10, -0.41421356
    plineObj.Update
    plineObj.SetBulge 12, -0.41421356
    plineObj.Update
    plineObj.Layer = "C-Mill"
    plineObj.Update
    plineObj.Closed = True
    
plineObj.Layer = "Ball-6"
plineObj.Update
  offsetObj = plineObj.Offset(-30)
plineObj.Layer = "C-Mill"
plineObj.Update
    
End If
  
  
  
End If
End If
  points2(1) = points2(3) + I:
  points2(3) = points2(3) + a + I:
  points2(5) = points2(3):
  points2(7) = points2(1):
      
Else: d = 1
d = d + 1

d = Cells(c, 4)
End If

Next d
c = c + 1

a = Cells(c, 2)
b = Cells(c, 3)

I = 100

Loop
              
ZoomAll
AP.Quit
End Sub

Sub MF120()
' Creating search objects
Dim AP As Excel.Application
Dim WB As Excel.Workbook
Dim WS As Excel.Worksheet

Set AP = Excel.Application
Set WB = AP.Workbooks.Open("C:\Users\������\Desktop\������\2.xlsx")
Set WS = WB.Worksheets("������ �˨������ ������")
c = 19
a = Cells(c, 2)
b = Cells(c, 3)
d = Cells(c, 4)
E = 1
I = 0
On Error Resume Next

' Description of all the points that will make up rectangles
  Dim plineObj As AcadLWPolyline
  Dim plineObjint1 As AcadLWPolyline
  Dim plineObjint2 As AcadLWPolyline
  Dim points(0 To 7) As Double
  Dim pointsrad(0 To 3) As Double
  Dim pointsrad2(0 To 3) As Double
  Dim pointswithin(0 To 27) As Double
  Dim pointswithin2(0 To 7) As Double
  Dim pointswithin436(0 To 35) As Double
  Dim circleObj1 As AcadCircle
  Dim circleObj2 As AcadCircle
  Dim circleObj3 As AcadCircle
  Dim intPoints1
  Dim intPoints2
  Dim intPoints3(0 To 1) As Variant
  Dim intPoints4(0 To 1) As Variant
  Dim intPoints5
  Dim intPoints6(0 To 1) As Variant
  Dim radius As Double
  Dim currentBulge As Double
  Dim distbtwnrad As Double
  Dim angle1 As Double
  Dim angle2 As Double
  Dim angle3 As Double
  Dim cosangle1 As Double
  Dim angle1rad As Double
  Dim offsetObj As Variant
  Dim cntr1(0 To 2) As Double
  Dim cntr2(0 To 2) As Double
  Dim cntr3(0 To 2) As Double
  
points(6) = 0

Do While a > 0

' Calculate the coordinates of each point
  points(0) = points(6) + I:    points(1) = 0
  points(2) = points(6) + I:    points(3) = a
  points(4) = points(0) + b:    points(5) = a
  points(6) = points(0) + b:    points(7) = 0

' Constructing four segments and connecting the corresponding points

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
  plineObj.Closed = True

If a > 180 Then
If b > 180 Then

 cntr1(0) = points(0) + 76: cntr1(1) = points(3) - 100
 radius = (((b - 172) ^ 2) / 96 + 24) / 2
 cntr2(0) = points(0) + (b / 2): cntr2(1) = points(3) - 66 - radius
 katg = cntr2(0) - cntr1(0)
 katv = cntr1(1) - cntr2(1)
 
 
 Set circleObj1 = ThisDrawing.ModelSpace.AddCircle(cntr1, 10)
 Set circleObj2 = ThisDrawing.ModelSpace.AddCircle(cntr2, radius)
 x = cntr2(0) - cntr1(0)
 y = cntr2(1) - cntr1(1)
 distbtwnrad = Sqr((x * x) + (y * y))
 outerradius1 = 10 + 57
 outerradius2 = radius + 57
 
 
 cosangle1 = (((outerradius1 * outerradius1) + (distbtwnrad * distbtwnrad)) - (outerradius2 * outerradius2)) / (2 * (outerradius1 * distbtwnrad))
 angle1rad = Atn(-cosangle1 / Sqr(-cosangle1 * cosangle1 + 1)) + 2 * Atn(1)
 angle1grad = angle1rad * (180 / 3.14159265358979)

 pointsrad(0) = cntr1(0):    pointsrad(1) = cntr1(1)
 pointsrad(2) = cntr2(0):    pointsrad(3) = cntr2(1)
 Set plineObjint1 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsrad)
  plineObjint1.Rotate cntr1, angle1rad
  plineObjint1.Update
cosangle2 = (((distbtwnrad * distbtwnrad) + (outerradius2 * outerradius2)) - (outerradius1 * outerradius1)) / (2 * (distbtwnrad * outerradius2))
angle2rad = Atn(-cosangle2 / Sqr(-cosangle2 * cosangle2 + 1)) + 2 * Atn(1)

pointsrad(0) = cntr2(0):         pointsrad(1) = cntr2(1)
pointsrad(2) = cntr1(0) - katg:  pointsrad(3) = cntr1(1) + katv
 Set plineObjint2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsrad)
  plineObjint2.Rotate cntr2, -angle2rad
  plineObjint2.Update
  
intPoints = plineObjint1.IntersectWith(plineObjint2, acExtendBoth)
intPoints1 = plineObjint1.IntersectWith(circleObj1, acExtendNone)
intPoints2 = plineObjint2.IntersectWith(circleObj2, acExtendThisEntity)
If intPoints2(1) < cntr2(1) Then
intPoints2 = plineObjint2.IntersectWith(circleObj2, acExtendOtherEntity)
End If
z1 = intPoints1(0) - points(0)
z2 = intPoints2(0) - points(0)
intPoints3(0) = points(4) - z1: intPoints3(1) = intPoints1(1)
intPoints4(0) = points(4) - z2: intPoints4(1) = intPoints2(1)

  pointswithin(0) = points(0) + 66:      pointswithin(1) = points(1) + 100
  pointswithin(2) = points(0) + 66:      pointswithin(3) = points(3) - 100
  pointswithin(4) = intPoints1(0):       pointswithin(5) = intPoints1(1)
  pointswithin(6) = intPoints2(0):       pointswithin(7) = intPoints2(1)
  pointswithin(8) = points(4) - (b / 2): pointswithin(9) = points(3) - 66
  pointswithin(10) = intPoints4(0):      pointswithin(11) = intPoints4(1)
  pointswithin(12) = intPoints3(0):      pointswithin(13) = intPoints3(1)
  pointswithin(14) = points(4) - 66:     pointswithin(15) = points(3) - 100
  pointswithin(16) = points(4) - 66:      pointswithin(17) = points(1) + 100
  pointswithin(18) = intPoints3(0):       pointswithin(19) = points(1) + (points(3) - intPoints1(1))
  pointswithin(20) = intPoints4(0):       pointswithin(21) = points(1) + (points(3) - intPoints2(1))
  pointswithin(22) = points(4) - (b / 2): pointswithin(23) = points(1) + 66
  pointswithin(24) = intPoints2(0):       pointswithin(25) = points(1) + (points(3) - intPoints4(1))
  pointswithin(26) = intPoints1(0):       pointswithin(27) = points(1) + (points(3) - intPoints3(1))
  
  
plineObjint1.Delete
plineObjint2.Delete
circleObj1.Delete
circleObj2.Delete
    
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointswithin)
    
  If b >= 436 Then
  gptnz4 = radius + 57
  kttv = radius + 33
  kttg = Sqr(((gptnz4) ^ 2) - ((kttv) ^ 2))
  cntr3(0) = points(4) - (b / 2) - kttg: cntr3(1) = points(3) - 33
  Set circleObj3 = ThisDrawing.ModelSpace.AddCircle(cntr3, 57)
  pointsrad2(0) = cntr3(0):         pointsrad2(1) = cntr3(1)
  pointsrad2(2) = cntr3(0):         pointsrad2(3) = cntr3(1) - radius
  Set plineObjint3 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsrad2)
  intPoints5 = plineObjint3.IntersectWith(circleObj3, acExtendNone)
  z3 = intPoints5(0) - points(0)
  intPoints6(0) = points(4) - z3: intPoints6(1) = intPoints5(1)
  circleObj3.Delete
  plineObjint3.Delete
  
  pointswithin436(0) = points(0) + 66:                          pointswithin436(1) = points(1) + 100
  pointswithin436(2) = points(0) + 66:                          pointswithin436(3) = points(3) - 100
  pointswithin436(4) = cntr1(0):                                pointswithin436(5) = cntr1(1) + 10
  pointswithin436(6) = intPoints5(0):                           pointswithin436(7) = intPoints5(1)
  pointswithin436(8) = intPoints2(0):                           pointswithin436(9) = intPoints2(1)
  pointswithin436(10) = points(4) - (b / 2):                    pointswithin436(11) = points(3) - 66
  pointswithin436(12) = intPoints4(0):                          pointswithin436(13) = intPoints4(1)
  pointswithin436(14) = intPoints6(0):                          pointswithin436(15) = intPoints6(1)
  pointswithin436(16) = points(4) - (cntr1(0) - points(0)):     pointswithin436(17) = cntr1(1) + 10
  pointswithin436(18) = points(4) - 66:                         pointswithin436(19) = points(3) - 100
  pointswithin436(20) = points(4) - 66:                         pointswithin436(21) = points(1) + 100
  pointswithin436(22) = points(4) - (cntr1(0) - points(0)):     pointswithin436(23) = points(1) + 90
  pointswithin436(24) = intPoints6(0):                          pointswithin436(25) = points(1) + (points(3) - intPoints6(1))
  pointswithin436(26) = intPoints4(0):                          pointswithin436(27) = points(1) + (points(3) - intPoints4(1))
  pointswithin436(28) = points(4) - (b / 2):                    pointswithin436(29) = points(1) + 66
  pointswithin436(30) = intPoints2(0):                           pointswithin436(31) = points(1) + (points(3) - intPoints2(1))
  pointswithin436(32) = intPoints5(0):                           pointswithin436(33) = points(1) + (points(3) - intPoints5(1))
  pointswithin436(34) = cntr1(0):                                pointswithin436(35) = points(1) + 90
 
  
  End If

gptnz1 = Sqr((pointswithin(4) - pointswithin(2)) ^ 2 + (pointswithin(5) - pointswithin(3)) ^ 2)
gptnz2 = Sqr((pointswithin(6) - pointswithin(4)) ^ 2 + (pointswithin(7) - pointswithin(5)) ^ 2)
gptnz3 = Sqr((pointswithin(8) - pointswithin(6)) ^ 2 + (pointswithin(9) - pointswithin(7)) ^ 2)

anglebulge1 = Atn((gptnz1 / 20) / Sqr(-(gptnz1 / 20) * (gptnz1 / 20) + 1))
anglebulge2 = Atn((gptnz2 / 114) / Sqr(-(gptnz2 / 114) * (gptnz2 / 114) + 1))
anglebulge3 = Atn((gptnz3 / (2 * radius)) / Sqr(-(gptnz3 / (2 * radius)) * (gptnz3 / (2 * radius)) + 1))
h = 20 * ((1 - Cos(anglebulge1)) / 2)
k = h / gptnz1
   ' Find the bulge of the third segment
    currentBulge = plineObj.GetBulge(2)
If b < 436 Then
    ' Change the bulge of the third segment
    plineObj.SetBulge 1, -k * 2
    plineObj.Update
    plineObj.SetBulge 6, -k * 2
    plineObj.Update
    plineObj.SetBulge 8, -k * 2
    plineObj.Update
    plineObj.SetBulge 13, -k * 2
    plineObj.Update
h = 114 * ((1 - Cos(anglebulge2)) / 2)
k = h / gptnz2
    plineObj.SetBulge 2, k * 2
    plineObj.Update
    plineObj.SetBulge 5, k * 2
    plineObj.Update
    plineObj.SetBulge 9, k * 2
    plineObj.Update
    plineObj.SetBulge 12, k * 2
    plineObj.Update
h = (2 * radius) * ((1 - Cos(anglebulge3)) / 2)
k = h / gptnz3
    plineObj.SetBulge 3, -k * 2
    plineObj.Update
    plineObj.SetBulge 4, -k * 2
    plineObj.Update
    plineObj.SetBulge 10, -k * 2
    plineObj.Update
    plineObj.SetBulge 11, -k * 2
    plineObj.Update
    
    plineObj.Layer = "C-Mill"
    plineObj.Update
    plineObj.Closed = True
    
plineObj.Layer = "Ball-6"
plineObj.Update
  offsetObj = plineObj.Offset(-30)
plineObj.Layer = "C-Mill"
plineObj.Update
    
End If
If b >= 436 Then
plineObj.Delete
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointswithin436)

gptnz1 = Sqr((pointswithin436(4) - pointswithin436(2)) ^ 2 + (pointswithin436(5) - pointswithin436(3)) ^ 2)
gptnz2 = Sqr((pointswithin436(8) - pointswithin436(6)) ^ 2 + (pointswithin436(9) - pointswithin436(7)) ^ 2)
anglebulge1 = Atn((gptnz1 / 20) / Sqr(-(gptnz1 / 20) * (gptnz1 / 20) + 1))
anglebulge2 = Atn((gptnz2 / 114) / Sqr(-(gptnz2 / 114) * (gptnz2 / 114) + 1))
h = 20 * ((1 - Cos(anglebulge1)) / 2)
k = h / gptnz1
    plineObj.SetBulge 1, -k * 2
    plineObj.Update
    plineObj.SetBulge 8, -k * 2
    plineObj.Update
    plineObj.SetBulge 10, -k * 2
    plineObj.Update
    plineObj.SetBulge 17, -k * 2
    plineObj.Update
h = 114 * ((1 - Cos(anglebulge2)) / 2)
k = h / gptnz2
    plineObj.SetBulge 3, k * 2
    plineObj.Update
    plineObj.SetBulge 6, k * 2
    plineObj.Update
    plineObj.SetBulge 12, k * 2
    plineObj.Update
    plineObj.SetBulge 15, k * 2
    plineObj.Update
    
h = (2 * radius) * ((1 - Cos(anglebulge3)) / 2)
k = h / gptnz3
    plineObj.SetBulge 4, -k * 2
    plineObj.Update
    plineObj.SetBulge 5, -k * 2
    plineObj.Update
    plineObj.SetBulge 13, -k * 2
    plineObj.Update
    plineObj.SetBulge 14, -k * 2
    plineObj.Update
    plineObj.Layer = "C-Mill"
    plineObj.Update
    plineObj.Closed = True
    
plineObj.Layer = "Ball-6"
plineObj.Update
  offsetObj = plineObj.Offset(-30)
plineObj.Layer = "C-Mill"
plineObj.Update
    
End If
End If
End If
                                 'plineObj.color = acByLayer
                                 'Dim layerObj As AcadLayer
                                 'Set layerObj = ThisDrawing.Layers.Add("Ball-6.1")
                                 'layerObj.color = acRed
   
I = 100
  
 Dim plineObj2 As AcadLWPolyline
  Dim points2(0 To 7) As Double
  
  points2(0) = points(0):    points2(1) = a + I:
  points2(2) = points(0):    points2(3) = (a * 2) + I:
  points2(4) = points(6):    points2(5) = (a * 2) + I:
  points2(6) = points(6):    points2(7) = a + I:
  
 For d = 2 To Cells(c, 4)
If d > 1 Then

Set plineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(points2)
  plineObj2.Closed = True

  pointswithin2(0) = points2(0) + 66:    pointswithin2(1) = points2(1) + 66
  pointswithin2(2) = points2(2) + 66:    pointswithin2(3) = points2(3) - 90
  pointswithin2(4) = points2(4) - 66:    pointswithin2(5) = points2(5) - 90
  pointswithin2(6) = points2(6) - 66:    pointswithin2(7) = points2(7) + 66

If a > 180 Then
If b > 180 Then

 cntr1(0) = points2(0) + 76: cntr1(1) = points2(3) - 100
 radius = (((b - 172) ^ 2) / 96 + 24) / 2
 cntr2(0) = points2(0) + (b / 2): cntr2(1) = points2(3) - 66 - radius
 katg = cntr2(0) - cntr1(0)
 katv = cntr1(1) - cntr2(1)
 
 
 Set circleObj1 = ThisDrawing.ModelSpace.AddCircle(cntr1, 10)
 Set circleObj2 = ThisDrawing.ModelSpace.AddCircle(cntr2, radius)
 x = cntr2(0) - cntr1(0)
 y = cntr2(1) - cntr1(1)
 distbtwnrad = Sqr((x * x) + (y * y))
 outerradius1 = 10 + 57
 outerradius2 = radius + 57
 
 
 cosangle1 = (((outerradius1 * outerradius1) + (distbtwnrad * distbtwnrad)) - (outerradius2 * outerradius2)) / (2 * (outerradius1 * distbtwnrad))
 angle1rad = Atn(-cosangle1 / Sqr(-cosangle1 * cosangle1 + 1)) + 2 * Atn(1)
 angle1grad = angle1rad * (180 / 3.14159265358979)

 pointsrad(0) = cntr1(0):    pointsrad(1) = cntr1(1)
 pointsrad(2) = cntr2(0):    pointsrad(3) = cntr2(1)
 Set plineObjint1 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsrad)
  plineObjint1.Rotate cntr1, angle1rad
  plineObjint1.Update
cosangle2 = (((distbtwnrad * distbtwnrad) + (outerradius2 * outerradius2)) - (outerradius1 * outerradius1)) / (2 * (distbtwnrad * outerradius2))
angle2rad = Atn(-cosangle2 / Sqr(-cosangle2 * cosangle2 + 1)) + 2 * Atn(1)

pointsrad(0) = cntr2(0):         pointsrad(1) = cntr2(1)
pointsrad(2) = cntr1(0) - katg:  pointsrad(3) = cntr1(1) + katv
 Set plineObjint2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsrad)
  plineObjint2.Rotate cntr2, -angle2rad
  plineObjint2.Update
  
intPoints = plineObjint1.IntersectWith(plineObjint2, acExtendBoth)
intPoints1 = plineObjint1.IntersectWith(circleObj1, acExtendNone)
intPoints2 = plineObjint2.IntersectWith(circleObj2, acExtendThisEntity)
If intPoints2(1) < cntr2(1) Then
intPoints2 = plineObjint2.IntersectWith(circleObj2, acExtendOtherEntity)
End If
z1 = intPoints1(0) - points2(0)
z2 = intPoints2(0) - points2(0)
intPoints3(0) = points2(4) - z1: intPoints3(1) = intPoints1(1)
intPoints4(0) = points2(4) - z2: intPoints4(1) = intPoints2(1)

  pointswithin(0) = points2(0) + 66:      pointswithin(1) = points2(1) + 100
  pointswithin(2) = points2(0) + 66:      pointswithin(3) = points2(3) - 100
  pointswithin(4) = intPoints1(0):       pointswithin(5) = intPoints1(1)
  pointswithin(6) = intPoints2(0):       pointswithin(7) = intPoints2(1)
  pointswithin(8) = points2(4) - (b / 2): pointswithin(9) = points2(3) - 66
  pointswithin(10) = intPoints4(0):      pointswithin(11) = intPoints4(1)
  pointswithin(12) = intPoints3(0):      pointswithin(13) = intPoints3(1)
  pointswithin(14) = points2(4) - 66:     pointswithin(15) = points2(3) - 100
  pointswithin(16) = points2(4) - 66:      pointswithin(17) = points2(1) + 100
  pointswithin(18) = intPoints3(0):       pointswithin(19) = points2(1) + (points2(3) - intPoints1(1))
  pointswithin(20) = intPoints4(0):       pointswithin(21) = points2(1) + (points2(3) - intPoints2(1))
  pointswithin(22) = points2(4) - (b / 2): pointswithin(23) = points2(1) + 66
  pointswithin(24) = intPoints2(0):       pointswithin(25) = points2(1) + (points2(3) - intPoints4(1))
  pointswithin(26) = intPoints1(0):       pointswithin(27) = points2(1) + (points2(3) - intPoints3(1))

plineObjint1.Delete
plineObjint2.Delete
circleObj1.Delete
circleObj2.Delete
    
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointswithin)
    
  If b >= 436 Then
  gptnz4 = radius + 57
  kttv = radius + 33
  kttg = Sqr(((gptnz4) ^ 2) - ((kttv) ^ 2))
  cntr3(0) = points2(4) - (b / 2) - kttg: cntr3(1) = points2(3) - 33
  Set circleObj3 = ThisDrawing.ModelSpace.AddCircle(cntr3, 57)
  pointsrad2(0) = cntr3(0):         pointsrad2(1) = cntr3(1)
  pointsrad2(2) = cntr3(0):         pointsrad2(3) = cntr3(1) - radius
  Set plineObjint3 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsrad2)
  intPoints5 = plineObjint3.IntersectWith(circleObj3, acExtendNone)
  z3 = intPoints5(0) - points2(0)
  intPoints6(0) = points2(4) - z3: intPoints6(1) = intPoints5(1)
  circleObj3.Delete
  plineObjint3.Delete
  
  pointswithin436(0) = points2(0) + 66:                          pointswithin436(1) = points2(1) + 100
  pointswithin436(2) = points2(0) + 66:                          pointswithin436(3) = points2(3) - 100
  pointswithin436(4) = cntr1(0):                                 pointswithin436(5) = cntr1(1) + 10
  pointswithin436(6) = intPoints5(0):                            pointswithin436(7) = intPoints5(1)
  pointswithin436(8) = intPoints2(0):                            pointswithin436(9) = intPoints2(1)
  pointswithin436(10) = points2(4) - (b / 2):                    pointswithin436(11) = points2(3) - 66
  pointswithin436(12) = intPoints4(0):                           pointswithin436(13) = intPoints4(1)
  pointswithin436(14) = intPoints6(0):                           pointswithin436(15) = intPoints6(1)
  pointswithin436(16) = points2(4) - (cntr1(0) - points2(0)):    pointswithin436(17) = cntr1(1) + 10
  pointswithin436(18) = points2(4) - 66:                         pointswithin436(19) = points2(3) - 100
  pointswithin436(20) = points2(4) - 66:                         pointswithin436(21) = points2(1) + 100
  pointswithin436(22) = points2(4) - (cntr1(0) - points2(0)):    pointswithin436(23) = points2(1) + 90
  pointswithin436(24) = intPoints6(0):                           pointswithin436(25) = points2(1) + (points2(3) - intPoints6(1))
  pointswithin436(26) = intPoints4(0):                           pointswithin436(27) = points2(1) + (points2(3) - intPoints4(1))
  pointswithin436(28) = points2(4) - (b / 2):                    pointswithin436(29) = points2(1) + 66
  pointswithin436(30) = intPoints2(0):                           pointswithin436(31) = points2(1) + (points2(3) - intPoints2(1))
  pointswithin436(32) = intPoints5(0):                           pointswithin436(33) = points2(1) + (points2(3) - intPoints5(1))
  pointswithin436(34) = cntr1(0):                                pointswithin436(35) = points2(1) + 90
  
  End If

gptnz1 = Sqr((pointswithin(4) - pointswithin(2)) ^ 2 + (pointswithin(5) - pointswithin(3)) ^ 2)
gptnz2 = Sqr((pointswithin(6) - pointswithin(4)) ^ 2 + (pointswithin(7) - pointswithin(5)) ^ 2)
gptnz3 = Sqr((pointswithin(8) - pointswithin(6)) ^ 2 + (pointswithin(9) - pointswithin(7)) ^ 2)

anglebulge1 = Atn((gptnz1 / 20) / Sqr(-(gptnz1 / 20) * (gptnz1 / 20) + 1))
anglebulge2 = Atn((gptnz2 / 114) / Sqr(-(gptnz2 / 114) * (gptnz2 / 114) + 1))
anglebulge3 = Atn((gptnz3 / (2 * radius)) / Sqr(-(gptnz3 / (2 * radius)) * (gptnz3 / (2 * radius)) + 1))
h = 20 * ((1 - Cos(anglebulge1)) / 2)
k = h / gptnz1
   ' Find the bulge of the third segment
    currentBulge = plineObj.GetBulge(2)
If b < 436 Then
    ' Change the bulge of the third segment
    plineObj.SetBulge 1, -k * 2
    plineObj.Update
    plineObj.SetBulge 6, -k * 2
    plineObj.Update
    plineObj.SetBulge 8, -k * 2
    plineObj.Update
    plineObj.SetBulge 13, -k * 2
    plineObj.Update
h = 114 * ((1 - Cos(anglebulge2)) / 2)
k = h / gptnz2
    plineObj.SetBulge 2, k * 2
    plineObj.Update
    plineObj.SetBulge 5, k * 2
    plineObj.Update
    plineObj.SetBulge 9, k * 2
    plineObj.Update
    plineObj.SetBulge 12, k * 2
    plineObj.Update
h = (2 * radius) * ((1 - Cos(anglebulge3)) / 2)
k = h / gptnz3
    plineObj.SetBulge 3, -k * 2
    plineObj.Update
    plineObj.SetBulge 4, -k * 2
    plineObj.Update
   plineObj.SetBulge 10, -k * 2
    plineObj.Update
    plineObj.SetBulge 11, -k * 2
    plineObj.Update
    plineObj.Layer = "C-Mill"
    plineObj.Update
    plineObj.Closed = True
    
plineObj.Layer = "Ball-6"
plineObj.Update
  offsetObj = plineObj.Offset(-30)
plineObj.Layer = "C-Mill"
plineObj.Update
    
End If
If b >= 436 Then
plineObj.Delete
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointswithin436)

gptnz1 = Sqr((pointswithin436(4) - pointswithin436(2)) ^ 2 + (pointswithin436(5) - pointswithin436(3)) ^ 2)
gptnz2 = Sqr((pointswithin436(8) - pointswithin436(6)) ^ 2 + (pointswithin436(9) - pointswithin436(7)) ^ 2)
anglebulge1 = Atn((gptnz1 / 20) / Sqr(-(gptnz1 / 20) * (gptnz1 / 20) + 1))
anglebulge2 = Atn((gptnz2 / 114) / Sqr(-(gptnz2 / 114) * (gptnz2 / 114) + 1))
h = 20 * ((1 - Cos(anglebulge1)) / 2)
k = h / gptnz1
    plineObj.SetBulge 1, -k * 2
    plineObj.Update
    plineObj.SetBulge 8, -k * 2
    plineObj.Update
    plineObj.SetBulge 10, -k * 2
    plineObj.Update
    plineObj.SetBulge 17, -k * 2
    plineObj.Update
h = 114 * ((1 - Cos(anglebulge2)) / 2)
k = h / gptnz2
    plineObj.SetBulge 3, k * 2
    plineObj.Update
    plineObj.SetBulge 6, k * 2
    plineObj.Update
    plineObj.SetBulge 12, k * 2
    plineObj.Update
    plineObj.SetBulge 15, k * 2
    plineObj.Update
h = (2 * radius) * ((1 - Cos(anglebulge3)) / 2)
k = h / gptnz3
    plineObj.SetBulge 4, -k * 2
    plineObj.Update
    plineObj.SetBulge 5, -k * 2
    plineObj.Update
    plineObj.SetBulge 13, -k * 2
    plineObj.Update
    plineObj.SetBulge 14, -k * 2
    plineObj.Update
    plineObj.Layer = "C-Mill"
    plineObj.Update
    plineObj.Closed = True
    
plineObj.Layer = "Ball-6"
plineObj.Update
  offsetObj = plineObj.Offset(-30)
plineObj.Layer = "C-Mill"
plineObj.Update
    
End If
  
  
  
End If
End If
  points2(1) = points2(3) + I:
  points2(3) = points2(3) + a + I:
  points2(5) = points2(3):
  points2(7) = points2(1):
      
Else: d = 1
d = d + 1

d = Cells(c, 4)
End If

Next d
c = c + 1

a = Cells(c, 2)
b = Cells(c, 3)

I = 100

Loop
              
ZoomAll
AP.Quit
End Sub

Sub MF121()
' Creating search objects
Dim AP As Excel.Application
Dim WB As Excel.Workbook
Dim WS As Excel.Worksheet

Set AP = Excel.Application
Set WB = AP.Workbooks.Open("C:\Users\������\Desktop\������\2.xlsx")
Set WS = WB.Worksheets("������ �˨������ ������")
c = 19
a = Cells(c, 2)
b = Cells(c, 3)
d = Cells(c, 4)
E = 1
I = 0
On Error Resume Next

' Description of all the points that will make up rectangles
  Dim plineObj As AcadLWPolyline
  Dim points(0 To 7) As Double
  Dim pointswithin(0 To 15) As Double
  Dim pointsrec(0 To 15) As Double
  Dim pointsrecl(0 To 9) As Double
  Dim pointsrecr(0 To 9) As Double
  Dim pointswithin2(0 To 15) As Double
  Dim pointsrec2(0 To 7) As Double
  Dim a1(0 To 2) As Double
  Dim a2(0 To 2) As Double
  Dim A3(0 To 2) As Double
  Dim A4(0 To 2) As Double
  Dim P As Variant
  
  
points(6) = 0

Do While a > 0

' Calculate the coordinates of each point
  points(0) = points(6) + I:    points(1) = 0
  points(2) = points(6) + I:    points(3) = a
  points(4) = points(0) + b:    points(5) = a
  points(6) = points(0) + b:    points(7) = 0

' Constructing four segments and connecting the corresponding points

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
  plineObj.Closed = True
  r = 30
  pointswithin(0) = points(0) + 66 + r:   pointswithin(1) = 66
  pointswithin(2) = points(0) + 66:       pointswithin(3) = 66 + r
  pointswithin(4) = points(2) + 66:       pointswithin(5) = a - 66 - r
  pointswithin(6) = points(2) + 66 + r:   pointswithin(7) = a - 66
  pointswithin(8) = points(4) - 66 - r:   pointswithin(9) = a - 66
  pointswithin(10) = points(4) - 66:      pointswithin(11) = a - 66 - r
  pointswithin(12) = points(4) - 66:      pointswithin(13) = 66 + r
  pointswithin(14) = points(4) - 66 - r:  pointswithin(15) = 66

' Condition for limiting the filling of narrow facades
If a > 220 Then
If b > 220 Then
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointswithin)
  ' Find the bulge of the third segment
    Dim currentBulge As Double
    currentBulge = plineObj.GetBulge(2)
    l = 2 * r * (Sqr(2) / 2)
    h = r * (1 - (Sqr(2) / 2))
    k = h / (l / 2)
    ' Change the bulge of the third segment
    plineObj.SetBulge 0, k
    plineObj.Update
    plineObj.SetBulge 2, k
    plineObj.Update
    plineObj.SetBulge 4, k
    plineObj.Update
    plineObj.SetBulge 6, k
    plineObj.Update
    plineObj.Layer = "C-Mill"
    plineObj.Update
  plineObj.Closed = True
 
  pointsrecl(0) = points(4) - 96:      pointsrecl(1) = points(1) + 117.96
  pointsrecl(2) = points(4) - 117.96:  pointsrecl(3) = points(1) + 96
  pointsrecl(4) = points(0) + 117.96:  pointsrecl(5) = points(1) + 96
  pointsrecl(6) = points(0) + 96:      pointsrecl(7) = points(1) + 117.96
  pointsrecl(8) = points(0) + 96:      pointsrecl(9) = points(3) - 117.96
  
  pointsrecr(0) = points(0) + 96:      pointsrecr(1) = points(3) - 117.96
  pointsrecr(2) = points(0) + 117.96:  pointsrecr(3) = points(3) - 96
  pointsrecr(4) = points(4) - 117.96:  pointsrecr(5) = points(3) - 96
  pointsrecr(6) = points(4) - 96:      pointsrecr(7) = points(3) - 117.96
  pointsrecr(8) = points(4) - 96:      pointsrecr(9) = points(1) + 117.96
  
  gptnz = Sqr((pointsrecr(2) - pointsrecr(0)) ^ 2 + (pointsrecr(3) - pointsrecr(1)) ^ 2)
  anglebulge = Atn((gptnz / 60) / Sqr(-(gptnz / 60) * (gptnz / 60) + 1))
  h = 60 * ((1 - Cos(anglebulge)) / 2)
  k = h / gptnz

 Set plineObjl = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsrecl)
  plineObjl.Layer = "C-Mill"
    plineObjl.Update
    'plineObj.Closed = True
     plineObjl.SetBulge 0, k
    plineObjl.Update
     plineObjl.SetBulge 2, k
    plineObjl.Update
    Set plineObjr = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsrecr)
  plineObjr.Layer = "C-Mill"
    plineObjr.Update
    'plineObj.Closed = True
     plineObjr.SetBulge 0, k
    plineObjl.Update
     plineObjr.SetBulge 2, k
    plineObjr.Update
                                 'plineObj.color = acByLayer
                                 'Dim layerObj As AcadLayer
                                 'Set layerObj = ThisDrawing.Layers.Add("Ball-6.1")
                                 'layerObj.color = acRed
  acmill = a - 192
  bcmill = b - 192
  
  P = 0
  
   Dim lineObj As AcadLine
  a1(0) = pointsrecl(6) + (bcmill / 2) - 30:     a1(1) = (pointsrecl(5) + (acmill / 2)) - 50:           a1(2) = 0
  a2(0) = pointsrecl(6) + (bcmill / 2) + 30:     a2(1) = (pointsrecl(5) + (acmill / 2)) + 50:           a2(2) = 0
  
  
  Do While a1(1) < a - 117.96

  lineObj.Layer = "Ball-6"
  lineObj.Update
  Set lineObj = ThisDrawing.ModelSpace.AddLine(a1, a2)
  lineObj.Layer = "Ball-6"
  lineObj.Update

  intPointsl = lineObj.IntersectWith(plineObjl, acExtendThisEntity)
  intPointsr = lineObj.IntersectWith(plineObjr, acExtendThisEntity)
  lineObj.Delete
  a1(0) = intPointsl(0):     a1(1) = intPointsl(1)
  a2(0) = intPointsr(0):     a2(1) = intPointsr(1)
  
  Set lineObj = ThisDrawing.ModelSpace.AddLine(a1, a2)
  lineObj.Layer = "Ball-6"
  lineObj.Update
  
   a1(0) = intPointsl(0):     a1(1) = intPointsl(1)
   a2(0) = intPointsr(0):     a2(1) = intPointsr(1)

  a1(1) = a1(1) + 100
  a2(1) = a2(1) + 100
  Loop
  
  a1(0) = pointsrecl(6) + (bcmill / 2) - 30:     a1(1) = ((pointsrecl(5) + (acmill / 2)) - 50) - 100:         a1(2) = 0
  a2(0) = pointsrecl(6) + (bcmill / 2) + 30:     a2(1) = ((pointsrecl(5) + (acmill / 2)) + 50) - 100:         a2(2) = 0
   
  Do While a2(1) > 117.96
 
  lineObj.Layer = "Ball-6"
  lineObj.Update
  Set lineObj = ThisDrawing.ModelSpace.AddLine(a1, a2)
  lineObj.Layer = "Ball-6"
  lineObj.Update

  intPointsl = lineObj.IntersectWith(plineObjl, acExtendThisEntity)
  intPointsr = lineObj.IntersectWith(plineObjr, acExtendThisEntity)
  lineObj.Delete
  a1(0) = intPointsl(0):     a1(1) = intPointsl(1)
  a2(0) = intPointsr(0):     a2(1) = intPointsr(1)
  
  Set lineObj = ThisDrawing.ModelSpace.AddLine(a1, a2)
  lineObj.Layer = "Ball-6"
  lineObj.Update
  
  a1(0) = intPointsl(0):     a1(1) = intPointsl(1)
  a2(0) = intPointsr(0):     a2(1) = intPointsr(1)
  a1(1) = a1(1) - 100
  a2(1) = a2(1) - 100
  Loop
 
 plineObjl.Delete
 plineObjr.Delete
 
  pointsrecl(0) = points(0) + 96:      pointsrecl(1) = points(1) + 117.96
  pointsrecl(2) = points(0) + 96:      pointsrecl(3) = points(3) - 117.96
  pointsrecl(4) = points(0) + 117.96:  pointsrecl(5) = points(3) - 96
  pointsrecl(6) = points(4) - 117.96:  pointsrecl(7) = points(3) - 96
  pointsrecl(8) = points(4) - 96:      pointsrecl(9) = points(3) - 117.96
  
  pointsrecr(0) = points(4) - 96:      pointsrecr(1) = points(3) - 117.96
  pointsrecr(2) = points(4) - 96:      pointsrecr(3) = points(1) + 117.96
  pointsrecr(4) = points(4) - 117.96:  pointsrecr(5) = points(1) + 96
  pointsrecr(6) = points(0) + 117.96:  pointsrecr(7) = points(1) + 96
  pointsrecr(8) = points(0) + 96:      pointsrecr(9) = points(1) + 117.96
  
    Set plineObjl = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsrecl)
  plineObjl.Layer = "Ball-6"
    plineObjl.Update
    'plineObj.Closed = True
  plineObjl.SetBulge 1, k
    plineObjl.Update
  plineObjl.SetBulge 3, k
    plineObjl.Update
    
    Set plineObjr = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsrecr)
  plineObjr.Layer = "Ball-6"
    plineObjr.Update
    'plineObj.Closed = True
  plineObjr.SetBulge 1, k
     plineObjr.Update
  plineObjr.SetBulge 3, k
     plineObjr.Update
                                 'plineObj.color = acByLayer
                                 'Dim layerObj As AcadLayer
                                 'Set layerObj = ThisDrawing.Layers.Add("Ball-6.1")
                                 'layerObj.color = acRed
  acmill = a - 192
  bcmill = b - 192
  
   
  A3(0) = (points(0) + 96) + (bcmill / 2) - 30:     A3(1) = ((points(1) + 96) + (acmill / 2)) + 50:           A3(2) = 0
  A4(0) = (points(0) + 96) + (bcmill / 2) + 30:     A4(1) = ((points(1) + 96) + (acmill / 2)) - 50:           A4(2) = 0
  
  
  Do While A4(1) < a - 117.96

  lineObj.Layer = "Ball-6"
  lineObj.Update
  Set lineObj = ThisDrawing.ModelSpace.AddLine(A3, A4)
  lineObj.Layer = "Ball-6"
  lineObj.Update

  intPointsl = lineObj.IntersectWith(plineObjl, acExtendThisEntity)
  intPointsr = lineObj.IntersectWith(plineObjr, acExtendThisEntity)
  lineObj.Delete
  A3(0) = intPointsl(0):     A3(1) = intPointsl(1)
  A4(0) = intPointsr(0):     A4(1) = intPointsr(1)
  
  Set lineObj = ThisDrawing.ModelSpace.AddLine(A3, A4)
  lineObj.Layer = "Ball-6"
  lineObj.Update
  
   A3(0) = intPointsl(0):     A3(1) = intPointsl(1)
   A4(0) = intPointsr(0):     A4(1) = intPointsr(1)

  A3(1) = A3(1) + 100
  A4(1) = A4(1) + 100
  Loop
  
  A3(0) = (points(0) + 96) + (bcmill / 2) - 30:    A3(1) = (((points(1) + 96) + (acmill / 2)) + 50) - 100:         A3(2) = 0
  A4(0) = (points(0) + 96) + (bcmill / 2) + 30:    A4(1) = (((points(1) + 96) + (acmill / 2)) - 50) - 100:         A4(2) = 0
   
  Do While A3(1) > 117.96
 
  lineObj.Layer = "Ball-6"
  lineObj.Update
  Set lineObj = ThisDrawing.ModelSpace.AddLine(A3, A4)
  lineObj.Layer = "Ball-6"
  lineObj.Update

  intPointsl = lineObj.IntersectWith(plineObjl, acExtendThisEntity)
  intPointsr = lineObj.IntersectWith(plineObjr, acExtendThisEntity)
  lineObj.Delete
  A3(0) = intPointsl(0):     A3(1) = intPointsl(1)
  A4(0) = intPointsr(0):     A4(1) = intPointsr(1)
  
  Set lineObj = ThisDrawing.ModelSpace.AddLine(A3, A4)
  lineObj.Layer = "Ball-6"
  lineObj.Update
  
  A3(0) = intPointsl(0):     A3(1) = intPointsl(1)
  A4(0) = intPointsr(0):     A4(1) = intPointsr(1)
  A3(1) = A3(1) - 100
  A4(1) = A4(1) - 100
  Loop
 
 plineObjl.Delete
 plineObjr.Delete
 
 pointsrec(0) = points(0) + 96:       pointsrec(1) = points(1) + 117.96
 pointsrec(2) = points(0) + 96:       pointsrec(3) = points(3) - 117.96
 pointsrec(4) = points(0) + 117.96:   pointsrec(5) = points(3) - 96
 pointsrec(6) = points(4) - 117.96:   pointsrec(7) = points(3) - 96
 pointsrec(8) = points(4) - 96:       pointsrec(9) = points(3) - 117.96
 pointsrec(10) = points(4) - 96:      pointsrec(11) = points(1) + 117.96
 pointsrec(12) = points(4) - 117.96:  pointsrec(13) = points(1) + 96
 pointsrec(14) = points(0) + 117.96:  pointsrec(15) = points(1) + 96
 
   Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsrec)
  plineObj.Layer = "Ball-6"
    plineObj.SetBulge 1, k
    plineObj.Update
    plineObj.SetBulge 3, k
    plineObj.Update
    plineObj.SetBulge 5, k
    plineObj.Update
    plineObj.SetBulge 7, k
    plineObj.Update
    plineObj.Update
    plineObj.Closed = True
 
' End of the condition of limiting the filling of narrow facades
  End If
  End If

I = 100

 Dim plineObj2 As AcadLWPolyline
  Dim points2(0 To 7) As Double
  
  points2(0) = points(0):    points2(1) = a + I:
  points2(2) = points(0):    points2(3) = (a * 2) + I:
  points2(4) = points(6):    points2(5) = (a * 2) + I:
  points2(6) = points(6):    points2(7) = a + I:
  
 For d = 2 To Cells(c, 4)
 If d > 1 Then

Set plineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(points2)
  plineObj2.Closed = True

  pointswithin2(0) = points2(0) + 66 + r:   pointswithin2(1) = points2(1) + 66
  pointswithin2(2) = points2(0) + 66:       pointswithin2(3) = points2(1) + 66 + r
  pointswithin2(4) = points2(2) + 66:       pointswithin2(5) = points2(3) - 66 - r
  pointswithin2(6) = points2(2) + 66 + r:   pointswithin2(7) = points2(3) - 66
  pointswithin2(8) = points2(4) - 66 - r:   pointswithin2(9) = points2(3) - 66
  pointswithin2(10) = points2(4) - 66:      pointswithin2(11) = points2(3) - 66 - r
  pointswithin2(12) = points2(4) - 66:      pointswithin2(13) = points2(1) + 66 + r
  pointswithin2(14) = points2(4) - 66 - r:  pointswithin2(15) = points2(1) + 66

' Condition for limiting the filling of narrow facades
If a > 220 Then
If b > 220 Then
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointswithin2)
  ' Find the bulge of the third segment
    currentBulge = plineObj.GetBulge(2)
    l = 2 * r * (Sqr(2) / 2)
    h = r * (1 - (Sqr(2) / 2))
    k = h / (l / 2)
    ' Change the bulge of the third segment
    plineObj.SetBulge 0, k
    plineObj.Update
    plineObj.SetBulge 2, k
    plineObj.Update
    plineObj.SetBulge 4, k
    plineObj.Update
    plineObj.SetBulge 6, k
    plineObj.Update
    plineObj.Layer = "C-Mill"
    plineObj.Update
  plineObj.Closed = True
 
  pointsrecl(0) = points2(4) - 96:      pointsrecl(1) = points2(1) + 117.96
  pointsrecl(2) = points2(4) - 117.96:  pointsrecl(3) = points2(1) + 96
  pointsrecl(4) = points2(0) + 117.96:  pointsrecl(5) = points2(1) + 96
  pointsrecl(6) = points2(0) + 96:      pointsrecl(7) = points2(1) + 117.96
  pointsrecl(8) = points2(0) + 96:      pointsrecl(9) = points2(3) - 117.96
  
  pointsrecr(0) = points2(0) + 96:      pointsrecr(1) = points2(3) - 117.96
  pointsrecr(2) = points2(0) + 117.96:  pointsrecr(3) = points2(3) - 96
  pointsrecr(4) = points2(4) - 117.96:  pointsrecr(5) = points2(3) - 96
  pointsrecr(6) = points2(4) - 96:      pointsrecr(7) = points2(3) - 117.96
  pointsrecr(8) = points2(4) - 96:      pointsrecr(9) = points2(1) + 117.96
  
  gptnz = Sqr((pointsrecr(2) - pointsrecr(0)) ^ 2 + (pointsrecr(3) - pointsrecr(1)) ^ 2)
  anglebulge = Atn((gptnz / 60) / Sqr(-(gptnz / 60) * (gptnz / 60) + 1))
  h = 60 * ((1 - Cos(anglebulge)) / 2)
  k = h / gptnz

 Set plineObjl = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsrecl)
  plineObjl.Layer = "C-Mill"
    plineObjl.Update
    'plineObj.Closed = True
     plineObjl.SetBulge 0, k
    plineObjl.Update
     plineObjl.SetBulge 2, k
    plineObjl.Update
    Set plineObjr = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsrecr)
  plineObjr.Layer = "C-Mill"
    plineObjr.Update
    'plineObj.Closed = True
     plineObjr.SetBulge 0, k
    plineObjl.Update
     plineObjr.SetBulge 2, k
    plineObjr.Update
                                 'plineObj.color = acByLayer
                                 'Dim layerObj As AcadLayer
                                 'Set layerObj = ThisDrawing.Layers.Add("Ball-6.1")
                                 'layerObj.color = acRed
  acmill = a - 192
  bcmill = b - 192
  
  P = 0
  
   
  a1(0) = pointsrecl(6) + (bcmill / 2) - 30:     a1(1) = (pointsrecl(5) + (acmill / 2)) - 50:           a1(2) = 0
  a2(0) = pointsrecl(6) + (bcmill / 2) + 30:     a2(1) = (pointsrecl(5) + (acmill / 2)) + 50:           a2(2) = 0
  
  
  Do While a1(1) < points2(3) - 117.96

  lineObj.Layer = "Ball-6"
  lineObj.Update
  Set lineObj = ThisDrawing.ModelSpace.AddLine(a1, a2)
  lineObj.Layer = "Ball-6"
  lineObj.Update

  intPointsl = lineObj.IntersectWith(plineObjl, acExtendThisEntity)
  intPointsr = lineObj.IntersectWith(plineObjr, acExtendThisEntity)
  lineObj.Delete
  a1(0) = intPointsl(0):     a1(1) = intPointsl(1)
  a2(0) = intPointsr(0):     a2(1) = intPointsr(1)
  
  Set lineObj = ThisDrawing.ModelSpace.AddLine(a1, a2)
  lineObj.Layer = "Ball-6"
  lineObj.Update
  
   a1(0) = intPointsl(0):     a1(1) = intPointsl(1)
   a2(0) = intPointsr(0):     a2(1) = intPointsr(1)

  a1(1) = a1(1) + 100
  a2(1) = a2(1) + 100
  Loop
  
  a1(0) = pointsrecl(6) + (bcmill / 2) - 30:     a1(1) = ((pointsrecl(5) + (acmill / 2)) - 50) - 100:         a1(2) = 0
  a2(0) = pointsrecl(6) + (bcmill / 2) + 30:     a2(1) = ((pointsrecl(5) + (acmill / 2)) + 50) - 100:         a2(2) = 0
   
  Do While a2(1) > points2(1) + 117.96
 
  lineObj.Layer = "Ball-6"
  lineObj.Update
  Set lineObj = ThisDrawing.ModelSpace.AddLine(a1, a2)
  lineObj.Layer = "Ball-6"
  lineObj.Update

  intPointsl = lineObj.IntersectWith(plineObjl, acExtendThisEntity)
  intPointsr = lineObj.IntersectWith(plineObjr, acExtendThisEntity)
  lineObj.Delete
  a1(0) = intPointsl(0):     a1(1) = intPointsl(1)
  a2(0) = intPointsr(0):     a2(1) = intPointsr(1)
  
  Set lineObj = ThisDrawing.ModelSpace.AddLine(a1, a2)
  lineObj.Layer = "Ball-6"
  lineObj.Update
  
  a1(0) = intPointsl(0):     a1(1) = intPointsl(1)
  a2(0) = intPointsr(0):     a2(1) = intPointsr(1)
  a1(1) = a1(1) - 100
  a2(1) = a2(1) - 100
  Loop
 
 plineObjl.Delete
 plineObjr.Delete
 
  
  pointsrecl(0) = points2(0) + 96:      pointsrecl(1) = points2(1) + 117.96
  pointsrecl(2) = points2(0) + 96:      pointsrecl(3) = points2(3) - 117.96
  pointsrecl(4) = points2(0) + 117.96:  pointsrecl(5) = points2(3) - 96
  pointsrecl(6) = points2(4) - 117.96:  pointsrecl(7) = points2(3) - 96
  pointsrecl(8) = points2(4) - 96:      pointsrecl(9) = points2(3) - 117.96
  
  pointsrecr(0) = points2(4) - 96:      pointsrecr(1) = points2(3) - 117.96
  pointsrecr(2) = points2(4) - 96:      pointsrecr(3) = points2(1) + 117.96
  pointsrecr(4) = points2(4) - 117.96:  pointsrecr(5) = points2(1) + 96
  pointsrecr(6) = points2(0) + 117.96:  pointsrecr(7) = points2(1) + 96
  pointsrecr(8) = points2(0) + 96:      pointsrecr(9) = points2(1) + 117.96
  
    Set plineObjl = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsrecl)
  plineObjl.Layer = "Ball-6"
    plineObjl.Update
    'plineObj.Closed = True
  plineObjl.SetBulge 1, k
    plineObjl.Update
  plineObjl.SetBulge 3, k
    plineObjl.Update
    
    Set plineObjr = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsrecr)
  plineObjr.Layer = "Ball-6"
    plineObjr.Update
    'plineObj.Closed = True
  plineObjr.SetBulge 1, k
     plineObjr.Update
  plineObjr.SetBulge 3, k
     plineObjr.Update
                                 'plineObj.color = acByLayer
                                 'Dim layerObj As AcadLayer
                                 'Set layerObj = ThisDrawing.Layers.Add("Ball-6.1")
                                 'layerObj.color = acRed
  acmill = a - 192
  bcmill = b - 192
  
   
  A3(0) = (points2(0) + 96) + (bcmill / 2) - 30:     A3(1) = ((points2(1) + 96) + (acmill / 2)) + 50:           A3(2) = 0
  A4(0) = (points2(0) + 96) + (bcmill / 2) + 30:     A4(1) = ((points2(1) + 96) + (acmill / 2)) - 50:           A4(2) = 0
  
  
  Do While A4(1) < points2(3) - 117.96

  lineObj.Layer = "Ball-6"
  lineObj.Update
  Set lineObj = ThisDrawing.ModelSpace.AddLine(A3, A4)
  lineObj.Layer = "Ball-6"
  lineObj.Update

  intPointsl = lineObj.IntersectWith(plineObjl, acExtendThisEntity)
  intPointsr = lineObj.IntersectWith(plineObjr, acExtendThisEntity)
  lineObj.Delete
  A3(0) = intPointsl(0):     A3(1) = intPointsl(1)
  A4(0) = intPointsr(0):     A4(1) = intPointsr(1)
  
  Set lineObj = ThisDrawing.ModelSpace.AddLine(A3, A4)
  lineObj.Layer = "Ball-6"
  lineObj.Update
  
   A3(0) = intPointsl(0):     A3(1) = intPointsl(1)
   A4(0) = intPointsr(0):     A4(1) = intPointsr(1)

  A3(1) = A3(1) + 100
  A4(1) = A4(1) + 100
  Loop
  
  A3(0) = (points2(0) + 96) + (bcmill / 2) - 30:    A3(1) = (((points2(1) + 96) + (acmill / 2)) + 50) - 100:         A3(2) = 0
  A4(0) = (points2(0) + 96) + (bcmill / 2) + 30:    A4(1) = (((points2(1) + 96) + (acmill / 2)) - 50) - 100:         A4(2) = 0
   
  Do While A3(1) > points2(1) + 117.96
 
  lineObj.Layer = "Ball-6"
  lineObj.Update
  Set lineObj = ThisDrawing.ModelSpace.AddLine(A3, A4)
  lineObj.Layer = "Ball-6"
  lineObj.Update

  intPointsl = lineObj.IntersectWith(plineObjl, acExtendThisEntity)
  intPointsr = lineObj.IntersectWith(plineObjr, acExtendThisEntity)
  lineObj.Delete
  A3(0) = intPointsl(0):     A3(1) = intPointsl(1)
  A4(0) = intPointsr(0):     A4(1) = intPointsr(1)
  
  Set lineObj = ThisDrawing.ModelSpace.AddLine(A3, A4)
  lineObj.Layer = "Ball-6"
  lineObj.Update
  
  A3(0) = intPointsl(0):     A3(1) = intPointsl(1)
  A4(0) = intPointsr(0):     A4(1) = intPointsr(1)
  A3(1) = A3(1) - 100
  A4(1) = A4(1) - 100
  Loop
 
 plineObjl.Delete
 plineObjr.Delete
 
 pointsrec(0) = points2(0) + 96:       pointsrec(1) = points2(1) + 117.96
 pointsrec(2) = points2(0) + 96:       pointsrec(3) = points2(3) - 117.96
 pointsrec(4) = points2(0) + 117.96:   pointsrec(5) = points2(3) - 96
 pointsrec(6) = points2(4) - 117.96:   pointsrec(7) = points2(3) - 96
 pointsrec(8) = points2(4) - 96:       pointsrec(9) = points2(3) - 117.96
 pointsrec(10) = points2(4) - 96:      pointsrec(11) = points2(1) + 117.96
 pointsrec(12) = points2(4) - 117.96:  pointsrec(13) = points2(1) + 96
 pointsrec(14) = points2(0) + 117.96:  pointsrec(15) = points2(1) + 96
 
   Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsrec)
  plineObj.Layer = "Ball-6"
    plineObj.SetBulge 1, k
    plineObj.Update
    plineObj.SetBulge 3, k
    plineObj.Update
    plineObj.SetBulge 5, k
    plineObj.Update
    plineObj.SetBulge 7, k
    plineObj.Update
    plineObj.Update
    plineObj.Closed = True

  End If
  End If
 
  points2(1) = points2(3) + I:
  points2(3) = points2(3) + a + I:
  points2(5) = points2(3):
  points2(7) = points2(1):
      
Else: d = 1
d = d + 1

d = Cells(c, 4)
End If

Next d
c = c + 1

a = Cells(c, 2)
b = Cells(c, 3)

I = 100

Loop
              
ZoomAll
AP.Quit
End Sub

Sub MF122()
' Creating search objects
Dim AP As Excel.Application
Dim WB As Excel.Workbook
Dim WS As Excel.Worksheet

Set AP = Excel.Application
Set WB = AP.Workbooks.Open("C:\Users\������\Desktop\������\2.xlsx")
Set WS = WB.Worksheets("������ �˨������ ������")
c = 19
a = Cells(c, 2)
b = Cells(c, 3)
d = Cells(c, 4)
E = 1
I = 0
On Error Resume Next

' Description of all the points that will make up rectangles
 Dim plineObj As AcadLWPolyline
  Dim points(0 To 7) As Double
  Dim pointswithin(0 To 15) As Double
points(6) = 0

Do While a > 0

' Calculate the coordinates of each point
  points(0) = points(6) + I:    points(1) = 0
  points(2) = points(6) + I:    points(3) = a
  points(4) = points(0) + b:    points(5) = a
  points(6) = points(0) + b:    points(7) = 0

' Constructing four segments and connecting the corresponding points

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
  plineObj.Closed = True
  
                                 'plineObj.color = acByLayer
                                 'Dim layerObj As AcadLayer
                                 'Set layerObj = ThisDrawing.Layers.Add("Ball-6.1")
                                 'layerObj.color = acRed
If a > 220 Then
If b > 220 Then

  pointswithin(0) = points(0) + 66:    pointswithin(1) = points(1) + 106
  pointswithin(2) = points(2) + 66:    pointswithin(3) = points(3) - 106
  pointswithin(4) = points(2) + 106:    pointswithin(5) = points(3) - 66
  pointswithin(6) = points(4) - 106:    pointswithin(7) = points(3) - 66
  pointswithin(8) = points(4) - 66:    pointswithin(9) = points(3) - 106
  pointswithin(10) = points(4) - 66:   pointswithin(11) = points(1) + 106
  pointswithin(12) = points(4) - 106:    pointswithin(13) = points(1) + 66
  pointswithin(14) = points(0) + 106:    pointswithin(15) = points(1) + 66
  
  Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointswithin)

 ' Find the bulge of the third segment
    currentBulge = plineObj.GetBulge(2)
   k = (pointswithin(4) - pointswithin(2)) / 2
    ' Change the bulge of the third segment
    plineObj.SetBulge 1, (24 / 57.9412)
    plineObj.Update
    plineObj.SetBulge 3, (24 / 57.9412)
    plineObj.Update
    plineObj.SetBulge 5, (24 / 57.9412)
    plineObj.Update
    plineObj.SetBulge 7, (24 / 57.9412)
    plineObj.Update
    plineObj.Layer = "C-Mill"
    plineObj.Update
  plineObj.Closed = True

End If
End If

I = 100
  
 Dim plineObj2 As AcadLWPolyline
  Dim points2(0 To 7) As Double
  
  points2(0) = points(0):    points2(1) = a + I:
  points2(2) = points(0):    points2(3) = (a * 2) + I:
  points2(4) = points(6):    points2(5) = (a * 2) + I:
  points2(6) = points(6):    points2(7) = a + I:
  
 For d = 2 To Cells(c, 4)
If d > 1 Then

Set plineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(points2)
  plineObj2.Closed = True
  
  If a > 220 Then
  If b > 220 Then
  
  pointswithin(0) = points2(0) + 66:      pointswithin(1) = points2(1) + 106
  pointswithin(2) = points2(2) + 66:      pointswithin(3) = points2(3) - 106
  pointswithin(4) = points2(2) + 106:     pointswithin(5) = points2(3) - 66
  pointswithin(6) = points2(4) - 106:     pointswithin(7) = points2(3) - 66
  pointswithin(8) = points2(4) - 66:      pointswithin(9) = points2(3) - 106
  pointswithin(10) = points2(4) - 66:     pointswithin(11) = points2(1) + 106
  pointswithin(12) = points2(4) - 106:    pointswithin(13) = points2(1) + 66
  pointswithin(14) = points2(0) + 106:    pointswithin(15) = points2(1) + 66
  
  Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointswithin)

 ' Find the bulge of the third segment
    currentBulge = plineObj.GetBulge(2)
    ' Change the bulge of the third segment
    plineObj.SetBulge 1, (24 / 57.9412)
    plineObj.Update
    plineObj.SetBulge 3, (24 / 57.9412)
    plineObj.Update
    plineObj.SetBulge 5, (24 / 57.9412)
    plineObj.Update
    plineObj.SetBulge 7, (24 / 57.9412)
    plineObj.Update
    plineObj.Layer = "C-Mill"
    plineObj.Update
    plineObj.Closed = True
    
    End If
    End If
  
  points2(1) = points2(3) + I:
  points2(3) = points2(3) + a + I:
  points2(5) = points2(3):
  points2(7) = points2(1):
  

Else: d = 1
d = d + 1

d = Cells(c, 4)
End If

Next d
c = c + 1

a = Cells(c, 2)
b = Cells(c, 3)

I = 100

Loop
              
ZoomAll
AP.Quit
End Sub

Sub MF123()
' Creating search objects
Dim AP As Excel.Application
Dim WB As Excel.Workbook
Dim WS As Excel.Worksheet

Set AP = Excel.Application
Set WB = AP.Workbooks.Open("C:\Users\������\Desktop\������\2.xlsx")
Set WS = WB.Worksheets("������ �˨������ ������")
c = 19
a = Cells(c, 2)
b = Cells(c, 3)
d = Cells(c, 4)
E = 1
I = 0
On Error Resume Next

' Description of all the points that will make up rectangles
  Dim plineObj As AcadLWPolyline
  Dim points(0 To 7) As Double
  Dim pointswithin(0 To 15) As Double
  Dim pointsrec(0 To 7) As Double
  Dim pointsrecl(0 To 5) As Double
  Dim pointsrecr(0 To 5) As Double
  Dim pointswithin2(0 To 15) As Double
  Dim pointsrec2(0 To 7) As Double
  Dim a1(0 To 2) As Double
  Dim a2(0 To 2) As Double
  Dim A3(0 To 2) As Double
  Dim A4(0 To 2) As Double
  Dim P As Variant
  
points(6) = 0

Do While a > 0

' Calculate the coordinates of each point
  points(0) = points(6) + I:    points(1) = 0
  points(2) = points(6) + I:    points(3) = a
  points(4) = points(0) + b:    points(5) = a
  points(6) = points(0) + b:    points(7) = 0

' Constructing four segments and connecting the corresponding points

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
  plineObj.Closed = True
  r = 36
  pointswithin(0) = points(0) + 30 + r:   pointswithin(1) = 30
  pointswithin(2) = points(0) + 30:       pointswithin(3) = 30 + r
  pointswithin(4) = points(2) + 30:       pointswithin(5) = a - 30 - r
  pointswithin(6) = points(2) + 30 + r:   pointswithin(7) = a - 30
  pointswithin(8) = points(4) - 30 - r:   pointswithin(9) = a - 30
  pointswithin(10) = points(4) - 30:      pointswithin(11) = a - 30 - r
  pointswithin(12) = points(4) - 30:      pointswithin(13) = 30 + r
  pointswithin(14) = points(4) - 30 - r:  pointswithin(15) = 30

' Condition for limiting the filling of narrow facades
If a > 132 Then
If b > 132 Then
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointswithin)
  ' Find the bulge of the third segment
    Dim currentBulge As Double
    currentBulge = plineObj.GetBulge(2)
    l = 2 * r * (Sqr(2) / 2)
    h = r * (1 - (Sqr(2) / 2))
    k = h / (l / 2)
    ' Change the bulge of the third segment
    plineObj.SetBulge 0, -k
    plineObj.Update
    plineObj.SetBulge 2, -k
    plineObj.Update
    plineObj.SetBulge 4, -k
    plineObj.Update
    plineObj.SetBulge 6, -k
    plineObj.Update
    plineObj.Layer = "Ball-6"
    plineObj.Update
  plineObj.Closed = True
 
  pointsrecl(0) = points(4) - 66:      pointsrecl(1) = 66
  pointsrecl(2) = points(0) + 66:      pointsrecl(3) = 66
  pointsrecl(4) = points(0) + 66:      pointsrecl(5) = a - 66
  
  pointsrecr(0) = points(0) + 66:      pointsrecr(1) = a - 66
  pointsrecr(2) = points(4) - 66:      pointsrecr(3) = a - 66
  pointsrecr(4) = points(4) - 66:      pointsrecr(5) = 66
  
  
  
 Set plineObjl = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsrecl)
  plineObjl.Layer = "C-Mill"
    plineObjl.Update
    'plineObj.Closed = True
    Set plineObjr = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsrecr)
  plineObjr.Layer = "C-Mill"
    plineObjr.Update
    'plineObj.Closed = True
                                 'plineObj.color = acByLayer
                                 'Dim layerObj As AcadLayer
                                 'Set layerObj = ThisDrawing.Layers.Add("Ball-6.1")
                                 'layerObj.color = acRed
  acmill = a - 132
  bcmill = b - 132
  
  P = 0
  
   Dim lineObj As AcadLine
  a1(0) = pointsrecl(2) + (bcmill / 2) - 21.75:     a1(1) = (pointsrecl(3) + (acmill / 2)) - 21.75:           a1(2) = 0
  a2(0) = pointsrecl(2) + (bcmill / 2) + 21.75:     a2(1) = (pointsrecl(3) + (acmill / 2)) + 21.75:           a2(2) = 0
  
  
  Do While a1(1) < a - 66

  lineObj.Layer = "C-Mill"
  lineObj.Update
  Set lineObj = ThisDrawing.ModelSpace.AddLine(a1, a2)
  lineObj.Layer = "C-Mill"
  lineObj.Update

  intPointsl = lineObj.IntersectWith(plineObjl, acExtendThisEntity)
  intPointsr = lineObj.IntersectWith(plineObjr, acExtendThisEntity)
  lineObj.Delete
  a1(0) = intPointsl(0):     a1(1) = intPointsl(1)
  a2(0) = intPointsr(0):     a2(1) = intPointsr(1)
  
  Set lineObj = ThisDrawing.ModelSpace.AddLine(a1, a2)
  lineObj.Layer = "C-Mill"
  lineObj.Update
  
   a1(0) = intPointsl(0):     a1(1) = intPointsl(1)
   a2(0) = intPointsr(0):     a2(1) = intPointsr(1)

  a1(1) = a1(1) + 87
  a2(1) = a2(1) + 87
  Loop
  
  a1(0) = pointsrecl(2) + (bcmill / 2) - 21.75:     a1(1) = ((pointsrecl(3) + (acmill / 2)) - 21.75) - 87:         a1(2) = 0
  a2(0) = pointsrecl(2) + (bcmill / 2) + 21.75:     a2(1) = ((pointsrecl(3) + (acmill / 2)) + 21.75) - 87:         a2(2) = 0
   
  Do While a2(1) > 66
 
  lineObj.Layer = "C-Mill"
  lineObj.Update
  Set lineObj = ThisDrawing.ModelSpace.AddLine(a1, a2)
  lineObj.Layer = "C-Mill"
  lineObj.Update

  intPointsl = lineObj.IntersectWith(plineObjl, acExtendThisEntity)
  intPointsr = lineObj.IntersectWith(plineObjr, acExtendThisEntity)
  lineObj.Delete
  a1(0) = intPointsl(0):     a1(1) = intPointsl(1)
  a2(0) = intPointsr(0):     a2(1) = intPointsr(1)
  
  Set lineObj = ThisDrawing.ModelSpace.AddLine(a1, a2)
  lineObj.Layer = "C-Mill"
  lineObj.Update
  
  a1(0) = intPointsl(0):     a1(1) = intPointsl(1)
  a2(0) = intPointsr(0):     a2(1) = intPointsr(1)
  
  a1(1) = a1(1) - 87
  a2(1) = a2(1) - 87

  Loop
 
 plineObjl.Delete
 plineObjr.Delete
 
  pointsrecl(0) = points(0) + 66:      pointsrecl(1) = 66
  pointsrecl(2) = points(0) + 66:      pointsrecl(3) = a - 66
  pointsrecl(4) = points(4) - 66:      pointsrecl(5) = a - 66
  
  pointsrecr(0) = points(4) - 66:      pointsrecr(1) = a - 66
  pointsrecr(2) = points(4) - 66:      pointsrecr(3) = 66
  pointsrecr(4) = points(0) + 66:      pointsrecr(5) = 66
   
  
    Set plineObjl = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsrecl)
  plineObjl.Layer = "C-Mill"
    plineObjl.Update
    'plineObj.Closed = True
    Set plineObjr = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsrecr)
  plineObjr.Layer = "C-Mill"
    plineObjr.Update
    'plineObj.Closed = True
                                 'plineObj.color = acByLayer
                                 'Dim layerObj As AcadLayer
                                 'Set layerObj = ThisDrawing.Layers.Add("Ball-6.1")
                                 'layerObj.color = acRed
  acmill = a - 132
  bcmill = b - 132
  
   
  A3(0) = pointsrecl(0) + (bcmill / 2) - 21.75:     A3(1) = (pointsrecl(1) + (acmill / 2)) + 21.75:           A3(2) = 0
  A4(0) = pointsrecl(0) + (bcmill / 2) + 21.75:     A4(1) = (pointsrecl(1) + (acmill / 2)) - 21.75:           A4(2) = 0
  
  
  Do While A4(1) < a - 66

  lineObj.Layer = "C-Mill"
  lineObj.Update
  Set lineObj = ThisDrawing.ModelSpace.AddLine(A3, A4)
  lineObj.Layer = "C-Mill"
  lineObj.Update

  intPointsl = lineObj.IntersectWith(plineObjl, acExtendThisEntity)
  intPointsr = lineObj.IntersectWith(plineObjr, acExtendThisEntity)
  lineObj.Delete
  A3(0) = intPointsl(0):     A3(1) = intPointsl(1)
  A4(0) = intPointsr(0):     A4(1) = intPointsr(1)
  
  Set lineObj = ThisDrawing.ModelSpace.AddLine(A3, A4)
  lineObj.Layer = "C-Mill"
  lineObj.Update
  
   A3(0) = intPointsl(0):     A3(1) = intPointsl(1)
   A4(0) = intPointsr(0):     A4(1) = intPointsr(1)

  A3(1) = A3(1) + 87
  A4(1) = A4(1) + 87
  Loop
  
  A3(0) = pointsrecl(0) + (bcmill / 2) - 21.75:     A3(1) = ((pointsrecl(1) + (acmill / 2)) + 21.75) - 87:         A3(2) = 0
  A4(0) = pointsrecl(0) + (bcmill / 2) + 21.75:     A4(1) = ((pointsrecl(1) + (acmill / 2)) - 21.75) - 87:         A4(2) = 0
   
  Do While A3(1) > 66
 
  lineObj.Layer = "C-Mill"
  lineObj.Update
  Set lineObj = ThisDrawing.ModelSpace.AddLine(A3, A4)
  lineObj.Layer = "C-Mill"
  lineObj.Update

  intPointsl = lineObj.IntersectWith(plineObjl, acExtendThisEntity)
  intPointsr = lineObj.IntersectWith(plineObjr, acExtendThisEntity)
  lineObj.Delete
  A3(0) = intPointsl(0):     A3(1) = intPointsl(1)
  A4(0) = intPointsr(0):     A4(1) = intPointsr(1)
  
  Set lineObj = ThisDrawing.ModelSpace.AddLine(A3, A4)
  lineObj.Layer = "C-Mill"
  lineObj.Update
  
  A3(0) = intPointsl(0):     A3(1) = intPointsl(1)
  A4(0) = intPointsr(0):     A4(1) = intPointsr(1)
  A3(1) = A3(1) - 87
  A4(1) = A4(1) - 87
  Loop
 
 plineObjl.Delete
 plineObjr.Delete
 
 pointsrec(0) = points(0) + 66:      pointsrec(1) = 66
 pointsrec(2) = points(0) + 66:      pointsrec(3) = a - 66
 pointsrec(4) = points(4) - 66:      pointsrec(5) = a - 66
 pointsrec(6) = points(4) - 66:      pointsrec(7) = 66
   
   Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsrec)
  plineObj.Layer = "C-Mill"
    plineObj.Update
    plineObj.Closed = True
' End of the condition of limiting the filling of narrow facades
  End If
  End If
I = 100
  
 Dim plineObj2 As AcadLWPolyline
  Dim points2(0 To 7) As Double
  
  points2(0) = points(0):    points2(1) = a + I:
  points2(2) = points(0):    points2(3) = (a * 2) + I:
  points2(4) = points(6):    points2(5) = (a * 2) + I:
  points2(6) = points(6):    points2(7) = a + I:
  
 For d = 2 To Cells(c, 4)
 If d > 1 Then

Set plineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(points2)
  plineObj2.Closed = True

  pointswithin2(0) = points2(0) + 30 + r:   pointswithin2(1) = points2(1) + 30
  pointswithin2(2) = points2(0) + 30:       pointswithin2(3) = points2(1) + 30 + r
  pointswithin2(4) = points2(2) + 30:       pointswithin2(5) = points2(3) - 30 - r
  pointswithin2(6) = points2(2) + 30 + r:   pointswithin2(7) = points2(3) - 30
  pointswithin2(8) = points2(4) - 30 - r:   pointswithin2(9) = points2(3) - 30
  pointswithin2(10) = points2(4) - 30:      pointswithin2(11) = points2(3) - 30 - r
  pointswithin2(12) = points2(4) - 30:      pointswithin2(13) = points2(1) + 30 + r
  pointswithin2(14) = points2(4) - 30 - r:  pointswithin2(15) = points2(1) + 30

' Condition for limiting the filling of narrow facades
If a > 132 Then
If b > 132 Then
  ' Find the bulge of the third segment
  Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointswithin2)
    currentBulge = plineObj.GetBulge(2)
    ' Change the bulge of the third segment
    plineObj.SetBulge 0, -k
    plineObj.Update
    plineObj.SetBulge 2, -k
    plineObj.Update
    plineObj.SetBulge 4, -k
    plineObj.Update
    plineObj.SetBulge 6, -k
    plineObj.Update
    plineObj.Layer = "Ball-6"
    plineObj.Update
  plineObj.Closed = True
  
  

  pointsrecl(0) = points2(4) - 66:      pointsrecl(1) = points2(1) + 66
  pointsrecl(2) = points2(0) + 66:      pointsrecl(3) = points2(1) + 66
  pointsrecl(4) = points2(0) + 66:      pointsrecl(5) = points2(3) - 66
  
  pointsrecr(0) = points2(0) + 66:      pointsrecr(1) = points2(3) - 66
  pointsrecr(2) = points2(4) - 66:      pointsrecr(3) = points2(3) - 66
  pointsrecr(4) = points2(4) - 66:      pointsrecr(5) = points2(1) + 66
  
  
  
    Set plineObjl = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsrecl)
  plineObjl.Layer = "C-Mill"
    plineObjl.Update
    'plineObj.Closed = True
    Set plineObjr = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsrecr)
  plineObjr.Layer = "C-Mill"
    plineObjr.Update
    'plineObj.Closed = True
                                 'plineObj.color = acByLayer
                                 'Dim layerObj As AcadLayer
                                 'Set layerObj = ThisDrawing.Layers.Add("Ball-6.1")
                                 'layerObj.color = acRed
  acmill = a - 132
  bcmill = b - 132
  
  
   
  a1(0) = pointsrecl(2) + (bcmill / 2) - 21.75:     a1(1) = (pointsrecl(3) + (acmill / 2)) - 21.75:           a1(2) = 0
  a2(0) = pointsrecl(2) + (bcmill / 2) + 21.75:     a2(1) = (pointsrecl(3) + (acmill / 2)) + 21.75:           a2(2) = 0
  
  
  Do While a1(1) < ((a * d) + ((I * d) - I)) - 66

  lineObj.Layer = "C-Mill"
  lineObj.Update
  Set lineObj = ThisDrawing.ModelSpace.AddLine(a1, a2)
  lineObj.Layer = "C-Mill"
  lineObj.Update

  intPointsl = lineObj.IntersectWith(plineObjl, acExtendThisEntity)
  intPointsr = lineObj.IntersectWith(plineObjr, acExtendThisEntity)
  lineObj.Delete
  a1(0) = intPointsl(0):     a1(1) = intPointsl(1)
  a2(0) = intPointsr(0):     a2(1) = intPointsr(1)
  
  Set lineObj = ThisDrawing.ModelSpace.AddLine(a1, a2)
  lineObj.Layer = "C-Mill"
  lineObj.Update
  
   a1(0) = intPointsl(0):     a1(1) = intPointsl(1)
   a2(0) = intPointsr(0):     a2(1) = intPointsr(1)

  a1(1) = a1(1) + 87
  a2(1) = a2(1) + 87
  Loop
  
  a1(0) = pointsrecl(2) + (bcmill / 2) - 21.75:     a1(1) = ((pointsrecl(3) + (acmill / 2)) - 21.75) - 87:         a1(2) = 0
  a2(0) = pointsrecl(2) + (bcmill / 2) + 21.75:     a2(1) = ((pointsrecl(3) + (acmill / 2)) + 21.75) - 87:         a2(2) = 0
   
  Do While a2(1) > (((a * d) + ((I * d) - I))) - a + 66
 
  lineObj.Layer = "C-Mill"
  lineObj.Update
  Set lineObj = ThisDrawing.ModelSpace.AddLine(a1, a2)
  lineObj.Layer = "C-Mill"
  lineObj.Update

  intPointsl = lineObj.IntersectWith(plineObjl, acExtendThisEntity)
  intPointsr = lineObj.IntersectWith(plineObjr, acExtendThisEntity)
  lineObj.Delete
  a1(0) = intPointsl(0):     a1(1) = intPointsl(1)
  a2(0) = intPointsr(0):     a2(1) = intPointsr(1)
  
  Set lineObj = ThisDrawing.ModelSpace.AddLine(a1, a2)
  lineObj.Layer = "C-Mill"
  lineObj.Update
  
  a1(0) = intPointsl(0):     a1(1) = intPointsl(1)
  a2(0) = intPointsr(0):     a2(1) = intPointsr(1)
  a1(1) = a1(1) - 87
  a2(1) = a2(1) - 87
  Loop
 
 plineObjl.Delete
 plineObjr.Delete
 
  pointsrecl(0) = points(0) + 66:      pointsrecl(1) = points2(1) + 66
  pointsrecl(2) = points(0) + 66:      pointsrecl(3) = points2(3) - 66
  pointsrecl(4) = points(4) - 66:      pointsrecl(5) = points2(3) - 66
  
  pointsrecr(0) = points(4) - 66:      pointsrecr(1) = points2(3) - 66
  pointsrecr(2) = points(4) - 66:      pointsrecr(3) = points2(1) + 66
  pointsrecr(4) = points(0) + 66:      pointsrecr(5) = points2(1) + 66
   
   
 Set plineObjl = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsrecl)
  plineObjl.Layer = "C-Mill"
    plineObjl.Update
    'plineObj.Closed = True
    Set plineObjr = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsrecr)
  plineObjr.Layer = "C-Mill"
    plineObjr.Update
    'plineObj.Closed = True
                                 'plineObj.color = acByLayer
                                 'Dim layerObj As AcadLayer
                                 'Set layerObj = ThisDrawing.Layers.Add("Ball-6.1")
                                 'layerObj.color = acRed
  acmill = a - 132
  bcmill = b - 132
  
   
  A3(0) = pointsrecl(0) + (bcmill / 2) - 21.75:     A3(1) = (pointsrecl(1) + (acmill / 2)) + 21.75:           A3(2) = 0
  A4(0) = pointsrecl(0) + (bcmill / 2) + 21.75:     A4(1) = (pointsrecl(1) + (acmill / 2)) - 21.75:           A4(2) = 0
  
 
  
  P = points2(1) - 0
  q = points2(3) - 0
  
  Do While A4(1) < ((a * d) + ((I * d) - I)) - 66

  lineObj.Layer = "C-Mill"
  lineObj.Update
  Set lineObj = ThisDrawing.ModelSpace.AddLine(A3, A4)
  lineObj.Layer = "C-Mill"
  lineObj.Update

  intPointsl = lineObj.IntersectWith(plineObjl, acExtendThisEntity)
  intPointsr = lineObj.IntersectWith(plineObjr, acExtendThisEntity)
  lineObj.Delete
  A3(0) = intPointsl(0):     A3(1) = intPointsl(1)
  A4(0) = intPointsr(0):     A4(1) = intPointsr(1)
  
  Set lineObj = ThisDrawing.ModelSpace.AddLine(A3, A4)
  lineObj.Layer = "C-Mill"
  lineObj.Update
  
   A3(0) = intPointsl(0):     A3(1) = intPointsl(1)
   A4(0) = intPointsr(0):     A4(1) = intPointsr(1)

  A3(1) = A3(1) + 87
  A4(1) = A4(1) + 87
  Loop
  
  A3(0) = pointsrecl(0) + (bcmill / 2) - 21.75:     A3(1) = ((pointsrecl(1) + (acmill / 2)) + 21.75) - 87:         A3(2) = 0
  A4(0) = pointsrecl(0) + (bcmill / 2) + 21.75:     A4(1) = ((pointsrecl(1) + (acmill / 2)) - 21.75) - 87:         A4(2) = 0
   
  Do While A3(1) > (((a * d) + ((I * d) - I))) - a + 66
 
  lineObj.Layer = "C-Mill"
  lineObj.Update
  Set lineObj = ThisDrawing.ModelSpace.AddLine(A3, A4)
  lineObj.Layer = "C-Mill"
  lineObj.Update

  intPointsl = lineObj.IntersectWith(plineObjl, acExtendThisEntity)
  intPointsr = lineObj.IntersectWith(plineObjr, acExtendThisEntity)
  lineObj.Delete
  A3(0) = intPointsl(0):     A3(1) = intPointsl(1)
  A4(0) = intPointsr(0):     A4(1) = intPointsr(1)
  
  Set lineObj = ThisDrawing.ModelSpace.AddLine(A3, A4)
  lineObj.Layer = "C-Mill"
  lineObj.Update
  
  A3(0) = intPointsl(0):     A3(1) = intPointsl(1)
  A4(0) = intPointsr(0):     A4(1) = intPointsr(1)
  A3(1) = A3(1) - 87
  A4(1) = A4(1) - 87
  Loop
 
 plineObjl.Delete
 plineObjr.Delete
 
 pointsrec(0) = points(0) + 66:      pointsrec(1) = points2(1) + 66
 pointsrec(2) = points(0) + 66:      pointsrec(3) = points2(3) - 66
 pointsrec(4) = points(4) - 66:      pointsrec(5) = points2(3) - 66
 pointsrec(6) = points(4) - 66:      pointsrec(7) = points2(1) + 66
   
   Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsrec)
  plineObj.Layer = "C-Mill"
    plineObj.Update
    plineObj.Closed = True
    
' End of the condition of limiting the filling of narrow facades
End If
End If
 
  points2(1) = points2(3) + I:
  points2(3) = points2(3) + a + I:
  points2(5) = points2(3):
  points2(7) = points2(1):
      
Else: d = 1
d = d + 1

d = Cells(c, 4)
End If

Next d
c = c + 1

a = Cells(c, 2)
b = Cells(c, 3)

I = 100

Loop
              
ZoomAll
AP.Quit
End Sub

Sub MF124()
' Creating search objects
Dim AP As Excel.Application
Dim WB As Excel.Workbook
Dim WS As Excel.Worksheet

Set AP = Excel.Application
Set WB = AP.Workbooks.Open("C:\Users\������\Desktop\������\2.xlsx")
Set WS = WB.Worksheets("������ �˨������ ������")
c = 19
a = Cells(c, 2)
b = Cells(c, 3)
d = Cells(c, 4)
E = 1
I = 0
On Error Resume Next

' Description of all the points that will make up rectangles
  Dim plineObj As AcadLWPolyline
  Dim points(0 To 7) As Double
  Dim pointswithin(0 To 7) As Double
  Dim pointswithin2(0 To 7) As Double
points(6) = 0
 
  Dim a1(0 To 2) As Double
  Dim a2(0 To 2) As Double
  Dim A3(0 To 2) As Double
  Dim A4(0 To 2) As Double
  Dim A5(0 To 2) As Double
  Dim A6(0 To 2) As Double
  

Do While a > 0

' Calculate the coordinates of each point
  points(0) = points(6) + I:    points(1) = 0
  points(2) = points(6) + I:    points(3) = a
  points(4) = points(0) + b:    points(5) = a
  points(6) = points(0) + b:    points(7) = 0

' Constructing four segments and connecting the corresponding points

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
  plineObj.Closed = True
 
 
  pointswithin(0) = points(0) + 65:    pointswithin(1) = 65
  pointswithin(2) = points(2) + 65:    pointswithin(3) = a - 89
  pointswithin(4) = points(4) - 65:    pointswithin(5) = a - 89
  pointswithin(6) = points(6) - 65:    pointswithin(7) = 65
    
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointswithin)

    Dim currentBulge As Double
    currentBulge = plineObj.GetBulge(2)
   k = (pointswithin(4) - pointswithin(2)) / 2
    ' Set the convexity of the 1st segment
    plineObj.SetBulge 1, -(24 / k)
    plineObj.Update
    plineObj.Layer = "C-Mill"
    plineObj.Update
  plineObj.Closed = True

 Dim lineObj As AcadLine
  
  a1(0) = points(0) + 47: a1(1) = 0:      a1(2) = 0
  a2(0) = points(2) + 47: a2(1) = a:      a2(2) = 0
  A3(0) = points(4) - 47: A3(1) = a:      A3(2) = 0
  A4(0) = points(6) - 47: A4(1) = 0:      A4(2) = 0
 
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(a1, a2)
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A3, A4)
lineObj.Layer = "Ball-6"
lineObj.Update
l = b - 94
k = l / 50
m = Round(k, 0)
n = l / m



A5(0) = a1(0) + n
A6(1) = a - 65:


   

Do While A5(0) < A3(0)

Dim intPoints As Variant


A5(0) = A5(0): A5(1) = 65:        A5(2) = 0
A6(0) = A5(0): A6(1) = a - 65:     A6(2) = 0

Set lineObj = ThisDrawing.ModelSpace.AddLine(A5, A6)
lineObj.Layer = "Ball-6"
lineObj.Update
 
intPoints = lineObj.IntersectWith(plineObj, acExtendNone)
lineObj.Delete

A6(1) = intPoints(1)

Set lineObj = ThisDrawing.ModelSpace.AddLine(A5, A6)
lineObj.Layer = "Ball-6"
lineObj.Update

A5(0) = A5(0) + n
A6(0) = A6(0) + n


Loop



                                 'plineObj.color = acByLayer
                                 'Dim layerObj As AcadLayer
                                 'Set layerObj = ThisDrawing.Layers.Add("Ball-6.1")
                                 'layerObj.color = acRed


I = 100
  
 Dim plineObj2 As AcadLWPolyline
  Dim points2(0 To 7) As Double
  
  points2(0) = points(0):    points2(1) = a + I:
  points2(2) = points(0):    points2(3) = (a * 2) + I:
  points2(4) = points(6):    points2(5) = (a * 2) + I:
  points2(6) = points(6):    points2(7) = a + I:
  
 For d = 2 To Cells(c, 4)
If d > 1 Then

Set plineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(points2)
  plineObj2.Closed = True

  pointswithin2(0) = points2(0) + 65:    pointswithin2(1) = points2(1) + 65
  pointswithin2(2) = points2(2) + 65:    pointswithin2(3) = points2(3) - 89
  pointswithin2(4) = points2(4) - 65:    pointswithin2(5) = points2(5) - 89
  pointswithin2(6) = points2(6) - 65:    pointswithin2(7) = points2(7) + 65
  
  
  Set plineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointswithin2)
    currentBulge = plineObj2.GetBulge(2)
   k = (pointswithin(4) - pointswithin(2)) / 2
    ' Set the convexity for the 1st segment
    plineObj2.SetBulge 1, -(24 / k)
    plineObj2.Update
    plineObj2.Layer = "C-Mill"
    plineObj2.Update
  plineObj2.Closed = True
  
  a1(0) = points2(0) + 47: a1(1) = points2(1):     a1(2) = 0
  a2(0) = points2(2) + 47: a2(1) = points2(3):     a2(2) = 0
  A3(0) = points2(4) - 47: A3(1) = points2(5):     A3(2) = 0
  A4(0) = points2(6) - 47: A4(1) = points2(7):     A4(2) = 0
  
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(a1, a2)
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A3, A4)
lineObj.Layer = "Ball-6"
lineObj.Update


A5(0) = a1(0) + n
A6(1) = a - 65:

Do While A5(0) < A3(0)

A5(0) = A5(0): A5(1) = points2(1) + 65:     A5(2) = 0
A6(0) = A5(0): A6(1) = points2(3) - 65:     A6(2) = 0

Set lineObj = ThisDrawing.ModelSpace.AddLine(A5, A6)
lineObj.Layer = "Ball-6"
lineObj.Update
 
intPoints = lineObj.IntersectWith(plineObj2, acExtendNone)
lineObj.Delete

A6(1) = intPoints(1)

Set lineObj = ThisDrawing.ModelSpace.AddLine(A5, A6)
lineObj.Layer = "Ball-6"
lineObj.Update

A5(0) = A5(0) + n
A6(0) = A6(0) + n

Loop

  points2(1) = points2(3) + I:
  points2(3) = points2(3) + a + I:
  points2(5) = points2(3):
  points2(7) = points2(1):
      
Else: d = 1
d = d + 1

d = Cells(c, 4)
End If

Next d
c = c + 1

a = Cells(c, 2)
b = Cells(c, 3)

I = 100

Loop
              
ZoomAll
AP.Quit
End Sub

Sub MF129()
' Creating search objects
Dim AP As Excel.Application
Dim WB As Excel.Workbook
Dim WS As Excel.Worksheet

Set AP = Excel.Application
Set WB = AP.Workbooks.Open("C:\Users\������\Desktop\������\2.xlsx")
Set WS = WB.Worksheets("������ �˨������ ������")
c = 19
a = Cells(c, 2)
b = Cells(c, 3)
d = Cells(c, 4)
E = 1
I = 0
On Error Resume Next

' Description of all the points that will make up rectangles
 Dim plineObj As AcadLWPolyline
  Dim points(0 To 7) As Double
points(6) = 0

Do While a > 0

' Calculate the coordinates of each point
  points(0) = points(6) + I:    points(1) = 0
  points(2) = points(6) + I:    points(3) = a
  points(4) = points(0) + b:    points(5) = a
  points(6) = points(0) + b:    points(7) = 0

' Constructing four segments and connecting the corresponding points

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
  plineObj.Closed = True
  
  ' Offset the polyline
Dim offsetObj As Variant
                                 'plineObj.color = acByLayer
                                 'Dim layerObj As AcadLayer
                                 'Set layerObj = ThisDrawing.Layers.Add("Ball-6.1")
                                 'layerObj.color = acRed


plineObj.Layer = "Ball-6"
plineObj.Update
  offsetObj = plineObj.Offset(35)
plineObj.Layer = "K-grav"
plineObj.Update
  offsetObj = plineObj.Offset(45)
plineObj.Layer = "C-Mill"
plineObj.Update
  offsetObj = plineObj.Offset(65)
plineObj.Layer = "0"
plineObj.Update

I = 100
  
 Dim plineObj2 As AcadLWPolyline
  Dim points2(0 To 7) As Double
  
  points2(0) = points(0):    points2(1) = a + I:
  points2(2) = points(0):    points2(3) = (a * 2) + I:
  points2(4) = points(6):    points2(5) = (a * 2) + I:
  points2(6) = points(6):    points2(7) = a + I:
  
 For d = 2 To Cells(c, 4)
If d > 1 Then

Set plineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(points2)
  plineObj2.Closed = True
   
  points2(1) = points2(3) + I:
  points2(3) = points2(3) + a + I:
  points2(5) = points2(3):
  points2(7) = points2(1):
  
 
  
  ' Offset the polyline

plineObj2.Layer = "Ball-6"
plineObj2.Update
  offsetObj = plineObj2.Offset(35)
plineObj2.Layer = "K-grav"
plineObj2.Update
  offsetObj = plineObj2.Offset(45)
plineObj2.Layer = "C-Mill"
plineObj2.Update
  offsetObj = plineObj2.Offset(65)
plineObj2.Layer = "0"
plineObj2.Update

Else: d = 1
d = d + 1

d = Cells(c, 4)
End If

Next d
c = c + 1

a = Cells(c, 2)
b = Cells(c, 3)

I = 100

Loop
              
ZoomAll
AP.Quit
End Sub

Sub MF130()
' Creating search objects
Dim AP As Excel.Application
Dim WB As Excel.Workbook
Dim WS As Excel.Worksheet

Set AP = Excel.Application
Set WB = AP.Workbooks.Open("C:\Users\������\Desktop\������\2.xlsx")
Set WS = WB.Worksheets("������ �˨������ ������")
c = 19
a = Cells(c, 2)
b = Cells(c, 3)
d = Cells(c, 4)
E = 1
I = 0
On Error Resume Next

' Description of all the points that will make up rectangles
 Dim plineObj As AcadLWPolyline
  Dim points(0 To 7) As Double
points(6) = 0

Do While a > 0

' Calculate the coordinates of each point
  points(0) = points(6) + I:    points(1) = 0
  points(2) = points(6) + I:    points(3) = a
  points(4) = points(0) + b:    points(5) = a
  points(6) = points(0) + b:    points(7) = 0

' Constructing four segments and connecting the corresponding points

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
  plineObj.Closed = True
  
  ' Offset the polyline
Dim offsetObj As Variant
                                 'plineObj.color = acByLayer
                                 'Dim layerObj As AcadLayer
                                 'Set layerObj = ThisDrawing.Layers.Add("Ball-6.1")
                                 'layerObj.color = acRed

plineObj.Layer = "K-grav"
plineObj.Update
  offsetObj = plineObj.Offset(50)
plineObj.Layer = "Ball-6"
plineObj.Update
  offsetObj = plineObj.Offset(51)
plineObj.Layer = "K-Mill"
plineObj.Update
  offsetObj = plineObj.Offset(68)
  offsetObj = plineObj.Offset(69)
  offsetObj = plineObj.Offset(70)
plineObj.Layer = "D-Mill"
plineObj.Update
  offsetObj = plineObj.Offset(74)
plineObj.Layer = "0"
plineObj.Update

I = 100
  
 Dim plineObj2 As AcadLWPolyline
  Dim points2(0 To 7) As Double
  
  points2(0) = points(0):    points2(1) = a + I:
  points2(2) = points(0):    points2(3) = (a * 2) + I:
  points2(4) = points(6):    points2(5) = (a * 2) + I:
  points2(6) = points(6):    points2(7) = a + I:
  
 For d = 2 To Cells(c, 4)
If d > 1 Then

Set plineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(points2)
  plineObj2.Closed = True
   
  points2(1) = points2(3) + I:
  points2(3) = points2(3) + a + I:
  points2(5) = points2(3):
  points2(7) = points2(1):
  
 
  
  ' Offset the polyline

plineObj2.Layer = "K-grav"
plineObj2.Update
  offsetObj = plineObj2.Offset(50)
plineObj2.Layer = "Ball-6"
plineObj2.Update
  offsetObj = plineObj2.Offset(51)
plineObj2.Layer = "K-Mill"
plineObj2.Update
  offsetObj = plineObj2.Offset(68)
  offsetObj = plineObj2.Offset(69)
  offsetObj = plineObj2.Offset(70)
plineObj2.Layer = "D-Mill"
plineObj2.Update
  offsetObj = plineObj2.Offset(74)
plineObj2.Layer = "0"
plineObj2.Update

Else: d = 1
d = d + 1

d = Cells(c, 4)
End If

Next d
c = c + 1

a = Cells(c, 2)
b = Cells(c, 3)

I = 100

Loop
              
ZoomAll
AP.Quit
End Sub

Sub MF131()
' Creating search objects
Dim AP As Excel.Application
Dim WB As Excel.Workbook
Dim WS As Excel.Worksheet

Set AP = Excel.Application
Set WB = AP.Workbooks.Open("C:\Users\������\Desktop\������\2.xlsx")
Set WS = WB.Worksheets("������ �˨������ ������")
c = 19
a = Cells(c, 2)
b = Cells(c, 3)
d = Cells(c, 4)
E = 1
I = 0
On Error Resume Next

' Description of all the points that will make up rectangles
 Dim plineObj As AcadLWPolyline
  Dim points(0 To 7) As Double
points(6) = 0

Do While a > 0

' Calculate the coordinates of each point
  points(0) = points(6) + I:    points(1) = 0
  points(2) = points(6) + I:    points(3) = a
  points(4) = points(0) + b:    points(5) = a
  points(6) = points(0) + b:    points(7) = 0

' Constructing four segments and connecting the corresponding points

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
  plineObj.Closed = True
  
Dim a1(0 To 2) As Double
Dim a2(0 To 2) As Double
Dim A3(0 To 2) As Double
Dim A4(0 To 2) As Double
Dim A5(0 To 2) As Double
Dim A6(0 To 2) As Double
Dim A7(0 To 2) As Double
Dim A8(0 To 2) As Double
Dim lineObj As AcadLine
  
  a1(0) = points(0):     a1(1) = 0:      a1(2) = 0
  a2(0) = points(0):     a2(1) = a:      a2(2) = 0
  A3(0) = points(6):     A3(1) = a:      A3(2) = 0
  A4(0) = points(6):     A4(1) = 0:      A4(2) = 0
  
  A5(0) = points(0) + 53: A5(1) = points(1) + 53:    A5(2) = 0
  A6(0) = points(2) + 53: A6(1) = points(3) - 53:    A6(2) = 0
  A7(0) = points(4) - 53: A7(1) = points(5) - 53:    A7(2) = 0
  A8(0) = points(6) - 53: A8(1) = points(7) + 53:    A8(2) = 0

   
If a > 100 Then
If b > 100 Then

lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(a1, A5)
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(a2, A6)
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A3, A7)
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A4, A8)
 
  ' Offset the polyline
  Dim offsetObj As Variant
                                 'plineObj.color = acByLayer
                                 'Dim layerObj As AcadLayer
                                 'Set layerObj = ThisDrawing.Layers.Add("Ball-6.1")
                                 'layerObj.color = acRed

plineObj.Layer = "K-grav"
plineObj.Update
  offsetObj = plineObj.Offset(50)
plineObj.Layer = "K-Mill"
plineObj.Update
  offsetObj = plineObj.Offset(68)
  offsetObj = plineObj.Offset(69)
  offsetObj = plineObj.Offset(70)
plineObj.Layer = "D-Mill"
plineObj.Update
  offsetObj = plineObj.Offset(74)
plineObj.Layer = "0"
plineObj.Update

End If
End If

I = 100
  
 Dim plineObj2 As AcadLWPolyline
  Dim points2(0 To 7) As Double
  
  points2(0) = points(0):    points2(1) = a + I:
  points2(2) = points(0):    points2(3) = (a * 2) + I:
  points2(4) = points(6):    points2(5) = (a * 2) + I:
  points2(6) = points(6):    points2(7) = a + I:
  
 For d = 2 To Cells(c, 4)
If d > 1 Then

Set plineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(points2)
  plineObj2.Closed = True
  
  a1(0) = points2(0):     a1(1) = points2(1):      a1(2) = 0
  a2(0) = points2(2):     a2(1) = points2(3):      a2(2) = 0
  A3(0) = points2(4):     A3(1) = points2(5):      A3(2) = 0
  A4(0) = points2(6):     A4(1) = points2(7):      A4(2) = 0
  
  A5(0) = points2(0) + 53: A5(1) = points2(1) + 53:    A5(2) = 0
  A6(0) = points2(2) + 53: A6(1) = points2(3) - 53:    A6(2) = 0
  A7(0) = points2(4) - 53: A7(1) = points2(5) - 53:    A7(2) = 0
  A8(0) = points2(6) - 53: A8(1) = points2(7) + 53:    A8(2) = 0
   
If a > 100 Then
If b > 100 Then

lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(a1, A5)
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(a2, A6)
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A3, A7)
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A4, A8)
lineObj.Layer = "Ball-6"
lineObj.Update

  ' Offset the polyline

plineObj2.Layer = "K-grav"
plineObj2.Update
  offsetObj = plineObj2.Offset(50)
plineObj2.Layer = "K-Mill"
plineObj2.Update
  offsetObj = plineObj2.Offset(68)
  offsetObj = plineObj2.Offset(69)
  offsetObj = plineObj2.Offset(70)
plineObj2.Layer = "D-Mill"
plineObj2.Update
  offsetObj = plineObj2.Offset(74)
plineObj2.Layer = "0"
plineObj2.Update

End If
End If

  points2(1) = points2(3) + I:
  points2(3) = points2(3) + a + I:
  points2(5) = points2(3):
  points2(7) = points2(1):


Else: d = 1
d = d + 1

d = Cells(c, 4)
End If

Next d
c = c + 1

a = Cells(c, 2)
b = Cells(c, 3)

I = 100

Loop
              
ZoomAll
AP.Quit
End Sub

Sub MF132()
' Creating search objects
Dim AP As Excel.Application
Dim WB As Excel.Workbook
Dim WS As Excel.Worksheet

Set AP = Excel.Application
Set WB = AP.Workbooks.Open("C:\Users\������\Desktop\������\2.xlsx")
Set WS = WB.Worksheets("������ �˨������ ������")
c = 19
a = Cells(c, 2)
b = Cells(c, 3)
d = Cells(c, 4)
E = 1
I = 0
On Error Resume Next

' Description of all the points that will make up rectangles
  Dim plineObj As AcadLWPolyline
  Dim points(0 To 7) As Double
points(6) = 0

Do While a > 0

' Calculate the coordinates of each point
  points(0) = points(6) + I:    points(1) = 0
  points(2) = points(6) + I:    points(3) = a
  points(4) = points(0) + b:    points(5) = a
  points(6) = points(0) + b:    points(7) = 0

' Constructing four segments and connecting the corresponding points

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
  plineObj.Closed = True
 
Dim a1(0 To 2) As Double
Dim a2(0 To 2) As Double
Dim A3(0 To 2) As Double
Dim A4(0 To 2) As Double
Dim A5(0 To 2) As Double
Dim A6(0 To 2) As Double
Dim A7(0 To 2) As Double
Dim A8(0 To 2) As Double
Dim lineObj As AcadLine
  
  a1(0) = points(0) + 51: a1(1) = 0:      a1(2) = 0
  a2(0) = points(2) + 51: a2(1) = a:      a2(2) = 0
  A3(0) = points(4) - 51: A3(1) = a:      A3(2) = 0
  A4(0) = points(6) - 51: A4(1) = 0:      A4(2) = 0
  
  A5(0) = points(0) + 51: A5(1) = 51:      A5(2) = 0
  A6(0) = points(2) + 51: A6(1) = a - 51:  A6(2) = 0
  A7(0) = points(4) - 51: A7(1) = a - 51:  A7(2) = 0
  A8(0) = points(6) - 51: A8(1) = 51:      A8(2) = 0
   
If a > 100 Then
If b > 100 Then
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(a1, a2)
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A6, A7)
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A3, A4)
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A8, A5)
lineObj.Layer = "Ball-6"
lineObj.Update

  ' Offset the polyline
  Dim offsetObj As Variant
 
                                 'plineObj.color = acByLayer
                                 'Dim layerObj As AcadLayer
                                 'Set layerObj = ThisDrawing.Layers.Add("Ball-6.1")
                                 'layerObj.color = acRed
plineObj.Layer = "K-grav"
plineObj.Update
  offsetObj = plineObj.Offset(50)
plineObj.Layer = "K-Mill"
plineObj.Update
  offsetObj = plineObj.Offset(66)
  offsetObj = plineObj.Offset(67)
  offsetObj = plineObj.Offset(68)
  offsetObj = plineObj.Offset(69)
plineObj.Layer = "0"
plineObj.Update

End If
End If

I = 100
  
 Dim plineObj2 As AcadLWPolyline
  Dim points2(0 To 7) As Double
  
  points2(0) = points(0):    points2(1) = a + I:
  points2(2) = points(0):    points2(3) = (a * 2) + I:
  points2(4) = points(6):    points2(5) = (a * 2) + I:
  points2(6) = points(6):    points2(7) = a + I:
  
 For d = 2 To Cells(c, 4)
If d > 1 Then

Set plineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(points2)
  plineObj2.Closed = True
If a > 100 Then
If b > 100 Then
  
  a1(0) = points2(0) + 51: a1(1) = points2(1):     a1(2) = 0
  a2(0) = points2(2) + 51: a2(1) = points2(3):     a2(2) = 0
  A3(0) = points2(4) - 51: A3(1) = points2(5):     A3(2) = 0
  A4(0) = points2(6) - 51: A4(1) = points2(7):     A4(2) = 0
  
  A5(0) = points2(0) + 51: A5(1) = points2(1) + 51:    A5(2) = 0
  A6(0) = points2(2) + 51: A6(1) = points2(3) - 51:    A6(2) = 0
  A7(0) = points2(4) - 51: A7(1) = points2(5) - 51:    A7(2) = 0
  A8(0) = points2(6) - 51: A8(1) = points2(7) + 51:    A8(2) = 0
  
     


lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(a1, a2)
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A6, A7)
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A3, A4)
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A5, A8)
lineObj.Layer = "Ball-6"
lineObj.Update
End If
End If

  points2(1) = points2(3) + I:
  points2(3) = points2(3) + a + I:
  points2(5) = points2(3):
  points2(7) = points2(1):
  
  ' Offset the polyline
plineObj2.Layer = "K-grav"
plineObj2.Update
  offsetObj = plineObj2.Offset(50)
plineObj2.Layer = "K-Mill"
plineObj2.Update
  offsetObj = plineObj2.Offset(66)
  offsetObj = plineObj2.Offset(67)
  offsetObj = plineObj2.Offset(68)
  offsetObj = plineObj2.Offset(69)
plineObj2.Layer = "0"
plineObj2.Update



Else: d = 1
d = d + 1

d = Cells(c, 4)
End If

Next d
c = c + 1

a = Cells(c, 2)
b = Cells(c, 3)

I = 100

Loop
              
ZoomAll
AP.Quit
End Sub

Sub MF133()
' Creating search objects
Dim AP As Excel.Application
Dim WB As Excel.Workbook
Dim WS As Excel.Worksheet

Set AP = Excel.Application
Set WB = AP.Workbooks.Open("C:\Users\������\Desktop\������\2.xlsx")
Set WS = WB.Worksheets("������ �˨������ ������")
c = 19
a = Cells(c, 2)
b = Cells(c, 3)
d = Cells(c, 4)
E = 1
I = 0
On Error Resume Next

' Description of all the points that will make up rectangles
 Dim plineObj As AcadLWPolyline
  Dim points(0 To 7) As Double
points(6) = 0

Do While a > 0

' Calculate the coordinates of each point
  points(0) = points(6) + I:    points(1) = 0
  points(2) = points(6) + I:    points(3) = a
  points(4) = points(0) + b:    points(5) = a
  points(6) = points(0) + b:    points(7) = 0

' Constructing four segments and connecting the corresponding points

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
  plineObj.Closed = True
If a > 184 Then
If b > 184 Then
  ' Offset the polyline
  Dim offsetObj As Variant
                                 'plineObj.color = acByLayer
                                 'Dim layerObj As AcadLayer
                                 'Set layerObj = ThisDrawing.Layers.Add("Ball-6.1")
                                 'layerObj.color = acRed

plineObj.Layer = "K-Mill"
plineObj.Update
  offsetObj = plineObj.Offset(66)
plineObj.Layer = "0"
plineObj.Update
  
  Dim droppointsld(0 To 7) As Double
  
  droppointsld(0) = points(0) + 11.26:      droppointsld(1) = points(1) + 11.82
  droppointsld(2) = points(0) + 53.01:      droppointsld(3) = points(1) + 74.63
  droppointsld(4) = points(0) + 74.63:      droppointsld(5) = points(1) + 53.01
  droppointsld(6) = points(0) + 11.82:      droppointsld(7) = points(1) + 11.26
  
    
  Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(droppointsld)
    plineObj.Layer = "K-grav"
    plineObj.Update
  ' Find the bulge of the third segment
    Dim currentBulge As Double
    currentBulge = plineObj.GetBulge(2)
    ' Change the bulge of the third segment
    plineObj.SetBulge 1, -1.222097
    plineObj.Update
    plineObj.SetBulge 3, -0.817647
    plineObj.Update
    plineObj.Layer = "K-grav"
    plineObj.Update
  plineObj.Closed = True
  
    Dim droppointslu(0 To 7) As Double
  
  droppointslu(0) = points(0) + 11.82:      droppointslu(1) = points(3) - 11.26
  droppointslu(2) = points(0) + 74.63:      droppointslu(3) = points(3) - 53.01
  droppointslu(4) = points(0) + 53.01:      droppointslu(5) = points(3) - 74.63
  droppointslu(6) = points(0) + 11.26:      droppointslu(7) = points(3) - 11.82
  
  Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(droppointslu)
  ' Find the bulge of the third segment
    currentBulge = plineObj.GetBulge(2)
    ' Change the bulge of the third segment
    plineObj.SetBulge 1, -1.222097
    plineObj.Update
    plineObj.SetBulge 3, -0.817647
    plineObj.Update
    plineObj.Layer = "K-grav"
    plineObj.Update
  plineObj.Closed = True
  
  Dim droppointsru(0 To 7) As Double
  
  droppointsru(0) = points(4) - 11.82:      droppointsru(1) = points(5) - 11.26
  droppointsru(2) = points(4) - 74.63:      droppointsru(3) = points(5) - 53.01
  droppointsru(4) = points(4) - 53.01:      droppointsru(5) = points(5) - 74.63
  droppointsru(6) = points(4) - 11.26:      droppointsru(7) = points(5) - 11.82
  
  Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(droppointsru)
  ' Find the bulge of the third segment
    currentBulge = plineObj.GetBulge(2)
    ' Change the bulge of the third segment
    plineObj.SetBulge 1, 1.222097
    plineObj.Update
    plineObj.SetBulge 3, 0.817647
    plineObj.Update
    plineObj.Layer = "K-grav"
    plineObj.Update
  plineObj.Closed = True
  
 Dim droppointsrd(0 To 7) As Double
  
  droppointsrd(0) = points(4) - 11.26:      droppointsrd(1) = points(7) + 11.82
  droppointsrd(2) = points(4) - 53.01:      droppointsrd(3) = points(7) + 74.63
  droppointsrd(4) = points(4) - 74.63:      droppointsrd(5) = points(7) + 53.01
  droppointsrd(6) = points(4) - 11.82:      droppointsrd(7) = points(7) + 11.26
  
    
  Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(droppointsrd)
    plineObj.Layer = "K-grav"
    plineObj.Update
  ' Find the bulge of the third segment
    currentBulge = plineObj.GetBulge(2)
    ' Change the bulge of the third segment
    plineObj.SetBulge 1, 1.222097
    plineObj.Update
    plineObj.SetBulge 3, 0.817647
    plineObj.Update
    plineObj.Layer = "K-grav"
    plineObj.Update
  plineObj.Closed = True
End If
End If
  
  

I = 100
  
 Dim plineObj2 As AcadLWPolyline
  Dim points2(0 To 7) As Double
  
  points2(0) = points(0):    points2(1) = a + I:
  points2(2) = points(0):    points2(3) = (a * 2) + I:
  points2(4) = points(6):    points2(5) = (a * 2) + I:
  points2(6) = points(6):    points2(7) = a + I:
  
 For d = 2 To Cells(c, 4)
If d > 1 Then

Set plineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(points2)
  plineObj2.Closed = True
If a > 184 Then
If b > 184 Then
  ' Offset the polyline
plineObj2.Layer = "K-Mill"
plineObj2.Update
  offsetObj = plineObj2.Offset(66)
plineObj2.Layer = "0"
plineObj2.Update
  
  droppointsld(0) = points2(0) + 11.26:      droppointsld(1) = points2(1) + 11.82
  droppointsld(2) = points2(0) + 53.01:      droppointsld(3) = points2(1) + 74.63
  droppointsld(4) = points2(0) + 74.63:      droppointsld(5) = points2(1) + 53.01
  droppointsld(6) = points2(0) + 11.82:      droppointsld(7) = points2(1) + 11.26
  
    
  Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(droppointsld)
    plineObj.Layer = "K-grav"
    plineObj.Update
  ' Find the bulge of the third segment
    currentBulge = plineObj.GetBulge(2)
    ' Change the bulge of the third segment
    plineObj.SetBulge 1, -1.222097
    plineObj.Update
    plineObj.SetBulge 3, -0.817647
    plineObj.Update
    plineObj.Layer = "K-grav"
    plineObj.Update
  plineObj.Closed = True
  
  droppointslu(0) = points2(0) + 11.82:      droppointslu(1) = points2(3) - 11.26
  droppointslu(2) = points2(0) + 74.63:      droppointslu(3) = points2(3) - 53.01
  droppointslu(4) = points2(0) + 53.01:      droppointslu(5) = points2(3) - 74.63
  droppointslu(6) = points2(0) + 11.26:      droppointslu(7) = points2(3) - 11.82
  
  Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(droppointslu)
  ' Find the bulge of the third segment
    currentBulge = plineObj.GetBulge(2)
    ' Change the bulge of the third segment
    plineObj.SetBulge 1, -1.222097
    plineObj.Update
    plineObj.SetBulge 3, -0.817647
    plineObj.Update
    plineObj.Layer = "K-grav"
    plineObj.Update
  plineObj.Closed = True
  
  droppointsru(0) = points2(4) - 11.82:      droppointsru(1) = points2(5) - 11.26
  droppointsru(2) = points2(4) - 74.63:      droppointsru(3) = points2(5) - 53.01
  droppointsru(4) = points2(4) - 53.01:      droppointsru(5) = points2(5) - 74.63
  droppointsru(6) = points2(4) - 11.26:      droppointsru(7) = points2(5) - 11.82
  
  Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(droppointsru)
  ' Find the bulge of the third segment
    currentBulge = plineObj.GetBulge(2)
    ' Change the bulge of the third segment
    plineObj.SetBulge 1, 1.222097
    plineObj.Update
    plineObj.SetBulge 3, 0.817647
    plineObj.Update
    plineObj.Layer = "K-grav"
    plineObj.Update
  plineObj.Closed = True

  droppointsrd(0) = points2(4) - 11.26:      droppointsrd(1) = points2(7) + 11.82
  droppointsrd(2) = points2(4) - 53.01:      droppointsrd(3) = points2(7) + 74.63
  droppointsrd(4) = points2(4) - 74.63:      droppointsrd(5) = points2(7) + 53.01
  droppointsrd(6) = points2(4) - 11.82:      droppointsrd(7) = points2(7) + 11.26
  
    
  Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(droppointsrd)
    plineObj.Layer = "K-grav"
    plineObj.Update
  ' Find the bulge of the third segment
    currentBulge = plineObj.GetBulge(2)
    ' Change the bulge of the third segment
    plineObj.SetBulge 1, 1.222097
    plineObj.Update
    plineObj.SetBulge 3, 0.817647
    plineObj.Update
    plineObj.Layer = "K-grav"
    plineObj.Update
  plineObj.Closed = True
End If
End If
  
  points2(1) = points2(3) + I:
  points2(3) = points2(3) + a + I:
  points2(5) = points2(3):
  points2(7) = points2(1):
  
  




Else: d = 1
d = d + 1

d = Cells(c, 4)
End If

Next d
c = c + 1

a = Cells(c, 2)
b = Cells(c, 3)

I = 100

Loop
              
ZoomAll
AP.Quit
End Sub

Sub MF134()
' Creating search objects
Dim AP As Excel.Application
Dim WB As Excel.Workbook
Dim WS As Excel.Worksheet

Set AP = Excel.Application
Set WB = AP.Workbooks.Open("C:\Users\������\Desktop\������\2.xlsx")
Set WS = WB.Worksheets("������ �˨������ ������")
c = 19
a = Cells(c, 2)
b = Cells(c, 3)
d = Cells(c, 4)
E = 1
I = 0
On Error Resume Next

' Description of all the points that will make up rectangles
 Dim plineObj As AcadLWPolyline
  Dim points(0 To 7) As Double
points(6) = 0

Do While a > 0

' Calculate the coordinates of each point
  points(0) = points(6) + I:    points(1) = 0
  points(2) = points(6) + I:    points(3) = a
  points(4) = points(0) + b:    points(5) = a
  points(6) = points(0) + b:    points(7) = 0

' Constructing four segments and connecting the corresponding points

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
  plineObj.Closed = True
  
  ' Offset the polyline
Dim offsetObj As Variant
                                 'plineObj.color = acByLayer
                                 'Dim layerObj As AcadLayer
                                 'Set layerObj = ThisDrawing.Layers.Add("Ball-6.1")
                                 'layerObj.color = acRed

plineObj.Layer = "K-grav"
plineObj.Update
  offsetObj = plineObj.Offset(40)
  offsetObj = plineObj.Offset(46.5)
  offsetObj = plineObj.Offset(58.5)
  offsetObj = plineObj.Offset(75.5)
plineObj.Layer = "0"
plineObj.Update

I = 100
  
 Dim plineObj2 As AcadLWPolyline
  Dim points2(0 To 7) As Double
  
  points2(0) = points(0):    points2(1) = a + I:
  points2(2) = points(0):    points2(3) = (a * 2) + I:
  points2(4) = points(6):    points2(5) = (a * 2) + I:
  points2(6) = points(6):    points2(7) = a + I:
  
 For d = 2 To Cells(c, 4)
If d > 1 Then

Set plineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(points2)
  plineObj2.Closed = True
   
  points2(1) = points2(3) + I:
  points2(3) = points2(3) + a + I:
  points2(5) = points2(3):
  points2(7) = points2(1):
  
 
  
  ' Offset the polyline

plineObj2.Layer = "K-grav"
plineObj2.Update
  offsetObj = plineObj2.Offset(40)
  offsetObj = plineObj2.Offset(46.5)
  offsetObj = plineObj2.Offset(58.5)
  offsetObj = plineObj2.Offset(75.5)
plineObj2.Layer = "0"
plineObj2.Update

Else: d = 1
d = d + 1

d = Cells(c, 4)
End If

Next d
c = c + 1

a = Cells(c, 2)
b = Cells(c, 3)

I = 100

Loop
              
ZoomAll
AP.Quit
End Sub

Sub MF135()
' Creating search objects
Dim AP As Excel.Application
Dim WB As Excel.Workbook
Dim WS As Excel.Worksheet

Set AP = Excel.Application
Set WB = AP.Workbooks.Open("C:\Users\������\Desktop\������\2.xlsx")
Set WS = WB.Worksheets("������ �˨������ ������")
c = 19
a = Cells(c, 2)
b = Cells(c, 3)
d = Cells(c, 4)
E = 1
I = 0
On Error Resume Next

' Description of all the points that will make up rectangles
  Dim plineObj As AcadLWPolyline
  Dim points(0 To 7) As Double
  Dim pointswithin(0 To 13) As Double
  Dim pointswithin2(0 To 13) As Double
  Dim P As Variant
  
points(6) = 0

Do While a > 0

' Calculate the coordinates of each point
  points(0) = points(6) + I:    points(1) = 0
  points(2) = points(6) + I:    points(3) = a
  points(4) = points(0) + b:    points(5) = a
  points(6) = points(0) + b:    points(7) = 0

' Constructing four segments and connecting the corresponding points

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
  plineObj.Closed = True
  P = (b - 90) / 4
  pointswithin(0) = points(0) + 45:           pointswithin(1) = points(1) + 45
  pointswithin(2) = points(2) + 45:           pointswithin(3) = points(3) - 69
  pointswithin(4) = points(2) + 45 + P:       pointswithin(5) = points(3) - 57
  pointswithin(6) = points(2) + 45 + (2 * P): pointswithin(7) = points(3) - 45
  pointswithin(8) = points(4) - 45 - P:       pointswithin(9) = points(3) - 57
  pointswithin(10) = points(4) - 45:          pointswithin(11) = points(3) - 69
  pointswithin(12) = points(4) - 45:          pointswithin(13) = points(7) + 45
  
If a > 180 Then
If b > 180 Then
 
    
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointswithin)
  
  ' Find the bulge of the third segment
    Dim currentBulge As Double
    currentBulge = plineObj.GetBulge(2)
   g = Sqr((P * P) + 144)
   angle = Atn((12 / g) / Sqr((-12 / g) * (12 / g) + 1))
   radius = (g / 2) / Sin(12 / g)
   h = radius * (1 - Cos(angle))
   k = h / g
    ' Change the bulge of the third segment
    plineObj.SetBulge 1, k * 2
    plineObj.Update
    plineObj.SetBulge 2, -k * 2
    plineObj.Update
    plineObj.SetBulge 3, -k * 2
    plineObj.Update
    plineObj.SetBulge 4, k * 2
    plineObj.Update
    plineObj.Layer = "K-grav"
    plineObj.Update
  plineObj.Closed = True
  
   ' Offset the polyline
Dim offsetObj As Variant
                                 'plineObj.color = acByLayer
                                 'Dim layerObj As AcadLayer
                                 'Set layerObj = ThisDrawing.Layers.Add("Ball-6.1")
                                 'layerObj.color = acRed

plineObj.Layer = "C-Mill"
plineObj.Update
  offsetObj = plineObj.Offset(20.5)
plineObj.Layer = "K-grav"
plineObj.Update
 
 End If
 End If
                                 'plineObj.color = acByLayer
                                 'Dim layerObj As AcadLayer
                                 'Set layerObj = ThisDrawing.Layers.Add("Ball-6.1")
                                 'layerObj.color = acRed
   
I = 100
  
 Dim plineObj2 As AcadLWPolyline
  Dim points2(0 To 7) As Double
  
  points2(0) = points(0):    points2(1) = a + I:
  points2(2) = points(0):    points2(3) = (a * 2) + I:
  points2(4) = points(6):    points2(5) = (a * 2) + I:
  points2(6) = points(6):    points2(7) = a + I:
  
 For d = 2 To Cells(c, 4)
If d > 1 Then

Set plineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(points2)
  plineObj2.Closed = True

  pointswithin2(0) = points2(0) + 45:             pointswithin2(1) = points2(1) + 45
  pointswithin2(2) = points2(2) + 45:             pointswithin2(3) = points2(3) - 69
  pointswithin2(4) = points2(2) + 45 + P:         pointswithin2(5) = points2(3) - 57
  pointswithin2(6) = points2(2) + 45 + (2 * P):   pointswithin2(7) = points2(3) - 45
  pointswithin2(8) = points2(4) - 45 - P:         pointswithin2(9) = points2(3) - 57
  pointswithin2(10) = points2(4) - 45:            pointswithin2(11) = points2(3) - 69
  pointswithin2(12) = points2(4) - 45:            pointswithin2(13) = points2(7) + 45
  
If a > 180 Then
If b > 180 Then
    plineObj.Layer = "K-grav"
    plineObj.Update
  ' Find the bulge of the third segment
  Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointswithin2)
    currentBulge = plineObj.GetBulge(2)
    ' Change the bulge of the third segment
    plineObj.SetBulge 1, k * 2
    plineObj.Update
    plineObj.SetBulge 2, -k * 2
    plineObj.Update
    plineObj.SetBulge 3, -k * 2
    plineObj.Update
    plineObj.SetBulge 4, k * 2
    plineObj.Update
    plineObj.Layer = "K-grav"
    plineObj.Update
  plineObj.Closed = True
 
plineObj.Layer = "C-Mill"
plineObj.Update
  offsetObj = plineObj.Offset(20.5)
plineObj.Layer = "K-grav"
plineObj.Update
 End If
 End If
  
  points2(1) = points2(3) + I:
  points2(3) = points2(3) + a + I:
  points2(5) = points2(3):
  points2(7) = points2(1):
      
Else: d = 1
d = d + 1

d = Cells(c, 4)
End If

Next d
c = c + 1

a = Cells(c, 2)
b = Cells(c, 3)

I = 100

Loop
              
ZoomAll
AP.Quit
End Sub

Sub MF137()
' Creating search objects
Dim AP As Excel.Application
Dim WB As Excel.Workbook
Dim WS As Excel.Worksheet

Set AP = Excel.Application
Set WB = AP.Workbooks.Open("C:\Users\������\Desktop\������\2.xlsx")
Set WS = WB.Worksheets("������ �˨������ ������")
c = 19
a = Cells(c, 2)
b = Cells(c, 3)
d = Cells(c, 4)
E = 1
I = 0
On Error Resume Next

' Description of all the points that will make up rectangles
 Dim plineObj As AcadLWPolyline
  Dim points(0 To 7) As Double
  Dim pointsktr(0 To 71) As Double
  Dim pointsktr2(0 To 39) As Double
  Dim pointsstr1(0 To 7) As Double
  Dim pointsstr2(0 To 11) As Double
  Dim pointscntr1(0 To 55) As Double
  Dim pointscntr2(0 To 5) As Double
  Dim pointscntr3(0 To 7) As Double
  Dim pointscntr4(0 To 9) As Double
points(6) = 0

Do While a > 0

' Calculate the coordinates of each point
  points(0) = points(6) + I:    points(1) = 0
  points(2) = points(6) + I:    points(3) = a
  points(4) = points(0) + b:    points(5) = a
  points(6) = points(0) + b:    points(7) = 0

' Constructing four segments and connecting the corresponding points

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
  plineObj.Closed = True

If a > 194 Then
If b > 194 Then
  ' Offset the polyline
  Dim offsetObj As Variant
                                 'plineObj.color = acByLayer
                                 'Dim layerObj As AcadLayer
                                 'Set layerObj = ThisDrawing.Layers.Add("Ball-6.1")
                                 'layerObj.color = acRed
plineObj.Layer = "K-grav"
plineObj.Update
  offsetObj = plineObj.Offset(49)
plineObj.Layer = "C-Mill"
plineObj.Update
  offsetObj = plineObj.Offset(69)
plineObj.Layer = "0"
plineObj.Update

  pointsktr(0) = points(0) + 30:      pointsktr(1) = points(1) + 30
  pointsktr(2) = points(0) + 30:      pointsktr(3) = points(1) + 44
  pointsktr(4) = points(0) + 35.6:    pointsktr(5) = points(1) + 44
  pointsktr(6) = points(0) + 35.6:    pointsktr(7) = points(1) + 46.8
  pointsktr(8) = points(0) + 30:      pointsktr(9) = points(1) + 46.8
  pointsktr(10) = points(0) + 30:     pointsktr(11) = points(3) - 46.8
  pointsktr(12) = points(0) + 35.6:   pointsktr(13) = points(3) - 46.8
  pointsktr(14) = points(0) + 35.6:   pointsktr(15) = points(3) - 44
  pointsktr(16) = points(0) + 30:      pointsktr(17) = points(3) - 44
  pointsktr(18) = points(0) + 30:      pointsktr(19) = points(3) - 30
  pointsktr(20) = points(0) + 44:      pointsktr(21) = points(3) - 30
  pointsktr(22) = points(0) + 44:      pointsktr(23) = points(3) - 35.6
  pointsktr(24) = points(0) + 46.8:    pointsktr(25) = points(3) - 35.6
  pointsktr(26) = points(0) + 46.8:    pointsktr(27) = points(3) - 30
  pointsktr(28) = points(4) - 46.8:    pointsktr(29) = points(3) - 30
  pointsktr(30) = points(4) - 46.8:    pointsktr(31) = points(3) - 35.6
  pointsktr(32) = points(4) - 44:      pointsktr(33) = points(3) - 35.6
  pointsktr(34) = points(4) - 44:      pointsktr(35) = points(3) - 30
  pointsktr(36) = points(4) - 30:      pointsktr(37) = points(3) - 30
  pointsktr(38) = points(4) - 30:      pointsktr(39) = points(3) - 44
  pointsktr(40) = points(4) - 35.6:    pointsktr(41) = points(3) - 44
  pointsktr(42) = points(4) - 35.6:    pointsktr(43) = points(3) - 46.8
  pointsktr(44) = points(4) - 30:      pointsktr(45) = points(3) - 46.8
  pointsktr(46) = points(4) - 30:      pointsktr(47) = points(1) + 46.8
  pointsktr(48) = points(4) - 35.6:    pointsktr(49) = points(1) + 46.8
  pointsktr(50) = points(4) - 35.6:    pointsktr(51) = points(1) + 44
  pointsktr(52) = points(4) - 30:      pointsktr(53) = points(1) + 44
  pointsktr(54) = points(4) - 30:      pointsktr(55) = points(1) + 30
  pointsktr(56) = points(4) - 44:      pointsktr(57) = points(1) + 30
  pointsktr(58) = points(4) - 44:      pointsktr(59) = points(1) + 35.6
  pointsktr(60) = points(4) - 46.8:    pointsktr(61) = points(1) + 35.6
  pointsktr(62) = points(4) - 46.8:    pointsktr(63) = points(1) + 30
  pointsktr(64) = points(0) + 46.8:    pointsktr(65) = points(1) + 30
  pointsktr(66) = points(0) + 46.8:    pointsktr(67) = points(1) + 35.6
  pointsktr(68) = points(0) + 44:      pointsktr(69) = points(1) + 35.6
  pointsktr(70) = points(0) + 44:      pointsktr(71) = points(1) + 30
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsktr)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update

  pointsktr2(0) = points(0) + 38.4:    pointsktr2(1) = points(1) + 49.6
  pointsktr2(2) = points(0) + 38.4:    pointsktr2(3) = points(3) - 49.6
  pointsktr2(4) = points(0) + 44:      pointsktr2(5) = points(3) - 49.6
  pointsktr2(6) = points(0) + 44:      pointsktr2(7) = points(3) - 44
  pointsktr2(8) = points(0) + 49.6:    pointsktr2(9) = points(3) - 44
  pointsktr2(10) = points(0) + 49.6:   pointsktr2(11) = points(3) - 38.4
  pointsktr2(12) = points(4) - 49.6:   pointsktr2(13) = points(3) - 38.4
  pointsktr2(14) = points(4) - 49.6:   pointsktr2(15) = points(3) - 44
  pointsktr2(16) = points(4) - 44:     pointsktr2(17) = points(3) - 44
  pointsktr2(18) = points(4) - 44:     pointsktr2(19) = points(3) - 49.6
  pointsktr2(20) = points(4) - 38.4:   pointsktr2(21) = points(3) - 49.6
  pointsktr2(22) = points(4) - 38.4:   pointsktr2(23) = points(1) + 49.6
  pointsktr2(24) = points(4) - 44:     pointsktr2(25) = points(1) + 49.6
  pointsktr2(26) = points(4) - 44:     pointsktr2(27) = points(1) + 44
  pointsktr2(28) = points(4) - 49.6:   pointsktr2(29) = points(1) + 44
  pointsktr2(30) = points(4) - 49.6:   pointsktr2(31) = points(1) + 38.4
  pointsktr2(32) = points(0) + 49.6:   pointsktr2(33) = points(1) + 38.4
  pointsktr2(34) = points(0) + 49.6:   pointsktr2(35) = points(1) + 44
  pointsktr2(36) = points(0) + 44:     pointsktr2(37) = points(1) + 44
  pointsktr2(38) = points(0) + 44:     pointsktr2(39) = points(1) + 49.6
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsktr2)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
'=========================Stripes======================================
  pointsstr1(0) = points(0) + 32.8:   pointsstr1(1) = points(1) + 49.6
  pointsstr1(2) = points(0) + 32.8:   pointsstr1(3) = points(3) - 49.6
  pointsstr1(4) = points(0) + 35.6:     pointsstr1(5) = points(3) - 49.6
  pointsstr1(6) = points(0) + 35.6:     pointsstr1(7) = points(1) + 49.6
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  pointsstr1(0) = points(0) + 49.6:   pointsstr1(1) = points(3) - 35.6
  pointsstr1(2) = points(0) + 49.6:   pointsstr1(3) = points(3) - 32.8
  pointsstr1(4) = points(4) - 49.6:   pointsstr1(5) = points(3) - 32.8
  pointsstr1(6) = points(4) - 49.6:   pointsstr1(7) = points(3) - 35.6
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  pointsstr1(0) = points(4) - 32.8:   pointsstr1(1) = points(1) + 49.6
  pointsstr1(2) = points(4) - 32.8:   pointsstr1(3) = points(3) - 49.6
  pointsstr1(4) = points(4) - 35.6:     pointsstr1(5) = points(3) - 49.6
  pointsstr1(6) = points(4) - 35.6:     pointsstr1(7) = points(1) + 49.6
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  pointsstr1(0) = points(0) + 49.6:   pointsstr1(1) = points(1) + 35.6
  pointsstr1(2) = points(0) + 49.6:   pointsstr1(3) = points(1) + 32.8
  pointsstr1(4) = points(4) - 49.6:   pointsstr1(5) = points(1) + 32.8
  pointsstr1(6) = points(4) - 49.6:   pointsstr1(7) = points(1) + 35.6
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
'=========================Squares======================================
  pointsstr1(0) = points(0) + 38.4:   pointsstr1(1) = points(1) + 38.4
  pointsstr1(2) = points(0) + 38.4:   pointsstr1(3) = points(1) + 41.2
  pointsstr1(4) = points(0) + 41.2:   pointsstr1(5) = points(1) + 41.2
  pointsstr1(6) = points(0) + 41.2:   pointsstr1(7) = points(1) + 38.4
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  pointsstr1(0) = points(0) + 38.4:   pointsstr1(1) = points(1) + 44
  pointsstr1(2) = points(0) + 38.4:   pointsstr1(3) = points(1) + 46.8
  pointsstr1(4) = points(0) + 41.2:   pointsstr1(5) = points(1) + 46.8
  pointsstr1(6) = points(0) + 41.2:   pointsstr1(7) = points(1) + 44
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  pointsstr1(0) = points(0) + 44:     pointsstr1(1) = points(1) + 38.4
  pointsstr1(2) = points(0) + 46.8:   pointsstr1(3) = points(1) + 38.4
  pointsstr1(4) = points(0) + 46.8:   pointsstr1(5) = points(1) + 41.2
  pointsstr1(6) = points(0) + 44:     pointsstr1(7) = points(1) + 41.2
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  '***********************
  pointsstr1(0) = points(4) - 38.4:   pointsstr1(1) = points(1) + 38.4
  pointsstr1(2) = points(4) - 38.4:   pointsstr1(3) = points(1) + 41.2
  pointsstr1(4) = points(4) - 41.2:   pointsstr1(5) = points(1) + 41.2
  pointsstr1(6) = points(4) - 41.2:   pointsstr1(7) = points(1) + 38.4
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  pointsstr1(0) = points(4) - 38.4:   pointsstr1(1) = points(1) + 44
  pointsstr1(2) = points(4) - 38.4:   pointsstr1(3) = points(1) + 46.8
  pointsstr1(4) = points(4) - 41.2:   pointsstr1(5) = points(1) + 46.8
  pointsstr1(6) = points(4) - 41.2:   pointsstr1(7) = points(1) + 44
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  pointsstr1(0) = points(4) - 44:     pointsstr1(1) = points(1) + 38.4
  pointsstr1(2) = points(4) - 46.8:   pointsstr1(3) = points(1) + 38.4
  pointsstr1(4) = points(4) - 46.8:   pointsstr1(5) = points(1) + 41.2
  pointsstr1(6) = points(4) - 44:     pointsstr1(7) = points(1) + 41.2
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  '***********************
  pointsstr1(0) = points(0) + 38.4:   pointsstr1(1) = points(3) - 38.4
  pointsstr1(2) = points(0) + 38.4:   pointsstr1(3) = points(3) - 41.2
  pointsstr1(4) = points(0) + 41.2:   pointsstr1(5) = points(3) - 41.2
  pointsstr1(6) = points(0) + 41.2:   pointsstr1(7) = points(3) - 38.4
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  pointsstr1(0) = points(0) + 38.4:   pointsstr1(1) = points(3) - 44
  pointsstr1(2) = points(0) + 38.4:   pointsstr1(3) = points(3) - 46.8
  pointsstr1(4) = points(0) + 41.2:   pointsstr1(5) = points(3) - 46.8
  pointsstr1(6) = points(0) + 41.2:   pointsstr1(7) = points(3) - 44
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  pointsstr1(0) = points(0) + 44:     pointsstr1(1) = points(3) - 38.4
  pointsstr1(2) = points(0) + 46.8:   pointsstr1(3) = points(3) - 38.4
  pointsstr1(4) = points(0) + 46.8:   pointsstr1(5) = points(3) - 41.2
  pointsstr1(6) = points(0) + 44:     pointsstr1(7) = points(3) - 41.2
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  '***********************
  pointsstr1(0) = points(4) - 38.4:   pointsstr1(1) = points(3) - 38.4
  pointsstr1(2) = points(4) - 38.4:   pointsstr1(3) = points(3) - 41.2
  pointsstr1(4) = points(4) - 41.2:   pointsstr1(5) = points(3) - 41.2
  pointsstr1(6) = points(4) - 41.2:   pointsstr1(7) = points(3) - 38.4
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  pointsstr1(0) = points(4) - 38.4:   pointsstr1(1) = points(3) - 44
  pointsstr1(2) = points(4) - 38.4:   pointsstr1(3) = points(3) - 46.8
  pointsstr1(4) = points(4) - 41.2:   pointsstr1(5) = points(3) - 46.8
  pointsstr1(6) = points(4) - 41.2:   pointsstr1(7) = points(3) - 44
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  pointsstr1(0) = points(4) - 44:     pointsstr1(1) = points(3) - 38.4
  pointsstr1(2) = points(4) - 46.8:   pointsstr1(3) = points(3) - 38.4
  pointsstr1(4) = points(4) - 46.8:   pointsstr1(5) = points(3) - 41.2
  pointsstr1(6) = points(4) - 44:     pointsstr1(7) = points(3) - 41.2
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
'=========================Angles======================================
  pointsstr2(0) = points(0) + 32.8:   pointsstr2(1) = points(1) + 32.8
  pointsstr2(2) = points(0) + 32.8:   pointsstr2(3) = points(1) + 41.2
  pointsstr2(4) = points(0) + 35.6:   pointsstr2(5) = points(1) + 41.2
  pointsstr2(6) = points(0) + 35.6:   pointsstr2(7) = points(1) + 35.6
  pointsstr2(8) = points(0) + 41.2:   pointsstr2(9) = points(1) + 35.6
  pointsstr2(10) = points(0) + 41.2:  pointsstr2(11) = points(1) + 32.8
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr2)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  pointsstr2(0) = points(0) + 32.8:   pointsstr2(1) = points(3) - 32.8
  pointsstr2(2) = points(0) + 32.8:   pointsstr2(3) = points(3) - 41.2
  pointsstr2(4) = points(0) + 35.6:   pointsstr2(5) = points(3) - 41.2
  pointsstr2(6) = points(0) + 35.6:   pointsstr2(7) = points(3) - 35.6
  pointsstr2(8) = points(0) + 41.2:   pointsstr2(9) = points(3) - 35.6
  pointsstr2(10) = points(0) + 41.2:  pointsstr2(11) = points(3) - 32.8
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr2)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  pointsstr2(0) = points(4) - 32.8:   pointsstr2(1) = points(3) - 32.8
  pointsstr2(2) = points(4) - 32.8:   pointsstr2(3) = points(3) - 41.2
  pointsstr2(4) = points(4) - 35.6:   pointsstr2(5) = points(3) - 41.2
  pointsstr2(6) = points(4) - 35.6:   pointsstr2(7) = points(3) - 35.6
  pointsstr2(8) = points(4) - 41.2:   pointsstr2(9) = points(3) - 35.6
  pointsstr2(10) = points(4) - 41.2:  pointsstr2(11) = points(3) - 32.8
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr2)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  pointsstr2(0) = points(4) - 32.8:   pointsstr2(1) = points(1) + 32.8
  pointsstr2(2) = points(4) - 32.8:   pointsstr2(3) = points(1) + 41.2
  pointsstr2(4) = points(4) - 35.6:   pointsstr2(5) = points(1) + 41.2
  pointsstr2(6) = points(4) - 35.6:   pointsstr2(7) = points(1) + 35.6
  pointsstr2(8) = points(4) - 41.2:   pointsstr2(9) = points(1) + 35.6
  pointsstr2(10) = points(4) - 41.2:  pointsstr2(11) = points(1) + 32.8
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr2)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  
  If a > 540 Then
  If b > 250 Then
'=========================CENTER======================================
  bcp = points(0) + (b / 2)
  acp = points(1) + (a / 2)
  pointscntr1(0) = bcp + 27.1207:    pointscntr1(1) = acp - 5.68434E-14
  pointscntr1(2) = bcp + 12.4984:    pointscntr1(3) = acp - 12.3831
  pointscntr1(4) = bcp + 17.5444:    pointscntr1(5) = acp - 19.634
  pointscntr1(6) = bcp + 27.5594:    pointscntr1(7) = acp - 42.4202
  pointscntr1(8) = bcp + 19.1501:    pointscntr1(9) = acp - 75.0415
  pointscntr1(10) = bcp + 3.99478:   pointscntr1(11) = acp - 103.894
  pointscntr1(12) = bcp + 7.9846:    pointscntr1(13) = acp - 104.96
  pointscntr1(14) = bcp + 0:         pointscntr1(15) = acp - 180.75
  pointscntr1(16) = bcp - 7.9846:    pointscntr1(17) = acp - 104.96
  pointscntr1(18) = bcp - 3.99478:   pointscntr1(19) = acp - 103.894
  pointscntr1(20) = bcp - 19.1501:   pointscntr1(21) = acp - 75.0415
  pointscntr1(22) = bcp - 27.5594:   pointscntr1(23) = acp - 42.4202
  pointscntr1(24) = bcp - 17.5444:   pointscntr1(25) = acp - 19.634
  pointscntr1(26) = bcp - 12.4984:   pointscntr1(27) = acp - 12.3831
  pointscntr1(28) = bcp - 27.1207:   pointscntr1(29) = acp + 5.68434E-14
  pointscntr1(30) = bcp - 12.4984:   pointscntr1(31) = acp + 12.3831
  pointscntr1(32) = bcp - 17.5444:   pointscntr1(33) = acp + 19.634
  pointscntr1(34) = bcp - 27.5594:   pointscntr1(35) = acp + 42.4202
  pointscntr1(36) = bcp - 19.1501:   pointscntr1(37) = acp + 75.0415
  pointscntr1(38) = bcp - 3.99478:   pointscntr1(39) = acp + 103.894
  pointscntr1(40) = bcp - 7.9846:    pointscntr1(41) = acp + 104.96
  pointscntr1(42) = bcp + 0:         pointscntr1(43) = acp + 180.75
  pointscntr1(44) = bcp + 7.9846:    pointscntr1(45) = acp + 104.96
  pointscntr1(46) = bcp + 3.99478:   pointscntr1(47) = acp + 103.894
  pointscntr1(48) = bcp + 19.1501:   pointscntr1(49) = acp + 75.0415
  pointscntr1(50) = bcp + 27.5594:   pointscntr1(51) = acp + 42.4202
  pointscntr1(52) = bcp + 17.5444:   pointscntr1(53) = acp + 19.634
  pointscntr1(54) = bcp + 12.4984:   pointscntr1(55) = acp + 12.3831
  
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  
  ' Find the bulge of the third segment
    currentBulge = plineObj.GetBulge(2)
    ' Change the bulge of the third segment
    plineObj.SetBulge 0, 0.0324358
    plineObj.Update
    plineObj.SetBulge 1, -0.00917899
    plineObj.Update
    plineObj.SetBulge 2, -0.0976773
    plineObj.Update
    plineObj.SetBulge 3, -0.221543
    plineObj.Update
    plineObj.SetBulge 4, -0.247182
    plineObj.Update
    plineObj.SetBulge 5, 0.035764
    plineObj.Update
    plineObj.SetBulge 8, 0.035764
    plineObj.Update
    plineObj.SetBulge 9, -0.247182
    plineObj.Update
    plineObj.SetBulge 10, -0.221543
    plineObj.Update
    plineObj.SetBulge 11, -0.0976773
    plineObj.Update
    plineObj.SetBulge 12, -0.00917899
    plineObj.Update
    plineObj.SetBulge 13, 0.0324358
    plineObj.Update
    plineObj.SetBulge 14, 0.0324358
    plineObj.Update
    plineObj.SetBulge 15, -0.00917899
    plineObj.Update
    plineObj.SetBulge 16, -0.0976773
    plineObj.Update
    plineObj.SetBulge 17, -0.221543
    plineObj.Update
    plineObj.SetBulge 18, -0.247182
    plineObj.Update
    plineObj.SetBulge 19, 0.035764
    plineObj.Update
    plineObj.SetBulge 22, 0.035764
    plineObj.Update
    plineObj.SetBulge 23, -0.247182
    plineObj.Update
    plineObj.SetBulge 24, -0.221543
    plineObj.Update
    plineObj.SetBulge 25, -0.0976773
    plineObj.Update
    plineObj.SetBulge 26, -0.00917899
    plineObj.Update
    plineObj.SetBulge 27, 0.0324358
    plineObj.Update
    
    plineObj.Layer = "K-grav"
    plineObj.Update
    plineObj.Closed = True
  
  pointscntr2(0) = bcp + 0:         pointscntr2(1) = acp + 102.168
  pointscntr2(2) = bcp - 15.2911:   pointscntr2(3) = acp + 78.5624
  pointscntr2(4) = bcp + 15.2911:   pointscntr2(5) = acp + 78.5624
  
 Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr2)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  
    plineObj.SetBulge 0, 0.250892
    plineObj.Update
    plineObj.SetBulge 1, -0.332226
    plineObj.Update
    plineObj.SetBulge 2, 0.250892
    plineObj.Update
  Dim b1(0 To 2) As Double
  Dim b2(0 To 2) As Double
  b1(0) = (b / 2) - 1:  b1(1) = (a / 2)
  b2(0) = (b / 2) + 1:  b2(1) = (a / 2)
  RetVal = plineObj.Mirror(b1, b2)

  pointscntr2(0) = bcp + 15.7341:  pointscntr2(1) = acp + 73.5418
  pointscntr2(2) = bcp - 15.7341:  pointscntr2(3) = acp + 73.5418
  pointscntr2(4) = bcp + 0:        pointscntr2(5) = acp + 26.2803
  
 Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr2)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  
    plineObj.SetBulge 0, 0.419513
    plineObj.Update
    plineObj.SetBulge 1, 0.182166
    plineObj.Update
    plineObj.SetBulge 2, 0.182166
    plineObj.Update
    RetVal = plineObj.Mirror(b1, b2)
  
  pointscntr3(0) = bcp + 0:             pointscntr3(1) = acp + 147.04
  pointscntr3(2) = bcp - 4.1528:        pointscntr3(3) = acp + 107.621
  pointscntr3(4) = bcp + 1.13687E-12:   pointscntr3(5) = acp + 106.064
  pointscntr3(6) = bcp + 4.1528:        pointscntr3(7) = acp + 107.621
  
 Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr3)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  
    plineObj.SetBulge 1, -0.0342582
    plineObj.Update
    plineObj.SetBulge 2, -0.0342582
    plineObj.Update
    RetVal = plineObj.Mirror(b1, b2)
    
  pointscntr3(0) = bcp + 0:             pointscntr3(1) = acp + 20.5705
  pointscntr3(2) = bcp + 7.88663:       pointscntr3(3) = acp + 12.0123
  pointscntr3(4) = bcp + 0:             pointscntr3(5) = acp + 1.97396
  pointscntr3(6) = bcp - 7.88663:       pointscntr3(7) = acp + 12.0123
  
 Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr3)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  
    plineObj.SetBulge 0, 0.0192317
    plineObj.Update
    plineObj.SetBulge 1, -0.0134617
    plineObj.Update
    plineObj.SetBulge 2, -0.0134617
    plineObj.Update
    plineObj.SetBulge 3, 0.0192317
    plineObj.Update
    RetVal = plineObj.Mirror(b1, b2)
  
  pointscntr3(0) = bcp - 21.3616:      pointscntr3(1) = acp + 0
  pointscntr3(2) = bcp - 10.4078:      pointscntr3(3) = acp - 9.53775
  pointscntr3(4) = bcp - 2.90986:      pointscntr3(5) = acp + 0
  pointscntr3(6) = bcp - 10.4078:      pointscntr3(7) = acp + 9.53775
  
 Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr3)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  
    plineObj.SetBulge 0, -0.0240064
    plineObj.Update
    plineObj.SetBulge 1, -0.012607
    plineObj.Update
    plineObj.SetBulge 2, -0.012607
    plineObj.Update
    plineObj.SetBulge 3, -0.0240064
    plineObj.Update
  Dim a1(0 To 2) As Double
  Dim a2(0 To 2) As Double
  a1(0) = points(4) - (b / 2):  a1(1) = (a / 2) + 1
  a2(0) = points(4) - (b / 2):  a2(1) = (a / 2) - 1
  RetVal = plineObj.Mirror(a1, a2)
  
  pointscntr4(0) = bcp - 24.1296:     pointscntr4(1) = acp + 43.1205
  pointscntr4(2) = bcp - 19.2635:     pointscntr4(3) = acp + 68.8882
  pointscntr4(4) = bcp - 2.17621:     pointscntr4(5) = acp + 23.4154
  pointscntr4(6) = bcp - 9.97906:     pointscntr4(7) = acp + 14.8729
  pointscntr4(8) = bcp - 14.6543:     pointscntr4(9) = acp + 21.6084
  
 Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr4)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  
    plineObj.SetBulge 0, -0.186466
    plineObj.Update
    plineObj.SetBulge 1, 0.168477
    plineObj.Update
    plineObj.SetBulge 2, -0.0195719
    plineObj.Update
    plineObj.SetBulge 3, -0.00864503
    plineObj.Update
    plineObj.SetBulge 4, -0.0975544
    plineObj.Update
  RetVal = plineObj.Copy
  RetVal = plineObj.Mirror(a1, a2)
  RetVal = plineObj.Mirror(b1, b2)
  
 ' Define the rotation of 45 degrees about a
  ' base point of (4, 4.25, 0)
  Dim basePoint(0 To 2) As Double
  Dim rotationAngle As Double
  basePoint(0) = points(4) - (b / 2): basePoint(1) = points(3) - (a / 2)
  rotationAngle = 3.14159   ' 45 degrees

  ' Rotate the polyline
  plineObj.Rotate basePoint, rotationAngle
  plineObj.Update
End If
End If
End If
End If
  
  
  
I = 100
  
 Dim plineObj2 As AcadLWPolyline
  Dim points2(0 To 7) As Double
  
  points2(0) = points(0):    points2(1) = a + I:
  points2(2) = points(0):    points2(3) = (a * 2) + I:
  points2(4) = points(6):    points2(5) = (a * 2) + I:
  points2(6) = points(6):    points2(7) = a + I:
  
 For d = 2 To Cells(c, 4)
If d > 1 Then

Set plineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(points2)
  plineObj2.Closed = True
  
If a > 194 Then
If b > 194 Then
  ' Offset the polyline
                                 'plineObj.color = acByLayer
                                 'Dim layerObj As AcadLayer
                                 'Set layerObj = ThisDrawing.Layers.Add("Ball-6.1")
                                 'layerObj.color = acRed
plineObj2.Layer = "K-grav"
plineObj2.Update
  offsetObj = plineObj2.Offset(49)
plineObj2.Layer = "C-Mill"
plineObj2.Update
  offsetObj = plineObj2.Offset(69)
plineObj2.Layer = "0"
plineObj2.Update

  pointsktr(0) = points2(0) + 30:      pointsktr(1) = points2(1) + 30
  pointsktr(2) = points2(0) + 30:      pointsktr(3) = points2(1) + 44
  pointsktr(4) = points2(0) + 35.6:    pointsktr(5) = points2(1) + 44
  pointsktr(6) = points2(0) + 35.6:    pointsktr(7) = points2(1) + 46.8
  pointsktr(8) = points2(0) + 30:      pointsktr(9) = points2(1) + 46.8
  pointsktr(10) = points2(0) + 30:     pointsktr(11) = points2(3) - 46.8
  pointsktr(12) = points2(0) + 35.6:   pointsktr(13) = points2(3) - 46.8
  pointsktr(14) = points2(0) + 35.6:   pointsktr(15) = points2(3) - 44
  pointsktr(16) = points2(0) + 30:      pointsktr(17) = points2(3) - 44
  pointsktr(18) = points2(0) + 30:      pointsktr(19) = points2(3) - 30
  pointsktr(20) = points2(0) + 44:      pointsktr(21) = points2(3) - 30
  pointsktr(22) = points2(0) + 44:      pointsktr(23) = points2(3) - 35.6
  pointsktr(24) = points2(0) + 46.8:    pointsktr(25) = points2(3) - 35.6
  pointsktr(26) = points2(0) + 46.8:    pointsktr(27) = points2(3) - 30
  pointsktr(28) = points2(4) - 46.8:    pointsktr(29) = points2(3) - 30
  pointsktr(30) = points2(4) - 46.8:    pointsktr(31) = points2(3) - 35.6
  pointsktr(32) = points2(4) - 44:      pointsktr(33) = points2(3) - 35.6
  pointsktr(34) = points2(4) - 44:      pointsktr(35) = points2(3) - 30
  pointsktr(36) = points2(4) - 30:      pointsktr(37) = points2(3) - 30
  pointsktr(38) = points2(4) - 30:      pointsktr(39) = points2(3) - 44
  pointsktr(40) = points2(4) - 35.6:    pointsktr(41) = points2(3) - 44
  pointsktr(42) = points2(4) - 35.6:    pointsktr(43) = points2(3) - 46.8
  pointsktr(44) = points2(4) - 30:      pointsktr(45) = points2(3) - 46.8
  pointsktr(46) = points2(4) - 30:      pointsktr(47) = points2(1) + 46.8
  pointsktr(48) = points2(4) - 35.6:    pointsktr(49) = points2(1) + 46.8
  pointsktr(50) = points2(4) - 35.6:    pointsktr(51) = points2(1) + 44
  pointsktr(52) = points2(4) - 30:      pointsktr(53) = points2(1) + 44
  pointsktr(54) = points2(4) - 30:      pointsktr(55) = points2(1) + 30
  pointsktr(56) = points2(4) - 44:      pointsktr(57) = points2(1) + 30
  pointsktr(58) = points2(4) - 44:      pointsktr(59) = points2(1) + 35.6
  pointsktr(60) = points2(4) - 46.8:    pointsktr(61) = points2(1) + 35.6
  pointsktr(62) = points2(4) - 46.8:    pointsktr(63) = points2(1) + 30
  pointsktr(64) = points2(0) + 46.8:    pointsktr(65) = points2(1) + 30
  pointsktr(66) = points2(0) + 46.8:    pointsktr(67) = points2(1) + 35.6
  pointsktr(68) = points2(0) + 44:      pointsktr(69) = points2(1) + 35.6
  pointsktr(70) = points2(0) + 44:      pointsktr(71) = points2(1) + 30
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsktr)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update

  pointsktr2(0) = points2(0) + 38.4:    pointsktr2(1) = points2(1) + 49.6
  pointsktr2(2) = points2(0) + 38.4:    pointsktr2(3) = points2(3) - 49.6
  pointsktr2(4) = points2(0) + 44:      pointsktr2(5) = points2(3) - 49.6
  pointsktr2(6) = points2(0) + 44:      pointsktr2(7) = points2(3) - 44
  pointsktr2(8) = points2(0) + 49.6:    pointsktr2(9) = points2(3) - 44
  pointsktr2(10) = points2(0) + 49.6:   pointsktr2(11) = points2(3) - 38.4
  pointsktr2(12) = points2(4) - 49.6:   pointsktr2(13) = points2(3) - 38.4
  pointsktr2(14) = points2(4) - 49.6:   pointsktr2(15) = points2(3) - 44
  pointsktr2(16) = points2(4) - 44:     pointsktr2(17) = points2(3) - 44
  pointsktr2(18) = points2(4) - 44:     pointsktr2(19) = points2(3) - 49.6
  pointsktr2(20) = points2(4) - 38.4:   pointsktr2(21) = points2(3) - 49.6
  pointsktr2(22) = points2(4) - 38.4:   pointsktr2(23) = points2(1) + 49.6
  pointsktr2(24) = points2(4) - 44:     pointsktr2(25) = points2(1) + 49.6
  pointsktr2(26) = points2(4) - 44:     pointsktr2(27) = points2(1) + 44
  pointsktr2(28) = points2(4) - 49.6:   pointsktr2(29) = points2(1) + 44
  pointsktr2(30) = points2(4) - 49.6:   pointsktr2(31) = points2(1) + 38.4
  pointsktr2(32) = points2(0) + 49.6:   pointsktr2(33) = points2(1) + 38.4
  pointsktr2(34) = points2(0) + 49.6:   pointsktr2(35) = points2(1) + 44
  pointsktr2(36) = points2(0) + 44:     pointsktr2(37) = points2(1) + 44
  pointsktr2(38) = points2(0) + 44:     pointsktr2(39) = points2(1) + 49.6
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsktr2)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
'=========================Stripes======================================
  pointsstr1(0) = points2(0) + 32.8:   pointsstr1(1) = points2(1) + 49.6
  pointsstr1(2) = points2(0) + 32.8:   pointsstr1(3) = points2(3) - 49.6
  pointsstr1(4) = points2(0) + 35.6:     pointsstr1(5) = points2(3) - 49.6
  pointsstr1(6) = points2(0) + 35.6:     pointsstr1(7) = points2(1) + 49.6
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  pointsstr1(0) = points2(0) + 49.6:   pointsstr1(1) = points2(3) - 35.6
  pointsstr1(2) = points2(0) + 49.6:   pointsstr1(3) = points2(3) - 32.8
  pointsstr1(4) = points2(4) - 49.6:   pointsstr1(5) = points2(3) - 32.8
  pointsstr1(6) = points2(4) - 49.6:   pointsstr1(7) = points2(3) - 35.6
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  pointsstr1(0) = points2(4) - 32.8:   pointsstr1(1) = points2(1) + 49.6
  pointsstr1(2) = points2(4) - 32.8:   pointsstr1(3) = points2(3) - 49.6
  pointsstr1(4) = points2(4) - 35.6:     pointsstr1(5) = points2(3) - 49.6
  pointsstr1(6) = points2(4) - 35.6:     pointsstr1(7) = points2(1) + 49.6
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  pointsstr1(0) = points2(0) + 49.6:   pointsstr1(1) = points2(1) + 35.6
  pointsstr1(2) = points2(0) + 49.6:   pointsstr1(3) = points2(1) + 32.8
  pointsstr1(4) = points2(4) - 49.6:   pointsstr1(5) = points2(1) + 32.8
  pointsstr1(6) = points2(4) - 49.6:   pointsstr1(7) = points2(1) + 35.6
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
'=========================Squares======================================
  pointsstr1(0) = points2(0) + 38.4:   pointsstr1(1) = points2(1) + 38.4
  pointsstr1(2) = points2(0) + 38.4:   pointsstr1(3) = points2(1) + 41.2
  pointsstr1(4) = points2(0) + 41.2:   pointsstr1(5) = points2(1) + 41.2
  pointsstr1(6) = points2(0) + 41.2:   pointsstr1(7) = points2(1) + 38.4
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  pointsstr1(0) = points2(0) + 38.4:   pointsstr1(1) = points2(1) + 44
  pointsstr1(2) = points2(0) + 38.4:   pointsstr1(3) = points2(1) + 46.8
  pointsstr1(4) = points2(0) + 41.2:   pointsstr1(5) = points2(1) + 46.8
  pointsstr1(6) = points2(0) + 41.2:   pointsstr1(7) = points2(1) + 44
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  pointsstr1(0) = points2(0) + 44:     pointsstr1(1) = points2(1) + 38.4
  pointsstr1(2) = points2(0) + 46.8:   pointsstr1(3) = points2(1) + 38.4
  pointsstr1(4) = points2(0) + 46.8:   pointsstr1(5) = points2(1) + 41.2
  pointsstr1(6) = points2(0) + 44:     pointsstr1(7) = points2(1) + 41.2
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  '***********************
  pointsstr1(0) = points2(4) - 38.4:   pointsstr1(1) = points2(1) + 38.4
  pointsstr1(2) = points2(4) - 38.4:   pointsstr1(3) = points2(1) + 41.2
  pointsstr1(4) = points2(4) - 41.2:   pointsstr1(5) = points2(1) + 41.2
  pointsstr1(6) = points2(4) - 41.2:   pointsstr1(7) = points2(1) + 38.4
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  pointsstr1(0) = points2(4) - 38.4:   pointsstr1(1) = points2(1) + 44
  pointsstr1(2) = points2(4) - 38.4:   pointsstr1(3) = points2(1) + 46.8
  pointsstr1(4) = points2(4) - 41.2:   pointsstr1(5) = points2(1) + 46.8
  pointsstr1(6) = points2(4) - 41.2:   pointsstr1(7) = points2(1) + 44
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  pointsstr1(0) = points2(4) - 44:     pointsstr1(1) = points2(1) + 38.4
  pointsstr1(2) = points2(4) - 46.8:   pointsstr1(3) = points2(1) + 38.4
  pointsstr1(4) = points2(4) - 46.8:   pointsstr1(5) = points2(1) + 41.2
  pointsstr1(6) = points2(4) - 44:     pointsstr1(7) = points2(1) + 41.2
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  '***********************
  pointsstr1(0) = points2(0) + 38.4:   pointsstr1(1) = points2(3) - 38.4
  pointsstr1(2) = points2(0) + 38.4:   pointsstr1(3) = points2(3) - 41.2
  pointsstr1(4) = points2(0) + 41.2:   pointsstr1(5) = points2(3) - 41.2
  pointsstr1(6) = points2(0) + 41.2:   pointsstr1(7) = points2(3) - 38.4
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  pointsstr1(0) = points2(0) + 38.4:   pointsstr1(1) = points2(3) - 44
  pointsstr1(2) = points2(0) + 38.4:   pointsstr1(3) = points2(3) - 46.8
  pointsstr1(4) = points2(0) + 41.2:   pointsstr1(5) = points2(3) - 46.8
  pointsstr1(6) = points2(0) + 41.2:   pointsstr1(7) = points2(3) - 44
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  pointsstr1(0) = points2(0) + 44:     pointsstr1(1) = points2(3) - 38.4
  pointsstr1(2) = points2(0) + 46.8:   pointsstr1(3) = points2(3) - 38.4
  pointsstr1(4) = points2(0) + 46.8:   pointsstr1(5) = points2(3) - 41.2
  pointsstr1(6) = points2(0) + 44:     pointsstr1(7) = points2(3) - 41.2
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  '***********************
  pointsstr1(0) = points2(4) - 38.4:   pointsstr1(1) = points2(3) - 38.4
  pointsstr1(2) = points2(4) - 38.4:   pointsstr1(3) = points2(3) - 41.2
  pointsstr1(4) = points2(4) - 41.2:   pointsstr1(5) = points2(3) - 41.2
  pointsstr1(6) = points2(4) - 41.2:   pointsstr1(7) = points2(3) - 38.4
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  pointsstr1(0) = points2(4) - 38.4:   pointsstr1(1) = points2(3) - 44
  pointsstr1(2) = points2(4) - 38.4:   pointsstr1(3) = points2(3) - 46.8
  pointsstr1(4) = points2(4) - 41.2:   pointsstr1(5) = points2(3) - 46.8
  pointsstr1(6) = points2(4) - 41.2:   pointsstr1(7) = points2(3) - 44
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  pointsstr1(0) = points2(4) - 44:     pointsstr1(1) = points2(3) - 38.4
  pointsstr1(2) = points2(4) - 46.8:   pointsstr1(3) = points2(3) - 38.4
  pointsstr1(4) = points2(4) - 46.8:   pointsstr1(5) = points2(3) - 41.2
  pointsstr1(6) = points2(4) - 44:     pointsstr1(7) = points2(3) - 41.2
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
'=========================Angles======================================
  pointsstr2(0) = points2(0) + 32.8:   pointsstr2(1) = points2(1) + 32.8
  pointsstr2(2) = points2(0) + 32.8:   pointsstr2(3) = points2(1) + 41.2
  pointsstr2(4) = points2(0) + 35.6:   pointsstr2(5) = points2(1) + 41.2
  pointsstr2(6) = points2(0) + 35.6:   pointsstr2(7) = points2(1) + 35.6
  pointsstr2(8) = points2(0) + 41.2:   pointsstr2(9) = points2(1) + 35.6
  pointsstr2(10) = points2(0) + 41.2:  pointsstr2(11) = points2(1) + 32.8
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr2)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  pointsstr2(0) = points2(0) + 32.8:   pointsstr2(1) = points2(3) - 32.8
  pointsstr2(2) = points2(0) + 32.8:   pointsstr2(3) = points2(3) - 41.2
  pointsstr2(4) = points2(0) + 35.6:   pointsstr2(5) = points2(3) - 41.2
  pointsstr2(6) = points2(0) + 35.6:   pointsstr2(7) = points2(3) - 35.6
  pointsstr2(8) = points2(0) + 41.2:   pointsstr2(9) = points2(3) - 35.6
  pointsstr2(10) = points2(0) + 41.2:  pointsstr2(11) = points2(3) - 32.8
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr2)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  pointsstr2(0) = points2(4) - 32.8:   pointsstr2(1) = points2(3) - 32.8
  pointsstr2(2) = points2(4) - 32.8:   pointsstr2(3) = points2(3) - 41.2
  pointsstr2(4) = points2(4) - 35.6:   pointsstr2(5) = points2(3) - 41.2
  pointsstr2(6) = points2(4) - 35.6:   pointsstr2(7) = points2(3) - 35.6
  pointsstr2(8) = points2(4) - 41.2:   pointsstr2(9) = points2(3) - 35.6
  pointsstr2(10) = points2(4) - 41.2:  pointsstr2(11) = points2(3) - 32.8
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr2)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  pointsstr2(0) = points2(4) - 32.8:   pointsstr2(1) = points2(1) + 32.8
  pointsstr2(2) = points2(4) - 32.8:   pointsstr2(3) = points2(1) + 41.2
  pointsstr2(4) = points2(4) - 35.6:   pointsstr2(5) = points2(1) + 41.2
  pointsstr2(6) = points2(4) - 35.6:   pointsstr2(7) = points2(1) + 35.6
  pointsstr2(8) = points2(4) - 41.2:   pointsstr2(9) = points2(1) + 35.6
  pointsstr2(10) = points2(4) - 41.2:  pointsstr2(11) = points2(1) + 32.8
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr2)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  
  If a > 540 Then
  If b > 250 Then
'=========================CENTER======================================
  bcp = points2(0) + (b / 2)
  acp = points2(1) + (a / 2)
  pointscntr1(0) = bcp + 27.1207:    pointscntr1(1) = acp - 5.68434E-14
  pointscntr1(2) = bcp + 12.4984:    pointscntr1(3) = acp - 12.3831
  pointscntr1(4) = bcp + 17.5444:    pointscntr1(5) = acp - 19.634
  pointscntr1(6) = bcp + 27.5594:    pointscntr1(7) = acp - 42.4202
  pointscntr1(8) = bcp + 19.1501:    pointscntr1(9) = acp - 75.0415
  pointscntr1(10) = bcp + 3.99478:   pointscntr1(11) = acp - 103.894
  pointscntr1(12) = bcp + 7.9846:    pointscntr1(13) = acp - 104.96
  pointscntr1(14) = bcp + 0:         pointscntr1(15) = acp - 180.75
  pointscntr1(16) = bcp - 7.9846:    pointscntr1(17) = acp - 104.96
  pointscntr1(18) = bcp - 3.99478:   pointscntr1(19) = acp - 103.894
  pointscntr1(20) = bcp - 19.1501:   pointscntr1(21) = acp - 75.0415
  pointscntr1(22) = bcp - 27.5594:   pointscntr1(23) = acp - 42.4202
  pointscntr1(24) = bcp - 17.5444:   pointscntr1(25) = acp - 19.634
  pointscntr1(26) = bcp - 12.4984:   pointscntr1(27) = acp - 12.3831
  pointscntr1(28) = bcp - 27.1207:   pointscntr1(29) = acp + 5.68434E-14
  pointscntr1(30) = bcp - 12.4984:   pointscntr1(31) = acp + 12.3831
  pointscntr1(32) = bcp - 17.5444:   pointscntr1(33) = acp + 19.634
  pointscntr1(34) = bcp - 27.5594:   pointscntr1(35) = acp + 42.4202
  pointscntr1(36) = bcp - 19.1501:   pointscntr1(37) = acp + 75.0415
  pointscntr1(38) = bcp - 3.99478:   pointscntr1(39) = acp + 103.894
  pointscntr1(40) = bcp - 7.9846:    pointscntr1(41) = acp + 104.96
  pointscntr1(42) = bcp + 0:         pointscntr1(43) = acp + 180.75
  pointscntr1(44) = bcp + 7.9846:    pointscntr1(45) = acp + 104.96
  pointscntr1(46) = bcp + 3.99478:   pointscntr1(47) = acp + 103.894
  pointscntr1(48) = bcp + 19.1501:   pointscntr1(49) = acp + 75.0415
  pointscntr1(50) = bcp + 27.5594:   pointscntr1(51) = acp + 42.4202
  pointscntr1(52) = bcp + 17.5444:   pointscntr1(53) = acp + 19.634
  pointscntr1(54) = bcp + 12.4984:   pointscntr1(55) = acp + 12.3831
  
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  
  ' Find the bulge of the third segment
    currentBulge = plineObj.GetBulge(2)
    ' Change the bulge of the third segment
    plineObj.SetBulge 0, 0.0324358
    plineObj.Update
    plineObj.SetBulge 1, -0.00917899
    plineObj.Update
    plineObj.SetBulge 2, -0.0976773
    plineObj.Update
    plineObj.SetBulge 3, -0.221543
    plineObj.Update
    plineObj.SetBulge 4, -0.247182
    plineObj.Update
    plineObj.SetBulge 5, 0.035764
    plineObj.Update
    plineObj.SetBulge 8, 0.035764
    plineObj.Update
    plineObj.SetBulge 9, -0.247182
    plineObj.Update
    plineObj.SetBulge 10, -0.221543
    plineObj.Update
    plineObj.SetBulge 11, -0.0976773
    plineObj.Update
    plineObj.SetBulge 12, -0.00917899
    plineObj.Update
    plineObj.SetBulge 13, 0.0324358
    plineObj.Update
    plineObj.SetBulge 14, 0.0324358
    plineObj.Update
    plineObj.SetBulge 15, -0.00917899
    plineObj.Update
    plineObj.SetBulge 16, -0.0976773
    plineObj.Update
    plineObj.SetBulge 17, -0.221543
    plineObj.Update
    plineObj.SetBulge 18, -0.247182
    plineObj.Update
    plineObj.SetBulge 19, 0.035764
    plineObj.Update
    plineObj.SetBulge 22, 0.035764
    plineObj.Update
    plineObj.SetBulge 23, -0.247182
    plineObj.Update
    plineObj.SetBulge 24, -0.221543
    plineObj.Update
    plineObj.SetBulge 25, -0.0976773
    plineObj.Update
    plineObj.SetBulge 26, -0.00917899
    plineObj.Update
    plineObj.SetBulge 27, 0.0324358
    plineObj.Update
    
    plineObj.Layer = "K-grav"
    plineObj.Update
    plineObj.Closed = True
  
  pointscntr2(0) = bcp + 0:         pointscntr2(1) = acp + 102.168
  pointscntr2(2) = bcp - 15.2911:   pointscntr2(3) = acp + 78.5624
  pointscntr2(4) = bcp + 15.2911:   pointscntr2(5) = acp + 78.5624
  
 Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr2)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  
    plineObj.SetBulge 0, 0.250892
    plineObj.Update
    plineObj.SetBulge 1, -0.332226
    plineObj.Update
    plineObj.SetBulge 2, 0.250892
    plineObj.Update
  b1(0) = points2(0) + (b / 2) - 1:  b1(1) = points2(3) - (a / 2)
  b2(0) = points2(0) + (b / 2) + 1:  b2(1) = points2(3) - (a / 2)
  RetVal = plineObj.Mirror(b1, b2)

  pointscntr2(0) = bcp + 15.7341:  pointscntr2(1) = acp + 73.5418
  pointscntr2(2) = bcp - 15.7341:  pointscntr2(3) = acp + 73.5418
  pointscntr2(4) = bcp + 0:        pointscntr2(5) = acp + 26.2803
  
 Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr2)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  
    plineObj.SetBulge 0, 0.419513
    plineObj.Update
    plineObj.SetBulge 1, 0.182166
    plineObj.Update
    plineObj.SetBulge 2, 0.182166
    plineObj.Update
    RetVal = plineObj.Mirror(b1, b2)
  
  pointscntr3(0) = bcp + 0:             pointscntr3(1) = acp + 147.04
  pointscntr3(2) = bcp - 4.1528:        pointscntr3(3) = acp + 107.621
  pointscntr3(4) = bcp + 1.13687E-12:   pointscntr3(5) = acp + 106.064
  pointscntr3(6) = bcp + 4.1528:        pointscntr3(7) = acp + 107.621
  
 Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr3)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  
    plineObj.SetBulge 1, -0.0342582
    plineObj.Update
    plineObj.SetBulge 2, -0.0342582
    plineObj.Update
    RetVal = plineObj.Mirror(b1, b2)
    
  pointscntr3(0) = bcp + 0:             pointscntr3(1) = acp + 20.5705
  pointscntr3(2) = bcp + 7.88663:       pointscntr3(3) = acp + 12.0123
  pointscntr3(4) = bcp + 0:             pointscntr3(5) = acp + 1.97396
  pointscntr3(6) = bcp - 7.88663:       pointscntr3(7) = acp + 12.0123
  
 Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr3)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  
    plineObj.SetBulge 0, 0.0192317
    plineObj.Update
    plineObj.SetBulge 1, -0.0134617
    plineObj.Update
    plineObj.SetBulge 2, -0.0134617
    plineObj.Update
    plineObj.SetBulge 3, 0.0192317
    plineObj.Update
    RetVal = plineObj.Mirror(b1, b2)
  
  pointscntr3(0) = bcp - 21.3616:      pointscntr3(1) = acp + 0
  pointscntr3(2) = bcp - 10.4078:      pointscntr3(3) = acp - 9.53775
  pointscntr3(4) = bcp - 2.90986:      pointscntr3(5) = acp + 0
  pointscntr3(6) = bcp - 10.4078:      pointscntr3(7) = acp + 9.53775
  
 Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr3)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  
    plineObj.SetBulge 0, -0.0240064
    plineObj.Update
    plineObj.SetBulge 1, -0.012607
    plineObj.Update
    plineObj.SetBulge 2, -0.012607
    plineObj.Update
    plineObj.SetBulge 3, -0.0240064
    plineObj.Update
  a1(0) = points2(4) - (b / 2):  a1(1) = points2(3) - (a / 2) + 1
  a2(0) = points2(4) - (b / 2):  a2(1) = points2(3) - (a / 2) - 1
  RetVal = plineObj.Mirror(a1, a2)
  
  pointscntr4(0) = bcp - 24.1296:     pointscntr4(1) = acp + 43.1205
  pointscntr4(2) = bcp - 19.2635:     pointscntr4(3) = acp + 68.8882
  pointscntr4(4) = bcp - 2.17621:     pointscntr4(5) = acp + 23.4154
  pointscntr4(6) = bcp - 9.97906:     pointscntr4(7) = acp + 14.8729
  pointscntr4(8) = bcp - 14.6543:     pointscntr4(9) = acp + 21.6084
  
 Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr4)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  
    plineObj.SetBulge 0, -0.186466
    plineObj.Update
    plineObj.SetBulge 1, 0.168477
    plineObj.Update
    plineObj.SetBulge 2, -0.0195719
    plineObj.Update
    plineObj.SetBulge 3, -0.00864503
    plineObj.Update
    plineObj.SetBulge 4, -0.0975544
    plineObj.Update
  RetVal = plineObj.Copy
  RetVal = plineObj.Mirror(a1, a2)
  RetVal = plineObj.Mirror(b1, b2)
  
 ' Define the rotation of 45 degrees about a
  ' base point of (4, 4.25, 0)
  basePoint(0) = points2(4) - (b / 2): basePoint(1) = points2(3) - (a / 2)
  rotationAngle = 3.14159   ' 45 degrees

  ' Rotate the polyline
  plineObj.Rotate basePoint, rotationAngle
  plineObj.Update
End If
End If
End If
End If
  
  points2(1) = points2(3) + I:
  points2(3) = points2(3) + a + I:
  points2(5) = points2(3):
  points2(7) = points2(1):
  

Else: d = 1
d = d + 1

d = Cells(c, 4)
End If

Next d
c = c + 1

a = Cells(c, 2)
b = Cells(c, 3)

I = 100

Loop
              
ZoomAll
AP.Quit
End Sub

Sub MF138()
' Creating search objects
Dim AP As Excel.Application
Dim WB As Excel.Workbook
Dim WS As Excel.Worksheet

Set AP = Excel.Application
Set WB = AP.Workbooks.Open("C:\Users\������\Desktop\������\2.xlsx")
Set WS = WB.Worksheets("������ �˨������ ������")
c = 19
a = Cells(c, 2)
b = Cells(c, 3)
d = Cells(c, 4)
E = 1
I = 0
On Error Resume Next

' Description of all the points that will make up rectangles
  Dim plineObj As AcadLWPolyline
  Dim plineObjw1 As AcadLWPolyline
  Dim plineObjw2 As AcadLWPolyline
  Dim points(0 To 7) As Double
  Dim pointsaktr1(0 To 209) As Double
  Dim pointsaktr2(0 To 17) As Double
  Dim pointsaktr3(0 To 39) As Double
  Dim pointsaktr4(0 To 47) As Double
  Dim pointsaktr5(0 To 43) As Double
  Dim pointscntr1(0 To 59) As Double
  Dim pointscntr2(0 To 55) As Double
  Dim pointscntr3(0 To 25) As Double
  Dim pointscntr4(0 To 39) As Double
  Dim pointscntr5(0 To 27) As Double
  Dim pointscntr6(0 To 11) As Double
  Dim pointswithin(0 To 31) As Double
  Dim pointswithin2(0 To 47) As Double
  Dim intPointsa
  Dim intPointsb
  Dim pointshelpa1(0 To 3) As Double
  Dim pointshelpa2(0 To 3) As Double
  Dim pointshelpb1(0 To 3) As Double
  Dim pointshelpb2(0 To 3) As Double
  Dim offsetObj As Variant
  Dim circleObj As AcadCircle
  Dim center(0 To 2) As Double
  Dim radius As Double

points(6) = 0

Do While a > 0

' Calculate the coordinates of each point
  points(0) = points(6) + I:    points(1) = 0
  points(2) = points(6) + I:    points(3) = a
  points(4) = points(0) + b:    points(5) = a
  points(6) = points(0) + b:    points(7) = 0

' Constructing four segments and connecting the corresponding points

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
  plineObj.Closed = True

If a > 280 Then
If b > 280 Then

  pointswithin(0) = points(0) + 100:                             pointswithin(1) = points(1) + 100
  pointswithin(2) = points(0) + 85:                              pointswithin(3) = points(1) + 100 + ((a - 200) / 4)
  pointswithin(4) = points(0) + 70:                              pointswithin(5) = points(1) + 100 + (2 * ((a - 200) / 4))
  pointswithin(6) = points(0) + 85:                              pointswithin(7) = points(1) + 100 + (3 * ((a - 200) / 4))
  pointswithin(8) = points(0) + 100:                             pointswithin(9) = points(3) - 100
  pointswithin(10) = points(0) + 100 + ((b - 200) / 4):          pointswithin(11) = points(3) - 85
  pointswithin(12) = points(0) + 100 + (2 * ((b - 200) / 4)):    pointswithin(13) = points(3) - 70
  pointswithin(14) = points(0) + 100 + (3 * ((b - 200) / 4)):    pointswithin(15) = points(3) - 85
  pointswithin(16) = points(4) - 100:                            pointswithin(17) = points(3) - 100
  pointswithin(18) = points(4) - 85:                             pointswithin(19) = points(3) - 100 - ((a - 200) / 4)
  pointswithin(20) = points(4) - 70:                             pointswithin(21) = points(3) - 100 - (2 * ((a - 200) / 4))
  pointswithin(22) = points(4) - 85:                             pointswithin(23) = points(3) - 100 - (3 * ((a - 200) / 4))
  pointswithin(24) = points(4) - 100:                            pointswithin(25) = points(1) + 100
  pointswithin(26) = points(4) - 100 - ((b - 200) / 4):          pointswithin(27) = points(1) + 85
  pointswithin(28) = points(4) - 100 - (2 * ((b - 200) / 4)):    pointswithin(29) = points(1) + 70
  pointswithin(30) = points(4) - 100 - (3 * ((b - 200) / 4)):    pointswithin(31) = points(1) + 85
  

If a > 400 Then
  

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointswithin)
  plineObj.Closed = True
   
   pb = (b - 200) / 4
   pa = (a - 200) / 4
ga = Sqr((pa * pa) + 225)
gb = Sqr((pb * pb) + 225)
   anglea = Atn((15 / ga) / Sqr((-15 / ga) * (15 / ga) + 1))
   angleb = Atn((15 / gb) / Sqr((-15 / gb) * (15 / gb) + 1))
   radiusa = (ga / 2) / Sin(15 / ga)
   radiusb = (gb / 2) / Sin(15 / gb)
   ha = radiusa * (1 - Cos(anglea))
   hb = radiusb * (1 - Cos(angleb))
   ka = ha / ga
   kb = hb / gb
   
    plineObj.SetBulge 0, ka * 2
    plineObj.SetBulge 1, -ka * 2
    plineObj.SetBulge 2, -ka * 2
    plineObj.SetBulge 3, ka * 2
    plineObj.SetBulge 4, kb * 2
    plineObj.SetBulge 5, -kb * 2
    plineObj.SetBulge 6, -kb * 2
    plineObj.SetBulge 7, kb * 2
    plineObj.SetBulge 8, ka * 2
    plineObj.SetBulge 9, -ka * 2
    plineObj.SetBulge 10, -ka * 2
    plineObj.SetBulge 11, ka * 2
    plineObj.SetBulge 12, kb * 2
    plineObj.SetBulge 13, -kb * 2
    plineObj.SetBulge 14, -kb * 2
    plineObj.SetBulge 15, kb * 2
    plineObj.Layer = "C-Mill"
    plineObj.Update
    plineObj.Closed = True

End If

  pointshelpa1(0) = points(0) + 79:                pointshelpa1(1) = points(1) + 100
  pointshelpa1(2) = points(0) + 49:                pointshelpa1(3) = points(1) + (a / 2)
  pointshelpa2(0) = points(0) + 100 - radiusa:     pointshelpa2(1) = points(1) + 100
  pointshelpa2(2) = pointswithin(2):               pointshelpa2(3) = pointswithin(3)
Set plineObjw1 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointshelpa1)
Set plineObjw2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointshelpa2)
intPointsa = plineObjw1.IntersectWith(plineObjw2, acExtendBoth)
Intersectionax = intPointsa(0) - points(0)
Intersectionay = intPointsa(1) - points(1)
plineObjw1.Delete
plineObjw2.Delete

  pointshelpb1(0) = points(0) + 100:     pointshelpb1(1) = points(3) - 79
  pointshelpb1(2) = points(0) + (b / 2): pointshelpb1(3) = points(3) - 49
  pointshelpb2(0) = points(0) + 100:     pointshelpb2(1) = points(3) - 100 + radiusb
  pointshelpb2(2) = pointswithin(10):    pointshelpb2(3) = pointswithin(11)
Set plineObjw1 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointshelpb1)
Set plineObjw2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointshelpb2)
intPointsb = plineObjw1.IntersectWith(plineObjw2, acExtendBoth)
Intersectionbx = intPointsb(0) - points(0)
Intersectionby = points(3) - intPointsb(1)
plineObjw1.Delete
plineObjw2.Delete
  
  pointswithin2(0) = points(0) + 79:                              pointswithin2(1) = points(1) + 79
  pointswithin2(2) = points(0) + 79:                              pointswithin2(3) = points(1) + 100
  pointswithin2(4) = points(0) + Intersectionax:                  pointswithin2(5) = points(1) + Intersectionay
  pointswithin2(6) = points(0) + 49:                              pointswithin2(7) = points(1) + 100 + (2 * ((a - 200) / 4))
  pointswithin2(8) = points(0) + Intersectionax:                  pointswithin2(9) = points(3) - Intersectionay
  pointswithin2(10) = points(0) + 79:                             pointswithin2(11) = points(3) - 100
  pointswithin2(12) = points(0) + 79:                             pointswithin2(13) = points(3) - 79
  pointswithin2(14) = points(0) + 100:                            pointswithin2(15) = points(3) - 79
  pointswithin2(16) = points(0) + Intersectionbx:                 pointswithin2(17) = points(3) - Intersectionby
  pointswithin2(18) = points(4) - 100 - (2 * ((b - 200) / 4)):    pointswithin2(19) = points(3) - 49
  pointswithin2(20) = points(4) - Intersectionbx:                 pointswithin2(21) = points(3) - Intersectionby
  pointswithin2(22) = points(4) - 100:                            pointswithin2(23) = points(3) - 79
  pointswithin2(24) = points(4) - 79:                             pointswithin2(25) = points(3) - 79
  pointswithin2(26) = points(4) - 79:                             pointswithin2(27) = points(3) - 100
  pointswithin2(28) = points(4) - Intersectionax:                 pointswithin2(29) = points(3) - Intersectionay
  pointswithin2(30) = points(4) - 49:                             pointswithin2(31) = points(1) + 100 + (2 * ((a - 200) / 4))
  pointswithin2(32) = points(4) - Intersectionax:                 pointswithin2(33) = points(1) + Intersectionay
  pointswithin2(34) = points(4) - 79:                             pointswithin2(35) = points(1) + 100
  pointswithin2(36) = points(4) - 79:                             pointswithin2(37) = points(1) + 79
  pointswithin2(38) = points(4) - 100:                            pointswithin2(39) = points(1) + 79
  pointswithin2(40) = points(4) - Intersectionbx:                 pointswithin2(41) = points(1) + Intersectionby
  pointswithin2(42) = points(4) - 100 - (2 * ((b - 200) / 4)):    pointswithin2(43) = points(1) + 49
  pointswithin2(44) = points(0) + Intersectionbx:                 pointswithin2(45) = points(1) + Intersectionby
  pointswithin2(46) = points(0) + 100:                            pointswithin2(47) = points(1) + 79

If a > 400 Then

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointswithin2)
  plineObj.Closed = True
  
    plineObj.SetBulge 1, ka * 2
    plineObj.SetBulge 2, -ka * 2
    plineObj.SetBulge 3, -ka * 2
    plineObj.SetBulge 4, ka * 2
    plineObj.SetBulge 7, kb * 2
    plineObj.SetBulge 8, -kb * 2
    plineObj.SetBulge 9, -kb * 2
    plineObj.SetBulge 10, kb * 2
    plineObj.SetBulge 13, ka * 2
    plineObj.SetBulge 14, -ka * 2
    plineObj.SetBulge 15, -ka * 2
    plineObj.SetBulge 16, ka * 2
    plineObj.SetBulge 19, kb * 2
    plineObj.SetBulge 20, -kb * 2
    plineObj.SetBulge 21, -kb * 2
    plineObj.SetBulge 22, kb * 2
    plineObj.Layer = "K-grav"
    plineObj.Update
    plineObj.Closed = True
  
End If


pointsaktr1(0) = points(0) + 41.3216:   pointsaktr1(1) = points(1) + 44.9698
pointsaktr1(2) = points(0) + 40.3478:   pointsaktr1(3) = points(1) + 43.1003
pointsaktr1(4) = points(0) + 46.1268:   pointsaktr1(5) = points(1) + 38.9822
pointsaktr1(6) = points(0) + 56.3386:   pointsaktr1(7) = points(1) + 50.9313
pointsaktr1(8) = points(0) + 44.8855:   pointsaktr1(9) = points(1) + 64.5693
pointsaktr1(10) = points(0) + 33.7952:  pointsaktr1(11) = points(1) + 64.5305
pointsaktr1(12) = points(0) + 24.0386:  pointsaktr1(13) = points(1) + 29.3573
pointsaktr1(14) = points(0) + 64.0284:  pointsaktr1(15) = points(1) + 21.9177
pointsaktr1(16) = points(0) + 127.102:  pointsaktr1(17) = points(1) + 21.8129
pointsaktr1(18) = points(0) + 127.198:  pointsaktr1(19) = points(1) + 21.6297
pointsaktr1(20) = points(0) + 127.198:  pointsaktr1(21) = points(1) + 21.1986
pointsaktr1(22) = points(0) + 127.382:  pointsaktr1(23) = points(1) + 20.9805
pointsaktr1(24) = points(0) + 128.388:  pointsaktr1(25) = points(1) + 20.3488
pointsaktr1(26) = points(0) + 130.165:  pointsaktr1(27) = points(1) + 19.536
pointsaktr1(28) = points(0) + 130.323:  pointsaktr1(29) = points(1) + 19.4341
pointsaktr1(30) = points(0) + 130.328:  pointsaktr1(31) = points(1) + 19.3365
pointsaktr1(32) = points(0) + 130.219:  pointsaktr1(33) = points(1) + 19.3111
pointsaktr1(34) = points(0) + 70.0257:  pointsaktr1(35) = points(1) + 19.4725
pointsaktr1(36) = points(0) + 29.517:   pointsaktr1(37) = points(1) + 20.6803
pointsaktr1(38) = points(0) + 17.9889:  pointsaktr1(39) = points(1) + 53.5889
pointsaktr1(40) = points(0) + 30.043:   pointsaktr1(41) = points(1) + 69.0807
pointsaktr1(42) = points(0) + 33.0582:  pointsaktr1(43) = points(1) + 72.7887
pointsaktr1(44) = points(0) + 33.8149:  pointsaktr1(45) = points(1) + 77.9079
pointsaktr1(46) = points(0) + 31.4709:  pointsaktr1(47) = points(1) + 79.0833
pointsaktr1(48) = points(0) + 29.5231:  pointsaktr1(49) = points(1) + 77.7696
pointsaktr1(50) = points(0) + 29.3443:  pointsaktr1(51) = points(1) + 77.788
pointsaktr1(52) = points(0) + 29.2837:  pointsaktr1(53) = points(1) + 78.7799
pointsaktr1(54) = points(0) + 30.6341:  pointsaktr1(55) = points(1) + 81.5798
pointsaktr1(56) = points(0) + 29.5258:  pointsaktr1(57) = points(1) + 87.1914
pointsaktr1(58) = points(0) + 24.6502:  pointsaktr1(59) = points(1) + 87.5247
pointsaktr1(60) = points(0) + 23.6658:  pointsaktr1(61) = points(1) + 83.7225
pointsaktr1(62) = points(0) + 25.7501:  pointsaktr1(63) = points(1) + 79.0141
pointsaktr1(64) = points(0) + 25.8284:  pointsaktr1(65) = points(1) + 78.4251
pointsaktr1(66) = points(0) + 25.6434:  pointsaktr1(67) = points(1) + 78.357
pointsaktr1(68) = points(0) + 22.852:   pointsaktr1(69) = points(1) + 79.0068
pointsaktr1(70) = points(0) + 20.5279:  pointsaktr1(71) = points(1) + 78.2691
pointsaktr1(72) = points(0) + 18.1972:  pointsaktr1(73) = points(1) + 77.4484
pointsaktr1(74) = points(0) + 16.347:   pointsaktr1(75) = points(1) + 77.9443
pointsaktr1(76) = points(0) + 15.3674:  pointsaktr1(77) = points(1) + 80.3287
pointsaktr1(78) = points(0) + 15.2094:  pointsaktr1(79) = points(1) + 89.2364
pointsaktr1(80) = points(0) + 15.7178:  pointsaktr1(81) = points(1) + 91.0514
pointsaktr1(82) = points(0) + 17.9348:  pointsaktr1(83) = points(1) + 92.3071
pointsaktr1(84) = points(0) + 19.838:   pointsaktr1(85) = points(1) + 91.9659
pointsaktr1(86) = points(0) + 21.9268:  pointsaktr1(87) = points(1) + 93.1289
pointsaktr1(88) = points(0) + 21.8411:  pointsaktr1(89) = points(1) + 93.7497
pointsaktr1(90) = points(0) + 17.2959:  pointsaktr1(91) = points(1) + 102.917
pointsaktr1(92) = points(0) + 15.5268:  pointsaktr1(93) = points(1) + 113.7
pointsaktr1(94) = points(0) + 20.041:   pointsaktr1(95) = points(1) + 120.229
pointsaktr1(96) = points(0) + 27.4677:  pointsaktr1(97) = points(1) + 119.229
pointsaktr1(98) = points(0) + 29.1061:  pointsaktr1(99) = points(1) + 114.516
pointsaktr1(100) = points(0) + 29.0205: pointsaktr1(101) = points(1) + 114.017
pointsaktr1(102) = points(0) + 24.0139: pointsaktr1(103) = points(1) + 112.877
pointsaktr1(104) = points(0) + 23.702:  pointsaktr1(105) = points(1) + 114.536
pointsaktr1(106) = points(0) + 23.5509: pointsaktr1(107) = points(1) + 116.32
pointsaktr1(108) = points(0) + 20.7646: pointsaktr1(109) = points(1) + 117.321
pointsaktr1(110) = points(0) + 17.4574: pointsaktr1(111) = points(1) + 113.936
pointsaktr1(112) = points(0) + 20.0439: pointsaktr1(113) = points(1) + 103.759
pointsaktr1(114) = points(0) + 28.9597: pointsaktr1(115) = points(1) + 90.4939
pointsaktr1(116) = points(0) + 40.2373: pointsaktr1(117) = points(1) + 78.0111
pointsaktr1(118) = points(0) + 42.7351: pointsaktr1(119) = points(1) + 73.81
pointsaktr1(120) = points(0) + 42.6941: pointsaktr1(121) = points(1) + 71.0615
pointsaktr1(122) = points(0) + 40.3146: pointsaktr1(123) = points(1) + 69.1187
pointsaktr1(124) = points(0) + 40.3297: pointsaktr1(125) = points(1) + 68.899
pointsaktr1(126) = points(0) + 55.6191: pointsaktr1(127) = points(1) + 56.4976
pointsaktr1(128) = points(0) + 58.2067: pointsaktr1(129) = points(1) + 47.14
pointsaktr1(130) = points(0) + 58.8506: pointsaktr1(131) = points(1) + 39.5231
pointsaktr1(132) = points(0) + 67.8108: pointsaktr1(133) = points(1) + 28.2563
pointsaktr1(134) = points(0) + 79.5835: pointsaktr1(135) = points(1) + 26.1017
pointsaktr1(136) = points(0) + 84.365:  pointsaktr1(137) = points(1) + 25.9125
pointsaktr1(138) = points(0) + 84.8383: pointsaktr1(139) = points(1) + 25.8106
pointsaktr1(140) = points(0) + 84.9932: pointsaktr1(141) = points(1) + 25.4054
pointsaktr1(142) = points(0) + 83.696:  pointsaktr1(143) = points(1) + 24.3908
pointsaktr1(144) = points(0) + 80.3696: pointsaktr1(145) = points(1) + 22.5492
pointsaktr1(146) = points(0) + 80.2486: pointsaktr1(147) = points(1) + 22.5346
pointsaktr1(148) = points(0) + 76.4568: pointsaktr1(149) = points(1) + 22.5275
pointsaktr1(150) = points(0) + 76.3508: pointsaktr1(151) = points(1) + 22.6946
pointsaktr1(152) = points(0) + 75.9307: pointsaktr1(153) = points(1) + 23.9924
pointsaktr1(154) = points(0) + 73.6332: pointsaktr1(155) = points(1) + 24.5453
pointsaktr1(156) = points(0) + 68.436:  pointsaktr1(157) = points(1) + 24.592
pointsaktr1(158) = points(0) + 63.6462: pointsaktr1(159) = points(1) + 24.8968
pointsaktr1(160) = points(0) + 58.4456: pointsaktr1(161) = points(1) + 28.7795
pointsaktr1(162) = points(0) + 58.1362: pointsaktr1(163) = points(1) + 29.3731
pointsaktr1(164) = points(0) + 58.0776: pointsaktr1(165) = points(1) + 29.4447
pointsaktr1(166) = points(0) + 57.4768: pointsaktr1(167) = points(1) + 29.4281
pointsaktr1(168) = points(0) + 57.0179: pointsaktr1(169) = points(1) + 28.8684
pointsaktr1(170) = points(0) + 56.6262: pointsaktr1(171) = points(1) + 28.0932
pointsaktr1(172) = points(0) + 56.4903: pointsaktr1(173) = points(1) + 27.8301
pointsaktr1(174) = points(0) + 55.7942: pointsaktr1(175) = points(1) + 27.1122
pointsaktr1(176) = points(0) + 55.2031: pointsaktr1(177) = points(1) + 27.0389
pointsaktr1(178) = points(0) + 55.0785: pointsaktr1(179) = points(1) + 27.1134
pointsaktr1(180) = points(0) + 53.4121: pointsaktr1(181) = points(1) + 33.4249
pointsaktr1(182) = points(0) + 54.6571: pointsaktr1(183) = points(1) + 38.7189
pointsaktr1(184) = points(0) + 54.4842: pointsaktr1(185) = points(1) + 38.8578
pointsaktr1(186) = points(0) + 50.9573: pointsaktr1(187) = points(1) + 37.2303
pointsaktr1(188) = points(0) + 39.1485: pointsaktr1(189) = points(1) + 40.5163
pointsaktr1(190) = points(0) + 36.3064: pointsaktr1(191) = points(1) + 45.3582
pointsaktr1(192) = points(0) + 36.6083: pointsaktr1(193) = points(1) + 47.1011
pointsaktr1(194) = points(0) + 37.8654: pointsaktr1(195) = points(1) + 48.2559
pointsaktr1(196) = points(0) + 44.2299: pointsaktr1(197) = points(1) + 54.6583
pointsaktr1(198) = points(0) + 48.0345: pointsaktr1(199) = points(1) + 62.0273
pointsaktr1(200) = points(0) + 48.2181: pointsaktr1(201) = points(1) + 62.0909
pointsaktr1(202) = points(0) + 54.9901: pointsaktr1(203) = points(1) + 53.5569
pointsaktr1(204) = points(0) + 54.9939: pointsaktr1(205) = points(1) + 53.4679
pointsaktr1(206) = points(0) + 50.0612: pointsaktr1(207) = points(1) + 46.176
pointsaktr1(208) = points(0) + 47.4251: pointsaktr1(209) = points(1) + 45.3091


Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsaktr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update
    
    ' Find the bulge of the third segment
    currentBulge = plineObj.GetBulge(2)
    ' Change the bulge of the third segment

plineObj.SetBulge 0, 0.409159
plineObj.SetBulge 1, 0.321957
plineObj.SetBulge 2, 0.468929
plineObj.SetBulge 3, 0.242406
plineObj.SetBulge 4, 0.181208
plineObj.SetBulge 5, 0.481392
plineObj.SetBulge 6, 0.380589
plineObj.SetBulge 7, 0
plineObj.SetBulge 8, -0.607953
plineObj.SetBulge 9, 0.32111
plineObj.SetBulge 10, 0.0696001
plineObj.SetBulge 11, 0.063976
plineObj.SetBulge 12, 0
plineObj.SetBulge 13, 0
plineObj.SetBulge 14, -0.502036
plineObj.SetBulge 15, -0.126379
plineObj.SetBulge 16, 0
plineObj.SetBulge 17, -0.246237
plineObj.SetBulge 18, -0.379673
plineObj.SetBulge 19, -0.102627
plineObj.SetBulge 20, 0.0894177
plineObj.SetBulge 21, 0.192392
plineObj.SetBulge 22, 0.46922
plineObj.SetBulge 23, 0.104315
plineObj.SetBulge 24, -0.437919
plineObj.SetBulge 25, -0.274694
plineObj.SetBulge 26, 0.0834997
plineObj.SetBulge 27, 0.259096
plineObj.SetBulge 28, 0.460234
plineObj.SetBulge 29, 0.247169
plineObj.SetBulge 30, 0.102358
plineObj.SetBulge 31, -0.197867
plineObj.SetBulge 32, -0.553072
plineObj.SetBulge 33, 0.213683
plineObj.SetBulge 34, 0.0574858
plineObj.SetBulge 35, -0.0523247
plineObj.SetBulge 36, -0.26884
plineObj.SetBulge 37, -0.136043
plineObj.SetBulge 38, -0.0629278
plineObj.SetBulge 39, -0.0675464
plineObj.SetBulge 40, -0.359276
plineObj.SetBulge 41, -0.00859763
plineObj.SetBulge 42, 0.355483
plineObj.SetBulge 43, 0.203986
plineObj.SetBulge 44, -0.039733
plineObj.SetBulge 45, -0.124994
plineObj.SetBulge 46, -0.254634
plineObj.SetBulge 47, -0.332616
plineObj.SetBulge 48, -0.179762
plineObj.SetBulge 49, -0.0847805
plineObj.SetBulge 50, -0.521548
plineObj.SetBulge 51, -0.278007
plineObj.SetBulge 52, 0.277679
plineObj.SetBulge 53, 0.264017
plineObj.SetBulge 54, 0.35711
plineObj.SetBulge 55, 0.152853
plineObj.SetBulge 56, 0.0530635
plineObj.SetBulge 57, 0
plineObj.SetBulge 58, -0.0707038
plineObj.SetBulge 59, -0.212213
plineObj.SetBulge 60, -0.171738
plineObj.SetBulge 61, 0.715161
plineObj.SetBulge 62, -0.217317
plineObj.SetBulge 63, -0.104569
plineObj.SetBulge 64, 0.0278538
plineObj.SetBulge 65, 0.303024
plineObj.SetBulge 66, 0.0691689
plineObj.SetBulge 67, 0
plineObj.SetBulge 68, -0.0324659
plineObj.SetBulge 69, -0.563877
plineObj.SetBulge 70, -0.090811
plineObj.SetBulge 71, 0
plineObj.SetBulge 72, -0.0878322
plineObj.SetBulge 73, 0
plineObj.SetBulge 74, -0.554362
plineObj.SetBulge 75, 0.465455
plineObj.SetBulge 76, 0.0895132
plineObj.SetBulge 77, 0.0326569
plineObj.SetBulge 78, -0.0784946
plineObj.SetBulge 79, -0.218538
plineObj.SetBulge 80, 0
plineObj.SetBulge 81, 0.0921223
plineObj.SetBulge 82, 0.256818
plineObj.SetBulge 83, 0.13945
plineObj.SetBulge 84, 0
plineObj.SetBulge 85, 0
plineObj.SetBulge 86, -0.160405
plineObj.SetBulge 87, -0.146033
plineObj.SetBulge 88, -0.191927
plineObj.SetBulge 89, -0.185968
plineObj.SetBulge 90, -0.062917
plineObj.SetBulge 91, 0.72199
plineObj.SetBulge 92, -0.0715105
plineObj.SetBulge 93, -0.271829
plineObj.SetBulge 94, -0.141216
plineObj.SetBulge 95, -0.209519
plineObj.SetBulge 96, -0.0278354
plineObj.SetBulge 97, 0.0505341
plineObj.SetBulge 98, 0.100553
plineObj.SetBulge 99, -0.53094
plineObj.SetBulge 100, -0.125904
plineObj.SetBulge 101, -0.0996162
plineObj.SetBulge 102, -0.148082
plineObj.SetBulge 103, -0.203841
plineObj.SetBulge 104, 0.135952
    
    plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True
  
  Dim b1(0 To 2) As Double
  Dim b2(0 To 2) As Double
  b1(0) = (b / 2) - 1:  b1(1) = (a / 2)
  b2(0) = (b / 2) + 1:  b2(1) = (a / 2)
  RetVal = plineObj.Mirror(b1, b2)
  Dim a1(0 To 2) As Double
  Dim a2(0 To 2) As Double
  a1(0) = points(4) - (b / 2): a1(1) = points(1) + (a / 2) - 1
  a2(0) = points(4) - (b / 2): a2(1) = points(1) + (a / 2) + 1
  RetVal = plineObj.Mirror(a1, a2)
  
  RetVal = plineObj.Copy
  ' Define the rotation of 45 degrees about a
  ' base point of (4, 4.25, 0)
  Dim basePoint(0 To 2) As Double
  Dim rotationAngle As Double
  basePoint(0) = points(4) - (b / 2): basePoint(1) = points(3) - (a / 2)
  rotationAngle = 3.14159   ' 45 degrees

  ' Rotate the polyline
  plineObj.Rotate basePoint, rotationAngle
  plineObj.Update

pointsaktr2(0) = points(0) + 36.8503:   pointsaktr2(1) = points(1) + 86.3405
pointsaktr2(2) = points(0) + 36.9528:   pointsaktr2(3) = points(1) + 86.5422
pointsaktr2(4) = points(0) + 46.0287:   pointsaktr2(5) = points(1) + 80.7521
pointsaktr2(6) = points(0) + 51.7303:   pointsaktr2(7) = points(1) + 74.7275
pointsaktr2(8) = points(0) + 56.0192:   pointsaktr2(9) = points(1) + 57.4278
pointsaktr2(10) = points(0) + 55.8011:  pointsaktr2(11) = points(1) + 57.3905
pointsaktr2(12) = points(0) + 48.6202:  pointsaktr2(13) = points(1) + 65.651
pointsaktr2(14) = points(0) + 48.5738:  pointsaktr2(15) = points(1) + 65.7471
pointsaktr2(16) = points(0) + 47.7344:  pointsaktr2(17) = points(1) + 71.4489

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsaktr2)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update
  
plineObj.SetBulge 0, -0.771138
plineObj.SetBulge 1, -0.0611678
plineObj.SetBulge 2, -0.0629979
plineObj.SetBulge 3, -0.201444
plineObj.SetBulge 4, -0.714409
plineObj.SetBulge 5, 0.106484
plineObj.SetBulge 6, -0.250546
plineObj.SetBulge 7, 0.0841751
plineObj.SetBulge 8, 0.130058

 plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True
  
  b1(0) = (b / 2) - 1:  b1(1) = (a / 2)
  b2(0) = (b / 2) + 1:  b2(1) = (a / 2)
  RetVal = plineObj.Mirror(b1, b2)
  a1(0) = points(4) - (b / 2): a1(1) = points(1) + (a / 2) - 1
  a2(0) = points(4) - (b / 2): a2(1) = points(1) + (a / 2) + 1
  RetVal = plineObj.Mirror(a1, a2)
  RetVal = plineObj.Copy
  ' Define the rotation of 45 degrees about a
  ' base point of (4, 4.25, 0)
  basePoint(0) = points(4) - (b / 2): basePoint(1) = points(3) - (a / 2)
  rotationAngle = 3.14159   ' 45 degrees

  ' Rotate the polyline
  plineObj.Rotate basePoint, rotationAngle
  plineObj.Update
  
pointsaktr3(0) = points(0) + 135.147:   pointsaktr3(1) = points(1) + 17.7302
pointsaktr3(2) = points(0) + 133.843:   pointsaktr3(3) = points(1) + 17.5545
pointsaktr3(4) = points(0) + 131.609:   pointsaktr3(5) = points(1) + 17.4606
pointsaktr3(6) = points(0) + 81.121:    pointsaktr3(7) = points(1) + 17.4146
pointsaktr3(8) = points(0) + 76.3148:   pointsaktr3(9) = points(1) + 17.9255
pointsaktr3(10) = points(0) + 75.7965:  pointsaktr3(11) = points(1) + 18.1653
pointsaktr3(12) = points(0) + 75.4241:  pointsaktr3(13) = points(1) + 18.7953
pointsaktr3(14) = points(0) + 75.3105:  pointsaktr3(15) = points(1) + 18.8943
pointsaktr3(16) = points(0) + 72.3454:  pointsaktr3(17) = points(1) + 18.8905
pointsaktr3(18) = points(0) + 72.2831:  pointsaktr3(19) = points(1) + 18.8724
pointsaktr3(20) = points(0) + 69.6728:  pointsaktr3(21) = points(1) + 16.5045
pointsaktr3(22) = points(0) + 70.0556:  pointsaktr3(23) = points(1) + 15.3657
pointsaktr3(24) = points(0) + 72.3199:  pointsaktr3(25) = points(1) + 15.0109
pointsaktr3(26) = points(0) + 137.437:  pointsaktr3(27) = points(1) + 15.0198
pointsaktr3(28) = points(0) + 138.883:  pointsaktr3(29) = points(1) + 15.2189
pointsaktr3(30) = points(0) + 138.929:  pointsaktr3(31) = points(1) + 15.4766
pointsaktr3(32) = points(0) + 138.302:  pointsaktr3(33) = points(1) + 15.8567
pointsaktr3(34) = points(0) + 136.44:   pointsaktr3(35) = points(1) + 16.6981
pointsaktr3(36) = points(0) + 135.583:  pointsaktr3(37) = points(1) + 17.2219
pointsaktr3(38) = points(0) + 135.283:  pointsaktr3(39) = points(1) + 17.6465

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsaktr3)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update
  
plineObj.SetBulge 0, -0.0317584
plineObj.SetBulge 1, -0.0137171
plineObj.SetBulge 2, 0
plineObj.SetBulge 3, -0.0603404
plineObj.SetBulge 4, -0.0618883
plineObj.SetBulge 5, -0.27718
plineObj.SetBulge 6, 0.388064
plineObj.SetBulge 7, 0
plineObj.SetBulge 8, 0.103341
plineObj.SetBulge 9, 0.124064
plineObj.SetBulge 10, 0.490174
plineObj.SetBulge 11, 0.0900062
plineObj.SetBulge 12, 0
plineObj.SetBulge 13, 0.0642505
plineObj.SetBulge 14, 0.673959
plineObj.SetBulge 15, 0.0711555
plineObj.SetBulge 16, 0
plineObj.SetBulge 17, -0.0526825
plineObj.SetBulge 18, -0.175037
plineObj.SetBulge 19, 0.393435

 plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True
  
  RetVal = plineObj.Mirror(b1, b2)
  RetVal = plineObj.Mirror(a1, a2)
  RetVal = plineObj.Copy
  ' Define the rotation of 45 degrees about a
  ' base point of (4, 4.25, 0)
  basePoint(0) = points(4) - (b / 2): basePoint(1) = points(3) - (a / 2)
  rotationAngle = 3.14159   ' 45 degrees
  plineObj.Rotate basePoint, rotationAngle
  plineObj.Update
  
pointsaktr4(0) = points(0) + 17.45:     pointsaktr4(1) = points(1) + 149.5
pointsaktr4(2) = points(0) + 17.45:     pointsaktr4(3) = points(1) + 132.254
pointsaktr4(4) = points(0) + 17.7239:   pointsaktr4(5) = points(1) + 128.781
pointsaktr4(6) = points(0) + 17.6402:   pointsaktr4(7) = points(1) + 128.65
pointsaktr4(8) = points(0) + 17.2757:   pointsaktr4(9) = points(1) + 128.422
pointsaktr4(10) = points(0) + 16.6942:  pointsaktr4(11) = points(1) + 127.512
pointsaktr4(12) = points(0) + 16.0544:  pointsaktr4(13) = points(1) + 126.126
pointsaktr4(14) = points(0) + 15.4754:  pointsaktr4(15) = points(1) + 125.07
pointsaktr4(16) = points(0) + 15.245:   pointsaktr4(17) = points(1) + 125.064
pointsaktr4(18) = points(0) + 15.1237:  pointsaktr4(19) = points(1) + 125.425
pointsaktr4(20) = points(0) + 15.0098:  pointsaktr4(21) = points(1) + 127.043
pointsaktr4(22) = points(0) + 15:       pointsaktr4(23) = points(1) + 149.5
pointsaktr4(24) = points(0) + 15:       pointsaktr4(25) = points(3) - 149.5
pointsaktr4(26) = points(0) + 15.0098:  pointsaktr4(27) = points(3) - 127.043
pointsaktr4(28) = points(0) + 15.1237:  pointsaktr4(29) = points(3) - 125.425
pointsaktr4(30) = points(0) + 15.245:   pointsaktr4(31) = points(3) - 125.064
pointsaktr4(32) = points(0) + 15.4754:  pointsaktr4(33) = points(3) - 125.07
pointsaktr4(34) = points(0) + 16.0544:  pointsaktr4(35) = points(3) - 126.126
pointsaktr4(36) = points(0) + 16.6942:  pointsaktr4(37) = points(3) - 127.512
pointsaktr4(38) = points(0) + 17.2757:   pointsaktr4(39) = points(3) - 128.422
pointsaktr4(40) = points(0) + 17.6402:   pointsaktr4(41) = points(3) - 128.65
pointsaktr4(42) = points(0) + 17.7239:   pointsaktr4(43) = points(3) - 128.781
pointsaktr4(44) = points(0) + 17.45:     pointsaktr4(45) = points(3) - 132.254
pointsaktr4(46) = points(0) + 17.45:     pointsaktr4(47) = points(3) - 149.5

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsaktr4)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update
  
plineObj.SetBulge 0, 0
plineObj.SetBulge 1, 0.0393737
plineObj.SetBulge 2, -0.380001
plineObj.SetBulge 3, 0.142137
plineObj.SetBulge 4, 0.0583339
plineObj.SetBulge 5, 0
plineObj.SetBulge 6, -0.0650531
plineObj.SetBulge 7, -0.486421
plineObj.SetBulge 8, -0.0826229
plineObj.SetBulge 9, -0.0364452
plineObj.SetBulge 10, 0
plineObj.SetBulge 11, 0
plineObj.SetBulge 12, 0
plineObj.SetBulge 13, -0.0364452
plineObj.SetBulge 14, -0.0826229
plineObj.SetBulge 15, -0.486421
plineObj.SetBulge 16, -0.0650531
plineObj.SetBulge 17, 0
plineObj.SetBulge 18, 0.0583339
plineObj.SetBulge 19, 0.142137
plineObj.SetBulge 20, -0.380001
plineObj.SetBulge 21, 0.0393737
plineObj.SetBulge 22, 0

RetVal = plineObj.Mirror(a1, a2)

pointsaktr5(0) = points(0) + 21.9:      pointsaktr5(1) = points(1) + 149.5
pointsaktr5(2) = points(0) + 22.1744:   pointsaktr5(3) = points(1) + 146.286
pointsaktr5(4) = points(0) + 22.1018:   pointsaktr5(5) = points(1) + 146.177
pointsaktr5(6) = points(0) + 21.7847:   pointsaktr5(7) = points(1) + 146.024
pointsaktr5(8) = points(0) + 21.0323:   pointsaktr5(9) = points(1) + 144.982
pointsaktr5(10) = points(0) + 20.3048:  pointsaktr5(11) = points(1) + 143.596
pointsaktr5(12) = points(0) + 19.8882:  pointsaktr5(13) = points(1) + 143.03
pointsaktr5(14) = points(0) + 19.6951:  pointsaktr5(15) = points(1) + 143.056
pointsaktr5(16) = points(0) + 19.5185:  pointsaktr5(17) = points(1) + 143.696
pointsaktr5(18) = points(0) + 19.45:    pointsaktr5(19) = points(1) + 145.219
pointsaktr5(20) = points(0) + 19.45:    pointsaktr5(21) = points(1) + 149.5
pointsaktr5(22) = points(0) + 19.45:    pointsaktr5(23) = points(3) - 149.5
pointsaktr5(24) = points(0) + 19.45:    pointsaktr5(25) = points(3) - 145.219
pointsaktr5(26) = points(0) + 19.5185:  pointsaktr5(27) = points(3) - 143.696
pointsaktr5(28) = points(0) + 19.6951:  pointsaktr5(29) = points(3) - 143.056
pointsaktr5(30) = points(0) + 19.8882:  pointsaktr5(31) = points(3) - 143.03
pointsaktr5(32) = points(0) + 20.3048:  pointsaktr5(33) = points(3) - 143.596
pointsaktr5(34) = points(0) + 21.0323:  pointsaktr5(35) = points(3) - 144.982
pointsaktr5(36) = points(0) + 21.7847:  pointsaktr5(37) = points(3) - 146.024
pointsaktr5(38) = points(0) + 22.1018:  pointsaktr5(39) = points(3) - 146.177
pointsaktr5(40) = points(0) + 22.1744:  pointsaktr5(41) = points(3) - 146.286
pointsaktr5(42) = points(0) + 21.9:     pointsaktr5(43) = points(3) - 149.5

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsaktr5)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update
  
plineObj.SetBulge 0, 0.0426022
plineObj.SetBulge 1, -0.398863
plineObj.SetBulge 2, 0.114352
plineObj.SetBulge 3, 0.0884516
plineObj.SetBulge 4, 0
plineObj.SetBulge 5, -0.0920327
plineObj.SetBulge 6, -0.457613
plineObj.SetBulge 7, -0.10162
plineObj.SetBulge 8, -0.0270408
plineObj.SetBulge 9, 0
plineObj.SetBulge 10, 0
plineObj.SetBulge 11, 0
plineObj.SetBulge 12, -0.0270408
plineObj.SetBulge 13, -0.10162
plineObj.SetBulge 14, -0.457613
plineObj.SetBulge 15, -0.0920327
plineObj.SetBulge 16, 0
plineObj.SetBulge 17, 0.0884516
plineObj.SetBulge 18, 0.114352
plineObj.SetBulge 19, -0.398863
plineObj.SetBulge 20, 0.0426022

RetVal = plineObj.Mirror(a1, a2)

'=============================CENTER===========================================
bcp = points(0) + (b / 2)
acp = points(1) + (a / 2)

pointscntr1(0) = bcp + 24.5833:     pointscntr1(1) = acp + 20.0873
pointscntr1(2) = bcp + 28.0246:     pointscntr1(3) = acp + 19.9972
pointscntr1(4) = bcp + 19.3482:     pointscntr1(5) = acp + 25.688
pointscntr1(6) = bcp + 26.8683:     pointscntr1(7) = acp + 32.6531
pointscntr1(8) = bcp + 27.427:      pointscntr1(9) = acp + 33.6769
pointscntr1(10) = bcp + 27.9934:    pointscntr1(11) = acp + 34.4973
pointscntr1(12) = bcp + 31.4343:    pointscntr1(13) = acp + 35.7712
pointscntr1(14) = bcp + 33.8541:    pointscntr1(15) = acp + 35.1915
pointscntr1(16) = bcp + 37:         pointscntr1(17) = acp + 33.4752
pointscntr1(18) = bcp + 30.8536:    pointscntr1(19) = acp + 38.2705
pointscntr1(20) = bcp + 23.058:     pointscntr1(21) = acp + 38.2446
pointscntr1(22) = bcp + 13.2405:    pointscntr1(23) = acp + 32.8606
pointscntr1(24) = bcp + 6.64503:    pointscntr1(25) = acp + 23.8124
pointscntr1(26) = bcp + 6.64503:    pointscntr1(27) = acp - 23.8124
pointscntr1(28) = bcp + 13.2405:    pointscntr1(29) = acp - 32.8606
pointscntr1(30) = bcp + 23.058:     pointscntr1(31) = acp - 38.2446
pointscntr1(32) = bcp + 30.8536:    pointscntr1(33) = acp - 38.2705
pointscntr1(34) = bcp + 37:         pointscntr1(35) = acp - 33.4752
pointscntr1(36) = bcp + 33.8541:    pointscntr1(37) = acp - 35.1915
pointscntr1(38) = bcp + 31.4343:    pointscntr1(39) = acp - 35.7712
pointscntr1(40) = bcp + 27.9934:    pointscntr1(41) = acp - 34.4973
pointscntr1(42) = bcp + 27.427:     pointscntr1(43) = acp - 33.6769
pointscntr1(44) = bcp + 26.8683:    pointscntr1(45) = acp - 32.6531
pointscntr1(46) = bcp + 19.3482:    pointscntr1(47) = acp - 25.688
pointscntr1(48) = bcp + 28.0246:    pointscntr1(49) = acp - 19.9972
pointscntr1(50) = bcp + 24.5833:    pointscntr1(51) = acp - 20.0873
pointscntr1(52) = bcp + 19.6471:    pointscntr1(53) = acp - 16.403
pointscntr1(54) = bcp + 6.28296:    pointscntr1(55) = acp - 9.466
pointscntr1(56) = bcp + 6.28296:    pointscntr1(57) = acp + 9.466
pointscntr1(58) = bcp + 19.6471:    pointscntr1(59) = acp + 16.403

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update

plineObj.SetBulge 0, -0.190956
plineObj.SetBulge 1, 0.493502
plineObj.SetBulge 2, 0.173886
plineObj.SetBulge 3, 0
plineObj.SetBulge 4, -0.0792127
plineObj.SetBulge 5, -0.236352
plineObj.SetBulge 6, -0.0741078
plineObj.SetBulge 7, -0.0595151
plineObj.SetBulge 8, 0.188765
plineObj.SetBulge 9, 0.132767
plineObj.SetBulge 10, 0.129543
plineObj.SetBulge 11, 0.0905785
plineObj.SetBulge 12, 0.219599
plineObj.SetBulge 13, 0.0905785
plineObj.SetBulge 14, 0.129543
plineObj.SetBulge 15, 0.132767
plineObj.SetBulge 16, 0.188765
plineObj.SetBulge 17, -0.0595151
plineObj.SetBulge 18, -0.0741078
plineObj.SetBulge 19, -0.236352
plineObj.SetBulge 20, -0.0792127
plineObj.SetBulge 21, 0
plineObj.SetBulge 22, 0.173886
plineObj.SetBulge 23, 0.493502
plineObj.SetBulge 24, -0.190956
plineObj.SetBulge 25, -0.133032
plineObj.SetBulge 26, -0.20958
plineObj.SetBulge 27, -0.334034
plineObj.SetBulge 28, -0.20958
plineObj.SetBulge 29, -0.133032

plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True

RetVal = plineObj.Mirror(a1, a2)

pointscntr2(0) = bcp + 6.58238: pointscntr2(1) = acp + 31.3788
pointscntr2(2) = bcp + 4.25321: pointscntr2(3) = acp + 33.6762
pointscntr2(4) = bcp + 0.898097:    pointscntr2(5) = acp + 33.4752
pointscntr2(6) = bcp + 1.68328: pointscntr2(7) = acp + 35.9463
pointscntr2(8) = bcp + 1.70714: pointscntr2(9) = acp + 37.0063
pointscntr2(10) = bcp + 0.782447:   pointscntr2(11) = acp + 37.995
pointscntr2(12) = bcp + 1.13687E-13:    pointscntr2(13) = acp + 37.9676
pointscntr2(14) = bcp + -0.782447:  pointscntr2(15) = acp + 37.995
pointscntr2(16) = bcp + -1.70714:   pointscntr2(17) = acp + 37.0063
pointscntr2(18) = bcp + -1.68328:   pointscntr2(19) = acp + 35.9463
pointscntr2(20) = bcp + -0.898097:  pointscntr2(21) = acp + 33.4752
pointscntr2(22) = bcp + -4.25321:   pointscntr2(23) = acp + 33.6762
pointscntr2(24) = bcp + -6.58238:   pointscntr2(25) = acp + 31.3788
pointscntr2(26) = bcp + -3.92174:   pointscntr2(27) = acp + 29.4618
pointscntr2(28) = bcp + -1.70295:   pointscntr2(29) = acp + 25.1
pointscntr2(30) = bcp + -0.455718:  pointscntr2(31) = acp + 18.477
pointscntr2(32) = bcp + -0.227606:  pointscntr2(33) = acp + 15.167
pointscntr2(34) = bcp + -0.110535:  pointscntr2(35) = acp + 11.4844
pointscntr2(36) = bcp + -0.0689475: pointscntr2(37) = acp + 10.7799
pointscntr2(38) = bcp + -0.0388717: pointscntr2(39) = acp + 10.5512
pointscntr2(40) = bcp + 1.13687E-13:    pointscntr2(41) = acp + 10.4124
pointscntr2(42) = bcp + 0.0388717:  pointscntr2(43) = acp + 10.5512
pointscntr2(44) = bcp + 0.0689475:  pointscntr2(45) = acp + 10.7799
pointscntr2(46) = bcp + 0.110535:   pointscntr2(47) = acp + 11.4844
pointscntr2(48) = bcp + 0.227606:   pointscntr2(49) = acp + 15.167
pointscntr2(50) = bcp + 0.455718:   pointscntr2(51) = acp + 18.477
pointscntr2(52) = bcp + 1.70295:    pointscntr2(53) = acp + 25.1
pointscntr2(54) = bcp + 3.92174:    pointscntr2(55) = acp + 29.4618


Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr2)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update

plineObj.SetBulge 0, 0.130293
plineObj.SetBulge 1, 0.286262
plineObj.SetBulge 2, 0.0628827
plineObj.SetBulge 3, 0.0664767
plineObj.SetBulge 4, 0.354054
plineObj.SetBulge 5, 0.0749942
plineObj.SetBulge 6, 0.0749942
plineObj.SetBulge 7, 0.354054
plineObj.SetBulge 8, 0.0664767
plineObj.SetBulge 9, 0.0628827
plineObj.SetBulge 10, 0.286262
plineObj.SetBulge 11, 0.130293
plineObj.SetBulge 12, -0.132961
plineObj.SetBulge 13, -0.0949138
plineObj.SetBulge 14, -0.0510213
plineObj.SetBulge 15, -0.0125717
plineObj.SetBulge 16, -0.00334278
plineObj.SetBulge 17, 0
plineObj.SetBulge 18, 0
plineObj.SetBulge 19, 0
plineObj.SetBulge 20, 0
plineObj.SetBulge 21, 0
plineObj.SetBulge 22, 0
plineObj.SetBulge 23, -0.00334278
plineObj.SetBulge 24, -0.0125717
plineObj.SetBulge 25, -0.0510213
plineObj.SetBulge 26, -0.0949138
plineObj.SetBulge 27, -0.132961

plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True

RetVal = plineObj.Mirror(b1, b2)

pointscntr3(0) = bcp + 17.03:       pointscntr3(1) = acp + 80.0281
pointscntr3(2) = bcp + 12.3194:     pointscntr3(3) = acp + 76.2511
pointscntr3(4) = bcp + 12.7291:     pointscntr3(5) = acp + 73.3674
pointscntr3(6) = bcp + 16.4548:     pointscntr3(7) = acp + 74.8083
pointscntr3(8) = bcp + 22.4096:     pointscntr3(9) = acp + 72.6438
pointscntr3(10) = bcp + 24.1819:    pointscntr3(11) = acp + 69.1113
pointscntr3(12) = bcp + 24.8313:    pointscntr3(13) = acp + 64.3255
pointscntr3(14) = bcp + 23.9272:    pointscntr3(15) = acp + 59.0018
pointscntr3(16) = bcp + 16.5438:    pointscntr3(17) = acp + 46.828
pointscntr3(18) = bcp + 6.28228:    pointscntr3(19) = acp + 36.4703
pointscntr3(20) = bcp + 22.0575:    pointscntr3(21) = acp + 48.4905
pointscntr3(22) = bcp + 28.7564:    pointscntr3(23) = acp + 65.9369
pointscntr3(24) = bcp + 25.5548:    pointscntr3(25) = acp + 74.964

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr3)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update

plineObj.SetBulge 0, 0.424127
plineObj.SetBulge 1, 0.161605
plineObj.SetBulge 2, 0.83194
plineObj.SetBulge 3, -0.267392
plineObj.SetBulge 4, -0.106986
plineObj.SetBulge 5, -0.0601812
plineObj.SetBulge 6, -0.0862239
plineObj.SetBulge 7, -0.0871434
plineObj.SetBulge 8, -0.0408194
plineObj.SetBulge 9, 0.0867111
plineObj.SetBulge 10, 0.208082
plineObj.SetBulge 11, 0.154888
plineObj.SetBulge 12, 0.189184

plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True
RetVal = plineObj.Mirror(b1, b2)
RetVal = plineObj.Mirror(a1, a2)
RetVal = plineObj.Copy
  ' Define the rotation of 45 degrees about a
  ' base point of (4, 4.25, 0)
  basePoint(0) = points(4) - (b / 2): basePoint(1) = points(3) - (a / 2)
  rotationAngle = 3.14159   ' 45 degrees
  plineObj.Rotate basePoint, rotationAngle
  plineObj.Update
  
pointscntr4(0) = bcp + 6.97983:      pointscntr4(1) = acp + 66.4733
pointscntr4(2) = bcp + 4.40328:      pointscntr4(3) = acp + 70.4044
pointscntr4(4) = bcp + 2.76705:      pointscntr4(5) = acp + 73.302
pointscntr4(6) = bcp + 2.03069:      pointscntr4(7) = acp + 75.4075
pointscntr4(8) = bcp + 1.13687E-13:  pointscntr4(9) = acp + 82.896
pointscntr4(10) = bcp + -2.03069:    pointscntr4(11) = acp + 75.4075
pointscntr4(12) = bcp + -2.76705:    pointscntr4(13) = acp + 73.302
pointscntr4(14) = bcp + -4.35149:    pointscntr4(15) = acp + 70.4834
pointscntr4(16) = bcp + -6.97983:    pointscntr4(17) = acp + 66.4733
pointscntr4(18) = bcp + -8.62407:    pointscntr4(19) = acp + 63.7027
pointscntr4(20) = bcp + -9.8543:     pointscntr4(21) = acp + 57.1808
pointscntr4(22) = bcp + -8.14995:    pointscntr4(23) = acp + 53.3016
pointscntr4(24) = bcp + -5.61959:    pointscntr4(25) = acp + 51.6618
pointscntr4(26) = bcp + -2.27229:    pointscntr4(27) = acp + 51.8809
pointscntr4(28) = bcp + 1.13687E-13: pointscntr4(29) = acp + 55.0398
pointscntr4(30) = bcp + 2.27229:     pointscntr4(31) = acp + 51.8809
pointscntr4(32) = bcp + 5.61959:     pointscntr4(33) = acp + 51.6618
pointscntr4(34) = bcp + 8.14995:     pointscntr4(35) = acp + 53.3016
pointscntr4(36) = bcp + 9.8543:      pointscntr4(37) = acp + 57.1808
pointscntr4(38) = bcp + 8.62407:     pointscntr4(39) = acp + 63.7027

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr4)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update

plineObj.SetBulge 0, 0
plineObj.SetBulge 1, -0.0340478
plineObj.SetBulge 2, -0.0404214
plineObj.SetBulge 3, 0
plineObj.SetBulge 4, 0
plineObj.SetBulge 5, -0.0404214
plineObj.SetBulge 6, -0.0330812
plineObj.SetBulge 7, 0
plineObj.SetBulge 8, 0.021196
plineObj.SetBulge 9, 0.173455
plineObj.SetBulge 10, 0.145317
plineObj.SetBulge 11, 0.143574
plineObj.SetBulge 12, 0.203485
plineObj.SetBulge 13, 0.200979
plineObj.SetBulge 14, 0.200979
plineObj.SetBulge 15, 0.203485
plineObj.SetBulge 16, 0.143574
plineObj.SetBulge 17, 0.145317
plineObj.SetBulge 18, 0.173455
plineObj.SetBulge 19, 0.021196

plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True
RetVal = plineObj.Mirror(b1, b2)

pointscntr5(0) = bcp + 8.66267:     pointscntr5(1) = acp + 94.5619
pointscntr5(2) = bcp + 12.0028:     pointscntr5(3) = acp + 94.7543
pointscntr5(4) = bcp + 14.9353:     pointscntr5(5) = acp + 93.6436
pointscntr5(6) = bcp + 16.3813:     pointscntr5(7) = acp + 91.4456
pointscntr5(8) = bcp + 16.3602:     pointscntr5(9) = acp + 89.5013
pointscntr5(10) = bcp + 14.3893:    pointscntr5(11) = acp + 86.8541
pointscntr5(12) = bcp + 11.3686:    pointscntr5(13) = acp + 86.1899
pointscntr5(14) = bcp + 13.0885:    pointscntr5(15) = acp + 87.6125
pointscntr5(16) = bcp + 12.9123:    pointscntr5(17) = acp + 89.239
pointscntr5(18) = bcp + 10.7834:    pointscntr5(19) = acp + 90.4247
pointscntr5(20) = bcp + 5.27888:    pointscntr5(21) = acp + 89.7626
pointscntr5(22) = bcp + 2.39309:    pointscntr5(23) = acp + 86.7891
pointscntr5(24) = bcp + 1.49609:    pointscntr5(25) = acp + 88.2863
pointscntr5(26) = bcp + 4.32014:    pointscntr5(27) = acp + 92.3133

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr5)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update
  
plineObj.SetBulge 0, -0.0883845
plineObj.SetBulge 1, -0.121995
plineObj.SetBulge 2, -0.193108
plineObj.SetBulge 3, -0.0993998
plineObj.SetBulge 4, -0.23388
plineObj.SetBulge 5, -0.125673
plineObj.SetBulge 6, 0.11811
plineObj.SetBulge 7, 0.412476
plineObj.SetBulge 8, 0.0962103
plineObj.SetBulge 9, 0.224454
plineObj.SetBulge 10, 0.11374
plineObj.SetBulge 11, 0
plineObj.SetBulge 12, -0.11562
plineObj.SetBulge 13, -0.128851

plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True
RetVal = plineObj.Mirror(b1, b2)
RetVal = plineObj.Mirror(a1, a2)
RetVal = plineObj.Copy
  ' Define the rotation of 45 degrees about a
  ' base point of (4, 4.25, 0)
  basePoint(0) = points(4) - (b / 2): basePoint(1) = points(3) - (a / 2)
  rotationAngle = 3.14159   ' 45 degrees
  plineObj.Rotate basePoint, rotationAngle
  plineObj.Update

pointscntr6(0) = bcp + 3.75304:       pointscntr6(1) = acp + 101.938
pointscntr6(2) = bcp + 1.13687E-13:   pointscntr6(3) = acp + 111.5
pointscntr6(4) = bcp + -3.75304:      pointscntr6(5) = acp + 101.938
pointscntr6(6) = bcp + -3.37708:      pointscntr6(7) = acp + 96.8397
pointscntr6(8) = bcp + 1.13687E-13:   pointscntr6(9) = acp + 88.5857
pointscntr6(10) = bcp + 3.37708:      pointscntr6(11) = acp + 96.8397

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr6)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update
  
plineObj.SetBulge 0, 0.06872
plineObj.SetBulge 1, 0.06872
plineObj.SetBulge 2, 0.171056
plineObj.SetBulge 3, 0
plineObj.SetBulge 4, 0
plineObj.SetBulge 5, 0.171056

    plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True
RetVal = plineObj.Mirror(b1, b2)

center(0) = a1(0): center(1) = b1(1) + 45.76: center(2) = 0: radius = 2.9613
Set circleObj = ThisDrawing.ModelSpace.AddCircle(center, radius)
circleObj.Layer = "K-grav_Pattern"
circleObj.Update
center(0) = a1(0): center(1) = b1(1) - 45.76: center(2) = 0: radius = 2.9613
Set circleObj = ThisDrawing.ModelSpace.AddCircle(center, radius)
circleObj.Layer = "K-grav_Pattern"
circleObj.Update

End If
End If
  
  
  
I = 100
  
 Dim plineObj2 As AcadLWPolyline
  Dim points2(0 To 7) As Double
  
  points2(0) = points(0):    points2(1) = a + I:
  points2(2) = points(0):    points2(3) = (a * 2) + I:
  points2(4) = points(6):    points2(5) = (a * 2) + I:
  points2(6) = points(6):    points2(7) = a + I:
  
'========================================================
'========================================================
'========================================================
  
  
 For d = 2 To Cells(c, 4)
If d > 1 Then

Set plineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(points2)
  plineObj2.Closed = True
  
If a > 280 Then
If b > 280 Then
  
  pointswithin(0) = points2(0) + 100:                             pointswithin(1) = points2(1) + 100
  pointswithin(2) = points2(0) + 85:                              pointswithin(3) = points2(1) + 100 + ((a - 200) / 4)
  pointswithin(4) = points2(0) + 70:                              pointswithin(5) = points2(1) + 100 + (2 * ((a - 200) / 4))
  pointswithin(6) = points2(0) + 85:                              pointswithin(7) = points2(1) + 100 + (3 * ((a - 200) / 4))
  pointswithin(8) = points2(0) + 100:                             pointswithin(9) = points2(3) - 100
  pointswithin(10) = points2(0) + 100 + ((b - 200) / 4):          pointswithin(11) = points2(3) - 85
  pointswithin(12) = points2(0) + 100 + (2 * ((b - 200) / 4)):    pointswithin(13) = points2(3) - 70
  pointswithin(14) = points2(0) + 100 + (3 * ((b - 200) / 4)):    pointswithin(15) = points2(3) - 85
  pointswithin(16) = points2(4) - 100:                            pointswithin(17) = points2(3) - 100
  pointswithin(18) = points2(4) - 85:                             pointswithin(19) = points2(3) - 100 - ((a - 200) / 4)
  pointswithin(20) = points2(4) - 70:                             pointswithin(21) = points2(3) - 100 - (2 * ((a - 200) / 4))
  pointswithin(22) = points2(4) - 85:                             pointswithin(23) = points2(3) - 100 - (3 * ((a - 200) / 4))
  pointswithin(24) = points2(4) - 100:                            pointswithin(25) = points2(1) + 100
  pointswithin(26) = points2(4) - 100 - ((b - 200) / 4):          pointswithin(27) = points2(1) + 85
  pointswithin(28) = points2(4) - 100 - (2 * ((b - 200) / 4)):    pointswithin(29) = points2(1) + 70
  pointswithin(30) = points2(4) - 100 - (3 * ((b - 200) / 4)):    pointswithin(31) = points2(1) + 85
  

If a > 400 Then
  

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointswithin)
  plineObj.Closed = True
   
   pb = (b - 200) / 4
   pa = (a - 200) / 4
ga = Sqr((pa * pa) + 225)
gb = Sqr((pb * pb) + 225)
   anglea = Atn((15 / ga) / Sqr((-15 / ga) * (15 / ga) + 1))
   angleb = Atn((15 / gb) / Sqr((-15 / gb) * (15 / gb) + 1))
   radiusa = (ga / 2) / Sin(15 / ga)
   radiusb = (gb / 2) / Sin(15 / gb)
   ha = radiusa * (1 - Cos(anglea))
   hb = radiusb * (1 - Cos(angleb))
   ka = ha / ga
   kb = hb / gb
   
    plineObj.SetBulge 0, ka * 2
    plineObj.SetBulge 1, -ka * 2
    plineObj.SetBulge 2, -ka * 2
    plineObj.SetBulge 3, ka * 2
    plineObj.SetBulge 4, kb * 2
    plineObj.SetBulge 5, -kb * 2
    plineObj.SetBulge 6, -kb * 2
    plineObj.SetBulge 7, kb * 2
    plineObj.SetBulge 8, ka * 2
    plineObj.SetBulge 9, -ka * 2
    plineObj.SetBulge 10, -ka * 2
    plineObj.SetBulge 11, ka * 2
    plineObj.SetBulge 12, kb * 2
    plineObj.SetBulge 13, -kb * 2
    plineObj.SetBulge 14, -kb * 2
    plineObj.SetBulge 15, kb * 2
    plineObj.Layer = "C-Mill"
    plineObj.Update
    plineObj.Closed = True

End If

  pointshelpa1(0) = points2(0) + 79:                pointshelpa1(1) = points2(1) + 100
  pointshelpa1(2) = points2(0) + 49:                pointshelpa1(3) = points2(1) + (a / 2)
  pointshelpa2(0) = points2(0) + 100 - radiusa:     pointshelpa2(1) = points2(1) + 100
  pointshelpa2(2) = pointswithin(2):               pointshelpa2(3) = pointswithin(3)
Set plineObjw1 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointshelpa1)
Set plineObjw2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointshelpa2)
intPointsa = plineObjw1.IntersectWith(plineObjw2, acExtendBoth)
Intersectionax = intPointsa(0) - points2(0)
Intersectionay = intPointsa(1) - points2(1)
plineObjw1.Delete
plineObjw2.Delete

  pointshelpb1(0) = points2(0) + 100:     pointshelpb1(1) = points2(3) - 79
  pointshelpb1(2) = points2(0) + (b / 2): pointshelpb1(3) = points2(3) - 49
  pointshelpb2(0) = points2(0) + 100:     pointshelpb2(1) = points2(3) - 100 + radiusb
  pointshelpb2(2) = pointswithin(10):    pointshelpb2(3) = pointswithin(11)
Set plineObjw1 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointshelpb1)
Set plineObjw2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointshelpb2)
intPointsb = plineObjw1.IntersectWith(plineObjw2, acExtendBoth)
Intersectionbx = intPointsb(0) - points2(0)
Intersectionby = points2(3) - intPointsb(1)
plineObjw1.Delete
plineObjw2.Delete
  
  pointswithin2(0) = points2(0) + 79:                              pointswithin2(1) = points2(1) + 79
  pointswithin2(2) = points2(0) + 79:                              pointswithin2(3) = points2(1) + 100
  pointswithin2(4) = points2(0) + Intersectionax:                  pointswithin2(5) = points2(1) + Intersectionay
  pointswithin2(6) = points2(0) + 49:                              pointswithin2(7) = points2(1) + 100 + (2 * ((a - 200) / 4))
  pointswithin2(8) = points2(0) + Intersectionax:                  pointswithin2(9) = points2(3) - Intersectionay
  pointswithin2(10) = points2(0) + 79:                             pointswithin2(11) = points2(3) - 100
  pointswithin2(12) = points2(0) + 79:                             pointswithin2(13) = points2(3) - 79
  pointswithin2(14) = points2(0) + 100:                            pointswithin2(15) = points2(3) - 79
  pointswithin2(16) = points2(0) + Intersectionbx:                 pointswithin2(17) = points2(3) - Intersectionby
  pointswithin2(18) = points2(4) - 100 - (2 * ((b - 200) / 4)):    pointswithin2(19) = points2(3) - 49
  pointswithin2(20) = points2(4) - Intersectionbx:                 pointswithin2(21) = points2(3) - Intersectionby
  pointswithin2(22) = points2(4) - 100:                            pointswithin2(23) = points2(3) - 79
  pointswithin2(24) = points2(4) - 79:                             pointswithin2(25) = points2(3) - 79
  pointswithin2(26) = points2(4) - 79:                             pointswithin2(27) = points2(3) - 100
  pointswithin2(28) = points2(4) - Intersectionax:                 pointswithin2(29) = points2(3) - Intersectionay
  pointswithin2(30) = points2(4) - 49:                             pointswithin2(31) = points2(1) + 100 + (2 * ((a - 200) / 4))
  pointswithin2(32) = points2(4) - Intersectionax:                 pointswithin2(33) = points2(1) + Intersectionay
  pointswithin2(34) = points2(4) - 79:                             pointswithin2(35) = points2(1) + 100
  pointswithin2(36) = points2(4) - 79:                             pointswithin2(37) = points2(1) + 79
  pointswithin2(38) = points2(4) - 100:                            pointswithin2(39) = points2(1) + 79
  pointswithin2(40) = points2(4) - Intersectionbx:                 pointswithin2(41) = points2(1) + Intersectionby
  pointswithin2(42) = points2(4) - 100 - (2 * ((b - 200) / 4)):    pointswithin2(43) = points2(1) + 49
  pointswithin2(44) = points2(0) + Intersectionbx:                 pointswithin2(45) = points2(1) + Intersectionby
  pointswithin2(46) = points2(0) + 100:                            pointswithin2(47) = points2(1) + 79

If a > 400 Then

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointswithin2)
  plineObj.Closed = True
  
    plineObj.SetBulge 1, ka * 2
    plineObj.SetBulge 2, -ka * 2
    plineObj.SetBulge 3, -ka * 2
    plineObj.SetBulge 4, ka * 2
    plineObj.SetBulge 7, kb * 2
    plineObj.SetBulge 8, -kb * 2
    plineObj.SetBulge 9, -kb * 2
    plineObj.SetBulge 10, kb * 2
    plineObj.SetBulge 13, ka * 2
    plineObj.SetBulge 14, -ka * 2
    plineObj.SetBulge 15, -ka * 2
    plineObj.SetBulge 16, ka * 2
    plineObj.SetBulge 19, kb * 2
    plineObj.SetBulge 20, -kb * 2
    plineObj.SetBulge 21, -kb * 2
    plineObj.SetBulge 22, kb * 2
    plineObj.Layer = "K-grav"
    plineObj.Update
    plineObj.Closed = True
  
End If


pointsaktr1(0) = points2(0) + 41.3216:   pointsaktr1(1) = points2(1) + 44.9698
pointsaktr1(2) = points2(0) + 40.3478:   pointsaktr1(3) = points2(1) + 43.1003
pointsaktr1(4) = points2(0) + 46.1268:   pointsaktr1(5) = points2(1) + 38.9822
pointsaktr1(6) = points2(0) + 56.3386:   pointsaktr1(7) = points2(1) + 50.9313
pointsaktr1(8) = points2(0) + 44.8855:   pointsaktr1(9) = points2(1) + 64.5693
pointsaktr1(10) = points2(0) + 33.7952:  pointsaktr1(11) = points2(1) + 64.5305
pointsaktr1(12) = points2(0) + 24.0386:  pointsaktr1(13) = points2(1) + 29.3573
pointsaktr1(14) = points2(0) + 64.0284:  pointsaktr1(15) = points2(1) + 21.9177
pointsaktr1(16) = points2(0) + 127.102:  pointsaktr1(17) = points2(1) + 21.8129
pointsaktr1(18) = points2(0) + 127.198:  pointsaktr1(19) = points2(1) + 21.6297
pointsaktr1(20) = points2(0) + 127.198:  pointsaktr1(21) = points2(1) + 21.1986
pointsaktr1(22) = points2(0) + 127.382:  pointsaktr1(23) = points2(1) + 20.9805
pointsaktr1(24) = points2(0) + 128.388:  pointsaktr1(25) = points2(1) + 20.3488
pointsaktr1(26) = points2(0) + 130.165:  pointsaktr1(27) = points2(1) + 19.536
pointsaktr1(28) = points2(0) + 130.323:  pointsaktr1(29) = points2(1) + 19.4341
pointsaktr1(30) = points2(0) + 130.328:  pointsaktr1(31) = points2(1) + 19.3365
pointsaktr1(32) = points2(0) + 130.219:  pointsaktr1(33) = points2(1) + 19.3111
pointsaktr1(34) = points2(0) + 70.0257:  pointsaktr1(35) = points2(1) + 19.4725
pointsaktr1(36) = points2(0) + 29.517:   pointsaktr1(37) = points2(1) + 20.6803
pointsaktr1(38) = points2(0) + 17.9889:  pointsaktr1(39) = points2(1) + 53.5889
pointsaktr1(40) = points2(0) + 30.043:   pointsaktr1(41) = points2(1) + 69.0807
pointsaktr1(42) = points2(0) + 33.0582:  pointsaktr1(43) = points2(1) + 72.7887
pointsaktr1(44) = points2(0) + 33.8149:  pointsaktr1(45) = points2(1) + 77.9079
pointsaktr1(46) = points2(0) + 31.4709:  pointsaktr1(47) = points2(1) + 79.0833
pointsaktr1(48) = points2(0) + 29.5231:  pointsaktr1(49) = points2(1) + 77.7696
pointsaktr1(50) = points2(0) + 29.3443:  pointsaktr1(51) = points2(1) + 77.788
pointsaktr1(52) = points2(0) + 29.2837:  pointsaktr1(53) = points2(1) + 78.7799
pointsaktr1(54) = points2(0) + 30.6341:  pointsaktr1(55) = points2(1) + 81.5798
pointsaktr1(56) = points2(0) + 29.5258:  pointsaktr1(57) = points2(1) + 87.1914
pointsaktr1(58) = points2(0) + 24.6502:  pointsaktr1(59) = points2(1) + 87.5247
pointsaktr1(60) = points2(0) + 23.6658:  pointsaktr1(61) = points2(1) + 83.7225
pointsaktr1(62) = points2(0) + 25.7501:  pointsaktr1(63) = points2(1) + 79.0141
pointsaktr1(64) = points2(0) + 25.8284:  pointsaktr1(65) = points2(1) + 78.4251
pointsaktr1(66) = points2(0) + 25.6434:  pointsaktr1(67) = points2(1) + 78.357
pointsaktr1(68) = points2(0) + 22.852:   pointsaktr1(69) = points2(1) + 79.0068
pointsaktr1(70) = points2(0) + 20.5279:  pointsaktr1(71) = points2(1) + 78.2691
pointsaktr1(72) = points2(0) + 18.1972:  pointsaktr1(73) = points2(1) + 77.4484
pointsaktr1(74) = points2(0) + 16.347:   pointsaktr1(75) = points2(1) + 77.9443
pointsaktr1(76) = points2(0) + 15.3674:  pointsaktr1(77) = points2(1) + 80.3287
pointsaktr1(78) = points2(0) + 15.2094:  pointsaktr1(79) = points2(1) + 89.2364
pointsaktr1(80) = points2(0) + 15.7178:  pointsaktr1(81) = points2(1) + 91.0514
pointsaktr1(82) = points2(0) + 17.9348:  pointsaktr1(83) = points2(1) + 92.3071
pointsaktr1(84) = points2(0) + 19.838:   pointsaktr1(85) = points2(1) + 91.9659
pointsaktr1(86) = points2(0) + 21.9268:  pointsaktr1(87) = points2(1) + 93.1289
pointsaktr1(88) = points2(0) + 21.8411:  pointsaktr1(89) = points2(1) + 93.7497
pointsaktr1(90) = points2(0) + 17.2959:  pointsaktr1(91) = points2(1) + 102.917
pointsaktr1(92) = points2(0) + 15.5268:  pointsaktr1(93) = points2(1) + 113.7
pointsaktr1(94) = points2(0) + 20.041:   pointsaktr1(95) = points2(1) + 120.229
pointsaktr1(96) = points2(0) + 27.4677:  pointsaktr1(97) = points2(1) + 119.229
pointsaktr1(98) = points2(0) + 29.1061:  pointsaktr1(99) = points2(1) + 114.516
pointsaktr1(100) = points2(0) + 29.0205: pointsaktr1(101) = points2(1) + 114.017
pointsaktr1(102) = points2(0) + 24.0139: pointsaktr1(103) = points2(1) + 112.877
pointsaktr1(104) = points2(0) + 23.702:  pointsaktr1(105) = points2(1) + 114.536
pointsaktr1(106) = points2(0) + 23.5509: pointsaktr1(107) = points2(1) + 116.32
pointsaktr1(108) = points2(0) + 20.7646: pointsaktr1(109) = points2(1) + 117.321
pointsaktr1(110) = points2(0) + 17.4574: pointsaktr1(111) = points2(1) + 113.936
pointsaktr1(112) = points2(0) + 20.0439: pointsaktr1(113) = points2(1) + 103.759
pointsaktr1(114) = points2(0) + 28.9597: pointsaktr1(115) = points2(1) + 90.4939
pointsaktr1(116) = points2(0) + 40.2373: pointsaktr1(117) = points2(1) + 78.0111
pointsaktr1(118) = points2(0) + 42.7351: pointsaktr1(119) = points2(1) + 73.81
pointsaktr1(120) = points2(0) + 42.6941: pointsaktr1(121) = points2(1) + 71.0615
pointsaktr1(122) = points2(0) + 40.3146: pointsaktr1(123) = points2(1) + 69.1187
pointsaktr1(124) = points2(0) + 40.3297: pointsaktr1(125) = points2(1) + 68.899
pointsaktr1(126) = points2(0) + 55.6191: pointsaktr1(127) = points2(1) + 56.4976
pointsaktr1(128) = points2(0) + 58.2067: pointsaktr1(129) = points2(1) + 47.14
pointsaktr1(130) = points2(0) + 58.8506: pointsaktr1(131) = points2(1) + 39.5231
pointsaktr1(132) = points2(0) + 67.8108: pointsaktr1(133) = points2(1) + 28.2563
pointsaktr1(134) = points2(0) + 79.5835: pointsaktr1(135) = points2(1) + 26.1017
pointsaktr1(136) = points2(0) + 84.365:  pointsaktr1(137) = points2(1) + 25.9125
pointsaktr1(138) = points2(0) + 84.8383: pointsaktr1(139) = points2(1) + 25.8106
pointsaktr1(140) = points2(0) + 84.9932: pointsaktr1(141) = points2(1) + 25.4054
pointsaktr1(142) = points2(0) + 83.696:  pointsaktr1(143) = points2(1) + 24.3908
pointsaktr1(144) = points2(0) + 80.3696: pointsaktr1(145) = points2(1) + 22.5492
pointsaktr1(146) = points2(0) + 80.2486: pointsaktr1(147) = points2(1) + 22.5346
pointsaktr1(148) = points2(0) + 76.4568: pointsaktr1(149) = points2(1) + 22.5275
pointsaktr1(150) = points2(0) + 76.3508: pointsaktr1(151) = points2(1) + 22.6946
pointsaktr1(152) = points2(0) + 75.9307: pointsaktr1(153) = points2(1) + 23.9924
pointsaktr1(154) = points2(0) + 73.6332: pointsaktr1(155) = points2(1) + 24.5453
pointsaktr1(156) = points2(0) + 68.436:  pointsaktr1(157) = points2(1) + 24.592
pointsaktr1(158) = points2(0) + 63.6462: pointsaktr1(159) = points2(1) + 24.8968
pointsaktr1(160) = points2(0) + 58.4456: pointsaktr1(161) = points2(1) + 28.7795
pointsaktr1(162) = points2(0) + 58.1362: pointsaktr1(163) = points2(1) + 29.3731
pointsaktr1(164) = points2(0) + 58.0776: pointsaktr1(165) = points2(1) + 29.4447
pointsaktr1(166) = points2(0) + 57.4768: pointsaktr1(167) = points2(1) + 29.4281
pointsaktr1(168) = points2(0) + 57.0179: pointsaktr1(169) = points2(1) + 28.8684
pointsaktr1(170) = points2(0) + 56.6262: pointsaktr1(171) = points2(1) + 28.0932
pointsaktr1(172) = points2(0) + 56.4903: pointsaktr1(173) = points2(1) + 27.8301
pointsaktr1(174) = points2(0) + 55.7942: pointsaktr1(175) = points2(1) + 27.1122
pointsaktr1(176) = points2(0) + 55.2031: pointsaktr1(177) = points2(1) + 27.0389
pointsaktr1(178) = points2(0) + 55.0785: pointsaktr1(179) = points2(1) + 27.1134
pointsaktr1(180) = points2(0) + 53.4121: pointsaktr1(181) = points2(1) + 33.4249
pointsaktr1(182) = points2(0) + 54.6571: pointsaktr1(183) = points2(1) + 38.7189
pointsaktr1(184) = points2(0) + 54.4842: pointsaktr1(185) = points2(1) + 38.8578
pointsaktr1(186) = points2(0) + 50.9573: pointsaktr1(187) = points2(1) + 37.2303
pointsaktr1(188) = points2(0) + 39.1485: pointsaktr1(189) = points2(1) + 40.5163
pointsaktr1(190) = points2(0) + 36.3064: pointsaktr1(191) = points2(1) + 45.3582
pointsaktr1(192) = points2(0) + 36.6083: pointsaktr1(193) = points2(1) + 47.1011
pointsaktr1(194) = points2(0) + 37.8654: pointsaktr1(195) = points2(1) + 48.2559
pointsaktr1(196) = points2(0) + 44.2299: pointsaktr1(197) = points2(1) + 54.6583
pointsaktr1(198) = points2(0) + 48.0345: pointsaktr1(199) = points2(1) + 62.0273
pointsaktr1(200) = points2(0) + 48.2181: pointsaktr1(201) = points2(1) + 62.0909
pointsaktr1(202) = points2(0) + 54.9901: pointsaktr1(203) = points2(1) + 53.5569
pointsaktr1(204) = points2(0) + 54.9939: pointsaktr1(205) = points2(1) + 53.4679
pointsaktr1(206) = points2(0) + 50.0612: pointsaktr1(207) = points2(1) + 46.176
pointsaktr1(208) = points2(0) + 47.4251: pointsaktr1(209) = points2(1) + 45.3091


Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsaktr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update
    
    ' Find the bulge of the third segment
    currentBulge = plineObj.GetBulge(2)
    ' Change the bulge of the third segment

plineObj.SetBulge 0, 0.409159
plineObj.SetBulge 1, 0.321957
plineObj.SetBulge 2, 0.468929
plineObj.SetBulge 3, 0.242406
plineObj.SetBulge 4, 0.181208
plineObj.SetBulge 5, 0.481392
plineObj.SetBulge 6, 0.380589
plineObj.SetBulge 7, 0
plineObj.SetBulge 8, -0.607953
plineObj.SetBulge 9, 0.32111
plineObj.SetBulge 10, 0.0696001
plineObj.SetBulge 11, 0.063976
plineObj.SetBulge 12, 0
plineObj.SetBulge 13, 0
plineObj.SetBulge 14, -0.502036
plineObj.SetBulge 15, -0.126379
plineObj.SetBulge 16, 0
plineObj.SetBulge 17, -0.246237
plineObj.SetBulge 18, -0.379673
plineObj.SetBulge 19, -0.102627
plineObj.SetBulge 20, 0.0894177
plineObj.SetBulge 21, 0.192392
plineObj.SetBulge 22, 0.46922
plineObj.SetBulge 23, 0.104315
plineObj.SetBulge 24, -0.437919
plineObj.SetBulge 25, -0.274694
plineObj.SetBulge 26, 0.0834997
plineObj.SetBulge 27, 0.259096
plineObj.SetBulge 28, 0.460234
plineObj.SetBulge 29, 0.247169
plineObj.SetBulge 30, 0.102358
plineObj.SetBulge 31, -0.197867
plineObj.SetBulge 32, -0.553072
plineObj.SetBulge 33, 0.213683
plineObj.SetBulge 34, 0.0574858
plineObj.SetBulge 35, -0.0523247
plineObj.SetBulge 36, -0.26884
plineObj.SetBulge 37, -0.136043
plineObj.SetBulge 38, -0.0629278
plineObj.SetBulge 39, -0.0675464
plineObj.SetBulge 40, -0.359276
plineObj.SetBulge 41, -0.00859763
plineObj.SetBulge 42, 0.355483
plineObj.SetBulge 43, 0.203986
plineObj.SetBulge 44, -0.039733
plineObj.SetBulge 45, -0.124994
plineObj.SetBulge 46, -0.254634
plineObj.SetBulge 47, -0.332616
plineObj.SetBulge 48, -0.179762
plineObj.SetBulge 49, -0.0847805
plineObj.SetBulge 50, -0.521548
plineObj.SetBulge 51, -0.278007
plineObj.SetBulge 52, 0.277679
plineObj.SetBulge 53, 0.264017
plineObj.SetBulge 54, 0.35711
plineObj.SetBulge 55, 0.152853
plineObj.SetBulge 56, 0.0530635
plineObj.SetBulge 57, 0
plineObj.SetBulge 58, -0.0707038
plineObj.SetBulge 59, -0.212213
plineObj.SetBulge 60, -0.171738
plineObj.SetBulge 61, 0.715161
plineObj.SetBulge 62, -0.217317
plineObj.SetBulge 63, -0.104569
plineObj.SetBulge 64, 0.0278538
plineObj.SetBulge 65, 0.303024
plineObj.SetBulge 66, 0.0691689
plineObj.SetBulge 67, 0
plineObj.SetBulge 68, -0.0324659
plineObj.SetBulge 69, -0.563877
plineObj.SetBulge 70, -0.090811
plineObj.SetBulge 71, 0
plineObj.SetBulge 72, -0.0878322
plineObj.SetBulge 73, 0
plineObj.SetBulge 74, -0.554362
plineObj.SetBulge 75, 0.465455
plineObj.SetBulge 76, 0.0895132
plineObj.SetBulge 77, 0.0326569
plineObj.SetBulge 78, -0.0784946
plineObj.SetBulge 79, -0.218538
plineObj.SetBulge 80, 0
plineObj.SetBulge 81, 0.0921223
plineObj.SetBulge 82, 0.256818
plineObj.SetBulge 83, 0.13945
plineObj.SetBulge 84, 0
plineObj.SetBulge 85, 0
plineObj.SetBulge 86, -0.160405
plineObj.SetBulge 87, -0.146033
plineObj.SetBulge 88, -0.191927
plineObj.SetBulge 89, -0.185968
plineObj.SetBulge 90, -0.062917
plineObj.SetBulge 91, 0.72199
plineObj.SetBulge 92, -0.0715105
plineObj.SetBulge 93, -0.271829
plineObj.SetBulge 94, -0.141216
plineObj.SetBulge 95, -0.209519
plineObj.SetBulge 96, -0.0278354
plineObj.SetBulge 97, 0.0505341
plineObj.SetBulge 98, 0.100553
plineObj.SetBulge 99, -0.53094
plineObj.SetBulge 100, -0.125904
plineObj.SetBulge 101, -0.0996162
plineObj.SetBulge 102, -0.148082
plineObj.SetBulge 103, -0.203841
plineObj.SetBulge 104, 0.135952
    
    plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True
  
  
  
  b1(0) = points2(4) - (b / 2) - 1: b1(1) = points2(1) + (a / 2)
  b2(0) = points2(4) - (b / 2) + 1:  b2(1) = points2(1) + (a / 2)
  RetVal = plineObj.Mirror(b1, b2)

 
  a1(0) = points2(4) - (b / 2): a1(1) = points2(1) + (a / 2) - 1
  a2(0) = points2(4) - (b / 2): a2(1) = points2(1) + (a / 2) + 1
  RetVal = plineObj.Mirror(a1, a2)
  
  RetVal = plineObj.Copy
  ' Define the rotation of 45 degrees about a
  ' base point of (4, 4.25, 0)
  
  basePoint(0) = points2(4) - (b / 2): basePoint(1) = points2(3) - (a / 2)
  rotationAngle = 3.14159   ' 45 degrees

  ' Rotate the polyline
  plineObj.Rotate basePoint, rotationAngle
  plineObj.Update

pointsaktr2(0) = points2(0) + 36.8503:   pointsaktr2(1) = points2(1) + 86.3405
pointsaktr2(2) = points2(0) + 36.9528:   pointsaktr2(3) = points2(1) + 86.5422
pointsaktr2(4) = points2(0) + 46.0287:   pointsaktr2(5) = points2(1) + 80.7521
pointsaktr2(6) = points2(0) + 51.7303:   pointsaktr2(7) = points2(1) + 74.7275
pointsaktr2(8) = points2(0) + 56.0192:   pointsaktr2(9) = points2(1) + 57.4278
pointsaktr2(10) = points2(0) + 55.8011:  pointsaktr2(11) = points2(1) + 57.3905
pointsaktr2(12) = points2(0) + 48.6202:  pointsaktr2(13) = points2(1) + 65.651
pointsaktr2(14) = points2(0) + 48.5738:  pointsaktr2(15) = points2(1) + 65.7471
pointsaktr2(16) = points2(0) + 47.7344:  pointsaktr2(17) = points2(1) + 71.4489

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsaktr2)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update
  
plineObj.SetBulge 0, -0.771138
plineObj.SetBulge 1, -0.0611678
plineObj.SetBulge 2, -0.0629979
plineObj.SetBulge 3, -0.201444
plineObj.SetBulge 4, -0.714409
plineObj.SetBulge 5, 0.106484
plineObj.SetBulge 6, -0.250546
plineObj.SetBulge 7, 0.0841751
plineObj.SetBulge 8, 0.130058

 plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True
  
  a1(0) = points2(4) - (b / 2): a1(1) = points2(1) + (a / 2) - 1
  a2(0) = points2(4) - (b / 2): a2(1) = points2(1) + (a / 2) + 1
  RetVal = plineObj.Mirror(b1, b2)
  a1(0) = points2(4) - (b / 2): a1(1) = points2(1) + (a / 2) - 1
  a2(0) = points2(4) - (b / 2): a2(1) = points2(1) + (a / 2) + 1
  RetVal = plineObj.Mirror(a1, a2)
  RetVal = plineObj.Copy
  ' Define the rotation of 45 degrees about a
  ' base point of (4, 4.25, 0)
  basePoint(0) = points2(4) - (b / 2): basePoint(1) = points2(3) - (a / 2)
  rotationAngle = 3.14159   ' 45 degrees

  ' Rotate the polyline
  plineObj.Rotate basePoint, rotationAngle
  plineObj.Update
  
pointsaktr3(0) = points2(0) + 135.147:   pointsaktr3(1) = points2(1) + 17.7302
pointsaktr3(2) = points2(0) + 133.843:   pointsaktr3(3) = points2(1) + 17.5545
pointsaktr3(4) = points2(0) + 131.609:   pointsaktr3(5) = points2(1) + 17.4606
pointsaktr3(6) = points2(0) + 81.121:    pointsaktr3(7) = points2(1) + 17.4146
pointsaktr3(8) = points2(0) + 76.3148:   pointsaktr3(9) = points2(1) + 17.9255
pointsaktr3(10) = points2(0) + 75.7965:  pointsaktr3(11) = points2(1) + 18.1653
pointsaktr3(12) = points2(0) + 75.4241:  pointsaktr3(13) = points2(1) + 18.7953
pointsaktr3(14) = points2(0) + 75.3105:  pointsaktr3(15) = points2(1) + 18.8943
pointsaktr3(16) = points2(0) + 72.3454:  pointsaktr3(17) = points2(1) + 18.8905
pointsaktr3(18) = points2(0) + 72.2831:  pointsaktr3(19) = points2(1) + 18.8724
pointsaktr3(20) = points2(0) + 69.6728:  pointsaktr3(21) = points2(1) + 16.5045
pointsaktr3(22) = points2(0) + 70.0556:  pointsaktr3(23) = points2(1) + 15.3657
pointsaktr3(24) = points2(0) + 72.3199:  pointsaktr3(25) = points2(1) + 15.0109
pointsaktr3(26) = points2(0) + 137.437:  pointsaktr3(27) = points2(1) + 15.0198
pointsaktr3(28) = points2(0) + 138.883:  pointsaktr3(29) = points2(1) + 15.2189
pointsaktr3(30) = points2(0) + 138.929:  pointsaktr3(31) = points2(1) + 15.4766
pointsaktr3(32) = points2(0) + 138.302:  pointsaktr3(33) = points2(1) + 15.8567
pointsaktr3(34) = points2(0) + 136.44:   pointsaktr3(35) = points2(1) + 16.6981
pointsaktr3(36) = points2(0) + 135.583:  pointsaktr3(37) = points2(1) + 17.2219
pointsaktr3(38) = points2(0) + 135.283:  pointsaktr3(39) = points2(1) + 17.6465

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsaktr3)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update
  
plineObj.SetBulge 0, -0.0317584
plineObj.SetBulge 1, -0.0137171
plineObj.SetBulge 2, 0
plineObj.SetBulge 3, -0.0603404
plineObj.SetBulge 4, -0.0618883
plineObj.SetBulge 5, -0.27718
plineObj.SetBulge 6, 0.388064
plineObj.SetBulge 7, 0
plineObj.SetBulge 8, 0.103341
plineObj.SetBulge 9, 0.124064
plineObj.SetBulge 10, 0.490174
plineObj.SetBulge 11, 0.0900062
plineObj.SetBulge 12, 0
plineObj.SetBulge 13, 0.0642505
plineObj.SetBulge 14, 0.673959
plineObj.SetBulge 15, 0.0711555
plineObj.SetBulge 16, 0
plineObj.SetBulge 17, -0.0526825
plineObj.SetBulge 18, -0.175037
plineObj.SetBulge 19, 0.393435

 plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True
  
  RetVal = plineObj.Mirror(b1, b2)
  RetVal = plineObj.Mirror(a1, a2)
  RetVal = plineObj.Copy
  ' Define the rotation of 45 degrees about a
  ' base point of (4, 4.25, 0)
  basePoint(0) = points2(4) - (b / 2): basePoint(1) = points2(3) - (a / 2)
  rotationAngle = 3.14159   ' 45 degrees
  plineObj.Rotate basePoint, rotationAngle
  plineObj.Update
  
pointsaktr4(0) = points2(0) + 17.45:      pointsaktr4(1) = points2(1) + 149.5
pointsaktr4(2) = points2(0) + 17.45:      pointsaktr4(3) = points2(1) + 132.254
pointsaktr4(4) = points2(0) + 17.7239:    pointsaktr4(5) = points2(1) + 128.781
pointsaktr4(6) = points2(0) + 17.6402:    pointsaktr4(7) = points2(1) + 128.65
pointsaktr4(8) = points2(0) + 17.2757:    pointsaktr4(9) = points2(1) + 128.422
pointsaktr4(10) = points2(0) + 16.6942:   pointsaktr4(11) = points2(1) + 127.512
pointsaktr4(12) = points2(0) + 16.0544:   pointsaktr4(13) = points2(1) + 126.126
pointsaktr4(14) = points2(0) + 15.4754:   pointsaktr4(15) = points2(1) + 125.07
pointsaktr4(16) = points2(0) + 15.245:    pointsaktr4(17) = points2(1) + 125.064
pointsaktr4(18) = points2(0) + 15.1237:   pointsaktr4(19) = points2(1) + 125.425
pointsaktr4(20) = points2(0) + 15.0098:   pointsaktr4(21) = points2(1) + 127.043
pointsaktr4(22) = points2(0) + 15:        pointsaktr4(23) = points2(1) + 149.5
pointsaktr4(24) = points2(0) + 15:        pointsaktr4(25) = points2(3) - 149.5
pointsaktr4(26) = points2(0) + 15.0098:   pointsaktr4(27) = points2(3) - 127.043
pointsaktr4(28) = points2(0) + 15.1237:   pointsaktr4(29) = points2(3) - 125.425
pointsaktr4(30) = points2(0) + 15.245:    pointsaktr4(31) = points2(3) - 125.064
pointsaktr4(32) = points2(0) + 15.4754:   pointsaktr4(33) = points2(3) - 125.07
pointsaktr4(34) = points2(0) + 16.0544:   pointsaktr4(35) = points2(3) - 126.126
pointsaktr4(36) = points2(0) + 16.6942:   pointsaktr4(37) = points2(3) - 127.512
pointsaktr4(38) = points2(0) + 17.2757:   pointsaktr4(39) = points2(3) - 128.422
pointsaktr4(40) = points2(0) + 17.6402:   pointsaktr4(41) = points2(3) - 128.65
pointsaktr4(42) = points2(0) + 17.7239:   pointsaktr4(43) = points2(3) - 128.781
pointsaktr4(44) = points2(0) + 17.45:     pointsaktr4(45) = points2(3) - 132.254
pointsaktr4(46) = points2(0) + 17.45:     pointsaktr4(47) = points2(3) - 149.5

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsaktr4)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update
  
plineObj.SetBulge 0, 0
plineObj.SetBulge 1, 0.0393737
plineObj.SetBulge 2, -0.380001
plineObj.SetBulge 3, 0.142137
plineObj.SetBulge 4, 0.0583339
plineObj.SetBulge 5, 0
plineObj.SetBulge 6, -0.0650531
plineObj.SetBulge 7, -0.486421
plineObj.SetBulge 8, -0.0826229
plineObj.SetBulge 9, -0.0364452
plineObj.SetBulge 10, 0
plineObj.SetBulge 11, 0
plineObj.SetBulge 12, 0
plineObj.SetBulge 13, -0.0364452
plineObj.SetBulge 14, -0.0826229
plineObj.SetBulge 15, -0.486421
plineObj.SetBulge 16, -0.0650531
plineObj.SetBulge 17, 0
plineObj.SetBulge 18, 0.0583339
plineObj.SetBulge 19, 0.142137
plineObj.SetBulge 20, -0.380001
plineObj.SetBulge 21, 0.0393737
plineObj.SetBulge 22, 0

RetVal = plineObj.Mirror(a1, a2)

pointsaktr5(0) = points2(0) + 21.9:      pointsaktr5(1) = points2(1) + 149.5
pointsaktr5(2) = points2(0) + 22.1744:   pointsaktr5(3) = points2(1) + 146.286
pointsaktr5(4) = points2(0) + 22.1018:   pointsaktr5(5) = points2(1) + 146.177
pointsaktr5(6) = points2(0) + 21.7847:   pointsaktr5(7) = points2(1) + 146.024
pointsaktr5(8) = points2(0) + 21.0323:   pointsaktr5(9) = points2(1) + 144.982
pointsaktr5(10) = points2(0) + 20.3048:  pointsaktr5(11) = points2(1) + 143.596
pointsaktr5(12) = points2(0) + 19.8882:  pointsaktr5(13) = points2(1) + 143.03
pointsaktr5(14) = points2(0) + 19.6951:  pointsaktr5(15) = points2(1) + 143.056
pointsaktr5(16) = points2(0) + 19.5185:  pointsaktr5(17) = points2(1) + 143.696
pointsaktr5(18) = points2(0) + 19.45:    pointsaktr5(19) = points2(1) + 145.219
pointsaktr5(20) = points2(0) + 19.45:    pointsaktr5(21) = points2(1) + 149.5
pointsaktr5(22) = points2(0) + 19.45:    pointsaktr5(23) = points2(3) - 149.5
pointsaktr5(24) = points2(0) + 19.45:    pointsaktr5(25) = points2(3) - 145.219
pointsaktr5(26) = points2(0) + 19.5185:  pointsaktr5(27) = points2(3) - 143.696
pointsaktr5(28) = points2(0) + 19.6951:  pointsaktr5(29) = points2(3) - 143.056
pointsaktr5(30) = points2(0) + 19.8882:  pointsaktr5(31) = points2(3) - 143.03
pointsaktr5(32) = points2(0) + 20.3048:  pointsaktr5(33) = points2(3) - 143.596
pointsaktr5(34) = points2(0) + 21.0323:  pointsaktr5(35) = points2(3) - 144.982
pointsaktr5(36) = points2(0) + 21.7847:  pointsaktr5(37) = points2(3) - 146.024
pointsaktr5(38) = points2(0) + 22.1018:  pointsaktr5(39) = points2(3) - 146.177
pointsaktr5(40) = points2(0) + 22.1744:  pointsaktr5(41) = points2(3) - 146.286
pointsaktr5(42) = points2(0) + 21.9:     pointsaktr5(43) = points2(3) - 149.5

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsaktr5)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update
  
plineObj.SetBulge 0, 0.0426022
plineObj.SetBulge 1, -0.398863
plineObj.SetBulge 2, 0.114352
plineObj.SetBulge 3, 0.0884516
plineObj.SetBulge 4, 0
plineObj.SetBulge 5, -0.0920327
plineObj.SetBulge 6, -0.457613
plineObj.SetBulge 7, -0.10162
plineObj.SetBulge 8, -0.0270408
plineObj.SetBulge 9, 0
plineObj.SetBulge 10, 0
plineObj.SetBulge 11, 0
plineObj.SetBulge 12, -0.0270408
plineObj.SetBulge 13, -0.10162
plineObj.SetBulge 14, -0.457613
plineObj.SetBulge 15, -0.0920327
plineObj.SetBulge 16, 0
plineObj.SetBulge 17, 0.0884516
plineObj.SetBulge 18, 0.114352
plineObj.SetBulge 19, -0.398863
plineObj.SetBulge 20, 0.0426022

RetVal = plineObj.Mirror(a1, a2)

'=============================CENTER===========================================
bcp = points2(0) + (b / 2)
acp = points2(1) + (a / 2)

pointscntr1(0) = bcp + 24.5833:     pointscntr1(1) = acp + 20.0873
pointscntr1(2) = bcp + 28.0246:     pointscntr1(3) = acp + 19.9972
pointscntr1(4) = bcp + 19.3482:     pointscntr1(5) = acp + 25.688
pointscntr1(6) = bcp + 26.8683:     pointscntr1(7) = acp + 32.6531
pointscntr1(8) = bcp + 27.427:      pointscntr1(9) = acp + 33.6769
pointscntr1(10) = bcp + 27.9934:    pointscntr1(11) = acp + 34.4973
pointscntr1(12) = bcp + 31.4343:    pointscntr1(13) = acp + 35.7712
pointscntr1(14) = bcp + 33.8541:    pointscntr1(15) = acp + 35.1915
pointscntr1(16) = bcp + 37:         pointscntr1(17) = acp + 33.4752
pointscntr1(18) = bcp + 30.8536:    pointscntr1(19) = acp + 38.2705
pointscntr1(20) = bcp + 23.058:     pointscntr1(21) = acp + 38.2446
pointscntr1(22) = bcp + 13.2405:    pointscntr1(23) = acp + 32.8606
pointscntr1(24) = bcp + 6.64503:    pointscntr1(25) = acp + 23.8124
pointscntr1(26) = bcp + 6.64503:    pointscntr1(27) = acp - 23.8124
pointscntr1(28) = bcp + 13.2405:    pointscntr1(29) = acp - 32.8606
pointscntr1(30) = bcp + 23.058:     pointscntr1(31) = acp - 38.2446
pointscntr1(32) = bcp + 30.8536:    pointscntr1(33) = acp - 38.2705
pointscntr1(34) = bcp + 37:         pointscntr1(35) = acp - 33.4752
pointscntr1(36) = bcp + 33.8541:    pointscntr1(37) = acp - 35.1915
pointscntr1(38) = bcp + 31.4343:    pointscntr1(39) = acp - 35.7712
pointscntr1(40) = bcp + 27.9934:    pointscntr1(41) = acp - 34.4973
pointscntr1(42) = bcp + 27.427:     pointscntr1(43) = acp - 33.6769
pointscntr1(44) = bcp + 26.8683:    pointscntr1(45) = acp - 32.6531
pointscntr1(46) = bcp + 19.3482:    pointscntr1(47) = acp - 25.688
pointscntr1(48) = bcp + 28.0246:    pointscntr1(49) = acp - 19.9972
pointscntr1(50) = bcp + 24.5833:    pointscntr1(51) = acp - 20.0873
pointscntr1(52) = bcp + 19.6471:    pointscntr1(53) = acp - 16.403
pointscntr1(54) = bcp + 6.28296:    pointscntr1(55) = acp - 9.466
pointscntr1(56) = bcp + 6.28296:    pointscntr1(57) = acp + 9.466
pointscntr1(58) = bcp + 19.6471:    pointscntr1(59) = acp + 16.403

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update

plineObj.SetBulge 0, -0.190956
plineObj.SetBulge 1, 0.493502
plineObj.SetBulge 2, 0.173886
plineObj.SetBulge 3, 0
plineObj.SetBulge 4, -0.0792127
plineObj.SetBulge 5, -0.236352
plineObj.SetBulge 6, -0.0741078
plineObj.SetBulge 7, -0.0595151
plineObj.SetBulge 8, 0.188765
plineObj.SetBulge 9, 0.132767
plineObj.SetBulge 10, 0.129543
plineObj.SetBulge 11, 0.0905785
plineObj.SetBulge 12, 0.219599
plineObj.SetBulge 13, 0.0905785
plineObj.SetBulge 14, 0.129543
plineObj.SetBulge 15, 0.132767
plineObj.SetBulge 16, 0.188765
plineObj.SetBulge 17, -0.0595151
plineObj.SetBulge 18, -0.0741078
plineObj.SetBulge 19, -0.236352
plineObj.SetBulge 20, -0.0792127
plineObj.SetBulge 21, 0
plineObj.SetBulge 22, 0.173886
plineObj.SetBulge 23, 0.493502
plineObj.SetBulge 24, -0.190956
plineObj.SetBulge 25, -0.133032
plineObj.SetBulge 26, -0.20958
plineObj.SetBulge 27, -0.334034
plineObj.SetBulge 28, -0.20958
plineObj.SetBulge 29, -0.133032

plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True

RetVal = plineObj.Mirror(a1, a2)

pointscntr2(0) = bcp + 6.58238:      pointscntr2(1) = acp + 31.3788
pointscntr2(2) = bcp + 4.25321:      pointscntr2(3) = acp + 33.6762
pointscntr2(4) = bcp + 0.898097:     pointscntr2(5) = acp + 33.4752
pointscntr2(6) = bcp + 1.68328:      pointscntr2(7) = acp + 35.9463
pointscntr2(8) = bcp + 1.70714:      pointscntr2(9) = acp + 37.0063
pointscntr2(10) = bcp + 0.782447:    pointscntr2(11) = acp + 37.995
pointscntr2(12) = bcp + 1.13687E-13: pointscntr2(13) = acp + 37.9676
pointscntr2(14) = bcp + -0.782447:   pointscntr2(15) = acp + 37.995
pointscntr2(16) = bcp + -1.70714:    pointscntr2(17) = acp + 37.0063
pointscntr2(18) = bcp + -1.68328:    pointscntr2(19) = acp + 35.9463
pointscntr2(20) = bcp + -0.898097:   pointscntr2(21) = acp + 33.4752
pointscntr2(22) = bcp + -4.25321:    pointscntr2(23) = acp + 33.6762
pointscntr2(24) = bcp + -6.58238:    pointscntr2(25) = acp + 31.3788
pointscntr2(26) = bcp + -3.92174:    pointscntr2(27) = acp + 29.4618
pointscntr2(28) = bcp + -1.70295:    pointscntr2(29) = acp + 25.1
pointscntr2(30) = bcp + -0.455718:   pointscntr2(31) = acp + 18.477
pointscntr2(32) = bcp + -0.227606:   pointscntr2(33) = acp + 15.167
pointscntr2(34) = bcp + -0.110535:   pointscntr2(35) = acp + 11.4844
pointscntr2(36) = bcp + -0.0689475:  pointscntr2(37) = acp + 10.7799
pointscntr2(38) = bcp + -0.0388717:  pointscntr2(39) = acp + 10.5512
pointscntr2(40) = bcp + 1.13687E-13: pointscntr2(41) = acp + 10.4124
pointscntr2(42) = bcp + 0.0388717:   pointscntr2(43) = acp + 10.5512
pointscntr2(44) = bcp + 0.0689475:   pointscntr2(45) = acp + 10.7799
pointscntr2(46) = bcp + 0.110535:    pointscntr2(47) = acp + 11.4844
pointscntr2(48) = bcp + 0.227606:    pointscntr2(49) = acp + 15.167
pointscntr2(50) = bcp + 0.455718:    pointscntr2(51) = acp + 18.477
pointscntr2(52) = bcp + 1.70295:     pointscntr2(53) = acp + 25.1
pointscntr2(54) = bcp + 3.92174:     pointscntr2(55) = acp + 29.4618


Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr2)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update

plineObj.SetBulge 0, 0.130293
plineObj.SetBulge 1, 0.286262
plineObj.SetBulge 2, 0.0628827
plineObj.SetBulge 3, 0.0664767
plineObj.SetBulge 4, 0.354054
plineObj.SetBulge 5, 0.0749942
plineObj.SetBulge 6, 0.0749942
plineObj.SetBulge 7, 0.354054
plineObj.SetBulge 8, 0.0664767
plineObj.SetBulge 9, 0.0628827
plineObj.SetBulge 10, 0.286262
plineObj.SetBulge 11, 0.130293
plineObj.SetBulge 12, -0.132961
plineObj.SetBulge 13, -0.0949138
plineObj.SetBulge 14, -0.0510213
plineObj.SetBulge 15, -0.0125717
plineObj.SetBulge 16, -0.00334278
plineObj.SetBulge 17, 0
plineObj.SetBulge 18, 0
plineObj.SetBulge 19, 0
plineObj.SetBulge 20, 0
plineObj.SetBulge 21, 0
plineObj.SetBulge 22, 0
plineObj.SetBulge 23, -0.00334278
plineObj.SetBulge 24, -0.0125717
plineObj.SetBulge 25, -0.0510213
plineObj.SetBulge 26, -0.0949138
plineObj.SetBulge 27, -0.132961

plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True

RetVal = plineObj.Mirror(b1, b2)

pointscntr3(0) = bcp + 17.03:       pointscntr3(1) = acp + 80.0281
pointscntr3(2) = bcp + 12.3194:     pointscntr3(3) = acp + 76.2511
pointscntr3(4) = bcp + 12.7291:     pointscntr3(5) = acp + 73.3674
pointscntr3(6) = bcp + 16.4548:     pointscntr3(7) = acp + 74.8083
pointscntr3(8) = bcp + 22.4096:     pointscntr3(9) = acp + 72.6438
pointscntr3(10) = bcp + 24.1819:    pointscntr3(11) = acp + 69.1113
pointscntr3(12) = bcp + 24.8313:    pointscntr3(13) = acp + 64.3255
pointscntr3(14) = bcp + 23.9272:    pointscntr3(15) = acp + 59.0018
pointscntr3(16) = bcp + 16.5438:    pointscntr3(17) = acp + 46.828
pointscntr3(18) = bcp + 6.28228:    pointscntr3(19) = acp + 36.4703
pointscntr3(20) = bcp + 22.0575:    pointscntr3(21) = acp + 48.4905
pointscntr3(22) = bcp + 28.7564:    pointscntr3(23) = acp + 65.9369
pointscntr3(24) = bcp + 25.5548:    pointscntr3(25) = acp + 74.964

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr3)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update

plineObj.SetBulge 0, 0.424127
plineObj.SetBulge 1, 0.161605
plineObj.SetBulge 2, 0.83194
plineObj.SetBulge 3, -0.267392
plineObj.SetBulge 4, -0.106986
plineObj.SetBulge 5, -0.0601812
plineObj.SetBulge 6, -0.0862239
plineObj.SetBulge 7, -0.0871434
plineObj.SetBulge 8, -0.0408194
plineObj.SetBulge 9, 0.0867111
plineObj.SetBulge 10, 0.208082
plineObj.SetBulge 11, 0.154888
plineObj.SetBulge 12, 0.189184

plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True
RetVal = plineObj.Mirror(b1, b2)
RetVal = plineObj.Mirror(a1, a2)
RetVal = plineObj.Copy
  ' Define the rotation of 45 degrees about a
  ' base point of (4, 4.25, 0)
  basePoint(0) = points2(4) - (b / 2): basePoint(1) = points2(3) - (a / 2)
  rotationAngle = 3.14159   ' 45 degrees
  plineObj.Rotate basePoint, rotationAngle
  plineObj.Update
  
pointscntr4(0) = bcp + 6.97983:      pointscntr4(1) = acp + 66.4733
pointscntr4(2) = bcp + 4.40328:      pointscntr4(3) = acp + 70.4044
pointscntr4(4) = bcp + 2.76705:      pointscntr4(5) = acp + 73.302
pointscntr4(6) = bcp + 2.03069:      pointscntr4(7) = acp + 75.4075
pointscntr4(8) = bcp + 1.13687E-13:  pointscntr4(9) = acp + 82.896
pointscntr4(10) = bcp + -2.03069:    pointscntr4(11) = acp + 75.4075
pointscntr4(12) = bcp + -2.76705:    pointscntr4(13) = acp + 73.302
pointscntr4(14) = bcp + -4.35149:    pointscntr4(15) = acp + 70.4834
pointscntr4(16) = bcp + -6.97983:    pointscntr4(17) = acp + 66.4733
pointscntr4(18) = bcp + -8.62407:    pointscntr4(19) = acp + 63.7027
pointscntr4(20) = bcp + -9.8543:     pointscntr4(21) = acp + 57.1808
pointscntr4(22) = bcp + -8.14995:    pointscntr4(23) = acp + 53.3016
pointscntr4(24) = bcp + -5.61959:    pointscntr4(25) = acp + 51.6618
pointscntr4(26) = bcp + -2.27229:    pointscntr4(27) = acp + 51.8809
pointscntr4(28) = bcp + 1.13687E-13: pointscntr4(29) = acp + 55.0398
pointscntr4(30) = bcp + 2.27229:     pointscntr4(31) = acp + 51.8809
pointscntr4(32) = bcp + 5.61959:     pointscntr4(33) = acp + 51.6618
pointscntr4(34) = bcp + 8.14995:     pointscntr4(35) = acp + 53.3016
pointscntr4(36) = bcp + 9.8543:      pointscntr4(37) = acp + 57.1808
pointscntr4(38) = bcp + 8.62407:     pointscntr4(39) = acp + 63.7027

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr4)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update

plineObj.SetBulge 0, 0
plineObj.SetBulge 1, -0.0340478
plineObj.SetBulge 2, -0.0404214
plineObj.SetBulge 3, 0
plineObj.SetBulge 4, 0
plineObj.SetBulge 5, -0.0404214
plineObj.SetBulge 6, -0.0330812
plineObj.SetBulge 7, 0
plineObj.SetBulge 8, 0.021196
plineObj.SetBulge 9, 0.173455
plineObj.SetBulge 10, 0.145317
plineObj.SetBulge 11, 0.143574
plineObj.SetBulge 12, 0.203485
plineObj.SetBulge 13, 0.200979
plineObj.SetBulge 14, 0.200979
plineObj.SetBulge 15, 0.203485
plineObj.SetBulge 16, 0.143574
plineObj.SetBulge 17, 0.145317
plineObj.SetBulge 18, 0.173455
plineObj.SetBulge 19, 0.021196

plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True
RetVal = plineObj.Mirror(b1, b2)

pointscntr5(0) = bcp + 8.66267:     pointscntr5(1) = acp + 94.5619
pointscntr5(2) = bcp + 12.0028:     pointscntr5(3) = acp + 94.7543
pointscntr5(4) = bcp + 14.9353:     pointscntr5(5) = acp + 93.6436
pointscntr5(6) = bcp + 16.3813:     pointscntr5(7) = acp + 91.4456
pointscntr5(8) = bcp + 16.3602:     pointscntr5(9) = acp + 89.5013
pointscntr5(10) = bcp + 14.3893:    pointscntr5(11) = acp + 86.8541
pointscntr5(12) = bcp + 11.3686:    pointscntr5(13) = acp + 86.1899
pointscntr5(14) = bcp + 13.0885:    pointscntr5(15) = acp + 87.6125
pointscntr5(16) = bcp + 12.9123:    pointscntr5(17) = acp + 89.239
pointscntr5(18) = bcp + 10.7834:    pointscntr5(19) = acp + 90.4247
pointscntr5(20) = bcp + 5.27888:    pointscntr5(21) = acp + 89.7626
pointscntr5(22) = bcp + 2.39309:    pointscntr5(23) = acp + 86.7891
pointscntr5(24) = bcp + 1.49609:    pointscntr5(25) = acp + 88.2863
pointscntr5(26) = bcp + 4.32014:    pointscntr5(27) = acp + 92.3133

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr5)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update
  
plineObj.SetBulge 0, -0.0883845
plineObj.SetBulge 1, -0.121995
plineObj.SetBulge 2, -0.193108
plineObj.SetBulge 3, -0.0993998
plineObj.SetBulge 4, -0.23388
plineObj.SetBulge 5, -0.125673
plineObj.SetBulge 6, 0.11811
plineObj.SetBulge 7, 0.412476
plineObj.SetBulge 8, 0.0962103
plineObj.SetBulge 9, 0.224454
plineObj.SetBulge 10, 0.11374
plineObj.SetBulge 11, 0
plineObj.SetBulge 12, -0.11562
plineObj.SetBulge 13, -0.128851

plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True
RetVal = plineObj.Mirror(b1, b2)
RetVal = plineObj.Mirror(a1, a2)
RetVal = plineObj.Copy
  ' Define the rotation of 45 degrees about a
  ' base point of (4, 4.25, 0)
  basePoint(0) = points2(4) - (b / 2): basePoint(1) = points2(3) - (a / 2)
  rotationAngle = 3.14159   ' 45 degrees
  plineObj.Rotate basePoint, rotationAngle
  plineObj.Update

pointscntr6(0) = bcp + 3.75304:       pointscntr6(1) = acp + 101.938
pointscntr6(2) = bcp + 1.13687E-13:   pointscntr6(3) = acp + 111.5
pointscntr6(4) = bcp + -3.75304:      pointscntr6(5) = acp + 101.938
pointscntr6(6) = bcp + -3.37708:      pointscntr6(7) = acp + 96.8397
pointscntr6(8) = bcp + 1.13687E-13:   pointscntr6(9) = acp + 88.5857
pointscntr6(10) = bcp + 3.37708:      pointscntr6(11) = acp + 96.8397

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr6)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update
  
plineObj.SetBulge 0, 0.06872
plineObj.SetBulge 1, 0.06872
plineObj.SetBulge 2, 0.171056
plineObj.SetBulge 3, 0
plineObj.SetBulge 4, 0
plineObj.SetBulge 5, 0.171056

    plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True
RetVal = plineObj.Mirror(b1, b2)

center(0) = a1(0): center(1) = b1(1) + 45.76: center(2) = 0: radius = 2.9613
Set circleObj = ThisDrawing.ModelSpace.AddCircle(center, radius)
circleObj.Layer = "K-grav_Pattern"
circleObj.Update
center(0) = a1(0): center(1) = b1(1) - 45.76: center(2) = 0: radius = 2.9613
Set circleObj = ThisDrawing.ModelSpace.AddCircle(center, radius)
circleObj.Layer = "K-grav_Pattern"
circleObj.Update

End If
End If
  
  
  points2(1) = points2(3) + I:
  points2(3) = points2(3) + a + I:
  points2(5) = points2(3):
  points2(7) = points2(1):
  

Else: d = 1
d = d + 1

d = Cells(c, 4)
End If

Next d
c = c + 1

a = Cells(c, 2)
b = Cells(c, 3)

I = 100

Loop
              
ZoomAll
AP.Quit
End Sub

Sub MF140()
' Creating search objects
Dim AP As Excel.Application
Dim WB As Excel.Workbook
Dim WS As Excel.Worksheet

Set AP = Excel.Application
 Set WB = AP.Workbooks.Open("C:\Users\������\Desktop\������\2.xlsx")
Set WS = WB.Worksheets("������ �˨������ ������")
c = 19
a = Cells(c, 2)
b = Cells(c, 3)
d = Cells(c, 4)
E = 1
I = 0
On Error Resume Next

' Description of all the points that will make up rectangles
  Dim plineObj As AcadLWPolyline
  Dim plineObjw1 As AcadLWPolyline
  Dim plineObjw2 As AcadLWPolyline
  Dim points(0 To 7) As Double
  Dim pointsaktr1(0 To 399) As Double
  Dim pointsaktr2(0 To 151) As Double
  Dim pointsaktr3(0 To 9) As Double
  Dim pointsaktr4(0 To 13) As Double
  Dim pointsaktr5(0 To 5) As Double
  Dim pointsaktr6(0 To 9) As Double
  Dim pointsaktr7(0 To 7) As Double
  Dim pointscntr1(0 To 55) As Double
  Dim pointscntr2(0 To 7) As Double
  Dim pointscntr3(0 To 9) As Double
  Dim pointscntr4(0 To 7) As Double
  Dim pointscntr5(0 To 5) As Double
  Dim pointscntr6(0 To 5) As Double
  Dim pointscntr7(0 To 7) As Double
  Dim pointswithin(0 To 19) As Double
  Dim pointswithin2(0 To 27) As Double
  Dim intPointsa
  Dim intPointsb
  Dim pointshelpa1(0 To 3) As Double
  Dim pointshelpa2(0 To 3) As Double
  Dim pointshelpb1(0 To 3) As Double
  Dim pointshelpb2(0 To 3) As Double
  Dim offsetObj As Variant
  Dim basePoint(0 To 2) As Double
  Dim rotationAngle As Double
  Dim b1(0 To 2) As Double
  Dim b2(0 To 2) As Double
  Dim a1(0 To 2) As Double
  Dim a2(0 To 2) As Double

points(6) = 0

Do While a > 0

' Calculate the coordinates of each point
  points(0) = points(6) + I:    points(1) = 0
  points(2) = points(6) + I:    points(3) = a
  points(4) = points(0) + b:    points(5) = a
  points(6) = points(0) + b:    points(7) = 0

' Constructing four segments and connecting the corresponding points

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
  plineObj.Closed = True

If a > 280 Then
If b >= 240 Then

  pointswithin(0) = points(0) + 70:                              pointswithin(1) = points(1) + 94
  pointswithin(2) = points(0) + 70:                              pointswithin(3) = points(3) - 94
  pointswithin(4) = points(0) + 70 + ((b - 140) / 4):            pointswithin(5) = points(3) - 82
  pointswithin(6) = points(0) + (b / 2):                         pointswithin(7) = points(3) - 70
  pointswithin(8) = points(4) - 70 - ((b - 140) / 4):            pointswithin(9) = points(3) - 82
  pointswithin(10) = points(4) - 70:                             pointswithin(11) = points(3) - 94
  pointswithin(12) = points(4) - 70:                             pointswithin(13) = points(1) + 94
  pointswithin(14) = points(4) - 70 - ((b - 140) / 4):           pointswithin(15) = points(1) + 82
  pointswithin(16) = points(0) + (b / 2):                        pointswithin(17) = points(1) + 70
  pointswithin(18) = points(0) + 70 + ((b - 140) / 4):           pointswithin(19) = points(1) + 82

  
If a > 400 Then
  

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointswithin)
  plineObj.Closed = True
   
   pb = (b - 140) / 4
   pa = (a - 140) / 4
ga = Sqr((pa * pa) + 144)
gb = Sqr((pb * pb) + 144)
   anglea = Atn((12 / ga) / Sqr((-12 / ga) * (12 / ga) + 1))
   angleb = Atn((12 / gb) / Sqr((-12 / gb) * (12 / gb) + 1))
   radiusa = (ga / 2) / Sin(12 / ga)
   radiusb = (gb / 2) / Sin(12 / gb)
   ha = radiusa * (1 - Cos(anglea))
   hb = radiusb * (1 - Cos(angleb))
   ka = ha / ga
   kb = hb / gb
    
    plineObj.SetBulge 1, kb * 2
    plineObj.SetBulge 2, -kb * 2
    plineObj.SetBulge 3, -kb * 2
    plineObj.SetBulge 4, kb * 2
    plineObj.SetBulge 6, kb * 2
    plineObj.SetBulge 7, -kb * 2
    plineObj.SetBulge 8, -kb * 2
    plineObj.SetBulge 9, kb * 2
    
    plineObj.Layer = "C-Mill"
    plineObj.Update
    plineObj.Closed = True

End If

  pointshelpb1(0) = points(0) + 70:      pointshelpb1(1) = points(3) - 73
  pointshelpb1(2) = points(0) + (b / 2): pointshelpb1(3) = points(3) - 49
  pointshelpb2(0) = points(0) + 70:      pointshelpb2(1) = points(3) - 94 + radiusb
  pointshelpb2(2) = pointswithin(4):    pointshelpb2(3) = pointswithin(5)
Set plineObjw1 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointshelpb1)
Set plineObjw2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointshelpb2)
intPointsb = plineObjw1.IntersectWith(plineObjw2, acExtendBoth)
Intersectionbx = intPointsb(0) - points(0)
Intersectionby = points(3) - intPointsb(1)
plineObjw1.Delete
plineObjw2.Delete
  
  pointswithin2(0) = points(0) + 49:                              pointswithin2(1) = points(1) + 73
  pointswithin2(2) = points(0) + 49:                              pointswithin2(3) = points(3) - 73
  pointswithin2(4) = points(0) + 70:                              pointswithin2(5) = points(3) - 73
  pointswithin2(6) = points(0) + Intersectionbx:                  pointswithin2(7) = points(3) - Intersectionby
  pointswithin2(8) = points(0) + (b / 2):                         pointswithin2(9) = points(3) - 49
  pointswithin2(10) = points(4) - Intersectionbx:                 pointswithin2(11) = points(3) - Intersectionby
  pointswithin2(12) = points(4) - 70:                             pointswithin2(13) = points(3) - 73
  pointswithin2(14) = points(4) - 49:                             pointswithin2(15) = points(3) - 73
  pointswithin2(16) = points(4) - 49:                             pointswithin2(17) = points(1) + 73
  pointswithin2(18) = points(4) - 70:                             pointswithin2(19) = points(1) + 73
  pointswithin2(20) = points(4) - Intersectionbx:                 pointswithin2(21) = points(1) + Intersectionby
  pointswithin2(22) = points(4) - (b / 2):                        pointswithin2(23) = points(1) + 49
  pointswithin2(24) = points(0) + Intersectionbx:                 pointswithin2(25) = points(1) + Intersectionby
  pointswithin2(26) = points(0) + 70:                             pointswithin2(27) = points(1) + 73

If a > 400 Then

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointswithin2)
  plineObj.Closed = True
    
    plineObj.SetBulge 2, kb * 2
    plineObj.SetBulge 3, -kb * 2
    plineObj.SetBulge 4, -kb * 2
    plineObj.SetBulge 5, kb * 2
    plineObj.SetBulge 9, kb * 2
    plineObj.SetBulge 10, -kb * 2
    plineObj.SetBulge 11, -kb * 2
    plineObj.SetBulge 12, kb * 2
  
   
    plineObj.Layer = "K-grav"
    plineObj.Update
    plineObj.Closed = True
  
End If

 
pointsaktr1(0) = points(0) + 117:       pointsaktr1(1) = points(1) + 20
pointsaktr1(2) = points(0) + 92.6462:   pointsaktr1(3) = points(1) + 20
pointsaktr1(4) = points(0) + 92.3323:   pointsaktr1(5) = points(1) + 19.9806
pointsaktr1(6) = points(0) + 91.3477:   pointsaktr1(7) = points(1) + 19.779
pointsaktr1(8) = points(0) + 91.1075:   pointsaktr1(9) = points(1) + 20
pointsaktr1(10) = points(0) + 93.4262:  pointsaktr1(11) = points(1) + 23.7281
pointsaktr1(12) = points(0) + 94.2774:  pointsaktr1(13) = points(1) + 24.9179
pointsaktr1(14) = points(0) + 110.71:   pointsaktr1(15) = points(1) + 31.0138
pointsaktr1(16) = points(0) + 89.6442:  pointsaktr1(17) = points(1) + 30.8565
pointsaktr1(18) = points(0) + 73.8143:  pointsaktr1(19) = points(1) + 22.1749
pointsaktr1(20) = points(0) + 59.8973:  pointsaktr1(21) = points(1) + 15.5509
pointsaktr1(22) = points(0) + 46.2255:  pointsaktr1(23) = points(1) + 17.2859
pointsaktr1(24) = points(0) + 36.6857:  pointsaktr1(25) = points(1) + 32.3791
pointsaktr1(26) = points(0) + 30.9681:  pointsaktr1(27) = points(1) + 39.9336
pointsaktr1(28) = points(0) + 26.2137:  pointsaktr1(29) = points(1) + 39.1948
pointsaktr1(30) = points(0) + 24.0654:  pointsaktr1(31) = points(1) + 39.2522
pointsaktr1(32) = points(0) + 25.0128:  pointsaktr1(33) = points(1) + 45.3756
pointsaktr1(34) = points(0) + 22.5192:  pointsaktr1(35) = points(1) + 47.4951
pointsaktr1(36) = points(0) + 20.0504:  pointsaktr1(37) = points(1) + 47.0064
pointsaktr1(38) = points(0) + 17.4885:  pointsaktr1(39) = points(1) + 45.2069
pointsaktr1(40) = points(0) + 17.0747:  pointsaktr1(41) = points(1) + 45.7786
pointsaktr1(42) = points(0) + 15.6084:  pointsaktr1(43) = points(1) + 47.1214
pointsaktr1(44) = points(0) + 15:       pointsaktr1(45) = points(1) + 47.9324
pointsaktr1(46) = points(0) + 19.7914:  pointsaktr1(47) = points(1) + 49.0263
pointsaktr1(48) = points(0) + 24.321:   pointsaktr1(49) = points(1) + 47.4196
pointsaktr1(50) = points(0) + 26.4955:  pointsaktr1(51) = points(1) + 44.2133
pointsaktr1(52) = points(0) + 27.7845:  pointsaktr1(53) = points(1) + 42.3558
pointsaktr1(54) = points(0) + 30.1043:  pointsaktr1(55) = points(1) + 42.1403
pointsaktr1(56) = points(0) + 29.1547:  pointsaktr1(57) = points(1) + 49.4473
pointsaktr1(58) = points(0) + 29.6646:  pointsaktr1(59) = points(1) + 52.8799
pointsaktr1(60) = points(0) + 34.0454:  pointsaktr1(61) = points(1) + 50.5427
pointsaktr1(62) = points(0) + 34.5844:  pointsaktr1(63) = points(1) + 43.5484
pointsaktr1(64) = points(0) + 32.4493:  pointsaktr1(65) = points(1) + 40.615
pointsaktr1(66) = points(0) + 36.7127:  pointsaktr1(67) = points(1) + 34.7043
pointsaktr1(68) = points(0) + 38.4184:  pointsaktr1(69) = points(1) + 45.122
pointsaktr1(70) = points(0) + 40.9074:  pointsaktr1(71) = points(1) + 58.1004
pointsaktr1(72) = points(0) + 40.6569:  pointsaktr1(73) = points(1) + 65.7011
pointsaktr1(74) = points(0) + 35.0593:  pointsaktr1(75) = points(1) + 58.0353
pointsaktr1(76) = points(0) + 34.7122:  pointsaktr1(77) = points(1) + 57.1492
pointsaktr1(78) = points(0) + 34.6115:  pointsaktr1(79) = points(1) + 57.1459
pointsaktr1(80) = points(0) + 34.5689:  pointsaktr1(81) = points(1) + 57.2598
pointsaktr1(82) = points(0) + 34.491:   pointsaktr1(83) = points(1) + 57.6791
pointsaktr1(84) = points(0) + 30.7607:  pointsaktr1(85) = points(1) + 72.4992
pointsaktr1(86) = points(0) + 29.4442:  pointsaktr1(87) = points(1) + 81.1508
pointsaktr1(88) = points(0) + 29.6822:  pointsaktr1(89) = points(1) + 84.3374
pointsaktr1(90) = points(0) + 30.3315:  pointsaktr1(91) = points(1) + 85.2978
pointsaktr1(92) = points(0) + 30.8867:  pointsaktr1(93) = points(1) + 85.4656
pointsaktr1(94) = points(0) + 33.0692:  pointsaktr1(95) = points(1) + 84.9308
pointsaktr1(96) = points(0) + 35:       pointsaktr1(97) = points(1) + 83.2755
pointsaktr1(98) = points(0) + 35:       pointsaktr1(99) = points(1) + 88
pointsaktr1(100) = points(0) + 35:      pointsaktr1(101) = points(3) - 88
pointsaktr1(102) = points(0) + 35:      pointsaktr1(103) = points(3) - 83.2755
pointsaktr1(104) = points(0) + 33.0692: pointsaktr1(105) = points(3) - 84.9308
pointsaktr1(106) = points(0) + 30.8867: pointsaktr1(107) = points(3) - 85.4656
pointsaktr1(108) = points(0) + 30.3315: pointsaktr1(109) = points(3) - 85.2978
pointsaktr1(110) = points(0) + 29.6822: pointsaktr1(111) = points(3) - 84.3374
pointsaktr1(112) = points(0) + 29.4442: pointsaktr1(113) = points(3) - 81.1508
pointsaktr1(114) = points(0) + 30.7607: pointsaktr1(115) = points(3) - 72.4992
pointsaktr1(116) = points(0) + 34.491:  pointsaktr1(117) = points(3) - 57.6791
pointsaktr1(118) = points(0) + 34.5689: pointsaktr1(119) = points(3) - 57.2598
pointsaktr1(120) = points(0) + 34.6115: pointsaktr1(121) = points(3) - 57.1459
pointsaktr1(122) = points(0) + 34.7122: pointsaktr1(123) = points(3) - 57.1492
pointsaktr1(124) = points(0) + 35.0593: pointsaktr1(125) = points(3) - 58.0353
pointsaktr1(126) = points(0) + 40.6569: pointsaktr1(127) = points(3) - 65.7011
pointsaktr1(128) = points(0) + 40.9074: pointsaktr1(129) = points(3) - 58.1004
pointsaktr1(130) = points(0) + 38.4184: pointsaktr1(131) = points(3) - 45.122
pointsaktr1(132) = points(0) + 36.7127: pointsaktr1(133) = points(3) - 34.7043
pointsaktr1(134) = points(0) + 32.4493: pointsaktr1(135) = points(3) - 40.615
pointsaktr1(136) = points(0) + 34.5844: pointsaktr1(137) = points(3) - 43.5484
pointsaktr1(138) = points(0) + 34.0454: pointsaktr1(139) = points(3) - 50.5427
pointsaktr1(140) = points(0) + 29.6646: pointsaktr1(141) = points(3) - 52.8799
pointsaktr1(142) = points(0) + 29.1547: pointsaktr1(143) = points(3) - 49.4473
pointsaktr1(144) = points(0) + 30.1043: pointsaktr1(145) = points(3) - 42.1403
pointsaktr1(146) = points(0) + 27.7845: pointsaktr1(147) = points(3) - 42.3558
pointsaktr1(148) = points(0) + 26.4955: pointsaktr1(149) = points(3) - 44.2133
pointsaktr1(150) = points(0) + 24.321:  pointsaktr1(151) = points(3) - 47.4196
pointsaktr1(152) = points(0) + 19.7914: pointsaktr1(153) = points(3) - 49.0263
pointsaktr1(154) = points(0) + 15:      pointsaktr1(155) = points(3) - 47.9324
pointsaktr1(156) = points(0) + 15.6084: pointsaktr1(157) = points(3) - 47.1214
pointsaktr1(158) = points(0) + 17.0747: pointsaktr1(159) = points(3) - 45.7786
pointsaktr1(160) = points(0) + 17.4885: pointsaktr1(161) = points(3) - 45.2069
pointsaktr1(162) = points(0) + 20.0504: pointsaktr1(163) = points(3) - 47.0064
pointsaktr1(164) = points(0) + 22.5192: pointsaktr1(165) = points(3) - 47.4951
pointsaktr1(166) = points(0) + 25.0128: pointsaktr1(167) = points(3) - 45.3756
pointsaktr1(168) = points(0) + 24.0654: pointsaktr1(169) = points(3) - 39.2522
pointsaktr1(170) = points(0) + 26.2137: pointsaktr1(171) = points(3) - 39.1948
pointsaktr1(172) = points(0) + 30.9681: pointsaktr1(173) = points(3) - 39.9336
pointsaktr1(174) = points(0) + 36.6857: pointsaktr1(175) = points(3) - 32.3791
pointsaktr1(176) = points(0) + 46.2255: pointsaktr1(177) = points(3) - 17.2859
pointsaktr1(178) = points(0) + 59.8973: pointsaktr1(179) = points(3) - 15.5509
pointsaktr1(180) = points(0) + 73.8143: pointsaktr1(181) = points(3) - 22.1749
pointsaktr1(182) = points(0) + 89.6442: pointsaktr1(183) = points(3) - 30.8565
pointsaktr1(184) = points(0) + 110.71:  pointsaktr1(185) = points(3) - 31.0138
pointsaktr1(186) = points(0) + 94.2774: pointsaktr1(187) = points(3) - 24.9179
pointsaktr1(188) = points(0) + 93.4262: pointsaktr1(189) = points(3) - 23.7281
pointsaktr1(190) = points(0) + 91.1075: pointsaktr1(191) = points(3) - 20
pointsaktr1(192) = points(0) + 91.3477: pointsaktr1(193) = points(3) - 19.779
pointsaktr1(194) = points(0) + 92.3323: pointsaktr1(195) = points(3) - 19.9806
pointsaktr1(196) = points(0) + 92.6462: pointsaktr1(197) = points(3) - 20
pointsaktr1(198) = points(0) + 117:     pointsaktr1(199) = points(3) - 20
pointsaktr1(200) = points(4) - 117:     pointsaktr1(201) = points(3) - 20
pointsaktr1(202) = points(4) - 92.6462: pointsaktr1(203) = points(3) - 20
pointsaktr1(204) = points(4) - 92.3323: pointsaktr1(205) = points(3) - 19.9806
pointsaktr1(206) = points(4) - 91.3477: pointsaktr1(207) = points(3) - 19.779
pointsaktr1(208) = points(4) - 91.1075: pointsaktr1(209) = points(3) - 20
pointsaktr1(210) = points(4) - 93.4262: pointsaktr1(211) = points(3) - 23.7281
pointsaktr1(212) = points(4) - 94.2774: pointsaktr1(213) = points(3) - 24.9179
pointsaktr1(214) = points(4) - 110.71:  pointsaktr1(215) = points(3) - 31.0138
pointsaktr1(216) = points(4) - 89.6442: pointsaktr1(217) = points(3) - 30.8565
pointsaktr1(218) = points(4) - 73.8143: pointsaktr1(219) = points(3) - 22.1749
pointsaktr1(220) = points(4) - 59.8973: pointsaktr1(221) = points(3) - 15.5509
pointsaktr1(222) = points(4) - 46.2255: pointsaktr1(223) = points(3) - 17.2859
pointsaktr1(224) = points(4) - 36.6857: pointsaktr1(225) = points(3) - 32.3791
pointsaktr1(226) = points(4) - 30.9681: pointsaktr1(227) = points(3) - 39.9336
pointsaktr1(228) = points(4) - 26.2137: pointsaktr1(229) = points(3) - 39.1948
pointsaktr1(230) = points(4) - 24.0654: pointsaktr1(231) = points(3) - 39.2522
pointsaktr1(232) = points(4) - 25.0128: pointsaktr1(233) = points(3) - 45.3756
pointsaktr1(234) = points(4) - 22.5192: pointsaktr1(235) = points(3) - 47.4951
pointsaktr1(236) = points(4) - 20.0504: pointsaktr1(237) = points(3) - 47.0064
pointsaktr1(238) = points(4) - 17.4885: pointsaktr1(239) = points(3) - 45.2069
pointsaktr1(240) = points(4) - 17.0747: pointsaktr1(241) = points(3) - 45.7786
pointsaktr1(242) = points(4) - 15.6084: pointsaktr1(243) = points(3) - 47.1214
pointsaktr1(244) = points(4) - 15:      pointsaktr1(245) = points(3) - 47.9324
pointsaktr1(246) = points(4) - 19.7914: pointsaktr1(247) = points(3) - 49.0263
pointsaktr1(248) = points(4) - 24.321:  pointsaktr1(249) = points(3) - 47.4196
pointsaktr1(250) = points(4) - 26.4955: pointsaktr1(251) = points(3) - 44.2133
pointsaktr1(252) = points(4) - 27.7845: pointsaktr1(253) = points(3) - 42.3558
pointsaktr1(254) = points(4) - 30.1043: pointsaktr1(255) = points(3) - 42.1403
pointsaktr1(256) = points(4) - 29.1547: pointsaktr1(257) = points(3) - 49.4473
pointsaktr1(258) = points(4) - 29.6646: pointsaktr1(259) = points(3) - 52.8799
pointsaktr1(260) = points(4) - 34.0454: pointsaktr1(261) = points(3) - 50.5427
pointsaktr1(262) = points(4) - 34.5844: pointsaktr1(263) = points(3) - 43.5484
pointsaktr1(264) = points(4) - 32.4493: pointsaktr1(265) = points(3) - 40.615
pointsaktr1(266) = points(4) - 36.7127: pointsaktr1(267) = points(3) - 34.7043
pointsaktr1(268) = points(4) - 38.4184: pointsaktr1(269) = points(3) - 45.122
pointsaktr1(270) = points(4) - 40.9074: pointsaktr1(271) = points(3) - 58.1004
pointsaktr1(272) = points(4) - 40.6569: pointsaktr1(273) = points(3) - 65.7011
pointsaktr1(274) = points(4) - 35.0593: pointsaktr1(275) = points(3) - 58.0353
pointsaktr1(276) = points(4) - 34.7122: pointsaktr1(277) = points(3) - 57.1492
pointsaktr1(278) = points(4) - 34.6115: pointsaktr1(279) = points(3) - 57.1459
pointsaktr1(280) = points(4) - 34.5689: pointsaktr1(281) = points(3) - 57.2598
pointsaktr1(282) = points(4) - 34.491:  pointsaktr1(283) = points(3) - 57.6791
pointsaktr1(284) = points(4) - 30.7607: pointsaktr1(285) = points(3) - 72.4992
pointsaktr1(286) = points(4) - 29.4442: pointsaktr1(287) = points(3) - 81.1508
pointsaktr1(288) = points(4) - 29.6822: pointsaktr1(289) = points(3) - 84.3374
pointsaktr1(290) = points(4) - 30.3315: pointsaktr1(291) = points(3) - 85.2978
pointsaktr1(292) = points(4) - 30.8867: pointsaktr1(293) = points(3) - 85.4656
pointsaktr1(294) = points(4) - 33.0692: pointsaktr1(295) = points(3) - 84.9308
pointsaktr1(296) = points(4) - 35:      pointsaktr1(297) = points(3) - 83.2755
pointsaktr1(298) = points(4) - 35:      pointsaktr1(299) = points(3) - 88
pointsaktr1(300) = points(4) - 35:      pointsaktr1(301) = points(1) + 88
pointsaktr1(302) = points(4) - 35:      pointsaktr1(303) = points(1) + 83.2755
pointsaktr1(304) = points(4) - 33.0692: pointsaktr1(305) = points(1) + 84.9308
pointsaktr1(306) = points(4) - 30.8867: pointsaktr1(307) = points(1) + 85.4656
pointsaktr1(308) = points(4) - 30.3315: pointsaktr1(309) = points(1) + 85.2978
pointsaktr1(310) = points(4) - 29.6822: pointsaktr1(311) = points(1) + 84.3374
pointsaktr1(312) = points(4) - 29.4442: pointsaktr1(313) = points(1) + 81.1508
pointsaktr1(314) = points(4) - 30.7607: pointsaktr1(315) = points(1) + 72.4992
pointsaktr1(316) = points(4) - 34.491:  pointsaktr1(317) = points(1) + 57.6791
pointsaktr1(318) = points(4) - 34.5689: pointsaktr1(319) = points(1) + 57.2598
pointsaktr1(320) = points(4) - 34.6115: pointsaktr1(321) = points(1) + 57.1459
pointsaktr1(322) = points(4) - 34.7122: pointsaktr1(323) = points(1) + 57.1492
pointsaktr1(324) = points(4) - 35.0593: pointsaktr1(325) = points(1) + 58.0353
pointsaktr1(326) = points(4) - 40.6569: pointsaktr1(327) = points(1) + 65.7011
pointsaktr1(328) = points(4) - 40.9074: pointsaktr1(329) = points(1) + 58.1004
pointsaktr1(330) = points(4) - 38.4184: pointsaktr1(331) = points(1) + 45.122
pointsaktr1(332) = points(4) - 36.7127: pointsaktr1(333) = points(1) + 34.7043
pointsaktr1(334) = points(4) - 32.4493: pointsaktr1(335) = points(1) + 40.615
pointsaktr1(336) = points(4) - 34.5844: pointsaktr1(337) = points(1) + 43.5484
pointsaktr1(338) = points(4) - 34.0454: pointsaktr1(339) = points(1) + 50.5427
pointsaktr1(340) = points(4) - 29.6646: pointsaktr1(341) = points(1) + 52.8799
pointsaktr1(342) = points(4) - 29.1547: pointsaktr1(343) = points(1) + 49.4473
pointsaktr1(344) = points(4) - 30.1043: pointsaktr1(345) = points(1) + 42.1403
pointsaktr1(346) = points(4) - 27.7845: pointsaktr1(347) = points(1) + 42.3558
pointsaktr1(348) = points(4) - 26.4955: pointsaktr1(349) = points(1) + 44.2133
pointsaktr1(350) = points(4) - 24.321:  pointsaktr1(351) = points(1) + 47.4196
pointsaktr1(352) = points(4) - 19.7914: pointsaktr1(353) = points(1) + 49.0263
pointsaktr1(354) = points(4) - 15:      pointsaktr1(355) = points(1) + 47.9324
pointsaktr1(356) = points(4) - 15.6084: pointsaktr1(357) = points(1) + 47.1214
pointsaktr1(358) = points(4) - 17.0747: pointsaktr1(359) = points(1) + 45.7786
pointsaktr1(360) = points(4) - 17.4885: pointsaktr1(361) = points(1) + 45.2069
pointsaktr1(362) = points(4) - 20.0504: pointsaktr1(363) = points(1) + 47.0064
pointsaktr1(364) = points(4) - 22.5192: pointsaktr1(365) = points(1) + 47.4951
pointsaktr1(366) = points(4) - 25.0128: pointsaktr1(367) = points(1) + 45.3756
pointsaktr1(368) = points(4) - 24.0654: pointsaktr1(369) = points(1) + 39.2522
pointsaktr1(370) = points(4) - 26.2137: pointsaktr1(371) = points(1) + 39.1948
pointsaktr1(372) = points(4) - 30.9681: pointsaktr1(373) = points(1) + 39.9336
pointsaktr1(374) = points(4) - 36.6857: pointsaktr1(375) = points(1) + 32.3791
pointsaktr1(376) = points(4) - 46.2255: pointsaktr1(377) = points(1) + 17.2859
pointsaktr1(378) = points(4) - 59.8973: pointsaktr1(379) = points(1) + 15.5509
pointsaktr1(380) = points(4) - 73.8143: pointsaktr1(381) = points(1) + 22.1749
pointsaktr1(382) = points(4) - 89.6442: pointsaktr1(383) = points(1) + 30.8565
pointsaktr1(384) = points(4) - 110.71:  pointsaktr1(385) = points(1) + 31.0138
pointsaktr1(386) = points(4) - 94.2774: pointsaktr1(387) = points(1) + 24.9179
pointsaktr1(388) = points(4) - 93.4262: pointsaktr1(389) = points(1) + 23.7281
pointsaktr1(390) = points(4) - 91.1075: pointsaktr1(391) = points(1) + 20
pointsaktr1(392) = points(4) - 91.3477: pointsaktr1(393) = points(1) + 19.779
pointsaktr1(394) = points(4) - 92.3323: pointsaktr1(395) = points(1) + 19.9806
pointsaktr1(396) = points(4) - 92.6462: pointsaktr1(397) = points(1) + 20
pointsaktr1(398) = points(4) - 117:     pointsaktr1(399) = points(1) + 20



Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsaktr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update
    
currentBulge = plineObj.GetBulge(2)

plineObj.SetBulge 0, 0
plineObj.SetBulge 1, 0.0177245
plineObj.SetBulge 2, 0.0523607
plineObj.SetBulge 3, -0.57942
plineObj.SetBulge 4, -0.168144
plineObj.SetBulge 5, 0.135111
plineObj.SetBulge 6, 0.269875
plineObj.SetBulge 7, 0.165606
plineObj.SetBulge 8, 0.0831599
plineObj.SetBulge 9, -0.112133
plineObj.SetBulge 10, -0.175336
plineObj.SetBulge 11, -0.273401
plineObj.SetBulge 12, -0.105483
plineObj.SetBulge 13, -0.135712
plineObj.SetBulge 14, 0.0444926
plineObj.SetBulge 15, 0.25299
plineObj.SetBulge 16, 0.2683
plineObj.SetBulge 17, 0.190049
plineObj.SetBulge 18, 0.0206626
plineObj.SetBulge 19, 0.0702883
plineObj.SetBulge 20, 0.0312334
plineObj.SetBulge 21, -0.124693
plineObj.SetBulge 22, -0.111948
plineObj.SetBulge 23, -0.172878
plineObj.SetBulge 24, -0.146843
plineObj.SetBulge 25, 0.152321
plineObj.SetBulge 26, 0.292528
plineObj.SetBulge 27, -0.0613614
plineObj.SetBulge 28, -0.0772198
plineObj.SetBulge 29, -0.322171
plineObj.SetBulge 30, -0.192556
plineObj.SetBulge 31, -0.164263
plineObj.SetBulge 32, 0.0753202
plineObj.SetBulge 33, -0.0424071
plineObj.SetBulge 34, 0.0288023
plineObj.SetBulge 35, 0.0825991
plineObj.SetBulge 36, 0.228651
plineObj.SetBulge 37, -0.0964151
plineObj.SetBulge 38, -0.528657
plineObj.SetBulge 39, -0.137223
plineObj.SetBulge 40, 0.049294
plineObj.SetBulge 41, -0.0178489
plineObj.SetBulge 42, -0.0299484
plineObj.SetBulge 43, -0.0830332
plineObj.SetBulge 44, -0.178953
plineObj.SetBulge 45, -0.165905
plineObj.SetBulge 46, -0.102816
plineObj.SetBulge 47, -0.132516
plineObj.SetBulge 48, 0
plineObj.SetBulge 49, 0
plineObj.SetBulge 50, 0
plineObj.SetBulge 51, -0.132516
plineObj.SetBulge 52, -0.102816
plineObj.SetBulge 53, -0.165905
plineObj.SetBulge 54, -0.178953
plineObj.SetBulge 55, -0.0830332
plineObj.SetBulge 56, -0.0299484
plineObj.SetBulge 57, -0.0178489
plineObj.SetBulge 58, 0.049294
plineObj.SetBulge 59, -0.137223
plineObj.SetBulge 60, -0.528657
plineObj.SetBulge 61, -0.0964151
plineObj.SetBulge 62, 0.228651
plineObj.SetBulge 63, 0.0825991
plineObj.SetBulge 64, 0.0288023
plineObj.SetBulge 65, -0.0424071
plineObj.SetBulge 66, 0.0753202
plineObj.SetBulge 67, -0.164263
plineObj.SetBulge 68, -0.192556
plineObj.SetBulge 69, -0.322171
plineObj.SetBulge 70, -0.0772198
plineObj.SetBulge 71, -0.0613614
plineObj.SetBulge 72, 0.292528
plineObj.SetBulge 73, 0.152321
plineObj.SetBulge 74, -0.146843
plineObj.SetBulge 75, -0.172878
plineObj.SetBulge 76, -0.111948
plineObj.SetBulge 77, -0.124693
plineObj.SetBulge 78, 0.0312334
plineObj.SetBulge 79, 0.0702883
plineObj.SetBulge 80, 0.0206626
plineObj.SetBulge 81, 0.190049
plineObj.SetBulge 82, 0.2683
plineObj.SetBulge 83, 0.25299
plineObj.SetBulge 84, 0.0444926
plineObj.SetBulge 85, -0.135712
plineObj.SetBulge 86, -0.105483
plineObj.SetBulge 87, -0.273401
plineObj.SetBulge 88, -0.175336
plineObj.SetBulge 89, -0.112133
plineObj.SetBulge 90, 0.0831599
plineObj.SetBulge 91, 0.165606
plineObj.SetBulge 92, 0.269875
plineObj.SetBulge 93, 0.135111
plineObj.SetBulge 94, -0.168144
plineObj.SetBulge 95, -0.57942
plineObj.SetBulge 96, 0.0523607
plineObj.SetBulge 97, 0.0177245
plineObj.SetBulge 98, 0
plineObj.SetBulge 99, 0
plineObj.SetBulge 100, 0
plineObj.SetBulge 101, 0.0177245
plineObj.SetBulge 102, 0.0523607
plineObj.SetBulge 103, -0.57942
plineObj.SetBulge 104, -0.168144
plineObj.SetBulge 105, 0.135111
plineObj.SetBulge 106, 0.269875
plineObj.SetBulge 107, 0.165606
plineObj.SetBulge 108, 0.0831599
plineObj.SetBulge 109, -0.112133
plineObj.SetBulge 110, -0.175336
plineObj.SetBulge 111, -0.273401
plineObj.SetBulge 112, -0.105483
plineObj.SetBulge 113, -0.135712
plineObj.SetBulge 114, 0.0444926
plineObj.SetBulge 115, 0.25299
plineObj.SetBulge 116, 0.2683
plineObj.SetBulge 117, 0.190049
plineObj.SetBulge 118, 0.0206626
plineObj.SetBulge 119, 0.0702883
plineObj.SetBulge 120, 0.0312334
plineObj.SetBulge 121, -0.124693
plineObj.SetBulge 122, -0.111948
plineObj.SetBulge 123, -0.172878
plineObj.SetBulge 124, -0.146843
plineObj.SetBulge 125, 0.152321
plineObj.SetBulge 126, 0.292528
plineObj.SetBulge 127, -0.0613614
plineObj.SetBulge 128, -0.0772198
plineObj.SetBulge 129, -0.322171
plineObj.SetBulge 130, -0.192556
plineObj.SetBulge 131, -0.164263
plineObj.SetBulge 132, 0.0753202
plineObj.SetBulge 133, -0.0424071
plineObj.SetBulge 134, 0.0288023
plineObj.SetBulge 135, 0.0825991
plineObj.SetBulge 136, 0.228651
plineObj.SetBulge 137, -0.0964151
plineObj.SetBulge 138, -0.528657
plineObj.SetBulge 139, -0.137223
plineObj.SetBulge 140, 0.049294
plineObj.SetBulge 141, -0.0178489
plineObj.SetBulge 142, -0.0299484
plineObj.SetBulge 143, -0.0830332
plineObj.SetBulge 144, -0.178953
plineObj.SetBulge 145, -0.165905
plineObj.SetBulge 146, -0.102816
plineObj.SetBulge 147, -0.132516
plineObj.SetBulge 148, 0
plineObj.SetBulge 149, 0
plineObj.SetBulge 150, 0
plineObj.SetBulge 151, -0.132516
plineObj.SetBulge 152, -0.102816
plineObj.SetBulge 153, -0.165905
plineObj.SetBulge 154, -0.178953
plineObj.SetBulge 155, -0.0830332
plineObj.SetBulge 156, -0.0299484
plineObj.SetBulge 157, -0.0178489
plineObj.SetBulge 158, 0.049294
plineObj.SetBulge 159, -0.137223
plineObj.SetBulge 160, -0.528657
plineObj.SetBulge 161, -0.0964151
plineObj.SetBulge 162, 0.228651
plineObj.SetBulge 163, 0.0825991
plineObj.SetBulge 164, 0.0288023
plineObj.SetBulge 165, -0.0424071
plineObj.SetBulge 166, 0.0753202
plineObj.SetBulge 167, -0.164263
plineObj.SetBulge 168, -0.192556
plineObj.SetBulge 169, -0.322171
plineObj.SetBulge 170, -0.0772198
plineObj.SetBulge 171, -0.0613614
plineObj.SetBulge 172, 0.292528
plineObj.SetBulge 173, 0.152321
plineObj.SetBulge 174, -0.146843
plineObj.SetBulge 175, -0.172878
plineObj.SetBulge 176, -0.111948
plineObj.SetBulge 177, -0.124693
plineObj.SetBulge 178, 0.0312334
plineObj.SetBulge 179, 0.0702883
plineObj.SetBulge 180, 0.0206626
plineObj.SetBulge 181, 0.190049
plineObj.SetBulge 182, 0.2683
plineObj.SetBulge 183, 0.25299
plineObj.SetBulge 184, 0.0444926
plineObj.SetBulge 185, -0.135712
plineObj.SetBulge 186, -0.105483
plineObj.SetBulge 187, -0.273401
plineObj.SetBulge 188, -0.175336
plineObj.SetBulge 189, -0.112133
plineObj.SetBulge 190, 0.0831599
plineObj.SetBulge 191, 0.165606
plineObj.SetBulge 192, 0.269875
plineObj.SetBulge 193, 0.135111
plineObj.SetBulge 194, -0.168144
plineObj.SetBulge 195, -0.57942
plineObj.SetBulge 196, 0.0523607
plineObj.SetBulge 197, 0.0177245
plineObj.SetBulge 198, 0
plineObj.SetBulge 199, 0
plineObj.SetBulge 200, 0

    
    plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True
  

pointsaktr2(0) = points(0) + 117:       pointsaktr2(1) = points(1) + 22.5
pointsaktr2(2) = points(0) + 99.158:    pointsaktr2(3) = points(1) + 22.5
pointsaktr2(4) = points(0) + 110.917:   pointsaktr2(5) = points(1) + 24.1929
pointsaktr2(6) = points(0) + 114.984:   pointsaktr2(7) = points(1) + 26.2407
pointsaktr2(8) = points(0) + 115.489:   pointsaktr2(9) = points(1) + 28.9732
pointsaktr2(10) = points(0) + 112.864:  pointsaktr2(11) = points(1) + 31.6992
pointsaktr2(12) = points(0) + 115.371:  pointsaktr2(13) = points(1) + 49.2106
pointsaktr2(14) = points(0) + 113.001:  pointsaktr2(15) = points(1) + 49.5365
pointsaktr2(16) = points(0) + 111.656:  pointsaktr2(17) = points(1) + 32.2762
pointsaktr2(18) = points(0) + 89.15:    pointsaktr2(19) = points(1) + 32.2728
pointsaktr2(20) = points(0) + 72.8856:  pointsaktr2(21) = points(1) + 23.3529
pointsaktr2(22) = points(0) + 59.5687:  pointsaktr2(23) = points(1) + 17.0144
pointsaktr2(24) = points(0) + 46.9093:  pointsaktr2(25) = points(1) + 18.621
pointsaktr2(26) = points(0) + 38.268:   pointsaktr2(27) = points(1) + 31.2047
pointsaktr2(28) = points(0) + 50.5636:  pointsaktr2(29) = points(1) + 28.0599
pointsaktr2(30) = points(0) + 49.9818:  pointsaktr2(31) = points(1) + 69.0067
pointsaktr2(32) = points(0) + 41.8998:  pointsaktr2(33) = points(1) + 67.5444
pointsaktr2(34) = points(0) + 37:       pointsaktr2(35) = points(1) + 80.4314
pointsaktr2(36) = points(0) + 37:       pointsaktr2(37) = points(1) + 88
pointsaktr2(38) = points(0) + 37:       pointsaktr2(39) = points(3) - 88
pointsaktr2(40) = points(0) + 37:       pointsaktr2(41) = points(3) - 80.4314
pointsaktr2(42) = points(0) + 41.8998:  pointsaktr2(43) = points(3) - 67.5444
pointsaktr2(44) = points(0) + 49.9818:  pointsaktr2(45) = points(3) - 69.0067
pointsaktr2(46) = points(0) + 50.5636:  pointsaktr2(47) = points(3) - 28.0599
pointsaktr2(48) = points(0) + 38.268:   pointsaktr2(49) = points(3) - 31.2047
pointsaktr2(50) = points(0) + 46.9093:  pointsaktr2(51) = points(3) - 18.621
pointsaktr2(52) = points(0) + 59.5687:  pointsaktr2(53) = points(3) - 17.0144
pointsaktr2(54) = points(0) + 72.8856:  pointsaktr2(55) = points(3) - 23.3529
pointsaktr2(56) = points(0) + 89.15:    pointsaktr2(57) = points(3) - 32.2728
pointsaktr2(58) = points(0) + 111.656:  pointsaktr2(59) = points(3) - 32.2762
pointsaktr2(60) = points(0) + 113.001:  pointsaktr2(61) = points(3) - 49.5365
pointsaktr2(62) = points(0) + 115.371:  pointsaktr2(63) = points(3) - 49.2106
pointsaktr2(64) = points(0) + 112.864:  pointsaktr2(65) = points(3) - 31.6992
pointsaktr2(66) = points(0) + 115.489:  pointsaktr2(67) = points(3) - 28.9732
pointsaktr2(68) = points(0) + 114.984:  pointsaktr2(69) = points(3) - 26.2407
pointsaktr2(70) = points(0) + 110.917:  pointsaktr2(71) = points(3) - 24.1929
pointsaktr2(72) = points(0) + 99.158:   pointsaktr2(73) = points(3) - 22.5
pointsaktr2(74) = points(0) + 117:      pointsaktr2(75) = points(3) - 22.5
pointsaktr2(76) = points(4) - 117:      pointsaktr2(77) = points(3) - 22.5
pointsaktr2(78) = points(4) - 99.158:   pointsaktr2(79) = points(3) - 22.5
pointsaktr2(80) = points(4) - 110.917:  pointsaktr2(81) = points(3) - 24.1929
pointsaktr2(82) = points(4) - 114.984:  pointsaktr2(83) = points(3) - 26.2407
pointsaktr2(84) = points(4) - 115.489:  pointsaktr2(85) = points(3) - 28.9732
pointsaktr2(86) = points(4) - 112.864:  pointsaktr2(87) = points(3) - 31.6992
pointsaktr2(88) = points(4) - 115.371:  pointsaktr2(89) = points(3) - 49.2106
pointsaktr2(90) = points(4) - 113.001:  pointsaktr2(91) = points(3) - 49.5365
pointsaktr2(92) = points(4) - 111.656:  pointsaktr2(93) = points(3) - 32.2762
pointsaktr2(94) = points(4) - 89.15:    pointsaktr2(95) = points(3) - 32.2728
pointsaktr2(96) = points(4) - 72.8856:  pointsaktr2(97) = points(3) - 23.3529
pointsaktr2(98) = points(4) - 59.5687:  pointsaktr2(99) = points(3) - 17.0144
pointsaktr2(100) = points(4) - 46.9093: pointsaktr2(101) = points(3) - 18.621
pointsaktr2(102) = points(4) - 38.268:  pointsaktr2(103) = points(3) - 31.2047
pointsaktr2(104) = points(4) - 50.5636: pointsaktr2(105) = points(3) - 28.0599
pointsaktr2(106) = points(4) - 49.9818: pointsaktr2(107) = points(3) - 69.0067
pointsaktr2(108) = points(4) - 41.8998: pointsaktr2(109) = points(3) - 67.5444
pointsaktr2(110) = points(4) - 37:      pointsaktr2(111) = points(3) - 80.4314
pointsaktr2(112) = points(4) - 37:      pointsaktr2(113) = points(3) - 88
pointsaktr2(114) = points(4) - 37:      pointsaktr2(115) = points(1) + 88
pointsaktr2(116) = points(4) - 37:      pointsaktr2(117) = points(1) + 80.4314
pointsaktr2(118) = points(4) - 41.8998: pointsaktr2(119) = points(1) + 67.5444
pointsaktr2(120) = points(4) - 49.9818: pointsaktr2(121) = points(1) + 69.0067
pointsaktr2(122) = points(4) - 50.5636: pointsaktr2(123) = points(1) + 28.0599
pointsaktr2(124) = points(4) - 38.268:  pointsaktr2(125) = points(1) + 31.2047
pointsaktr2(126) = points(4) - 46.9093: pointsaktr2(127) = points(1) + 18.621
pointsaktr2(128) = points(4) - 59.5687: pointsaktr2(129) = points(1) + 17.0144
pointsaktr2(130) = points(4) - 72.8856: pointsaktr2(131) = points(1) + 23.3529
pointsaktr2(132) = points(4) - 89.15:   pointsaktr2(133) = points(1) + 32.2728
pointsaktr2(134) = points(4) - 111.656: pointsaktr2(135) = points(1) + 32.2762
pointsaktr2(136) = points(4) - 113.001: pointsaktr2(137) = points(1) + 49.5365
pointsaktr2(138) = points(4) - 115.371: pointsaktr2(139) = points(1) + 49.2106
pointsaktr2(140) = points(4) - 112.864: pointsaktr2(141) = points(1) + 31.6992
pointsaktr2(142) = points(4) - 115.489: pointsaktr2(143) = points(1) + 28.9732
pointsaktr2(144) = points(4) - 114.984: pointsaktr2(145) = points(1) + 26.2407
pointsaktr2(146) = points(4) - 110.917: pointsaktr2(147) = points(1) + 24.1929
pointsaktr2(148) = points(4) - 99.158:  pointsaktr2(149) = points(1) + 22.5
pointsaktr2(150) = points(4) - 117:     pointsaktr2(151) = points(1) + 22.5
 
If b < 300 Then
pointsaktr2(13) = points(1) + 43
pointsaktr2(15) = points(1) + 43.4
pointsaktr2(61) = points(3) - 43.4
pointsaktr2(63) = points(3) - 43
pointsaktr2(89) = points(3) - 43
pointsaktr2(91) = points(3) - 43.4
pointsaktr2(137) = points(1) + 43.4
pointsaktr2(139) = points(1) + 43
End If

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsaktr2)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update
  
plineObj.SetBulge 0, 0
plineObj.SetBulge 1, 0.0427012
plineObj.SetBulge 2, 0.119624
plineObj.SetBulge 3, 0.355722
plineObj.SetBulge 4, 0.13362
plineObj.SetBulge 5, 0.234279
plineObj.SetBulge 6, 0
plineObj.SetBulge 7, -0.266061
plineObj.SetBulge 8, 0.169365
plineObj.SetBulge 9, 0.0831599
plineObj.SetBulge 10, -0.112133
plineObj.SetBulge 11, -0.175336
plineObj.SetBulge 12, -0.253039
plineObj.SetBulge 13, 0.149933
plineObj.SetBulge 14, 0.967494
plineObj.SetBulge 15, 0.0992359
plineObj.SetBulge 16, 0.0970037
plineObj.SetBulge 17, 0
plineObj.SetBulge 18, 0
plineObj.SetBulge 19, 0
plineObj.SetBulge 20, 0.0970037
plineObj.SetBulge 21, 0.0992359
plineObj.SetBulge 22, 0.967494
plineObj.SetBulge 23, 0.149933
plineObj.SetBulge 24, -0.253039
plineObj.SetBulge 25, -0.175336
plineObj.SetBulge 26, -0.112133
plineObj.SetBulge 27, 0.0831599
plineObj.SetBulge 28, 0.169365
plineObj.SetBulge 29, -0.266061
plineObj.SetBulge 30, 0
plineObj.SetBulge 31, 0.234279
plineObj.SetBulge 32, 0.13362
plineObj.SetBulge 33, 0.355722
plineObj.SetBulge 34, 0.119624
plineObj.SetBulge 35, 0.0427012
plineObj.SetBulge 36, 0
plineObj.SetBulge 37, 0
plineObj.SetBulge 38, 0
plineObj.SetBulge 39, 0.0427012
plineObj.SetBulge 40, 0.119624
plineObj.SetBulge 41, 0.355722
plineObj.SetBulge 42, 0.13362
plineObj.SetBulge 43, 0.234279
plineObj.SetBulge 44, 0
plineObj.SetBulge 45, -0.266061
plineObj.SetBulge 46, 0.169365
plineObj.SetBulge 47, 0.0831599
plineObj.SetBulge 48, -0.112133
plineObj.SetBulge 49, -0.175336
plineObj.SetBulge 50, -0.253039
plineObj.SetBulge 51, 0.149933
plineObj.SetBulge 52, 0.967494
plineObj.SetBulge 53, 0.0992359
plineObj.SetBulge 54, 0.0970037
plineObj.SetBulge 55, 0
plineObj.SetBulge 56, 0
plineObj.SetBulge 57, 0
plineObj.SetBulge 58, 0.0970037
plineObj.SetBulge 59, 0.0992359
plineObj.SetBulge 60, 0.967494
plineObj.SetBulge 61, 0.149933
plineObj.SetBulge 62, -0.253039
plineObj.SetBulge 63, -0.175336
plineObj.SetBulge 64, -0.112133
plineObj.SetBulge 65, 0.0831599
plineObj.SetBulge 66, 0.169365
plineObj.SetBulge 67, -0.266061
plineObj.SetBulge 68, 0
plineObj.SetBulge 69, 0.234279
plineObj.SetBulge 70, 0.13362
plineObj.SetBulge 71, 0.355722
plineObj.SetBulge 72, 0.119624
plineObj.SetBulge 73, 0.0427012
plineObj.SetBulge 74, 0
plineObj.SetBulge 75, 0

 plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True
  
pointsaktr3(0) = points(0) + 31.6684:   pointsaktr3(1) = points(1) + 42.8836
pointsaktr3(2) = points(0) + 30.5655:   pointsaktr3(3) = points(1) + 48.9538
pointsaktr3(4) = points(0) + 30.9385:   pointsaktr3(5) = points(1) + 51.4282
pointsaktr3(6) = points(0) + 33.3093:   pointsaktr3(7) = points(1) + 49.2005
pointsaktr3(8) = points(0) + 33.752:    pointsaktr3(9) = points(1) + 46.0948


Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsaktr3)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update
  
plineObj.SetBulge 0, -0.0681083
plineObj.SetBulge 1, -0.0969811
plineObj.SetBulge 2, -0.21617
plineObj.SetBulge 3, -0.125224
plineObj.SetBulge 4, -0.238364


 plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True
    
  b1(0) = (b / 2) - 1:  b1(1) = (a / 2)
  b2(0) = (b / 2) + 1:  b2(1) = (a / 2)
  a1(0) = points(4) - (b / 2): a1(1) = points(1) + (a / 2) - 1
  a2(0) = points(4) - (b / 2): a2(1) = points(1) + (a / 2) + 1
  RetVal = plineObj.Mirror(b1, b2)
  RetVal = plineObj.Mirror(a1, a2)
  RetVal = plineObj.Copy
  ' Define the rotation of 180 degrees about a
  ' base point of (4, 4.25, 0)
  basePoint(0) = points(4) - (b / 2): basePoint(1) = points(3) - (a / 2)
  rotationAngle = 3.14159   ' 180 degrees
  plineObj.Rotate basePoint, rotationAngle
  plineObj.Update
  
pointsaktr4(0) = points(0) + 32.2044:   pointsaktr4(1) = points(1) + 70.283
pointsaktr4(2) = points(0) + 34.6713:   pointsaktr4(3) = points(1) + 64.8485
pointsaktr4(4) = points(0) + 34.6713:   pointsaktr4(5) = points(1) + 79.75
pointsaktr4(6) = points(0) + 32.688:    pointsaktr4(7) = points(1) + 81.4884
pointsaktr4(8) = points(0) + 31.48:     pointsaktr4(9) = points(1) + 82.0684
pointsaktr4(10) = points(0) + 30.6893:  pointsaktr4(11) = points(1) + 81.5751
pointsaktr4(12) = points(0) + 30.6422:  pointsaktr4(13) = points(1) + 80.7573


Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsaktr4)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update
  
plineObj.SetBulge 0, 0.0818639
plineObj.SetBulge 1, 0
plineObj.SetBulge 2, 0.065715
plineObj.SetBulge 3, 0.0705315
plineObj.SetBulge 4, 0.461421
plineObj.SetBulge 5, 0.0454552
plineObj.SetBulge 6, 0.0574167


RetVal = plineObj.Mirror(a1, a2)
RetVal = plineObj.Mirror(b1, b2)
RetVal = plineObj.Copy
plineObj.Rotate basePoint, rotationAngle

pointsaktr5(0) = points(0) + 40.4778:   pointsaktr5(1) = points(1) + 66.8334
pointsaktr5(2) = points(0) + 36.4488:   pointsaktr5(3) = points(1) + 63.9005
pointsaktr5(4) = points(0) + 36.4488:   pointsaktr5(5) = points(1) + 77.5281


Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsaktr5)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update
  
plineObj.SetBulge 0, 0.0833799
plineObj.SetBulge 1, 0
plineObj.SetBulge 2, -0.0920228


RetVal = plineObj.Mirror(a1, a2)
RetVal = plineObj.Mirror(b1, b2)
RetVal = plineObj.Copy
plineObj.Rotate basePoint, rotationAngle

pointsaktr6(0) = points(0) + 50.5066:   pointsaktr6(1) = points(1) + 29.4087
pointsaktr6(2) = points(0) + 38.1331:   pointsaktr6(3) = points(1) + 33.8528
pointsaktr6(4) = points(0) + 40.8016:   pointsaktr6(5) = points(1) + 49.9506
pointsaktr6(6) = points(0) + 42.1013:   pointsaktr6(7) = points(1) + 66.3594
pointsaktr6(8) = points(0) + 49.9632:   pointsaktr6(9) = points(1) + 67.6568

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsaktr6)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update

plineObj.SetBulge 0, -0.19595
plineObj.SetBulge 1, -0.0514386
plineObj.SetBulge 2, 0.0942854
plineObj.SetBulge 3, -0.0888928
plineObj.SetBulge 4, -0.972406

RetVal = plineObj.Mirror(a1, a2)
RetVal = plineObj.Mirror(b1, b2)
RetVal = plineObj.Copy
plineObj.Rotate basePoint, rotationAngle

pointsaktr7(0) = points(0) + 105.446:   pointsaktr7(1) = points(1) + 25.2141
pointsaktr7(2) = points(0) + 113.409:   pointsaktr7(3) = points(1) + 26.5497
pointsaktr7(4) = points(0) + 113.971:   pointsaktr7(5) = points(1) + 28.6347
pointsaktr7(6) = points(0) + 111.928:   pointsaktr7(7) = points(1) + 30.4463

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsaktr7)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update
  
plineObj.SetBulge 0, 0.114003
plineObj.SetBulge 1, 0.49186
plineObj.SetBulge 2, 0.09754
plineObj.SetBulge 3, -0.105753


RetVal = plineObj.Mirror(a1, a2)
RetVal = plineObj.Mirror(b1, b2)
RetVal = plineObj.Copy
plineObj.Rotate basePoint, rotationAngle

'=============================CENTER===========================================
bcp = points(0) + (b / 2)
acp = points(1) + (a / 2)

pointscntr1(0) = bcp + 8.33227:     pointscntr1(1) = acp - 8.2554
pointscntr1(2) = bcp + 11.6963:     pointscntr1(3) = acp - 13.0893
pointscntr1(4) = bcp + 18.3729:     pointscntr1(5) = acp - 28.2801
pointscntr1(6) = bcp + 12.7667:     pointscntr1(7) = acp - 50.0277
pointscntr1(8) = bcp + 2.66319:     pointscntr1(9) = acp - 69.2627
pointscntr1(10) = bcp + 5.32307:    pointscntr1(11) = acp - 69.9733
pointscntr1(12) = bcp + 0:          pointscntr1(13) = acp - 120.5
pointscntr1(14) = bcp - 5.32307:  pointscntr1(15) = acp - 69.9733
pointscntr1(16) = bcp - 2.66319:  pointscntr1(17) = acp - 69.2627
pointscntr1(18) = bcp - 12.7667:  pointscntr1(19) = acp - 50.0277
pointscntr1(20) = bcp - 18.3729:  pointscntr1(21) = acp - 28.2801
pointscntr1(22) = bcp - 11.6963:  pointscntr1(23) = acp - 13.0893
pointscntr1(24) = bcp - 8.33227:  pointscntr1(25) = acp - 8.2554
pointscntr1(26) = bcp - 18.0805:  pointscntr1(27) = acp + 5.68434E-14
pointscntr1(28) = bcp - 8.33227:  pointscntr1(29) = acp + 8.2554
pointscntr1(30) = bcp - 11.6963:  pointscntr1(31) = acp + 13.0893
pointscntr1(32) = bcp - 18.3729:  pointscntr1(33) = acp + 28.2801
pointscntr1(34) = bcp - 12.7667:  pointscntr1(35) = acp + 50.0277
pointscntr1(36) = bcp - 2.66319:  pointscntr1(37) = acp + 69.2627
pointscntr1(38) = bcp - 5.32307:  pointscntr1(39) = acp + 69.9733
pointscntr1(40) = bcp + 0:          pointscntr1(41) = acp + 120.5
pointscntr1(42) = bcp + 5.32307:    pointscntr1(43) = acp + 69.9733
pointscntr1(44) = bcp + 2.66319:    pointscntr1(45) = acp + 69.2627
pointscntr1(46) = bcp + 12.7667:    pointscntr1(47) = acp + 50.0277
pointscntr1(48) = bcp + 18.3729:    pointscntr1(49) = acp + 28.2801
pointscntr1(50) = bcp + 11.6963:    pointscntr1(51) = acp + 13.0893
pointscntr1(52) = bcp + 8.33227:    pointscntr1(53) = acp + 8.2554
pointscntr1(54) = bcp + 18.0805:    pointscntr1(55) = acp - 4.73695E-14


Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update

plineObj.SetBulge 0, -0.00917899
plineObj.SetBulge 1, -0.0976773
plineObj.SetBulge 2, -0.221543
plineObj.SetBulge 3, -0.247182
plineObj.SetBulge 4, 0.035764
plineObj.SetBulge 5, 0
plineObj.SetBulge 6, 0
plineObj.SetBulge 7, 0.035764
plineObj.SetBulge 8, -0.247182
plineObj.SetBulge 9, -0.221543
plineObj.SetBulge 10, -0.0976773
plineObj.SetBulge 11, -0.00917899
plineObj.SetBulge 12, 0.0324358
plineObj.SetBulge 13, 0.0324358
plineObj.SetBulge 14, -0.00917899
plineObj.SetBulge 15, -0.0976773
plineObj.SetBulge 16, -0.221543
plineObj.SetBulge 17, -0.247182
plineObj.SetBulge 18, 0.035764
plineObj.SetBulge 19, 0
plineObj.SetBulge 20, 0
plineObj.SetBulge 21, 0.035764
plineObj.SetBulge 22, -0.247182
plineObj.SetBulge 23, -0.221543
plineObj.SetBulge 24, -0.0976773
plineObj.SetBulge 25, -0.00917899
plineObj.SetBulge 26, 0.0324358
plineObj.SetBulge 27, 0.0324358

plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True

pointscntr2(0) = bcp - 14.2411:   pointscntr2(1) = acp - 9.4739E-15
pointscntr2(2) = bcp - 6.93853:   pointscntr2(3) = acp + 6.3585
pointscntr2(4) = bcp - 1.93991:   pointscntr2(5) = acp + 3.78956E-14
pointscntr2(6) = bcp - 6.93853:   pointscntr2(7) = acp - 6.3585



Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr2)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update

plineObj.SetBulge 0, 0.0240064
plineObj.SetBulge 1, 0.012607
plineObj.SetBulge 2, 0.012607
plineObj.SetBulge 3, 0.0240064


plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True

RetVal = plineObj.Mirror(a1, a2)

pointscntr3(0) = bcp - 16.0864:   pointscntr3(1) = acp + 28.747
pointscntr3(2) = bcp - 9.76953:   pointscntr3(3) = acp + 14.4056
pointscntr3(4) = bcp - 6.65271:   pointscntr3(5) = acp + 9.91527
pointscntr3(6) = bcp - 1.45081:   pointscntr3(7) = acp + 15.6103
pointscntr3(8) = bcp - 12.8423:   pointscntr3(9) = acp + 45.9255

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr3)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update

plineObj.SetBulge 0, 0.0975544
plineObj.SetBulge 1, 0.00864503
plineObj.SetBulge 2, 0.0195719
plineObj.SetBulge 3, -0.168477
plineObj.SetBulge 4, 0.186466


plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True
RetVal = plineObj.Mirror(b1, b2)
RetVal = plineObj.Mirror(a1, a2)
RetVal = plineObj.Copy
  ' Define the rotation of 180 degrees about a
  ' base point of (4, 4.25, 0)
  basePoint(0) = points(4) - (b / 2): basePoint(1) = points(3) - (a / 2)
  rotationAngle = 3.14159   ' 180 degrees
  plineObj.Rotate basePoint, rotationAngle
  plineObj.Update
  
pointscntr4(0) = bcp + 5.25775:         pointscntr4(1) = acp + 8.0082
pointscntr4(2) = bcp - 5.68434E-14:   pointscntr4(3) = acp + 1.31597
pointscntr4(4) = bcp - 5.25775:       pointscntr4(5) = acp + 8.0082
pointscntr4(6) = bcp + 3.78956E-14:     pointscntr4(7) = acp + 13.7137


Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr4)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update

plineObj.SetBulge 0, -0.0134617
plineObj.SetBulge 1, -0.0134617
plineObj.SetBulge 2, 0.0192317
plineObj.SetBulge 3, 0.0192317


plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True
RetVal = plineObj.Mirror(b1, b2)

pointscntr5(0) = bcp - 10.4894:         pointscntr5(1) = acp + 49.0279
pointscntr5(2) = bcp + 1.89478E-14:     pointscntr5(3) = acp + 17.5202
pointscntr5(4) = bcp + 10.4894:         pointscntr5(5) = acp + 49.0279


Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr5)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update
  
plineObj.SetBulge 0, 0.182166
plineObj.SetBulge 1, 0.182166
plineObj.SetBulge 2, 0.419513


plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True
RetVal = plineObj.Mirror(b1, b2)


pointscntr6(0) = bcp - 10.1941:    pointscntr6(1) = acp + 52.3749
pointscntr6(2) = bcp + 10.1941:    pointscntr6(3) = acp + 52.3749
pointscntr6(4) = bcp + 0:          pointscntr6(5) = acp + 68.112

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr6)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update
  
plineObj.SetBulge 0, -0.332226
plineObj.SetBulge 1, 0.250892
plineObj.SetBulge 2, 0.250892


    plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True
RetVal = plineObj.Mirror(b1, b2)

pointscntr7(0) = bcp - 2.76853:         pointscntr7(1) = acp + 71.7473
pointscntr7(2) = bcp + 7.56728E-13:     pointscntr7(3) = acp + 70.7093
pointscntr7(4) = bcp + 2.76853:         pointscntr7(5) = acp + 71.7473
pointscntr7(6) = bcp + 0:               pointscntr7(7) = acp + 98.0267


Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr7)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update
  
plineObj.SetBulge 0, -0.0342582
plineObj.SetBulge 1, -0.0342582
plineObj.SetBulge 2, 0
plineObj.SetBulge 3, 0

    plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True
RetVal = plineObj.Mirror(b1, b2)

End If
End If
  
  
  
I = 100
  
 Dim plineObj2 As AcadLWPolyline
  Dim points2(0 To 7) As Double
  
  points2(0) = points(0):    points2(1) = a + I:
  points2(2) = points(0):    points2(3) = (a * 2) + I:
  points2(4) = points(6):    points2(5) = (a * 2) + I:
  points2(6) = points(6):    points2(7) = a + I:
  
'========================================================
'========================================================
'========================================================
  
  
 For d = 2 To Cells(c, 4)
If d > 1 Then

Set plineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(points2)
  plineObj2.Closed = True
  
If a > 280 Then
If b >= 240 Then
  
  pointswithin(0) = points2(0) + 70:                              pointswithin(1) = points2(1) + 94
  pointswithin(2) = points2(0) + 70:                              pointswithin(3) = points2(3) - 94
  pointswithin(4) = points2(0) + 70 + ((b - 140) / 4):            pointswithin(5) = points2(3) - 82
  pointswithin(6) = points2(0) + (b / 2):                         pointswithin(7) = points2(3) - 70
  pointswithin(8) = points2(4) - 70 - ((b - 140) / 4):            pointswithin(9) = points2(3) - 82
  pointswithin(10) = points2(4) - 70:                             pointswithin(11) = points2(3) - 94
  pointswithin(12) = points2(4) - 70:                             pointswithin(13) = points2(1) + 94
  pointswithin(14) = points2(4) - 70 - ((b - 140) / 4):           pointswithin(15) = points2(1) + 82
  pointswithin(16) = points2(0) + (b / 2):                        pointswithin(17) = points2(1) + 70
  pointswithin(18) = points2(0) + 70 + ((b - 140) / 4):           pointswithin(19) = points2(1) + 82
  

If a > 400 Then
  

 Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointswithin)
  plineObj.Closed = True
   
   pb = (b - 140) / 4
   pa = (a - 140) / 4
ga = Sqr((pa * pa) + 144)
gb = Sqr((pb * pb) + 144)
   anglea = Atn((12 / ga) / Sqr((-12 / ga) * (12 / ga) + 1))
   angleb = Atn((12 / gb) / Sqr((-12 / gb) * (12 / gb) + 1))
   radiusa = (ga / 2) / Sin(12 / ga)
   radiusb = (gb / 2) / Sin(12 / gb)
   ha = radiusa * (1 - Cos(anglea))
   hb = radiusb * (1 - Cos(angleb))
   ka = ha / ga
   kb = hb / gb
   
    plineObj.SetBulge 1, kb * 2
    plineObj.SetBulge 2, -kb * 2
    plineObj.SetBulge 3, -kb * 2
    plineObj.SetBulge 4, kb * 2
    plineObj.SetBulge 6, kb * 2
    plineObj.SetBulge 7, -kb * 2
    plineObj.SetBulge 8, -kb * 2
    plineObj.SetBulge 9, kb * 2
    
    plineObj.Layer = "C-Mill"
    plineObj.Update
    plineObj.Closed = True

End If


  pointshelpb1(0) = points2(0) + 70:      pointshelpb1(1) = points2(3) - 73
  pointshelpb1(2) = points2(0) + (b / 2): pointshelpb1(3) = points2(3) - 49
  pointshelpb2(0) = points2(0) + 70:      pointshelpb2(1) = points2(3) - 94 + radiusb
  pointshelpb2(2) = pointswithin(4):      pointshelpb2(3) = pointswithin(5)
Set plineObjw1 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointshelpb1)
Set plineObjw2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointshelpb2)
intPointsb = plineObjw1.IntersectWith(plineObjw2, acExtendBoth)
Intersectionbx = intPointsb(0) - points2(0)
Intersectionby = points2(3) - intPointsb(1)
plineObjw1.Delete
plineObjw2.Delete
  
  pointswithin2(0) = points2(0) + 49:                              pointswithin2(1) = points2(1) + 73
  pointswithin2(2) = points2(0) + 49:                              pointswithin2(3) = points2(3) - 73
  pointswithin2(4) = points2(0) + 70:                              pointswithin2(5) = points2(3) - 73
  pointswithin2(6) = points2(0) + Intersectionbx:                  pointswithin2(7) = points2(3) - Intersectionby
  pointswithin2(8) = points2(0) + (b / 2):                         pointswithin2(9) = points2(3) - 49
  pointswithin2(10) = points2(4) - Intersectionbx:                 pointswithin2(11) = points2(3) - Intersectionby
  pointswithin2(12) = points2(4) - 70:                             pointswithin2(13) = points2(3) - 73
  pointswithin2(14) = points2(4) - 49:                             pointswithin2(15) = points2(3) - 73
  pointswithin2(16) = points2(4) - 49:                             pointswithin2(17) = points2(1) + 73
  pointswithin2(18) = points2(4) - 70:                             pointswithin2(19) = points2(1) + 73
  pointswithin2(20) = points2(4) - Intersectionbx:                 pointswithin2(21) = points2(1) + Intersectionby
  pointswithin2(22) = points2(4) - (b / 2):                        pointswithin2(23) = points2(1) + 49
  pointswithin2(24) = points2(0) + Intersectionbx:                 pointswithin2(25) = points2(1) + Intersectionby
  pointswithin2(26) = points2(0) + 70:                             pointswithin2(27) = points2(1) + 73


If a > 400 Then

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointswithin2)
  plineObj.Closed = True
  
    plineObj.SetBulge 2, kb * 2
    plineObj.SetBulge 3, -kb * 2
    plineObj.SetBulge 4, -kb * 2
    plineObj.SetBulge 5, kb * 2
    plineObj.SetBulge 9, kb * 2
    plineObj.SetBulge 10, -kb * 2
    plineObj.SetBulge 11, -kb * 2
    plineObj.SetBulge 12, kb * 2

    plineObj.Layer = "K-grav"
    plineObj.Update
    plineObj.Closed = True
  
End If

pointsaktr1(0) = points2(0) + 117:      pointsaktr1(1) = points2(1) + 20
pointsaktr1(2) = points2(0) + 92.6462:   pointsaktr1(3) = points2(1) + 20
pointsaktr1(4) = points2(0) + 92.3323:   pointsaktr1(5) = points2(1) + 19.9806
pointsaktr1(6) = points2(0) + 91.3477:   pointsaktr1(7) = points2(1) + 19.779
pointsaktr1(8) = points2(0) + 91.1075:   pointsaktr1(9) = points2(1) + 20
pointsaktr1(10) = points2(0) + 93.4262:  pointsaktr1(11) = points2(1) + 23.7281
pointsaktr1(12) = points2(0) + 94.2774:  pointsaktr1(13) = points2(1) + 24.9179
pointsaktr1(14) = points2(0) + 110.71:   pointsaktr1(15) = points2(1) + 31.0138
pointsaktr1(16) = points2(0) + 89.6442:  pointsaktr1(17) = points2(1) + 30.8565
pointsaktr1(18) = points2(0) + 73.8143:  pointsaktr1(19) = points2(1) + 22.1749
pointsaktr1(20) = points2(0) + 59.8973:  pointsaktr1(21) = points2(1) + 15.5509
pointsaktr1(22) = points2(0) + 46.2255:  pointsaktr1(23) = points2(1) + 17.2859
pointsaktr1(24) = points2(0) + 36.6857:  pointsaktr1(25) = points2(1) + 32.3791
pointsaktr1(26) = points2(0) + 30.9681:  pointsaktr1(27) = points2(1) + 39.9336
pointsaktr1(28) = points2(0) + 26.2137:  pointsaktr1(29) = points2(1) + 39.1948
pointsaktr1(30) = points2(0) + 24.0654:  pointsaktr1(31) = points2(1) + 39.2522
pointsaktr1(32) = points2(0) + 25.0128:  pointsaktr1(33) = points2(1) + 45.3756
pointsaktr1(34) = points2(0) + 22.5192:  pointsaktr1(35) = points2(1) + 47.4951
pointsaktr1(36) = points2(0) + 20.0504:  pointsaktr1(37) = points2(1) + 47.0064
pointsaktr1(38) = points2(0) + 17.4885:  pointsaktr1(39) = points2(1) + 45.2069
pointsaktr1(40) = points2(0) + 17.0747:  pointsaktr1(41) = points2(1) + 45.7786
pointsaktr1(42) = points2(0) + 15.6084:  pointsaktr1(43) = points2(1) + 47.1214
pointsaktr1(44) = points2(0) + 15:       pointsaktr1(45) = points2(1) + 47.9324
pointsaktr1(46) = points2(0) + 19.7914:  pointsaktr1(47) = points2(1) + 49.0263
pointsaktr1(48) = points2(0) + 24.321:   pointsaktr1(49) = points2(1) + 47.4196
pointsaktr1(50) = points2(0) + 26.4955:  pointsaktr1(51) = points2(1) + 44.2133
pointsaktr1(52) = points2(0) + 27.7845:  pointsaktr1(53) = points2(1) + 42.3558
pointsaktr1(54) = points2(0) + 30.1043:  pointsaktr1(55) = points2(1) + 42.1403
pointsaktr1(56) = points2(0) + 29.1547:  pointsaktr1(57) = points2(1) + 49.4473
pointsaktr1(58) = points2(0) + 29.6646:  pointsaktr1(59) = points2(1) + 52.8799
pointsaktr1(60) = points2(0) + 34.0454:  pointsaktr1(61) = points2(1) + 50.5427
pointsaktr1(62) = points2(0) + 34.5844:  pointsaktr1(63) = points2(1) + 43.5484
pointsaktr1(64) = points2(0) + 32.4493:  pointsaktr1(65) = points2(1) + 40.615
pointsaktr1(66) = points2(0) + 36.7127:  pointsaktr1(67) = points2(1) + 34.7043
pointsaktr1(68) = points2(0) + 38.4184:  pointsaktr1(69) = points2(1) + 45.122
pointsaktr1(70) = points2(0) + 40.9074:  pointsaktr1(71) = points2(1) + 58.1004
pointsaktr1(72) = points2(0) + 40.6569:  pointsaktr1(73) = points2(1) + 65.7011
pointsaktr1(74) = points2(0) + 35.0593:  pointsaktr1(75) = points2(1) + 58.0353
pointsaktr1(76) = points2(0) + 34.7122:  pointsaktr1(77) = points2(1) + 57.1492
pointsaktr1(78) = points2(0) + 34.6115:  pointsaktr1(79) = points2(1) + 57.1459
pointsaktr1(80) = points2(0) + 34.5689:  pointsaktr1(81) = points2(1) + 57.2598
pointsaktr1(82) = points2(0) + 34.491:   pointsaktr1(83) = points2(1) + 57.6791
pointsaktr1(84) = points2(0) + 30.7607:  pointsaktr1(85) = points2(1) + 72.4992
pointsaktr1(86) = points2(0) + 29.4442:  pointsaktr1(87) = points2(1) + 81.1508
pointsaktr1(88) = points2(0) + 29.6822:  pointsaktr1(89) = points2(1) + 84.3374
pointsaktr1(90) = points2(0) + 30.3315:  pointsaktr1(91) = points2(1) + 85.2978
pointsaktr1(92) = points2(0) + 30.8867:  pointsaktr1(93) = points2(1) + 85.4656
pointsaktr1(94) = points2(0) + 33.0692:  pointsaktr1(95) = points2(1) + 84.9308
pointsaktr1(96) = points2(0) + 35:       pointsaktr1(97) = points2(1) + 83.2755
pointsaktr1(98) = points2(0) + 35:       pointsaktr1(99) = points2(1) + 88
pointsaktr1(100) = points2(0) + 35:     pointsaktr1(101) = points2(3) - 88
pointsaktr1(102) = points2(0) + 35:     pointsaktr1(103) = points2(3) - 83.2755
pointsaktr1(104) = points2(0) + 33.0692: pointsaktr1(105) = points2(3) - 84.9308
pointsaktr1(106) = points2(0) + 30.8867: pointsaktr1(107) = points2(3) - 85.4656
pointsaktr1(108) = points2(0) + 30.3315: pointsaktr1(109) = points2(3) - 85.2978
pointsaktr1(110) = points2(0) + 29.6822: pointsaktr1(111) = points2(3) - 84.3374
pointsaktr1(112) = points2(0) + 29.4442: pointsaktr1(113) = points2(3) - 81.1508
pointsaktr1(114) = points2(0) + 30.7607: pointsaktr1(115) = points2(3) - 72.4992
pointsaktr1(116) = points2(0) + 34.491:  pointsaktr1(117) = points2(3) - 57.6791
pointsaktr1(118) = points2(0) + 34.5689: pointsaktr1(119) = points2(3) - 57.2598
pointsaktr1(120) = points2(0) + 34.6115: pointsaktr1(121) = points2(3) - 57.1459
pointsaktr1(122) = points2(0) + 34.7122: pointsaktr1(123) = points2(3) - 57.1492
pointsaktr1(124) = points2(0) + 35.0593: pointsaktr1(125) = points2(3) - 58.0353
pointsaktr1(126) = points2(0) + 40.6569: pointsaktr1(127) = points2(3) - 65.7011
pointsaktr1(128) = points2(0) + 40.9074: pointsaktr1(129) = points2(3) - 58.1004
pointsaktr1(130) = points2(0) + 38.4184: pointsaktr1(131) = points2(3) - 45.122
pointsaktr1(132) = points2(0) + 36.7127: pointsaktr1(133) = points2(3) - 34.7043
pointsaktr1(134) = points2(0) + 32.4493: pointsaktr1(135) = points2(3) - 40.615
pointsaktr1(136) = points2(0) + 34.5844: pointsaktr1(137) = points2(3) - 43.5484
pointsaktr1(138) = points2(0) + 34.0454: pointsaktr1(139) = points2(3) - 50.5427
pointsaktr1(140) = points2(0) + 29.6646: pointsaktr1(141) = points2(3) - 52.8799
pointsaktr1(142) = points2(0) + 29.1547: pointsaktr1(143) = points2(3) - 49.4473
pointsaktr1(144) = points2(0) + 30.1043: pointsaktr1(145) = points2(3) - 42.1403
pointsaktr1(146) = points2(0) + 27.7845: pointsaktr1(147) = points2(3) - 42.3558
pointsaktr1(148) = points2(0) + 26.4955: pointsaktr1(149) = points2(3) - 44.2133
pointsaktr1(150) = points2(0) + 24.321:  pointsaktr1(151) = points2(3) - 47.4196
pointsaktr1(152) = points2(0) + 19.7914: pointsaktr1(153) = points2(3) - 49.0263
pointsaktr1(154) = points2(0) + 15:     pointsaktr1(155) = points2(3) - 47.9324
pointsaktr1(156) = points2(0) + 15.6084: pointsaktr1(157) = points2(3) - 47.1214
pointsaktr1(158) = points2(0) + 17.0747: pointsaktr1(159) = points2(3) - 45.7786
pointsaktr1(160) = points2(0) + 17.4885: pointsaktr1(161) = points2(3) - 45.2069
pointsaktr1(162) = points2(0) + 20.0504: pointsaktr1(163) = points2(3) - 47.0064
pointsaktr1(164) = points2(0) + 22.5192: pointsaktr1(165) = points2(3) - 47.4951
pointsaktr1(166) = points2(0) + 25.0128: pointsaktr1(167) = points2(3) - 45.3756
pointsaktr1(168) = points2(0) + 24.0654: pointsaktr1(169) = points2(3) - 39.2522
pointsaktr1(170) = points2(0) + 26.2137: pointsaktr1(171) = points2(3) - 39.1948
pointsaktr1(172) = points2(0) + 30.9681: pointsaktr1(173) = points2(3) - 39.9336
pointsaktr1(174) = points2(0) + 36.6857: pointsaktr1(175) = points2(3) - 32.3791
pointsaktr1(176) = points2(0) + 46.2255: pointsaktr1(177) = points2(3) - 17.2859
pointsaktr1(178) = points2(0) + 59.8973: pointsaktr1(179) = points2(3) - 15.5509
pointsaktr1(180) = points2(0) + 73.8143: pointsaktr1(181) = points2(3) - 22.1749
pointsaktr1(182) = points2(0) + 89.6442: pointsaktr1(183) = points2(3) - 30.8565
pointsaktr1(184) = points2(0) + 110.71:  pointsaktr1(185) = points2(3) - 31.0138
pointsaktr1(186) = points2(0) + 94.2774: pointsaktr1(187) = points2(3) - 24.9179
pointsaktr1(188) = points2(0) + 93.4262: pointsaktr1(189) = points2(3) - 23.7281
pointsaktr1(190) = points2(0) + 91.1075: pointsaktr1(191) = points2(3) - 20
pointsaktr1(192) = points2(0) + 91.3477: pointsaktr1(193) = points2(3) - 19.779
pointsaktr1(194) = points2(0) + 92.3323: pointsaktr1(195) = points2(3) - 19.9806
pointsaktr1(196) = points2(0) + 92.6462: pointsaktr1(197) = points2(3) - 20
pointsaktr1(198) = points2(0) + 117:    pointsaktr1(199) = points2(3) - 20
pointsaktr1(200) = points2(4) - 117:     pointsaktr1(201) = points2(3) - 20
pointsaktr1(202) = points2(4) - 92.6462: pointsaktr1(203) = points2(3) - 20
pointsaktr1(204) = points2(4) - 92.3323: pointsaktr1(205) = points2(3) - 19.9806
pointsaktr1(206) = points2(4) - 91.3477: pointsaktr1(207) = points2(3) - 19.779
pointsaktr1(208) = points2(4) - 91.1075: pointsaktr1(209) = points2(3) - 20
pointsaktr1(210) = points2(4) - 93.4262: pointsaktr1(211) = points2(3) - 23.7281
pointsaktr1(212) = points2(4) - 94.2774: pointsaktr1(213) = points2(3) - 24.9179
pointsaktr1(214) = points2(4) - 110.71:  pointsaktr1(215) = points2(3) - 31.0138
pointsaktr1(216) = points2(4) - 89.6442: pointsaktr1(217) = points2(3) - 30.8565
pointsaktr1(218) = points2(4) - 73.8143: pointsaktr1(219) = points2(3) - 22.1749
pointsaktr1(220) = points2(4) - 59.8973: pointsaktr1(221) = points2(3) - 15.5509
pointsaktr1(222) = points2(4) - 46.2255: pointsaktr1(223) = points2(3) - 17.2859
pointsaktr1(224) = points2(4) - 36.6857: pointsaktr1(225) = points2(3) - 32.3791
pointsaktr1(226) = points2(4) - 30.9681: pointsaktr1(227) = points2(3) - 39.9336
pointsaktr1(228) = points2(4) - 26.2137: pointsaktr1(229) = points2(3) - 39.1948
pointsaktr1(230) = points2(4) - 24.0654: pointsaktr1(231) = points2(3) - 39.2522
pointsaktr1(232) = points2(4) - 25.0128: pointsaktr1(233) = points2(3) - 45.3756
pointsaktr1(234) = points2(4) - 22.5192: pointsaktr1(235) = points2(3) - 47.4951
pointsaktr1(236) = points2(4) - 20.0504: pointsaktr1(237) = points2(3) - 47.0064
pointsaktr1(238) = points2(4) - 17.4885: pointsaktr1(239) = points2(3) - 45.2069
pointsaktr1(240) = points2(4) - 17.0747: pointsaktr1(241) = points2(3) - 45.7786
pointsaktr1(242) = points2(4) - 15.6084: pointsaktr1(243) = points2(3) - 47.1214
pointsaktr1(244) = points2(4) - 15:      pointsaktr1(245) = points2(3) - 47.9324
pointsaktr1(246) = points2(4) - 19.7914: pointsaktr1(247) = points2(3) - 49.0263
pointsaktr1(248) = points2(4) - 24.321:  pointsaktr1(249) = points2(3) - 47.4196
pointsaktr1(250) = points2(4) - 26.4955: pointsaktr1(251) = points2(3) - 44.2133
pointsaktr1(252) = points2(4) - 27.7845: pointsaktr1(253) = points2(3) - 42.3558
pointsaktr1(254) = points2(4) - 30.1043: pointsaktr1(255) = points2(3) - 42.1403
pointsaktr1(256) = points2(4) - 29.1547: pointsaktr1(257) = points2(3) - 49.4473
pointsaktr1(258) = points2(4) - 29.6646: pointsaktr1(259) = points2(3) - 52.8799
pointsaktr1(260) = points2(4) - 34.0454: pointsaktr1(261) = points2(3) - 50.5427
pointsaktr1(262) = points2(4) - 34.5844: pointsaktr1(263) = points2(3) - 43.5484
pointsaktr1(264) = points2(4) - 32.4493: pointsaktr1(265) = points2(3) - 40.615
pointsaktr1(266) = points2(4) - 36.7127: pointsaktr1(267) = points2(3) - 34.7043
pointsaktr1(268) = points2(4) - 38.4184: pointsaktr1(269) = points2(3) - 45.122
pointsaktr1(270) = points2(4) - 40.9074: pointsaktr1(271) = points2(3) - 58.1004
pointsaktr1(272) = points2(4) - 40.6569: pointsaktr1(273) = points2(3) - 65.7011
pointsaktr1(274) = points2(4) - 35.0593: pointsaktr1(275) = points2(3) - 58.0353
pointsaktr1(276) = points2(4) - 34.7122: pointsaktr1(277) = points2(3) - 57.1492
pointsaktr1(278) = points2(4) - 34.6115: pointsaktr1(279) = points2(3) - 57.1459
pointsaktr1(280) = points2(4) - 34.5689: pointsaktr1(281) = points2(3) - 57.2598
pointsaktr1(282) = points2(4) - 34.491:  pointsaktr1(283) = points2(3) - 57.6791
pointsaktr1(284) = points2(4) - 30.7607: pointsaktr1(285) = points2(3) - 72.4992
pointsaktr1(286) = points2(4) - 29.4442: pointsaktr1(287) = points2(3) - 81.1508
pointsaktr1(288) = points2(4) - 29.6822: pointsaktr1(289) = points2(3) - 84.3374
pointsaktr1(290) = points2(4) - 30.3315: pointsaktr1(291) = points2(3) - 85.2978
pointsaktr1(292) = points2(4) - 30.8867: pointsaktr1(293) = points2(3) - 85.4656
pointsaktr1(294) = points2(4) - 33.0692: pointsaktr1(295) = points2(3) - 84.9308
pointsaktr1(296) = points2(4) - 35:      pointsaktr1(297) = points2(3) - 83.2755
pointsaktr1(298) = points2(4) - 35:      pointsaktr1(299) = points2(3) - 88
pointsaktr1(300) = points2(4) - 35:     pointsaktr1(301) = points2(1) + 88
pointsaktr1(302) = points2(4) - 35:     pointsaktr1(303) = points2(1) + 83.2755
pointsaktr1(304) = points2(4) - 33.0692: pointsaktr1(305) = points2(1) + 84.9308
pointsaktr1(306) = points2(4) - 30.8867: pointsaktr1(307) = points2(1) + 85.4656
pointsaktr1(308) = points2(4) - 30.3315: pointsaktr1(309) = points2(1) + 85.2978
pointsaktr1(310) = points2(4) - 29.6822: pointsaktr1(311) = points2(1) + 84.3374
pointsaktr1(312) = points2(4) - 29.4442: pointsaktr1(313) = points2(1) + 81.1508
pointsaktr1(314) = points2(4) - 30.7607: pointsaktr1(315) = points2(1) + 72.4992
pointsaktr1(316) = points2(4) - 34.491:  pointsaktr1(317) = points2(1) + 57.6791
pointsaktr1(318) = points2(4) - 34.5689: pointsaktr1(319) = points2(1) + 57.2598
pointsaktr1(320) = points2(4) - 34.6115: pointsaktr1(321) = points2(1) + 57.1459
pointsaktr1(322) = points2(4) - 34.7122: pointsaktr1(323) = points2(1) + 57.1492
pointsaktr1(324) = points2(4) - 35.0593: pointsaktr1(325) = points2(1) + 58.0353
pointsaktr1(326) = points2(4) - 40.6569: pointsaktr1(327) = points2(1) + 65.7011
pointsaktr1(328) = points2(4) - 40.9074: pointsaktr1(329) = points2(1) + 58.1004
pointsaktr1(330) = points2(4) - 38.4184: pointsaktr1(331) = points2(1) + 45.122
pointsaktr1(332) = points2(4) - 36.7127: pointsaktr1(333) = points2(1) + 34.7043
pointsaktr1(334) = points2(4) - 32.4493: pointsaktr1(335) = points2(1) + 40.615
pointsaktr1(336) = points2(4) - 34.5844: pointsaktr1(337) = points2(1) + 43.5484
pointsaktr1(338) = points2(4) - 34.0454: pointsaktr1(339) = points2(1) + 50.5427
pointsaktr1(340) = points2(4) - 29.6646: pointsaktr1(341) = points2(1) + 52.8799
pointsaktr1(342) = points2(4) - 29.1547: pointsaktr1(343) = points2(1) + 49.4473
pointsaktr1(344) = points2(4) - 30.1043: pointsaktr1(345) = points2(1) + 42.1403
pointsaktr1(346) = points2(4) - 27.7845: pointsaktr1(347) = points2(1) + 42.3558
pointsaktr1(348) = points2(4) - 26.4955: pointsaktr1(349) = points2(1) + 44.2133
pointsaktr1(350) = points2(4) - 24.321:  pointsaktr1(351) = points2(1) + 47.4196
pointsaktr1(352) = points2(4) - 19.7914: pointsaktr1(353) = points2(1) + 49.0263
pointsaktr1(354) = points2(4) - 15:     pointsaktr1(355) = points2(1) + 47.9324
pointsaktr1(356) = points2(4) - 15.6084: pointsaktr1(357) = points2(1) + 47.1214
pointsaktr1(358) = points2(4) - 17.0747: pointsaktr1(359) = points2(1) + 45.7786
pointsaktr1(360) = points2(4) - 17.4885: pointsaktr1(361) = points2(1) + 45.2069
pointsaktr1(362) = points2(4) - 20.0504: pointsaktr1(363) = points2(1) + 47.0064
pointsaktr1(364) = points2(4) - 22.5192: pointsaktr1(365) = points2(1) + 47.4951
pointsaktr1(366) = points2(4) - 25.0128: pointsaktr1(367) = points2(1) + 45.3756
pointsaktr1(368) = points2(4) - 24.0654: pointsaktr1(369) = points2(1) + 39.2522
pointsaktr1(370) = points2(4) - 26.2137: pointsaktr1(371) = points2(1) + 39.1948
pointsaktr1(372) = points2(4) - 30.9681: pointsaktr1(373) = points2(1) + 39.9336
pointsaktr1(374) = points2(4) - 36.6857: pointsaktr1(375) = points2(1) + 32.3791
pointsaktr1(376) = points2(4) - 46.2255: pointsaktr1(377) = points2(1) + 17.2859
pointsaktr1(378) = points2(4) - 59.8973: pointsaktr1(379) = points2(1) + 15.5509
pointsaktr1(380) = points2(4) - 73.8143: pointsaktr1(381) = points2(1) + 22.1749
pointsaktr1(382) = points2(4) - 89.6442: pointsaktr1(383) = points2(1) + 30.8565
pointsaktr1(384) = points2(4) - 110.71:  pointsaktr1(385) = points2(1) + 31.0138
pointsaktr1(386) = points2(4) - 94.2774: pointsaktr1(387) = points2(1) + 24.9179
pointsaktr1(388) = points2(4) - 93.4262: pointsaktr1(389) = points2(1) + 23.7281
pointsaktr1(390) = points2(4) - 91.1075: pointsaktr1(391) = points2(1) + 20
pointsaktr1(392) = points2(4) - 91.3477: pointsaktr1(393) = points2(1) + 19.779
pointsaktr1(394) = points2(4) - 92.3323: pointsaktr1(395) = points2(1) + 19.9806
pointsaktr1(396) = points2(4) - 92.6462: pointsaktr1(397) = points2(1) + 20
pointsaktr1(398) = points2(4) - 117:    pointsaktr1(399) = points2(1) + 20



Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsaktr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update
    
currentBulge = plineObj.GetBulge(2)

plineObj.SetBulge 0, 0
plineObj.SetBulge 1, 0.0177245
plineObj.SetBulge 2, 0.0523607
plineObj.SetBulge 3, -0.57942
plineObj.SetBulge 4, -0.168144
plineObj.SetBulge 5, 0.135111
plineObj.SetBulge 6, 0.269875
plineObj.SetBulge 7, 0.165606
plineObj.SetBulge 8, 0.0831599
plineObj.SetBulge 9, -0.112133
plineObj.SetBulge 10, -0.175336
plineObj.SetBulge 11, -0.273401
plineObj.SetBulge 12, -0.105483
plineObj.SetBulge 13, -0.135712
plineObj.SetBulge 14, 0.0444926
plineObj.SetBulge 15, 0.25299
plineObj.SetBulge 16, 0.2683
plineObj.SetBulge 17, 0.190049
plineObj.SetBulge 18, 0.0206626
plineObj.SetBulge 19, 0.0702883
plineObj.SetBulge 20, 0.0312334
plineObj.SetBulge 21, -0.124693
plineObj.SetBulge 22, -0.111948
plineObj.SetBulge 23, -0.172878
plineObj.SetBulge 24, -0.146843
plineObj.SetBulge 25, 0.152321
plineObj.SetBulge 26, 0.292528
plineObj.SetBulge 27, -0.0613614
plineObj.SetBulge 28, -0.0772198
plineObj.SetBulge 29, -0.322171
plineObj.SetBulge 30, -0.192556
plineObj.SetBulge 31, -0.164263
plineObj.SetBulge 32, 0.0753202
plineObj.SetBulge 33, -0.0424071
plineObj.SetBulge 34, 0.0288023
plineObj.SetBulge 35, 0.0825991
plineObj.SetBulge 36, 0.228651
plineObj.SetBulge 37, -0.0964151
plineObj.SetBulge 38, -0.528657
plineObj.SetBulge 39, -0.137223
plineObj.SetBulge 40, 0.049294
plineObj.SetBulge 41, -0.0178489
plineObj.SetBulge 42, -0.0299484
plineObj.SetBulge 43, -0.0830332
plineObj.SetBulge 44, -0.178953
plineObj.SetBulge 45, -0.165905
plineObj.SetBulge 46, -0.102816
plineObj.SetBulge 47, -0.132516
plineObj.SetBulge 48, 0
plineObj.SetBulge 49, 0
plineObj.SetBulge 50, 0
plineObj.SetBulge 51, -0.132516
plineObj.SetBulge 52, -0.102816
plineObj.SetBulge 53, -0.165905
plineObj.SetBulge 54, -0.178953
plineObj.SetBulge 55, -0.0830332
plineObj.SetBulge 56, -0.0299484
plineObj.SetBulge 57, -0.0178489
plineObj.SetBulge 58, 0.049294
plineObj.SetBulge 59, -0.137223
plineObj.SetBulge 60, -0.528657
plineObj.SetBulge 61, -0.0964151
plineObj.SetBulge 62, 0.228651
plineObj.SetBulge 63, 0.0825991
plineObj.SetBulge 64, 0.0288023
plineObj.SetBulge 65, -0.0424071
plineObj.SetBulge 66, 0.0753202
plineObj.SetBulge 67, -0.164263
plineObj.SetBulge 68, -0.192556
plineObj.SetBulge 69, -0.322171
plineObj.SetBulge 70, -0.0772198
plineObj.SetBulge 71, -0.0613614
plineObj.SetBulge 72, 0.292528
plineObj.SetBulge 73, 0.152321
plineObj.SetBulge 74, -0.146843
plineObj.SetBulge 75, -0.172878
plineObj.SetBulge 76, -0.111948
plineObj.SetBulge 77, -0.124693
plineObj.SetBulge 78, 0.0312334
plineObj.SetBulge 79, 0.0702883
plineObj.SetBulge 80, 0.0206626
plineObj.SetBulge 81, 0.190049
plineObj.SetBulge 82, 0.2683
plineObj.SetBulge 83, 0.25299
plineObj.SetBulge 84, 0.0444926
plineObj.SetBulge 85, -0.135712
plineObj.SetBulge 86, -0.105483
plineObj.SetBulge 87, -0.273401
plineObj.SetBulge 88, -0.175336
plineObj.SetBulge 89, -0.112133
plineObj.SetBulge 90, 0.0831599
plineObj.SetBulge 91, 0.165606
plineObj.SetBulge 92, 0.269875
plineObj.SetBulge 93, 0.135111
plineObj.SetBulge 94, -0.168144
plineObj.SetBulge 95, -0.57942
plineObj.SetBulge 96, 0.0523607
plineObj.SetBulge 97, 0.0177245
plineObj.SetBulge 98, 0
plineObj.SetBulge 99, 0
plineObj.SetBulge 100, 0
plineObj.SetBulge 101, 0.0177245
plineObj.SetBulge 102, 0.0523607
plineObj.SetBulge 103, -0.57942
plineObj.SetBulge 104, -0.168144
plineObj.SetBulge 105, 0.135111
plineObj.SetBulge 106, 0.269875
plineObj.SetBulge 107, 0.165606
plineObj.SetBulge 108, 0.0831599
plineObj.SetBulge 109, -0.112133
plineObj.SetBulge 110, -0.175336
plineObj.SetBulge 111, -0.273401
plineObj.SetBulge 112, -0.105483
plineObj.SetBulge 113, -0.135712
plineObj.SetBulge 114, 0.0444926
plineObj.SetBulge 115, 0.25299
plineObj.SetBulge 116, 0.2683
plineObj.SetBulge 117, 0.190049
plineObj.SetBulge 118, 0.0206626
plineObj.SetBulge 119, 0.0702883
plineObj.SetBulge 120, 0.0312334
plineObj.SetBulge 121, -0.124693
plineObj.SetBulge 122, -0.111948
plineObj.SetBulge 123, -0.172878
plineObj.SetBulge 124, -0.146843
plineObj.SetBulge 125, 0.152321
plineObj.SetBulge 126, 0.292528
plineObj.SetBulge 127, -0.0613614
plineObj.SetBulge 128, -0.0772198
plineObj.SetBulge 129, -0.322171
plineObj.SetBulge 130, -0.192556
plineObj.SetBulge 131, -0.164263
plineObj.SetBulge 132, 0.0753202
plineObj.SetBulge 133, -0.0424071
plineObj.SetBulge 134, 0.0288023
plineObj.SetBulge 135, 0.0825991
plineObj.SetBulge 136, 0.228651
plineObj.SetBulge 137, -0.0964151
plineObj.SetBulge 138, -0.528657
plineObj.SetBulge 139, -0.137223
plineObj.SetBulge 140, 0.049294
plineObj.SetBulge 141, -0.0178489
plineObj.SetBulge 142, -0.0299484
plineObj.SetBulge 143, -0.0830332
plineObj.SetBulge 144, -0.178953
plineObj.SetBulge 145, -0.165905
plineObj.SetBulge 146, -0.102816
plineObj.SetBulge 147, -0.132516
plineObj.SetBulge 148, 0
plineObj.SetBulge 149, 0
plineObj.SetBulge 150, 0
plineObj.SetBulge 151, -0.132516
plineObj.SetBulge 152, -0.102816
plineObj.SetBulge 153, -0.165905
plineObj.SetBulge 154, -0.178953
plineObj.SetBulge 155, -0.0830332
plineObj.SetBulge 156, -0.0299484
plineObj.SetBulge 157, -0.0178489
plineObj.SetBulge 158, 0.049294
plineObj.SetBulge 159, -0.137223
plineObj.SetBulge 160, -0.528657
plineObj.SetBulge 161, -0.0964151
plineObj.SetBulge 162, 0.228651
plineObj.SetBulge 163, 0.0825991
plineObj.SetBulge 164, 0.0288023
plineObj.SetBulge 165, -0.0424071
plineObj.SetBulge 166, 0.0753202
plineObj.SetBulge 167, -0.164263
plineObj.SetBulge 168, -0.192556
plineObj.SetBulge 169, -0.322171
plineObj.SetBulge 170, -0.0772198
plineObj.SetBulge 171, -0.0613614
plineObj.SetBulge 172, 0.292528
plineObj.SetBulge 173, 0.152321
plineObj.SetBulge 174, -0.146843
plineObj.SetBulge 175, -0.172878
plineObj.SetBulge 176, -0.111948
plineObj.SetBulge 177, -0.124693
plineObj.SetBulge 178, 0.0312334
plineObj.SetBulge 179, 0.0702883
plineObj.SetBulge 180, 0.0206626
plineObj.SetBulge 181, 0.190049
plineObj.SetBulge 182, 0.2683
plineObj.SetBulge 183, 0.25299
plineObj.SetBulge 184, 0.0444926
plineObj.SetBulge 185, -0.135712
plineObj.SetBulge 186, -0.105483
plineObj.SetBulge 187, -0.273401
plineObj.SetBulge 188, -0.175336
plineObj.SetBulge 189, -0.112133
plineObj.SetBulge 190, 0.0831599
plineObj.SetBulge 191, 0.165606
plineObj.SetBulge 192, 0.269875
plineObj.SetBulge 193, 0.135111
plineObj.SetBulge 194, -0.168144
plineObj.SetBulge 195, -0.57942
plineObj.SetBulge 196, 0.0523607
plineObj.SetBulge 197, 0.0177245
plineObj.SetBulge 198, 0
plineObj.SetBulge 199, 0
plineObj.SetBulge 200, 0

    
    plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True
  

pointsaktr2(0) = points2(0) + 117:       pointsaktr2(1) = points2(1) + 22.5
pointsaktr2(2) = points2(0) + 99.158:    pointsaktr2(3) = points2(1) + 22.5
pointsaktr2(4) = points2(0) + 110.917:   pointsaktr2(5) = points2(1) + 24.1929
pointsaktr2(6) = points2(0) + 114.984:   pointsaktr2(7) = points2(1) + 26.2407
pointsaktr2(8) = points2(0) + 115.489:   pointsaktr2(9) = points2(1) + 28.9732
pointsaktr2(10) = points2(0) + 112.864:  pointsaktr2(11) = points2(1) + 31.6992
pointsaktr2(12) = points2(0) + 115.371:  pointsaktr2(13) = points2(1) + 49.2106
pointsaktr2(14) = points2(0) + 113.001:  pointsaktr2(15) = points2(1) + 49.5365
pointsaktr2(16) = points2(0) + 111.656:  pointsaktr2(17) = points2(1) + 32.2762
pointsaktr2(18) = points2(0) + 89.15:    pointsaktr2(19) = points2(1) + 32.2728
pointsaktr2(20) = points2(0) + 72.8856:  pointsaktr2(21) = points2(1) + 23.3529
pointsaktr2(22) = points2(0) + 59.5687:  pointsaktr2(23) = points2(1) + 17.0144
pointsaktr2(24) = points2(0) + 46.9093:  pointsaktr2(25) = points2(1) + 18.621
pointsaktr2(26) = points2(0) + 38.268:   pointsaktr2(27) = points2(1) + 31.2047
pointsaktr2(28) = points2(0) + 50.5636:  pointsaktr2(29) = points2(1) + 28.0599
pointsaktr2(30) = points2(0) + 49.9818:  pointsaktr2(31) = points2(1) + 69.0067
pointsaktr2(32) = points2(0) + 41.8998:  pointsaktr2(33) = points2(1) + 67.5444
pointsaktr2(34) = points2(0) + 37:       pointsaktr2(35) = points2(1) + 80.4314
pointsaktr2(36) = points2(0) + 37:       pointsaktr2(37) = points2(1) + 88
pointsaktr2(38) = points2(0) + 37:       pointsaktr2(39) = points2(3) - 88#
pointsaktr2(40) = points2(0) + 37:       pointsaktr2(41) = points2(3) - 80.4314
pointsaktr2(42) = points2(0) + 41.8998:  pointsaktr2(43) = points2(3) - 67.5444
pointsaktr2(44) = points2(0) + 49.9818:  pointsaktr2(45) = points2(3) - 69.0067
pointsaktr2(46) = points2(0) + 50.5636:  pointsaktr2(47) = points2(3) - 28.0599
pointsaktr2(48) = points2(0) + 38.268:   pointsaktr2(49) = points2(3) - 31.2047
pointsaktr2(50) = points2(0) + 46.9093:  pointsaktr2(51) = points2(3) - 18.621
pointsaktr2(52) = points2(0) + 59.5687:  pointsaktr2(53) = points2(3) - 17.0144
pointsaktr2(54) = points2(0) + 72.8856:  pointsaktr2(55) = points2(3) - 23.3529
pointsaktr2(56) = points2(0) + 89.15:    pointsaktr2(57) = points2(3) - 32.2728
pointsaktr2(58) = points2(0) + 111.656:  pointsaktr2(59) = points2(3) - 32.2762
pointsaktr2(60) = points2(0) + 113.001:  pointsaktr2(61) = points2(3) - 49.5365
pointsaktr2(62) = points2(0) + 115.371:  pointsaktr2(63) = points2(3) - 49.2106
pointsaktr2(64) = points2(0) + 112.864:  pointsaktr2(65) = points2(3) - 31.6992
pointsaktr2(66) = points2(0) + 115.489:  pointsaktr2(67) = points2(3) - 28.9732
pointsaktr2(68) = points2(0) + 114.984:  pointsaktr2(69) = points2(3) - 26.2407
pointsaktr2(70) = points2(0) + 110.917:  pointsaktr2(71) = points2(3) - 24.1929
pointsaktr2(72) = points2(0) + 99.158:   pointsaktr2(73) = points2(3) - 22.5
pointsaktr2(74) = points2(0) + 117:      pointsaktr2(75) = points2(3) - 22.5
pointsaktr2(76) = points2(4) - 117:      pointsaktr2(77) = points2(3) - 22.5
pointsaktr2(78) = points2(4) - 99.158:   pointsaktr2(79) = points2(3) - 22.5
pointsaktr2(80) = points2(4) - 110.917:  pointsaktr2(81) = points2(3) - 24.1929
pointsaktr2(82) = points2(4) - 114.984:  pointsaktr2(83) = points2(3) - 26.2407
pointsaktr2(84) = points2(4) - 115.489:  pointsaktr2(85) = points2(3) - 28.9732
pointsaktr2(86) = points2(4) - 112.864:  pointsaktr2(87) = points2(3) - 31.6992
pointsaktr2(88) = points2(4) - 115.371:  pointsaktr2(89) = points2(3) - 49.2106
pointsaktr2(90) = points2(4) - 113.001:  pointsaktr2(91) = points2(3) - 49.5365
pointsaktr2(92) = points2(4) - 111.656:  pointsaktr2(93) = points2(3) - 32.2762
pointsaktr2(94) = points2(4) - 89.15:    pointsaktr2(95) = points2(3) - 32.2728
pointsaktr2(96) = points2(4) - 72.8856:  pointsaktr2(97) = points2(3) - 23.3529
pointsaktr2(98) = points2(4) - 59.5687:  pointsaktr2(99) = points2(3) - 17.0144
pointsaktr2(100) = points2(4) - 46.9093: pointsaktr2(101) = points2(3) - 18.621
pointsaktr2(102) = points2(4) - 38.268:  pointsaktr2(103) = points2(3) - 31.2047
pointsaktr2(104) = points2(4) - 50.5636: pointsaktr2(105) = points2(3) - 28.0599
pointsaktr2(106) = points2(4) - 49.9818: pointsaktr2(107) = points2(3) - 69.0067
pointsaktr2(108) = points2(4) - 41.8998: pointsaktr2(109) = points2(3) - 67.5444
pointsaktr2(110) = points2(4) - 37:      pointsaktr2(111) = points2(3) - 80.4314
pointsaktr2(112) = points2(4) - 37:      pointsaktr2(113) = points2(3) - 88
pointsaktr2(114) = points2(4) - 37:      pointsaktr2(115) = points2(1) + 88#
pointsaktr2(116) = points2(4) - 37:      pointsaktr2(117) = points2(1) + 80.4314
pointsaktr2(118) = points2(4) - 41.8998: pointsaktr2(119) = points2(1) + 67.5444
pointsaktr2(120) = points2(4) - 49.9818: pointsaktr2(121) = points2(1) + 69.0067
pointsaktr2(122) = points2(4) - 50.5636: pointsaktr2(123) = points2(1) + 28.0599
pointsaktr2(124) = points2(4) - 38.268:  pointsaktr2(125) = points2(1) + 31.2047
pointsaktr2(126) = points2(4) - 46.9093: pointsaktr2(127) = points2(1) + 18.621
pointsaktr2(128) = points2(4) - 59.5687: pointsaktr2(129) = points2(1) + 17.0144
pointsaktr2(130) = points2(4) - 72.8856: pointsaktr2(131) = points2(1) + 23.3529
pointsaktr2(132) = points2(4) - 89.15:   pointsaktr2(133) = points2(1) + 32.2728
pointsaktr2(134) = points2(4) - 111.656: pointsaktr2(135) = points2(1) + 32.2762
pointsaktr2(136) = points2(4) - 113.001: pointsaktr2(137) = points2(1) + 49.5365
pointsaktr2(138) = points2(4) - 115.371: pointsaktr2(139) = points2(1) + 49.2106
pointsaktr2(140) = points2(4) - 112.864: pointsaktr2(141) = points2(1) + 31.6992
pointsaktr2(142) = points2(4) - 115.489: pointsaktr2(143) = points2(1) + 28.9732
pointsaktr2(144) = points2(4) - 114.984: pointsaktr2(145) = points2(1) + 26.2407
pointsaktr2(146) = points2(4) - 110.917: pointsaktr2(147) = points2(1) + 24.1929
pointsaktr2(148) = points2(4) - 99.158:  pointsaktr2(149) = points2(1) + 22.5
pointsaktr2(150) = points2(4) - 117:     pointsaktr2(151) = points2(1) + 22.5

If b < 300 Then
pointsaktr2(13) = points2(1) + 43
pointsaktr2(15) = points2(1) + 43.4
pointsaktr2(61) = points2(3) - 43.4
pointsaktr2(63) = points2(3) - 43
pointsaktr2(89) = points2(3) - 43
pointsaktr2(91) = points2(3) - 43.4
pointsaktr2(137) = points2(1) + 43.4
pointsaktr2(139) = points2(1) + 43
End If

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsaktr2)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update
  
plineObj.SetBulge 0, 0
plineObj.SetBulge 1, 0.0427012
plineObj.SetBulge 2, 0.119624
plineObj.SetBulge 3, 0.355722
plineObj.SetBulge 4, 0.13362
plineObj.SetBulge 5, 0.234279
plineObj.SetBulge 6, 0
plineObj.SetBulge 7, -0.266061
plineObj.SetBulge 8, 0.169365
plineObj.SetBulge 9, 0.0831599
plineObj.SetBulge 10, -0.112133
plineObj.SetBulge 11, -0.175336
plineObj.SetBulge 12, -0.253039
plineObj.SetBulge 13, 0.149933
plineObj.SetBulge 14, 0.967494
plineObj.SetBulge 15, 0.0992359
plineObj.SetBulge 16, 0.0970037
plineObj.SetBulge 17, 0
plineObj.SetBulge 18, 0
plineObj.SetBulge 19, 0
plineObj.SetBulge 20, 0.0970037
plineObj.SetBulge 21, 0.0992359
plineObj.SetBulge 22, 0.967494
plineObj.SetBulge 23, 0.149933
plineObj.SetBulge 24, -0.253039
plineObj.SetBulge 25, -0.175336
plineObj.SetBulge 26, -0.112133
plineObj.SetBulge 27, 0.0831599
plineObj.SetBulge 28, 0.169365
plineObj.SetBulge 29, -0.266061
plineObj.SetBulge 30, 0
plineObj.SetBulge 31, 0.234279
plineObj.SetBulge 32, 0.13362
plineObj.SetBulge 33, 0.355722
plineObj.SetBulge 34, 0.119624
plineObj.SetBulge 35, 0.0427012
plineObj.SetBulge 36, 0
plineObj.SetBulge 37, 0
plineObj.SetBulge 38, 0
plineObj.SetBulge 39, 0.0427012
plineObj.SetBulge 40, 0.119624
plineObj.SetBulge 41, 0.355722
plineObj.SetBulge 42, 0.13362
plineObj.SetBulge 43, 0.234279
plineObj.SetBulge 44, 0
plineObj.SetBulge 45, -0.266061
plineObj.SetBulge 46, 0.169365
plineObj.SetBulge 47, 0.0831599
plineObj.SetBulge 48, -0.112133
plineObj.SetBulge 49, -0.175336
plineObj.SetBulge 50, -0.253039
plineObj.SetBulge 51, 0.149933
plineObj.SetBulge 52, 0.967494
plineObj.SetBulge 53, 0.0992359
plineObj.SetBulge 54, 0.0970037
plineObj.SetBulge 55, 0
plineObj.SetBulge 56, 0
plineObj.SetBulge 57, 0
plineObj.SetBulge 58, 0.0970037
plineObj.SetBulge 59, 0.0992359
plineObj.SetBulge 60, 0.967494
plineObj.SetBulge 61, 0.149933
plineObj.SetBulge 62, -0.253039
plineObj.SetBulge 63, -0.175336
plineObj.SetBulge 64, -0.112133
plineObj.SetBulge 65, 0.0831599
plineObj.SetBulge 66, 0.169365
plineObj.SetBulge 67, -0.266061
plineObj.SetBulge 68, 0
plineObj.SetBulge 69, 0.234279
plineObj.SetBulge 70, 0.13362
plineObj.SetBulge 71, 0.355722
plineObj.SetBulge 72, 0.119624
plineObj.SetBulge 73, 0.0427012
plineObj.SetBulge 74, 0
plineObj.SetBulge 75, 0

 plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True
  
pointsaktr3(0) = points2(0) + 31.6684:   pointsaktr3(1) = points2(1) + 42.8836
pointsaktr3(2) = points2(0) + 30.5655:   pointsaktr3(3) = points2(1) + 48.9538
pointsaktr3(4) = points2(0) + 30.9385:   pointsaktr3(5) = points2(1) + 51.4282
pointsaktr3(6) = points2(0) + 33.3093:   pointsaktr3(7) = points2(1) + 49.2005
pointsaktr3(8) = points2(0) + 33.752:    pointsaktr3(9) = points2(1) + 46.0948


Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsaktr3)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update
  
plineObj.SetBulge 0, -0.0681083
plineObj.SetBulge 1, -0.0969811
plineObj.SetBulge 2, -0.21617
plineObj.SetBulge 3, -0.125224
plineObj.SetBulge 4, -0.238364


 plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True
    
  b1(0) = (b / 2) - 1:  b1(1) = points2(3) - (a / 2)
  b2(0) = (b / 2) + 1:  b2(1) = points2(3) - (a / 2)
  a1(0) = points2(4) - (b / 2): a1(1) = points2(1) + (a / 2) - 1
  a2(0) = points2(4) - (b / 2): a2(1) = points2(1) + (a / 2) + 1
  RetVal = plineObj.Mirror(b1, b2)
  RetVal = plineObj.Mirror(a1, a2)
  RetVal = plineObj.Copy
  ' Define the rotation of 180 degrees about a
  ' base point of (4, 4.25, 0)
  basePoint(0) = points2(4) - (b / 2): basePoint(1) = points2(3) - (a / 2)
  rotationAngle = 3.14159   ' 180 degrees
  plineObj.Rotate basePoint, rotationAngle
  plineObj.Update
  
pointsaktr4(0) = points2(0) + 32.2044:   pointsaktr4(1) = points2(1) + 70.283
pointsaktr4(2) = points2(0) + 34.6713:   pointsaktr4(3) = points2(1) + 64.8485
pointsaktr4(4) = points2(0) + 34.6713:   pointsaktr4(5) = points2(1) + 79.75
pointsaktr4(6) = points2(0) + 32.688:    pointsaktr4(7) = points2(1) + 81.4884
pointsaktr4(8) = points2(0) + 31.48:     pointsaktr4(9) = points2(1) + 82.0684
pointsaktr4(10) = points2(0) + 30.6893:  pointsaktr4(11) = points2(1) + 81.5751
pointsaktr4(12) = points2(0) + 30.6422:  pointsaktr4(13) = points2(1) + 80.7573


Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsaktr4)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update
  
plineObj.SetBulge 0, 0.0818639
plineObj.SetBulge 1, 0
plineObj.SetBulge 2, 0.065715
plineObj.SetBulge 3, 0.0705315
plineObj.SetBulge 4, 0.461421
plineObj.SetBulge 5, 0.0454552
plineObj.SetBulge 6, 0.0574167


RetVal = plineObj.Mirror(a1, a2)
RetVal = plineObj.Mirror(b1, b2)
RetVal = plineObj.Copy
plineObj.Rotate basePoint, rotationAngle

pointsaktr5(0) = points2(0) + 40.4778:   pointsaktr5(1) = points2(1) + 66.8334
pointsaktr5(2) = points2(0) + 36.4488:   pointsaktr5(3) = points2(1) + 63.9005
pointsaktr5(4) = points2(0) + 36.4488:   pointsaktr5(5) = points2(1) + 77.5281


Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsaktr5)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update
  
plineObj.SetBulge 0, 0.0833799
plineObj.SetBulge 1, 0
plineObj.SetBulge 2, -0.0920228


RetVal = plineObj.Mirror(a1, a2)
RetVal = plineObj.Mirror(b1, b2)
RetVal = plineObj.Copy
plineObj.Rotate basePoint, rotationAngle

pointsaktr6(0) = points2(0) + 50.5066:   pointsaktr6(1) = points2(1) + 29.4087
pointsaktr6(2) = points2(0) + 38.1331:   pointsaktr6(3) = points2(1) + 33.8528
pointsaktr6(4) = points2(0) + 40.8016:   pointsaktr6(5) = points2(1) + 49.9506
pointsaktr6(6) = points2(0) + 42.1013:   pointsaktr6(7) = points2(1) + 66.3594
pointsaktr6(8) = points2(0) + 49.9632:   pointsaktr6(9) = points2(1) + 67.6568

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsaktr6)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update

plineObj.SetBulge 0, -0.19595
plineObj.SetBulge 1, -0.0514386
plineObj.SetBulge 2, 0.0942854
plineObj.SetBulge 3, -0.0888928
plineObj.SetBulge 4, -0.972406

RetVal = plineObj.Mirror(a1, a2)
RetVal = plineObj.Mirror(b1, b2)
RetVal = plineObj.Copy
plineObj.Rotate basePoint, rotationAngle

pointsaktr7(0) = points2(0) + 105.446:   pointsaktr7(1) = points2(1) + 25.2141
pointsaktr7(2) = points2(0) + 113.409:   pointsaktr7(3) = points2(1) + 26.5497
pointsaktr7(4) = points2(0) + 113.971:   pointsaktr7(5) = points2(1) + 28.6347
pointsaktr7(6) = points2(0) + 111.928:   pointsaktr7(7) = points2(1) + 30.4463

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsaktr7)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update
  
plineObj.SetBulge 0, 0.114003
plineObj.SetBulge 1, 0.49186
plineObj.SetBulge 2, 0.09754
plineObj.SetBulge 3, -0.105753

RetVal = plineObj.Mirror(a1, a2)
RetVal = plineObj.Mirror(b1, b2)
RetVal = plineObj.Copy
plineObj.Rotate basePoint, rotationAngle

'=============================CENTER===========================================
bcp = points2(0) + (b / 2)
acp = points2(1) + (a / 2)

pointscntr1(0) = bcp + 8.33227:     pointscntr1(1) = acp - 8.2554
pointscntr1(2) = bcp + 11.6963:     pointscntr1(3) = acp - 13.0893
pointscntr1(4) = bcp + 18.3729:     pointscntr1(5) = acp - 28.2801
pointscntr1(6) = bcp + 12.7667:     pointscntr1(7) = acp - 50.0277
pointscntr1(8) = bcp + 2.66319:     pointscntr1(9) = acp - 69.2627
pointscntr1(10) = bcp + 5.32307:    pointscntr1(11) = acp - 69.9733
pointscntr1(12) = bcp + 0:          pointscntr1(13) = acp - 120.5
pointscntr1(14) = bcp - 5.32307:  pointscntr1(15) = acp - 69.9733
pointscntr1(16) = bcp - 2.66319:  pointscntr1(17) = acp - 69.2627
pointscntr1(18) = bcp - 12.7667:  pointscntr1(19) = acp - 50.0277
pointscntr1(20) = bcp - 18.3729:  pointscntr1(21) = acp - 28.2801
pointscntr1(22) = bcp - 11.6963:  pointscntr1(23) = acp - 13.0893
pointscntr1(24) = bcp - 8.33227:  pointscntr1(25) = acp - 8.2554
pointscntr1(26) = bcp - 18.0805:  pointscntr1(27) = acp + 5.68434E-14
pointscntr1(28) = bcp - 8.33227:  pointscntr1(29) = acp + 8.2554
pointscntr1(30) = bcp - 11.6963:  pointscntr1(31) = acp + 13.0893
pointscntr1(32) = bcp - 18.3729:  pointscntr1(33) = acp + 28.2801
pointscntr1(34) = bcp - 12.7667:  pointscntr1(35) = acp + 50.0277
pointscntr1(36) = bcp - 2.66319:  pointscntr1(37) = acp + 69.2627
pointscntr1(38) = bcp - 5.32307:  pointscntr1(39) = acp + 69.9733
pointscntr1(40) = bcp + 0:          pointscntr1(41) = acp + 120.5
pointscntr1(42) = bcp + 5.32307:    pointscntr1(43) = acp + 69.9733
pointscntr1(44) = bcp + 2.66319:    pointscntr1(45) = acp + 69.2627
pointscntr1(46) = bcp + 12.7667:    pointscntr1(47) = acp + 50.0277
pointscntr1(48) = bcp + 18.3729:    pointscntr1(49) = acp + 28.2801
pointscntr1(50) = bcp + 11.6963:    pointscntr1(51) = acp + 13.0893
pointscntr1(52) = bcp + 8.33227:    pointscntr1(53) = acp + 8.2554
pointscntr1(54) = bcp + 18.0805:    pointscntr1(55) = acp - 4.73695E-14


Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update

plineObj.SetBulge 0, -0.00917899
plineObj.SetBulge 1, -0.0976773
plineObj.SetBulge 2, -0.221543
plineObj.SetBulge 3, -0.247182
plineObj.SetBulge 4, 0.035764
plineObj.SetBulge 5, 0
plineObj.SetBulge 6, 0
plineObj.SetBulge 7, 0.035764
plineObj.SetBulge 8, -0.247182
plineObj.SetBulge 9, -0.221543
plineObj.SetBulge 10, -0.0976773
plineObj.SetBulge 11, -0.00917899
plineObj.SetBulge 12, 0.0324358
plineObj.SetBulge 13, 0.0324358
plineObj.SetBulge 14, -0.00917899
plineObj.SetBulge 15, -0.0976773
plineObj.SetBulge 16, -0.221543
plineObj.SetBulge 17, -0.247182
plineObj.SetBulge 18, 0.035764
plineObj.SetBulge 19, 0
plineObj.SetBulge 20, 0
plineObj.SetBulge 21, 0.035764
plineObj.SetBulge 22, -0.247182
plineObj.SetBulge 23, -0.221543
plineObj.SetBulge 24, -0.0976773
plineObj.SetBulge 25, -0.00917899
plineObj.SetBulge 26, 0.0324358
plineObj.SetBulge 27, 0.0324358

plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True

pointscntr2(0) = bcp - 14.2411:   pointscntr2(1) = acp - 9.4739E-15
pointscntr2(2) = bcp - 6.93853:   pointscntr2(3) = acp + 6.3585
pointscntr2(4) = bcp - 1.93991:   pointscntr2(5) = acp + 3.78956E-14
pointscntr2(6) = bcp - 6.93853:   pointscntr2(7) = acp - 6.3585



Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr2)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update

plineObj.SetBulge 0, 0.0240064
plineObj.SetBulge 1, 0.012607
plineObj.SetBulge 2, 0.012607
plineObj.SetBulge 3, 0.0240064


plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True

RetVal = plineObj.Mirror(a1, a2)

pointscntr3(0) = bcp - 16.0864:   pointscntr3(1) = acp + 28.747
pointscntr3(2) = bcp - 9.76953:   pointscntr3(3) = acp + 14.4056
pointscntr3(4) = bcp - 6.65271:   pointscntr3(5) = acp + 9.91527
pointscntr3(6) = bcp - 1.45081:   pointscntr3(7) = acp + 15.6103
pointscntr3(8) = bcp - 12.8423:   pointscntr3(9) = acp + 45.9255

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr3)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update

plineObj.SetBulge 0, 0.0975544
plineObj.SetBulge 1, 0.00864503
plineObj.SetBulge 2, 0.0195719
plineObj.SetBulge 3, -0.168477
plineObj.SetBulge 4, 0.186466


plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True
RetVal = plineObj.Mirror(b1, b2)
RetVal = plineObj.Mirror(a1, a2)
RetVal = plineObj.Copy
  ' Define the rotation of 180 degrees about a
  ' base point of (4, 4.25, 0)
  basePoint(0) = points2(4) - (b / 2): basePoint(1) = points2(3) - (a / 2)
  rotationAngle = 3.14159   ' 180 degrees
  plineObj.Rotate basePoint, rotationAngle
  plineObj.Update
  
pointscntr4(0) = bcp + 5.25775:         pointscntr4(1) = acp + 8.0082
pointscntr4(2) = bcp - 5.68434E-14:   pointscntr4(3) = acp + 1.31597
pointscntr4(4) = bcp - 5.25775:       pointscntr4(5) = acp + 8.0082
pointscntr4(6) = bcp + 3.78956E-14:     pointscntr4(7) = acp + 13.7137


Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr4)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update

plineObj.SetBulge 0, -0.0134617
plineObj.SetBulge 1, -0.0134617
plineObj.SetBulge 2, 0.0192317
plineObj.SetBulge 3, 0.0192317


plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True
RetVal = plineObj.Mirror(b1, b2)

pointscntr5(0) = bcp - 10.4894:         pointscntr5(1) = acp + 49.0279
pointscntr5(2) = bcp + 1.89478E-14:     pointscntr5(3) = acp + 17.5202
pointscntr5(4) = bcp + 10.4894:         pointscntr5(5) = acp + 49.0279


Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr5)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update
  
plineObj.SetBulge 0, 0.182166
plineObj.SetBulge 1, 0.182166
plineObj.SetBulge 2, 0.419513


plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True
RetVal = plineObj.Mirror(b1, b2)


pointscntr6(0) = bcp - 10.1941:    pointscntr6(1) = acp + 52.3749
pointscntr6(2) = bcp + 10.1941:    pointscntr6(3) = acp + 52.3749
pointscntr6(4) = bcp + 0:          pointscntr6(5) = acp + 68.112

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr6)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update
  
plineObj.SetBulge 0, -0.332226
plineObj.SetBulge 1, 0.250892
plineObj.SetBulge 2, 0.250892


    plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True
RetVal = plineObj.Mirror(b1, b2)

pointscntr7(0) = bcp - 2.76853:         pointscntr7(1) = acp + 71.7473
pointscntr7(2) = bcp + 7.56728E-13:     pointscntr7(3) = acp + 70.7093
pointscntr7(4) = bcp + 2.76853:         pointscntr7(5) = acp + 71.7473
pointscntr7(6) = bcp + 0:               pointscntr7(7) = acp + 98.0267


Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr7)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update
  
plineObj.SetBulge 0, -0.0342582
plineObj.SetBulge 1, -0.0342582
plineObj.SetBulge 2, 0
plineObj.SetBulge 3, 0

    plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True
RetVal = plineObj.Mirror(b1, b2)


End If
End If
  
  
  points2(1) = points2(3) + I:
  points2(3) = points2(3) + a + I:
  points2(5) = points2(3):
  points2(7) = points2(1):
  

Else: d = 1
d = d + 1

d = Cells(c, 4)
End If

Next d
c = c + 1

a = Cells(c, 2)
b = Cells(c, 3)

I = 100

Loop
              
ZoomAll
AP.Quit
End Sub

Sub MF142()
' Creating search objects
Dim AP As Excel.Application
Dim WB As Excel.Workbook
Dim WS As Excel.Worksheet

Set AP = Excel.Application
Set WB = AP.Workbooks.Open("C:\Users\������\Desktop\������\2.xlsx")
Set WS = WB.Worksheets("������ �˨������ ������")
c = 19
a = Cells(c, 2)
b = Cells(c, 3)
d = Cells(c, 4)
E = 1
I = 0
On Error Resume Next

' Description of all the points that will make up rectangles
 Dim plineObj As AcadLWPolyline
 Dim plineObjdeep As AcadLWPolyline
  Dim points(0 To 7) As Double
  Dim pointsdeep(0 To 15) As Double
  Dim pointsktr(0 To 71) As Double
  Dim pointsktr2(0 To 279) As Double
  Dim pointsstr1(0 To 7) As Double
  Dim pointsstr2(0 To 71) As Double
  Dim pointscntr1(0 To 55) As Double
  Dim pointscntr2(0 To 5) As Double
  Dim pointscntr3(0 To 7) As Double
  Dim pointscntr4(0 To 9) As Double
  Dim b1(0 To 2) As Double
  Dim b2(0 To 2) As Double
  Dim a1(0 To 2) As Double
  Dim a2(0 To 2) As Double
  Dim basePoint(0 To 2) As Double
  Dim rotationAngle As Double
  
  
points(6) = 0
 

Do While a > 0

' Calculate the coordinates of each point
  points(0) = points(6) + I:    points(1) = 0
  points(2) = points(6) + I:    points(3) = a
  points(4) = points(0) + b:    points(5) = a
  points(6) = points(0) + b:    points(7) = 0

' Constructing four segments and connecting the corresponding points
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
  plineObj.Closed = True

If a >= 185 Then
If b >= 185 Then
  ' Offset the polyline
  Dim offsetObj As Variant
                                 'plineObj.color = acByLayer
                                 'Dim layerObj As AcadLayer
                                 'Set layerObj = ThisDrawing.Layers.Add("Ball-6.1")
                                 'layerObj.color = acRed
 '===================Deep engraving===================================
  
  pointsdeep(0) = points(0) + 49:      pointsdeep(1) = points(1) + 91.3
  pointsdeep(2) = points(0) + 49:      pointsdeep(3) = points(3) - 91.3
  pointsdeep(4) = points(0) + 91.3:    pointsdeep(5) = points(3) - 49
  pointsdeep(6) = points(4) - 91.3:    pointsdeep(7) = points(3) - 49
  pointsdeep(8) = points(4) - 49:      pointsdeep(9) = points(3) - 91.3
  pointsdeep(10) = points(4) - 49:     pointsdeep(11) = points(1) + 91.3
  pointsdeep(12) = points(4) - 91.3:   pointsdeep(13) = points(1) + 49
  pointsdeep(14) = points(0) + 91.3:   pointsdeep(15) = points(1) + 49

plineObjdeep.Layer = "K-grav"
plineObjdeep.Update
Set plineObjdeep = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsdeep)
  plineObjdeep.Closed = True

plineObjdeep.Layer = "C-Mill"
plineObjdeep.Update
  offsetObj = plineObjdeep.Offset(20)
plineObjdeep.Layer = "0"
plineObjdeep.Update

End If
End If

If a >= 145 Then
If b >= 145 Then

'===================Pattern outline on the frame===================================

pointsktr(0) = points(0) + 42:    pointsktr(1) = points(1) + 30
pointsktr(2) = points(0) + 42:    pointsktr(3) = points(1) + 36
pointsktr(4) = points(0) + 39:    pointsktr(5) = points(1) + 36
pointsktr(6) = points(0) + 39:    pointsktr(7) = points(1) + 30
pointsktr(8) = points(0) + 30:    pointsktr(9) = points(1) + 30
pointsktr(10) = points(0) + 30:   pointsktr(11) = points(1) + 39
pointsktr(12) = points(0) + 36:   pointsktr(13) = points(1) + 39
pointsktr(14) = points(0) + 36:   pointsktr(15) = points(1) + 42
pointsktr(16) = points(0) + 30:   pointsktr(17) = points(1) + 42
pointsktr(18) = points(0) + 30:   pointsktr(19) = points(3) - 42
pointsktr(20) = points(0) + 36:   pointsktr(21) = points(3) - 42
pointsktr(22) = points(0) + 36:   pointsktr(23) = points(3) - 39
pointsktr(24) = points(0) + 30:   pointsktr(25) = points(3) - 39
pointsktr(26) = points(0) + 30:   pointsktr(27) = points(3) - 30
pointsktr(28) = points(0) + 39:   pointsktr(29) = points(3) - 30
pointsktr(30) = points(0) + 39:   pointsktr(31) = points(3) - 36
pointsktr(32) = points(0) + 42:   pointsktr(33) = points(3) - 36
pointsktr(34) = points(0) + 42:   pointsktr(35) = points(3) - 30
pointsktr(36) = points(4) - 42:   pointsktr(37) = points(3) - 30
pointsktr(38) = points(4) - 42:   pointsktr(39) = points(3) - 36
pointsktr(40) = points(4) - 39:   pointsktr(41) = points(3) - 36
pointsktr(42) = points(4) - 39:   pointsktr(43) = points(3) - 30
pointsktr(44) = points(4) - 30:   pointsktr(45) = points(3) - 30
pointsktr(46) = points(4) - 30:   pointsktr(47) = points(3) - 39
pointsktr(48) = points(4) - 36:   pointsktr(49) = points(3) - 39
pointsktr(50) = points(4) - 36:   pointsktr(51) = points(3) - 42
pointsktr(52) = points(4) - 30:   pointsktr(53) = points(3) - 42
pointsktr(54) = points(4) - 30:   pointsktr(55) = points(1) + 42
pointsktr(56) = points(4) - 36:   pointsktr(57) = points(1) + 42
pointsktr(58) = points(4) - 36:   pointsktr(59) = points(1) + 39
pointsktr(60) = points(4) - 30:   pointsktr(61) = points(1) + 39
pointsktr(62) = points(4) - 30:   pointsktr(63) = points(1) + 30
pointsktr(64) = points(4) - 39:   pointsktr(65) = points(1) + 30
pointsktr(66) = points(4) - 39:   pointsktr(67) = points(1) + 36
pointsktr(68) = points(4) - 42:   pointsktr(69) = points(1) + 36
pointsktr(70) = points(4) - 42:   pointsktr(71) = points(1) + 30

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsktr)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update

pointsktr2(0) = points(0) + 45:         pointsktr2(1) = points(1) + 33
pointsktr2(2) = points(0) + 45:         pointsktr2(3) = points(1) + 36
pointsktr2(4) = points(0) + 70.15:      pointsktr2(5) = points(1) + 36
pointsktr2(6) = points(0) + 68.8884:    pointsktr2(7) = points(1) + 36.914
pointsktr2(8) = points(0) + 67.4818:    pointsktr2(9) = points(1) + 37.4156
pointsktr2(10) = points(0) + 65.6556:   pointsktr2(11) = points(1) + 38.4118
pointsktr2(12) = points(0) + 64.0803:   pointsktr2(13) = points(1) + 39.9661
pointsktr2(14) = points(0) + 61.0444:   pointsktr2(15) = points(1) + 42.5889
pointsktr2(16) = points(0) + 57.0324:   pointsktr2(17) = points(1) + 44.4867
pointsktr2(18) = points(0) + 52.0524:   pointsktr2(19) = points(1) + 45.3159
pointsktr2(20) = points(0) + 52.7992:   pointsktr2(21) = points(1) + 46.9026
pointsktr2(22) = points(0) + 53.236:    pointsktr2(23) = points(1) + 48.4263
pointsktr2(24) = points(0) + 53.3051:   pointsktr2(25) = points(1) + 49.0892
pointsktr2(26) = points(0) + 53.5541:   pointsktr2(27) = points(1) + 50.0059
pointsktr2(28) = points(0) + 53.773:    pointsktr2(29) = points(1) + 50.7406
pointsktr2(30) = points(0) + 54.3842:   pointsktr2(31) = points(1) + 53.0168
pointsktr2(32) = points(0) + 54.5717:   pointsktr2(33) = points(1) + 54.154
pointsktr2(34) = points(0) + 54.5256:   pointsktr2(35) = points(1) + 54.5256
pointsktr2(36) = points(0) + 54.154:    pointsktr2(37) = points(1) + 54.5717
pointsktr2(38) = points(0) + 53.0168:   pointsktr2(39) = points(1) + 54.3842
pointsktr2(40) = points(0) + 50.7406:   pointsktr2(41) = points(1) + 53.773
pointsktr2(42) = points(0) + 50.0059:   pointsktr2(43) = points(1) + 53.5541
pointsktr2(44) = points(0) + 49.0892:   pointsktr2(45) = points(1) + 53.3051
pointsktr2(46) = points(0) + 48.4263:   pointsktr2(47) = points(1) + 53.236
pointsktr2(48) = points(0) + 46.9026:   pointsktr2(49) = points(1) + 52.7992
pointsktr2(50) = points(0) + 45.3159:   pointsktr2(51) = points(1) + 52.0524
pointsktr2(52) = points(0) + 44.4867:   pointsktr2(53) = points(1) + 57.0324
pointsktr2(54) = points(0) + 42.5889:   pointsktr2(55) = points(1) + 61.0444
pointsktr2(56) = points(0) + 39.9661:   pointsktr2(57) = points(1) + 64.0803
pointsktr2(58) = points(0) + 38.4118:   pointsktr2(59) = points(1) + 65.6556
pointsktr2(60) = points(0) + 37.4156:   pointsktr2(61) = points(1) + 67.4818
pointsktr2(62) = points(0) + 36.914:    pointsktr2(63) = points(1) + 68.8884
pointsktr2(64) = points(0) + 36:        pointsktr2(65) = points(1) + 70.15
pointsktr2(66) = points(0) + 36:        pointsktr2(67) = points(1) + 45
pointsktr2(68) = points(0) + 33:        pointsktr2(69) = points(1) + 45
pointsktr2(70) = points(0) + 33:        pointsktr2(71) = points(3) - 45
pointsktr2(72) = points(0) + 36:       pointsktr2(73) = points(3) - 45
pointsktr2(74) = points(0) + 36:       pointsktr2(75) = points(3) - 70.15
pointsktr2(76) = points(0) + 36.914:   pointsktr2(77) = points(3) - 68.8884
pointsktr2(78) = points(0) + 37.4156:  pointsktr2(79) = points(3) - 67.4818
pointsktr2(80) = points(0) + 38.4118:  pointsktr2(81) = points(3) - 65.6556
pointsktr2(82) = points(0) + 39.9661:  pointsktr2(83) = points(3) - 64.0803
pointsktr2(84) = points(0) + 42.5889:  pointsktr2(85) = points(3) - 61.0444
pointsktr2(86) = points(0) + 44.4867:  pointsktr2(87) = points(3) - 57.0324
pointsktr2(88) = points(0) + 45.3159:  pointsktr2(89) = points(3) - 52.0524
pointsktr2(90) = points(0) + 46.9026:  pointsktr2(91) = points(3) - 52.7992
pointsktr2(92) = points(0) + 48.4263:  pointsktr2(93) = points(3) - 53.236
pointsktr2(94) = points(0) + 49.0892:  pointsktr2(95) = points(3) - 53.3051
pointsktr2(96) = points(0) + 50.0059:  pointsktr2(97) = points(3) - 53.5541
pointsktr2(98) = points(0) + 50.7406:  pointsktr2(99) = points(3) - 53.773
pointsktr2(100) = points(0) + 53.0168: pointsktr2(101) = points(3) - 54.3842
pointsktr2(102) = points(0) + 54.154:  pointsktr2(103) = points(3) - 54.5717
pointsktr2(104) = points(0) + 54.5256: pointsktr2(105) = points(3) - 54.5256
pointsktr2(106) = points(0) + 54.5717: pointsktr2(107) = points(3) - 54.154
pointsktr2(108) = points(0) + 54.3842: pointsktr2(109) = points(3) - 53.0168
pointsktr2(110) = points(0) + 53.773:  pointsktr2(111) = points(3) - 50.7406
pointsktr2(112) = points(0) + 53.5541: pointsktr2(113) = points(3) - 50.0059
pointsktr2(114) = points(0) + 53.3051: pointsktr2(115) = points(3) - 49.0892
pointsktr2(116) = points(0) + 53.236:  pointsktr2(117) = points(3) - 48.4263
pointsktr2(118) = points(0) + 52.7992: pointsktr2(119) = points(3) - 46.9026
pointsktr2(120) = points(0) + 52.0524: pointsktr2(121) = points(3) - 45.3159
pointsktr2(122) = points(0) + 57.0324: pointsktr2(123) = points(3) - 44.4867
pointsktr2(124) = points(0) + 61.0444: pointsktr2(125) = points(3) - 42.5889
pointsktr2(126) = points(0) + 64.0803: pointsktr2(127) = points(3) - 39.9661
pointsktr2(128) = points(0) + 65.6556: pointsktr2(129) = points(3) - 38.4118
pointsktr2(130) = points(0) + 67.4818: pointsktr2(131) = points(3) - 37.4156
pointsktr2(132) = points(0) + 68.8884: pointsktr2(133) = points(3) - 36.914
pointsktr2(134) = points(0) + 70.15:   pointsktr2(135) = points(3) - 36
pointsktr2(136) = points(0) + 45:      pointsktr2(137) = points(3) - 36
pointsktr2(138) = points(0) + 45:      pointsktr2(139) = points(3) - 33
pointsktr2(140) = points(4) - 45:      pointsktr2(141) = points(3) - 33
pointsktr2(142) = points(4) - 45:      pointsktr2(143) = points(3) - 36
pointsktr2(144) = points(4) - 70.15:   pointsktr2(145) = points(3) - 36
pointsktr2(146) = points(4) - 68.8884: pointsktr2(147) = points(3) - 36.914
pointsktr2(148) = points(4) - 67.4818: pointsktr2(149) = points(3) - 37.4156
pointsktr2(150) = points(4) - 65.6556: pointsktr2(151) = points(3) - 38.4118
pointsktr2(152) = points(4) - 64.0803: pointsktr2(153) = points(3) - 39.9661
pointsktr2(154) = points(4) - 61.0444: pointsktr2(155) = points(3) - 42.5889
pointsktr2(156) = points(4) - 57.0324: pointsktr2(157) = points(3) - 44.4867
pointsktr2(158) = points(4) - 52.0524: pointsktr2(159) = points(3) - 45.3159
pointsktr2(160) = points(4) - 52.7992: pointsktr2(161) = points(3) - 46.9026
pointsktr2(162) = points(4) - 53.236:  pointsktr2(163) = points(3) - 48.4263
pointsktr2(164) = points(4) - 53.3051: pointsktr2(165) = points(3) - 49.0892
pointsktr2(166) = points(4) - 53.5541: pointsktr2(167) = points(3) - 50.0059
pointsktr2(168) = points(4) - 53.773:  pointsktr2(169) = points(3) - 50.7406
pointsktr2(170) = points(4) - 54.3842: pointsktr2(171) = points(3) - 53.0168
pointsktr2(172) = points(4) - 54.5717: pointsktr2(173) = points(3) - 54.154
pointsktr2(174) = points(4) - 54.5256: pointsktr2(175) = points(3) - 54.5256
pointsktr2(176) = points(4) - 54.154:  pointsktr2(177) = points(3) - 54.5717
pointsktr2(178) = points(4) - 53.0168: pointsktr2(179) = points(3) - 54.3842
pointsktr2(180) = points(4) - 50.7406: pointsktr2(181) = points(3) - 53.773
pointsktr2(182) = points(4) - 50.0059: pointsktr2(183) = points(3) - 53.5541
pointsktr2(184) = points(4) - 49.0892: pointsktr2(185) = points(3) - 53.3051
pointsktr2(186) = points(4) - 48.4263: pointsktr2(187) = points(3) - 53.236
pointsktr2(188) = points(4) - 46.9026: pointsktr2(189) = points(3) - 52.7992
pointsktr2(190) = points(4) - 45.3159: pointsktr2(191) = points(3) - 52.0524
pointsktr2(192) = points(4) - 44.4867: pointsktr2(193) = points(3) - 57.0324
pointsktr2(194) = points(4) - 42.5889: pointsktr2(195) = points(3) - 61.0444
pointsktr2(196) = points(4) - 39.9661: pointsktr2(197) = points(3) - 64.0803
pointsktr2(198) = points(4) - 38.4118: pointsktr2(199) = points(3) - 65.6556
pointsktr2(200) = points(4) - 37.4156: pointsktr2(201) = points(3) - 67.4818
pointsktr2(202) = points(4) - 36.914:  pointsktr2(203) = points(3) - 68.8884
pointsktr2(204) = points(4) - 36:      pointsktr2(205) = points(3) - 70.15
pointsktr2(206) = points(4) - 36:      pointsktr2(207) = points(3) - 45
pointsktr2(208) = points(4) - 33:      pointsktr2(209) = points(3) - 45
pointsktr2(210) = points(4) - 33:      pointsktr2(211) = points(1) + 45
pointsktr2(212) = points(4) - 36:      pointsktr2(213) = points(1) + 45
pointsktr2(214) = points(4) - 36:      pointsktr2(215) = points(1) + 70.15
pointsktr2(216) = points(4) - 36.914:  pointsktr2(217) = points(1) + 68.8884
pointsktr2(218) = points(4) - 37.4156: pointsktr2(219) = points(1) + 67.4818
pointsktr2(220) = points(4) - 38.4118: pointsktr2(221) = points(1) + 65.6556
pointsktr2(222) = points(4) - 39.9661: pointsktr2(223) = points(1) + 64.0803
pointsktr2(224) = points(4) - 42.5889: pointsktr2(225) = points(1) + 61.0444
pointsktr2(226) = points(4) - 44.4867: pointsktr2(227) = points(1) + 57.0324
pointsktr2(228) = points(4) - 45.3159: pointsktr2(229) = points(1) + 52.0524
pointsktr2(230) = points(4) - 46.9026: pointsktr2(231) = points(1) + 52.7992
pointsktr2(232) = points(4) - 48.4263: pointsktr2(233) = points(1) + 53.236
pointsktr2(234) = points(4) - 49.0892: pointsktr2(235) = points(1) + 53.3051
pointsktr2(236) = points(4) - 50.0059: pointsktr2(237) = points(1) + 53.5541
pointsktr2(238) = points(4) - 50.7406: pointsktr2(239) = points(1) + 53.773
pointsktr2(240) = points(4) - 53.0168: pointsktr2(241) = points(1) + 54.3842
pointsktr2(242) = points(4) - 54.154:  pointsktr2(243) = points(1) + 54.5717
pointsktr2(244) = points(4) - 54.5256: pointsktr2(245) = points(1) + 54.5256
pointsktr2(246) = points(4) - 54.5717: pointsktr2(247) = points(1) + 54.154
pointsktr2(248) = points(4) - 54.3842: pointsktr2(249) = points(1) + 53.0168
pointsktr2(250) = points(4) - 53.773:  pointsktr2(251) = points(1) + 50.7406
pointsktr2(252) = points(4) - 53.5541: pointsktr2(253) = points(1) + 50.0059
pointsktr2(254) = points(4) - 53.3051: pointsktr2(255) = points(1) + 49.0892
pointsktr2(256) = points(4) - 53.236:  pointsktr2(257) = points(1) + 48.4263
pointsktr2(258) = points(4) - 52.7992: pointsktr2(259) = points(1) + 46.9026
pointsktr2(260) = points(4) - 52.0524: pointsktr2(261) = points(1) + 45.3159
pointsktr2(262) = points(4) - 57.0324: pointsktr2(263) = points(1) + 44.4867
pointsktr2(264) = points(4) - 61.0444: pointsktr2(265) = points(1) + 42.5889
pointsktr2(266) = points(4) - 64.0803: pointsktr2(267) = points(1) + 39.9661
pointsktr2(268) = points(4) - 65.6556: pointsktr2(269) = points(1) + 38.4118
pointsktr2(270) = points(4) - 67.4818: pointsktr2(271) = points(1) + 37.4156
pointsktr2(272) = points(4) - 68.8884: pointsktr2(273) = points(1) + 36.914
pointsktr2(274) = points(4) - 70.15:   pointsktr2(275) = points(1) + 36
pointsktr2(276) = points(4) - 45:      pointsktr2(277) = points(1) + 36
pointsktr2(278) = points(4) - 45:      pointsktr2(279) = points(1) + 33

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsktr2)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  
plineObj.SetBulge 0, 0
plineObj.SetBulge 1, 0
plineObj.SetBulge 2, 0.0452849
plineObj.SetBulge 3, 0.0767779
plineObj.SetBulge 4, -0.125063
plineObj.SetBulge 5, -0.0179955
plineObj.SetBulge 6, 0.0434551
plineObj.SetBulge 7, 0.0833797
plineObj.SetBulge 8, 0.0493602
plineObj.SetBulge 9, -0.0221733
plineObj.SetBulge 10, 0.0927081
plineObj.SetBulge 11, 0
plineObj.SetBulge 12, -0.104648
plineObj.SetBulge 13, 0.0311187
plineObj.SetBulge 14, -0.0452272
plineObj.SetBulge 15, 0.0707349
plineObj.SetBulge 16, 0.0483396
plineObj.SetBulge 17, 0.0483396
plineObj.SetBulge 18, 0.0707349
plineObj.SetBulge 19, -0.0452272
plineObj.SetBulge 20, 0.0311187
plineObj.SetBulge 21, -0.104648
plineObj.SetBulge 22, 0
plineObj.SetBulge 23, 0.0927081
plineObj.SetBulge 24, -0.0221733
plineObj.SetBulge 25, 0.0493602
plineObj.SetBulge 26, 0.0833797
plineObj.SetBulge 27, 0.0434551
plineObj.SetBulge 28, -0.0179955
plineObj.SetBulge 29, -0.125063
plineObj.SetBulge 30, 0.0767779
plineObj.SetBulge 31, 0.0452849
plineObj.SetBulge 32, 0
plineObj.SetBulge 33, 0
plineObj.SetBulge 34, 0
plineObj.SetBulge 35, 0
plineObj.SetBulge 36, 0
plineObj.SetBulge 37, 0.0452849
plineObj.SetBulge 38, 0.0767779
plineObj.SetBulge 39, -0.125063
plineObj.SetBulge 40, -0.0179955
plineObj.SetBulge 41, 0.0434551
plineObj.SetBulge 42, 0.0833797
plineObj.SetBulge 43, 0.0493602
plineObj.SetBulge 44, -0.0221733
plineObj.SetBulge 45, 0.0927081
plineObj.SetBulge 46, 0
plineObj.SetBulge 47, -0.104648
plineObj.SetBulge 48, 0.0311187
plineObj.SetBulge 49, -0.0452272
plineObj.SetBulge 50, 0.0707349
plineObj.SetBulge 51, 0.0483396
plineObj.SetBulge 52, 0.0483396
plineObj.SetBulge 53, 0.0707349
plineObj.SetBulge 54, -0.0452272
plineObj.SetBulge 55, 0.0311187
plineObj.SetBulge 56, -0.104648
plineObj.SetBulge 57, 0
plineObj.SetBulge 58, 0.0927081
plineObj.SetBulge 59, -0.0221733
plineObj.SetBulge 60, 0.0493602
plineObj.SetBulge 61, 0.0833797
plineObj.SetBulge 62, 0.0434551
plineObj.SetBulge 63, -0.0179955
plineObj.SetBulge 64, -0.125063
plineObj.SetBulge 65, 0.0767779
plineObj.SetBulge 66, 0.0452849
plineObj.SetBulge 67, 0
plineObj.SetBulge 68, 0
plineObj.SetBulge 69, 0
plineObj.SetBulge 70, 0
plineObj.SetBulge 71, 0
plineObj.SetBulge 72, 0.0452849
plineObj.SetBulge 73, 0.0767779
plineObj.SetBulge 74, -0.125063
plineObj.SetBulge 75, -0.0179955
plineObj.SetBulge 76, 0.0434551
plineObj.SetBulge 77, 0.0833797
plineObj.SetBulge 78, 0.0493602
plineObj.SetBulge 79, -0.0221733
plineObj.SetBulge 80, 0.0927081
plineObj.SetBulge 81, 0
plineObj.SetBulge 82, -0.104648
plineObj.SetBulge 83, 0.0311187
plineObj.SetBulge 84, -0.0452272
plineObj.SetBulge 85, 0.0707349
plineObj.SetBulge 86, 0.0483396
plineObj.SetBulge 87, 0.0483396
plineObj.SetBulge 88, 0.0707349
plineObj.SetBulge 89, -0.0452272
plineObj.SetBulge 90, 0.0311187
plineObj.SetBulge 91, -0.104648
plineObj.SetBulge 92, 0
plineObj.SetBulge 93, 0.0927081
plineObj.SetBulge 94, -0.0221733
plineObj.SetBulge 95, 0.0493602
plineObj.SetBulge 96, 0.0833797
plineObj.SetBulge 97, 0.0434551
plineObj.SetBulge 98, -0.0179955
plineObj.SetBulge 99, -0.125063
plineObj.SetBulge 100, 0.0767779
plineObj.SetBulge 101, 0.0452849
plineObj.SetBulge 102, 0
plineObj.SetBulge 103, 0
plineObj.SetBulge 104, 0
plineObj.SetBulge 105, 0
plineObj.SetBulge 106, 0
plineObj.SetBulge 107, 0.0452849
plineObj.SetBulge 108, 0.0767779
plineObj.SetBulge 109, -0.125063
plineObj.SetBulge 110, -0.0179955
plineObj.SetBulge 111, 0.0434551
plineObj.SetBulge 112, 0.0833797
plineObj.SetBulge 113, 0.0493602
plineObj.SetBulge 114, -0.0221733
plineObj.SetBulge 115, 0.0927081
plineObj.SetBulge 116, 0
plineObj.SetBulge 117, -0.104648
plineObj.SetBulge 118, 0.0311187
plineObj.SetBulge 119, -0.0452272
plineObj.SetBulge 120, 0.0707349
plineObj.SetBulge 121, 0.0483396
plineObj.SetBulge 122, 0.0483396
plineObj.SetBulge 123, 0.0707349
plineObj.SetBulge 124, -0.0452272
plineObj.SetBulge 125, 0.0311187
plineObj.SetBulge 126, -0.104648
plineObj.SetBulge 127, 0
plineObj.SetBulge 128, 0.0927081
plineObj.SetBulge 129, -0.0221733
plineObj.SetBulge 130, 0.0493602
plineObj.SetBulge 131, 0.0833797
plineObj.SetBulge 132, 0.0434551
plineObj.SetBulge 133, -0.0179955
plineObj.SetBulge 134, -0.125063
plineObj.SetBulge 135, 0.0767779
plineObj.SetBulge 136, 0.0452849
plineObj.SetBulge 137, 0
plineObj.SetBulge 138, 0
plineObj.SetBulge 139, 0
plineObj.Update
  
'=========================Stripes======================================
  a1(0) = points(4) - (b / 2):  a1(1) = (a / 2) + 1
  a2(0) = points(4) - (b / 2):  a2(1) = (a / 2) - 1
  b1(0) = (b / 2) - 1:  b1(1) = (a / 2)
  b2(0) = (b / 2) + 1:  b2(1) = (a / 2)

  pointsstr1(0) = points(0) + 42:     pointsstr1(1) = points(1) + 55.6455
  pointsstr1(2) = points(0) + 42:     pointsstr1(3) = points(1) + 45
  pointsstr1(4) = points(0) + 39:     pointsstr1(5) = points(1) + 45
  pointsstr1(6) = points(0) + 39:     pointsstr1(7) = points(1) + 60.5518
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  plineObj.SetBulge 3, -0.133444
  plineObj.Update
  RetVal = plineObj.Mirror(a1, a2)
  RetVal = plineObj.Mirror(b1, b2)
  RetVal = plineObj.Copy
  basePoint(0) = points(4) - (b / 2): basePoint(1) = points(3) - (a / 2)
  rotationAngle = 3.14159   ' 45 degrees
  plineObj.Rotate basePoint, rotationAngle
  plineObj.Update
  
  pointsstr1(0) = points(0) + 55.6455:   pointsstr1(1) = points(1) + 42
  pointsstr1(2) = points(0) + 45:        pointsstr1(3) = points(1) + 42
  pointsstr1(4) = points(0) + 45:        pointsstr1(5) = points(1) + 39
  pointsstr1(6) = points(0) + 60.5518:   pointsstr1(7) = points(1) + 39
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  plineObj.SetBulge 3, 0.133444
  plineObj.Update
  RetVal = plineObj.Mirror(a1, a2)
  RetVal = plineObj.Mirror(b1, b2)
  RetVal = plineObj.Copy
  basePoint(0) = points(4) - (b / 2): basePoint(1) = points(3) - (a / 2)
  rotationAngle = 3.14159   ' 45 degrees
  plineObj.Rotate basePoint, rotationAngle
  plineObj.Update
  
'=========================Squares======================================
  pointsstr1(0) = points(0) + 33:   pointsstr1(1) = points(1) + 33
  pointsstr1(2) = points(0) + 33:   pointsstr1(3) = points(1) + 36
  pointsstr1(4) = points(0) + 36:   pointsstr1(5) = points(1) + 36
  pointsstr1(6) = points(0) + 36:   pointsstr1(7) = points(1) + 33
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  pointsstr1(0) = points(0) + 39:   pointsstr1(1) = points(1) + 39
  pointsstr1(2) = points(0) + 39:   pointsstr1(3) = points(1) + 42
  pointsstr1(4) = points(0) + 42:   pointsstr1(5) = points(1) + 42
  pointsstr1(6) = points(0) + 42:   pointsstr1(7) = points(1) + 39
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  '***********************
  pointsstr1(0) = points(4) - 33:   pointsstr1(1) = points(1) + 33
  pointsstr1(2) = points(4) - 33:   pointsstr1(3) = points(1) + 36
  pointsstr1(4) = points(4) - 36:   pointsstr1(5) = points(1) + 36
  pointsstr1(6) = points(4) - 36:   pointsstr1(7) = points(1) + 33
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  pointsstr1(0) = points(4) - 39:   pointsstr1(1) = points(1) + 39
  pointsstr1(2) = points(4) - 39:   pointsstr1(3) = points(1) + 42
  pointsstr1(4) = points(4) - 42:   pointsstr1(5) = points(1) + 42
  pointsstr1(6) = points(4) - 42:   pointsstr1(7) = points(1) + 39
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  '***********************
  pointsstr1(0) = points(0) + 33:   pointsstr1(1) = points(3) - 33
  pointsstr1(2) = points(0) + 33:   pointsstr1(3) = points(3) - 36
  pointsstr1(4) = points(0) + 36:   pointsstr1(5) = points(3) - 36
  pointsstr1(6) = points(0) + 36:   pointsstr1(7) = points(3) - 33
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  pointsstr1(0) = points(0) + 39:   pointsstr1(1) = points(3) - 39
  pointsstr1(2) = points(0) + 39:   pointsstr1(3) = points(3) - 42
  pointsstr1(4) = points(0) + 42:   pointsstr1(5) = points(3) - 42
  pointsstr1(6) = points(0) + 42:   pointsstr1(7) = points(3) - 39
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  '***********************
  pointsstr1(0) = points(4) - 33:   pointsstr1(1) = points(3) - 33
  pointsstr1(2) = points(4) - 33:   pointsstr1(3) = points(3) - 36
  pointsstr1(4) = points(4) - 36:   pointsstr1(5) = points(3) - 36
  pointsstr1(6) = points(4) - 36:   pointsstr1(7) = points(3) - 33
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  pointsstr1(0) = points(4) - 39:   pointsstr1(1) = points(3) - 39
  pointsstr1(2) = points(4) - 39:   pointsstr1(3) = points(3) - 42
  pointsstr1(4) = points(4) - 42:   pointsstr1(5) = points(3) - 42
  pointsstr1(6) = points(4) - 42:   pointsstr1(7) = points(3) - 39
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
'===========================Oval====================================

pointsstr2(0) = points(0) + 50.5415:    pointsstr2(1) = points(1) + 48.0666
pointsstr2(2) = points(0) + 50.1973:    pointsstr2(3) = points(1) + 47.76
pointsstr2(4) = points(0) + 49.8259:    pointsstr2(5) = points(1) + 47.5003
pointsstr2(6) = points(0) + 49.4388:    pointsstr2(7) = points(1) + 47.2954
pointsstr2(8) = points(0) + 49.0475:    pointsstr2(9) = points(1) + 47.1516
pointsstr2(10) = points(0) + 48.664:    pointsstr2(11) = points(1) + 47.0732
pointsstr2(12) = points(0) + 48.2999:   pointsstr2(13) = points(1) + 47.0625
pointsstr2(14) = points(0) + 47.9664:   pointsstr2(15) = points(1) + 47.12
pointsstr2(16) = points(0) + 47.6735:   pointsstr2(17) = points(1) + 47.2438
pointsstr2(18) = points(0) + 47.4302:   pointsstr2(19) = points(1) + 47.4302
pointsstr2(20) = points(0) + 47.2438:   pointsstr2(21) = points(1) + 47.6735
pointsstr2(22) = points(0) + 47.12:     pointsstr2(23) = points(1) + 47.9664
pointsstr2(24) = points(0) + 47.0625:   pointsstr2(25) = points(1) + 48.2999
pointsstr2(26) = points(0) + 47.0732:   pointsstr2(27) = points(1) + 48.664
pointsstr2(28) = points(0) + 47.1516:   pointsstr2(29) = points(1) + 49.0475
pointsstr2(30) = points(0) + 47.2954:   pointsstr2(31) = points(1) + 49.4388
pointsstr2(32) = points(0) + 47.5003:   pointsstr2(33) = points(1) + 49.8259
pointsstr2(34) = points(0) + 47.76:     pointsstr2(35) = points(1) + 50.1973
pointsstr2(36) = points(0) + 48.0666:   pointsstr2(37) = points(1) + 50.5415
pointsstr2(38) = points(0) + 48.4108:   pointsstr2(39) = points(1) + 50.848
pointsstr2(40) = points(0) + 48.7821:   pointsstr2(41) = points(1) + 51.1077
pointsstr2(42) = points(0) + 49.1693:   pointsstr2(43) = points(1) + 51.3126
pointsstr2(44) = points(0) + 49.5606:   pointsstr2(45) = points(1) + 51.4564
pointsstr2(46) = points(0) + 49.944:    pointsstr2(47) = points(1) + 51.5349
pointsstr2(48) = points(0) + 50.3081:   pointsstr2(49) = points(1) + 51.5455
pointsstr2(50) = points(0) + 50.6416:   pointsstr2(51) = points(1) + 51.4881
pointsstr2(52) = points(0) + 50.9345:   pointsstr2(53) = points(1) + 51.3643
pointsstr2(54) = points(0) + 51.1778:   pointsstr2(55) = points(1) + 51.1778
pointsstr2(56) = points(0) + 51.3643:   pointsstr2(57) = points(1) + 50.9345
pointsstr2(58) = points(0) + 51.4881:   pointsstr2(59) = points(1) + 50.6416
pointsstr2(60) = points(0) + 51.5455:   pointsstr2(61) = points(1) + 50.3081
pointsstr2(62) = points(0) + 51.5349:   pointsstr2(63) = points(1) + 49.944
pointsstr2(64) = points(0) + 51.4564:   pointsstr2(65) = points(1) + 49.5606
pointsstr2(66) = points(0) + 51.3126:   pointsstr2(67) = points(1) + 49.1693
pointsstr2(68) = points(0) + 51.1077:   pointsstr2(69) = points(1) + 48.7821
pointsstr2(70) = points(0) + 50.848:    pointsstr2(71) = points(1) + 48.4108

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr2)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update

plineObj.SetBulge 0, -0.0288637
plineObj.SetBulge 1, -0.0299658
plineObj.SetBulge 2, -0.0320654
plineObj.SetBulge 3, -0.0354072
plineObj.SetBulge 4, -0.040166
plineObj.SetBulge 5, -0.0463995
plineObj.SetBulge 6, -0.0537273
plineObj.SetBulge 7, -0.0608551
plineObj.SetBulge 8, -0.0659533
plineObj.SetBulge 9, -0.0659533
plineObj.SetBulge 10, -0.0608551
plineObj.SetBulge 11, -0.0537273
plineObj.SetBulge 12, -0.0463995
plineObj.SetBulge 13, -0.040166
plineObj.SetBulge 14, -0.0354072
plineObj.SetBulge 15, -0.0320654
plineObj.SetBulge 16, -0.0299658
plineObj.SetBulge 17, -0.0288637
plineObj.SetBulge 18, -0.0288637
plineObj.SetBulge 19, -0.0299658
plineObj.SetBulge 20, -0.0320654
plineObj.SetBulge 21, -0.0354072
plineObj.SetBulge 22, -0.040166
plineObj.SetBulge 23, -0.0463995
plineObj.SetBulge 24, -0.0537273
plineObj.SetBulge 25, -0.0608551
plineObj.SetBulge 26, -0.0659533
plineObj.SetBulge 27, -0.0659533
plineObj.SetBulge 28, -0.0608551
plineObj.SetBulge 29, -0.0537273
plineObj.SetBulge 30, -0.0463995
plineObj.SetBulge 31, -0.040166
plineObj.SetBulge 32, -0.0354072
plineObj.SetBulge 33, -0.0320654
plineObj.SetBulge 34, -0.0299658
plineObj.SetBulge 35, -0.0288637

  RetVal = plineObj.Mirror(a1, a2)
  RetVal = plineObj.Mirror(b1, b2)
  RetVal = plineObj.Copy
  basePoint(0) = points(4) - (b / 2): basePoint(1) = points(3) - (a / 2)
  rotationAngle = 3.14159   ' 45 degrees
  plineObj.Rotate basePoint, rotationAngle
  plineObj.Update
  
  If a > 540 Then
  If b > 250 Then
'=========================CENTER======================================
  
  bcp = points(0) + (b / 2)
  acp = points(1) + (a / 2)
  a1(0) = points(4) - (b / 2):  a1(1) = (a / 2) + 1
  a2(0) = points(4) - (b / 2):  a2(1) = (a / 2) - 1
  b1(0) = (b / 2) - 1:  b1(1) = (a / 2)
  b2(0) = (b / 2) + 1:  b2(1) = (a / 2)
 
  pointscntr1(0) = bcp + 27.1207:    pointscntr1(1) = acp - 5.68434E-14
  pointscntr1(2) = bcp + 12.4984:    pointscntr1(3) = acp - 12.3831
  pointscntr1(4) = bcp + 17.5444:    pointscntr1(5) = acp - 19.634
  pointscntr1(6) = bcp + 27.5594:    pointscntr1(7) = acp - 42.4202
  pointscntr1(8) = bcp + 19.1501:    pointscntr1(9) = acp - 75.0415
  pointscntr1(10) = bcp + 3.99478:   pointscntr1(11) = acp - 103.894
  pointscntr1(12) = bcp + 7.9846:    pointscntr1(13) = acp - 104.96
  pointscntr1(14) = bcp + 0:         pointscntr1(15) = acp - 180.75
  pointscntr1(16) = bcp - 7.9846:    pointscntr1(17) = acp - 104.96
  pointscntr1(18) = bcp - 3.99478:   pointscntr1(19) = acp - 103.894
  pointscntr1(20) = bcp - 19.1501:   pointscntr1(21) = acp - 75.0415
  pointscntr1(22) = bcp - 27.5594:   pointscntr1(23) = acp - 42.4202
  pointscntr1(24) = bcp - 17.5444:   pointscntr1(25) = acp - 19.634
  pointscntr1(26) = bcp - 12.4984:   pointscntr1(27) = acp - 12.3831
  pointscntr1(28) = bcp - 27.1207:   pointscntr1(29) = acp + 5.68434E-14
  pointscntr1(30) = bcp - 12.4984:   pointscntr1(31) = acp + 12.3831
  pointscntr1(32) = bcp - 17.5444:   pointscntr1(33) = acp + 19.634
  pointscntr1(34) = bcp - 27.5594:   pointscntr1(35) = acp + 42.4202
  pointscntr1(36) = bcp - 19.1501:   pointscntr1(37) = acp + 75.0415
  pointscntr1(38) = bcp - 3.99478:   pointscntr1(39) = acp + 103.894
  pointscntr1(40) = bcp - 7.9846:    pointscntr1(41) = acp + 104.96
  pointscntr1(42) = bcp + 0:         pointscntr1(43) = acp + 180.75
  pointscntr1(44) = bcp + 7.9846:    pointscntr1(45) = acp + 104.96
  pointscntr1(46) = bcp + 3.99478:   pointscntr1(47) = acp + 103.894
  pointscntr1(48) = bcp + 19.1501:   pointscntr1(49) = acp + 75.0415
  pointscntr1(50) = bcp + 27.5594:   pointscntr1(51) = acp + 42.4202
  pointscntr1(52) = bcp + 17.5444:   pointscntr1(53) = acp + 19.634
  pointscntr1(54) = bcp + 12.4984:   pointscntr1(55) = acp + 12.3831
  
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  
  ' Find the bulge of the third segment
    currentBulge = plineObj.GetBulge(2)
    ' Change the bulge of the third segment
    plineObj.SetBulge 0, 0.0324358
    plineObj.Update
    plineObj.SetBulge 1, -0.00917899
    plineObj.Update
    plineObj.SetBulge 2, -0.0976773
    plineObj.Update
    plineObj.SetBulge 3, -0.221543
    plineObj.Update
    plineObj.SetBulge 4, -0.247182
    plineObj.Update
    plineObj.SetBulge 5, 0.035764
    plineObj.Update
    plineObj.SetBulge 8, 0.035764
    plineObj.Update
    plineObj.SetBulge 9, -0.247182
    plineObj.Update
    plineObj.SetBulge 10, -0.221543
    plineObj.Update
    plineObj.SetBulge 11, -0.0976773
    plineObj.Update
    plineObj.SetBulge 12, -0.00917899
    plineObj.Update
    plineObj.SetBulge 13, 0.0324358
    plineObj.Update
    plineObj.SetBulge 14, 0.0324358
    plineObj.Update
    plineObj.SetBulge 15, -0.00917899
    plineObj.Update
    plineObj.SetBulge 16, -0.0976773
    plineObj.Update
    plineObj.SetBulge 17, -0.221543
    plineObj.Update
    plineObj.SetBulge 18, -0.247182
    plineObj.Update
    plineObj.SetBulge 19, 0.035764
    plineObj.Update
    plineObj.SetBulge 22, 0.035764
    plineObj.Update
    plineObj.SetBulge 23, -0.247182
    plineObj.Update
    plineObj.SetBulge 24, -0.221543
    plineObj.Update
    plineObj.SetBulge 25, -0.0976773
    plineObj.Update
    plineObj.SetBulge 26, -0.00917899
    plineObj.Update
    plineObj.SetBulge 27, 0.0324358
    plineObj.Update
    
    plineObj.Layer = "K-grav"
    plineObj.Update
    plineObj.Closed = True
  
  pointscntr2(0) = bcp + 0:         pointscntr2(1) = acp + 102.168
  pointscntr2(2) = bcp - 15.2911:   pointscntr2(3) = acp + 78.5624
  pointscntr2(4) = bcp + 15.2911:   pointscntr2(5) = acp + 78.5624
  
 Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr2)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  
    plineObj.SetBulge 0, 0.250892
    plineObj.Update
    plineObj.SetBulge 1, -0.332226
    plineObj.Update
    plineObj.SetBulge 2, 0.250892
    plineObj.Update
  
 
  RetVal = plineObj.Mirror(b1, b2)

  pointscntr2(0) = bcp + 15.7341:  pointscntr2(1) = acp + 73.5418
  pointscntr2(2) = bcp - 15.7341:  pointscntr2(3) = acp + 73.5418
  pointscntr2(4) = bcp + 0:        pointscntr2(5) = acp + 26.2803
  
 Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr2)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  
    plineObj.SetBulge 0, 0.419513
    plineObj.Update
    plineObj.SetBulge 1, 0.182166
    plineObj.Update
    plineObj.SetBulge 2, 0.182166
    plineObj.Update
    RetVal = plineObj.Mirror(b1, b2)
  
  pointscntr3(0) = bcp + 0:             pointscntr3(1) = acp + 147.04
  pointscntr3(2) = bcp - 4.1528:        pointscntr3(3) = acp + 107.621
  pointscntr3(4) = bcp + 1.13687E-12:   pointscntr3(5) = acp + 106.064
  pointscntr3(6) = bcp + 4.1528:        pointscntr3(7) = acp + 107.621
  
 Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr3)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  
    plineObj.SetBulge 1, -0.0342582
    plineObj.Update
    plineObj.SetBulge 2, -0.0342582
    plineObj.Update
    RetVal = plineObj.Mirror(b1, b2)
    
  pointscntr3(0) = bcp + 0:             pointscntr3(1) = acp + 20.5705
  pointscntr3(2) = bcp + 7.88663:       pointscntr3(3) = acp + 12.0123
  pointscntr3(4) = bcp + 0:             pointscntr3(5) = acp + 1.97396
  pointscntr3(6) = bcp - 7.88663:       pointscntr3(7) = acp + 12.0123
  
 Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr3)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  
    plineObj.SetBulge 0, 0.0192317
    plineObj.Update
    plineObj.SetBulge 1, -0.0134617
    plineObj.Update
    plineObj.SetBulge 2, -0.0134617
    plineObj.Update
    plineObj.SetBulge 3, 0.0192317
    plineObj.Update
    RetVal = plineObj.Mirror(b1, b2)
  
  pointscntr3(0) = bcp - 21.3616:      pointscntr3(1) = acp + 0
  pointscntr3(2) = bcp - 10.4078:      pointscntr3(3) = acp - 9.53775
  pointscntr3(4) = bcp - 2.90986:      pointscntr3(5) = acp + 0
  pointscntr3(6) = bcp - 10.4078:      pointscntr3(7) = acp + 9.53775
  
 Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr3)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  
    plineObj.SetBulge 0, -0.0240064
    plineObj.Update
    plineObj.SetBulge 1, -0.012607
    plineObj.Update
    plineObj.SetBulge 2, -0.012607
    plineObj.Update
    plineObj.SetBulge 3, -0.0240064
    plineObj.Update
  
  RetVal = plineObj.Mirror(a1, a2)
  
  pointscntr4(0) = bcp - 24.1296:     pointscntr4(1) = acp + 43.1205
  pointscntr4(2) = bcp - 19.2635:     pointscntr4(3) = acp + 68.8882
  pointscntr4(4) = bcp - 2.17621:     pointscntr4(5) = acp + 23.4154
  pointscntr4(6) = bcp - 9.97906:     pointscntr4(7) = acp + 14.8729
  pointscntr4(8) = bcp - 14.6543:     pointscntr4(9) = acp + 21.6084
  
 Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr4)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  
    plineObj.SetBulge 0, -0.186466
    plineObj.Update
    plineObj.SetBulge 1, 0.168477
    plineObj.Update
    plineObj.SetBulge 2, -0.0195719
    plineObj.Update
    plineObj.SetBulge 3, -0.00864503
    plineObj.Update
    plineObj.SetBulge 4, -0.0975544
    plineObj.Update
  RetVal = plineObj.Copy
  RetVal = plineObj.Mirror(a1, a2)
  RetVal = plineObj.Mirror(b1, b2)
  
 ' Define the rotation of 45 degrees about a
  ' base point of (4, 4.25, 0)
  basePoint(0) = points(4) - (b / 2): basePoint(1) = points(3) - (a / 2)
  rotationAngle = 3.14159   ' 45 degrees

  ' Rotate the polyline
  plineObj.Rotate basePoint, rotationAngle
  plineObj.Update
End If
End If
End If
End If

  
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++

I = 100
  
 Dim plineObj2 As AcadLWPolyline
  Dim points2(0 To 7) As Double
  
  points2(0) = points(0):    points2(1) = a + I:
  points2(2) = points(0):    points2(3) = (a * 2) + I:
  points2(4) = points(6):    points2(5) = (a * 2) + I:
  points2(6) = points(6):    points2(7) = a + I:
  
 For d = 2 To Cells(c, 4)
If d > 1 Then

Set plineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(points2)
  plineObj2.Closed = True
  
If a >= 185 Then
If b >= 185 Then
  
'=========================Deep engraving===========================================
  
  pointsdeep(0) = points2(0) + 49:      pointsdeep(1) = points2(1) + 91.3
  pointsdeep(2) = points2(0) + 49:      pointsdeep(3) = points2(3) - 91.3
  pointsdeep(4) = points2(0) + 91.3:    pointsdeep(5) = points2(3) - 49
  pointsdeep(6) = points2(4) - 91.3:    pointsdeep(7) = points2(3) - 49
  pointsdeep(8) = points2(4) - 49:      pointsdeep(9) = points2(3) - 91.3
  pointsdeep(10) = points2(4) - 49:     pointsdeep(11) = points2(1) + 91.3
  pointsdeep(12) = points2(4) - 91.3:   pointsdeep(13) = points2(1) + 49
  pointsdeep(14) = points2(0) + 91.3:   pointsdeep(15) = points2(1) + 49

plineObjdeep.Layer = "K-grav"
plineObjdeep.Update
Set plineObjdeep = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsdeep)
  plineObjdeep.Closed = True

plineObjdeep.Layer = "C-Mill"
plineObjdeep.Update
  offsetObj = plineObjdeep.Offset(20)
plineObjdeep.Layer = "0"
plineObjdeep.Update

End If
End If

If a >= 145 Then
If b >= 145 Then

'===================Pattern outline on the frame===================================

pointsktr(0) = points2(0) + 42:    pointsktr(1) = points2(1) + 30
pointsktr(2) = points2(0) + 42:    pointsktr(3) = points2(1) + 36
pointsktr(4) = points2(0) + 39:    pointsktr(5) = points2(1) + 36
pointsktr(6) = points2(0) + 39:    pointsktr(7) = points2(1) + 30
pointsktr(8) = points2(0) + 30:    pointsktr(9) = points2(1) + 30
pointsktr(10) = points2(0) + 30:   pointsktr(11) = points2(1) + 39
pointsktr(12) = points2(0) + 36:   pointsktr(13) = points2(1) + 39
pointsktr(14) = points2(0) + 36:   pointsktr(15) = points2(1) + 42
pointsktr(16) = points2(0) + 30:   pointsktr(17) = points2(1) + 42
pointsktr(18) = points2(0) + 30:   pointsktr(19) = points2(3) - 42
pointsktr(20) = points2(0) + 36:   pointsktr(21) = points2(3) - 42
pointsktr(22) = points2(0) + 36:   pointsktr(23) = points2(3) - 39
pointsktr(24) = points2(0) + 30:   pointsktr(25) = points2(3) - 39
pointsktr(26) = points2(0) + 30:   pointsktr(27) = points2(3) - 30
pointsktr(28) = points2(0) + 39:   pointsktr(29) = points2(3) - 30
pointsktr(30) = points2(0) + 39:   pointsktr(31) = points2(3) - 36
pointsktr(32) = points2(0) + 42:   pointsktr(33) = points2(3) - 36
pointsktr(34) = points2(0) + 42:   pointsktr(35) = points2(3) - 30
pointsktr(36) = points2(4) - 42:   pointsktr(37) = points2(3) - 30
pointsktr(38) = points2(4) - 42:   pointsktr(39) = points2(3) - 36
pointsktr(40) = points2(4) - 39:   pointsktr(41) = points2(3) - 36
pointsktr(42) = points2(4) - 39:   pointsktr(43) = points2(3) - 30
pointsktr(44) = points2(4) - 30:   pointsktr(45) = points2(3) - 30
pointsktr(46) = points2(4) - 30:   pointsktr(47) = points2(3) - 39
pointsktr(48) = points2(4) - 36:   pointsktr(49) = points2(3) - 39
pointsktr(50) = points2(4) - 36:   pointsktr(51) = points2(3) - 42
pointsktr(52) = points2(4) - 30:   pointsktr(53) = points2(3) - 42
pointsktr(54) = points2(4) - 30:   pointsktr(55) = points2(1) + 42
pointsktr(56) = points2(4) - 36:   pointsktr(57) = points2(1) + 42
pointsktr(58) = points2(4) - 36:   pointsktr(59) = points2(1) + 39
pointsktr(60) = points2(4) - 30:   pointsktr(61) = points2(1) + 39
pointsktr(62) = points2(4) - 30:   pointsktr(63) = points2(1) + 30
pointsktr(64) = points2(4) - 39:   pointsktr(65) = points2(1) + 30
pointsktr(66) = points2(4) - 39:   pointsktr(67) = points2(1) + 36
pointsktr(68) = points2(4) - 42:   pointsktr(69) = points2(1) + 36
pointsktr(70) = points2(4) - 42:   pointsktr(71) = points2(1) + 30

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsktr)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update

pointsktr2(0) = points2(0) + 45:         pointsktr2(1) = points2(1) + 33
pointsktr2(2) = points2(0) + 45:         pointsktr2(3) = points2(1) + 36
pointsktr2(4) = points2(0) + 70.15:      pointsktr2(5) = points2(1) + 36
pointsktr2(6) = points2(0) + 68.8884:    pointsktr2(7) = points2(1) + 36.914
pointsktr2(8) = points2(0) + 67.4818:    pointsktr2(9) = points2(1) + 37.4156
pointsktr2(10) = points2(0) + 65.6556:   pointsktr2(11) = points2(1) + 38.4118
pointsktr2(12) = points2(0) + 64.0803:   pointsktr2(13) = points2(1) + 39.9661
pointsktr2(14) = points2(0) + 61.0444:   pointsktr2(15) = points2(1) + 42.5889
pointsktr2(16) = points2(0) + 57.0324:   pointsktr2(17) = points2(1) + 44.4867
pointsktr2(18) = points2(0) + 52.0524:   pointsktr2(19) = points2(1) + 45.3159
pointsktr2(20) = points2(0) + 52.7992:   pointsktr2(21) = points2(1) + 46.9026
pointsktr2(22) = points2(0) + 53.236:    pointsktr2(23) = points2(1) + 48.4263
pointsktr2(24) = points2(0) + 53.3051:   pointsktr2(25) = points2(1) + 49.0892
pointsktr2(26) = points2(0) + 53.5541:   pointsktr2(27) = points2(1) + 50.0059
pointsktr2(28) = points2(0) + 53.773:    pointsktr2(29) = points2(1) + 50.7406
pointsktr2(30) = points2(0) + 54.3842:   pointsktr2(31) = points2(1) + 53.0168
pointsktr2(32) = points2(0) + 54.5717:   pointsktr2(33) = points2(1) + 54.154
pointsktr2(34) = points2(0) + 54.5256:   pointsktr2(35) = points2(1) + 54.5256
pointsktr2(36) = points2(0) + 54.154:    pointsktr2(37) = points2(1) + 54.5717
pointsktr2(38) = points2(0) + 53.0168:   pointsktr2(39) = points2(1) + 54.3842
pointsktr2(40) = points2(0) + 50.7406:   pointsktr2(41) = points2(1) + 53.773
pointsktr2(42) = points2(0) + 50.0059:   pointsktr2(43) = points2(1) + 53.5541
pointsktr2(44) = points2(0) + 49.0892:   pointsktr2(45) = points2(1) + 53.3051
pointsktr2(46) = points2(0) + 48.4263:   pointsktr2(47) = points2(1) + 53.236
pointsktr2(48) = points2(0) + 46.9026:   pointsktr2(49) = points2(1) + 52.7992
pointsktr2(50) = points2(0) + 45.3159:   pointsktr2(51) = points2(1) + 52.0524
pointsktr2(52) = points2(0) + 44.4867:   pointsktr2(53) = points2(1) + 57.0324
pointsktr2(54) = points2(0) + 42.5889:   pointsktr2(55) = points2(1) + 61.0444
pointsktr2(56) = points2(0) + 39.9661:   pointsktr2(57) = points2(1) + 64.0803
pointsktr2(58) = points2(0) + 38.4118:   pointsktr2(59) = points2(1) + 65.6556
pointsktr2(60) = points2(0) + 37.4156:   pointsktr2(61) = points2(1) + 67.4818
pointsktr2(62) = points2(0) + 36.914:    pointsktr2(63) = points2(1) + 68.8884
pointsktr2(64) = points2(0) + 36:        pointsktr2(65) = points2(1) + 70.15
pointsktr2(66) = points2(0) + 36:        pointsktr2(67) = points2(1) + 45
pointsktr2(68) = points2(0) + 33:        pointsktr2(69) = points2(1) + 45
pointsktr2(70) = points2(0) + 33:        pointsktr2(71) = points2(3) - 45
pointsktr2(72) = points2(0) + 36:       pointsktr2(73) = points2(3) - 45
pointsktr2(74) = points2(0) + 36:       pointsktr2(75) = points2(3) - 70.15
pointsktr2(76) = points2(0) + 36.914:   pointsktr2(77) = points2(3) - 68.8884
pointsktr2(78) = points2(0) + 37.4156:  pointsktr2(79) = points2(3) - 67.4818
pointsktr2(80) = points2(0) + 38.4118:  pointsktr2(81) = points2(3) - 65.6556
pointsktr2(82) = points2(0) + 39.9661:  pointsktr2(83) = points2(3) - 64.0803
pointsktr2(84) = points2(0) + 42.5889:  pointsktr2(85) = points2(3) - 61.0444
pointsktr2(86) = points2(0) + 44.4867:  pointsktr2(87) = points2(3) - 57.0324
pointsktr2(88) = points2(0) + 45.3159:  pointsktr2(89) = points2(3) - 52.0524
pointsktr2(90) = points2(0) + 46.9026:  pointsktr2(91) = points2(3) - 52.7992
pointsktr2(92) = points2(0) + 48.4263:  pointsktr2(93) = points2(3) - 53.236
pointsktr2(94) = points2(0) + 49.0892:  pointsktr2(95) = points2(3) - 53.3051
pointsktr2(96) = points2(0) + 50.0059:  pointsktr2(97) = points2(3) - 53.5541
pointsktr2(98) = points2(0) + 50.7406:  pointsktr2(99) = points2(3) - 53.773
pointsktr2(100) = points2(0) + 53.0168: pointsktr2(101) = points2(3) - 54.3842
pointsktr2(102) = points2(0) + 54.154:  pointsktr2(103) = points2(3) - 54.5717
pointsktr2(104) = points2(0) + 54.5256: pointsktr2(105) = points2(3) - 54.5256
pointsktr2(106) = points2(0) + 54.5717: pointsktr2(107) = points2(3) - 54.154
pointsktr2(108) = points2(0) + 54.3842: pointsktr2(109) = points2(3) - 53.0168
pointsktr2(110) = points2(0) + 53.773:  pointsktr2(111) = points2(3) - 50.7406
pointsktr2(112) = points2(0) + 53.5541: pointsktr2(113) = points2(3) - 50.0059
pointsktr2(114) = points2(0) + 53.3051: pointsktr2(115) = points2(3) - 49.0892
pointsktr2(116) = points2(0) + 53.236:  pointsktr2(117) = points2(3) - 48.4263
pointsktr2(118) = points2(0) + 52.7992: pointsktr2(119) = points2(3) - 46.9026
pointsktr2(120) = points2(0) + 52.0524: pointsktr2(121) = points2(3) - 45.3159
pointsktr2(122) = points2(0) + 57.0324: pointsktr2(123) = points2(3) - 44.4867
pointsktr2(124) = points2(0) + 61.0444: pointsktr2(125) = points2(3) - 42.5889
pointsktr2(126) = points2(0) + 64.0803: pointsktr2(127) = points2(3) - 39.9661
pointsktr2(128) = points2(0) + 65.6556: pointsktr2(129) = points2(3) - 38.4118
pointsktr2(130) = points2(0) + 67.4818: pointsktr2(131) = points2(3) - 37.4156
pointsktr2(132) = points2(0) + 68.8884: pointsktr2(133) = points2(3) - 36.914
pointsktr2(134) = points2(0) + 70.15:   pointsktr2(135) = points2(3) - 36
pointsktr2(136) = points2(0) + 45:      pointsktr2(137) = points2(3) - 36
pointsktr2(138) = points2(0) + 45:      pointsktr2(139) = points2(3) - 33
pointsktr2(140) = points2(4) - 45:      pointsktr2(141) = points2(3) - 33
pointsktr2(142) = points2(4) - 45:      pointsktr2(143) = points2(3) - 36
pointsktr2(144) = points2(4) - 70.15:   pointsktr2(145) = points2(3) - 36
pointsktr2(146) = points2(4) - 68.8884: pointsktr2(147) = points2(3) - 36.914
pointsktr2(148) = points2(4) - 67.4818: pointsktr2(149) = points2(3) - 37.4156
pointsktr2(150) = points2(4) - 65.6556: pointsktr2(151) = points2(3) - 38.4118
pointsktr2(152) = points2(4) - 64.0803: pointsktr2(153) = points2(3) - 39.9661
pointsktr2(154) = points2(4) - 61.0444: pointsktr2(155) = points2(3) - 42.5889
pointsktr2(156) = points2(4) - 57.0324: pointsktr2(157) = points2(3) - 44.4867
pointsktr2(158) = points2(4) - 52.0524: pointsktr2(159) = points2(3) - 45.3159
pointsktr2(160) = points2(4) - 52.7992: pointsktr2(161) = points2(3) - 46.9026
pointsktr2(162) = points2(4) - 53.236:  pointsktr2(163) = points2(3) - 48.4263
pointsktr2(164) = points2(4) - 53.3051: pointsktr2(165) = points2(3) - 49.0892
pointsktr2(166) = points2(4) - 53.5541: pointsktr2(167) = points2(3) - 50.0059
pointsktr2(168) = points2(4) - 53.773:  pointsktr2(169) = points2(3) - 50.7406
pointsktr2(170) = points2(4) - 54.3842: pointsktr2(171) = points2(3) - 53.0168
pointsktr2(172) = points2(4) - 54.5717: pointsktr2(173) = points2(3) - 54.154
pointsktr2(174) = points2(4) - 54.5256: pointsktr2(175) = points2(3) - 54.5256
pointsktr2(176) = points2(4) - 54.154:  pointsktr2(177) = points2(3) - 54.5717
pointsktr2(178) = points2(4) - 53.0168: pointsktr2(179) = points2(3) - 54.3842
pointsktr2(180) = points2(4) - 50.7406: pointsktr2(181) = points2(3) - 53.773
pointsktr2(182) = points2(4) - 50.0059: pointsktr2(183) = points2(3) - 53.5541
pointsktr2(184) = points2(4) - 49.0892: pointsktr2(185) = points2(3) - 53.3051
pointsktr2(186) = points2(4) - 48.4263: pointsktr2(187) = points2(3) - 53.236
pointsktr2(188) = points2(4) - 46.9026: pointsktr2(189) = points2(3) - 52.7992
pointsktr2(190) = points2(4) - 45.3159: pointsktr2(191) = points2(3) - 52.0524
pointsktr2(192) = points2(4) - 44.4867: pointsktr2(193) = points2(3) - 57.0324
pointsktr2(194) = points2(4) - 42.5889: pointsktr2(195) = points2(3) - 61.0444
pointsktr2(196) = points2(4) - 39.9661: pointsktr2(197) = points2(3) - 64.0803
pointsktr2(198) = points2(4) - 38.4118: pointsktr2(199) = points2(3) - 65.6556
pointsktr2(200) = points2(4) - 37.4156: pointsktr2(201) = points2(3) - 67.4818
pointsktr2(202) = points2(4) - 36.914:  pointsktr2(203) = points2(3) - 68.8884
pointsktr2(204) = points2(4) - 36:      pointsktr2(205) = points2(3) - 70.15
pointsktr2(206) = points2(4) - 36:      pointsktr2(207) = points2(3) - 45
pointsktr2(208) = points2(4) - 33:      pointsktr2(209) = points2(3) - 45
pointsktr2(210) = points2(4) - 33:      pointsktr2(211) = points2(1) + 45
pointsktr2(212) = points2(4) - 36:      pointsktr2(213) = points2(1) + 45
pointsktr2(214) = points2(4) - 36:      pointsktr2(215) = points2(1) + 70.15
pointsktr2(216) = points2(4) - 36.914:  pointsktr2(217) = points2(1) + 68.8884
pointsktr2(218) = points2(4) - 37.4156: pointsktr2(219) = points2(1) + 67.4818
pointsktr2(220) = points2(4) - 38.4118: pointsktr2(221) = points2(1) + 65.6556
pointsktr2(222) = points2(4) - 39.9661: pointsktr2(223) = points2(1) + 64.0803
pointsktr2(224) = points2(4) - 42.5889: pointsktr2(225) = points2(1) + 61.0444
pointsktr2(226) = points2(4) - 44.4867: pointsktr2(227) = points2(1) + 57.0324
pointsktr2(228) = points2(4) - 45.3159: pointsktr2(229) = points2(1) + 52.0524
pointsktr2(230) = points2(4) - 46.9026: pointsktr2(231) = points2(1) + 52.7992
pointsktr2(232) = points2(4) - 48.4263: pointsktr2(233) = points2(1) + 53.236
pointsktr2(234) = points2(4) - 49.0892: pointsktr2(235) = points2(1) + 53.3051
pointsktr2(236) = points2(4) - 50.0059: pointsktr2(237) = points2(1) + 53.5541
pointsktr2(238) = points2(4) - 50.7406: pointsktr2(239) = points2(1) + 53.773
pointsktr2(240) = points2(4) - 53.0168: pointsktr2(241) = points2(1) + 54.3842
pointsktr2(242) = points2(4) - 54.154:  pointsktr2(243) = points2(1) + 54.5717
pointsktr2(244) = points2(4) - 54.5256: pointsktr2(245) = points2(1) + 54.5256
pointsktr2(246) = points2(4) - 54.5717: pointsktr2(247) = points2(1) + 54.154
pointsktr2(248) = points2(4) - 54.3842: pointsktr2(249) = points2(1) + 53.0168
pointsktr2(250) = points2(4) - 53.773:  pointsktr2(251) = points2(1) + 50.7406
pointsktr2(252) = points2(4) - 53.5541: pointsktr2(253) = points2(1) + 50.0059
pointsktr2(254) = points2(4) - 53.3051: pointsktr2(255) = points2(1) + 49.0892
pointsktr2(256) = points2(4) - 53.236:  pointsktr2(257) = points2(1) + 48.4263
pointsktr2(258) = points2(4) - 52.7992: pointsktr2(259) = points2(1) + 46.9026
pointsktr2(260) = points2(4) - 52.0524: pointsktr2(261) = points2(1) + 45.3159
pointsktr2(262) = points2(4) - 57.0324: pointsktr2(263) = points2(1) + 44.4867
pointsktr2(264) = points2(4) - 61.0444: pointsktr2(265) = points2(1) + 42.5889
pointsktr2(266) = points2(4) - 64.0803: pointsktr2(267) = points2(1) + 39.9661
pointsktr2(268) = points2(4) - 65.6556: pointsktr2(269) = points2(1) + 38.4118
pointsktr2(270) = points2(4) - 67.4818: pointsktr2(271) = points2(1) + 37.4156
pointsktr2(272) = points2(4) - 68.8884: pointsktr2(273) = points2(1) + 36.914
pointsktr2(274) = points2(4) - 70.15:   pointsktr2(275) = points2(1) + 36
pointsktr2(276) = points2(4) - 45:      pointsktr2(277) = points2(1) + 36
pointsktr2(278) = points2(4) - 45:      pointsktr2(279) = points2(1) + 33

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsktr2)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  
plineObj.SetBulge 0, 0
plineObj.SetBulge 1, 0
plineObj.SetBulge 2, 0.0452849
plineObj.SetBulge 3, 0.0767779
plineObj.SetBulge 4, -0.125063
plineObj.SetBulge 5, -0.0179955
plineObj.SetBulge 6, 0.0434551
plineObj.SetBulge 7, 0.0833797
plineObj.SetBulge 8, 0.0493602
plineObj.SetBulge 9, -0.0221733
plineObj.SetBulge 10, 0.0927081
plineObj.SetBulge 11, 0
plineObj.SetBulge 12, -0.104648
plineObj.SetBulge 13, 0.0311187
plineObj.SetBulge 14, -0.0452272
plineObj.SetBulge 15, 0.0707349
plineObj.SetBulge 16, 0.0483396
plineObj.SetBulge 17, 0.0483396
plineObj.SetBulge 18, 0.0707349
plineObj.SetBulge 19, -0.0452272
plineObj.SetBulge 20, 0.0311187
plineObj.SetBulge 21, -0.104648
plineObj.SetBulge 22, 0
plineObj.SetBulge 23, 0.0927081
plineObj.SetBulge 24, -0.0221733
plineObj.SetBulge 25, 0.0493602
plineObj.SetBulge 26, 0.0833797
plineObj.SetBulge 27, 0.0434551
plineObj.SetBulge 28, -0.0179955
plineObj.SetBulge 29, -0.125063
plineObj.SetBulge 30, 0.0767779
plineObj.SetBulge 31, 0.0452849
plineObj.SetBulge 32, 0
plineObj.SetBulge 33, 0
plineObj.SetBulge 34, 0
plineObj.SetBulge 35, 0
plineObj.SetBulge 36, 0
plineObj.SetBulge 37, 0.0452849
plineObj.SetBulge 38, 0.0767779
plineObj.SetBulge 39, -0.125063
plineObj.SetBulge 40, -0.0179955
plineObj.SetBulge 41, 0.0434551
plineObj.SetBulge 42, 0.0833797
plineObj.SetBulge 43, 0.0493602
plineObj.SetBulge 44, -0.0221733
plineObj.SetBulge 45, 0.0927081
plineObj.SetBulge 46, 0
plineObj.SetBulge 47, -0.104648
plineObj.SetBulge 48, 0.0311187
plineObj.SetBulge 49, -0.0452272
plineObj.SetBulge 50, 0.0707349
plineObj.SetBulge 51, 0.0483396
plineObj.SetBulge 52, 0.0483396
plineObj.SetBulge 53, 0.0707349
plineObj.SetBulge 54, -0.0452272
plineObj.SetBulge 55, 0.0311187
plineObj.SetBulge 56, -0.104648
plineObj.SetBulge 57, 0
plineObj.SetBulge 58, 0.0927081
plineObj.SetBulge 59, -0.0221733
plineObj.SetBulge 60, 0.0493602
plineObj.SetBulge 61, 0.0833797
plineObj.SetBulge 62, 0.0434551
plineObj.SetBulge 63, -0.0179955
plineObj.SetBulge 64, -0.125063
plineObj.SetBulge 65, 0.0767779
plineObj.SetBulge 66, 0.0452849
plineObj.SetBulge 67, 0
plineObj.SetBulge 68, 0
plineObj.SetBulge 69, 0
plineObj.SetBulge 70, 0
plineObj.SetBulge 71, 0
plineObj.SetBulge 72, 0.0452849
plineObj.SetBulge 73, 0.0767779
plineObj.SetBulge 74, -0.125063
plineObj.SetBulge 75, -0.0179955
plineObj.SetBulge 76, 0.0434551
plineObj.SetBulge 77, 0.0833797
plineObj.SetBulge 78, 0.0493602
plineObj.SetBulge 79, -0.0221733
plineObj.SetBulge 80, 0.0927081
plineObj.SetBulge 81, 0
plineObj.SetBulge 82, -0.104648
plineObj.SetBulge 83, 0.0311187
plineObj.SetBulge 84, -0.0452272
plineObj.SetBulge 85, 0.0707349
plineObj.SetBulge 86, 0.0483396
plineObj.SetBulge 87, 0.0483396
plineObj.SetBulge 88, 0.0707349
plineObj.SetBulge 89, -0.0452272
plineObj.SetBulge 90, 0.0311187
plineObj.SetBulge 91, -0.104648
plineObj.SetBulge 92, 0
plineObj.SetBulge 93, 0.0927081
plineObj.SetBulge 94, -0.0221733
plineObj.SetBulge 95, 0.0493602
plineObj.SetBulge 96, 0.0833797
plineObj.SetBulge 97, 0.0434551
plineObj.SetBulge 98, -0.0179955
plineObj.SetBulge 99, -0.125063
plineObj.SetBulge 100, 0.0767779
plineObj.SetBulge 101, 0.0452849
plineObj.SetBulge 102, 0
plineObj.SetBulge 103, 0
plineObj.SetBulge 104, 0
plineObj.SetBulge 105, 0
plineObj.SetBulge 106, 0
plineObj.SetBulge 107, 0.0452849
plineObj.SetBulge 108, 0.0767779
plineObj.SetBulge 109, -0.125063
plineObj.SetBulge 110, -0.0179955
plineObj.SetBulge 111, 0.0434551
plineObj.SetBulge 112, 0.0833797
plineObj.SetBulge 113, 0.0493602
plineObj.SetBulge 114, -0.0221733
plineObj.SetBulge 115, 0.0927081
plineObj.SetBulge 116, 0
plineObj.SetBulge 117, -0.104648
plineObj.SetBulge 118, 0.0311187
plineObj.SetBulge 119, -0.0452272
plineObj.SetBulge 120, 0.0707349
plineObj.SetBulge 121, 0.0483396
plineObj.SetBulge 122, 0.0483396
plineObj.SetBulge 123, 0.0707349
plineObj.SetBulge 124, -0.0452272
plineObj.SetBulge 125, 0.0311187
plineObj.SetBulge 126, -0.104648
plineObj.SetBulge 127, 0
plineObj.SetBulge 128, 0.0927081
plineObj.SetBulge 129, -0.0221733
plineObj.SetBulge 130, 0.0493602
plineObj.SetBulge 131, 0.0833797
plineObj.SetBulge 132, 0.0434551
plineObj.SetBulge 133, -0.0179955
plineObj.SetBulge 134, -0.125063
plineObj.SetBulge 135, 0.0767779
plineObj.SetBulge 136, 0.0452849
plineObj.SetBulge 137, 0
plineObj.SetBulge 138, 0
plineObj.SetBulge 139, 0
plineObj.Update
  
'=========================Stripes======================================
  a1(0) = points2(4) - (b / 2):       a1(1) = points2(1) + ((a / 2) + 1)
  a2(0) = points2(4) - (b / 2):       a2(1) = points2(1) + ((a / 2) - 1)
  b1(0) = points2(0) + ((b / 2) - 1): b1(1) = points2(1) + (a / 2)
  b2(0) = points2(0) + ((b / 2) + 1): b2(1) = points2(1) + (a / 2)

  pointsstr1(0) = points2(0) + 42:     pointsstr1(1) = points2(1) + 55.6455
  pointsstr1(2) = points2(0) + 42:     pointsstr1(3) = points2(1) + 45
  pointsstr1(4) = points2(0) + 39:     pointsstr1(5) = points2(1) + 45
  pointsstr1(6) = points2(0) + 39:     pointsstr1(7) = points2(1) + 60.5518
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  plineObj.SetBulge 3, -0.133444
  plineObj.Update
  RetVal = plineObj.Mirror(a1, a2)
  RetVal = plineObj.Mirror(b1, b2)
  RetVal = plineObj.Copy
  basePoint(0) = points2(4) - (b / 2): basePoint(1) = points2(3) - (a / 2)
  rotationAngle = 3.14159   ' 45 degrees
  plineObj.Rotate basePoint, rotationAngle
  plineObj.Update
  
  pointsstr1(0) = points2(0) + 55.6455:   pointsstr1(1) = points2(1) + 42
  pointsstr1(2) = points2(0) + 45:        pointsstr1(3) = points2(1) + 42
  pointsstr1(4) = points2(0) + 45:        pointsstr1(5) = points2(1) + 39
  pointsstr1(6) = points2(0) + 60.5518:   pointsstr1(7) = points2(1) + 39
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  plineObj.SetBulge 3, 0.133444
  plineObj.Update
  RetVal = plineObj.Mirror(a1, a2)
  RetVal = plineObj.Mirror(b1, b2)
  RetVal = plineObj.Copy
  basePoint(0) = points2(4) - (b / 2): basePoint(1) = points2(3) - (a / 2)
  rotationAngle = 3.14159   ' 45 degrees
  plineObj.Rotate basePoint, rotationAngle
  plineObj.Update
  
'=========================Squares======================================
  pointsstr1(0) = points2(0) + 33:   pointsstr1(1) = points2(1) + 33
  pointsstr1(2) = points2(0) + 33:   pointsstr1(3) = points2(1) + 36
  pointsstr1(4) = points2(0) + 36:   pointsstr1(5) = points2(1) + 36
  pointsstr1(6) = points2(0) + 36:   pointsstr1(7) = points2(1) + 33
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  pointsstr1(0) = points2(0) + 39:   pointsstr1(1) = points2(1) + 39
  pointsstr1(2) = points2(0) + 39:   pointsstr1(3) = points2(1) + 42
  pointsstr1(4) = points2(0) + 42:   pointsstr1(5) = points2(1) + 42
  pointsstr1(6) = points2(0) + 42:   pointsstr1(7) = points2(1) + 39
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  '***********************
  pointsstr1(0) = points2(4) - 33:   pointsstr1(1) = points2(1) + 33
  pointsstr1(2) = points2(4) - 33:   pointsstr1(3) = points2(1) + 36
  pointsstr1(4) = points2(4) - 36:   pointsstr1(5) = points2(1) + 36
  pointsstr1(6) = points2(4) - 36:   pointsstr1(7) = points2(1) + 33
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  pointsstr1(0) = points2(4) - 39:   pointsstr1(1) = points2(1) + 39
  pointsstr1(2) = points2(4) - 39:   pointsstr1(3) = points2(1) + 42
  pointsstr1(4) = points2(4) - 42:   pointsstr1(5) = points2(1) + 42
  pointsstr1(6) = points2(4) - 42:   pointsstr1(7) = points2(1) + 39
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  '***********************
  pointsstr1(0) = points2(0) + 33:   pointsstr1(1) = points2(3) - 33
  pointsstr1(2) = points2(0) + 33:   pointsstr1(3) = points2(3) - 36
  pointsstr1(4) = points2(0) + 36:   pointsstr1(5) = points2(3) - 36
  pointsstr1(6) = points2(0) + 36:   pointsstr1(7) = points2(3) - 33
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  pointsstr1(0) = points2(0) + 39:   pointsstr1(1) = points2(3) - 39
  pointsstr1(2) = points2(0) + 39:   pointsstr1(3) = points2(3) - 42
  pointsstr1(4) = points2(0) + 42:   pointsstr1(5) = points2(3) - 42
  pointsstr1(6) = points2(0) + 42:   pointsstr1(7) = points2(3) - 39
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  '***********************
  pointsstr1(0) = points2(4) - 33:   pointsstr1(1) = points2(3) - 33
  pointsstr1(2) = points2(4) - 33:   pointsstr1(3) = points2(3) - 36
  pointsstr1(4) = points2(4) - 36:   pointsstr1(5) = points2(3) - 36
  pointsstr1(6) = points2(4) - 36:   pointsstr1(7) = points2(3) - 33
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  pointsstr1(0) = points2(4) - 39:   pointsstr1(1) = points2(3) - 39
  pointsstr1(2) = points2(4) - 39:   pointsstr1(3) = points2(3) - 42
  pointsstr1(4) = points2(4) - 42:   pointsstr1(5) = points2(3) - 42
  pointsstr1(6) = points2(4) - 42:   pointsstr1(7) = points2(3) - 39
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
'===========================Oval====================================

pointsstr2(0) = points2(0) + 50.5415:    pointsstr2(1) = points2(1) + 48.0666
pointsstr2(2) = points2(0) + 50.1973:    pointsstr2(3) = points2(1) + 47.76
pointsstr2(4) = points2(0) + 49.8259:    pointsstr2(5) = points2(1) + 47.5003
pointsstr2(6) = points2(0) + 49.4388:    pointsstr2(7) = points2(1) + 47.2954
pointsstr2(8) = points2(0) + 49.0475:    pointsstr2(9) = points2(1) + 47.1516
pointsstr2(10) = points2(0) + 48.664:    pointsstr2(11) = points2(1) + 47.0732
pointsstr2(12) = points2(0) + 48.2999:   pointsstr2(13) = points2(1) + 47.0625
pointsstr2(14) = points2(0) + 47.9664:   pointsstr2(15) = points2(1) + 47.12
pointsstr2(16) = points2(0) + 47.6735:   pointsstr2(17) = points2(1) + 47.2438
pointsstr2(18) = points2(0) + 47.4302:   pointsstr2(19) = points2(1) + 47.4302
pointsstr2(20) = points2(0) + 47.2438:   pointsstr2(21) = points2(1) + 47.6735
pointsstr2(22) = points2(0) + 47.12:     pointsstr2(23) = points2(1) + 47.9664
pointsstr2(24) = points2(0) + 47.0625:   pointsstr2(25) = points2(1) + 48.2999
pointsstr2(26) = points2(0) + 47.0732:   pointsstr2(27) = points2(1) + 48.664
pointsstr2(28) = points2(0) + 47.1516:   pointsstr2(29) = points2(1) + 49.0475
pointsstr2(30) = points2(0) + 47.2954:   pointsstr2(31) = points2(1) + 49.4388
pointsstr2(32) = points2(0) + 47.5003:   pointsstr2(33) = points2(1) + 49.8259
pointsstr2(34) = points2(0) + 47.76:     pointsstr2(35) = points2(1) + 50.1973
pointsstr2(36) = points2(0) + 48.0666:   pointsstr2(37) = points2(1) + 50.5415
pointsstr2(38) = points2(0) + 48.4108:   pointsstr2(39) = points2(1) + 50.848
pointsstr2(40) = points2(0) + 48.7821:   pointsstr2(41) = points2(1) + 51.1077
pointsstr2(42) = points2(0) + 49.1693:   pointsstr2(43) = points2(1) + 51.3126
pointsstr2(44) = points2(0) + 49.5606:   pointsstr2(45) = points2(1) + 51.4564
pointsstr2(46) = points2(0) + 49.944:    pointsstr2(47) = points2(1) + 51.5349
pointsstr2(48) = points2(0) + 50.3081:   pointsstr2(49) = points2(1) + 51.5455
pointsstr2(50) = points2(0) + 50.6416:   pointsstr2(51) = points2(1) + 51.4881
pointsstr2(52) = points2(0) + 50.9345:   pointsstr2(53) = points2(1) + 51.3643
pointsstr2(54) = points2(0) + 51.1778:   pointsstr2(55) = points2(1) + 51.1778
pointsstr2(56) = points2(0) + 51.3643:   pointsstr2(57) = points2(1) + 50.9345
pointsstr2(58) = points2(0) + 51.4881:   pointsstr2(59) = points2(1) + 50.6416
pointsstr2(60) = points2(0) + 51.5455:   pointsstr2(61) = points2(1) + 50.3081
pointsstr2(62) = points2(0) + 51.5349:   pointsstr2(63) = points2(1) + 49.944
pointsstr2(64) = points2(0) + 51.4564:   pointsstr2(65) = points2(1) + 49.5606
pointsstr2(66) = points2(0) + 51.3126:   pointsstr2(67) = points2(1) + 49.1693
pointsstr2(68) = points2(0) + 51.1077:   pointsstr2(69) = points2(1) + 48.7821
pointsstr2(70) = points2(0) + 50.848:    pointsstr2(71) = points2(1) + 48.4108

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsstr2)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update

plineObj.SetBulge 0, -0.0288637
plineObj.SetBulge 1, -0.0299658
plineObj.SetBulge 2, -0.0320654
plineObj.SetBulge 3, -0.0354072
plineObj.SetBulge 4, -0.040166
plineObj.SetBulge 5, -0.0463995
plineObj.SetBulge 6, -0.0537273
plineObj.SetBulge 7, -0.0608551
plineObj.SetBulge 8, -0.0659533
plineObj.SetBulge 9, -0.0659533
plineObj.SetBulge 10, -0.0608551
plineObj.SetBulge 11, -0.0537273
plineObj.SetBulge 12, -0.0463995
plineObj.SetBulge 13, -0.040166
plineObj.SetBulge 14, -0.0354072
plineObj.SetBulge 15, -0.0320654
plineObj.SetBulge 16, -0.0299658
plineObj.SetBulge 17, -0.0288637
plineObj.SetBulge 18, -0.0288637
plineObj.SetBulge 19, -0.0299658
plineObj.SetBulge 20, -0.0320654
plineObj.SetBulge 21, -0.0354072
plineObj.SetBulge 22, -0.040166
plineObj.SetBulge 23, -0.0463995
plineObj.SetBulge 24, -0.0537273
plineObj.SetBulge 25, -0.0608551
plineObj.SetBulge 26, -0.0659533
plineObj.SetBulge 27, -0.0659533
plineObj.SetBulge 28, -0.0608551
plineObj.SetBulge 29, -0.0537273
plineObj.SetBulge 30, -0.0463995
plineObj.SetBulge 31, -0.040166
plineObj.SetBulge 32, -0.0354072
plineObj.SetBulge 33, -0.0320654
plineObj.SetBulge 34, -0.0299658
plineObj.SetBulge 35, -0.0288637

  RetVal = plineObj.Mirror(a1, a2)
  RetVal = plineObj.Mirror(b1, b2)
  RetVal = plineObj.Copy
  basePoint(0) = points2(4) - (b / 2): basePoint(1) = points2(3) - (a / 2)
  rotationAngle = 3.14159   ' 45 degrees
  plineObj.Rotate basePoint, rotationAngle
  plineObj.Update
  
  If a > 540 Then
  If b > 250 Then
'=========================CENTER======================================
  
  bcp = points2(0) + (b / 2)
  acp = points2(1) + (a / 2)
  a1(0) = points2(4) - (b / 2):       a1(1) = points2(1) + ((a / 2) + 1)
  a2(0) = points2(4) - (b / 2):       a2(1) = points2(1) + ((a / 2) - 1)
  b1(0) = points2(0) + ((b / 2) - 1): b1(1) = points2(1) + (a / 2)
  b2(0) = points2(0) + ((b / 2) + 1): b2(1) = points2(1) + (a / 2)
 
  pointscntr1(0) = bcp + 27.1207:    pointscntr1(1) = acp - 5.68434E-14
  pointscntr1(2) = bcp + 12.4984:    pointscntr1(3) = acp - 12.3831
  pointscntr1(4) = bcp + 17.5444:    pointscntr1(5) = acp - 19.634
  pointscntr1(6) = bcp + 27.5594:    pointscntr1(7) = acp - 42.4202
  pointscntr1(8) = bcp + 19.1501:    pointscntr1(9) = acp - 75.0415
  pointscntr1(10) = bcp + 3.99478:   pointscntr1(11) = acp - 103.894
  pointscntr1(12) = bcp + 7.9846:    pointscntr1(13) = acp - 104.96
  pointscntr1(14) = bcp + 0:         pointscntr1(15) = acp - 180.75
  pointscntr1(16) = bcp - 7.9846:    pointscntr1(17) = acp - 104.96
  pointscntr1(18) = bcp - 3.99478:   pointscntr1(19) = acp - 103.894
  pointscntr1(20) = bcp - 19.1501:   pointscntr1(21) = acp - 75.0415
  pointscntr1(22) = bcp - 27.5594:   pointscntr1(23) = acp - 42.4202
  pointscntr1(24) = bcp - 17.5444:   pointscntr1(25) = acp - 19.634
  pointscntr1(26) = bcp - 12.4984:   pointscntr1(27) = acp - 12.3831
  pointscntr1(28) = bcp - 27.1207:   pointscntr1(29) = acp + 5.68434E-14
  pointscntr1(30) = bcp - 12.4984:   pointscntr1(31) = acp + 12.3831
  pointscntr1(32) = bcp - 17.5444:   pointscntr1(33) = acp + 19.634
  pointscntr1(34) = bcp - 27.5594:   pointscntr1(35) = acp + 42.4202
  pointscntr1(36) = bcp - 19.1501:   pointscntr1(37) = acp + 75.0415
  pointscntr1(38) = bcp - 3.99478:   pointscntr1(39) = acp + 103.894
  pointscntr1(40) = bcp - 7.9846:    pointscntr1(41) = acp + 104.96
  pointscntr1(42) = bcp + 0:         pointscntr1(43) = acp + 180.75
  pointscntr1(44) = bcp + 7.9846:    pointscntr1(45) = acp + 104.96
  pointscntr1(46) = bcp + 3.99478:   pointscntr1(47) = acp + 103.894
  pointscntr1(48) = bcp + 19.1501:   pointscntr1(49) = acp + 75.0415
  pointscntr1(50) = bcp + 27.5594:   pointscntr1(51) = acp + 42.4202
  pointscntr1(52) = bcp + 17.5444:   pointscntr1(53) = acp + 19.634
  pointscntr1(54) = bcp + 12.4984:   pointscntr1(55) = acp + 12.3831
  
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  
  ' Find the bulge of the third segment
    currentBulge = plineObj.GetBulge(2)
    ' Change the bulge of the third segment
    plineObj.SetBulge 0, 0.0324358
    plineObj.Update
    plineObj.SetBulge 1, -0.00917899
    plineObj.Update
    plineObj.SetBulge 2, -0.0976773
    plineObj.Update
    plineObj.SetBulge 3, -0.221543
    plineObj.Update
    plineObj.SetBulge 4, -0.247182
    plineObj.Update
    plineObj.SetBulge 5, 0.035764
    plineObj.Update
    plineObj.SetBulge 8, 0.035764
    plineObj.Update
    plineObj.SetBulge 9, -0.247182
    plineObj.Update
    plineObj.SetBulge 10, -0.221543
    plineObj.Update
    plineObj.SetBulge 11, -0.0976773
    plineObj.Update
    plineObj.SetBulge 12, -0.00917899
    plineObj.Update
    plineObj.SetBulge 13, 0.0324358
    plineObj.Update
    plineObj.SetBulge 14, 0.0324358
    plineObj.Update
    plineObj.SetBulge 15, -0.00917899
    plineObj.Update
    plineObj.SetBulge 16, -0.0976773
    plineObj.Update
    plineObj.SetBulge 17, -0.221543
    plineObj.Update
    plineObj.SetBulge 18, -0.247182
    plineObj.Update
    plineObj.SetBulge 19, 0.035764
    plineObj.Update
    plineObj.SetBulge 22, 0.035764
    plineObj.Update
    plineObj.SetBulge 23, -0.247182
    plineObj.Update
    plineObj.SetBulge 24, -0.221543
    plineObj.Update
    plineObj.SetBulge 25, -0.0976773
    plineObj.Update
    plineObj.SetBulge 26, -0.00917899
    plineObj.Update
    plineObj.SetBulge 27, 0.0324358
    plineObj.Update
    
    plineObj.Layer = "K-grav"
    plineObj.Update
    plineObj.Closed = True
  
  pointscntr2(0) = bcp + 0:         pointscntr2(1) = acp + 102.168
  pointscntr2(2) = bcp - 15.2911:   pointscntr2(3) = acp + 78.5624
  pointscntr2(4) = bcp + 15.2911:   pointscntr2(5) = acp + 78.5624
  
 Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr2)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  
    plineObj.SetBulge 0, 0.250892
    plineObj.Update
    plineObj.SetBulge 1, -0.332226
    plineObj.Update
    plineObj.SetBulge 2, 0.250892
    plineObj.Update
  
 
  RetVal = plineObj.Mirror(b1, b2)

  pointscntr2(0) = bcp + 15.7341:  pointscntr2(1) = acp + 73.5418
  pointscntr2(2) = bcp - 15.7341:  pointscntr2(3) = acp + 73.5418
  pointscntr2(4) = bcp + 0:        pointscntr2(5) = acp + 26.2803
  
 Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr2)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  
    plineObj.SetBulge 0, 0.419513
    plineObj.Update
    plineObj.SetBulge 1, 0.182166
    plineObj.Update
    plineObj.SetBulge 2, 0.182166
    plineObj.Update
    RetVal = plineObj.Mirror(b1, b2)
  
  pointscntr3(0) = bcp + 0:             pointscntr3(1) = acp + 147.04
  pointscntr3(2) = bcp - 4.1528:        pointscntr3(3) = acp + 107.621
  pointscntr3(4) = bcp + 1.13687E-12:   pointscntr3(5) = acp + 106.064
  pointscntr3(6) = bcp + 4.1528:        pointscntr3(7) = acp + 107.621
  
 Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr3)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  
    plineObj.SetBulge 1, -0.0342582
    plineObj.Update
    plineObj.SetBulge 2, -0.0342582
    plineObj.Update
    RetVal = plineObj.Mirror(b1, b2)
    
  pointscntr3(0) = bcp + 0:             pointscntr3(1) = acp + 20.5705
  pointscntr3(2) = bcp + 7.88663:       pointscntr3(3) = acp + 12.0123
  pointscntr3(4) = bcp + 0:             pointscntr3(5) = acp + 1.97396
  pointscntr3(6) = bcp - 7.88663:       pointscntr3(7) = acp + 12.0123
  
 Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr3)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  
    plineObj.SetBulge 0, 0.0192317
    plineObj.Update
    plineObj.SetBulge 1, -0.0134617
    plineObj.Update
    plineObj.SetBulge 2, -0.0134617
    plineObj.Update
    plineObj.SetBulge 3, 0.0192317
    plineObj.Update
    RetVal = plineObj.Mirror(b1, b2)
  
  pointscntr3(0) = bcp - 21.3616:      pointscntr3(1) = acp + 0
  pointscntr3(2) = bcp - 10.4078:      pointscntr3(3) = acp - 9.53775
  pointscntr3(4) = bcp - 2.90986:      pointscntr3(5) = acp + 0
  pointscntr3(6) = bcp - 10.4078:      pointscntr3(7) = acp + 9.53775
  
 Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr3)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  
    plineObj.SetBulge 0, -0.0240064
    plineObj.Update
    plineObj.SetBulge 1, -0.012607
    plineObj.Update
    plineObj.SetBulge 2, -0.012607
    plineObj.Update
    plineObj.SetBulge 3, -0.0240064
    plineObj.Update
  
  RetVal = plineObj.Mirror(a1, a2)
  
  pointscntr4(0) = bcp - 24.1296:     pointscntr4(1) = acp + 43.1205
  pointscntr4(2) = bcp - 19.2635:     pointscntr4(3) = acp + 68.8882
  pointscntr4(4) = bcp - 2.17621:     pointscntr4(5) = acp + 23.4154
  pointscntr4(6) = bcp - 9.97906:     pointscntr4(7) = acp + 14.8729
  pointscntr4(8) = bcp - 14.6543:     pointscntr4(9) = acp + 21.6084
  
 Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr4)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  plineObj.Update
  
    plineObj.SetBulge 0, -0.186466
    plineObj.Update
    plineObj.SetBulge 1, 0.168477
    plineObj.Update
    plineObj.SetBulge 2, -0.0195719
    plineObj.Update
    plineObj.SetBulge 3, -0.00864503
    plineObj.Update
    plineObj.SetBulge 4, -0.0975544
    plineObj.Update
  RetVal = plineObj.Copy
  RetVal = plineObj.Mirror(a1, a2)
  RetVal = plineObj.Mirror(b1, b2)
  
 ' Define the rotation of 45 degrees about a
  ' base point of (4, 4.25, 0)
  basePoint(0) = points2(4) - (b / 2): basePoint(1) = points2(3) - (a / 2)
  rotationAngle = 3.14159   ' 45 degrees

  ' Rotate the polyline
  plineObj.Rotate basePoint, rotationAngle
  plineObj.Update
End If
End If
End If
End If

  
  points2(1) = points2(3) + I:
  points2(3) = points2(3) + a + I:
  points2(5) = points2(3):
  points2(7) = points2(1):
  

Else: d = 1
d = d + 1

d = Cells(c, 4)
End If

Next d
c = c + 1

a = Cells(c, 2)
b = Cells(c, 3)

I = 100

Loop
              
ZoomAll
AP.Quit
End Sub

Sub MF143()
' Creating search objects
Dim AP As Excel.Application
Dim WB As Excel.Workbook
Dim WS As Excel.Worksheet

Set AP = Excel.Application
 Set WB = AP.Workbooks.Open("C:\Users\������\Desktop\������\2.xlsx")
Set WS = WB.Worksheets("������ �˨������ ������")
c = 19
a = Cells(c, 2)
b = Cells(c, 3)
d = Cells(c, 4)
E = 1
I = 0
On Error Resume Next

' Description of all the points that will make up rectangles
  Dim plineObj As AcadLWPolyline
  Dim plineObjw1 As AcadLWPolyline
  Dim plineObjw2 As AcadLWPolyline
  Dim points(0 To 7) As Double
  Dim pointsaktr1(0 To 503) As Double
  Dim pointsaktr2(0 To 39) As Double
  Dim pointsaktr3(0 To 9) As Double
  Dim pointsaktr4(0 To 13) As Double
  Dim pointsaktr5(0 To 5) As Double
  Dim pointsaktr6(0 To 9) As Double
  Dim pointsaktr7(0 To 7) As Double
  Dim pointscntr1(0 To 59) As Double
  Dim pointscntr2(0 To 55) As Double
  Dim pointscntr3(0 To 25) As Double
  Dim pointscntr4(0 To 39) As Double
  Dim pointscntr5(0 To 27) As Double
  Dim pointscntr6(0 To 11) As Double
  Dim pointswithin(0 To 15) As Double
  Dim pointswithin2(0 To 31) As Double
  Dim intPointsa
  Dim intPointsb
  Dim pointshelpa1(0 To 3) As Double
  Dim pointshelpa2(0 To 3) As Double
  Dim pointshelpb1(0 To 3) As Double
  Dim pointshelpb2(0 To 3) As Double
  Dim offsetObj As Variant
  Dim basePoint(0 To 2) As Double
  Dim rotationAngle As Double
  Dim b1(0 To 2) As Double
  Dim b2(0 To 2) As Double
  Dim a1(0 To 2) As Double
  Dim a2(0 To 2) As Double
  Dim circleObj As AcadCircle
  Dim center(0 To 2) As Double
  Dim radius As Double

points(6) = 0

Do While a > 0

' Calculate the coordinates of each point
  points(0) = points(6) + I:    points(1) = 0
  points(2) = points(6) + I:    points(3) = a
  points(4) = points(0) + b:    points(5) = a
  points(6) = points(0) + b:    points(7) = 0

' Constructing four segments and connecting the corresponding points

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
  plineObj.Closed = True

If a >= 140 Then
If b >= 140 Then

r = 30
  pointswithin(0) = points(0) + 70:                              pointswithin(1) = points(1) + 100
  pointswithin(2) = points(0) + 70:                              pointswithin(3) = points(3) - 100
  pointswithin(4) = points(0) + 100:                             pointswithin(5) = points(3) - 70
  pointswithin(6) = points(4) - 100:                             pointswithin(7) = points(3) - 70
  pointswithin(8) = points(4) - 70:                              pointswithin(9) = points(3) - 100
  pointswithin(10) = points(4) - 70:                             pointswithin(11) = points(1) + 100
  pointswithin(12) = points(4) - 100:                            pointswithin(13) = points(1) + 70
  pointswithin(14) = points(0) + 100:                            pointswithin(15) = points(1) + 70

  
If a > 241 Then
If b > 241 Then
  

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointswithin)
  plineObj.Closed = True
  ' Find the bulge of the third segment
    Dim currentBulge As Double
    currentBulge = plineObj.GetBulge(2)
    l = 2 * r * (Sqr(2) / 2)
    h = r * (1 - (Sqr(2) / 2))
    k = h / (l / 2)
    
    ' Change the bulge of the third segment
    plineObj.SetBulge 1, k
    plineObj.SetBulge 3, k
    plineObj.SetBulge 5, k
    plineObj.SetBulge 7, k
    plineObj.Layer = "C_2"
    plineObj.Update
    plineObj.Closed = True
    
  offsetObj = plineObj.Offset(6)
    plineObj.Layer = "C-Mill"
    plineObj.Update
    
End If
End If
  
  pointswithin2(0) = points(0) + 64:                              pointswithin2(1) = points(1) + 100
  pointswithin2(2) = points(0) + 64:                              pointswithin2(3) = points(3) - 100
  pointswithin2(4) = points(0) + 70:                              pointswithin2(5) = points(3) - 94
  pointswithin2(6) = points(0) + 94:                              pointswithin2(7) = points(3) - 70
  pointswithin2(8) = points(0) + 100:                             pointswithin2(9) = points(3) - 64
  pointswithin2(10) = points(4) - 100:                            pointswithin2(11) = points(3) - 64
  pointswithin2(12) = points(4) - 94:                             pointswithin2(13) = points(3) - 70
  pointswithin2(14) = points(4) - 70:                             pointswithin2(15) = points(3) - 94
  pointswithin2(16) = points(4) - 64:                             pointswithin2(17) = points(3) - 100
  pointswithin2(18) = points(4) - 64:                             pointswithin2(19) = points(1) + 100
  pointswithin2(20) = points(4) - 70:                             pointswithin2(21) = points(1) + 94
  pointswithin2(22) = points(4) - 94:                             pointswithin2(23) = points(1) + 70
  pointswithin2(24) = points(4) - 100:                            pointswithin2(25) = points(1) + 64
  pointswithin2(26) = points(0) + 100:                            pointswithin2(27) = points(1) + 64
  pointswithin2(28) = points(0) + 94:                             pointswithin2(29) = points(1) + 70
  pointswithin2(30) = points(0) + 70:                             pointswithin2(31) = points(1) + 94
  
If a > 241 Then
If b > 241 Then

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointswithin2)
  plineObj.Closed = True
    
    plineObj.SetBulge 1, -k
    plineObj.SetBulge 2, k
    plineObj.SetBulge 3, -k
    plineObj.SetBulge 5, -k
    plineObj.SetBulge 6, k
    plineObj.SetBulge 7, -k
    plineObj.SetBulge 9, -k
    plineObj.SetBulge 10, k
    plineObj.SetBulge 11, -k
    plineObj.SetBulge 13, -k
    plineObj.SetBulge 14, k
    plineObj.SetBulge 15, -k
  
   
    plineObj.Layer = "C_2"
    plineObj.Update
    plineObj.Closed = True
  
End If
End If


pointsaktr1(0) = points(0) + 60.5:      pointsaktr1(1) = points(1) + 33
pointsaktr1(2) = points(0) + 62.8818:   pointsaktr1(3) = points(1) + 37.6743
pointsaktr1(4) = points(0) + 59.9507:   pointsaktr1(5) = points(1) + 41.9916
pointsaktr1(6) = points(0) + 58.0126:   pointsaktr1(7) = points(1) + 41.8673
pointsaktr1(8) = points(0) + 56.7296:   pointsaktr1(9) = points(1) + 39.5607
pointsaktr1(10) = points(0) + 58.5907:  pointsaktr1(11) = points(1) + 37.5362
pointsaktr1(12) = points(0) + 59.6991:  pointsaktr1(13) = points(1) + 38.013
pointsaktr1(14) = points(0) + 60.022:   pointsaktr1(15) = points(1) + 40.2373
pointsaktr1(16) = points(0) + 61.2812:  pointsaktr1(17) = points(1) + 38.1637
pointsaktr1(18) = points(0) + 59.0042:  pointsaktr1(19) = points(1) + 34.633
pointsaktr1(20) = points(0) + 55.4873:  pointsaktr1(21) = points(1) + 33.7236
pointsaktr1(22) = points(0) + 49.9128:  pointsaktr1(23) = points(1) + 36.6409
pointsaktr1(24) = points(0) + 47.6162:  pointsaktr1(25) = points(1) + 44.2001
pointsaktr1(26) = points(0) + 51.7037:  pointsaktr1(27) = points(1) + 49.0761
pointsaktr1(28) = points(0) + 54.3447:  pointsaktr1(29) = points(1) + 48.7943
pointsaktr1(30) = points(0) + 55.7883:  pointsaktr1(31) = points(1) + 48.3252
pointsaktr1(32) = points(0) + 57.2677:  pointsaktr1(33) = points(1) + 48.0411
pointsaktr1(34) = points(0) + 59.2488:  pointsaktr1(35) = points(1) + 48.7514
pointsaktr1(36) = points(0) + 59.4835:  pointsaktr1(37) = points(1) + 50.9641
pointsaktr1(38) = points(0) + 57.5849:  pointsaktr1(39) = points(1) + 52.2794
pointsaktr1(40) = points(0) + 54.413:   pointsaktr1(41) = points(1) + 52.5139
pointsaktr1(42) = points(0) + 59.6861:  pointsaktr1(43) = points(1) + 53.7975
pointsaktr1(44) = points(0) + 61.8737:  pointsaktr1(45) = points(1) + 53.3042
pointsaktr1(46) = points(0) + 62.2884:  pointsaktr1(47) = points(1) + 53.366
pointsaktr1(48) = points(0) + 62.614:   pointsaktr1(49) = points(1) + 53.7765
pointsaktr1(50) = points(0) + 62.5956:  pointsaktr1(51) = points(1) + 54.4812
pointsaktr1(52) = points(0) + 62.0533:  pointsaktr1(53) = points(1) + 55.3326
pointsaktr1(54) = points(0) + 61.8627:  pointsaktr1(55) = points(1) + 55.9772
pointsaktr1(56) = points(0) + 62.0122:  pointsaktr1(57) = points(1) + 56.5534
pointsaktr1(58) = points(0) + 62.2419:  pointsaktr1(59) = points(1) + 57.4669
pointsaktr1(60) = points(0) + 61.7068:  pointsaktr1(61) = points(1) + 61.3268
pointsaktr1(62) = points(0) + 61.5862:  pointsaktr1(63) = points(1) + 61.5862
pointsaktr1(64) = points(0) + 61.3268:  pointsaktr1(65) = points(1) + 61.7068
pointsaktr1(66) = points(0) + 57.4669:  pointsaktr1(67) = points(1) + 62.2419
pointsaktr1(68) = points(0) + 56.5534:  pointsaktr1(69) = points(1) + 62.0122
pointsaktr1(70) = points(0) + 55.9772:  pointsaktr1(71) = points(1) + 61.8627
pointsaktr1(72) = points(0) + 55.3326:  pointsaktr1(73) = points(1) + 62.0533
pointsaktr1(74) = points(0) + 54.4812:  pointsaktr1(75) = points(1) + 62.5956
pointsaktr1(76) = points(0) + 53.7765:  pointsaktr1(77) = points(1) + 62.614
pointsaktr1(78) = points(0) + 53.366:   pointsaktr1(79) = points(1) + 62.2884
pointsaktr1(80) = points(0) + 53.3042:  pointsaktr1(81) = points(1) + 61.8737
pointsaktr1(82) = points(0) + 53.7975:  pointsaktr1(83) = points(1) + 59.6861
pointsaktr1(84) = points(0) + 52.5139:  pointsaktr1(85) = points(1) + 54.413
pointsaktr1(86) = points(0) + 52.2794:  pointsaktr1(87) = points(1) + 57.5849
pointsaktr1(88) = points(0) + 50.9641:  pointsaktr1(89) = points(1) + 59.4835
pointsaktr1(90) = points(0) + 48.7514:  pointsaktr1(91) = points(1) + 59.2488
pointsaktr1(92) = points(0) + 48.0411:  pointsaktr1(93) = points(1) + 57.2677
pointsaktr1(94) = points(0) + 48.3252:  pointsaktr1(95) = points(1) + 55.7883
pointsaktr1(96) = points(0) + 48.7943:  pointsaktr1(97) = points(1) + 54.3447
pointsaktr1(98) = points(0) + 49.0761:  pointsaktr1(99) = points(1) + 51.7037
pointsaktr1(100) = points(0) + 44.2001: pointsaktr1(101) = points(1) + 47.6162
pointsaktr1(102) = points(0) + 36.6409: pointsaktr1(103) = points(1) + 49.9128
pointsaktr1(104) = points(0) + 33.7236: pointsaktr1(105) = points(1) + 55.4873
pointsaktr1(106) = points(0) + 34.633:  pointsaktr1(107) = points(1) + 59.0042
pointsaktr1(108) = points(0) + 38.1637: pointsaktr1(109) = points(1) + 61.2812
pointsaktr1(110) = points(0) + 40.2373: pointsaktr1(111) = points(1) + 60.022
pointsaktr1(112) = points(0) + 38.013:  pointsaktr1(113) = points(1) + 59.6991
pointsaktr1(114) = points(0) + 37.5362: pointsaktr1(115) = points(1) + 58.5907
pointsaktr1(116) = points(0) + 39.5607: pointsaktr1(117) = points(1) + 56.7296
pointsaktr1(118) = points(0) + 41.8673: pointsaktr1(119) = points(1) + 58.0126
pointsaktr1(120) = points(0) + 41.9916: pointsaktr1(121) = points(1) + 59.9507
pointsaktr1(122) = points(0) + 37.6743: pointsaktr1(123) = points(1) + 62.8818
pointsaktr1(124) = points(0) + 33:      pointsaktr1(125) = points(1) + 60.5
pointsaktr1(126) = points(0) + 33:      pointsaktr1(127) = points(3) - 60.5
pointsaktr1(128) = points(0) + 37.6743: pointsaktr1(129) = points(3) - 62.8818
pointsaktr1(130) = points(0) + 41.9916: pointsaktr1(131) = points(3) - 59.9507
pointsaktr1(132) = points(0) + 41.8673: pointsaktr1(133) = points(3) - 58.0126
pointsaktr1(134) = points(0) + 39.5607: pointsaktr1(135) = points(3) - 56.7296
pointsaktr1(136) = points(0) + 37.5362: pointsaktr1(137) = points(3) - 58.5907
pointsaktr1(138) = points(0) + 38.013:  pointsaktr1(139) = points(3) - 59.6991
pointsaktr1(140) = points(0) + 40.2373: pointsaktr1(141) = points(3) - 60.022
pointsaktr1(142) = points(0) + 38.1637: pointsaktr1(143) = points(3) - 61.2812
pointsaktr1(144) = points(0) + 34.633:  pointsaktr1(145) = points(3) - 59.0042
pointsaktr1(146) = points(0) + 33.7236: pointsaktr1(147) = points(3) - 55.4873
pointsaktr1(148) = points(0) + 36.6409: pointsaktr1(149) = points(3) - 49.9128
pointsaktr1(150) = points(0) + 44.2001: pointsaktr1(151) = points(3) - 47.6162
pointsaktr1(152) = points(0) + 49.0761: pointsaktr1(153) = points(3) - 51.7037
pointsaktr1(154) = points(0) + 48.7943: pointsaktr1(155) = points(3) - 54.3447
pointsaktr1(156) = points(0) + 48.3252: pointsaktr1(157) = points(3) - 55.7883
pointsaktr1(158) = points(0) + 48.0411: pointsaktr1(159) = points(3) - 57.2677
pointsaktr1(160) = points(0) + 48.7514: pointsaktr1(161) = points(3) - 59.2488
pointsaktr1(162) = points(0) + 50.9641: pointsaktr1(163) = points(3) - 59.4835
pointsaktr1(164) = points(0) + 52.2794: pointsaktr1(165) = points(3) - 57.5849
pointsaktr1(166) = points(0) + 52.5139: pointsaktr1(167) = points(3) - 54.413
pointsaktr1(168) = points(0) + 53.7975: pointsaktr1(169) = points(3) - 59.6861
pointsaktr1(170) = points(0) + 53.3042: pointsaktr1(171) = points(3) - 61.8737
pointsaktr1(172) = points(0) + 53.366:  pointsaktr1(173) = points(3) - 62.2884
pointsaktr1(174) = points(0) + 53.7765: pointsaktr1(175) = points(3) - 62.614
pointsaktr1(176) = points(0) + 54.4812: pointsaktr1(177) = points(3) - 62.5956
pointsaktr1(178) = points(0) + 55.3326: pointsaktr1(179) = points(3) - 62.0533
pointsaktr1(180) = points(0) + 55.9772: pointsaktr1(181) = points(3) - 61.8627
pointsaktr1(182) = points(0) + 56.5534: pointsaktr1(183) = points(3) - 62.0122
pointsaktr1(184) = points(0) + 57.4669: pointsaktr1(185) = points(3) - 62.2419
pointsaktr1(186) = points(0) + 61.3268: pointsaktr1(187) = points(3) - 61.7068
pointsaktr1(188) = points(0) + 61.5862: pointsaktr1(189) = points(3) - 61.5862
pointsaktr1(190) = points(0) + 61.7068: pointsaktr1(191) = points(3) - 61.3268
pointsaktr1(192) = points(0) + 62.2419: pointsaktr1(193) = points(3) - 57.4669
pointsaktr1(194) = points(0) + 62.0122: pointsaktr1(195) = points(3) - 56.5534
pointsaktr1(196) = points(0) + 61.8627: pointsaktr1(197) = points(3) - 55.9772
pointsaktr1(198) = points(0) + 62.0533: pointsaktr1(199) = points(3) - 55.3326
pointsaktr1(200) = points(0) + 62.5956: pointsaktr1(201) = points(3) - 54.4812
pointsaktr1(202) = points(0) + 62.614:  pointsaktr1(203) = points(3) - 53.7765
pointsaktr1(204) = points(0) + 62.2884: pointsaktr1(205) = points(3) - 53.366
pointsaktr1(206) = points(0) + 61.8737: pointsaktr1(207) = points(3) - 53.3042
pointsaktr1(208) = points(0) + 59.6861: pointsaktr1(209) = points(3) - 53.7975
pointsaktr1(210) = points(0) + 54.413:  pointsaktr1(211) = points(3) - 52.5139
pointsaktr1(212) = points(0) + 57.5849: pointsaktr1(213) = points(3) - 52.2794
pointsaktr1(214) = points(0) + 59.4835: pointsaktr1(215) = points(3) - 50.9641
pointsaktr1(216) = points(0) + 59.2488: pointsaktr1(217) = points(3) - 48.7514
pointsaktr1(218) = points(0) + 57.2677: pointsaktr1(219) = points(3) - 48.0411
pointsaktr1(220) = points(0) + 55.7883: pointsaktr1(221) = points(3) - 48.3252
pointsaktr1(222) = points(0) + 54.3447: pointsaktr1(223) = points(3) - 48.7943
pointsaktr1(224) = points(0) + 51.7037: pointsaktr1(225) = points(3) - 49.0761
pointsaktr1(226) = points(0) + 47.6162: pointsaktr1(227) = points(3) - 44.2001
pointsaktr1(228) = points(0) + 49.9128: pointsaktr1(229) = points(3) - 36.6409
pointsaktr1(230) = points(0) + 55.4873: pointsaktr1(231) = points(3) - 33.7236
pointsaktr1(232) = points(0) + 59.0042: pointsaktr1(233) = points(3) - 34.633
pointsaktr1(234) = points(0) + 61.2812: pointsaktr1(235) = points(3) - 38.1637
pointsaktr1(236) = points(0) + 60.022:  pointsaktr1(237) = points(3) - 40.2373
pointsaktr1(238) = points(0) + 59.6991: pointsaktr1(239) = points(3) - 38.013
pointsaktr1(240) = points(0) + 58.5907: pointsaktr1(241) = points(3) - 37.5362
pointsaktr1(242) = points(0) + 56.7296: pointsaktr1(243) = points(3) - 39.5607
pointsaktr1(244) = points(0) + 58.0126: pointsaktr1(245) = points(3) - 41.8673
pointsaktr1(246) = points(0) + 59.9507: pointsaktr1(247) = points(3) - 41.9916
pointsaktr1(248) = points(0) + 62.8818: pointsaktr1(249) = points(3) - 37.6743
pointsaktr1(250) = points(0) + 60.5:    pointsaktr1(251) = points(3) - 33
pointsaktr1(252) = points(4) - 60.5:    pointsaktr1(253) = points(3) - 33
pointsaktr1(254) = points(4) - 62.8818: pointsaktr1(255) = points(3) - 37.6743
pointsaktr1(256) = points(4) - 59.9507: pointsaktr1(257) = points(3) - 41.9916
pointsaktr1(258) = points(4) - 58.0126: pointsaktr1(259) = points(3) - 41.8673
pointsaktr1(260) = points(4) - 56.7296: pointsaktr1(261) = points(3) - 39.5607
pointsaktr1(262) = points(4) - 58.5907: pointsaktr1(263) = points(3) - 37.5362
pointsaktr1(264) = points(4) - 59.6991: pointsaktr1(265) = points(3) - 38.013
pointsaktr1(266) = points(4) - 60.022:  pointsaktr1(267) = points(3) - 40.2373
pointsaktr1(268) = points(4) - 61.2812: pointsaktr1(269) = points(3) - 38.1637
pointsaktr1(270) = points(4) - 59.0042: pointsaktr1(271) = points(3) - 34.633
pointsaktr1(272) = points(4) - 55.4873: pointsaktr1(273) = points(3) - 33.7236
pointsaktr1(274) = points(4) - 49.9128: pointsaktr1(275) = points(3) - 36.6409
pointsaktr1(276) = points(4) - 47.6162: pointsaktr1(277) = points(3) - 44.2001
pointsaktr1(278) = points(4) - 51.7037: pointsaktr1(279) = points(3) - 49.0761
pointsaktr1(280) = points(4) - 54.3447: pointsaktr1(281) = points(3) - 48.7943
pointsaktr1(282) = points(4) - 55.7883: pointsaktr1(283) = points(3) - 48.3252
pointsaktr1(284) = points(4) - 57.2677: pointsaktr1(285) = points(3) - 48.0411
pointsaktr1(286) = points(4) - 59.2488: pointsaktr1(287) = points(3) - 48.7514
pointsaktr1(288) = points(4) - 59.4835: pointsaktr1(289) = points(3) - 50.9641
pointsaktr1(290) = points(4) - 57.5849: pointsaktr1(291) = points(3) - 52.2794
pointsaktr1(292) = points(4) - 54.413:  pointsaktr1(293) = points(3) - 52.5139
pointsaktr1(294) = points(4) - 59.6861: pointsaktr1(295) = points(3) - 53.7975
pointsaktr1(296) = points(4) - 61.8737: pointsaktr1(297) = points(3) - 53.3042
pointsaktr1(298) = points(4) - 62.2884: pointsaktr1(299) = points(3) - 53.366
pointsaktr1(300) = points(4) - 62.614:  pointsaktr1(301) = points(3) - 53.7765
pointsaktr1(302) = points(4) - 62.5956: pointsaktr1(303) = points(3) - 54.4812
pointsaktr1(304) = points(4) - 62.0533: pointsaktr1(305) = points(3) - 55.3326
pointsaktr1(306) = points(4) - 61.8627: pointsaktr1(307) = points(3) - 55.9772
pointsaktr1(308) = points(4) - 62.0122: pointsaktr1(309) = points(3) - 56.5534
pointsaktr1(310) = points(4) - 62.2419: pointsaktr1(311) = points(3) - 57.4669
pointsaktr1(312) = points(4) - 61.7068: pointsaktr1(313) = points(3) - 61.3268
pointsaktr1(314) = points(4) - 61.5862: pointsaktr1(315) = points(3) - 61.5862
pointsaktr1(316) = points(4) - 61.3268: pointsaktr1(317) = points(3) - 61.7068
pointsaktr1(318) = points(4) - 57.4669: pointsaktr1(319) = points(3) - 62.2419
pointsaktr1(320) = points(4) - 56.5534: pointsaktr1(321) = points(3) - 62.0122
pointsaktr1(322) = points(4) - 55.9772: pointsaktr1(323) = points(3) - 61.8627
pointsaktr1(324) = points(4) - 55.3326: pointsaktr1(325) = points(3) - 62.0533
pointsaktr1(326) = points(4) - 54.4812: pointsaktr1(327) = points(3) - 62.5956
pointsaktr1(328) = points(4) - 53.7765: pointsaktr1(329) = points(3) - 62.614
pointsaktr1(330) = points(4) - 53.366:  pointsaktr1(331) = points(3) - 62.2884
pointsaktr1(332) = points(4) - 53.3042: pointsaktr1(333) = points(3) - 61.8737
pointsaktr1(334) = points(4) - 53.7975: pointsaktr1(335) = points(3) - 59.6861
pointsaktr1(336) = points(4) - 52.5139: pointsaktr1(337) = points(3) - 54.413
pointsaktr1(338) = points(4) - 52.2794: pointsaktr1(339) = points(3) - 57.5849
pointsaktr1(340) = points(4) - 50.9641: pointsaktr1(341) = points(3) - 59.4835
pointsaktr1(342) = points(4) - 48.7514: pointsaktr1(343) = points(3) - 59.2488
pointsaktr1(344) = points(4) - 48.0411: pointsaktr1(345) = points(3) - 57.2677
pointsaktr1(346) = points(4) - 48.3252: pointsaktr1(347) = points(3) - 55.7883
pointsaktr1(348) = points(4) - 48.7943: pointsaktr1(349) = points(3) - 54.3447
pointsaktr1(350) = points(4) - 49.0761: pointsaktr1(351) = points(3) - 51.7037
pointsaktr1(352) = points(4) - 44.2001: pointsaktr1(353) = points(3) - 47.6162
pointsaktr1(354) = points(4) - 36.6409: pointsaktr1(355) = points(3) - 49.9128
pointsaktr1(356) = points(4) - 33.7236: pointsaktr1(357) = points(3) - 55.4873
pointsaktr1(358) = points(4) - 34.633:  pointsaktr1(359) = points(3) - 59.0042
pointsaktr1(360) = points(4) - 38.1637: pointsaktr1(361) = points(3) - 61.2812
pointsaktr1(362) = points(4) - 40.2373: pointsaktr1(363) = points(3) - 60.022
pointsaktr1(364) = points(4) - 38.013:  pointsaktr1(365) = points(3) - 59.6991
pointsaktr1(366) = points(4) - 37.5362: pointsaktr1(367) = points(3) - 58.5907
pointsaktr1(368) = points(4) - 39.5607: pointsaktr1(369) = points(3) - 56.7296
pointsaktr1(370) = points(4) - 41.8673: pointsaktr1(371) = points(3) - 58.0126
pointsaktr1(372) = points(4) - 41.9916: pointsaktr1(373) = points(3) - 59.9507
pointsaktr1(374) = points(4) - 37.6743: pointsaktr1(375) = points(3) - 62.8818
pointsaktr1(376) = points(4) - 33:      pointsaktr1(377) = points(3) - 60.5
pointsaktr1(378) = points(4) - 33:      pointsaktr1(379) = points(1) + 60.5
pointsaktr1(380) = points(4) - 37.6743: pointsaktr1(381) = points(1) + 62.8818
pointsaktr1(382) = points(4) - 41.9916: pointsaktr1(383) = points(1) + 59.9507
pointsaktr1(384) = points(4) - 41.8673: pointsaktr1(385) = points(1) + 58.0126
pointsaktr1(386) = points(4) - 39.5607: pointsaktr1(387) = points(1) + 56.7296
pointsaktr1(388) = points(4) - 37.5362: pointsaktr1(389) = points(1) + 58.5907
pointsaktr1(390) = points(4) - 38.013:  pointsaktr1(391) = points(1) + 59.6991
pointsaktr1(392) = points(4) - 40.2373: pointsaktr1(393) = points(1) + 60.022
pointsaktr1(394) = points(4) - 38.1637: pointsaktr1(395) = points(1) + 61.2812
pointsaktr1(396) = points(4) - 34.633:  pointsaktr1(397) = points(1) + 59.0042
pointsaktr1(398) = points(4) - 33.7236: pointsaktr1(399) = points(1) + 55.4873
pointsaktr1(400) = points(4) - 36.6409: pointsaktr1(401) = points(1) + 49.9128
pointsaktr1(402) = points(4) - 44.2001: pointsaktr1(403) = points(1) + 47.6162
pointsaktr1(404) = points(4) - 49.0761: pointsaktr1(405) = points(1) + 51.7037
pointsaktr1(406) = points(4) - 48.7943: pointsaktr1(407) = points(1) + 54.3447
pointsaktr1(408) = points(4) - 48.3252: pointsaktr1(409) = points(1) + 55.7883
pointsaktr1(410) = points(4) - 48.0411: pointsaktr1(411) = points(1) + 57.2677
pointsaktr1(412) = points(4) - 48.7514: pointsaktr1(413) = points(1) + 59.2488
pointsaktr1(414) = points(4) - 50.9641: pointsaktr1(415) = points(1) + 59.4835
pointsaktr1(416) = points(4) - 52.2794: pointsaktr1(417) = points(1) + 57.5849
pointsaktr1(418) = points(4) - 52.5139: pointsaktr1(419) = points(1) + 54.413
pointsaktr1(420) = points(4) - 53.7975: pointsaktr1(421) = points(1) + 59.6861
pointsaktr1(422) = points(4) - 53.3042: pointsaktr1(423) = points(1) + 61.8737
pointsaktr1(424) = points(4) - 53.366:  pointsaktr1(425) = points(1) + 62.2884
pointsaktr1(426) = points(4) - 53.7765: pointsaktr1(427) = points(1) + 62.614
pointsaktr1(428) = points(4) - 54.4812: pointsaktr1(429) = points(1) + 62.5956
pointsaktr1(430) = points(4) - 55.3326: pointsaktr1(431) = points(1) + 62.0533
pointsaktr1(432) = points(4) - 55.9772: pointsaktr1(433) = points(1) + 61.8627
pointsaktr1(434) = points(4) - 56.5534: pointsaktr1(435) = points(1) + 62.0122
pointsaktr1(436) = points(4) - 57.4669: pointsaktr1(437) = points(1) + 62.2419
pointsaktr1(438) = points(4) - 61.3268: pointsaktr1(439) = points(1) + 61.7068
pointsaktr1(440) = points(4) - 61.5862: pointsaktr1(441) = points(1) + 61.5862
pointsaktr1(442) = points(4) - 61.7068: pointsaktr1(443) = points(1) + 61.3268
pointsaktr1(444) = points(4) - 62.2419: pointsaktr1(445) = points(1) + 57.4669
pointsaktr1(446) = points(4) - 62.0122: pointsaktr1(447) = points(1) + 56.5534
pointsaktr1(448) = points(4) - 61.8627: pointsaktr1(449) = points(1) + 55.9772
pointsaktr1(450) = points(4) - 62.0533: pointsaktr1(451) = points(1) + 55.3326
pointsaktr1(452) = points(4) - 62.5956: pointsaktr1(453) = points(1) + 54.4812
pointsaktr1(454) = points(4) - 62.614:  pointsaktr1(455) = points(1) + 53.7765
pointsaktr1(456) = points(4) - 62.2884: pointsaktr1(457) = points(1) + 53.366
pointsaktr1(458) = points(4) - 61.8737: pointsaktr1(459) = points(1) + 53.3042
pointsaktr1(460) = points(4) - 59.6861: pointsaktr1(461) = points(1) + 53.7975
pointsaktr1(462) = points(4) - 54.413:  pointsaktr1(463) = points(1) + 52.5139
pointsaktr1(464) = points(4) - 57.5849: pointsaktr1(465) = points(1) + 52.2794
pointsaktr1(466) = points(4) - 59.4835: pointsaktr1(467) = points(1) + 50.9641
pointsaktr1(468) = points(4) - 59.2488: pointsaktr1(469) = points(1) + 48.7514
pointsaktr1(470) = points(4) - 57.2677: pointsaktr1(471) = points(1) + 48.0411
pointsaktr1(472) = points(4) - 55.7883: pointsaktr1(473) = points(1) + 48.3252
pointsaktr1(474) = points(4) - 54.3447: pointsaktr1(475) = points(1) + 48.7943
pointsaktr1(476) = points(4) - 51.7037: pointsaktr1(477) = points(1) + 49.0761
pointsaktr1(478) = points(4) - 47.6162: pointsaktr1(479) = points(1) + 44.2001
pointsaktr1(480) = points(4) - 49.9128: pointsaktr1(481) = points(1) + 36.6409
pointsaktr1(482) = points(4) - 55.4873: pointsaktr1(483) = points(1) + 33.7236
pointsaktr1(484) = points(4) - 59.0042: pointsaktr1(485) = points(1) + 34.633
pointsaktr1(486) = points(4) - 61.2812: pointsaktr1(487) = points(1) + 38.1637
pointsaktr1(488) = points(4) - 60.022:  pointsaktr1(489) = points(1) + 40.2373
pointsaktr1(490) = points(4) - 59.6991: pointsaktr1(491) = points(1) + 38.013
pointsaktr1(492) = points(4) - 58.5907: pointsaktr1(493) = points(1) + 37.5362
pointsaktr1(494) = points(4) - 56.7296: pointsaktr1(495) = points(1) + 39.5607
pointsaktr1(496) = points(4) - 58.0126: pointsaktr1(497) = points(1) + 41.8673
pointsaktr1(498) = points(4) - 59.9507: pointsaktr1(499) = points(1) + 41.9916
pointsaktr1(500) = points(4) - 62.8818: pointsaktr1(501) = points(1) + 37.6743
pointsaktr1(502) = points(4) - 60.5:    pointsaktr1(503) = points(1) + 33

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsaktr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update
    
currentBulge = plineObj.GetBulge(2)

plineObj.SetBulge 0, 0.240087
plineObj.SetBulge 1, 0.307386
plineObj.SetBulge 2, 0.224633
plineObj.SetBulge 3, 0.286077
plineObj.SetBulge 4, 0.36144
plineObj.SetBulge 5, 0.276755
plineObj.SetBulge 6, 0.244948
plineObj.SetBulge 7, -0.314822
plineObj.SetBulge 8, -0.259895
plineObj.SetBulge 9, -0.118763
plineObj.SetBulge 10, -0.254683
plineObj.SetBulge 11, -0.148555
plineObj.SetBulge 12, -0.363708
plineObj.SetBulge 13, -0.141808
plineObj.SetBulge 14, 0.0369261
plineObj.SetBulge 15, 0.0253271
plineObj.SetBulge 16, 0.246491
plineObj.SetBulge 17, 0.33
plineObj.SetBulge 18, 0.220023
plineObj.SetBulge 19, 0.0495029
plineObj.SetBulge 20, -0.130373
plineObj.SetBulge 21, -0.100964
plineObj.SetBulge 22, 0.29339
plineObj.SetBulge 23, 0.0911692
plineObj.SetBulge 24, 0.263146
plineObj.SetBulge 25, 0.0132959
plineObj.SetBulge 26, -0.154396
plineObj.SetBulge 27, -0.117989
plineObj.SetBulge 28, 0.121785
plineObj.SetBulge 29, 0.0709732
plineObj.SetBulge 30, 0.0779699
plineObj.SetBulge 31, 0.0779699
plineObj.SetBulge 32, 0.0709732
plineObj.SetBulge 33, 0.121785
plineObj.SetBulge 34, -0.117989
plineObj.SetBulge 35, -0.154396
plineObj.SetBulge 36, 0.0132959
plineObj.SetBulge 37, 0.263146
plineObj.SetBulge 38, 0.0911692
plineObj.SetBulge 39, 0.29339
plineObj.SetBulge 40, -0.100964
plineObj.SetBulge 41, -0.130373
plineObj.SetBulge 42, 0.0495029
plineObj.SetBulge 43, 0.220023
plineObj.SetBulge 44, 0.33
plineObj.SetBulge 45, 0.246491
plineObj.SetBulge 46, 0.0253271
plineObj.SetBulge 47, 0.0369261
plineObj.SetBulge 48, -0.141808
plineObj.SetBulge 49, -0.363708
plineObj.SetBulge 50, -0.148555
plineObj.SetBulge 51, -0.254683
plineObj.SetBulge 52, -0.118763
plineObj.SetBulge 53, -0.259895
plineObj.SetBulge 54, -0.314822
plineObj.SetBulge 55, 0.244948
plineObj.SetBulge 56, 0.276755
plineObj.SetBulge 57, 0.36144
plineObj.SetBulge 58, 0.286077
plineObj.SetBulge 59, 0.224633
plineObj.SetBulge 60, 0.307386
plineObj.SetBulge 61, 0.240087
plineObj.SetBulge 62, 0
plineObj.SetBulge 63, 0.240087
plineObj.SetBulge 64, 0.307386
plineObj.SetBulge 65, 0.224633
plineObj.SetBulge 66, 0.286077
plineObj.SetBulge 67, 0.36144
plineObj.SetBulge 68, 0.276755
plineObj.SetBulge 69, 0.244948
plineObj.SetBulge 70, -0.314822
plineObj.SetBulge 71, -0.259895
plineObj.SetBulge 72, -0.118763
plineObj.SetBulge 73, -0.254683
plineObj.SetBulge 74, -0.148555
plineObj.SetBulge 75, -0.363708
plineObj.SetBulge 76, -0.141808
plineObj.SetBulge 77, 0.0369261
plineObj.SetBulge 78, 0.0253271
plineObj.SetBulge 79, 0.246491
plineObj.SetBulge 80, 0.33
plineObj.SetBulge 81, 0.220023
plineObj.SetBulge 82, 0.0495029
plineObj.SetBulge 83, -0.130373
plineObj.SetBulge 84, -0.100964
plineObj.SetBulge 85, 0.29339
plineObj.SetBulge 86, 0.0911692
plineObj.SetBulge 87, 0.263146
plineObj.SetBulge 88, 0.0132959
plineObj.SetBulge 89, -0.154396
plineObj.SetBulge 90, -0.117989
plineObj.SetBulge 91, 0.121785
plineObj.SetBulge 92, 0.0709732
plineObj.SetBulge 93, 0.0779699
plineObj.SetBulge 94, 0.0779699
plineObj.SetBulge 95, 0.0709732
plineObj.SetBulge 96, 0.121785
plineObj.SetBulge 97, -0.117989
plineObj.SetBulge 98, -0.154396
plineObj.SetBulge 99, 0.0132959
plineObj.SetBulge 100, 0.263146
plineObj.SetBulge 101, 0.0911692
plineObj.SetBulge 102, 0.29339
plineObj.SetBulge 103, -0.100964
plineObj.SetBulge 104, -0.130373
plineObj.SetBulge 105, 0.0495029
plineObj.SetBulge 106, 0.220023
plineObj.SetBulge 107, 0.33
plineObj.SetBulge 108, 0.246491
plineObj.SetBulge 109, 0.0253271
plineObj.SetBulge 110, 0.0369261
plineObj.SetBulge 111, -0.141808
plineObj.SetBulge 112, -0.363708
plineObj.SetBulge 113, -0.148555
plineObj.SetBulge 114, -0.254683
plineObj.SetBulge 115, -0.118763
plineObj.SetBulge 116, -0.259895
plineObj.SetBulge 117, -0.314822
plineObj.SetBulge 118, 0.244948
plineObj.SetBulge 119, 0.276755
plineObj.SetBulge 120, 0.36144
plineObj.SetBulge 121, 0.286077
plineObj.SetBulge 122, 0.224633
plineObj.SetBulge 123, 0.307386
plineObj.SetBulge 124, 0.240087
plineObj.SetBulge 125, 0
plineObj.SetBulge 126, 0.240087
plineObj.SetBulge 127, 0.307386
plineObj.SetBulge 128, 0.224633
plineObj.SetBulge 129, 0.286077
plineObj.SetBulge 130, 0.36144
plineObj.SetBulge 131, 0.276755
plineObj.SetBulge 132, 0.244948
plineObj.SetBulge 133, -0.314822
plineObj.SetBulge 134, -0.259895
plineObj.SetBulge 135, -0.118763
plineObj.SetBulge 136, -0.254683
plineObj.SetBulge 137, -0.148555
plineObj.SetBulge 138, -0.363708
plineObj.SetBulge 139, -0.141808
plineObj.SetBulge 140, 0.0369261
plineObj.SetBulge 141, 0.0253271
plineObj.SetBulge 142, 0.246491
plineObj.SetBulge 143, 0.33
plineObj.SetBulge 144, 0.220023
plineObj.SetBulge 145, 0.0495029
plineObj.SetBulge 146, -0.130373
plineObj.SetBulge 147, -0.100964
plineObj.SetBulge 148, 0.29339
plineObj.SetBulge 149, 0.0911692
plineObj.SetBulge 150, 0.263146
plineObj.SetBulge 151, 0.0132959
plineObj.SetBulge 152, -0.154396
plineObj.SetBulge 153, -0.117989
plineObj.SetBulge 154, 0.121785
plineObj.SetBulge 155, 0.0709732
plineObj.SetBulge 156, 0.0779699
plineObj.SetBulge 157, 0.0779699
plineObj.SetBulge 158, 0.0709732
plineObj.SetBulge 159, 0.121785
plineObj.SetBulge 160, -0.117989
plineObj.SetBulge 161, -0.154396
plineObj.SetBulge 162, 0.0132959
plineObj.SetBulge 163, 0.263146
plineObj.SetBulge 164, 0.0911692
plineObj.SetBulge 165, 0.29339
plineObj.SetBulge 166, -0.100964
plineObj.SetBulge 167, -0.130373
plineObj.SetBulge 168, 0.0495029
plineObj.SetBulge 169, 0.220023
plineObj.SetBulge 170, 0.33
plineObj.SetBulge 171, 0.246491
plineObj.SetBulge 172, 0.0253271
plineObj.SetBulge 173, 0.0369261
plineObj.SetBulge 174, -0.141808
plineObj.SetBulge 175, -0.363708
plineObj.SetBulge 176, -0.148555
plineObj.SetBulge 177, -0.254683
plineObj.SetBulge 178, -0.118763
plineObj.SetBulge 179, -0.259895
plineObj.SetBulge 180, -0.314822
plineObj.SetBulge 181, 0.244948
plineObj.SetBulge 182, 0.276755
plineObj.SetBulge 183, 0.36144
plineObj.SetBulge 184, 0.286077
plineObj.SetBulge 185, 0.224633
plineObj.SetBulge 186, 0.307386
plineObj.SetBulge 187, 0.240087
plineObj.SetBulge 188, 0
plineObj.SetBulge 189, 0.240087
plineObj.SetBulge 190, 0.307386
plineObj.SetBulge 191, 0.224633
plineObj.SetBulge 192, 0.286077
plineObj.SetBulge 193, 0.36144
plineObj.SetBulge 194, 0.276755
plineObj.SetBulge 195, 0.244948
plineObj.SetBulge 196, -0.314822
plineObj.SetBulge 197, -0.259895
plineObj.SetBulge 198, -0.118763
plineObj.SetBulge 199, -0.254683
plineObj.SetBulge 200, -0.148555
plineObj.SetBulge 201, -0.363708
plineObj.SetBulge 202, -0.141808
plineObj.SetBulge 203, 0.0369261
plineObj.SetBulge 204, 0.0253271
plineObj.SetBulge 205, 0.246491
plineObj.SetBulge 206, 0.33
plineObj.SetBulge 207, 0.220023
plineObj.SetBulge 208, 0.0495029
plineObj.SetBulge 209, -0.130373
plineObj.SetBulge 210, -0.100964
plineObj.SetBulge 211, 0.29339
plineObj.SetBulge 212, 0.0911692
plineObj.SetBulge 213, 0.263146
plineObj.SetBulge 214, 0.0132959
plineObj.SetBulge 215, -0.154396
plineObj.SetBulge 216, -0.117989
plineObj.SetBulge 217, 0.121785
plineObj.SetBulge 218, 0.0709732
plineObj.SetBulge 219, 0.0779699
plineObj.SetBulge 220, 0.0779699
plineObj.SetBulge 221, 0.0709732
plineObj.SetBulge 222, 0.121785
plineObj.SetBulge 223, -0.117989
plineObj.SetBulge 224, -0.154396
plineObj.SetBulge 225, 0.0132959
plineObj.SetBulge 226, 0.263146
plineObj.SetBulge 227, 0.0911692
plineObj.SetBulge 228, 0.29339
plineObj.SetBulge 229, -0.100964
plineObj.SetBulge 230, -0.130373
plineObj.SetBulge 231, 0.0495029
plineObj.SetBulge 232, 0.220023
plineObj.SetBulge 233, 0.33
plineObj.SetBulge 234, 0.246491
plineObj.SetBulge 235, 0.0253271
plineObj.SetBulge 236, 0.0369261
plineObj.SetBulge 237, -0.141808
plineObj.SetBulge 238, -0.363708
plineObj.SetBulge 239, -0.148555
plineObj.SetBulge 240, -0.254683
plineObj.SetBulge 241, -0.118763
plineObj.SetBulge 242, -0.259895
plineObj.SetBulge 243, -0.314822
plineObj.SetBulge 244, 0.244948
plineObj.SetBulge 245, 0.276755
plineObj.SetBulge 246, 0.36144
plineObj.SetBulge 247, 0.286077
plineObj.SetBulge 248, 0.224633
plineObj.SetBulge 249, 0.307386
plineObj.SetBulge 250, 0.240087
plineObj.SetBulge 251, 0
 
    plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True
  
pointsaktr2(0) = points(0) + 56.1843:   pointsaktr2(1) = points(1) + 30
pointsaktr2(2) = points(0) + 45.3574:   pointsaktr2(3) = points(1) + 41.645
pointsaktr2(4) = points(0) + 46.1831:   pointsaktr2(5) = points(1) + 46.1831
pointsaktr2(6) = points(0) + 41.645: pointsaktr2(7) = points(1) + 45.3574
pointsaktr2(8) = points(0) + 30: pointsaktr2(9) = points(1) + 56.1843
pointsaktr2(10) = points(0) + 30:   pointsaktr2(11) = points(3) - 56.1843
pointsaktr2(12) = points(0) + 41.645:   pointsaktr2(13) = points(3) - 45.3574
pointsaktr2(14) = points(0) + 46.1831:  pointsaktr2(15) = points(3) - 46.1831
pointsaktr2(16) = points(0) + 45.3574:  pointsaktr2(17) = points(3) - 41.645
pointsaktr2(18) = points(0) + 56.1843:  pointsaktr2(19) = points(3) - 30
pointsaktr2(20) = points(4) - 56.1843:  pointsaktr2(21) = points(3) - 30
pointsaktr2(22) = points(4) - 45.3574:  pointsaktr2(23) = points(3) - 41.645
pointsaktr2(24) = points(4) - 46.1831:  pointsaktr2(25) = points(3) - 46.1831
pointsaktr2(26) = points(4) - 41.645:   pointsaktr2(27) = points(3) - 45.3574
pointsaktr2(28) = points(4) - 30:   pointsaktr2(29) = points(3) - 56.1843
pointsaktr2(30) = points(4) - 30:   pointsaktr2(31) = points(1) + 56.1843
pointsaktr2(32) = points(4) - 41.645:   pointsaktr2(33) = points(1) + 45.3574
pointsaktr2(34) = points(4) - 46.1831:  pointsaktr2(35) = points(1) + 46.1831
pointsaktr2(36) = points(4) - 45.3574:  pointsaktr2(37) = points(1) + 41.645
pointsaktr2(38) = points(4) - 56.1843:  pointsaktr2(39) = points(1) + 30


Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsaktr2)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update
  
plineObj.SetBulge 0, -0.435695
plineObj.SetBulge 1, -0.0902271
plineObj.SetBulge 2, -0.0902271
plineObj.SetBulge 3, -0.435695
plineObj.SetBulge 4, 0
plineObj.SetBulge 5, -0.435695
plineObj.SetBulge 6, -0.0902271
plineObj.SetBulge 7, -0.0902271
plineObj.SetBulge 8, -0.435695
plineObj.SetBulge 9, 0
plineObj.SetBulge 10, -0.435695
plineObj.SetBulge 11, -0.0902271
plineObj.SetBulge 12, -0.0902271
plineObj.SetBulge 13, -0.435695
plineObj.SetBulge 14, 0
plineObj.SetBulge 15, -0.435695
plineObj.SetBulge 16, -0.0902271
plineObj.SetBulge 17, -0.0902271
plineObj.SetBulge 18, -0.435695
plineObj.SetBulge 19, 0

 plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True
  

'=============================CENTER===========================================
  
  b1(0) = (b / 2) - 1:  b1(1) = (a / 2)
  b2(0) = (b / 2) + 1:  b2(1) = (a / 2)
 
  a1(0) = points(4) - (b / 2): a1(1) = points(1) + (a / 2) - 1
  a2(0) = points(4) - (b / 2): a2(1) = points(1) + (a / 2) + 1
  
bcp = points(0) + (b / 2)
acp = points(1) + (a / 2)

If a >= 320 Then
If b >= 170 Then

pointscntr1(0) = bcp + 24.5833:     pointscntr1(1) = acp + 20.0873
pointscntr1(2) = bcp + 28.0246:     pointscntr1(3) = acp + 19.9972
pointscntr1(4) = bcp + 19.3482:     pointscntr1(5) = acp + 25.688
pointscntr1(6) = bcp + 26.8683:     pointscntr1(7) = acp + 32.6531
pointscntr1(8) = bcp + 27.427:      pointscntr1(9) = acp + 33.6769
pointscntr1(10) = bcp + 27.9934:    pointscntr1(11) = acp + 34.4973
pointscntr1(12) = bcp + 31.4343:    pointscntr1(13) = acp + 35.7712
pointscntr1(14) = bcp + 33.8541:    pointscntr1(15) = acp + 35.1915
pointscntr1(16) = bcp + 37:         pointscntr1(17) = acp + 33.4752
pointscntr1(18) = bcp + 30.8536:    pointscntr1(19) = acp + 38.2705
pointscntr1(20) = bcp + 23.058:     pointscntr1(21) = acp + 38.2446
pointscntr1(22) = bcp + 13.2405:    pointscntr1(23) = acp + 32.8606
pointscntr1(24) = bcp + 6.64503:    pointscntr1(25) = acp + 23.8124
pointscntr1(26) = bcp + 6.64503:    pointscntr1(27) = acp - 23.8124
pointscntr1(28) = bcp + 13.2405:    pointscntr1(29) = acp - 32.8606
pointscntr1(30) = bcp + 23.058:     pointscntr1(31) = acp - 38.2446
pointscntr1(32) = bcp + 30.8536:    pointscntr1(33) = acp - 38.2705
pointscntr1(34) = bcp + 37:         pointscntr1(35) = acp - 33.4752
pointscntr1(36) = bcp + 33.8541:    pointscntr1(37) = acp - 35.1915
pointscntr1(38) = bcp + 31.4343:    pointscntr1(39) = acp - 35.7712
pointscntr1(40) = bcp + 27.9934:    pointscntr1(41) = acp - 34.4973
pointscntr1(42) = bcp + 27.427:     pointscntr1(43) = acp - 33.6769
pointscntr1(44) = bcp + 26.8683:    pointscntr1(45) = acp - 32.6531
pointscntr1(46) = bcp + 19.3482:    pointscntr1(47) = acp - 25.688
pointscntr1(48) = bcp + 28.0246:    pointscntr1(49) = acp - 19.9972
pointscntr1(50) = bcp + 24.5833:    pointscntr1(51) = acp - 20.0873
pointscntr1(52) = bcp + 19.6471:    pointscntr1(53) = acp - 16.403
pointscntr1(54) = bcp + 6.28296:    pointscntr1(55) = acp - 9.466
pointscntr1(56) = bcp + 6.28296:    pointscntr1(57) = acp + 9.466
pointscntr1(58) = bcp + 19.6471:    pointscntr1(59) = acp + 16.403

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update

plineObj.SetBulge 0, -0.190956
plineObj.SetBulge 1, 0.493502
plineObj.SetBulge 2, 0.173886
plineObj.SetBulge 3, 0
plineObj.SetBulge 4, -0.0792127
plineObj.SetBulge 5, -0.236352
plineObj.SetBulge 6, -0.0741078
plineObj.SetBulge 7, -0.0595151
plineObj.SetBulge 8, 0.188765
plineObj.SetBulge 9, 0.132767
plineObj.SetBulge 10, 0.129543
plineObj.SetBulge 11, 0.0905785
plineObj.SetBulge 12, 0.219599
plineObj.SetBulge 13, 0.0905785
plineObj.SetBulge 14, 0.129543
plineObj.SetBulge 15, 0.132767
plineObj.SetBulge 16, 0.188765
plineObj.SetBulge 17, -0.0595151
plineObj.SetBulge 18, -0.0741078
plineObj.SetBulge 19, -0.236352
plineObj.SetBulge 20, -0.0792127
plineObj.SetBulge 21, 0
plineObj.SetBulge 22, 0.173886
plineObj.SetBulge 23, 0.493502
plineObj.SetBulge 24, -0.190956
plineObj.SetBulge 25, -0.133032
plineObj.SetBulge 26, -0.20958
plineObj.SetBulge 27, -0.334034
plineObj.SetBulge 28, -0.20958
plineObj.SetBulge 29, -0.133032

plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True

RetVal = plineObj.Mirror(a1, a2)

pointscntr2(0) = bcp + 6.58238:       pointscntr2(1) = acp + 31.3788
pointscntr2(2) = bcp + 4.25321:       pointscntr2(3) = acp + 33.6762
pointscntr2(4) = bcp + 0.898097:      pointscntr2(5) = acp + 33.4752
pointscntr2(6) = bcp + 1.68328:       pointscntr2(7) = acp + 35.9463
pointscntr2(8) = bcp + 1.70714:       pointscntr2(9) = acp + 37.0063
pointscntr2(10) = bcp + 0.782447:     pointscntr2(11) = acp + 37.995
pointscntr2(12) = bcp + 1.13687E-13:  pointscntr2(13) = acp + 37.9676
pointscntr2(14) = bcp + -0.782447:    pointscntr2(15) = acp + 37.995
pointscntr2(16) = bcp + -1.70714:     pointscntr2(17) = acp + 37.0063
pointscntr2(18) = bcp + -1.68328:     pointscntr2(19) = acp + 35.9463
pointscntr2(20) = bcp + -0.898097:    pointscntr2(21) = acp + 33.4752
pointscntr2(22) = bcp + -4.25321:     pointscntr2(23) = acp + 33.6762
pointscntr2(24) = bcp + -6.58238:     pointscntr2(25) = acp + 31.3788
pointscntr2(26) = bcp + -3.92174:     pointscntr2(27) = acp + 29.4618
pointscntr2(28) = bcp + -1.70295:     pointscntr2(29) = acp + 25.1
pointscntr2(30) = bcp + -0.455718:    pointscntr2(31) = acp + 18.477
pointscntr2(32) = bcp + -0.227606:    pointscntr2(33) = acp + 15.167
pointscntr2(34) = bcp + -0.110535:    pointscntr2(35) = acp + 11.4844
pointscntr2(36) = bcp + -0.0689475:   pointscntr2(37) = acp + 10.7799
pointscntr2(38) = bcp + -0.0388717:   pointscntr2(39) = acp + 10.5512
pointscntr2(40) = bcp + 1.13687E-13:  pointscntr2(41) = acp + 10.4124
pointscntr2(42) = bcp + 0.0388717:    pointscntr2(43) = acp + 10.5512
pointscntr2(44) = bcp + 0.0689475:    pointscntr2(45) = acp + 10.7799
pointscntr2(46) = bcp + 0.110535:     pointscntr2(47) = acp + 11.4844
pointscntr2(48) = bcp + 0.227606:     pointscntr2(49) = acp + 15.167
pointscntr2(50) = bcp + 0.455718:     pointscntr2(51) = acp + 18.477
pointscntr2(52) = bcp + 1.70295:      pointscntr2(53) = acp + 25.1
pointscntr2(54) = bcp + 3.92174:      pointscntr2(55) = acp + 29.4618


Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr2)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update

plineObj.SetBulge 0, 0.130293
plineObj.SetBulge 1, 0.286262
plineObj.SetBulge 2, 0.0628827
plineObj.SetBulge 3, 0.0664767
plineObj.SetBulge 4, 0.354054
plineObj.SetBulge 5, 0.0749942
plineObj.SetBulge 6, 0.0749942
plineObj.SetBulge 7, 0.354054
plineObj.SetBulge 8, 0.0664767
plineObj.SetBulge 9, 0.0628827
plineObj.SetBulge 10, 0.286262
plineObj.SetBulge 11, 0.130293
plineObj.SetBulge 12, -0.132961
plineObj.SetBulge 13, -0.0949138
plineObj.SetBulge 14, -0.0510213
plineObj.SetBulge 15, -0.0125717
plineObj.SetBulge 16, -0.00334278
plineObj.SetBulge 17, 0
plineObj.SetBulge 18, 0
plineObj.SetBulge 19, 0
plineObj.SetBulge 20, 0
plineObj.SetBulge 21, 0
plineObj.SetBulge 22, 0
plineObj.SetBulge 23, -0.00334278
plineObj.SetBulge 24, -0.0125717
plineObj.SetBulge 25, -0.0510213
plineObj.SetBulge 26, -0.0949138
plineObj.SetBulge 27, -0.132961

plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True

RetVal = plineObj.Mirror(b1, b2)

pointscntr3(0) = bcp + 17.03:       pointscntr3(1) = acp + 80.0281
pointscntr3(2) = bcp + 12.3194:     pointscntr3(3) = acp + 76.2511
pointscntr3(4) = bcp + 12.7291:     pointscntr3(5) = acp + 73.3674
pointscntr3(6) = bcp + 16.4548:     pointscntr3(7) = acp + 74.8083
pointscntr3(8) = bcp + 22.4096:     pointscntr3(9) = acp + 72.6438
pointscntr3(10) = bcp + 24.1819:    pointscntr3(11) = acp + 69.1113
pointscntr3(12) = bcp + 24.8313:    pointscntr3(13) = acp + 64.3255
pointscntr3(14) = bcp + 23.9272:    pointscntr3(15) = acp + 59.0018
pointscntr3(16) = bcp + 16.5438:    pointscntr3(17) = acp + 46.828
pointscntr3(18) = bcp + 6.28228:    pointscntr3(19) = acp + 36.4703
pointscntr3(20) = bcp + 22.0575:    pointscntr3(21) = acp + 48.4905
pointscntr3(22) = bcp + 28.7564:    pointscntr3(23) = acp + 65.9369
pointscntr3(24) = bcp + 25.5548:    pointscntr3(25) = acp + 74.964

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr3)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update

plineObj.SetBulge 0, 0.424127
plineObj.SetBulge 1, 0.161605
plineObj.SetBulge 2, 0.83194
plineObj.SetBulge 3, -0.267392
plineObj.SetBulge 4, -0.106986
plineObj.SetBulge 5, -0.0601812
plineObj.SetBulge 6, -0.0862239
plineObj.SetBulge 7, -0.0871434
plineObj.SetBulge 8, -0.0408194
plineObj.SetBulge 9, 0.0867111
plineObj.SetBulge 10, 0.208082
plineObj.SetBulge 11, 0.154888
plineObj.SetBulge 12, 0.189184

plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True
RetVal = plineObj.Mirror(b1, b2)
RetVal = plineObj.Mirror(a1, a2)
RetVal = plineObj.Copy
  ' Define the rotation of 45 degrees about a
  ' base point of (4, 4.25, 0)
  basePoint(0) = points(4) - (b / 2): basePoint(1) = points(3) - (a / 2)
  rotationAngle = 3.14159   ' 45 degrees
  plineObj.Rotate basePoint, rotationAngle
  plineObj.Update
  
pointscntr4(0) = bcp + 6.97983:      pointscntr4(1) = acp + 66.4733
pointscntr4(2) = bcp + 4.40328:      pointscntr4(3) = acp + 70.4044
pointscntr4(4) = bcp + 2.76705:      pointscntr4(5) = acp + 73.302
pointscntr4(6) = bcp + 2.03069:      pointscntr4(7) = acp + 75.4075
pointscntr4(8) = bcp + 1.13687E-13:  pointscntr4(9) = acp + 82.896
pointscntr4(10) = bcp + -2.03069:    pointscntr4(11) = acp + 75.4075
pointscntr4(12) = bcp + -2.76705:    pointscntr4(13) = acp + 73.302
pointscntr4(14) = bcp + -4.35149:    pointscntr4(15) = acp + 70.4834
pointscntr4(16) = bcp + -6.97983:    pointscntr4(17) = acp + 66.4733
pointscntr4(18) = bcp + -8.62407:    pointscntr4(19) = acp + 63.7027
pointscntr4(20) = bcp + -9.8543:     pointscntr4(21) = acp + 57.1808
pointscntr4(22) = bcp + -8.14995:    pointscntr4(23) = acp + 53.3016
pointscntr4(24) = bcp + -5.61959:    pointscntr4(25) = acp + 51.6618
pointscntr4(26) = bcp + -2.27229:    pointscntr4(27) = acp + 51.8809
pointscntr4(28) = bcp + 1.13687E-13: pointscntr4(29) = acp + 55.0398
pointscntr4(30) = bcp + 2.27229:     pointscntr4(31) = acp + 51.8809
pointscntr4(32) = bcp + 5.61959:     pointscntr4(33) = acp + 51.6618
pointscntr4(34) = bcp + 8.14995:     pointscntr4(35) = acp + 53.3016
pointscntr4(36) = bcp + 9.8543:      pointscntr4(37) = acp + 57.1808
pointscntr4(38) = bcp + 8.62407:     pointscntr4(39) = acp + 63.7027

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr4)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update

plineObj.SetBulge 0, 0
plineObj.SetBulge 1, -0.0340478
plineObj.SetBulge 2, -0.0404214
plineObj.SetBulge 3, 0
plineObj.SetBulge 4, 0
plineObj.SetBulge 5, -0.0404214
plineObj.SetBulge 6, -0.0330812
plineObj.SetBulge 7, 0
plineObj.SetBulge 8, 0.021196
plineObj.SetBulge 9, 0.173455
plineObj.SetBulge 10, 0.145317
plineObj.SetBulge 11, 0.143574
plineObj.SetBulge 12, 0.203485
plineObj.SetBulge 13, 0.200979
plineObj.SetBulge 14, 0.200979
plineObj.SetBulge 15, 0.203485
plineObj.SetBulge 16, 0.143574
plineObj.SetBulge 17, 0.145317
plineObj.SetBulge 18, 0.173455
plineObj.SetBulge 19, 0.021196

plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True
RetVal = plineObj.Mirror(b1, b2)

pointscntr5(0) = bcp + 8.66267:     pointscntr5(1) = acp + 94.5619
pointscntr5(2) = bcp + 12.0028:     pointscntr5(3) = acp + 94.7543
pointscntr5(4) = bcp + 14.9353:     pointscntr5(5) = acp + 93.6436
pointscntr5(6) = bcp + 16.3813:     pointscntr5(7) = acp + 91.4456
pointscntr5(8) = bcp + 16.3602:     pointscntr5(9) = acp + 89.5013
pointscntr5(10) = bcp + 14.3893:    pointscntr5(11) = acp + 86.8541
pointscntr5(12) = bcp + 11.3686:    pointscntr5(13) = acp + 86.1899
pointscntr5(14) = bcp + 13.0885:    pointscntr5(15) = acp + 87.6125
pointscntr5(16) = bcp + 12.9123:    pointscntr5(17) = acp + 89.239
pointscntr5(18) = bcp + 10.7834:    pointscntr5(19) = acp + 90.4247
pointscntr5(20) = bcp + 5.27888:    pointscntr5(21) = acp + 89.7626
pointscntr5(22) = bcp + 2.39309:    pointscntr5(23) = acp + 86.7891
pointscntr5(24) = bcp + 1.49609:    pointscntr5(25) = acp + 88.2863
pointscntr5(26) = bcp + 4.32014:    pointscntr5(27) = acp + 92.3133

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr5)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update
  
plineObj.SetBulge 0, -0.0883845
plineObj.SetBulge 1, -0.121995
plineObj.SetBulge 2, -0.193108
plineObj.SetBulge 3, -0.0993998
plineObj.SetBulge 4, -0.23388
plineObj.SetBulge 5, -0.125673
plineObj.SetBulge 6, 0.11811
plineObj.SetBulge 7, 0.412476
plineObj.SetBulge 8, 0.0962103
plineObj.SetBulge 9, 0.224454
plineObj.SetBulge 10, 0.11374
plineObj.SetBulge 11, 0
plineObj.SetBulge 12, -0.11562
plineObj.SetBulge 13, -0.128851

plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True
RetVal = plineObj.Mirror(b1, b2)
RetVal = plineObj.Mirror(a1, a2)
RetVal = plineObj.Copy
  ' Define the rotation of 45 degrees about a
  ' base point of (4, 4.25, 0)
  basePoint(0) = points(4) - (b / 2): basePoint(1) = points(3) - (a / 2)
  rotationAngle = 3.14159   ' 45 degrees
  plineObj.Rotate basePoint, rotationAngle
  plineObj.Update

pointscntr6(0) = bcp + 3.75304:       pointscntr6(1) = acp + 101.938
pointscntr6(2) = bcp + 1.13687E-13:   pointscntr6(3) = acp + 111.5
pointscntr6(4) = bcp + -3.75304:      pointscntr6(5) = acp + 101.938
pointscntr6(6) = bcp + -3.37708:      pointscntr6(7) = acp + 96.8397
pointscntr6(8) = bcp + 1.13687E-13:   pointscntr6(9) = acp + 88.5857
pointscntr6(10) = bcp + 3.37708:      pointscntr6(11) = acp + 96.8397

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr6)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update
  
plineObj.SetBulge 0, 0.06872
plineObj.SetBulge 1, 0.06872
plineObj.SetBulge 2, 0.171056
plineObj.SetBulge 3, 0
plineObj.SetBulge 4, 0
plineObj.SetBulge 5, 0.171056

    plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True
RetVal = plineObj.Mirror(b1, b2)

center(0) = a1(0): center(1) = b1(1) + 45.76: center(2) = 0: radius = 2.9613
Set circleObj = ThisDrawing.ModelSpace.AddCircle(center, radius)
circleObj.Layer = "K-grav_Pattern"
circleObj.Update
center(0) = a1(0): center(1) = b1(1) - 45.76: center(2) = 0: radius = 2.9613
Set circleObj = ThisDrawing.ModelSpace.AddCircle(center, radius)
circleObj.Layer = "K-grav_Pattern"
circleObj.Update

End If
End If
End If
End If
  
I = 100
  
 Dim plineObj2 As AcadLWPolyline
  Dim points2(0 To 7) As Double
  
  points2(0) = points(0):    points2(1) = a + I:
  points2(2) = points(0):    points2(3) = (a * 2) + I:
  points2(4) = points(6):    points2(5) = (a * 2) + I:
  points2(6) = points(6):    points2(7) = a + I:
  
'========================================================
'========================================================
'========================================================
  
 For d = 2 To Cells(c, 4)
If d > 1 Then

Set plineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(points2)
  plineObj2.Closed = True
  
If a >= 140 Then
If b >= 140 Then
  
  pointswithin(0) = points2(0) + 70:                              pointswithin(1) = points2(1) + 100
  pointswithin(2) = points2(0) + 70:                              pointswithin(3) = points2(3) - 100
  pointswithin(4) = points2(0) + 100:                             pointswithin(5) = points2(3) - 70
  pointswithin(6) = points2(4) - 100:                             pointswithin(7) = points2(3) - 70
  pointswithin(8) = points2(4) - 70:                              pointswithin(9) = points2(3) - 100
  pointswithin(10) = points2(4) - 70:                             pointswithin(11) = points2(1) + 100
  pointswithin(12) = points2(4) - 100:                            pointswithin(13) = points2(1) + 70
  pointswithin(14) = points2(0) + 100:                            pointswithin(15) = points2(1) + 70

If a > 241 Then
If b > 241 Then

 Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointswithin)
  plineObj.Closed = True
   
   ' Find the bulge of the third segment
    currentBulge = plineObj.GetBulge(2)
    l = 2 * r * (Sqr(2) / 2)
    h = r * (1 - (Sqr(2) / 2))
    k = h / (l / 2)
    
    ' Change the bulge of the third segment
    plineObj.SetBulge 1, k
    plineObj.SetBulge 3, k
    plineObj.SetBulge 5, k
    plineObj.SetBulge 7, k
    plineObj.Layer = "C_2"
    plineObj.Update
    plineObj.Closed = True
    
  offsetObj = plineObj.Offset(6)
    plineObj.Layer = "C-Mill"
    plineObj.Update

End If
End If
  
  pointswithin2(0) = points2(0) + 64:                              pointswithin2(1) = points2(1) + 100
  pointswithin2(2) = points2(0) + 64:                              pointswithin2(3) = points2(3) - 100
  pointswithin2(4) = points2(0) + 70:                              pointswithin2(5) = points2(3) - 94
  pointswithin2(6) = points2(0) + 94:                              pointswithin2(7) = points2(3) - 70
  pointswithin2(8) = points2(0) + 100:                             pointswithin2(9) = points2(3) - 64
  pointswithin2(10) = points2(4) - 100:                            pointswithin2(11) = points2(3) - 64
  pointswithin2(12) = points2(4) - 94:                             pointswithin2(13) = points2(3) - 70
  pointswithin2(14) = points2(4) - 70:                             pointswithin2(15) = points2(3) - 94
  pointswithin2(16) = points2(4) - 64:                             pointswithin2(17) = points2(3) - 100
  pointswithin2(18) = points2(4) - 64:                             pointswithin2(19) = points2(1) + 100
  pointswithin2(20) = points2(4) - 70:                             pointswithin2(21) = points2(1) + 94
  pointswithin2(22) = points2(4) - 94:                             pointswithin2(23) = points2(1) + 70
  pointswithin2(24) = points2(4) - 100:                            pointswithin2(25) = points2(1) + 64
  pointswithin2(26) = points2(0) + 100:                            pointswithin2(27) = points2(1) + 64
  pointswithin2(28) = points2(0) + 94:                             pointswithin2(29) = points2(1) + 70
  pointswithin2(30) = points2(0) + 70:                             pointswithin2(31) = points2(1) + 94


If a > 241 Then
If b > 241 Then

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointswithin2)
  plineObj.Closed = True
  
   plineObj.SetBulge 1, -k
    plineObj.SetBulge 2, k
    plineObj.SetBulge 3, -k
    plineObj.SetBulge 5, -k
    plineObj.SetBulge 6, k
    plineObj.SetBulge 7, -k
    plineObj.SetBulge 9, -k
    plineObj.SetBulge 10, k
    plineObj.SetBulge 11, -k
    plineObj.SetBulge 13, -k
    plineObj.SetBulge 14, k
    plineObj.SetBulge 15, -k
   
    plineObj.Layer = "C_2"
    plineObj.Update
    plineObj.Closed = True
  
End If
End If

pointsaktr1(0) = points2(0) + 60.5:      pointsaktr1(1) = points2(1) + 33
pointsaktr1(2) = points2(0) + 62.8818:   pointsaktr1(3) = points2(1) + 37.6743
pointsaktr1(4) = points2(0) + 59.9507:   pointsaktr1(5) = points2(1) + 41.9916
pointsaktr1(6) = points2(0) + 58.0126:   pointsaktr1(7) = points2(1) + 41.8673
pointsaktr1(8) = points2(0) + 56.7296:   pointsaktr1(9) = points2(1) + 39.5607
pointsaktr1(10) = points2(0) + 58.5907:  pointsaktr1(11) = points2(1) + 37.5362
pointsaktr1(12) = points2(0) + 59.6991:  pointsaktr1(13) = points2(1) + 38.013
pointsaktr1(14) = points2(0) + 60.022:   pointsaktr1(15) = points2(1) + 40.2373
pointsaktr1(16) = points2(0) + 61.2812:  pointsaktr1(17) = points2(1) + 38.1637
pointsaktr1(18) = points2(0) + 59.0042:  pointsaktr1(19) = points2(1) + 34.633
pointsaktr1(20) = points2(0) + 55.4873:  pointsaktr1(21) = points2(1) + 33.7236
pointsaktr1(22) = points2(0) + 49.9128:  pointsaktr1(23) = points2(1) + 36.6409
pointsaktr1(24) = points2(0) + 47.6162:  pointsaktr1(25) = points2(1) + 44.2001
pointsaktr1(26) = points2(0) + 51.7037:  pointsaktr1(27) = points2(1) + 49.0761
pointsaktr1(28) = points2(0) + 54.3447:  pointsaktr1(29) = points2(1) + 48.7943
pointsaktr1(30) = points2(0) + 55.7883:  pointsaktr1(31) = points2(1) + 48.3252
pointsaktr1(32) = points2(0) + 57.2677:  pointsaktr1(33) = points2(1) + 48.0411
pointsaktr1(34) = points2(0) + 59.2488:  pointsaktr1(35) = points2(1) + 48.7514
pointsaktr1(36) = points2(0) + 59.4835:  pointsaktr1(37) = points2(1) + 50.9641
pointsaktr1(38) = points2(0) + 57.5849:  pointsaktr1(39) = points2(1) + 52.2794
pointsaktr1(40) = points2(0) + 54.413:   pointsaktr1(41) = points2(1) + 52.5139
pointsaktr1(42) = points2(0) + 59.6861:  pointsaktr1(43) = points2(1) + 53.7975
pointsaktr1(44) = points2(0) + 61.8737:  pointsaktr1(45) = points2(1) + 53.3042
pointsaktr1(46) = points2(0) + 62.2884:  pointsaktr1(47) = points2(1) + 53.366
pointsaktr1(48) = points2(0) + 62.614:   pointsaktr1(49) = points2(1) + 53.7765
pointsaktr1(50) = points2(0) + 62.5956:  pointsaktr1(51) = points2(1) + 54.4812
pointsaktr1(52) = points2(0) + 62.0533:  pointsaktr1(53) = points2(1) + 55.3326
pointsaktr1(54) = points2(0) + 61.8627:  pointsaktr1(55) = points2(1) + 55.9772
pointsaktr1(56) = points2(0) + 62.0122:  pointsaktr1(57) = points2(1) + 56.5534
pointsaktr1(58) = points2(0) + 62.2419:  pointsaktr1(59) = points2(1) + 57.4669
pointsaktr1(60) = points2(0) + 61.7068:  pointsaktr1(61) = points2(1) + 61.3268
pointsaktr1(62) = points2(0) + 61.5862:  pointsaktr1(63) = points2(1) + 61.5862
pointsaktr1(64) = points2(0) + 61.3268:  pointsaktr1(65) = points2(1) + 61.7068
pointsaktr1(66) = points2(0) + 57.4669:  pointsaktr1(67) = points2(1) + 62.2419
pointsaktr1(68) = points2(0) + 56.5534:  pointsaktr1(69) = points2(1) + 62.0122
pointsaktr1(70) = points2(0) + 55.9772:  pointsaktr1(71) = points2(1) + 61.8627
pointsaktr1(72) = points2(0) + 55.3326:  pointsaktr1(73) = points2(1) + 62.0533
pointsaktr1(74) = points2(0) + 54.4812:  pointsaktr1(75) = points2(1) + 62.5956
pointsaktr1(76) = points2(0) + 53.7765:  pointsaktr1(77) = points2(1) + 62.614
pointsaktr1(78) = points2(0) + 53.366:   pointsaktr1(79) = points2(1) + 62.2884
pointsaktr1(80) = points2(0) + 53.3042:  pointsaktr1(81) = points2(1) + 61.8737
pointsaktr1(82) = points2(0) + 53.7975:  pointsaktr1(83) = points2(1) + 59.6861
pointsaktr1(84) = points2(0) + 52.5139:  pointsaktr1(85) = points2(1) + 54.413
pointsaktr1(86) = points2(0) + 52.2794:  pointsaktr1(87) = points2(1) + 57.5849
pointsaktr1(88) = points2(0) + 50.9641:  pointsaktr1(89) = points2(1) + 59.4835
pointsaktr1(90) = points2(0) + 48.7514:  pointsaktr1(91) = points2(1) + 59.2488
pointsaktr1(92) = points2(0) + 48.0411:  pointsaktr1(93) = points2(1) + 57.2677
pointsaktr1(94) = points2(0) + 48.3252:  pointsaktr1(95) = points2(1) + 55.7883
pointsaktr1(96) = points2(0) + 48.7943:  pointsaktr1(97) = points2(1) + 54.3447
pointsaktr1(98) = points2(0) + 49.0761:  pointsaktr1(99) = points2(1) + 51.7037
pointsaktr1(100) = points2(0) + 44.2001: pointsaktr1(101) = points2(1) + 47.6162
pointsaktr1(102) = points2(0) + 36.6409: pointsaktr1(103) = points2(1) + 49.9128
pointsaktr1(104) = points2(0) + 33.7236: pointsaktr1(105) = points2(1) + 55.4873
pointsaktr1(106) = points2(0) + 34.633:  pointsaktr1(107) = points2(1) + 59.0042
pointsaktr1(108) = points2(0) + 38.1637: pointsaktr1(109) = points2(1) + 61.2812
pointsaktr1(110) = points2(0) + 40.2373: pointsaktr1(111) = points2(1) + 60.022
pointsaktr1(112) = points2(0) + 38.013:  pointsaktr1(113) = points2(1) + 59.6991
pointsaktr1(114) = points2(0) + 37.5362: pointsaktr1(115) = points2(1) + 58.5907
pointsaktr1(116) = points2(0) + 39.5607: pointsaktr1(117) = points2(1) + 56.7296
pointsaktr1(118) = points2(0) + 41.8673: pointsaktr1(119) = points2(1) + 58.0126
pointsaktr1(120) = points2(0) + 41.9916: pointsaktr1(121) = points2(1) + 59.9507
pointsaktr1(122) = points2(0) + 37.6743: pointsaktr1(123) = points2(1) + 62.8818
pointsaktr1(124) = points2(0) + 33:      pointsaktr1(125) = points2(1) + 60.5
pointsaktr1(126) = points2(0) + 33:      pointsaktr1(127) = points2(3) - 60.5
pointsaktr1(128) = points2(0) + 37.6743: pointsaktr1(129) = points2(3) - 62.8818
pointsaktr1(130) = points2(0) + 41.9916: pointsaktr1(131) = points2(3) - 59.9507
pointsaktr1(132) = points2(0) + 41.8673: pointsaktr1(133) = points2(3) - 58.0126
pointsaktr1(134) = points2(0) + 39.5607: pointsaktr1(135) = points2(3) - 56.7296
pointsaktr1(136) = points2(0) + 37.5362: pointsaktr1(137) = points2(3) - 58.5907
pointsaktr1(138) = points2(0) + 38.013:  pointsaktr1(139) = points2(3) - 59.6991
pointsaktr1(140) = points2(0) + 40.2373: pointsaktr1(141) = points2(3) - 60.022
pointsaktr1(142) = points2(0) + 38.1637: pointsaktr1(143) = points2(3) - 61.2812
pointsaktr1(144) = points2(0) + 34.633:  pointsaktr1(145) = points2(3) - 59.0042
pointsaktr1(146) = points2(0) + 33.7236: pointsaktr1(147) = points2(3) - 55.4873
pointsaktr1(148) = points2(0) + 36.6409: pointsaktr1(149) = points2(3) - 49.9128
pointsaktr1(150) = points2(0) + 44.2001: pointsaktr1(151) = points2(3) - 47.6162
pointsaktr1(152) = points2(0) + 49.0761: pointsaktr1(153) = points2(3) - 51.7037
pointsaktr1(154) = points2(0) + 48.7943: pointsaktr1(155) = points2(3) - 54.3447
pointsaktr1(156) = points2(0) + 48.3252: pointsaktr1(157) = points2(3) - 55.7883
pointsaktr1(158) = points2(0) + 48.0411: pointsaktr1(159) = points2(3) - 57.2677
pointsaktr1(160) = points2(0) + 48.7514: pointsaktr1(161) = points2(3) - 59.2488
pointsaktr1(162) = points2(0) + 50.9641: pointsaktr1(163) = points2(3) - 59.4835
pointsaktr1(164) = points2(0) + 52.2794: pointsaktr1(165) = points2(3) - 57.5849
pointsaktr1(166) = points2(0) + 52.5139: pointsaktr1(167) = points2(3) - 54.413
pointsaktr1(168) = points2(0) + 53.7975: pointsaktr1(169) = points2(3) - 59.6861
pointsaktr1(170) = points2(0) + 53.3042: pointsaktr1(171) = points2(3) - 61.8737
pointsaktr1(172) = points2(0) + 53.366:  pointsaktr1(173) = points2(3) - 62.2884
pointsaktr1(174) = points2(0) + 53.7765: pointsaktr1(175) = points2(3) - 62.614
pointsaktr1(176) = points2(0) + 54.4812: pointsaktr1(177) = points2(3) - 62.5956
pointsaktr1(178) = points2(0) + 55.3326: pointsaktr1(179) = points2(3) - 62.0533
pointsaktr1(180) = points2(0) + 55.9772: pointsaktr1(181) = points2(3) - 61.8627
pointsaktr1(182) = points2(0) + 56.5534: pointsaktr1(183) = points2(3) - 62.0122
pointsaktr1(184) = points2(0) + 57.4669: pointsaktr1(185) = points2(3) - 62.2419
pointsaktr1(186) = points2(0) + 61.3268: pointsaktr1(187) = points2(3) - 61.7068
pointsaktr1(188) = points2(0) + 61.5862: pointsaktr1(189) = points2(3) - 61.5862
pointsaktr1(190) = points2(0) + 61.7068: pointsaktr1(191) = points2(3) - 61.3268
pointsaktr1(192) = points2(0) + 62.2419: pointsaktr1(193) = points2(3) - 57.4669
pointsaktr1(194) = points2(0) + 62.0122: pointsaktr1(195) = points2(3) - 56.5534
pointsaktr1(196) = points2(0) + 61.8627: pointsaktr1(197) = points2(3) - 55.9772
pointsaktr1(198) = points2(0) + 62.0533: pointsaktr1(199) = points2(3) - 55.3326
pointsaktr1(200) = points2(0) + 62.5956: pointsaktr1(201) = points2(3) - 54.4812
pointsaktr1(202) = points2(0) + 62.614:  pointsaktr1(203) = points2(3) - 53.7765
pointsaktr1(204) = points2(0) + 62.2884: pointsaktr1(205) = points2(3) - 53.366
pointsaktr1(206) = points2(0) + 61.8737: pointsaktr1(207) = points2(3) - 53.3042
pointsaktr1(208) = points2(0) + 59.6861: pointsaktr1(209) = points2(3) - 53.7975
pointsaktr1(210) = points2(0) + 54.413:  pointsaktr1(211) = points2(3) - 52.5139
pointsaktr1(212) = points2(0) + 57.5849: pointsaktr1(213) = points2(3) - 52.2794
pointsaktr1(214) = points2(0) + 59.4835: pointsaktr1(215) = points2(3) - 50.9641
pointsaktr1(216) = points2(0) + 59.2488: pointsaktr1(217) = points2(3) - 48.7514
pointsaktr1(218) = points2(0) + 57.2677: pointsaktr1(219) = points2(3) - 48.0411
pointsaktr1(220) = points2(0) + 55.7883: pointsaktr1(221) = points2(3) - 48.3252
pointsaktr1(222) = points2(0) + 54.3447: pointsaktr1(223) = points2(3) - 48.7943
pointsaktr1(224) = points2(0) + 51.7037: pointsaktr1(225) = points2(3) - 49.0761
pointsaktr1(226) = points2(0) + 47.6162: pointsaktr1(227) = points2(3) - 44.2001
pointsaktr1(228) = points2(0) + 49.9128: pointsaktr1(229) = points2(3) - 36.6409
pointsaktr1(230) = points2(0) + 55.4873: pointsaktr1(231) = points2(3) - 33.7236
pointsaktr1(232) = points2(0) + 59.0042: pointsaktr1(233) = points2(3) - 34.633
pointsaktr1(234) = points2(0) + 61.2812: pointsaktr1(235) = points2(3) - 38.1637
pointsaktr1(236) = points2(0) + 60.022:  pointsaktr1(237) = points2(3) - 40.2373
pointsaktr1(238) = points2(0) + 59.6991: pointsaktr1(239) = points2(3) - 38.013
pointsaktr1(240) = points2(0) + 58.5907: pointsaktr1(241) = points2(3) - 37.5362
pointsaktr1(242) = points2(0) + 56.7296: pointsaktr1(243) = points2(3) - 39.5607
pointsaktr1(244) = points2(0) + 58.0126: pointsaktr1(245) = points2(3) - 41.8673
pointsaktr1(246) = points2(0) + 59.9507: pointsaktr1(247) = points2(3) - 41.9916
pointsaktr1(248) = points2(0) + 62.8818: pointsaktr1(249) = points2(3) - 37.6743
pointsaktr1(250) = points2(0) + 60.5:    pointsaktr1(251) = points2(3) - 33
pointsaktr1(252) = points2(4) - 60.5:    pointsaktr1(253) = points2(3) - 33
pointsaktr1(254) = points2(4) - 62.8818: pointsaktr1(255) = points2(3) - 37.6743
pointsaktr1(256) = points2(4) - 59.9507: pointsaktr1(257) = points2(3) - 41.9916
pointsaktr1(258) = points2(4) - 58.0126: pointsaktr1(259) = points2(3) - 41.8673
pointsaktr1(260) = points2(4) - 56.7296: pointsaktr1(261) = points2(3) - 39.5607
pointsaktr1(262) = points2(4) - 58.5907: pointsaktr1(263) = points2(3) - 37.5362
pointsaktr1(264) = points2(4) - 59.6991: pointsaktr1(265) = points2(3) - 38.013
pointsaktr1(266) = points2(4) - 60.022:  pointsaktr1(267) = points2(3) - 40.2373
pointsaktr1(268) = points2(4) - 61.2812: pointsaktr1(269) = points2(3) - 38.1637
pointsaktr1(270) = points2(4) - 59.0042: pointsaktr1(271) = points2(3) - 34.633
pointsaktr1(272) = points2(4) - 55.4873: pointsaktr1(273) = points2(3) - 33.7236
pointsaktr1(274) = points2(4) - 49.9128: pointsaktr1(275) = points2(3) - 36.6409
pointsaktr1(276) = points2(4) - 47.6162: pointsaktr1(277) = points2(3) - 44.2001
pointsaktr1(278) = points2(4) - 51.7037: pointsaktr1(279) = points2(3) - 49.0761
pointsaktr1(280) = points2(4) - 54.3447: pointsaktr1(281) = points2(3) - 48.7943
pointsaktr1(282) = points2(4) - 55.7883: pointsaktr1(283) = points2(3) - 48.3252
pointsaktr1(284) = points2(4) - 57.2677: pointsaktr1(285) = points2(3) - 48.0411
pointsaktr1(286) = points2(4) - 59.2488: pointsaktr1(287) = points2(3) - 48.7514
pointsaktr1(288) = points2(4) - 59.4835: pointsaktr1(289) = points2(3) - 50.9641
pointsaktr1(290) = points2(4) - 57.5849: pointsaktr1(291) = points2(3) - 52.2794
pointsaktr1(292) = points2(4) - 54.413:  pointsaktr1(293) = points2(3) - 52.5139
pointsaktr1(294) = points2(4) - 59.6861: pointsaktr1(295) = points2(3) - 53.7975
pointsaktr1(296) = points2(4) - 61.8737: pointsaktr1(297) = points2(3) - 53.3042
pointsaktr1(298) = points2(4) - 62.2884: pointsaktr1(299) = points2(3) - 53.366
pointsaktr1(300) = points2(4) - 62.614:  pointsaktr1(301) = points2(3) - 53.7765
pointsaktr1(302) = points2(4) - 62.5956: pointsaktr1(303) = points2(3) - 54.4812
pointsaktr1(304) = points2(4) - 62.0533: pointsaktr1(305) = points2(3) - 55.3326
pointsaktr1(306) = points2(4) - 61.8627: pointsaktr1(307) = points2(3) - 55.9772
pointsaktr1(308) = points2(4) - 62.0122: pointsaktr1(309) = points2(3) - 56.5534
pointsaktr1(310) = points2(4) - 62.2419: pointsaktr1(311) = points2(3) - 57.4669
pointsaktr1(312) = points2(4) - 61.7068: pointsaktr1(313) = points2(3) - 61.3268
pointsaktr1(314) = points2(4) - 61.5862: pointsaktr1(315) = points2(3) - 61.5862
pointsaktr1(316) = points2(4) - 61.3268: pointsaktr1(317) = points2(3) - 61.7068
pointsaktr1(318) = points2(4) - 57.4669: pointsaktr1(319) = points2(3) - 62.2419
pointsaktr1(320) = points2(4) - 56.5534: pointsaktr1(321) = points2(3) - 62.0122
pointsaktr1(322) = points2(4) - 55.9772: pointsaktr1(323) = points2(3) - 61.8627
pointsaktr1(324) = points2(4) - 55.3326: pointsaktr1(325) = points2(3) - 62.0533
pointsaktr1(326) = points2(4) - 54.4812: pointsaktr1(327) = points2(3) - 62.5956
pointsaktr1(328) = points2(4) - 53.7765: pointsaktr1(329) = points2(3) - 62.614
pointsaktr1(330) = points2(4) - 53.366:  pointsaktr1(331) = points2(3) - 62.2884
pointsaktr1(332) = points2(4) - 53.3042: pointsaktr1(333) = points2(3) - 61.8737
pointsaktr1(334) = points2(4) - 53.7975: pointsaktr1(335) = points2(3) - 59.6861
pointsaktr1(336) = points2(4) - 52.5139: pointsaktr1(337) = points2(3) - 54.413
pointsaktr1(338) = points2(4) - 52.2794: pointsaktr1(339) = points2(3) - 57.5849
pointsaktr1(340) = points2(4) - 50.9641: pointsaktr1(341) = points2(3) - 59.4835
pointsaktr1(342) = points2(4) - 48.7514: pointsaktr1(343) = points2(3) - 59.2488
pointsaktr1(344) = points2(4) - 48.0411: pointsaktr1(345) = points2(3) - 57.2677
pointsaktr1(346) = points2(4) - 48.3252: pointsaktr1(347) = points2(3) - 55.7883
pointsaktr1(348) = points2(4) - 48.7943: pointsaktr1(349) = points2(3) - 54.3447
pointsaktr1(350) = points2(4) - 49.0761: pointsaktr1(351) = points2(3) - 51.7037
pointsaktr1(352) = points2(4) - 44.2001: pointsaktr1(353) = points2(3) - 47.6162
pointsaktr1(354) = points2(4) - 36.6409: pointsaktr1(355) = points2(3) - 49.9128
pointsaktr1(356) = points2(4) - 33.7236: pointsaktr1(357) = points2(3) - 55.4873
pointsaktr1(358) = points2(4) - 34.633:  pointsaktr1(359) = points2(3) - 59.0042
pointsaktr1(360) = points2(4) - 38.1637: pointsaktr1(361) = points2(3) - 61.2812
pointsaktr1(362) = points2(4) - 40.2373: pointsaktr1(363) = points2(3) - 60.022
pointsaktr1(364) = points2(4) - 38.013:  pointsaktr1(365) = points2(3) - 59.6991
pointsaktr1(366) = points2(4) - 37.5362: pointsaktr1(367) = points2(3) - 58.5907
pointsaktr1(368) = points2(4) - 39.5607: pointsaktr1(369) = points2(3) - 56.7296
pointsaktr1(370) = points2(4) - 41.8673: pointsaktr1(371) = points2(3) - 58.0126
pointsaktr1(372) = points2(4) - 41.9916: pointsaktr1(373) = points2(3) - 59.9507
pointsaktr1(374) = points2(4) - 37.6743: pointsaktr1(375) = points2(3) - 62.8818
pointsaktr1(376) = points2(4) - 33:  pointsaktr1(377) = points2(3) - 60.5
pointsaktr1(378) = points2(4) - 33:  pointsaktr1(379) = points2(1) + 60.5
pointsaktr1(380) = points2(4) - 37.6743: pointsaktr1(381) = points2(1) + 62.8818
pointsaktr1(382) = points2(4) - 41.9916: pointsaktr1(383) = points2(1) + 59.9507
pointsaktr1(384) = points2(4) - 41.8673: pointsaktr1(385) = points2(1) + 58.0126
pointsaktr1(386) = points2(4) - 39.5607: pointsaktr1(387) = points2(1) + 56.7296
pointsaktr1(388) = points2(4) - 37.5362: pointsaktr1(389) = points2(1) + 58.5907
pointsaktr1(390) = points2(4) - 38.013:  pointsaktr1(391) = points2(1) + 59.6991
pointsaktr1(392) = points2(4) - 40.2373: pointsaktr1(393) = points2(1) + 60.022
pointsaktr1(394) = points2(4) - 38.1637: pointsaktr1(395) = points2(1) + 61.2812
pointsaktr1(396) = points2(4) - 34.633:  pointsaktr1(397) = points2(1) + 59.0042
pointsaktr1(398) = points2(4) - 33.7236: pointsaktr1(399) = points2(1) + 55.4873
pointsaktr1(400) = points2(4) - 36.6409:   pointsaktr1(401) = points2(1) + 49.9128
pointsaktr1(402) = points2(4) - 44.2001: pointsaktr1(403) = points2(1) + 47.6162
pointsaktr1(404) = points2(4) - 49.0761: pointsaktr1(405) = points2(1) + 51.7037
pointsaktr1(406) = points2(4) - 48.7943:  pointsaktr1(407) = points2(1) + 54.3447
pointsaktr1(408) = points2(4) - 48.3252:  pointsaktr1(409) = points2(1) + 55.7883
pointsaktr1(410) = points2(4) - 48.0411:  pointsaktr1(411) = points2(1) + 57.2677
pointsaktr1(412) = points2(4) - 48.7514:  pointsaktr1(413) = points2(1) + 59.2488
pointsaktr1(414) = points2(4) - 50.9641:  pointsaktr1(415) = points2(1) + 59.4835
pointsaktr1(416) = points2(4) - 52.2794: pointsaktr1(417) = points2(1) + 57.5849
pointsaktr1(418) = points2(4) - 52.5139: pointsaktr1(419) = points2(1) + 54.413
pointsaktr1(420) = points2(4) - 53.7975: pointsaktr1(421) = points2(1) + 59.6861
pointsaktr1(422) = points2(4) - 53.3042:  pointsaktr1(423) = points2(1) + 61.8737
pointsaktr1(424) = points2(4) - 53.366: pointsaktr1(425) = points2(1) + 62.2884
pointsaktr1(426) = points2(4) - 53.7765:  pointsaktr1(427) = points2(1) + 62.614
pointsaktr1(428) = points2(4) - 54.4812:  pointsaktr1(429) = points2(1) + 62.5956
pointsaktr1(430) = points2(4) - 55.3326:  pointsaktr1(431) = points2(1) + 62.0533
pointsaktr1(432) = points2(4) - 55.9772:  pointsaktr1(433) = points2(1) + 61.8627
pointsaktr1(434) = points2(4) - 56.5534: pointsaktr1(435) = points2(1) + 62.0122
pointsaktr1(436) = points2(4) - 57.4669: pointsaktr1(437) = points2(1) + 62.2419
pointsaktr1(438) = points2(4) - 61.3268: pointsaktr1(439) = points2(1) + 61.7068
pointsaktr1(440) = points2(4) - 61.5862: pointsaktr1(441) = points2(1) + 61.5862
pointsaktr1(442) = points2(4) - 61.7068: pointsaktr1(443) = points2(1) + 61.3268
pointsaktr1(444) = points2(4) - 62.2419: pointsaktr1(445) = points2(1) + 57.4669
pointsaktr1(446) = points2(4) - 62.0122: pointsaktr1(447) = points2(1) + 56.5534
pointsaktr1(448) = points2(4) - 61.8627: pointsaktr1(449) = points2(1) + 55.9772
pointsaktr1(450) = points2(4) - 62.0533: pointsaktr1(451) = points2(1) + 55.3326
pointsaktr1(452) = points2(4) - 62.5956: pointsaktr1(453) = points2(1) + 54.4812
pointsaktr1(454) = points2(4) - 62.614: pointsaktr1(455) = points2(1) + 53.7765
pointsaktr1(456) = points2(4) - 62.2884: pointsaktr1(457) = points2(1) + 53.366
pointsaktr1(458) = points2(4) - 61.8737: pointsaktr1(459) = points2(1) + 53.3042
pointsaktr1(460) = points2(4) - 59.6861: pointsaktr1(461) = points2(1) + 53.7975
pointsaktr1(462) = points2(4) - 54.413: pointsaktr1(463) = points2(1) + 52.5139
pointsaktr1(464) = points2(4) - 57.5849: pointsaktr1(465) = points2(1) + 52.2794
pointsaktr1(466) = points2(4) - 59.4835:  pointsaktr1(467) = points2(1) + 50.9641
pointsaktr1(468) = points2(4) - 59.2488:  pointsaktr1(469) = points2(1) + 48.7514
pointsaktr1(470) = points2(4) - 57.2677:  pointsaktr1(471) = points2(1) + 48.0411
pointsaktr1(472) = points2(4) - 55.7883:  pointsaktr1(473) = points2(1) + 48.3252
pointsaktr1(474) = points2(4) - 54.3447:  pointsaktr1(475) = points2(1) + 48.7943
pointsaktr1(476) = points2(4) - 51.7037:  pointsaktr1(477) = points2(1) + 49.0761
pointsaktr1(478) = points2(4) - 47.6162:  pointsaktr1(479) = points2(1) + 44.2001
pointsaktr1(480) = points2(4) - 49.9128:  pointsaktr1(481) = points2(1) + 36.6409
pointsaktr1(482) = points2(4) - 55.4873:  pointsaktr1(483) = points2(1) + 33.7236
pointsaktr1(484) = points2(4) - 59.0042:  pointsaktr1(485) = points2(1) + 34.633
pointsaktr1(486) = points2(4) - 61.2812:  pointsaktr1(487) = points2(1) + 38.1637
pointsaktr1(488) = points2(4) - 60.022:   pointsaktr1(489) = points2(1) + 40.2373
pointsaktr1(490) = points2(4) - 59.6991:  pointsaktr1(491) = points2(1) + 38.013
pointsaktr1(492) = points2(4) - 58.5907:  pointsaktr1(493) = points2(1) + 37.5362
pointsaktr1(494) = points2(4) - 56.7296:  pointsaktr1(495) = points2(1) + 39.5607
pointsaktr1(496) = points2(4) - 58.0126:  pointsaktr1(497) = points2(1) + 41.8673
pointsaktr1(498) = points2(4) - 59.9507:  pointsaktr1(499) = points2(1) + 41.9916
pointsaktr1(500) = points2(4) - 62.8818:  pointsaktr1(501) = points2(1) + 37.6743
pointsaktr1(502) = points2(4) - 60.5:     pointsaktr1(503) = points2(1) + 33

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsaktr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update
    
currentBulge = plineObj.GetBulge(2)

plineObj.SetBulge 0, 0.240087
plineObj.SetBulge 1, 0.307386
plineObj.SetBulge 2, 0.224633
plineObj.SetBulge 3, 0.286077
plineObj.SetBulge 4, 0.36144
plineObj.SetBulge 5, 0.276755
plineObj.SetBulge 6, 0.244948
plineObj.SetBulge 7, -0.314822
plineObj.SetBulge 8, -0.259895
plineObj.SetBulge 9, -0.118763
plineObj.SetBulge 10, -0.254683
plineObj.SetBulge 11, -0.148555
plineObj.SetBulge 12, -0.363708
plineObj.SetBulge 13, -0.141808
plineObj.SetBulge 14, 0.0369261
plineObj.SetBulge 15, 0.0253271
plineObj.SetBulge 16, 0.246491
plineObj.SetBulge 17, 0.33
plineObj.SetBulge 18, 0.220023
plineObj.SetBulge 19, 0.0495029
plineObj.SetBulge 20, -0.130373
plineObj.SetBulge 21, -0.100964
plineObj.SetBulge 22, 0.29339
plineObj.SetBulge 23, 0.0911692
plineObj.SetBulge 24, 0.263146
plineObj.SetBulge 25, 0.0132959
plineObj.SetBulge 26, -0.154396
plineObj.SetBulge 27, -0.117989
plineObj.SetBulge 28, 0.121785
plineObj.SetBulge 29, 0.0709732
plineObj.SetBulge 30, 0.0779699
plineObj.SetBulge 31, 0.0779699
plineObj.SetBulge 32, 0.0709732
plineObj.SetBulge 33, 0.121785
plineObj.SetBulge 34, -0.117989
plineObj.SetBulge 35, -0.154396
plineObj.SetBulge 36, 0.0132959
plineObj.SetBulge 37, 0.263146
plineObj.SetBulge 38, 0.0911692
plineObj.SetBulge 39, 0.29339
plineObj.SetBulge 40, -0.100964
plineObj.SetBulge 41, -0.130373
plineObj.SetBulge 42, 0.0495029
plineObj.SetBulge 43, 0.220023
plineObj.SetBulge 44, 0.33
plineObj.SetBulge 45, 0.246491
plineObj.SetBulge 46, 0.0253271
plineObj.SetBulge 47, 0.0369261
plineObj.SetBulge 48, -0.141808
plineObj.SetBulge 49, -0.363708
plineObj.SetBulge 50, -0.148555
plineObj.SetBulge 51, -0.254683
plineObj.SetBulge 52, -0.118763
plineObj.SetBulge 53, -0.259895
plineObj.SetBulge 54, -0.314822
plineObj.SetBulge 55, 0.244948
plineObj.SetBulge 56, 0.276755
plineObj.SetBulge 57, 0.36144
plineObj.SetBulge 58, 0.286077
plineObj.SetBulge 59, 0.224633
plineObj.SetBulge 60, 0.307386
plineObj.SetBulge 61, 0.240087
plineObj.SetBulge 62, 0
plineObj.SetBulge 63, 0.240087
plineObj.SetBulge 64, 0.307386
plineObj.SetBulge 65, 0.224633
plineObj.SetBulge 66, 0.286077
plineObj.SetBulge 67, 0.36144
plineObj.SetBulge 68, 0.276755
plineObj.SetBulge 69, 0.244948
plineObj.SetBulge 70, -0.314822
plineObj.SetBulge 71, -0.259895
plineObj.SetBulge 72, -0.118763
plineObj.SetBulge 73, -0.254683
plineObj.SetBulge 74, -0.148555
plineObj.SetBulge 75, -0.363708
plineObj.SetBulge 76, -0.141808
plineObj.SetBulge 77, 0.0369261
plineObj.SetBulge 78, 0.0253271
plineObj.SetBulge 79, 0.246491
plineObj.SetBulge 80, 0.33
plineObj.SetBulge 81, 0.220023
plineObj.SetBulge 82, 0.0495029
plineObj.SetBulge 83, -0.130373
plineObj.SetBulge 84, -0.100964
plineObj.SetBulge 85, 0.29339
plineObj.SetBulge 86, 0.0911692
plineObj.SetBulge 87, 0.263146
plineObj.SetBulge 88, 0.0132959
plineObj.SetBulge 89, -0.154396
plineObj.SetBulge 90, -0.117989
plineObj.SetBulge 91, 0.121785
plineObj.SetBulge 92, 0.0709732
plineObj.SetBulge 93, 0.0779699
plineObj.SetBulge 94, 0.0779699
plineObj.SetBulge 95, 0.0709732
plineObj.SetBulge 96, 0.121785
plineObj.SetBulge 97, -0.117989
plineObj.SetBulge 98, -0.154396
plineObj.SetBulge 99, 0.0132959
plineObj.SetBulge 100, 0.263146
plineObj.SetBulge 101, 0.0911692
plineObj.SetBulge 102, 0.29339
plineObj.SetBulge 103, -0.100964
plineObj.SetBulge 104, -0.130373
plineObj.SetBulge 105, 0.0495029
plineObj.SetBulge 106, 0.220023
plineObj.SetBulge 107, 0.33
plineObj.SetBulge 108, 0.246491
plineObj.SetBulge 109, 0.0253271
plineObj.SetBulge 110, 0.0369261
plineObj.SetBulge 111, -0.141808
plineObj.SetBulge 112, -0.363708
plineObj.SetBulge 113, -0.148555
plineObj.SetBulge 114, -0.254683
plineObj.SetBulge 115, -0.118763
plineObj.SetBulge 116, -0.259895
plineObj.SetBulge 117, -0.314822
plineObj.SetBulge 118, 0.244948
plineObj.SetBulge 119, 0.276755
plineObj.SetBulge 120, 0.36144
plineObj.SetBulge 121, 0.286077
plineObj.SetBulge 122, 0.224633
plineObj.SetBulge 123, 0.307386
plineObj.SetBulge 124, 0.240087
plineObj.SetBulge 125, 0
plineObj.SetBulge 126, 0.240087
plineObj.SetBulge 127, 0.307386
plineObj.SetBulge 128, 0.224633
plineObj.SetBulge 129, 0.286077
plineObj.SetBulge 130, 0.36144
plineObj.SetBulge 131, 0.276755
plineObj.SetBulge 132, 0.244948
plineObj.SetBulge 133, -0.314822
plineObj.SetBulge 134, -0.259895
plineObj.SetBulge 135, -0.118763
plineObj.SetBulge 136, -0.254683
plineObj.SetBulge 137, -0.148555
plineObj.SetBulge 138, -0.363708
plineObj.SetBulge 139, -0.141808
plineObj.SetBulge 140, 0.0369261
plineObj.SetBulge 141, 0.0253271
plineObj.SetBulge 142, 0.246491
plineObj.SetBulge 143, 0.33
plineObj.SetBulge 144, 0.220023
plineObj.SetBulge 145, 0.0495029
plineObj.SetBulge 146, -0.130373
plineObj.SetBulge 147, -0.100964
plineObj.SetBulge 148, 0.29339
plineObj.SetBulge 149, 0.0911692
plineObj.SetBulge 150, 0.263146
plineObj.SetBulge 151, 0.0132959
plineObj.SetBulge 152, -0.154396
plineObj.SetBulge 153, -0.117989
plineObj.SetBulge 154, 0.121785
plineObj.SetBulge 155, 0.0709732
plineObj.SetBulge 156, 0.0779699
plineObj.SetBulge 157, 0.0779699
plineObj.SetBulge 158, 0.0709732
plineObj.SetBulge 159, 0.121785
plineObj.SetBulge 160, -0.117989
plineObj.SetBulge 161, -0.154396
plineObj.SetBulge 162, 0.0132959
plineObj.SetBulge 163, 0.263146
plineObj.SetBulge 164, 0.0911692
plineObj.SetBulge 165, 0.29339
plineObj.SetBulge 166, -0.100964
plineObj.SetBulge 167, -0.130373
plineObj.SetBulge 168, 0.0495029
plineObj.SetBulge 169, 0.220023
plineObj.SetBulge 170, 0.33
plineObj.SetBulge 171, 0.246491
plineObj.SetBulge 172, 0.0253271
plineObj.SetBulge 173, 0.0369261
plineObj.SetBulge 174, -0.141808
plineObj.SetBulge 175, -0.363708
plineObj.SetBulge 176, -0.148555
plineObj.SetBulge 177, -0.254683
plineObj.SetBulge 178, -0.118763
plineObj.SetBulge 179, -0.259895
plineObj.SetBulge 180, -0.314822
plineObj.SetBulge 181, 0.244948
plineObj.SetBulge 182, 0.276755
plineObj.SetBulge 183, 0.36144
plineObj.SetBulge 184, 0.286077
plineObj.SetBulge 185, 0.224633
plineObj.SetBulge 186, 0.307386
plineObj.SetBulge 187, 0.240087
plineObj.SetBulge 188, 0
plineObj.SetBulge 189, 0.240087
plineObj.SetBulge 190, 0.307386
plineObj.SetBulge 191, 0.224633
plineObj.SetBulge 192, 0.286077
plineObj.SetBulge 193, 0.36144
plineObj.SetBulge 194, 0.276755
plineObj.SetBulge 195, 0.244948
plineObj.SetBulge 196, -0.314822
plineObj.SetBulge 197, -0.259895
plineObj.SetBulge 198, -0.118763
plineObj.SetBulge 199, -0.254683
plineObj.SetBulge 200, -0.148555
plineObj.SetBulge 201, -0.363708
plineObj.SetBulge 202, -0.141808
plineObj.SetBulge 203, 0.0369261
plineObj.SetBulge 204, 0.0253271
plineObj.SetBulge 205, 0.246491
plineObj.SetBulge 206, 0.33
plineObj.SetBulge 207, 0.220023
plineObj.SetBulge 208, 0.0495029
plineObj.SetBulge 209, -0.130373
plineObj.SetBulge 210, -0.100964
plineObj.SetBulge 211, 0.29339
plineObj.SetBulge 212, 0.0911692
plineObj.SetBulge 213, 0.263146
plineObj.SetBulge 214, 0.0132959
plineObj.SetBulge 215, -0.154396
plineObj.SetBulge 216, -0.117989
plineObj.SetBulge 217, 0.121785
plineObj.SetBulge 218, 0.0709732
plineObj.SetBulge 219, 0.0779699
plineObj.SetBulge 220, 0.0779699
plineObj.SetBulge 221, 0.0709732
plineObj.SetBulge 222, 0.121785
plineObj.SetBulge 223, -0.117989
plineObj.SetBulge 224, -0.154396
plineObj.SetBulge 225, 0.0132959
plineObj.SetBulge 226, 0.263146
plineObj.SetBulge 227, 0.0911692
plineObj.SetBulge 228, 0.29339
plineObj.SetBulge 229, -0.100964
plineObj.SetBulge 230, -0.130373
plineObj.SetBulge 231, 0.0495029
plineObj.SetBulge 232, 0.220023
plineObj.SetBulge 233, 0.33
plineObj.SetBulge 234, 0.246491
plineObj.SetBulge 235, 0.0253271
plineObj.SetBulge 236, 0.0369261
plineObj.SetBulge 237, -0.141808
plineObj.SetBulge 238, -0.363708
plineObj.SetBulge 239, -0.148555
plineObj.SetBulge 240, -0.254683
plineObj.SetBulge 241, -0.118763
plineObj.SetBulge 242, -0.259895
plineObj.SetBulge 243, -0.314822
plineObj.SetBulge 244, 0.244948
plineObj.SetBulge 245, 0.276755
plineObj.SetBulge 246, 0.36144
plineObj.SetBulge 247, 0.286077
plineObj.SetBulge 248, 0.224633
plineObj.SetBulge 249, 0.307386
plineObj.SetBulge 250, 0.240087
plineObj.SetBulge 251, 0

    plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True

pointsaktr2(0) = points2(0) + 56.1843:   pointsaktr2(1) = points2(1) + 30
pointsaktr2(2) = points2(0) + 45.3574:   pointsaktr2(3) = points2(1) + 41.645
pointsaktr2(4) = points2(0) + 46.1831:   pointsaktr2(5) = points2(1) + 46.1831
pointsaktr2(6) = points2(0) + 41.645:    pointsaktr2(7) = points2(1) + 45.3574
pointsaktr2(8) = points2(0) + 30:        pointsaktr2(9) = points2(1) + 56.1843
pointsaktr2(10) = points2(0) + 30:       pointsaktr2(11) = points2(3) - 56.1843
pointsaktr2(12) = points2(0) + 41.645:   pointsaktr2(13) = points2(3) - 45.3574
pointsaktr2(14) = points2(0) + 46.1831:  pointsaktr2(15) = points2(3) - 46.1831
pointsaktr2(16) = points2(0) + 45.3574:  pointsaktr2(17) = points2(3) - 41.645
pointsaktr2(18) = points2(0) + 56.1843:  pointsaktr2(19) = points2(3) - 30
pointsaktr2(20) = points2(4) - 56.1843:  pointsaktr2(21) = points2(3) - 30
pointsaktr2(22) = points2(4) - 45.3574:  pointsaktr2(23) = points2(3) - 41.645
pointsaktr2(24) = points2(4) - 46.1831:  pointsaktr2(25) = points2(3) - 46.1831
pointsaktr2(26) = points2(4) - 41.645:   pointsaktr2(27) = points2(3) - 45.3574
pointsaktr2(28) = points2(4) - 30:       pointsaktr2(29) = points2(3) - 56.1843
pointsaktr2(30) = points2(4) - 30:       pointsaktr2(31) = points2(1) + 56.1843
pointsaktr2(32) = points2(4) - 41.645:   pointsaktr2(33) = points2(1) + 45.3574
pointsaktr2(34) = points2(4) - 46.1831:  pointsaktr2(35) = points2(1) + 46.1831
pointsaktr2(36) = points2(4) - 45.3574:  pointsaktr2(37) = points2(1) + 41.645
pointsaktr2(38) = points2(4) - 56.1843:  pointsaktr2(39) = points2(1) + 30


Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsaktr2)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update
  
plineObj.SetBulge 0, -0.435695
plineObj.SetBulge 1, -0.0902271
plineObj.SetBulge 2, -0.0902271
plineObj.SetBulge 3, -0.435695
plineObj.SetBulge 4, 0
plineObj.SetBulge 5, -0.435695
plineObj.SetBulge 6, -0.0902271
plineObj.SetBulge 7, -0.0902271
plineObj.SetBulge 8, -0.435695
plineObj.SetBulge 9, 0
plineObj.SetBulge 10, -0.435695
plineObj.SetBulge 11, -0.0902271
plineObj.SetBulge 12, -0.0902271
plineObj.SetBulge 13, -0.435695
plineObj.SetBulge 14, 0
plineObj.SetBulge 15, -0.435695
plineObj.SetBulge 16, -0.0902271
plineObj.SetBulge 17, -0.0902271
plineObj.SetBulge 18, -0.435695
plineObj.SetBulge 19, 0

 plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True

'=============================CENTER===========================================
  b1(0) = (points2(4) - (b / 2)) - 1: b1(1) = points2(3) - (a / 2)
  b2(0) = (points2(4) - (b / 2)) + 1: b2(1) = points2(3) - (a / 2)
 
  a1(0) = points2(4) - (b / 2): a1(1) = points2(1) + (a / 2) - 1
  a2(0) = points2(4) - (b / 2): a2(1) = points2(1) + (a / 2) + 1

bcp = points2(0) + (b / 2)
acp = points2(1) + (a / 2)

If a >= 320 Then
If b >= 170 Then

pointscntr1(0) = bcp + 24.5833:     pointscntr1(1) = acp + 20.0873
pointscntr1(2) = bcp + 28.0246:     pointscntr1(3) = acp + 19.9972
pointscntr1(4) = bcp + 19.3482:     pointscntr1(5) = acp + 25.688
pointscntr1(6) = bcp + 26.8683:     pointscntr1(7) = acp + 32.6531
pointscntr1(8) = bcp + 27.427:      pointscntr1(9) = acp + 33.6769
pointscntr1(10) = bcp + 27.9934:    pointscntr1(11) = acp + 34.4973
pointscntr1(12) = bcp + 31.4343:    pointscntr1(13) = acp + 35.7712
pointscntr1(14) = bcp + 33.8541:    pointscntr1(15) = acp + 35.1915
pointscntr1(16) = bcp + 37:         pointscntr1(17) = acp + 33.4752
pointscntr1(18) = bcp + 30.8536:    pointscntr1(19) = acp + 38.2705
pointscntr1(20) = bcp + 23.058:     pointscntr1(21) = acp + 38.2446
pointscntr1(22) = bcp + 13.2405:    pointscntr1(23) = acp + 32.8606
pointscntr1(24) = bcp + 6.64503:    pointscntr1(25) = acp + 23.8124
pointscntr1(26) = bcp + 6.64503:    pointscntr1(27) = acp - 23.8124
pointscntr1(28) = bcp + 13.2405:    pointscntr1(29) = acp - 32.8606
pointscntr1(30) = bcp + 23.058:     pointscntr1(31) = acp - 38.2446
pointscntr1(32) = bcp + 30.8536:    pointscntr1(33) = acp - 38.2705
pointscntr1(34) = bcp + 37:         pointscntr1(35) = acp - 33.4752
pointscntr1(36) = bcp + 33.8541:    pointscntr1(37) = acp - 35.1915
pointscntr1(38) = bcp + 31.4343:    pointscntr1(39) = acp - 35.7712
pointscntr1(40) = bcp + 27.9934:    pointscntr1(41) = acp - 34.4973
pointscntr1(42) = bcp + 27.427:     pointscntr1(43) = acp - 33.6769
pointscntr1(44) = bcp + 26.8683:    pointscntr1(45) = acp - 32.6531
pointscntr1(46) = bcp + 19.3482:    pointscntr1(47) = acp - 25.688
pointscntr1(48) = bcp + 28.0246:    pointscntr1(49) = acp - 19.9972
pointscntr1(50) = bcp + 24.5833:    pointscntr1(51) = acp - 20.0873
pointscntr1(52) = bcp + 19.6471:    pointscntr1(53) = acp - 16.403
pointscntr1(54) = bcp + 6.28296:    pointscntr1(55) = acp - 9.466
pointscntr1(56) = bcp + 6.28296:    pointscntr1(57) = acp + 9.466
pointscntr1(58) = bcp + 19.6471:    pointscntr1(59) = acp + 16.403

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update

plineObj.SetBulge 0, -0.190956
plineObj.SetBulge 1, 0.493502
plineObj.SetBulge 2, 0.173886
plineObj.SetBulge 3, 0
plineObj.SetBulge 4, -0.0792127
plineObj.SetBulge 5, -0.236352
plineObj.SetBulge 6, -0.0741078
plineObj.SetBulge 7, -0.0595151
plineObj.SetBulge 8, 0.188765
plineObj.SetBulge 9, 0.132767
plineObj.SetBulge 10, 0.129543
plineObj.SetBulge 11, 0.0905785
plineObj.SetBulge 12, 0.219599
plineObj.SetBulge 13, 0.0905785
plineObj.SetBulge 14, 0.129543
plineObj.SetBulge 15, 0.132767
plineObj.SetBulge 16, 0.188765
plineObj.SetBulge 17, -0.0595151
plineObj.SetBulge 18, -0.0741078
plineObj.SetBulge 19, -0.236352
plineObj.SetBulge 20, -0.0792127
plineObj.SetBulge 21, 0
plineObj.SetBulge 22, 0.173886
plineObj.SetBulge 23, 0.493502
plineObj.SetBulge 24, -0.190956
plineObj.SetBulge 25, -0.133032
plineObj.SetBulge 26, -0.20958
plineObj.SetBulge 27, -0.334034
plineObj.SetBulge 28, -0.20958
plineObj.SetBulge 29, -0.133032

plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True

RetVal = plineObj.Mirror(a1, a2)

pointscntr2(0) = bcp + 6.58238:      pointscntr2(1) = acp + 31.3788
pointscntr2(2) = bcp + 4.25321:      pointscntr2(3) = acp + 33.6762
pointscntr2(4) = bcp + 0.898097:     pointscntr2(5) = acp + 33.4752
pointscntr2(6) = bcp + 1.68328:      pointscntr2(7) = acp + 35.9463
pointscntr2(8) = bcp + 1.70714:      pointscntr2(9) = acp + 37.0063
pointscntr2(10) = bcp + 0.782447:    pointscntr2(11) = acp + 37.995
pointscntr2(12) = bcp + 1.13687E-13: pointscntr2(13) = acp + 37.9676
pointscntr2(14) = bcp + -0.782447:   pointscntr2(15) = acp + 37.995
pointscntr2(16) = bcp + -1.70714:    pointscntr2(17) = acp + 37.0063
pointscntr2(18) = bcp + -1.68328:    pointscntr2(19) = acp + 35.9463
pointscntr2(20) = bcp + -0.898097:   pointscntr2(21) = acp + 33.4752
pointscntr2(22) = bcp + -4.25321:    pointscntr2(23) = acp + 33.6762
pointscntr2(24) = bcp + -6.58238:    pointscntr2(25) = acp + 31.3788
pointscntr2(26) = bcp + -3.92174:    pointscntr2(27) = acp + 29.4618
pointscntr2(28) = bcp + -1.70295:    pointscntr2(29) = acp + 25.1
pointscntr2(30) = bcp + -0.455718:   pointscntr2(31) = acp + 18.477
pointscntr2(32) = bcp + -0.227606:   pointscntr2(33) = acp + 15.167
pointscntr2(34) = bcp + -0.110535:   pointscntr2(35) = acp + 11.4844
pointscntr2(36) = bcp + -0.0689475:  pointscntr2(37) = acp + 10.7799
pointscntr2(38) = bcp + -0.0388717:  pointscntr2(39) = acp + 10.5512
pointscntr2(40) = bcp + 1.13687E-13: pointscntr2(41) = acp + 10.4124
pointscntr2(42) = bcp + 0.0388717:   pointscntr2(43) = acp + 10.5512
pointscntr2(44) = bcp + 0.0689475:   pointscntr2(45) = acp + 10.7799
pointscntr2(46) = bcp + 0.110535:    pointscntr2(47) = acp + 11.4844
pointscntr2(48) = bcp + 0.227606:    pointscntr2(49) = acp + 15.167
pointscntr2(50) = bcp + 0.455718:    pointscntr2(51) = acp + 18.477
pointscntr2(52) = bcp + 1.70295:     pointscntr2(53) = acp + 25.1
pointscntr2(54) = bcp + 3.92174:     pointscntr2(55) = acp + 29.4618


Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr2)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update

plineObj.SetBulge 0, 0.130293
plineObj.SetBulge 1, 0.286262
plineObj.SetBulge 2, 0.0628827
plineObj.SetBulge 3, 0.0664767
plineObj.SetBulge 4, 0.354054
plineObj.SetBulge 5, 0.0749942
plineObj.SetBulge 6, 0.0749942
plineObj.SetBulge 7, 0.354054
plineObj.SetBulge 8, 0.0664767
plineObj.SetBulge 9, 0.0628827
plineObj.SetBulge 10, 0.286262
plineObj.SetBulge 11, 0.130293
plineObj.SetBulge 12, -0.132961
plineObj.SetBulge 13, -0.0949138
plineObj.SetBulge 14, -0.0510213
plineObj.SetBulge 15, -0.0125717
plineObj.SetBulge 16, -0.00334278
plineObj.SetBulge 17, 0
plineObj.SetBulge 18, 0
plineObj.SetBulge 19, 0
plineObj.SetBulge 20, 0
plineObj.SetBulge 21, 0
plineObj.SetBulge 22, 0
plineObj.SetBulge 23, -0.00334278
plineObj.SetBulge 24, -0.0125717
plineObj.SetBulge 25, -0.0510213
plineObj.SetBulge 26, -0.0949138
plineObj.SetBulge 27, -0.132961

plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True

RetVal = plineObj.Mirror(b1, b2)

pointscntr3(0) = bcp + 17.03:       pointscntr3(1) = acp + 80.0281
pointscntr3(2) = bcp + 12.3194:     pointscntr3(3) = acp + 76.2511
pointscntr3(4) = bcp + 12.7291:     pointscntr3(5) = acp + 73.3674
pointscntr3(6) = bcp + 16.4548:     pointscntr3(7) = acp + 74.8083
pointscntr3(8) = bcp + 22.4096:     pointscntr3(9) = acp + 72.6438
pointscntr3(10) = bcp + 24.1819:    pointscntr3(11) = acp + 69.1113
pointscntr3(12) = bcp + 24.8313:    pointscntr3(13) = acp + 64.3255
pointscntr3(14) = bcp + 23.9272:    pointscntr3(15) = acp + 59.0018
pointscntr3(16) = bcp + 16.5438:    pointscntr3(17) = acp + 46.828
pointscntr3(18) = bcp + 6.28228:    pointscntr3(19) = acp + 36.4703
pointscntr3(20) = bcp + 22.0575:    pointscntr3(21) = acp + 48.4905
pointscntr3(22) = bcp + 28.7564:    pointscntr3(23) = acp + 65.9369
pointscntr3(24) = bcp + 25.5548:    pointscntr3(25) = acp + 74.964

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr3)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update

plineObj.SetBulge 0, 0.424127
plineObj.SetBulge 1, 0.161605
plineObj.SetBulge 2, 0.83194
plineObj.SetBulge 3, -0.267392
plineObj.SetBulge 4, -0.106986
plineObj.SetBulge 5, -0.0601812
plineObj.SetBulge 6, -0.0862239
plineObj.SetBulge 7, -0.0871434
plineObj.SetBulge 8, -0.0408194
plineObj.SetBulge 9, 0.0867111
plineObj.SetBulge 10, 0.208082
plineObj.SetBulge 11, 0.154888
plineObj.SetBulge 12, 0.189184

plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True
RetVal = plineObj.Mirror(b1, b2)
RetVal = plineObj.Mirror(a1, a2)
RetVal = plineObj.Copy
  ' Define the rotation of 45 degrees about a
  ' base point of (4, 4.25, 0)
  basePoint(0) = points2(0) + (b / 2): basePoint(1) = points2(1) + (a / 2)
  rotationAngle = 3.14159   ' 45 degrees
  plineObj.Rotate basePoint, rotationAngle
  plineObj.Update
  
pointscntr4(0) = bcp + 6.97983:      pointscntr4(1) = acp + 66.4733
pointscntr4(2) = bcp + 4.40328:      pointscntr4(3) = acp + 70.4044
pointscntr4(4) = bcp + 2.76705:      pointscntr4(5) = acp + 73.302
pointscntr4(6) = bcp + 2.03069:      pointscntr4(7) = acp + 75.4075
pointscntr4(8) = bcp + 1.13687E-13:  pointscntr4(9) = acp + 82.896
pointscntr4(10) = bcp + -2.03069:    pointscntr4(11) = acp + 75.4075
pointscntr4(12) = bcp + -2.76705:    pointscntr4(13) = acp + 73.302
pointscntr4(14) = bcp + -4.35149:    pointscntr4(15) = acp + 70.4834
pointscntr4(16) = bcp + -6.97983:    pointscntr4(17) = acp + 66.4733
pointscntr4(18) = bcp + -8.62407:    pointscntr4(19) = acp + 63.7027
pointscntr4(20) = bcp + -9.8543:     pointscntr4(21) = acp + 57.1808
pointscntr4(22) = bcp + -8.14995:    pointscntr4(23) = acp + 53.3016
pointscntr4(24) = bcp + -5.61959:    pointscntr4(25) = acp + 51.6618
pointscntr4(26) = bcp + -2.27229:    pointscntr4(27) = acp + 51.8809
pointscntr4(28) = bcp + 1.13687E-13: pointscntr4(29) = acp + 55.0398
pointscntr4(30) = bcp + 2.27229:     pointscntr4(31) = acp + 51.8809
pointscntr4(32) = bcp + 5.61959:     pointscntr4(33) = acp + 51.6618
pointscntr4(34) = bcp + 8.14995:     pointscntr4(35) = acp + 53.3016
pointscntr4(36) = bcp + 9.8543:      pointscntr4(37) = acp + 57.1808
pointscntr4(38) = bcp + 8.62407:     pointscntr4(39) = acp + 63.7027

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr4)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update

plineObj.SetBulge 0, 0
plineObj.SetBulge 1, -0.0340478
plineObj.SetBulge 2, -0.0404214
plineObj.SetBulge 3, 0
plineObj.SetBulge 4, 0
plineObj.SetBulge 5, -0.0404214
plineObj.SetBulge 6, -0.0330812
plineObj.SetBulge 7, 0
plineObj.SetBulge 8, 0.021196
plineObj.SetBulge 9, 0.173455
plineObj.SetBulge 10, 0.145317
plineObj.SetBulge 11, 0.143574
plineObj.SetBulge 12, 0.203485
plineObj.SetBulge 13, 0.200979
plineObj.SetBulge 14, 0.200979
plineObj.SetBulge 15, 0.203485
plineObj.SetBulge 16, 0.143574
plineObj.SetBulge 17, 0.145317
plineObj.SetBulge 18, 0.173455
plineObj.SetBulge 19, 0.021196

plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True
RetVal = plineObj.Mirror(b1, b2)


pointscntr5(0) = bcp + 8.66267:     pointscntr5(1) = acp + 94.5619
pointscntr5(2) = bcp + 12.0028:     pointscntr5(3) = acp + 94.7543
pointscntr5(4) = bcp + 14.9353:     pointscntr5(5) = acp + 93.6436
pointscntr5(6) = bcp + 16.3813:     pointscntr5(7) = acp + 91.4456
pointscntr5(8) = bcp + 16.3602:     pointscntr5(9) = acp + 89.5013
pointscntr5(10) = bcp + 14.3893:    pointscntr5(11) = acp + 86.8541
pointscntr5(12) = bcp + 11.3686:    pointscntr5(13) = acp + 86.1899
pointscntr5(14) = bcp + 13.0885:    pointscntr5(15) = acp + 87.6125
pointscntr5(16) = bcp + 12.9123:    pointscntr5(17) = acp + 89.239
pointscntr5(18) = bcp + 10.7834:    pointscntr5(19) = acp + 90.4247
pointscntr5(20) = bcp + 5.27888:    pointscntr5(21) = acp + 89.7626
pointscntr5(22) = bcp + 2.39309:    pointscntr5(23) = acp + 86.7891
pointscntr5(24) = bcp + 1.49609:    pointscntr5(25) = acp + 88.2863
pointscntr5(26) = bcp + 4.32014:    pointscntr5(27) = acp + 92.3133

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr5)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update
  
plineObj.SetBulge 0, -0.0883845
plineObj.SetBulge 1, -0.121995
plineObj.SetBulge 2, -0.193108
plineObj.SetBulge 3, -0.0993998
plineObj.SetBulge 4, -0.23388
plineObj.SetBulge 5, -0.125673
plineObj.SetBulge 6, 0.11811
plineObj.SetBulge 7, 0.412476
plineObj.SetBulge 8, 0.0962103
plineObj.SetBulge 9, 0.224454
plineObj.SetBulge 10, 0.11374
plineObj.SetBulge 11, 0
plineObj.SetBulge 12, -0.11562
plineObj.SetBulge 13, -0.128851

plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True
RetVal = plineObj.Mirror(b1, b2)
RetVal = plineObj.Mirror(a1, a2)
RetVal = plineObj.Copy
  ' Define the rotation of 45 degrees about a
  ' base point of (4, 4.25, 0)
  basePoint(0) = points2(0) + (b / 2): basePoint(1) = points2(1) + (a / 2)
  rotationAngle = 3.14159   ' 45 degrees
  plineObj.Rotate basePoint, rotationAngle
  plineObj.Update

pointscntr6(0) = bcp + 3.75304:       pointscntr6(1) = acp + 101.938
pointscntr6(2) = bcp + 1.13687E-13:   pointscntr6(3) = acp + 111.5
pointscntr6(4) = bcp + -3.75304:      pointscntr6(5) = acp + 101.938
pointscntr6(6) = bcp + -3.37708:      pointscntr6(7) = acp + 96.8397
pointscntr6(8) = bcp + 1.13687E-13:   pointscntr6(9) = acp + 88.5857
pointscntr6(10) = bcp + 3.37708:      pointscntr6(11) = acp + 96.8397

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointscntr6)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update
  
plineObj.SetBulge 0, 0.06872
plineObj.SetBulge 1, 0.06872
plineObj.SetBulge 2, 0.171056
plineObj.SetBulge 3, 0
plineObj.SetBulge 4, 0
plineObj.SetBulge 5, 0.171056

    plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True
RetVal = plineObj.Mirror(b1, b2)

center(0) = a1(0): center(1) = b1(1) + 45.76: center(2) = 0: radius = 2.9613
Set circleObj = ThisDrawing.ModelSpace.AddCircle(center, radius)
circleObj.Layer = "K-grav_Pattern"
circleObj.Update
center(0) = a1(0): center(1) = b1(1) - 45.76: center(2) = 0: radius = 2.9613
Set circleObj = ThisDrawing.ModelSpace.AddCircle(center, radius)
circleObj.Layer = "K-grav_Pattern"
circleObj.Update

End If
End If
End If
End If
  
  points2(1) = points2(3) + I:
  points2(3) = points2(3) + a + I:
  points2(5) = points2(3):
  points2(7) = points2(1):
  

Else: d = 1
d = d + 1

d = Cells(c, 4)
End If

Next d
c = c + 1

a = Cells(c, 2)
b = Cells(c, 3)

I = 100

Loop
              
ZoomAll
AP.Quit
End Sub

Sub MF145()
' Creating search objects
Dim AP As Excel.Application
Dim WB As Excel.Workbook
Dim WS As Excel.Worksheet

Set AP = Excel.Application
Set WB = AP.Workbooks.Open("C:\Users\������\Desktop\������\2.xlsx")
Set WS = WB.Worksheets("������ �˨������ ������")
c = 19
a = Cells(c, 2)
b = Cells(c, 3)
d = Cells(c, 4)
E = 1
I = 0
On Error Resume Next

' Description of all the points that will make up rectangles
  Dim plineObj As AcadLWPolyline
  Dim points(0 To 7) As Double
  Dim pointsk(0 To 7) As Double
  Dim pointsark(0 To 3) As Double
  Dim pointsktr1(0 To 95) As Double
  Dim pointsktr2(0 To 51) As Double
  Dim pointsktr3(0 To 15) As Double
  Dim pointsktr4(0 To 5) As Double
  Dim pointsktr5(0 To 5) As Double
  Dim M1(0 To 2) As Double
  Dim M2(0 To 2) As Double
  
points(6) = 0

Do While a > 0

' Calculate the coordinates of each point
  points(0) = points(6) + I:    points(1) = 0
  points(2) = points(6) + I:    points(3) = a
  points(4) = points(0) + b:    points(5) = a
  points(6) = points(0) + b:    points(7) = 0

' Constructing four segments and connecting the corresponding points

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
  plineObj.Closed = True
 
Dim a1(0 To 2) As Double
Dim a2(0 To 2) As Double
Dim A3(0 To 2) As Double
Dim A4(0 To 2) As Double
Dim A5(0 To 2) As Double
Dim A6(0 To 2) As Double
Dim A7(0 To 2) As Double
Dim A8(0 To 2) As Double
Dim lineObj As AcadLine
  
  a1(0) = points(0) + 47: a1(1) = 0:      a1(2) = 0
  a2(0) = points(2) + 47: a2(1) = a:      a2(2) = 0
  A3(0) = points(4) - 47: A3(1) = a:      A3(2) = 0
  A4(0) = points(6) - 47: A4(1) = 0:      A4(2) = 0
  
  A5(0) = points(0) + 47: A5(1) = 77:      A5(2) = 0
  A6(0) = points(2) + 47: A6(1) = a - 77:  A6(2) = 0
  A7(0) = points(4) - 47: A7(1) = a - 77:  A7(2) = 0
  A8(0) = points(6) - 47: A8(1) = 77:      A8(2) = 0
   
If a >= 105 Then
If b >= 105 Then
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(a1, a2)
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A3, A4)
lineObj.Layer = "Ball-6"
lineObj.Update

If a >= 180 Then

Set lineObj = ThisDrawing.ModelSpace.AddLine(A6, A7)
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A8, A5)
lineObj.Layer = "Ball-6"
lineObj.Update

End If

  pointsark(0) = points(0) + 47:    pointsark(1) = points(1) + 45
  pointsark(2) = points(4) - 47:    pointsark(3) = points(1) + 45

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsark)
k = (pointsark(2) - pointsark(0)) / 2
    ' Change the bulge of the third segment
    plineObj.SetBulge 0, 21 / k
    plineObj.Update
plineObj.Layer = "Ball-6"
plineObj.Update


  pointsark(0) = points(0) + 47:    pointsark(1) = points(3) - 45
  pointsark(2) = points(4) - 47:    pointsark(3) = points(3) - 45

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsark)
k = (pointsark(2) - pointsark(0)) / 2
    ' Change the bulge of the third segment
    plineObj.SetBulge 0, -(21 / k)
    plineObj.Update
plineObj.Layer = "Ball-6"
plineObj.Update
 
                                 'plineObj.color = acByLayer
                                 'Dim layerObj As AcadLayer
                                 'Set layerObj = ThisDrawing.Layers.Add("Ball-6.1")
                                 'layerObj.color = acRed
If a >= 254 Then
If b >= 194 Then


  pointsk(0) = points(0) + 47:    pointsk(1) = points(1) + 77
  pointsk(2) = points(0) + 47:    pointsk(3) = points(3) - 77
  pointsk(4) = points(4) - 47:    pointsk(5) = points(3) - 77
  pointsk(6) = points(4) - 47:    pointsk(7) = points(1) + 77
  
  
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsk)
plineObj.Closed = True

plineObj.Layer = "K-grav"
plineObj.Update

 ' Offset the polyline
  Dim offsetObj As Variant
  
  plineObj.Layer = "C-mill"
plineObj.Update
  offsetObj = plineObj.Offset(20)
plineObj.Layer = "K-grav"
plineObj.Update

End If
End If

'============================ENGRAVING============================

If a >= 146 Then
If b >= 295 Then

bcp = points(0) + (b / 2)
acp = points(1)
M1(0) = points(4) - (b / 2) - 1: M1(1) = points(1) + (a / 2)
M2(0) = points(4) - (b / 2) + 1: M2(1) = points(1) + (a / 2)

pointsktr1(0) = bcp + 5.68434E-14:    pointsktr1(1) = acp + 60.7166
pointsktr1(2) = bcp + 9.63955:        pointsktr1(3) = acp + 59.23
pointsktr1(4) = bcp + 18.3303:        pointsktr1(5) = acp + 54.1882
pointsktr1(6) = bcp + 33.5681:        pointsktr1(7) = acp + 60.6556
pointsktr1(8) = bcp + 50.0241:        pointsktr1(9) = acp + 59.1337
pointsktr1(10) = bcp + 76.6522:       pointsktr1(11) = acp + 48.2634
pointsktr1(12) = bcp + 95:            pointsktr1(13) = acp + 50.8008
pointsktr1(14) = bcp + 94.1946:       pointsktr1(15) = acp + 49.3519
pointsktr1(16) = bcp + 89.9012:       pointsktr1(17) = acp + 46.7215
pointsktr1(18) = bcp + 84.9367:       pointsktr1(19) = acp + 45.8814
pointsktr1(20) = bcp + 73.728:        pointsktr1(21) = acp + 47.1349
pointsktr1(22) = bcp + 60.6079:       pointsktr1(23) = acp + 52.0547
pointsktr1(24) = bcp + 47.6654:       pointsktr1(25) = acp + 57.4244
pointsktr1(26) = bcp + 36.6652:       pointsktr1(27) = acp + 59.5536
pointsktr1(28) = bcp + 30.7517:       pointsktr1(29) = acp + 59.3172
pointsktr1(30) = bcp + 25.4212:       pointsktr1(31) = acp + 57.9987
pointsktr1(32) = bcp + 19.3968:       pointsktr1(33) = acp + 53.1661
pointsktr1(34) = bcp + 20.1065:       pointsktr1(35) = acp + 43.4681
pointsktr1(36) = bcp + 25.1381:       pointsktr1(37) = acp + 41.104
pointsktr1(38) = bcp + 30.2788:       pointsktr1(39) = acp + 43.2313
pointsktr1(40) = bcp + 28.3471:       pointsktr1(41) = acp + 40.8538
pointsktr1(42) = bcp + 25.3618:       pointsktr1(43) = acp + 40.1632
pointsktr1(44) = bcp + 19.3968:       pointsktr1(45) = acp + 42.2855
pointsktr1(46) = bcp + 9.94206:       pointsktr1(47) = acp + 38.2695
pointsktr1(48) = bcp + 5.68434E-14:   pointsktr1(49) = acp + 40.8656
pointsktr1(50) = bcp - 9.94206:       pointsktr1(51) = acp + 38.2695
pointsktr1(52) = bcp - 19.3968:       pointsktr1(53) = acp + 42.2855
pointsktr1(54) = bcp - 25.3618:       pointsktr1(55) = acp + 40.1632
pointsktr1(56) = bcp - 28.3471:       pointsktr1(57) = acp + 40.8538
pointsktr1(58) = bcp - 30.2788:       pointsktr1(59) = acp + 43.2313
pointsktr1(60) = bcp - 25.1381:       pointsktr1(61) = acp + 41.104
pointsktr1(62) = bcp - 20.1065:       pointsktr1(63) = acp + 43.4681
pointsktr1(64) = bcp - 19.3968:       pointsktr1(65) = acp + 53.1661
pointsktr1(66) = bcp - 25.4212:       pointsktr1(67) = acp + 57.9987
pointsktr1(68) = bcp - 30.7517:       pointsktr1(69) = acp + 59.3172
pointsktr1(70) = bcp - 36.6652:       pointsktr1(71) = acp + 59.5536
pointsktr1(72) = bcp - 47.6654:       pointsktr1(73) = acp + 57.4244
pointsktr1(74) = bcp - 60.6079:       pointsktr1(75) = acp + 52.0547
pointsktr1(76) = bcp - 73.728:        pointsktr1(77) = acp + 47.1349
pointsktr1(78) = bcp - 84.9367:       pointsktr1(79) = acp + 45.8814
pointsktr1(80) = bcp - 89.9012:       pointsktr1(81) = acp + 46.7215
pointsktr1(82) = bcp - 94.1946:       pointsktr1(83) = acp + 49.3519
pointsktr1(84) = bcp - 95:            pointsktr1(85) = acp + 50.8008
pointsktr1(86) = bcp - 76.6522:       pointsktr1(87) = acp + 48.2634
pointsktr1(88) = bcp - 50.0241:       pointsktr1(89) = acp + 59.1337
pointsktr1(90) = bcp - 33.5681:       pointsktr1(91) = acp + 60.6556
pointsktr1(92) = bcp - 18.3303:       pointsktr1(93) = acp + 54.1882
pointsktr1(94) = bcp - 9.63955:       pointsktr1(95) = acp + 59.23

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsktr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update
  
plineObj.SetBulge 0, -0.0694346
plineObj.SetBulge 1, -0.143464
plineObj.SetBulge 2, -0.205036
plineObj.SetBulge 3, -0.0755612
plineObj.SetBulge 4, -0.0499684
plineObj.SetBulge 5, 0.307878
plineObj.SetBulge 6, -0.174812
plineObj.SetBulge 7, -0.109731
plineObj.SetBulge 8, -0.0671936
plineObj.SetBulge 9, -0.080044
plineObj.SetBulge 10, -0.0258583
plineObj.SetBulge 11, 0.0120843
plineObj.SetBulge 12, 0.0920047
plineObj.SetBulge 13, 0.0206841
plineObj.SetBulge 14, 0.075674
plineObj.SetBulge 15, 0.163488
plineObj.SetBulge 16, -0.309365
plineObj.SetBulge 17, 0.242665
plineObj.SetBulge 18, 0.171279
plineObj.SetBulge 19, -0.21038
plineObj.SetBulge 20, -0.118469
plineObj.SetBulge 21, -0.187197
plineObj.SetBulge 22, -0.203467
plineObj.SetBulge 23, -0.123761
plineObj.SetBulge 24, -0.123761
plineObj.SetBulge 25, -0.203467
plineObj.SetBulge 26, -0.187197
plineObj.SetBulge 27, -0.118469
plineObj.SetBulge 28, -0.21038
plineObj.SetBulge 29, 0.171279
plineObj.SetBulge 30, 0.242665
plineObj.SetBulge 31, -0.309365
plineObj.SetBulge 32, 0.163488
plineObj.SetBulge 33, 0.075674
plineObj.SetBulge 34, 0.0206841
plineObj.SetBulge 35, 0.0920047
plineObj.SetBulge 36, 0.0120843
plineObj.SetBulge 37, -0.0258583
plineObj.SetBulge 38, -0.080044
plineObj.SetBulge 39, -0.0671936
plineObj.SetBulge 40, -0.109731
plineObj.SetBulge 41, -0.174812
plineObj.SetBulge 42, 0.307878
plineObj.SetBulge 43, -0.0499684
plineObj.SetBulge 44, -0.0755612
plineObj.SetBulge 45, -0.205036
plineObj.SetBulge 46, -0.143464
plineObj.SetBulge 47, -0.0694346
   
  plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True
    RetVal = plineObj.Mirror(M1, M2)

pointsktr2(0) = bcp + 5.68434E-14:    pointsktr2(1) = acp + 52.6932
pointsktr2(2) = bcp + 2.50477:        pointsktr2(3) = acp + 51.9151
pointsktr2(4) = bcp + 3.90531:        pointsktr2(5) = acp + 50.1933
pointsktr2(6) = bcp + 4.4078:         pointsktr2(7) = acp + 48.158
pointsktr2(8) = bcp + 4.1626:         pointsktr2(9) = acp + 44.8378
pointsktr2(10) = bcp + 2.1293:        pointsktr2(11) = acp + 42.2855
pointsktr2(12) = bcp + 6.07129:       pointsktr2(13) = acp + 41.0498
pointsktr2(14) = bcp + 10.2004:       pointsktr2(15) = acp + 40.9088
pointsktr2(16) = bcp + 14.3105:       pointsktr2(17) = acp + 41.5164
pointsktr2(18) = bcp + 18.2265:       pointsktr2(19) = acp + 43.6516
pointsktr2(20) = bcp + 16.73:         pointsktr2(21) = acp + 48.2914
pointsktr2(22) = bcp + 17.7417:       pointsktr2(23) = acp + 53.1661
pointsktr2(24) = bcp + 9.51482:       pointsktr2(25) = acp + 58.0991
pointsktr2(26) = bcp + 5.68434E-14:   pointsktr2(27) = acp + 59.3172
pointsktr2(28) = bcp - 9.51482:       pointsktr2(29) = acp + 58.0991
pointsktr2(30) = bcp - 17.7417:       pointsktr2(31) = acp + 53.1661
pointsktr2(32) = bcp - 16.73:         pointsktr2(33) = acp + 48.2914
pointsktr2(34) = bcp - 18.2265:       pointsktr2(35) = acp + 43.6516
pointsktr2(36) = bcp - 14.3105:       pointsktr2(37) = acp + 41.5164
pointsktr2(38) = bcp - 10.2004:       pointsktr2(39) = acp + 40.9088
pointsktr2(40) = bcp - 6.07129:       pointsktr2(41) = acp + 41.0498
pointsktr2(42) = bcp - 2.1293:        pointsktr2(43) = acp + 42.2855
pointsktr2(44) = bcp - 4.1626:        pointsktr2(45) = acp + 44.8378
pointsktr2(46) = bcp - 4.4078:        pointsktr2(47) = acp + 48.158
pointsktr2(48) = bcp - 3.90531:       pointsktr2(49) = acp + 50.1933
pointsktr2(50) = bcp - 2.50477:       pointsktr2(51) = acp + 51.9151

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsktr2)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update
  
plineObj.SetBulge 0, -0.164816
plineObj.SetBulge 1, -0.123711
plineObj.SetBulge 2, -0.0740984
plineObj.SetBulge 3, -0.123082
plineObj.SetBulge 4, -0.129395
plineObj.SetBulge 5, 0.0762057
plineObj.SetBulge 6, 0.0503136
plineObj.SetBulge 7, 0.0523918
plineObj.SetBulge 8, 0.128222
plineObj.SetBulge 9, -0.181721
plineObj.SetBulge 10, -0.115229
plineObj.SetBulge 11, 0.147569
plineObj.SetBulge 12, 0.0693834
plineObj.SetBulge 13, 0.0693834
plineObj.SetBulge 14, 0.147569
plineObj.SetBulge 15, -0.115229
plineObj.SetBulge 16, -0.181721
plineObj.SetBulge 17, 0.128222
plineObj.SetBulge 18, 0.0523918
plineObj.SetBulge 19, 0.0503136
plineObj.SetBulge 20, 0.0762057
plineObj.SetBulge 21, -0.129395
plineObj.SetBulge 22, -0.123082
plineObj.SetBulge 23, -0.0740984
plineObj.SetBulge 24, -0.123711
plineObj.SetBulge 25, -0.164816

plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True
    RetVal = plineObj.Mirror(M1, M2)
    
pointsktr3(0) = bcp + 5.68434E-14:  pointsktr3(1) = acp + 50.5643
pointsktr3(2) = bcp + 1.65434:      pointsktr3(3) = acp + 49.1014
pointsktr3(4) = bcp + 1.88941:      pointsktr3(5) = acp + 46.8978
pointsktr3(6) = bcp + 1.30833:      pointsktr3(7) = acp + 44.8785
pointsktr3(8) = bcp + 5.68434E-14:  pointsktr3(9) = acp + 43.2314
pointsktr3(10) = bcp - 1.30833:     pointsktr3(11) = acp + 44.8785
pointsktr3(12) = bcp - 1.88941:     pointsktr3(13) = acp + 46.8978
pointsktr3(14) = bcp - 1.65434:     pointsktr3(15) = acp + 49.1014

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsktr3)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update
  
plineObj.SetBulge 0, -0.294708
plineObj.SetBulge 1, -0.0635714
plineObj.SetBulge 2, -0.14288
plineObj.SetBulge 3, -0.061403
plineObj.SetBulge 4, -0.061403
plineObj.SetBulge 5, -0.14288
plineObj.SetBulge 6, -0.0635714
plineObj.SetBulge 7, -0.294708

plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True
    RetVal = plineObj.Mirror(M1, M2)
   
pointsktr4(0) = bcp - 19.1607: pointsktr4(1) = acp + 44.6506
pointsktr4(2) = bcp - 18.6875: pointsktr4(3) = acp + 51.7474
pointsktr4(4) = bcp - 19.1607: pointsktr4(5) = acp + 44.6506
   
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsktr4)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update
  
plineObj.SetBulge 0, -0.261219
plineObj.SetBulge 1, -0.267301
plineObj.SetBulge 2, 0

plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True
    RetVal = plineObj.Mirror(M1, M2)

pointsktr5(0) = bcp + 19.1607: pointsktr5(1) = acp + 44.6506
pointsktr5(2) = bcp + 18.6875: pointsktr5(3) = acp + 51.7474
pointsktr5(4) = bcp + 19.1607: pointsktr5(5) = acp + 44.6506
    
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsktr5)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update
  
plineObj.SetBulge 0, 0.261219
plineObj.SetBulge 1, 0.267301
plineObj.SetBulge 2, 0

plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True
    RetVal = plineObj.Mirror(M1, M2)



End If
End If
End If
End If

'=======================================================================================================================
'==========================REPEATING ELEMENTS=========================================================================

I = 100
  
 Dim plineObj2 As AcadLWPolyline
  Dim points2(0 To 7) As Double
  
  points2(0) = points(0):    points2(1) = a + I:
  points2(2) = points(0):    points2(3) = (a * 2) + I:
  points2(4) = points(6):    points2(5) = (a * 2) + I:
  points2(6) = points(6):    points2(7) = a + I:
  
 For d = 2 To Cells(c, 4)
If d > 1 Then

Set plineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(points2)
  plineObj2.Closed = True

  
  a1(0) = points2(0) + 47: a1(1) = points2(1):        a1(2) = 0
  a2(0) = points2(2) + 47: a2(1) = points2(1) + a:    a2(2) = 0
  A3(0) = points2(4) - 47: A3(1) = points2(1) + a:    A3(2) = 0
  A4(0) = points2(6) - 47: A4(1) = points2(1):        A4(2) = 0
  
  A5(0) = points2(0) + 47: A5(1) = points2(1) + 77:    A5(2) = 0
  A6(0) = points2(2) + 47: A6(1) = points2(3) - 77:    A6(2) = 0
  A7(0) = points2(4) - 47: A7(1) = points2(3) - 77:    A7(2) = 0
  A8(0) = points2(6) - 47: A8(1) = points2(1) + 77:    A8(2) = 0
   
If a >= 105 Then
If b >= 105 Then
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(a1, a2)
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A3, A4)
lineObj.Layer = "Ball-6"
lineObj.Update

If a >= 180 Then

Set lineObj = ThisDrawing.ModelSpace.AddLine(A6, A7)
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A8, A5)
lineObj.Layer = "Ball-6"
lineObj.Update

End If

  pointsark(0) = points2(0) + 47:    pointsark(1) = points2(1) + 45
  pointsark(2) = points2(4) - 47:    pointsark(3) = points2(1) + 45

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsark)
k = (pointsark(2) - pointsark(0)) / 2
    ' Change the bulge of the third segment
    plineObj.SetBulge 0, 21 / k
    plineObj.Update
plineObj.Layer = "Ball-6"
plineObj.Update


  pointsark(0) = points2(0) + 47:    pointsark(1) = points2(3) - 45
  pointsark(2) = points2(4) - 47:    pointsark(3) = points2(3) - 45

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsark)
k = (pointsark(2) - pointsark(0)) / 2
    ' Change the bulge of the third segment
    plineObj.SetBulge 0, -(21 / k)
    plineObj.Update
plineObj.Layer = "Ball-6"
plineObj.Update

                                 'plineObj.color = acByLayer
                                 'Dim layerObj As AcadLayer
                                 'Set layerObj = ThisDrawing.Layers.Add("Ball-6.1")
                                 'layerObj.color = acRed
If a >= 254 Then
If b >= 194 Then

  pointsk(0) = points2(0) + 47:    pointsk(1) = points2(1) + 77
  pointsk(2) = points2(0) + 47:    pointsk(3) = points2(3) - 77
  pointsk(4) = points2(4) - 47:    pointsk(5) = points2(3) - 77
  pointsk(6) = points2(4) - 47:    pointsk(7) = points2(1) + 77
  
  
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsk)
plineObj.Closed = True

plineObj.Layer = "K-grav"
plineObj.Update
  
  plineObj.Layer = "C-mill"
plineObj.Update
  offsetObj = plineObj.Offset(20)
plineObj.Layer = "K-grav"
plineObj.Update

End If
End If

'============================ENGRAVING============================

If a >= 146 Then
If b >= 295 Then

bcp = points2(0) + (b / 2)
acp = points2(1)
M1(0) = points2(4) - (b / 2) - 1: M1(1) = points2(1) + (a / 2)
M2(0) = points2(4) - (b / 2) + 1: M2(1) = points2(1) + (a / 2)

pointsktr1(0) = bcp + 5.68434E-14:    pointsktr1(1) = acp + 60.7166
pointsktr1(2) = bcp + 9.63955:        pointsktr1(3) = acp + 59.23
pointsktr1(4) = bcp + 18.3303:        pointsktr1(5) = acp + 54.1882
pointsktr1(6) = bcp + 33.5681:        pointsktr1(7) = acp + 60.6556
pointsktr1(8) = bcp + 50.0241:        pointsktr1(9) = acp + 59.1337
pointsktr1(10) = bcp + 76.6522:       pointsktr1(11) = acp + 48.2634
pointsktr1(12) = bcp + 95:            pointsktr1(13) = acp + 50.8008
pointsktr1(14) = bcp + 94.1946:       pointsktr1(15) = acp + 49.3519
pointsktr1(16) = bcp + 89.9012:       pointsktr1(17) = acp + 46.7215
pointsktr1(18) = bcp + 84.9367:       pointsktr1(19) = acp + 45.8814
pointsktr1(20) = bcp + 73.728:        pointsktr1(21) = acp + 47.1349
pointsktr1(22) = bcp + 60.6079:       pointsktr1(23) = acp + 52.0547
pointsktr1(24) = bcp + 47.6654:       pointsktr1(25) = acp + 57.4244
pointsktr1(26) = bcp + 36.6652:       pointsktr1(27) = acp + 59.5536
pointsktr1(28) = bcp + 30.7517:       pointsktr1(29) = acp + 59.3172
pointsktr1(30) = bcp + 25.4212:       pointsktr1(31) = acp + 57.9987
pointsktr1(32) = bcp + 19.3968:       pointsktr1(33) = acp + 53.1661
pointsktr1(34) = bcp + 20.1065:       pointsktr1(35) = acp + 43.4681
pointsktr1(36) = bcp + 25.1381:       pointsktr1(37) = acp + 41.104
pointsktr1(38) = bcp + 30.2788:       pointsktr1(39) = acp + 43.2313
pointsktr1(40) = bcp + 28.3471:       pointsktr1(41) = acp + 40.8538
pointsktr1(42) = bcp + 25.3618:       pointsktr1(43) = acp + 40.1632
pointsktr1(44) = bcp + 19.3968:       pointsktr1(45) = acp + 42.2855
pointsktr1(46) = bcp + 9.94206:       pointsktr1(47) = acp + 38.2695
pointsktr1(48) = bcp + 5.68434E-14:   pointsktr1(49) = acp + 40.8656
pointsktr1(50) = bcp - 9.94206:       pointsktr1(51) = acp + 38.2695
pointsktr1(52) = bcp - 19.3968:       pointsktr1(53) = acp + 42.2855
pointsktr1(54) = bcp - 25.3618:       pointsktr1(55) = acp + 40.1632
pointsktr1(56) = bcp - 28.3471:       pointsktr1(57) = acp + 40.8538
pointsktr1(58) = bcp - 30.2788:       pointsktr1(59) = acp + 43.2313
pointsktr1(60) = bcp - 25.1381:       pointsktr1(61) = acp + 41.104
pointsktr1(62) = bcp - 20.1065:       pointsktr1(63) = acp + 43.4681
pointsktr1(64) = bcp - 19.3968:       pointsktr1(65) = acp + 53.1661
pointsktr1(66) = bcp - 25.4212:       pointsktr1(67) = acp + 57.9987
pointsktr1(68) = bcp - 30.7517:       pointsktr1(69) = acp + 59.3172
pointsktr1(70) = bcp - 36.6652:       pointsktr1(71) = acp + 59.5536
pointsktr1(72) = bcp - 47.6654:       pointsktr1(73) = acp + 57.4244
pointsktr1(74) = bcp - 60.6079:       pointsktr1(75) = acp + 52.0547
pointsktr1(76) = bcp - 73.728:        pointsktr1(77) = acp + 47.1349
pointsktr1(78) = bcp - 84.9367:       pointsktr1(79) = acp + 45.8814
pointsktr1(80) = bcp - 89.9012:       pointsktr1(81) = acp + 46.7215
pointsktr1(82) = bcp - 94.1946:       pointsktr1(83) = acp + 49.3519
pointsktr1(84) = bcp - 95:            pointsktr1(85) = acp + 50.8008
pointsktr1(86) = bcp - 76.6522:       pointsktr1(87) = acp + 48.2634
pointsktr1(88) = bcp - 50.0241:       pointsktr1(89) = acp + 59.1337
pointsktr1(90) = bcp - 33.5681:       pointsktr1(91) = acp + 60.6556
pointsktr1(92) = bcp - 18.3303:       pointsktr1(93) = acp + 54.1882
pointsktr1(94) = bcp - 9.63955:       pointsktr1(95) = acp + 59.23

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsktr1)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update
  
plineObj.SetBulge 0, -0.0694346
plineObj.SetBulge 1, -0.143464
plineObj.SetBulge 2, -0.205036
plineObj.SetBulge 3, -0.0755612
plineObj.SetBulge 4, -0.0499684
plineObj.SetBulge 5, 0.307878
plineObj.SetBulge 6, -0.174812
plineObj.SetBulge 7, -0.109731
plineObj.SetBulge 8, -0.0671936
plineObj.SetBulge 9, -0.080044
plineObj.SetBulge 10, -0.0258583
plineObj.SetBulge 11, 0.0120843
plineObj.SetBulge 12, 0.0920047
plineObj.SetBulge 13, 0.0206841
plineObj.SetBulge 14, 0.075674
plineObj.SetBulge 15, 0.163488
plineObj.SetBulge 16, -0.309365
plineObj.SetBulge 17, 0.242665
plineObj.SetBulge 18, 0.171279
plineObj.SetBulge 19, -0.21038
plineObj.SetBulge 20, -0.118469
plineObj.SetBulge 21, -0.187197
plineObj.SetBulge 22, -0.203467
plineObj.SetBulge 23, -0.123761
plineObj.SetBulge 24, -0.123761
plineObj.SetBulge 25, -0.203467
plineObj.SetBulge 26, -0.187197
plineObj.SetBulge 27, -0.118469
plineObj.SetBulge 28, -0.21038
plineObj.SetBulge 29, 0.171279
plineObj.SetBulge 30, 0.242665
plineObj.SetBulge 31, -0.309365
plineObj.SetBulge 32, 0.163488
plineObj.SetBulge 33, 0.075674
plineObj.SetBulge 34, 0.0206841
plineObj.SetBulge 35, 0.0920047
plineObj.SetBulge 36, 0.0120843
plineObj.SetBulge 37, -0.0258583
plineObj.SetBulge 38, -0.080044
plineObj.SetBulge 39, -0.0671936
plineObj.SetBulge 40, -0.109731
plineObj.SetBulge 41, -0.174812
plineObj.SetBulge 42, 0.307878
plineObj.SetBulge 43, -0.0499684
plineObj.SetBulge 44, -0.0755612
plineObj.SetBulge 45, -0.205036
plineObj.SetBulge 46, -0.143464
plineObj.SetBulge 47, -0.0694346
   
  plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True
    RetVal = plineObj.Mirror(M1, M2)

pointsktr2(0) = bcp + 5.68434E-14:    pointsktr2(1) = acp + 52.6932
pointsktr2(2) = bcp + 2.50477:        pointsktr2(3) = acp + 51.9151
pointsktr2(4) = bcp + 3.90531:        pointsktr2(5) = acp + 50.1933
pointsktr2(6) = bcp + 4.4078:         pointsktr2(7) = acp + 48.158
pointsktr2(8) = bcp + 4.1626:         pointsktr2(9) = acp + 44.8378
pointsktr2(10) = bcp + 2.1293:        pointsktr2(11) = acp + 42.2855
pointsktr2(12) = bcp + 6.07129:       pointsktr2(13) = acp + 41.0498
pointsktr2(14) = bcp + 10.2004:       pointsktr2(15) = acp + 40.9088
pointsktr2(16) = bcp + 14.3105:       pointsktr2(17) = acp + 41.5164
pointsktr2(18) = bcp + 18.2265:       pointsktr2(19) = acp + 43.6516
pointsktr2(20) = bcp + 16.73:         pointsktr2(21) = acp + 48.2914
pointsktr2(22) = bcp + 17.7417:       pointsktr2(23) = acp + 53.1661
pointsktr2(24) = bcp + 9.51482:       pointsktr2(25) = acp + 58.0991
pointsktr2(26) = bcp + 5.68434E-14:   pointsktr2(27) = acp + 59.3172
pointsktr2(28) = bcp - 9.51482:       pointsktr2(29) = acp + 58.0991
pointsktr2(30) = bcp - 17.7417:       pointsktr2(31) = acp + 53.1661
pointsktr2(32) = bcp - 16.73:         pointsktr2(33) = acp + 48.2914
pointsktr2(34) = bcp - 18.2265:       pointsktr2(35) = acp + 43.6516
pointsktr2(36) = bcp - 14.3105:       pointsktr2(37) = acp + 41.5164
pointsktr2(38) = bcp - 10.2004:       pointsktr2(39) = acp + 40.9088
pointsktr2(40) = bcp - 6.07129:       pointsktr2(41) = acp + 41.0498
pointsktr2(42) = bcp - 2.1293:        pointsktr2(43) = acp + 42.2855
pointsktr2(44) = bcp - 4.1626:        pointsktr2(45) = acp + 44.8378
pointsktr2(46) = bcp - 4.4078:        pointsktr2(47) = acp + 48.158
pointsktr2(48) = bcp - 3.90531:       pointsktr2(49) = acp + 50.1933
pointsktr2(50) = bcp - 2.50477:       pointsktr2(51) = acp + 51.9151

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsktr2)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update
  
plineObj.SetBulge 0, -0.164816
plineObj.SetBulge 1, -0.123711
plineObj.SetBulge 2, -0.0740984
plineObj.SetBulge 3, -0.123082
plineObj.SetBulge 4, -0.129395
plineObj.SetBulge 5, 0.0762057
plineObj.SetBulge 6, 0.0503136
plineObj.SetBulge 7, 0.0523918
plineObj.SetBulge 8, 0.128222
plineObj.SetBulge 9, -0.181721
plineObj.SetBulge 10, -0.115229
plineObj.SetBulge 11, 0.147569
plineObj.SetBulge 12, 0.0693834
plineObj.SetBulge 13, 0.0693834
plineObj.SetBulge 14, 0.147569
plineObj.SetBulge 15, -0.115229
plineObj.SetBulge 16, -0.181721
plineObj.SetBulge 17, 0.128222
plineObj.SetBulge 18, 0.0523918
plineObj.SetBulge 19, 0.0503136
plineObj.SetBulge 20, 0.0762057
plineObj.SetBulge 21, -0.129395
plineObj.SetBulge 22, -0.123082
plineObj.SetBulge 23, -0.0740984
plineObj.SetBulge 24, -0.123711
plineObj.SetBulge 25, -0.164816

plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True
    RetVal = plineObj.Mirror(M1, M2)
    
pointsktr3(0) = bcp + 5.68434E-14:  pointsktr3(1) = acp + 50.5643
pointsktr3(2) = bcp + 1.65434:      pointsktr3(3) = acp + 49.1014
pointsktr3(4) = bcp + 1.88941:      pointsktr3(5) = acp + 46.8978
pointsktr3(6) = bcp + 1.30833:      pointsktr3(7) = acp + 44.8785
pointsktr3(8) = bcp + 5.68434E-14:  pointsktr3(9) = acp + 43.2314
pointsktr3(10) = bcp - 1.30833:     pointsktr3(11) = acp + 44.8785
pointsktr3(12) = bcp - 1.88941:     pointsktr3(13) = acp + 46.8978
pointsktr3(14) = bcp - 1.65434:     pointsktr3(15) = acp + 49.1014

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsktr3)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update
  
plineObj.SetBulge 0, -0.294708
plineObj.SetBulge 1, -0.0635714
plineObj.SetBulge 2, -0.14288
plineObj.SetBulge 3, -0.061403
plineObj.SetBulge 4, -0.061403
plineObj.SetBulge 5, -0.14288
plineObj.SetBulge 6, -0.0635714
plineObj.SetBulge 7, -0.294708

plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True
    RetVal = plineObj.Mirror(M1, M2)
   
pointsktr4(0) = bcp - 19.1607: pointsktr4(1) = acp + 44.6506
pointsktr4(2) = bcp - 18.6875: pointsktr4(3) = acp + 51.7474
pointsktr4(4) = bcp - 19.1607: pointsktr4(5) = acp + 44.6506
   
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsktr4)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update
  
plineObj.SetBulge 0, -0.261219
plineObj.SetBulge 1, -0.267301
plineObj.SetBulge 2, 0

plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True
    RetVal = plineObj.Mirror(M1, M2)

pointsktr5(0) = bcp + 19.1607: pointsktr5(1) = acp + 44.6506
pointsktr5(2) = bcp + 18.6875: pointsktr5(3) = acp + 51.7474
pointsktr5(4) = bcp + 19.1607: pointsktr5(5) = acp + 44.6506
    
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsktr5)
  plineObj.Closed = True
  plineObj.Layer = "K-grav_Pattern"
  plineObj.Update
  
plineObj.SetBulge 0, 0.261219
plineObj.SetBulge 1, 0.267301
plineObj.SetBulge 2, 0

plineObj.Layer = "K-grav_Pattern"
    plineObj.Update
    plineObj.Closed = True
    RetVal = plineObj.Mirror(M1, M2)

End If
End If
End If
End If

  points2(1) = points2(3) + I:
  points2(3) = points2(3) + a + I:
  points2(5) = points2(3):
  points2(7) = points2(1):

Else: d = 1
d = d + 1

d = Cells(c, 4)
End If

Next d
c = c + 1

a = Cells(c, 2)
b = Cells(c, 3)

I = 100

Loop
              
ZoomAll
AP.Quit
End Sub

Sub MF162()
' Creating search objects
Dim AP As Excel.Application
Dim WB As Excel.Workbook
Dim WS As Excel.Worksheet

Set AP = Excel.Application
Set WB = AP.Workbooks.Open("C:\Users\������\Desktop\������\2.xlsx")
Set WS = WB.Worksheets("������ �˨������ ������")
c = 19
a = Cells(c, 2)
b = Cells(c, 3)
d = Cells(c, 4)
E = 1
I = 0
On Error Resume Next

' Description of all the points that will make up rectangles
 Dim plineObj As AcadLWPolyline
 Dim plineObjis1 As AcadLWPolyline
 Dim plineObjis2 As AcadLWPolyline
 Dim plineObjis3 As AcadLWPolyline
 Dim plineObjis4 As AcadLWPolyline
 Dim plineObjis5 As AcadLWPolyline
 Dim plineObjis6 As AcadLWPolyline
 Dim plineObjis7 As AcadLWPolyline
 Dim plineObjis8 As AcadLWPolyline
 Dim plineObjisw As AcadLWPolyline
  Dim points(0 To 7) As Double
  Dim pointswithin(0 To 7) As Double
  Dim pointsfill(0 To 23) As Double
  Dim pointsleaf1(0 To 11) As Double
  Dim pointsleaf2(0 To 11) As Double
  Dim pointsleaf3(0 To 11) As Double
  Dim pointsleaf4(0 To 11) As Double
  Dim pointsleaf5(0 To 11) As Double
  Dim pointsleaf6(0 To 11) As Double
  Dim pointsis1(0 To 3) As Double
  Dim pointsis2(0 To 3) As Double
  Dim pointsis3(0 To 3) As Double
  Dim pointsis4(0 To 3) As Double
  Dim pointsis5(0 To 3) As Double
  Dim pointsis6(0 To 3) As Double
  Dim pointsis7(0 To 3) As Double
  Dim pointsis8(0 To 3) As Double
  Dim circleObj As AcadCircle
  Dim circleObj2 As AcadCircle
  Dim center(0 To 2) As Double
  Dim radius As Double
  Dim intPoints1 As Variant
  Dim intPoints2 As Variant
  Dim intPoints3 As Variant
  Dim intPoints4 As Variant
  Dim intPoints5 As Variant
  Dim intPoints6 As Variant
  Dim intPoints7 As Variant
  Dim intPoints8 As Variant
  Dim io As Integer
  Dim pointsarc(0 To 3) As Double
  Dim intPointsfillet1 As Variant
  Dim intPointsfillet2 As Variant
  Dim intPointsfillet3 As Variant
  Dim intPointsfillet4 As Variant
  Dim pointsoffset(0 To 15) As Double
  Dim distx As Double
  Dim disty As Double
  Dim P As Variant
points(6) = 0

Do While a > 0

' Calculate the coordinates of each point
  points(0) = points(6) + I:    points(1) = 0
  points(2) = points(6) + I:    points(3) = a
  points(4) = points(0) + b:    points(5) = a
  points(6) = points(0) + b:    points(7) = 0

' Constructing four segments and connecting the corresponding points

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
  plineObj.Closed = True
  plineObj.Layer = "0"
plineObj.Update
  
  ' Offset the polyline
  Dim offsetObj As Variant
                                 'plineObj.color = acByLayer
                                 'Dim layerObj As AcadLayer
                                 'Set layerObj = ThisDrawing.Layers.Add("Ball-6.1")
                                 'layerObj.color = acRed

If a >= 220 Then
If b >= 202 Then


  pointswithin(0) = points(0) + 66:    pointswithin(1) = points(1) + 102
  pointswithin(2) = points(2) + 66:    pointswithin(3) = points(3) - 102
  pointswithin(4) = points(4) - 66:    pointswithin(5) = points(3) - 102
  pointswithin(6) = points(6) - 66:    pointswithin(7) = points(1) + 102

   
   Set plineObjw = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointswithin)
  ' Find the bulge of the third segment
    Dim currentBulge As Double
    currentBulge = plineObjw.GetBulge(2)
   k = (pointswithin(4) - pointswithin(2)) / 2
    ' Change the bulge of the third segment
    plineObjw.SetBulge 1, -(36 / k)
    plineObjw.Update
    plineObjw.SetBulge 3, -(36 / k)
    plineObjw.Update
  plineObjw.Closed = True
  plineObjw.Layer = "C-Mill"
  plineObjw.Update
  plineObjw.Layer = "C2"
  offsetObj = plineObjw.Offset(6)
  plineObjw.Update
  offsetObj = plineObjw.Offset(-6)
  plineObjw.Update

If a >= 220 Then
If b >= 220 Then
For io = 0 To UBound(offsetObj)
offsetObj(io).Delete
Next
 '' ROUNDINGS (TOP)
  pointsarc(0) = points(0) + 66:              pointsarc(1) = points(3) - 102
  pointsarc(2) = points(4) - 66:              pointsarc(3) = points(3) - 102
  Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsarc)
  currentBulge = plineObj.GetBulge(1)
   plineObj.SetBulge 0, -(36 / k)
    plineObj.Update
  plineObj.Layer = "C2"
  plineObj.Update
   offsetObj = plineObj.Offset(-6)
   plineObj.Delete
   
center(0) = points(0) + 66: center(1) = points(3) - 102: center(2) = 0: radius = 6
Set circleObj = ThisDrawing.ModelSpace.AddCircle(center, radius)
For io = 0 To UBound(offsetObj)
intPointsfillet1 = offsetObj(io).IntersectWith(circleObj, acExtendBoth)
circleObj.Delete
center(0) = points(4) - 66: center(1) = points(3) - 102: center(2) = 0: radius = 6
Set circleObj = ThisDrawing.ModelSpace.AddCircle(center, radius)
intPointsfillet2 = offsetObj(io).IntersectWith(circleObj, acExtendBoth)
circleObj.Delete
offsetObj(io).Delete
Next

'center(0) = intPointsfillet2(0): center(1) = intPointsfillet2(1): center(2) = 0: radius = 1
'Set circleObj = ThisDrawing.ModelSpace.AddCircle(center, radius)
'center(0) = intPointsfillet1(0): center(1) = intPointsfillet1(1): center(2) = 0: radius = 1
'Set circleObj = ThisDrawing.ModelSpace.AddCircle(center, radius)

' ROUNDINGS (BOTTOM)
  pointsarc(0) = points(0) + 66:              pointsarc(1) = points(1) + 102
  pointsarc(2) = points(4) - 66:              pointsarc(3) = points(1) + 102
  Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsarc)
  currentBulge = plineObj.GetBulge(1)
   plineObj.SetBulge 0, (36 / k)
    plineObj.Update
  plineObj.Layer = "C2"
  plineObj.Update
   offsetObj = plineObj.Offset(6)
   plineObj.Delete
   
center(0) = points(0) + 66: center(1) = points(1) + 102: center(2) = 0: radius = 6
Set circleObj = ThisDrawing.ModelSpace.AddCircle(center, radius)
For io = 0 To UBound(offsetObj)
intPointsfillet3 = offsetObj(io).IntersectWith(circleObj, acExtendBoth)
circleObj.Delete
center(0) = points(4) - 66: center(1) = points(1) + 102: center(2) = 0: radius = 6
Set circleObj = ThisDrawing.ModelSpace.AddCircle(center, radius)
intPointsfillet4 = offsetObj(io).IntersectWith(circleObj, acExtendBoth)
circleObj.Delete
offsetObj(io).Delete
Next

'center(0) = intPointsfillet3(0): center(1) = intPointsfillet3(1): center(2) = 0: radius = 1
'Set circleObj = ThisDrawing.ModelSpace.AddCircle(center, radius)
'center(0) = intPointsfillet4(0): center(1) = intPointsfillet4(1): center(2) = 0: radius = 1
'Set circleObj = ThisDrawing.ModelSpace.AddCircle(center, radius)

  pointsoffset(0) = points(0) + 60:              pointsoffset(1) = points(1) + 102
  pointsoffset(2) = points(0) + 60:              pointsoffset(3) = points(3) - 102
  pointsoffset(4) = intPointsfillet1(0):         pointsoffset(5) = intPointsfillet1(1)
  pointsoffset(6) = intPointsfillet2(0):         pointsoffset(7) = intPointsfillet2(1)
  pointsoffset(8) = points(4) - 60:              pointsoffset(9) = points(3) - 102
  pointsoffset(10) = points(4) - 60:             pointsoffset(11) = points(1) + 102
  pointsoffset(12) = intPointsfillet4(0):        pointsoffset(13) = intPointsfillet4(1)
  pointsoffset(14) = intPointsfillet3(0):        pointsoffset(15) = intPointsfillet3(1)
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsoffset)
plineObj.Closed = True
distx = intPointsfillet1(0) - pointsoffset(2)
disty = intPointsfillet1(1) - pointsoffset(3)
dist = Sqr((distx ^ 2) + (disty ^ 2))
h = 6 - Sqr(36 - (dist ^ 2) / 4)
kr = h / dist


    plineObj.SetBulge 1, -2 * kr
    plineObj.Update
        plineObj.SetBulge 2, -(36 / k)
        plineObj.Update
    plineObj.SetBulge 3, -2 * kr
    plineObj.Update
    plineObj.SetBulge 5, -2 * kr
    plineObj.Update
        plineObj.SetBulge 6, -(36 / k)
        plineObj.Update
    plineObj.SetBulge 7, -2 * kr
    plineObj.Update
    plineObj.Layer = "C2"
    plineObj.Update
  plineObj.Closed = True

End If
End If


If a >= 235 Then
If b >= 235 Then

 plineObjw.Layer = "Ball-6"
 plineObjw.Update
   offsetObj = plineObjw.Offset(30)
For io = 0 To UBound(offsetObj)
offsetObj(io).Delete
Next
End If
End If

plineObjw.Layer = "C-Mill"
plineObjw.Update

r = (0.25 * (pointswithin(4) - pointswithin(2)) ^ 2 / 36 + 36) / 2
r1 = r - 30
 ' Creating a circle

  center(0) = points(0) + (b / 2): center(1) = points(3) - r - 66: center(2) = 0: radius = r1
  Set circleObj = ThisDrawing.ModelSpace.AddCircle(center, radius)
  
  pointsis1(0) = points(0) + (b / 2) - 5:     pointsis1(1) = points(3) - 130
  pointsis1(2) = points(0) + (b / 2) - 5:     pointsis1(3) = points(3)
  pointsis2(0) = points(0) + (b / 2) + 5:     pointsis2(1) = points(3) - 130
  pointsis2(2) = points(0) + (b / 2) + 5:     pointsis2(3) = points(3)
  pointsis3(0) = points(0) + 96:              pointsis3(1) = points(3) - 130
  pointsis3(2) = points(0) + 96:              pointsis3(3) = points(3)
  pointsis4(0) = points(4) - 96:              pointsis4(1) = points(3) - 130
  pointsis4(2) = points(4) - 96:              pointsis4(3) = points(3)
  Set plineObjis1 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsis1)
  Set plineObjis2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsis2)
  Set plineObjis3 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsis3)
  Set plineObjis4 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsis4)
  
  
  ' Find the intersection points between the line and the circle
    
    intPoints1 = plineObjis1.IntersectWith(circleObj, acExtendNone)
    intPoints2 = plineObjis2.IntersectWith(circleObj, acExtendNone)
    intPoints3 = plineObjis3.IntersectWith(circleObj, acExtendNone)
    intPoints4 = plineObjis4.IntersectWith(circleObj, acExtendNone)
    
   circleObj.Delete
    
  center(0) = points(0) + (b / 2): center(1) = points(1) + r + 66: center(2) = 0: radius = r1
  Set circleObj2 = ThisDrawing.ModelSpace.AddCircle(center, radius)
  
  pointsis5(0) = points(0) + (b / 2) - 5:     pointsis5(1) = points(1)
  pointsis5(2) = points(0) + (b / 2) - 5:     pointsis5(3) = points(1) + 130
  pointsis6(0) = points(0) + (b / 2) + 5:     pointsis6(1) = points(1)
  pointsis6(2) = points(0) + (b / 2) + 5:     pointsis6(3) = points(1) + 130
  pointsis7(0) = points(0) + 96:              pointsis7(1) = points(1)
  pointsis7(2) = points(0) + 96:              pointsis7(3) = points(1) + 130
  pointsis8(0) = points(4) - 96:              pointsis8(1) = points(1)
  pointsis8(2) = points(4) - 96:              pointsis8(3) = points(1) + 130
  Set plineObjis5 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsis5)
  Set plineObjis6 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsis6)
  Set plineObjis7 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsis7)
  Set plineObjis8 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsis8)
    
    intPoints5 = plineObjis5.IntersectWith(circleObj2, acExtendNone)
    intPoints6 = plineObjis6.IntersectWith(circleObj2, acExtendNone)
    intPoints7 = plineObjis7.IntersectWith(circleObj2, acExtendNone)
    intPoints8 = plineObjis8.IntersectWith(circleObj2, acExtendNone)
  
    plineObjis1.Delete
    plineObjis2.Delete
    plineObjis3.Delete
    plineObjis4.Delete
    plineObjis5.Delete
    plineObjis6.Delete
    plineObjis7.Delete
    plineObjis8.Delete
      
    circleObj2.Delete
    
pointsfill(0) = points(0) + 96:              pointsfill(1) = intPoints8(1)
pointsfill(2) = points(0) + 96:              pointsfill(3) = intPoints4(1)
pointsfill(4) = intPoints1(0):               pointsfill(5) = intPoints1(1)
pointsfill(6) = points(0) + (b / 2) + 4.12:  pointsfill(7) = intPoints1(1) - 21.41
pointsfill(8) = points(0) + (b / 2) - 4.12:  pointsfill(9) = intPoints1(1) - 21.41
pointsfill(10) = intPoints2(0):              pointsfill(11) = intPoints2(1)
pointsfill(12) = points(4) - 96:             pointsfill(13) = intPoints3(1)
pointsfill(14) = points(4) - 96:             pointsfill(15) = intPoints7(1)
pointsfill(16) = intPoints6(0):              pointsfill(17) = intPoints6(1)
pointsfill(18) = points(0) + (b / 2) - 4.12: pointsfill(19) = intPoints6(1) + 21.41
pointsfill(20) = points(0) + (b / 2) + 4.12: pointsfill(21) = intPoints6(1) + 21.41
pointsfill(22) = intPoints5(0):              pointsfill(23) = intPoints5(1)

If a >= 250 Then
If b >= 250 Then

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsfill)
plineObj.Closed = True

P = intPoints1(0) - intPoints4(0)
p1 = intPoints1(1) - intPoints4(1)
g = Sqr((P * P) + (p1 * p1))
   angle = Atn((p1 / g) / Sqr((-p1 / g) * (p1 / g) + 1))
   radius = (g / 2) / Sin(p1 / g)
   h = radius * (1 - Cos(angle))
   k = h / g

   currentBulge = plineObj.GetBulge(2)
   K1 = (intPoints6(0) - intPoints5(0)) / 2
   K2 = 124 - (points(3) - intPoints1(1))
    ' Change the bulge of the third segment
    plineObj.SetBulge 1, -2 * k
    plineObj.Update
    plineObj.SetBulge 2, -0.54
    plineObj.Update
    plineObj.SetBulge 3, -0.54
    plineObj.Update
    plineObj.SetBulge 4, -0.54
    plineObj.Update
    plineObj.SetBulge 5, -2 * k
    plineObj.Update
    plineObj.SetBulge 7, -2 * k
    plineObj.Update
    plineObj.SetBulge 8, -0.54
    plineObj.Update
    plineObj.SetBulge 9, -0.54
    plineObj.Update
    plineObj.SetBulge 10, -0.54
    plineObj.Update
    plineObj.SetBulge 11, -2 * k
    plineObj.Update
    plineObj.Layer = "Ball-6"
    plineObj.Update
  plineObj.Closed = True


End If
End If
End If
End If

   
I = 100
  
 Dim plineObj2 As AcadLWPolyline
  Dim points2(0 To 7) As Double
  
  points2(0) = points(0):    points2(1) = a + I:
  points2(2) = points(0):    points2(3) = (a * 2) + I:
  points2(4) = points(6):    points2(5) = (a * 2) + I:
  points2(6) = points(6):    points2(7) = a + I:
  
 For d = 2 To Cells(c, 4)
If d > 1 Then

Set plineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(points2)
  plineObj2.Closed = True
  
If a >= 220 Then
If b >= 202 Then

  pointswithin(0) = points2(0) + 66:    pointswithin(1) = points2(1) + 102
  pointswithin(2) = points2(2) + 66:    pointswithin(3) = points2(3) - 102
  pointswithin(4) = points2(4) - 66:    pointswithin(5) = points2(3) - 102
  pointswithin(6) = points2(6) - 66:    pointswithin(7) = points2(1) + 102

   
   Set plineObjw = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointswithin)
  ' Find the bulge of the third segment
    currentBulge = plineObjw.GetBulge(2)
   k = (pointswithin(4) - pointswithin(2)) / 2
    ' Change the bulge of the third segment
    plineObjw.SetBulge 1, -(36 / k)
    plineObjw.Update
    plineObjw.SetBulge 3, -(36 / k)
    plineObjw.Update
  plineObjw.Closed = True
  plineObjw.Layer = "C-Mill"
  plineObjw.Update
  plineObjw.Layer = "C2"
  offsetObj = plineObjw.Offset(6)
  plineObjw.Update
  offsetObj = plineObjw.Offset(-6)
  plineObjw.Update
  
If a >= 220 Then
If b >= 220 Then
For io = 0 To UBound(offsetObj)
offsetObj(io).Delete
Next
 '' ROUNDINGS (TOP)
  pointsarc(0) = points2(0) + 66:              pointsarc(1) = points2(3) - 102
  pointsarc(2) = points2(4) - 66:              pointsarc(3) = points2(3) - 102
  Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsarc)
  currentBulge = plineObj.GetBulge(1)
   plineObj.SetBulge 0, -(36 / k)
    plineObj.Update
  plineObj.Layer = "C2"
  plineObj.Update
   offsetObj = plineObj.Offset(-6)
   plineObj.Delete
   
center(0) = points2(0) + 66: center(1) = points2(3) - 102: center(2) = 0: radius = 6
Set circleObj = ThisDrawing.ModelSpace.AddCircle(center, radius)
For io = 0 To UBound(offsetObj)
intPointsfillet1 = offsetObj(io).IntersectWith(circleObj, acExtendBoth)
circleObj.Delete
center(0) = points2(4) - 66: center(1) = points2(3) - 102: center(2) = 0: radius = 6
Set circleObj = ThisDrawing.ModelSpace.AddCircle(center, radius)
intPointsfillet2 = offsetObj(io).IntersectWith(circleObj, acExtendBoth)
circleObj.Delete
offsetObj(io).Delete
Next

'center(0) = intPointsfillet2(0): center(1) = intPointsfillet2(1): center(2) = 0: radius = 1
'Set circleObj = ThisDrawing.ModelSpace.AddCircle(center, radius)
'center(0) = intPointsfillet1(0): center(1) = intPointsfillet1(1): center(2) = 0: radius = 1
'Set circleObj = ThisDrawing.ModelSpace.AddCircle(center, radius)

' ROUNDINGS (BOTTOM)
  pointsarc(0) = points2(0) + 66:              pointsarc(1) = points2(1) + 102
  pointsarc(2) = points2(4) - 66:              pointsarc(3) = points2(1) + 102
  Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsarc)
  currentBulge = plineObj.GetBulge(1)
   plineObj.SetBulge 0, (36 / k)
    plineObj.Update
  plineObj.Layer = "C2"
  plineObj.Update
   offsetObj = plineObj.Offset(6)
   plineObj.Delete
   
center(0) = points2(0) + 66: center(1) = points2(1) + 102: center(2) = 0: radius = 6
Set circleObj = ThisDrawing.ModelSpace.AddCircle(center, radius)
For io = 0 To UBound(offsetObj)
intPointsfillet3 = offsetObj(io).IntersectWith(circleObj, acExtendBoth)
circleObj.Delete
center(0) = points2(4) - 66: center(1) = points2(1) + 102: center(2) = 0: radius = 6
Set circleObj = ThisDrawing.ModelSpace.AddCircle(center, radius)
intPointsfillet4 = offsetObj(io).IntersectWith(circleObj, acExtendBoth)
circleObj.Delete
offsetObj(io).Delete
Next

'center(0) = intPointsfillet3(0): center(1) = intPointsfillet3(1): center(2) = 0: radius = 1
'Set circleObj = ThisDrawing.ModelSpace.AddCircle(center, radius)
'center(0) = intPointsfillet4(0): center(1) = intPointsfillet4(1): center(2) = 0: radius = 1
'Set circleObj = ThisDrawing.ModelSpace.AddCircle(center, radius)

  pointsoffset(0) = points2(0) + 60:              pointsoffset(1) = points2(1) + 102
  pointsoffset(2) = points2(0) + 60:              pointsoffset(3) = points2(3) - 102
  pointsoffset(4) = intPointsfillet1(0):         pointsoffset(5) = intPointsfillet1(1)
  pointsoffset(6) = intPointsfillet2(0):         pointsoffset(7) = intPointsfillet2(1)
  pointsoffset(8) = points2(4) - 60:              pointsoffset(9) = points2(3) - 102
  pointsoffset(10) = points2(4) - 60:             pointsoffset(11) = points2(1) + 102
  pointsoffset(12) = intPointsfillet4(0):        pointsoffset(13) = intPointsfillet4(1)
  pointsoffset(14) = intPointsfillet3(0):        pointsoffset(15) = intPointsfillet3(1)
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsoffset)
plineObj.Closed = True
distx = intPointsfillet1(0) - pointsoffset(2)
disty = intPointsfillet1(1) - pointsoffset(3)
dist = Sqr((distx ^ 2) + (disty ^ 2))
h = 6 - Sqr(36 - (dist ^ 2) / 4)
kr = h / dist


    plineObj.SetBulge 1, -2 * kr
    plineObj.Update
        plineObj.SetBulge 2, -(36 / k)
        plineObj.Update
    plineObj.SetBulge 3, -2 * kr
    plineObj.Update
    plineObj.SetBulge 5, -2 * kr
    plineObj.Update
        plineObj.SetBulge 6, -(36 / k)
        plineObj.Update
    plineObj.SetBulge 7, -2 * kr
    plineObj.Update
    plineObj.Layer = "C2"
    plineObj.Update
  plineObj.Closed = True

End If
End If


If a >= 235 Then
If b >= 235 Then
plineObjw.Layer = "Ball-6"
 plineObjw.Update
   offsetObj = plineObjw.Offset(30)

End If
End If
  
plineObjw.Layer = "C-Mill"
plineObjw.Update

r = (0.25 * (pointswithin(4) - pointswithin(2)) ^ 2 / 36 + 36) / 2
r1 = r - 30
 ' Creating a circle
  center(0) = points2(0) + (b / 2): center(1) = points2(3) - r - 66: center(2) = 0: radius = r1
  Set circleObj = ThisDrawing.ModelSpace.AddCircle(center, radius)
  
  pointsis1(0) = points2(0) + (b / 2) - 5:     pointsis1(1) = points2(3) - 130
  pointsis1(2) = points2(0) + (b / 2) - 5:     pointsis1(3) = points2(3)
  pointsis2(0) = points2(0) + (b / 2) + 5:     pointsis2(1) = points2(3) - 130
  pointsis2(2) = points2(0) + (b / 2) + 5:     pointsis2(3) = points2(3)
  pointsis3(0) = points2(0) + 96:              pointsis3(1) = points2(3) - 130
  pointsis3(2) = points2(0) + 96:              pointsis3(3) = points2(3)
  pointsis4(0) = points2(4) - 96:              pointsis4(1) = points2(3) - 130
  pointsis4(2) = points2(4) - 96:              pointsis4(3) = points2(3)
  Set plineObjis1 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsis1)
  Set plineObjis2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsis2)
  Set plineObjis3 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsis3)
  Set plineObjis4 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsis4)
  
  
  ' Find the intersection points between the line and the circle
    intPoints1 = plineObjis1.IntersectWith(circleObj, acExtendNone)
    intPoints2 = plineObjis2.IntersectWith(circleObj, acExtendNone)
    intPoints3 = plineObjis3.IntersectWith(circleObj, acExtendNone)
    intPoints4 = plineObjis4.IntersectWith(circleObj, acExtendNone)
    
    circleObj.Delete
    
  center(0) = points2(0) + (b / 2): center(1) = points2(1) + r + 66: center(2) = 0: radius = r1
  Set circleObj2 = ThisDrawing.ModelSpace.AddCircle(center, radius)
  
  pointsis5(0) = points2(0) + (b / 2) - 5:     pointsis5(1) = points2(1)
  pointsis5(2) = points2(0) + (b / 2) - 5:     pointsis5(3) = points2(1) + 130
  pointsis6(0) = points2(0) + (b / 2) + 5:     pointsis6(1) = points2(1)
  pointsis6(2) = points2(0) + (b / 2) + 5:     pointsis6(3) = points2(1) + 130
  pointsis7(0) = points2(0) + 96:              pointsis7(1) = points2(1)
  pointsis7(2) = points2(0) + 96:              pointsis7(3) = points2(1) + 130
  pointsis8(0) = points2(4) - 96:              pointsis8(1) = points2(1)
  pointsis8(2) = points2(4) - 96:              pointsis8(3) = points2(1) + 130
  Set plineObjis5 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsis5)
  Set plineObjis6 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsis6)
  Set plineObjis7 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsis7)
  Set plineObjis8 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsis8)
    
    intPoints5 = plineObjis5.IntersectWith(circleObj2, acExtendNone)
    intPoints6 = plineObjis6.IntersectWith(circleObj2, acExtendNone)
    intPoints7 = plineObjis7.IntersectWith(circleObj2, acExtendNone)
    intPoints8 = plineObjis8.IntersectWith(circleObj2, acExtendNone)
  
    plineObjis1.Delete
    plineObjis2.Delete
    plineObjis3.Delete
    plineObjis4.Delete
    plineObjis5.Delete
    plineObjis6.Delete
    plineObjis7.Delete
    plineObjis8.Delete
      
    
    circleObj2.Delete
    
pointsfill(0) = points2(0) + 96:              pointsfill(1) = intPoints8(1)
pointsfill(2) = points2(0) + 96:              pointsfill(3) = intPoints4(1)
pointsfill(4) = intPoints1(0):                pointsfill(5) = intPoints1(1)
pointsfill(6) = points2(0) + (b / 2) + 4.12:  pointsfill(7) = intPoints1(1) - 21.41
pointsfill(8) = points2(0) + (b / 2) - 4.12:  pointsfill(9) = intPoints1(1) - 21.41
pointsfill(10) = intPoints2(0):               pointsfill(11) = intPoints2(1)
pointsfill(12) = points2(4) - 96:             pointsfill(13) = intPoints3(1)
pointsfill(14) = points2(4) - 96:             pointsfill(15) = intPoints7(1)
pointsfill(16) = intPoints6(0):               pointsfill(17) = intPoints6(1)
pointsfill(18) = points2(0) + (b / 2) - 4.12: pointsfill(19) = intPoints6(1) + 21.41
pointsfill(20) = points2(0) + (b / 2) + 4.12: pointsfill(21) = intPoints6(1) + 21.41
pointsfill(22) = intPoints5(0):               pointsfill(23) = intPoints5(1)

If a >= 250 Then
If b >= 250 Then
For io = 0 To UBound(offsetObj)
offsetObj(io).Delete
Next
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsfill)
plineObj.Closed = True

P = intPoints1(0) - intPoints4(0)
p1 = intPoints1(1) - intPoints4(1)
g = Sqr((P * P) + (p1 * p1))
   angle = Atn((p1 / g) / Sqr((-p1 / g) * (p1 / g) + 1))
   radius = (g / 2) / Sin(p1 / g)
   h = radius * (1 - Cos(angle))
   k = h / g

   currentBulge = plineObj.GetBulge(2)
   K1 = (intPoints6(0) - intPoints5(0)) / 2
   K2 = 124 - (points2(3) - intPoints1(1))
    ' Change the bulge of the third segment
 plineObj.SetBulge 1, -2 * k
    plineObj.Update
    plineObj.SetBulge 2, -0.54
    plineObj.Update
    plineObj.SetBulge 3, -0.54
    plineObj.Update
    plineObj.SetBulge 4, -0.54
    plineObj.Update
    plineObj.SetBulge 5, -2 * k
    plineObj.Update
    plineObj.SetBulge 7, -2 * k
    plineObj.Update
    plineObj.SetBulge 8, -0.54
    plineObj.Update
    plineObj.SetBulge 9, -0.54
    plineObj.Update
    plineObj.SetBulge 10, -0.54
    plineObj.Update
    plineObj.SetBulge 11, -2 * k
    plineObj.Update
    plineObj.Layer = "Ball-6"
    plineObj.Update
  plineObj.Closed = True

End If
End If
End If
End If
  
  points2(1) = points2(3) + I:
  points2(3) = points2(3) + a + I:
  points2(5) = points2(3):
  points2(7) = points2(1):
  
  

Else: d = 1
d = d + 1

d = Cells(c, 4)
End If

Next d
c = c + 1

a = Cells(c, 2)
b = Cells(c, 3)

I = 100

Loop
              
ZoomAll
AP.Quit
End Sub

Sub MF165()
' Creating search objects
Dim AP As Excel.Application
Dim WB As Excel.Workbook
Dim WS As Excel.Worksheet

Set AP = Excel.Application
Set WB = AP.Workbooks.Open("C:\Users\������\Desktop\������\2.xlsx")
Set WS = WB.Worksheets("������ �˨������ ������")
c = 19
a = Cells(c, 2)
b = Cells(c, 3)
d = Cells(c, 4)
E = 1
I = 0
On Error Resume Next

' Description of all the points that will make up rectangles
 Dim plineObj As AcadLWPolyline
 Dim plineObjw As AcadLWPolyline
  Dim points(0 To 7) As Double
  Dim pointswithin(0 To 19) As Double
  Dim pointsqrt(0 To 9) As Double
  Dim pointsqrt2(0 To 9) As Double
  Dim pointsarrow1(0 To 5) As Double
  Dim pointsarrow2(0 To 5) As Double
  Dim pointsarrow3(0 To 5) As Double
points(6) = 0

Do While a > 0

' Calculate the coordinates of each point
  points(0) = points(6) + I:    points(1) = 0
  points(2) = points(6) + I:    points(3) = a
  points(4) = points(0) + b:    points(5) = a
  points(6) = points(0) + b:    points(7) = 0

' Constructing four segments and connecting the corresponding points

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
  plineObj.Closed = True
 
  
pointswithin(0) = points(0) + 37:         pointswithin(1) = points(1) + 37:
pointswithin(2) = points(0) + 37:         pointswithin(3) = points(3) - 61:
pointswithin(4) = points(0) + 62:         pointswithin(5) = points(3) - 61:
pointswithin(6) = points(0) + 73.7805:    pointswithin(7) = points(3) - 54.1041:
pointswithin(8) = points(0) + 103:        pointswithin(9) = points(3) - 37:
pointswithin(10) = points(4) - 103:       pointswithin(11) = points(3) - 37:
pointswithin(12) = points(4) - 73.7805:   pointswithin(13) = points(3) - 54.1041:
pointswithin(14) = points(4) - 62:        pointswithin(15) = points(3) - 61:
pointswithin(16) = points(4) - 37:        pointswithin(17) = points(3) - 61:
pointswithin(18) = points(4) - 37:        pointswithin(19) = points(1) + 37:


If a >= 260 Then
If b >= 260 Then

plineObjw.Layer = "Ball-6"
plineObjw.Update
Set plineObjw = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointswithin)
  ' Find the bulge of the third segment
    Dim currentBulge As Double
    currentBulge = plineObjw.GetBulge(2)
  ' Change the bulge of the third segment
    plineObjw.SetBulge 2, 0.27116
    plineObjw.Update
    plineObjw.SetBulge 3, -0.27116
    plineObjw.Update
    plineObjw.SetBulge 5, -0.27116
    plineObjw.Update
    plineObjw.SetBulge 6, 0.27116
    plineObjw.Update
    plineObjw.Closed = True

  ' Offset the polyline
  Dim offsetObj As Variant
                                 'plineObj.color = acByLayer
                                 'Dim layerObj As AcadLayer
                                 'Set layerObj = ThisDrawing.Layers.Add("Ball-6.1")
                                 'layerObj.color = acRed

plineObjw.Layer = "K-grav"
plineObjw.Update
  offsetObj = plineObjw.Offset(10)
plineObjw.Layer = "K-Mill"
plineObjw.Update
  offsetObj = plineObjw.Offset(28)
  offsetObj = plineObjw.Offset(29)
  offsetObj = plineObjw.Offset(30)
plineObjw.Layer = "D-Mill"
plineObjw.Update
  offsetObj = plineObjw.Offset(34)
plineObjw.Layer = "Ball-6"
plineObjw.Update

pointsqrt(0) = points(0) + 88:   pointsqrt(1) = points(1) + 88:
pointsqrt(2) = points(0) + 88:   pointsqrt(3) = points(1) + 126:
pointsqrt(4) = points(0) + 94:   pointsqrt(5) = points(1) + 126:
pointsqrt(6) = points(0) + 126:  pointsqrt(7) = points(1) + 94:
pointsqrt(8) = points(0) + 126:  pointsqrt(9) = points(1) + 88:


Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsqrt)
' Find the bulge of the segment
    currentBulge = plineObj.GetBulge(2)
  ' Change the bulge of the third segment
    plineObj.SetBulge 2, -0.41421356
    plineObj.Layer = "AreaClear"
plineObj.Update
plineObj.Closed = True

pointsqrt2(0) = points(4) - 88:    pointsqrt2(1) = points(1) + 88:
pointsqrt2(2) = points(4) - 88:    pointsqrt2(3) = points(1) + 126:
pointsqrt2(4) = points(4) - 94:     pointsqrt2(5) = points(1) + 126:
pointsqrt2(6) = points(4) - 126:   pointsqrt2(7) = points(1) + 94:
pointsqrt2(8) = points(4) - 126:   pointsqrt2(9) = points(1) + 88:


Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsqrt2)
' Find the bulge of the segment
    currentBulge = plineObj.GetBulge(2)
  ' Change the bulge of the third segment
    plineObj.SetBulge 2, 0.41421356
    plineObj.Update
    plineObj.Layer = "AreaClear"
plineObj.Update
plineObj.Closed = True

'============================Arrows on the left============================
pointsarrow1(0) = points(0) + 98:    pointsarrow1(1) = points(1) + 100:
pointsarrow1(2) = points(0) + 96:    pointsarrow1(3) = points(1) + 118:
pointsarrow1(4) = points(0) + 100:   pointsarrow1(5) = points(1) + 118:

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsarrow1)
plineObj.Layer = "K-grav"
plineObj.Update
plineObj.Closed = True

pointsarrow2(0) = points(0) + 99.4142:    pointsarrow2(1) = points(1) + 99.4142:
pointsarrow2(2) = points(0) + 110.7279:   pointsarrow2(3) = points(1) + 113.5564:
pointsarrow2(4) = points(0) + 113.5564:   pointsarrow2(5) = points(1) + 110.7279:

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsarrow2)
plineObj.Layer = "K-grav"
plineObj.Update
plineObj.Closed = True

pointsarrow3(0) = points(0) + 100:    pointsarrow3(1) = points(1) + 98:
pointsarrow3(2) = points(0) + 118:    pointsarrow3(3) = points(1) + 96:
pointsarrow3(4) = points(0) + 118:    pointsarrow3(5) = points(1) + 100:

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsarrow3)
plineObj.Layer = "K-grav"
plineObj.Update
plineObj.Closed = True

'============================Arrows on the right============================
pointsarrow1(0) = points(4) - 98:    pointsarrow1(1) = points(1) + 100:
pointsarrow1(2) = points(4) - 96:    pointsarrow1(3) = points(1) + 118:
pointsarrow1(4) = points(4) - 100:   pointsarrow1(5) = points(1) + 118:

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsarrow1)
plineObj.Layer = "K-grav"
plineObj.Update
plineObj.Closed = True

pointsarrow2(0) = points(4) - 99.4142:    pointsarrow2(1) = points(1) + 99.4142:
pointsarrow2(2) = points(4) - 110.7279:   pointsarrow2(3) = points(1) + 113.5564:
pointsarrow2(4) = points(4) - 113.5564:   pointsarrow2(5) = points(1) + 110.7279:

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsarrow2)
plineObj.Layer = "K-grav"
plineObj.Update
plineObj.Closed = True

pointsarrow3(0) = points(4) - 100:    pointsarrow3(1) = points(1) + 98:
pointsarrow3(2) = points(4) - 118:    pointsarrow3(3) = points(1) + 96:
pointsarrow3(4) = points(4) - 118:    pointsarrow3(5) = points(1) + 100:

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsarrow3)
plineObj.Layer = "K-grav"
plineObj.Update
plineObj.Closed = True

End If
End If

I = 100
  
 Dim plineObj2 As AcadLWPolyline
  Dim points2(0 To 7) As Double
  
  points2(0) = points(0):    points2(1) = a + I:
  points2(2) = points(0):    points2(3) = (a * 2) + I:
  points2(4) = points(6):    points2(5) = (a * 2) + I:
  points2(6) = points(6):    points2(7) = a + I:
  
 For d = 2 To Cells(c, 4)
If d > 1 Then

Set plineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(points2)
  plineObj2.Closed = True
  

pointswithin(0) = points2(0) + 37:         pointswithin(1) = points2(1) + 37:
pointswithin(2) = points2(0) + 37:         pointswithin(3) = points2(3) - 61:
pointswithin(4) = points2(0) + 62:         pointswithin(5) = points2(3) - 61:
pointswithin(6) = points2(0) + 73.7805:    pointswithin(7) = points2(3) - 54.1041:
pointswithin(8) = points2(0) + 103:        pointswithin(9) = points2(3) - 37:
pointswithin(10) = points2(4) - 103:       pointswithin(11) = points2(3) - 37:
pointswithin(12) = points2(4) - 73.7805:   pointswithin(13) = points2(3) - 54.1041:
pointswithin(14) = points2(4) - 62:        pointswithin(15) = points2(3) - 61:
pointswithin(16) = points2(4) - 37:        pointswithin(17) = points2(3) - 61:
pointswithin(18) = points2(4) - 37:        pointswithin(19) = points2(1) + 37:




If a >= 260 Then
If b >= 260 Then

plineObjw.Layer = "Ball-6"
plineObjw.Update
Set plineObjw = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointswithin)
  ' Find the bulge of the third segment
 
    currentBulge = plineObjw.GetBulge(2)
  ' Change the bulge of the third segment
    plineObjw.SetBulge 2, 0.27116
    plineObjw.Update
    plineObjw.SetBulge 3, -0.27116
    plineObjw.Update
    plineObjw.SetBulge 5, -0.27116
    plineObjw.Update
    plineObjw.SetBulge 6, 0.27116
    plineObjw.Update
    plineObjw.Update
    plineObjw.Closed = True

  ' Offset the polyline

                                 'plineObj.color = acByLayer
                                 'Dim layerObj As AcadLayer
                                 'Set layerObj = ThisDrawing.Layers.Add("Ball-6.1")
                                 'layerObj.color = acRed

plineObjw.Layer = "K-grav"
plineObjw.Update
  offsetObj = plineObjw.Offset(10)
plineObjw.Layer = "K-Mill"
plineObjw.Update
  offsetObj = plineObjw.Offset(28)
  offsetObj = plineObjw.Offset(29)
  offsetObj = plineObjw.Offset(30)
plineObjw.Layer = "D-Mill"
plineObjw.Update
  offsetObj = plineObjw.Offset(34)
plineObjw.Layer = "Ball-6"
plineObjw.Update

pointsqrt(0) = points2(0) + 88:   pointsqrt(1) = points2(1) + 88:
pointsqrt(2) = points2(0) + 88:   pointsqrt(3) = points2(1) + 126:
pointsqrt(4) = points2(0) + 94:   pointsqrt(5) = points2(1) + 126:
pointsqrt(6) = points2(0) + 126:  pointsqrt(7) = points2(1) + 94:
pointsqrt(8) = points2(0) + 126:  pointsqrt(9) = points2(1) + 88:


Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsqrt)
' Find the bulge of the segment
    currentBulge = plineObj.GetBulge(2)
  ' Change the bulge of the third segment
    plineObj.SetBulge 2, -0.41421356
    plineObj.Layer = "AreaClear"
plineObj.Update
plineObj.Closed = True

pointsqrt2(0) = points2(4) - 88:    pointsqrt2(1) = points2(1) + 88:
pointsqrt2(2) = points2(4) - 88:    pointsqrt2(3) = points2(1) + 126:
pointsqrt2(4) = points2(4) - 94:    pointsqrt2(5) = points2(1) + 126:
pointsqrt2(6) = points2(4) - 126:   pointsqrt2(7) = points2(1) + 94:
pointsqrt2(8) = points2(4) - 126:   pointsqrt2(9) = points2(1) + 88:


Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsqrt2)
' Find the bulge of the segment
    currentBulge = plineObj.GetBulge(2)
  ' Change the bulge of the third segment
    plineObj.SetBulge 2, 0.41421356
    plineObj.Update
    plineObj.Layer = "AreaClear"
plineObj.Update
plineObj.Closed = True

'============================Arrows on the left============================
pointsarrow1(0) = points2(0) + 98:    pointsarrow1(1) = points2(1) + 100:
pointsarrow1(2) = points2(0) + 96:    pointsarrow1(3) = points2(1) + 118:
pointsarrow1(4) = points2(0) + 100:   pointsarrow1(5) = points2(1) + 118:

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsarrow1)
plineObj.Layer = "K-grav"
plineObj.Update
plineObj.Closed = True

pointsarrow2(0) = points2(0) + 99.4142:    pointsarrow2(1) = points2(1) + 99.4142:
pointsarrow2(2) = points2(0) + 110.7279:   pointsarrow2(3) = points2(1) + 113.5564:
pointsarrow2(4) = points2(0) + 113.5564:   pointsarrow2(5) = points2(1) + 110.7279:

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsarrow2)
plineObj.Layer = "K-grav"
plineObj.Update
plineObj.Closed = True

pointsarrow3(0) = points2(0) + 100:    pointsarrow3(1) = points2(1) + 98:
pointsarrow3(2) = points2(0) + 118:    pointsarrow3(3) = points2(1) + 96:
pointsarrow3(4) = points2(0) + 118:    pointsarrow3(5) = points2(1) + 100:

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsarrow3)
plineObj.Layer = "K-grav"
plineObj.Update
plineObj.Closed = True

'============================Arrows on the right============================
pointsarrow1(0) = points2(4) - 98:    pointsarrow1(1) = points2(1) + 100:
pointsarrow1(2) = points2(4) - 96:    pointsarrow1(3) = points2(1) + 118:
pointsarrow1(4) = points2(4) - 100:   pointsarrow1(5) = points2(1) + 118:

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsarrow1)
plineObj.Layer = "K-grav"
plineObj.Update
plineObj.Closed = True

pointsarrow2(0) = points2(4) - 99.4142:    pointsarrow2(1) = points2(1) + 99.4142:
pointsarrow2(2) = points2(4) - 110.7279:   pointsarrow2(3) = points2(1) + 113.5564:
pointsarrow2(4) = points2(4) - 113.5564:   pointsarrow2(5) = points2(1) + 110.7279:

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsarrow2)
plineObj.Layer = "K-grav"
plineObj.Update
plineObj.Closed = True

pointsarrow3(0) = points2(4) - 100:    pointsarrow3(1) = points2(1) + 98:
pointsarrow3(2) = points2(4) - 118:    pointsarrow3(3) = points2(1) + 96:
pointsarrow3(4) = points2(4) - 118:    pointsarrow3(5) = points2(1) + 100:

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsarrow3)
plineObj.Layer = "K-grav"
plineObj.Update
plineObj.Closed = True




End If
End If

  points2(1) = points2(3) + I:
  points2(3) = points2(3) + a + I:
  points2(5) = points2(3):
  points2(7) = points2(1):


Else: d = 1
d = d + 1

d = Cells(c, 4)
End If

Next d
c = c + 1

a = Cells(c, 2)
b = Cells(c, 3)

I = 100

Loop
              
ZoomAll
AP.Quit
End Sub

Sub MF166()
' Creating search objects
Dim AP As Excel.Application
Dim WB As Excel.Workbook
Dim WS As Excel.Worksheet

Set AP = Excel.Application
Set WB = AP.Workbooks.Open("C:\Users\������\Desktop\������\2.xlsx")
Set WS = WB.Worksheets("������ �˨������ ������")
c = 19
a = Cells(c, 2)
b = Cells(c, 3)
d = Cells(c, 4)
E = 1
I = 0
On Error Resume Next

' Description of all the points that will make up rectangles
 Dim plineObj As AcadLWPolyline
 Dim plineObjw As AcadLWPolyline
  Dim points(0 To 7) As Double
  Dim pointswithin(0 To 31) As Double
points(6) = 0

Do While a > 0

' Calculate the coordinates of each point
  points(0) = points(6) + I:    points(1) = 0
  points(2) = points(6) + I:    points(3) = a
  points(4) = points(0) + b:    points(5) = a
  points(6) = points(0) + b:    points(7) = 0

' Constructing four segments and connecting the corresponding points

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
  plineObj.Closed = True
 
  
pointswithin(0) = points(0) + 37:         pointswithin(1) = points(1) + 61:
pointswithin(2) = points(0) + 37:         pointswithin(3) = points(3) - 61:
pointswithin(4) = points(0) + 62:         pointswithin(5) = points(3) - 61:
pointswithin(6) = points(0) + 73.7805:    pointswithin(7) = points(3) - 54.1041:
pointswithin(8) = points(0) + 103:        pointswithin(9) = points(3) - 37:
pointswithin(10) = points(4) - 103:       pointswithin(11) = points(3) - 37:
pointswithin(12) = points(4) - 73.7805:   pointswithin(13) = points(3) - 54.1041:
pointswithin(14) = points(4) - 62:        pointswithin(15) = points(3) - 61:
pointswithin(16) = points(4) - 37:        pointswithin(17) = points(3) - 61:
pointswithin(18) = points(4) - 37:        pointswithin(19) = points(1) + 61:
pointswithin(20) = points(4) - 62:        pointswithin(21) = points(1) + 61:
pointswithin(22) = points(4) - 73.7805:   pointswithin(23) = points(1) + 54.1041:
pointswithin(24) = points(4) - 103:       pointswithin(25) = points(1) + 37:
pointswithin(26) = points(0) + 103:       pointswithin(27) = points(1) + 37:
pointswithin(28) = points(0) + 73.7805:   pointswithin(29) = points(1) + 54.1041:
pointswithin(30) = points(0) + 62:        pointswithin(31) = points(1) + 61:

If a >= 220 Then
If b >= 220 Then

plineObjw.Layer = "Ball-6"
plineObjw.Update
Set plineObjw = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointswithin)
  ' Find the bulge of the third segment
    Dim currentBulge As Double
    currentBulge = plineObjw.GetBulge(2)
  ' Change the bulge of the third segment
    plineObjw.SetBulge 2, 0.27116
    plineObjw.Update
    plineObjw.SetBulge 3, -0.27116
    plineObjw.Update
    plineObjw.SetBulge 5, -0.27116
    plineObjw.Update
    plineObjw.SetBulge 6, 0.27116
    plineObjw.Update
    plineObjw.SetBulge 10, 0.27116
    plineObjw.Update
    plineObjw.SetBulge 11, -0.27116
    plineObjw.Update
    plineObjw.SetBulge 13, -0.27116
    plineObjw.Update
    plineObjw.SetBulge 14, 0.27116
    plineObjw.Update
    plineObjw.Closed = True




  ' Offset the polyline
  Dim offsetObj As Variant
                                 'plineObj.color = acByLayer
                                 'Dim layerObj As AcadLayer
                                 'Set layerObj = ThisDrawing.Layers.Add("Ball-6.1")
                                 'layerObj.color = acRed

plineObjw.Layer = "K-grav"
plineObjw.Update
  offsetObj = plineObjw.Offset(10)
plineObjw.Layer = "K-Mill"
plineObjw.Update
  offsetObj = plineObjw.Offset(28)
  offsetObj = plineObjw.Offset(29)
  offsetObj = plineObjw.Offset(30)
plineObjw.Layer = "D-Mill"
plineObjw.Update
  offsetObj = plineObjw.Offset(34)
plineObjw.Layer = "Ball-6"
plineObjw.Update

End If
End If

I = 100
  
 Dim plineObj2 As AcadLWPolyline
  Dim points2(0 To 7) As Double
  
  points2(0) = points(0):    points2(1) = a + I:
  points2(2) = points(0):    points2(3) = (a * 2) + I:
  points2(4) = points(6):    points2(5) = (a * 2) + I:
  points2(6) = points(6):    points2(7) = a + I:
  
 For d = 2 To Cells(c, 4)
If d > 1 Then

Set plineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(points2)
  plineObj2.Closed = True
  

pointswithin(0) = points2(0) + 37:         pointswithin(1) = points2(1) + 61:
pointswithin(2) = points2(0) + 37:         pointswithin(3) = points2(3) - 61:
pointswithin(4) = points2(0) + 62:         pointswithin(5) = points2(3) - 61:
pointswithin(6) = points2(0) + 73.7805:    pointswithin(7) = points2(3) - 54.1041:
pointswithin(8) = points2(0) + 103:        pointswithin(9) = points2(3) - 37:
pointswithin(10) = points2(4) - 103:       pointswithin(11) = points2(3) - 37:
pointswithin(12) = points2(4) - 73.7805:   pointswithin(13) = points2(3) - 54.1041:
pointswithin(14) = points2(4) - 62:        pointswithin(15) = points2(3) - 61:
pointswithin(16) = points2(4) - 37:        pointswithin(17) = points2(3) - 61:
pointswithin(18) = points2(4) - 37:        pointswithin(19) = points2(1) + 61:
pointswithin(20) = points2(4) - 62:        pointswithin(21) = points2(1) + 61:
pointswithin(22) = points2(4) - 73.7805:   pointswithin(23) = points2(1) + 54.1041:
pointswithin(24) = points2(4) - 103:       pointswithin(25) = points2(1) + 37:
pointswithin(26) = points2(0) + 103:       pointswithin(27) = points2(1) + 37:
pointswithin(28) = points2(0) + 73.7805:   pointswithin(29) = points2(1) + 54.1041:
pointswithin(30) = points2(0) + 62:        pointswithin(31) = points2(1) + 61:

If a >= 220 Then
If b >= 220 Then

plineObjw.Layer = "Ball-6"
plineObjw.Update
Set plineObjw = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointswithin)
  ' Find the bulge of the third segment
 
    currentBulge = plineObjw.GetBulge(2)
  ' Change the bulge of the third segment
    plineObjw.SetBulge 2, 0.27116
    plineObjw.Update
    plineObjw.SetBulge 3, -0.27116
    plineObjw.Update
    plineObjw.SetBulge 5, -0.27116
    plineObjw.Update
    plineObjw.SetBulge 6, 0.27116
    plineObjw.Update
    plineObjw.SetBulge 10, 0.27116
    plineObjw.Update
    plineObjw.SetBulge 11, -0.27116
    plineObjw.Update
    plineObjw.SetBulge 13, -0.27116
    plineObjw.Update
    plineObjw.SetBulge 14, 0.27116
    plineObjw.Update
    plineObjw.Closed = True

  ' Offset the polyline

                                 'plineObj.color = acByLayer
                                 'Dim layerObj As AcadLayer
                                 'Set layerObj = ThisDrawing.Layers.Add("Ball-6.1")
                                 'layerObj.color = acRed

plineObjw.Layer = "K-grav"
plineObjw.Update
  offsetObj = plineObjw.Offset(10)
plineObjw.Layer = "K-Mill"
plineObjw.Update
  offsetObj = plineObjw.Offset(28)
  offsetObj = plineObjw.Offset(29)
  offsetObj = plineObjw.Offset(30)
plineObjw.Layer = "D-Mill"
plineObjw.Update
  offsetObj = plineObjw.Offset(34)
plineObjw.Layer = "Ball-6"
plineObjw.Update

End If
End If

  points2(1) = points2(3) + I:
  points2(3) = points2(3) + a + I:
  points2(5) = points2(3):
  points2(7) = points2(1):


Else: d = 1
d = d + 1

d = Cells(c, 4)
End If

Next d
c = c + 1

a = Cells(c, 2)
b = Cells(c, 3)

I = 100

Loop
              
ZoomAll
AP.Quit
End Sub

Sub MF167()
' Creating search objects
Dim AP As Excel.Application
Dim WB As Excel.Workbook
Dim WS As Excel.Worksheet

Set AP = Excel.Application
Set WB = AP.Workbooks.Open("C:\Users\������\Desktop\������\2.xlsx")
Set WS = WB.Worksheets("������ �˨������ ������")
c = 19
a = Cells(c, 2)
b = Cells(c, 3)
d = Cells(c, 4)
E = 1
I = 0
On Error Resume Next

' Description of all the points that will make up rectangles
 Dim plineObj As AcadLWPolyline
 Dim plineObjis1 As AcadLWPolyline
 Dim plineObjis2 As AcadLWPolyline
 Dim plineObjis3 As AcadLWPolyline
 Dim plineObjis4 As AcadLWPolyline
 Dim plineObjis5 As AcadLWPolyline
 Dim plineObjis6 As AcadLWPolyline
 Dim plineObjis7 As AcadLWPolyline
 Dim plineObjis8 As AcadLWPolyline
 Dim plineObjisw As AcadLWPolyline
  Dim points(0 To 7) As Double
  Dim pointswithin(0 To 7) As Double
  Dim pointsfill(0 To 15) As Double
  Dim pointsleaf1(0 To 11) As Double
  Dim pointsleaf2(0 To 11) As Double
  Dim pointsleaf3(0 To 11) As Double
  Dim pointsleaf4(0 To 11) As Double
  Dim pointsleaf5(0 To 11) As Double
  Dim pointsleaf6(0 To 11) As Double
  Dim pointsis1(0 To 3) As Double
  Dim pointsis2(0 To 3) As Double
  Dim pointsis3(0 To 3) As Double
  Dim pointsis4(0 To 3) As Double
  Dim pointsis5(0 To 3) As Double
  Dim pointsis6(0 To 3) As Double
  Dim pointsis7(0 To 3) As Double
  Dim pointsis8(0 To 3) As Double
  Dim P As Variant
points(6) = 0

Do While a > 0

' Calculate the coordinates of each point
  points(0) = points(6) + I:    points(1) = 0
  points(2) = points(6) + I:    points(3) = a
  points(4) = points(0) + b:    points(5) = a
  points(6) = points(0) + b:    points(7) = 0

' Constructing four segments and connecting the corresponding points

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
  plineObj.Closed = True
  plineObj.Layer = "0"
plineObj.Update
  
  ' Offset the polyline
  Dim offsetObj As Variant
                                 'plineObj.color = acByLayer
                                 'Dim layerObj As AcadLayer
                                 'Set layerObj = ThisDrawing.Layers.Add("Ball-6.1")
                                 'layerObj.color = acRed

If a >= 110 Then
If b >= 110 Then

plineObj.Layer = "Ball-6"
plineObj.Update
  offsetObj = plineObj.Offset(42)

plineObj.Layer = "0"
plineObj.Update


If a >= 150 Then
If b >= 146 Then


  pointswithin(0) = points(0) + 49:    pointswithin(1) = points(1) + 68
  pointswithin(2) = points(2) + 49:    pointswithin(3) = points(3) - 68
  pointswithin(4) = points(4) - 49:    pointswithin(5) = points(3) - 68
  pointswithin(6) = points(6) - 49:    pointswithin(7) = points(1) + 68

   
   Set plineObjw = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointswithin)
  ' Find the bulge of the third segment
    Dim currentBulge As Double
    currentBulge = plineObjw.GetBulge(2)
   k = (pointswithin(4) - pointswithin(2)) / 2
    ' Change the bulge of the third segment
    plineObjw.SetBulge 1, -(19 / k)
    plineObjw.Update
    plineObjw.SetBulge 3, -(19 / k)
    plineObjw.Update
  plineObjw.Closed = True
  plineObjw.Layer = "K-grav"
  plineObjw.Update

If a >= 198 Then
If b >= 198 Then
plineObjw.Layer = "C-Mill"
plineObjw.Update
  offsetObj = plineObjw.Offset(20)
If a >= 236 Then
If b >= 226 Then

 plineObjw.Layer = "Ball-6"
 plineObjw.Update
   offsetObj = plineObjw.Offset(50)

End If
End If

plineObjw.Layer = "K-grav"
plineObjw.Update

r = (0.25 * (pointswithin(4) - pointswithin(2)) ^ 2 / 19 + 19) / 2
r1 = r - 50
 ' Creating a circle
  Dim circleObj As AcadCircle
  Dim center(0 To 2) As Double
  Dim radius As Double
  center(0) = points(0) + (b / 2): center(1) = points(3) - r - 49: center(2) = 0: radius = r1
  Set circleObj = ThisDrawing.ModelSpace.AddCircle(center, radius)
  
  pointsis1(0) = points(0) + (b / 2) - 24:    pointsis1(1) = points(1) + (a / 2)
  pointsis1(2) = points(0) + (b / 2) - 24:    pointsis1(3) = points(3)
  pointsis2(0) = points(0) + (b / 2) + 24:    pointsis2(1) = points(1) + (a / 2)
  pointsis2(2) = points(0) + (b / 2) + 24:    pointsis2(3) = points(3)
  pointsis3(0) = points(0) + 99:              pointsis3(1) = points(1) + (a / 2)
  pointsis3(2) = points(0) + 99:              pointsis3(3) = points(3)
  pointsis4(0) = points(4) - 99:              pointsis4(1) = points(1) + (a / 2)
  pointsis4(2) = points(4) - 99:              pointsis4(3) = points(3)
  Set plineObjis1 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsis1)
  Set plineObjis2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsis2)
  Set plineObjis3 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsis3)
  Set plineObjis4 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsis4)
  
  
  ' Find the intersection points between the line and the circle
    Dim intPoints1 As Variant
    Dim intPoints2 As Variant
    Dim intPoints3 As Variant
    Dim intPoints4 As Variant
    intPoints1 = plineObjis1.IntersectWith(circleObj, acExtendNone)
    intPoints2 = plineObjis2.IntersectWith(circleObj, acExtendNone)
    intPoints3 = plineObjis3.IntersectWith(circleObj, acExtendNone)
    intPoints4 = plineObjis4.IntersectWith(circleObj, acExtendNone)
    
    circleObj.Delete
    
  Dim circleObj2 As AcadCircle
  center(0) = points(0) + (b / 2): center(1) = points(1) + r + 49: center(2) = 0: radius = r1
  Set circleObj2 = ThisDrawing.ModelSpace.AddCircle(center, radius)
  
  pointsis5(0) = points(0) + (b / 2) - 24:    pointsis5(1) = points(1)
  pointsis5(2) = points(0) + (b / 2) - 24:    pointsis5(3) = points(3) - (a / 2)
  pointsis6(0) = points(0) + (b / 2) + 24:    pointsis6(1) = points(1)
  pointsis6(2) = points(0) + (b / 2) + 24:    pointsis6(3) = points(3) - (a / 2)
  pointsis7(0) = points(0) + 99:              pointsis7(1) = points(1)
  pointsis7(2) = points(0) + 99:              pointsis7(3) = points(3) - (a / 2)
  pointsis8(0) = points(4) - 99:              pointsis8(1) = points(1)
  pointsis8(2) = points(4) - 99:              pointsis8(3) = points(3) - (a / 2)
  Set plineObjis5 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsis5)
  Set plineObjis6 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsis6)
  Set plineObjis7 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsis7)
  Set plineObjis8 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsis8)
    
    Dim intPoints5 As Variant
    Dim intPoints6 As Variant
    Dim intPoints7 As Variant
    Dim intPoints8 As Variant
    
    intPoints5 = plineObjis5.IntersectWith(circleObj2, acExtendNone)
    intPoints6 = plineObjis6.IntersectWith(circleObj2, acExtendNone)
    intPoints7 = plineObjis7.IntersectWith(circleObj2, acExtendNone)
    intPoints8 = plineObjis8.IntersectWith(circleObj2, acExtendNone)
  
    plineObjis1.Delete
    plineObjis2.Delete
    plineObjis3.Delete
    plineObjis4.Delete
    plineObjis5.Delete
    plineObjis6.Delete
    plineObjis7.Delete
    plineObjis8.Delete
      
    
    circleObj2.Delete
    
pointsfill(0) = points(0) + 99:  pointsfill(1) = intPoints8(1)
pointsfill(2) = points(0) + 99:  pointsfill(3) = intPoints4(1)
pointsfill(4) = intPoints1(0):   pointsfill(5) = intPoints1(1)
pointsfill(6) = intPoints2(0):   pointsfill(7) = intPoints2(1)
pointsfill(8) = points(4) - 99:  pointsfill(9) = intPoints3(1)
pointsfill(10) = points(4) - 99: pointsfill(11) = intPoints7(1)
pointsfill(12) = intPoints6(0):  pointsfill(13) = intPoints6(1)
pointsfill(14) = intPoints5(0):  pointsfill(15) = intPoints5(1)

If a >= 260 Then
If b >= 260 Then
Dim io As Integer
For io = 0 To UBound(offsetObj)
offsetObj(io).Delete
Next
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsfill)
plineObj.Closed = True

P = intPoints1(0) - intPoints4(0)
p1 = intPoints1(1) - intPoints4(1)
g = Sqr((P * P) + (p1 * p1))
   angle = Atn((p1 / g) / Sqr((-p1 / g) * (p1 / g) + 1))
   radius = (g / 2) / Sin(p1 / g)
   h = radius * (1 - Cos(angle))
   k = h / g

   currentBulge = plineObj.GetBulge(2)
   K1 = (intPoints6(0) - intPoints5(0)) / 2
   K2 = 124 - (points(3) - intPoints1(1))
    ' Change the bulge of the third segment
    plineObj.SetBulge 1, -2 * k
    plineObj.Update
    plineObj.SetBulge 2, (K2 / K1)
    plineObj.Update
    plineObj.SetBulge 3, -2 * k
    plineObj.Update
    plineObj.SetBulge 5, -2 * k
    plineObj.Update
    plineObj.SetBulge 6, (K2 / K1)
    plineObj.Update
    plineObj.SetBulge 7, -2 * k
    plineObj.Update
    plineObj.Layer = "Ball-6"
    plineObj.Update
  plineObj.Closed = True
'===============================Top pattern===============================================
  pointsleaf1(0) = points(0) + (b / 2) - 0.3166:    pointsleaf1(1) = points(3) - 116
  pointsleaf1(2) = points(0) + (b / 2) - 1.8092:    pointsleaf1(3) = pointsleaf1(1) + 7.75
  pointsleaf1(4) = points(0) + (b / 2) - 0.3166:    pointsleaf1(5) = pointsleaf1(3) + 7.75
  pointsleaf1(6) = points(0) + (b / 2) + 0.3166:    pointsleaf1(7) = pointsleaf1(5)
  pointsleaf1(8) = points(0) + (b / 2) + 1.8092:    pointsleaf1(9) = pointsleaf1(3)
  pointsleaf1(10) = points(0) + (b / 2) + 0.3166:   pointsleaf1(11) = pointsleaf1(1)

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsleaf1)
plineObj.Closed = True
plineObj.Layer = "K-grav_Pattern"
plineObj.Update

   pointsleaf2(0) = points(0) + (b / 2) - 2.7129:     pointsleaf2(1) = points(3) - 117.1026
   pointsleaf2(2) = points(0) + (b / 2) - 8.8379:     pointsleaf2(3) = pointsleaf2(1) + 4.9774
   pointsleaf2(4) = points(0) + (b / 2) - 12.6761:    pointsleaf2(5) = pointsleaf2(1) + 11.8773
   pointsleaf2(6) = points(0) + (b / 2) - 12.1911:    pointsleaf2(7) = pointsleaf2(1) + 12.2807
   pointsleaf2(8) = points(0) + (b / 2) - 6.0661:     pointsleaf2(9) = pointsleaf2(1) + 7.3033
   pointsleaf2(10) = points(0) + (b / 2) - 2.2278:    pointsleaf2(11) = pointsleaf2(1) + 0.407

 Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsleaf2)
 plineObj.Closed = True
 plineObj.Layer = "K-grav_Pattern"
 plineObj.Update

   pointsleaf3(0) = points(0) + (b / 2) + 2.7129:     pointsleaf3(1) = points(3) - 117.1026
   pointsleaf3(2) = points(0) + (b / 2) + 8.8379:     pointsleaf3(3) = pointsleaf2(1) + 4.9774
   pointsleaf3(4) = points(0) + (b / 2) + 12.6761:    pointsleaf3(5) = pointsleaf2(1) + 11.8773
   pointsleaf3(6) = points(0) + (b / 2) + 12.1911:    pointsleaf3(7) = pointsleaf2(1) + 12.2807
   pointsleaf3(8) = points(0) + (b / 2) + 6.0661:     pointsleaf3(9) = pointsleaf2(1) + 7.3033
   pointsleaf3(10) = points(0) + (b / 2) + 2.2278:    pointsleaf3(11) = pointsleaf2(1) + 0.407

 Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsleaf3)
 plineObj.Closed = True
 plineObj.Layer = "K-grav_Pattern"
 plineObj.Update
 
'===============================Bottom pattern===============================================
  pointsleaf4(0) = points(0) + (b / 2) - 0.3166:    pointsleaf4(1) = points(1) + 116
  pointsleaf4(2) = points(0) + (b / 2) - 1.8092:    pointsleaf4(3) = pointsleaf4(1) - 7.75
  pointsleaf4(4) = points(0) + (b / 2) - 0.3166:    pointsleaf4(5) = pointsleaf4(3) - 7.75
  pointsleaf4(6) = points(0) + (b / 2) + 0.3166:    pointsleaf4(7) = pointsleaf4(5)
  pointsleaf4(8) = points(0) + (b / 2) + 1.8092:    pointsleaf4(9) = pointsleaf4(3)
  pointsleaf4(10) = points(0) + (b / 2) + 0.3166:   pointsleaf4(11) = pointsleaf4(1)

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsleaf4)
plineObj.Closed = True
plineObj.Layer = "K-grav_Pattern"
plineObj.Update

   pointsleaf5(0) = points(0) + (b / 2) - 2.7129:     pointsleaf5(1) = points(1) + 117.1026
   pointsleaf5(2) = points(0) + (b / 2) - 8.8379:     pointsleaf5(3) = pointsleaf5(1) - 4.9774
   pointsleaf5(4) = points(0) + (b / 2) - 12.6761:    pointsleaf5(5) = pointsleaf5(1) - 11.8773
   pointsleaf5(6) = points(0) + (b / 2) - 12.1911:    pointsleaf5(7) = pointsleaf5(1) - 12.2807
   pointsleaf5(8) = points(0) + (b / 2) - 6.0661:     pointsleaf5(9) = pointsleaf5(1) - 7.3033
   pointsleaf5(10) = points(0) + (b / 2) - 2.2278:    pointsleaf5(11) = pointsleaf5(1) - 0.407

 Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsleaf5)
 plineObj.Closed = True
 plineObj.Layer = "K-grav_Pattern"
 plineObj.Update

   pointsleaf6(0) = points(0) + (b / 2) + 2.7129:     pointsleaf6(1) = points(1) + 117.1026
   pointsleaf6(2) = points(0) + (b / 2) + 8.8379:     pointsleaf6(3) = pointsleaf6(1) - 4.9774
   pointsleaf6(4) = points(0) + (b / 2) + 12.6761:    pointsleaf6(5) = pointsleaf6(1) - 11.8773
   pointsleaf6(6) = points(0) + (b / 2) + 12.1911:    pointsleaf6(7) = pointsleaf6(1) - 12.2807
   pointsleaf6(8) = points(0) + (b / 2) + 6.0661:     pointsleaf6(9) = pointsleaf6(1) - 7.3033
   pointsleaf6(10) = points(0) + (b / 2) + 2.2278:    pointsleaf6(11) = pointsleaf6(1) - 0.407

 Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsleaf6)
 plineObj.Closed = True
 plineObj.Layer = "K-grav_Pattern"
 plineObj.Update



End If
End If
End If
End If
End If
End If
End If
End If
   
I = 100
  
 Dim plineObj2 As AcadLWPolyline
  Dim points2(0 To 7) As Double
  
  points2(0) = points(0):    points2(1) = a + I:
  points2(2) = points(0):    points2(3) = (a * 2) + I:
  points2(4) = points(6):    points2(5) = (a * 2) + I:
  points2(6) = points(6):    points2(7) = a + I:
  
 For d = 2 To Cells(c, 4)
If d > 1 Then

Set plineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(points2)
  plineObj2.Closed = True
  
If a >= 110 Then
If b >= 110 Then

plineObj2.Layer = "Ball-6"
plineObj2.Update
  offsetObj = plineObj2.Offset(42)
plineObj2.Layer = "0"
plineObj2.Update

If a >= 150 Then
If b >= 146 Then

  pointswithin(0) = points2(0) + 49:    pointswithin(1) = points2(1) + 68
  pointswithin(2) = points2(2) + 49:    pointswithin(3) = points2(3) - 68
  pointswithin(4) = points2(4) - 49:    pointswithin(5) = points2(3) - 68
  pointswithin(6) = points2(6) - 49:    pointswithin(7) = points2(1) + 68

   
   Set plineObjw = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointswithin)
  ' Find the bulge of the third segment
    currentBulge = plineObjw.GetBulge(2)
   k = (pointswithin(4) - pointswithin(2)) / 2
    ' Change the bulge of the third segment
    plineObjw.SetBulge 1, -(19 / k)
    plineObjw.Update
    plineObjw.SetBulge 3, -(19 / k)
    plineObjw.Update
  plineObjw.Closed = True
  plineObjw.Layer = "K-grav"
  plineObjw.Update

If a >= 198 Then
If b >= 198 Then
plineObjw.Layer = "C-Mill"
plineObjw.Update
  offsetObj = plineObjw.Offset(20)
If a >= 236 Then
If b >= 226 Then
plineObjw.Layer = "Ball-6"
 plineObjw.Update
   offsetObj = plineObjw.Offset(50)

End If
End If
  
plineObjw.Layer = "K-grav"
plineObjw.Update

r = (0.25 * (pointswithin(4) - pointswithin(2)) ^ 2 / 19 + 19) / 2
r1 = r - 50
 ' Creating a circle
  center(0) = points2(0) + (b / 2): center(1) = points2(3) - r - 49: center(2) = 0: radius = r1
  Set circleObj = ThisDrawing.ModelSpace.AddCircle(center, radius)
  
  pointsis1(0) = points2(0) + (b / 2) - 24:    pointsis1(1) = points2(1) + (a / 2)
  pointsis1(2) = points2(0) + (b / 2) - 24:    pointsis1(3) = points2(3)
  pointsis2(0) = points2(0) + (b / 2) + 24:    pointsis2(1) = points2(1) + (a / 2)
  pointsis2(2) = points2(0) + (b / 2) + 24:    pointsis2(3) = points2(3)
  pointsis3(0) = points2(0) + 99:              pointsis3(1) = points2(1) + (a / 2)
  pointsis3(2) = points2(0) + 99:              pointsis3(3) = points2(3)
  pointsis4(0) = points2(4) - 99:              pointsis4(1) = points2(1) + (a / 2)
  pointsis4(2) = points2(4) - 99:              pointsis4(3) = points2(3)
  Set plineObjis1 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsis1)
  Set plineObjis2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsis2)
  Set plineObjis3 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsis3)
  Set plineObjis4 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsis4)
  
  
  ' Find the intersection points between the line and the circle
    intPoints1 = plineObjis1.IntersectWith(circleObj, acExtendNone)
    intPoints2 = plineObjis2.IntersectWith(circleObj, acExtendNone)
    intPoints3 = plineObjis3.IntersectWith(circleObj, acExtendNone)
    intPoints4 = plineObjis4.IntersectWith(circleObj, acExtendNone)
    
    circleObj.Delete
    
  center(0) = points2(0) + (b / 2): center(1) = points2(1) + r + 49: center(2) = 0: radius = r1
  Set circleObj2 = ThisDrawing.ModelSpace.AddCircle(center, radius)
  
  pointsis5(0) = points2(0) + (b / 2) - 24:    pointsis5(1) = points2(1)
  pointsis5(2) = points2(0) + (b / 2) - 24:    pointsis5(3) = points2(3) - (a / 2)
  pointsis6(0) = points2(0) + (b / 2) + 24:    pointsis6(1) = points2(1)
  pointsis6(2) = points2(0) + (b / 2) + 24:    pointsis6(3) = points2(3) - (a / 2)
  pointsis7(0) = points2(0) + 99:              pointsis7(1) = points2(1)
  pointsis7(2) = points2(0) + 99:              pointsis7(3) = points2(3) - (a / 2)
  pointsis8(0) = points2(4) - 99:              pointsis8(1) = points2(1)
  pointsis8(2) = points2(4) - 99:              pointsis8(3) = points2(3) - (a / 2)
  Set plineObjis5 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsis5)
  Set plineObjis6 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsis6)
  Set plineObjis7 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsis7)
  Set plineObjis8 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsis8)
    
    intPoints5 = plineObjis5.IntersectWith(circleObj2, acExtendNone)
    intPoints6 = plineObjis6.IntersectWith(circleObj2, acExtendNone)
    intPoints7 = plineObjis7.IntersectWith(circleObj2, acExtendNone)
    intPoints8 = plineObjis8.IntersectWith(circleObj2, acExtendNone)
  
    plineObjis1.Delete
    plineObjis2.Delete
    plineObjis3.Delete
    plineObjis4.Delete
    plineObjis5.Delete
    plineObjis6.Delete
    plineObjis7.Delete
    plineObjis8.Delete
      
    
    circleObj2.Delete
    
pointsfill(0) = points2(0) + 99:  pointsfill(1) = intPoints8(1)
pointsfill(2) = points2(0) + 99:  pointsfill(3) = intPoints4(1)
pointsfill(4) = intPoints1(0):   pointsfill(5) = intPoints1(1)
pointsfill(6) = intPoints2(0):   pointsfill(7) = intPoints2(1)
pointsfill(8) = points2(4) - 99:  pointsfill(9) = intPoints3(1)
pointsfill(10) = points2(4) - 99: pointsfill(11) = intPoints7(1)
pointsfill(12) = intPoints6(0):  pointsfill(13) = intPoints6(1)
pointsfill(14) = intPoints5(0):  pointsfill(15) = intPoints5(1)




If a >= 260 Then
If b >= 260 Then
For io = 0 To UBound(offsetObj)
offsetObj(io).Delete
Next
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsfill)
plineObj.Closed = True

P = intPoints1(0) - intPoints4(0)
p1 = intPoints1(1) - intPoints4(1)
g = Sqr((P * P) + (p1 * p1))
   angle = Atn((p1 / g) / Sqr((-p1 / g) * (p1 / g) + 1))
   radius = (g / 2) / Sin(p1 / g)
   h = radius * (1 - Cos(angle))
   k = h / g

   currentBulge = plineObj.GetBulge(2)
   K1 = (intPoints6(0) - intPoints5(0)) / 2
   K2 = 124 - (points2(3) - intPoints1(1))
    ' Change the bulge of the third segment
    plineObj.SetBulge 1, -2 * k
    plineObj.Update
    plineObj.SetBulge 2, (K2 / K1)
    plineObj.Update
    plineObj.SetBulge 3, -2 * k
    plineObj.Update
    plineObj.SetBulge 5, -2 * k
    plineObj.Update
    plineObj.SetBulge 6, (K2 / K1)
    plineObj.Update
    plineObj.SetBulge 7, -2 * k
    plineObj.Update
    plineObj.Layer = "Ball-6"
    plineObj.Update
  plineObj.Closed = True
'===============================Top pattern===============================================
  pointsleaf1(0) = points2(0) + (b / 2) - 0.3166:    pointsleaf1(1) = points2(3) - 116
  pointsleaf1(2) = points2(0) + (b / 2) - 1.8092:    pointsleaf1(3) = pointsleaf1(1) + 7.75
  pointsleaf1(4) = points2(0) + (b / 2) - 0.3166:    pointsleaf1(5) = pointsleaf1(3) + 7.75
  pointsleaf1(6) = points2(0) + (b / 2) + 0.3166:    pointsleaf1(7) = pointsleaf1(5)
  pointsleaf1(8) = points2(0) + (b / 2) + 1.8092:    pointsleaf1(9) = pointsleaf1(3)
  pointsleaf1(10) = points2(0) + (b / 2) + 0.3166:   pointsleaf1(11) = pointsleaf1(1)

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsleaf1)
plineObj.Closed = True
plineObj.Layer = "K-grav_Pattern"
plineObj.Update

   pointsleaf2(0) = points2(0) + (b / 2) - 2.7129:     pointsleaf2(1) = points2(3) - 117.1026
   pointsleaf2(2) = points2(0) + (b / 2) - 8.8379:     pointsleaf2(3) = pointsleaf2(1) + 4.9774
   pointsleaf2(4) = points2(0) + (b / 2) - 12.6761:    pointsleaf2(5) = pointsleaf2(1) + 11.8773
   pointsleaf2(6) = points2(0) + (b / 2) - 12.1911:    pointsleaf2(7) = pointsleaf2(1) + 12.2807
   pointsleaf2(8) = points2(0) + (b / 2) - 6.0661:     pointsleaf2(9) = pointsleaf2(1) + 7.3033
   pointsleaf2(10) = points2(0) + (b / 2) - 2.2278:    pointsleaf2(11) = pointsleaf2(1) + 0.407

 Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsleaf2)
 plineObj.Closed = True
 plineObj.Layer = "K-grav_Pattern"
 plineObj.Update

   pointsleaf3(0) = points2(0) + (b / 2) + 2.7129:     pointsleaf3(1) = points2(3) - 117.1026
   pointsleaf3(2) = points2(0) + (b / 2) + 8.8379:     pointsleaf3(3) = pointsleaf2(1) + 4.9774
   pointsleaf3(4) = points2(0) + (b / 2) + 12.6761:    pointsleaf3(5) = pointsleaf2(1) + 11.8773
   pointsleaf3(6) = points2(0) + (b / 2) + 12.1911:    pointsleaf3(7) = pointsleaf2(1) + 12.2807
   pointsleaf3(8) = points2(0) + (b / 2) + 6.0661:     pointsleaf3(9) = pointsleaf2(1) + 7.3033
   pointsleaf3(10) = points2(0) + (b / 2) + 2.2278:    pointsleaf3(11) = pointsleaf2(1) + 0.407

 Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsleaf3)
 plineObj.Closed = True
 plineObj.Layer = "K-grav_Pattern"
 plineObj.Update
 
'===============================Bottom pattern===============================================
  pointsleaf4(0) = points2(0) + (b / 2) - 0.3166:    pointsleaf4(1) = points2(1) + 116
  pointsleaf4(2) = points2(0) + (b / 2) - 1.8092:    pointsleaf4(3) = pointsleaf4(1) - 7.75
  pointsleaf4(4) = points2(0) + (b / 2) - 0.3166:    pointsleaf4(5) = pointsleaf4(3) - 7.75
  pointsleaf4(6) = points2(0) + (b / 2) + 0.3166:    pointsleaf4(7) = pointsleaf4(5)
  pointsleaf4(8) = points2(0) + (b / 2) + 1.8092:    pointsleaf4(9) = pointsleaf4(3)
  pointsleaf4(10) = points2(0) + (b / 2) + 0.3166:   pointsleaf4(11) = pointsleaf4(1)

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsleaf4)
plineObj.Closed = True
plineObj.Layer = "K-grav_Pattern"
plineObj.Update

   pointsleaf5(0) = points2(0) + (b / 2) - 2.7129:     pointsleaf5(1) = points2(1) + 117.1026
   pointsleaf5(2) = points2(0) + (b / 2) - 8.8379:     pointsleaf5(3) = pointsleaf5(1) - 4.9774
   pointsleaf5(4) = points2(0) + (b / 2) - 12.6761:    pointsleaf5(5) = pointsleaf5(1) - 11.8773
   pointsleaf5(6) = points2(0) + (b / 2) - 12.1911:    pointsleaf5(7) = pointsleaf5(1) - 12.2807
   pointsleaf5(8) = points2(0) + (b / 2) - 6.0661:     pointsleaf5(9) = pointsleaf5(1) - 7.3033
   pointsleaf5(10) = points2(0) + (b / 2) - 2.2278:    pointsleaf5(11) = pointsleaf5(1) - 0.407

 Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsleaf5)
 plineObj.Closed = True
 plineObj.Layer = "K-grav_Pattern"
 plineObj.Update

   pointsleaf6(0) = points2(0) + (b / 2) + 2.7129:     pointsleaf6(1) = points2(1) + 117.1026
   pointsleaf6(2) = points2(0) + (b / 2) + 8.8379:     pointsleaf6(3) = pointsleaf6(1) - 4.9774
   pointsleaf6(4) = points2(0) + (b / 2) + 12.6761:    pointsleaf6(5) = pointsleaf6(1) - 11.8773
   pointsleaf6(6) = points2(0) + (b / 2) + 12.1911:    pointsleaf6(7) = pointsleaf6(1) - 12.2807
   pointsleaf6(8) = points2(0) + (b / 2) + 6.0661:     pointsleaf6(9) = pointsleaf6(1) - 7.3033
   pointsleaf6(10) = points2(0) + (b / 2) + 2.2278:    pointsleaf6(11) = pointsleaf6(1) - 0.407

 Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsleaf6)
 plineObj.Closed = True
 plineObj.Layer = "K-grav_Pattern"
 plineObj.Update

End If
End If
End If
End If
End If
End If
End If
End If
  
  points2(1) = points2(3) + I:
  points2(3) = points2(3) + a + I:
  points2(5) = points2(3):
  points2(7) = points2(1):
  
  

Else: d = 1
d = d + 1

d = Cells(c, 4)
End If

Next d
c = c + 1

a = Cells(c, 2)
b = Cells(c, 3)

I = 100

Loop
              
ZoomAll
AP.Quit
End Sub

Sub MF168()
' Creating search objects
Dim AP As Excel.Application
Dim WB As Excel.Workbook
Dim WS As Excel.Worksheet

Set AP = Excel.Application
Set WB = AP.Workbooks.Open("C:\Users\������\Desktop\������\2.xlsx")
Set WS = WB.Worksheets("������ �˨������ ������")
c = 19
a = Cells(c, 2)
b = Cells(c, 3)
d = Cells(c, 4)
E = 1
I = 0
On Error Resume Next

' Description of all the points that will make up rectangles
 Dim plineObj As AcadLWPolyline
  Dim points(0 To 7) As Double
points(6) = 0

Do While a > 0

' Calculate the coordinates of each point
  points(0) = points(6) + I:    points(1) = 0
  points(2) = points(6) + I:    points(3) = a
  points(4) = points(0) + b:    points(5) = a
  points(6) = points(0) + b:    points(7) = 0

' Constructing four segments and connecting the corresponding points

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
  plineObj.Closed = True
  
Dim a1(0 To 2) As Double
Dim a2(0 To 2) As Double
Dim A3(0 To 2) As Double
Dim A4(0 To 2) As Double
Dim A5(0 To 2) As Double
Dim A6(0 To 2) As Double
Dim A7(0 To 2) As Double
Dim A8(0 To 2) As Double
Dim A9(0 To 2) As Double
Dim A10(0 To 2) As Double
Dim A11(0 To 2) As Double
Dim A12(0 To 2) As Double
Dim A13(0 To 2) As Double
Dim A14(0 To 2) As Double
Dim A15(0 To 2) As Double
Dim A16(0 To 2) As Double
Dim lineObj As AcadLine
  
  a1(0) = points(0):     a1(1) = 0:      a1(2) = 0
  a2(0) = points(0):     a2(1) = a:      a2(2) = 0
  A3(0) = points(6):     A3(1) = a:      A3(2) = 0
  A4(0) = points(6):     A4(1) = 0:      A4(2) = 0
  
  A5(0) = points(0) + 21.4142: A5(1) = points(1) + 21.4142:    A5(2) = 0
  A6(0) = points(2) + 21.4142: A6(1) = points(3) - 21.4142:    A6(2) = 0
  A7(0) = points(4) - 21.4142: A7(1) = points(5) - 21.4142:    A7(2) = 0
  A8(0) = points(6) - 21.4142: A8(1) = points(7) + 21.4142:    A8(2) = 0
  
  A9(0) = points(0) + 45:   A9(1) = points(1) + 45:    A9(2) = 0
  A10(0) = points(2) + 45: A10(1) = points(3) - 45:    A10(2) = 0
  A11(0) = points(4) - 45: A11(1) = points(5) - 45:    A11(2) = 0
  A12(0) = points(6) - 45: A12(1) = points(7) + 45:    A12(2) = 0
  
  A13(0) = points(0) + 55: A13(1) = points(1) + 55:    A13(2) = 0
  A14(0) = points(2) + 55: A14(1) = points(3) - 55:    A14(2) = 0
  A15(0) = points(4) - 55: A15(1) = points(5) - 55:    A15(2) = 0
  A16(0) = points(6) - 55: A16(1) = points(7) + 55:    A16(2) = 0

   
If a > 100 Then
If b > 100 Then

lineObj.Layer = "Ball-4"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A5, A9)
lineObj.Layer = "Ball-4"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A6, A10)
lineObj.Layer = "Ball-4"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A7, A11)
lineObj.Layer = "Ball-4"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A8, A12)
lineObj.Layer = "Ball-4"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A9, A13)
lineObj.Layer = "Ball-4"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A10, A14)
lineObj.Layer = "Ball-4"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A11, A15)
lineObj.Layer = "Ball-4"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A12, A16)
lineObj.Layer = "Ball-4"
lineObj.Update
 
  ' Offset the polyline
  Dim offsetObj As Variant
                                 'plineObj.color = acByLayer
                                 'Dim layerObj As AcadLayer
                                 'Set layerObj = ThisDrawing.Layers.Add("Ball-6.1")
                                 'layerObj.color = acRed

plineObj.Layer = "Ball-4"
plineObj.Update
  offsetObj = plineObj.Offset(20)
plineObj.Layer = "AreaClear"
plineObj.Update
  offsetObj = plineObj.Offset(20)
plineObj.Layer = "AreaClear"
plineObj.Update
  offsetObj = plineObj.Offset(45)
plineObj.Layer = "Ball-4"
plineObj.Update
  offsetObj = plineObj.Offset(55)
plineObj.Layer = "Koln-R5"
plineObj.Update
  offsetObj = plineObj.Offset(60)
plineObj.Layer = "Koln-R5"
plineObj.Update
  offsetObj = plineObj.Offset(65)
plineObj.Layer = "EndMill-12"
plineObj.Update
  offsetObj = plineObj.Offset(71)
  offsetObj = plineObj.Offset(73.5)
plineObj.Layer = "Ball-20"
plineObj.Update
  offsetObj = plineObj.Offset(79.5)
plineObj.Layer = "Koln-R5"
plineObj.Update
  offsetObj = plineObj.Offset(88.5)
  offsetObj = plineObj.Offset(89.5)
plineObj.Layer = "N-Mill"
plineObj.Update

End If
End If

I = 100
  
 Dim plineObj2 As AcadLWPolyline
  Dim points2(0 To 7) As Double
  
  points2(0) = points(0):    points2(1) = a + I:
  points2(2) = points(0):    points2(3) = (a * 2) + I:
  points2(4) = points(6):    points2(5) = (a * 2) + I:
  points2(6) = points(6):    points2(7) = a + I:
  
 For d = 2 To Cells(c, 4)
If d > 1 Then

Set plineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(points2)
  plineObj2.Closed = True
  
  a1(0) = points2(0):     a1(1) = points2(1):      a1(2) = 0
  a2(0) = points2(2):     a2(1) = points2(3):      a2(2) = 0
  A3(0) = points2(4):     A3(1) = points2(5):      A3(2) = 0
  A4(0) = points2(6):     A4(1) = points2(7):      A4(2) = 0
  
    
  A5(0) = points2(0) + 21.4142: A5(1) = points2(1) + 21.4142:    A5(2) = 0
  A6(0) = points2(2) + 21.4142: A6(1) = points2(3) - 21.4142:    A6(2) = 0
  A7(0) = points2(4) - 21.4142: A7(1) = points2(5) - 21.4142:    A7(2) = 0
  A8(0) = points2(6) - 21.4142: A8(1) = points2(7) + 21.4142:    A8(2) = 0
  
  A9(0) = points2(0) + 45:   A9(1) = points2(1) + 45:    A9(2) = 0
  A10(0) = points2(2) + 45: A10(1) = points2(3) - 45:    A10(2) = 0
  A11(0) = points2(4) - 45: A11(1) = points2(5) - 45:    A11(2) = 0
  A12(0) = points2(6) - 45: A12(1) = points2(7) + 45:    A12(2) = 0
  
  A13(0) = points2(0) + 55: A13(1) = points2(1) + 55:    A13(2) = 0
  A14(0) = points2(2) + 55: A14(1) = points2(3) - 55:    A14(2) = 0
  A15(0) = points2(4) - 55: A15(1) = points2(5) - 55:    A15(2) = 0
  A16(0) = points2(6) - 55: A16(1) = points2(7) + 55:    A16(2) = 0

   
If a > 100 Then
If b > 100 Then

lineObj.Layer = "Ball-4"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A5, A9)
lineObj.Layer = "Ball-4"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A6, A10)
lineObj.Layer = "Ball-4"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A7, A11)
lineObj.Layer = "Ball-4"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A8, A12)
lineObj.Layer = "Ball-4"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A9, A13)
lineObj.Layer = "Ball-4"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A10, A14)
lineObj.Layer = "Ball-4"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A11, A15)
lineObj.Layer = "Ball-4"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A12, A16)
lineObj.Layer = "Ball-4"
lineObj.Update
 
  ' Offset the polyline

                                 'plineObj.color = acByLayer
                                 'Dim layerObj As AcadLayer
                                 'Set layerObj = ThisDrawing.Layers.Add("Ball-6.1")
                                 'layerObj.color = acRed

plineObj2.Layer = "Ball-4"
plineObj2.Update
  offsetObj = plineObj2.Offset(20)
plineObj2.Layer = "AreaClear"
plineObj2.Update
  offsetObj = plineObj2.Offset(20)
plineObj2.Layer = "AreaClear"
plineObj2.Update
  offsetObj = plineObj2.Offset(45)
plineObj2.Layer = "Ball-4"
plineObj2.Update
  offsetObj = plineObj2.Offset(55)
plineObj2.Layer = "Koln-R5"
plineObj2.Update
  offsetObj = plineObj2.Offset(60)
plineObj2.Layer = "Koln-R5"
plineObj2.Update
  offsetObj = plineObj2.Offset(65)
plineObj2.Layer = "EndMill-12"
plineObj2.Update
  offsetObj = plineObj2.Offset(71)
  offsetObj = plineObj2.Offset(73.5)
plineObj2.Layer = "Ball-20"
plineObj2.Update
  offsetObj = plineObj2.Offset(79.5)
plineObj2.Layer = "Koln-R5"
plineObj2.Update
  offsetObj = plineObj2.Offset(88.5)
  offsetObj = plineObj2.Offset(89.5)
plineObj2.Layer = "N-Mill"
plineObj2.Update

End If
End If
  points2(1) = points2(3) + I:
  points2(3) = points2(3) + a + I:
  points2(5) = points2(3):
  points2(7) = points2(1):
  
 
Else: d = 1
d = d + 1

d = Cells(c, 4)
End If

Next d
c = c + 1

a = Cells(c, 2)
b = Cells(c, 3)

I = 100

Loop
              
ZoomAll
AP.Quit
End Sub

Sub MF174()
' Creating search objects
Dim AP As Excel.Application
Dim WB As Excel.Workbook
Dim WS As Excel.Worksheet

Set AP = Excel.Application
Set WB = AP.Workbooks.Open("C:\Users\������\Desktop\������\2.xlsx")
Set WS = WB.Worksheets("������ �˨������ ������")
c = 19
a = Cells(c, 2)
b = Cells(c, 3)
d = Cells(c, 4)
E = 1
I = 0
On Error Resume Next

' Description of all the points that will make up rectangles
  Dim plineObj As AcadLWPolyline
  Dim points(0 To 7) As Double
  Dim pointsendmill1(0 To 47) As Double
  Dim pointsendmill2(0 To 15) As Double
  Dim pointsta(0 To 5) As Double
  Dim basePoint(0 To 2) As Double
  Dim rotationAngle As Double
  Dim b01(0 To 2) As Double
  Dim b02(0 To 2) As Double
  Dim a01(0 To 2) As Double
  Dim a02(0 To 2) As Double
points(6) = 0

Do While a > 0

' Calculate the coordinates of each point
  points(0) = points(6) + I:    points(1) = 0
  points(2) = points(6) + I:    points(3) = a
  points(4) = points(0) + b:    points(5) = a
  points(6) = points(0) + b:    points(7) = 0

' Constructing four segments and connecting the corresponding points

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
  plineObj.Closed = True
 
Dim a1(0 To 2) As Double
Dim a2(0 To 2) As Double
Dim A3(0 To 2) As Double
Dim A4(0 To 2) As Double
Dim A5(0 To 2) As Double
Dim A6(0 To 2) As Double
Dim A7(0 To 2) As Double
Dim A8(0 To 2) As Double
Dim lineObj As AcadLine
  
  a1(0) = points(0) + 51: a1(1) = 0:      a1(2) = 0
  a2(0) = points(2) + 51: a2(1) = a:      a2(2) = 0
  A3(0) = points(4) - 51: A3(1) = a:      A3(2) = 0
  A4(0) = points(6) - 51: A4(1) = 0:      A4(2) = 0
  
  A5(0) = points(0) + 51: A5(1) = 51:      A5(2) = 0
  A6(0) = points(2) + 51: A6(1) = a - 51:  A6(2) = 0
  A7(0) = points(4) - 51: A7(1) = a - 51:  A7(2) = 0
  A8(0) = points(6) - 51: A8(1) = 51:      A8(2) = 0
   
If a > 100 Then
If b > 100 Then
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(a1, a2)
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A6, A7)
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A3, A4)
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A8, A5)
lineObj.Layer = "Ball-6"
lineObj.Update

  ' Offset the polyline
  Dim offsetObj As Variant
 
                                 'plineObj.color = acByLayer
                                 'Dim layerObj As AcadLayer
                                 'Set layerObj = ThisDrawing.Layers.Add("Ball-6.1")
                                 'layerObj.color = acRed
plineObj.Layer = "K-grav"
plineObj.Update
  offsetObj = plineObj.Offset(50)
  offsetObj = plineObj.Offset(57)
plineObj.Layer = "K-Mill"
plineObj.Update
  offsetObj = plineObj.Offset(73)
  offsetObj = plineObj.Offset(74)
  offsetObj = plineObj.Offset(75)
  offsetObj = plineObj.Offset(76)
  offsetObj = plineObj.Offset(77)
  offsetObj = plineObj.Offset(78)
plineObj.Layer = "AreaClear"
plineObj.Update
  offsetObj = plineObj.Offset(84)
plineObj.Layer = "0"
plineObj.Update


pointsendmill1(0) = points(0) + 6:       pointsendmill1(1) = points(1) + 42.35
pointsendmill1(2) = points(0) + 6:       pointsendmill1(3) = points(3) - 42.35
pointsendmill1(4) = points(0) + 7.42:    pointsendmill1(5) = points(3) - 38.48
pointsendmill1(6) = points(0) + 7.91:    pointsendmill1(7) = points(3) - 28.77
pointsendmill1(8) = points(0) + 28.77:   pointsendmill1(9) = points(3) - 7.91
pointsendmill1(10) = points(0) + 38.48:  pointsendmill1(11) = points(3) - 7.42
pointsendmill1(12) = points(0) + 42.35:  pointsendmill1(13) = points(3) - 6
pointsendmill1(14) = points(4) - 42.35:  pointsendmill1(15) = points(3) - 6
pointsendmill1(16) = points(4) - 38.48:  pointsendmill1(17) = points(3) - 7.42
pointsendmill1(18) = points(4) - 28.77:  pointsendmill1(19) = points(3) - 7.91
pointsendmill1(20) = points(4) - 7.91:   pointsendmill1(21) = points(3) - 28.77
pointsendmill1(22) = points(4) - 7.42:   pointsendmill1(23) = points(3) - 38.48
pointsendmill1(24) = points(4) - 6:      pointsendmill1(25) = points(3) - 42.35
pointsendmill1(26) = points(4) - 6:      pointsendmill1(27) = points(1) + 42.35
pointsendmill1(28) = points(4) - 7.42:   pointsendmill1(29) = points(1) + 38.48
pointsendmill1(30) = points(4) - 7.91:   pointsendmill1(31) = points(1) + 28.77
pointsendmill1(32) = points(4) - 28.77:  pointsendmill1(33) = points(1) + 7.91
pointsendmill1(34) = points(4) - 38.48:  pointsendmill1(35) = points(1) + 7.42
pointsendmill1(36) = points(4) - 42.35:  pointsendmill1(37) = points(1) + 6
pointsendmill1(38) = points(0) + 42.35:  pointsendmill1(39) = points(1) + 6
pointsendmill1(40) = points(0) + 38.48:  pointsendmill1(41) = points(1) + 7.42
pointsendmill1(42) = points(0) + 28.77:  pointsendmill1(43) = points(1) + 7.91
pointsendmill1(44) = points(0) + 7.91:   pointsendmill1(45) = points(1) + 28.77
pointsendmill1(46) = points(0) + 7.42:   pointsendmill1(47) = points(1) + 38.48


Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsendmill1)
  plineObj.Closed = True
  ' Find the bulge of the third segment
    currentBulge = plineObj.GetBulge(1)
    ' Change the bulge of the third segment
    plineObj.SetBulge 1, -0.17737
    plineObj.Update
    plineObj.SetBulge 2, 0.33851
    plineObj.Update
    plineObj.SetBulge 3, -0.8321
    plineObj.Update
    plineObj.SetBulge 4, 0.33851
    plineObj.Update
    plineObj.SetBulge 5, -0.17737
    plineObj.Update
    plineObj.SetBulge 7, -0.17737
    plineObj.Update
    plineObj.SetBulge 8, 0.33851
    plineObj.Update
    plineObj.SetBulge 9, -0.8321
    plineObj.Update
    plineObj.SetBulge 10, 0.33851
    plineObj.Update
    plineObj.SetBulge 11, -0.17737
    plineObj.Update
    plineObj.SetBulge 13, -0.17737
    plineObj.Update
    plineObj.SetBulge 14, 0.33851
    plineObj.Update
    plineObj.SetBulge 15, -0.8321
    plineObj.Update
    plineObj.SetBulge 16, 0.33851
    plineObj.Update
    plineObj.SetBulge 17, -0.17737
    plineObj.Update
    plineObj.SetBulge 19, -0.17737
    plineObj.Update
    plineObj.SetBulge 20, 0.33851
    plineObj.Update
    plineObj.SetBulge 21, -0.8321
    plineObj.Update
    plineObj.SetBulge 22, 0.33851
    plineObj.Update
    plineObj.SetBulge 23, -0.17737
    plineObj.Update
    plineObj.Layer = "EndMill-12"
    plineObj.Update
  plineObj.Closed = True

If a > 260 Then
If b > 260 Then

pointsendmill2(0) = points(0) + 100:       pointsendmill2(1) = points(1) + 125
pointsendmill2(2) = points(0) + 100:       pointsendmill2(3) = points(3) - 125
pointsendmill2(4) = points(0) + 124.5:     pointsendmill2(5) = points(3) - 105
pointsendmill2(6) = points(4) - 124.5:     pointsendmill2(7) = points(3) - 105
pointsendmill2(8) = points(4) - 100:       pointsendmill2(9) = points(3) - 125
pointsendmill2(10) = points(4) - 100:      pointsendmill2(11) = points(1) + 125
pointsendmill2(12) = points(4) - 124.5:    pointsendmill2(13) = points(1) + 105
pointsendmill2(14) = points(0) + 124.5:    pointsendmill2(15) = points(1) + 105

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsendmill2)
  plineObj.Closed = True

    k = (b - (124.5 * 2)) / 2
    plineObj.SetBulge 1, 0.35624
    plineObj.Update
    plineObj.SetBulge 2, -(5 / k)
    plineObj.Update
    plineObj.SetBulge 3, 0.35624
    plineObj.Update
    plineObj.SetBulge 5, 0.35624
    plineObj.Update
    plineObj.SetBulge 6, -(5 / k)
    plineObj.Update
    plineObj.SetBulge 7, 0.35624
    plineObj.Update
    plineObj.Layer = "EndMill-12"
    plineObj.Update
  plineObj.Closed = True

pointsta(0) = points(0) + 100.8638: pointsta(1) = points(1) + 104.8991
pointsta(2) = points(0) + 101.5206: pointsta(3) = points(1) + 120.4453
pointsta(4) = points(0) + 105.5638: pointsta(5) = points(1) + 119.7324
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsta)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  b01(0) = (b / 2) - 1:  b01(1) = (a / 2)
  b02(0) = (b / 2) + 1:  b02(1) = (a / 2)
  a01(0) = points(4) - (b / 2): a01(1) = points(1) + (a / 2) - 1
  a02(0) = points(4) - (b / 2): a02(1) = points(1) + (a / 2) + 1
  RetVal = plineObj.Mirror(b01, b02)
  RetVal = plineObj.Mirror(a01, a02)
  RetVal = plineObj.Copy
  ' Define the rotation of 180 degrees about a
  ' base point of (4, 4.25, 0)
  basePoint(0) = points(4) - (b / 2): basePoint(1) = points(3) - (a / 2)
  rotationAngle = 3.14159   ' 180 degrees
  plineObj.Rotate basePoint, rotationAngle
  plineObj.Update

pointsta(0) = points(0) + 112.9725: pointsta(1) = points(1) + 115.8756
pointsta(2) = points(0) + 103.5176: pointsta(3) = points(1) + 103.5176
pointsta(4) = points(0) + 115.8756: pointsta(5) = points(1) + 112.9725
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsta)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  b01(0) = (b / 2) - 1:  b01(1) = (a / 2)
  b02(0) = (b / 2) + 1:  b02(1) = (a / 2)
  a01(0) = points(4) - (b / 2): a01(1) = points(1) + (a / 2) - 1
  a02(0) = points(4) - (b / 2): a02(1) = points(1) + (a / 2) + 1
  RetVal = plineObj.Mirror(b01, b02)
  RetVal = plineObj.Mirror(a01, a02)
  RetVal = plineObj.Copy
  ' Define the rotation of 180 degrees about a
  ' base point of (4, 4.25, 0)
  basePoint(0) = points(4) - (b / 2): basePoint(1) = points(3) - (a / 2)
  rotationAngle = 3.14159   ' 180 degrees
  plineObj.Rotate basePoint, rotationAngle
  plineObj.Update

pointsta(0) = points(0) + 104.8991: pointsta(1) = points(1) + 100.8638
pointsta(2) = points(0) + 119.7924: pointsta(3) = points(1) + 105.5638
pointsta(4) = points(0) + 120.4453: pointsta(5) = points(1) + 101.5206
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsta)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  b01(0) = (b / 2) - 1:  b01(1) = (a / 2)
  b02(0) = (b / 2) + 1:  b02(1) = (a / 2)
  a01(0) = points(4) - (b / 2): a01(1) = points(1) + (a / 2) - 1
  a02(0) = points(4) - (b / 2): a02(1) = points(1) + (a / 2) + 1
  RetVal = plineObj.Mirror(b01, b02)
  RetVal = plineObj.Mirror(a01, a02)
  RetVal = plineObj.Copy
  ' Define the rotation of 180 degrees about a
  ' base point of (4, 4.25, 0)
  basePoint(0) = points(4) - (b / 2): basePoint(1) = points(3) - (a / 2)
  rotationAngle = 3.14159   ' 180 degrees
  plineObj.Rotate basePoint, rotationAngle
  plineObj.Update
End If
End If
End If
End If

I = 100
  
'========================================================================
  
 Dim plineObj2 As AcadLWPolyline
  Dim points2(0 To 7) As Double
  
  points2(0) = points(0):    points2(1) = a + I:
  points2(2) = points(0):    points2(3) = (a * 2) + I:
  points2(4) = points(6):    points2(5) = (a * 2) + I:
  points2(6) = points(6):    points2(7) = a + I:
  
 For d = 2 To Cells(c, 4)
If d > 1 Then

Set plineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(points2)
  plineObj2.Closed = True
If a > 100 Then
If b > 100 Then
  
  a1(0) = points2(0) + 51: a1(1) = points2(1):     a1(2) = 0
  a2(0) = points2(2) + 51: a2(1) = points2(3):     a2(2) = 0
  A3(0) = points2(4) - 51: A3(1) = points2(5):     A3(2) = 0
  A4(0) = points2(6) - 51: A4(1) = points2(7):     A4(2) = 0
  
  A5(0) = points2(0) + 51: A5(1) = points2(1) + 51:    A5(2) = 0
  A6(0) = points2(2) + 51: A6(1) = points2(3) - 51:    A6(2) = 0
  A7(0) = points2(4) - 51: A7(1) = points2(5) - 51:    A7(2) = 0
  A8(0) = points2(6) - 51: A8(1) = points2(7) + 51:    A8(2) = 0
  
     


lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(a1, a2)
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A6, A7)
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A3, A4)
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A5, A8)
lineObj.Layer = "Ball-6"
lineObj.Update

  
  ' Offset the polyline
plineObj2.Layer = "K-grav"
plineObj2.Update
  offsetObj = plineObj2.Offset(50)
  offsetObj = plineObj2.Offset(57)
plineObj2.Layer = "K-Mill"
plineObj2.Update
  offsetObj = plineObj2.Offset(73)
  offsetObj = plineObj2.Offset(74)
  offsetObj = plineObj2.Offset(75)
  offsetObj = plineObj2.Offset(76)
  offsetObj = plineObj2.Offset(77)
  offsetObj = plineObj2.Offset(78)
plineObj2.Layer = "AreaClear"
plineObj2.Update
  offsetObj = plineObj2.Offset(84)
plineObj2.Layer = "0"
plineObj2.Update

pointsendmill1(0) = points2(0) + 6:       pointsendmill1(1) = points2(1) + 42.35
pointsendmill1(2) = points2(0) + 6:       pointsendmill1(3) = points2(3) - 42.35
pointsendmill1(4) = points2(0) + 7.42:    pointsendmill1(5) = points2(3) - 38.48
pointsendmill1(6) = points2(0) + 7.91:    pointsendmill1(7) = points2(3) - 28.77
pointsendmill1(8) = points2(0) + 28.77:   pointsendmill1(9) = points2(3) - 7.91
pointsendmill1(10) = points2(0) + 38.48:  pointsendmill1(11) = points2(3) - 7.42
pointsendmill1(12) = points2(0) + 42.35:  pointsendmill1(13) = points2(3) - 6
pointsendmill1(14) = points2(4) - 42.35:  pointsendmill1(15) = points2(3) - 6
pointsendmill1(16) = points2(4) - 38.48:  pointsendmill1(17) = points2(3) - 7.42
pointsendmill1(18) = points2(4) - 28.77:  pointsendmill1(19) = points2(3) - 7.91
pointsendmill1(20) = points2(4) - 7.91:   pointsendmill1(21) = points2(3) - 28.77
pointsendmill1(22) = points2(4) - 7.42:   pointsendmill1(23) = points2(3) - 38.48
pointsendmill1(24) = points2(4) - 6:      pointsendmill1(25) = points2(3) - 42.35
pointsendmill1(26) = points2(4) - 6:      pointsendmill1(27) = points2(1) + 42.35
pointsendmill1(28) = points2(4) - 7.42:   pointsendmill1(29) = points2(1) + 38.48
pointsendmill1(30) = points2(4) - 7.91:   pointsendmill1(31) = points2(1) + 28.77
pointsendmill1(32) = points2(4) - 28.77:  pointsendmill1(33) = points2(1) + 7.91
pointsendmill1(34) = points2(4) - 38.48:  pointsendmill1(35) = points2(1) + 7.42
pointsendmill1(36) = points2(4) - 42.35:  pointsendmill1(37) = points2(1) + 6
pointsendmill1(38) = points2(0) + 42.35:  pointsendmill1(39) = points2(1) + 6
pointsendmill1(40) = points2(0) + 38.48:  pointsendmill1(41) = points2(1) + 7.42
pointsendmill1(42) = points2(0) + 28.77:  pointsendmill1(43) = points2(1) + 7.91
pointsendmill1(44) = points2(0) + 7.91:   pointsendmill1(45) = points2(1) + 28.77
pointsendmill1(46) = points2(0) + 7.42:   pointsendmill1(47) = points2(1) + 38.48


Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsendmill1)
  plineObj.Closed = True
  ' Find the bulge of the third segment
    currentBulge = plineObj.GetBulge(1)
    ' Change the bulge of the third segment
    plineObj.SetBulge 1, -0.17737
    plineObj.Update
    plineObj.SetBulge 2, 0.33851
    plineObj.Update
    plineObj.SetBulge 3, -0.8321
    plineObj.Update
    plineObj.SetBulge 4, 0.33851
    plineObj.Update
    plineObj.SetBulge 5, -0.17737
    plineObj.Update
    plineObj.SetBulge 7, -0.17737
    plineObj.Update
    plineObj.SetBulge 8, 0.33851
    plineObj.Update
    plineObj.SetBulge 9, -0.8321
    plineObj.Update
    plineObj.SetBulge 10, 0.33851
    plineObj.Update
    plineObj.SetBulge 11, -0.17737
    plineObj.Update
    plineObj.SetBulge 13, -0.17737
    plineObj.Update
    plineObj.SetBulge 14, 0.33851
    plineObj.Update
    plineObj.SetBulge 15, -0.8321
    plineObj.Update
    plineObj.SetBulge 16, 0.33851
    plineObj.Update
    plineObj.SetBulge 17, -0.17737
    plineObj.Update
    plineObj.SetBulge 19, -0.17737
    plineObj.Update
    plineObj.SetBulge 20, 0.33851
    plineObj.Update
    plineObj.SetBulge 21, -0.8321
    plineObj.Update
    plineObj.SetBulge 22, 0.33851
    plineObj.Update
    plineObj.SetBulge 23, -0.17737
    plineObj.Update
    plineObj.Layer = "EndMill-12"
    plineObj.Update
  plineObj.Closed = True

If a > 260 Then
If b > 260 Then

pointsendmill2(0) = points2(0) + 100:       pointsendmill2(1) = points2(1) + 125
pointsendmill2(2) = points2(0) + 100:       pointsendmill2(3) = points2(3) - 125
pointsendmill2(4) = points2(0) + 124.5:     pointsendmill2(5) = points2(3) - 105
pointsendmill2(6) = points2(4) - 124.5:     pointsendmill2(7) = points2(3) - 105
pointsendmill2(8) = points2(4) - 100:       pointsendmill2(9) = points2(3) - 125
pointsendmill2(10) = points2(4) - 100:      pointsendmill2(11) = points2(1) + 125
pointsendmill2(12) = points2(4) - 124.5:    pointsendmill2(13) = points2(1) + 105
pointsendmill2(14) = points2(0) + 124.5:    pointsendmill2(15) = points2(1) + 105

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsendmill2)
  plineObj.Closed = True

    k = (b - (124.5 * 2)) / 2
    plineObj.SetBulge 1, 0.35624
    plineObj.Update
    plineObj.SetBulge 2, -(5 / k)
    plineObj.Update
    plineObj.SetBulge 3, 0.35624
    plineObj.Update
    plineObj.SetBulge 5, 0.35624
    plineObj.Update
    plineObj.SetBulge 6, -(5 / k)
    plineObj.Update
    plineObj.SetBulge 7, 0.35624
    plineObj.Update
    plineObj.Layer = "EndMill-12"
    plineObj.Update
  plineObj.Closed = True

pointsta(0) = points2(0) + 100.8638: pointsta(1) = points2(1) + 104.8991
pointsta(2) = points2(0) + 101.5206: pointsta(3) = points2(1) + 120.4453
pointsta(4) = points2(0) + 105.5638: pointsta(5) = points2(1) + 119.7324
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsta)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  b01(0) = points2(0) + ((b / 2) - 1): b01(1) = points2(1) + (a / 2)
  b02(0) = points2(0) + ((b / 2) + 1): b02(1) = points2(1) + (a / 2)
  a01(0) = points2(4) - (b / 2): a01(1) = points2(1) + (a / 2) - 1
  a02(0) = points2(4) - (b / 2): a02(1) = points2(1) + (a / 2) + 1
  RetVal = plineObj.Mirror(b01, b02)
  RetVal = plineObj.Mirror(a01, a02)
  RetVal = plineObj.Copy
  ' Define the rotation of 180 degrees about a
  ' base point of (4, 4.25, 0)
  basePoint(0) = points2(4) - (b / 2): basePoint(1) = points2(3) - (a / 2)
  rotationAngle = 3.14159   ' 180 degrees
  plineObj.Rotate basePoint, rotationAngle
  plineObj.Update

pointsta(0) = points2(0) + 112.9725: pointsta(1) = points2(1) + 115.8756
pointsta(2) = points2(0) + 103.5176: pointsta(3) = points2(1) + 103.5176
pointsta(4) = points2(0) + 115.8756: pointsta(5) = points2(1) + 112.9725
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsta)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  b01(0) = points2(0) + ((b / 2) - 1): b01(1) = points2(1) + (a / 2)
  b02(0) = points2(0) + ((b / 2) + 1): b02(1) = points2(1) + (a / 2)
  a01(0) = points2(4) - (b / 2): a01(1) = points2(1) + (a / 2) - 1
  a02(0) = points2(4) - (b / 2): a02(1) = points2(1) + (a / 2) + 1
  RetVal = plineObj.Mirror(b01, b02)
  RetVal = plineObj.Mirror(a01, a02)
  RetVal = plineObj.Copy
  ' Define the rotation of 180 degrees about a
  ' base point of (4, 4.25, 0)
  basePoint(0) = points2(4) - (b / 2): basePoint(1) = points2(3) - (a / 2)
  rotationAngle = 3.14159   ' 180 degrees
  plineObj.Rotate basePoint, rotationAngle
  plineObj.Update

pointsta(0) = points2(0) + 104.8991: pointsta(1) = points2(1) + 100.8638
pointsta(2) = points2(0) + 119.7924: pointsta(3) = points2(1) + 105.5638
pointsta(4) = points2(0) + 120.4453: pointsta(5) = points2(1) + 101.5206
Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsta)
  plineObj.Closed = True
  plineObj.Layer = "K-grav"
  b01(0) = points2(0) + ((b / 2) - 1): b01(1) = points2(1) + (a / 2)
  b02(0) = points2(0) + ((b / 2) + 1): b02(1) = points2(1) + (a / 2)
  a01(0) = points2(4) - (b / 2): a01(1) = points2(1) + (a / 2) - 1
  a02(0) = points2(4) - (b / 2): a02(1) = points2(1) + (a / 2) + 1
  RetVal = plineObj.Mirror(b01, b02)
  RetVal = plineObj.Mirror(a01, a02)
  RetVal = plineObj.Copy
  ' Define the rotation of 180 degrees about a
  ' base point of (4, 4.25, 0)
  basePoint(0) = points2(4) - (b / 2): basePoint(1) = points2(3) - (a / 2)
  rotationAngle = 3.14159   ' 180 degrees
  plineObj.Rotate basePoint, rotationAngle
  plineObj.Update
  
End If
End If
End If
End If

  points2(1) = points2(3) + I:
  points2(3) = points2(3) + a + I:
  points2(5) = points2(3):
  points2(7) = points2(1):

Else: d = 1
d = d + 1

d = Cells(c, 4)
End If

Next d
c = c + 1

a = Cells(c, 2)
b = Cells(c, 3)

I = 100

Loop
              
ZoomAll
AP.Quit
End Sub

Sub MF175()
' Creating search objects
Dim AP As Excel.Application
Dim WB As Excel.Workbook
Dim WS As Excel.Worksheet

Set AP = Excel.Application
 Set WB = AP.Workbooks.Open("C:\Users\������\Desktop\������\2.xlsx")
Set WS = WB.Worksheets("������ �˨������ ������")
c = 19
a = Cells(c, 2)
b = Cells(c, 3)
d = Cells(c, 4)
E = 1
I = 0
On Error Resume Next

' Description of all the points that will make up rectangles
  Dim plineObj As AcadLWPolyline
  Dim plineObjw1 As AcadLWPolyline
  Dim plineObjw2 As AcadLWPolyline
  Dim points(0 To 7) As Double
  Dim pointswithin(0 To 13) As Double
  Dim pointswithin2(0 To 17) As Double
  Dim intPointsa
  Dim intPointsb
  Dim pointshelpa1(0 To 3) As Double
  Dim pointshelpa2(0 To 3) As Double
  Dim pointshelpb1(0 To 3) As Double
  Dim pointshelpb2(0 To 3) As Double
  Dim offsetObj As Variant
  Dim basePoint(0 To 2) As Double
  Dim rotationAngle As Double
  Dim b1(0 To 2) As Double
  Dim b2(0 To 2) As Double
  Dim a1(0 To 2) As Double
  Dim a2(0 To 2) As Double
  Dim l1(0 To 2) As Double
  Dim l2(0 To 2) As Double
  Dim l3(0 To 2) As Double
  Dim l4(0 To 2) As Double

points(6) = 0

Do While a > 0

' Calculate the coordinates of each point
  points(0) = points(6) + I:    points(1) = 0
  points(2) = points(6) + I:    points(3) = a
  points(4) = points(0) + b:    points(5) = a
  points(6) = points(0) + b:    points(7) = 0

' Constructing four segments and connecting the corresponding points

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
  plineObj.Closed = True

If a > 170 Then
If b >= 170 Then

  pointswithin(0) = points(0) + 66:                              pointswithin(1) = points(1) + 66
  pointswithin(2) = points(0) + 66:                              pointswithin(3) = points(3) - 90
  pointswithin(4) = points(0) + 66 + ((b - 132) / 4):            pointswithin(5) = points(3) - 78
  pointswithin(6) = points(0) + (b / 2):                         pointswithin(7) = points(3) - 66
  pointswithin(8) = points(4) - 66 - ((b - 132) / 4):            pointswithin(9) = points(3) - 78
  pointswithin(10) = points(4) - 66:                             pointswithin(11) = points(3) - 90
  pointswithin(12) = points(4) - 66:                             pointswithin(13) = points(1) + 66
   
If a > 170 Then
  

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointswithin)
  plineObj.Closed = True
   
   pb = (b - 132) / 4
   pa = (a - 132) / 4
ga = Sqr((pa * pa) + 144)
gb = Sqr((pb * pb) + 144)
   anglea = Atn((12 / ga) / Sqr((-12 / ga) * (12 / ga) + 1))
   angleb = Atn((12 / gb) / Sqr((-12 / gb) * (12 / gb) + 1))
   radiusa = (ga / 2) / Sin(12 / ga)
   radiusb = (gb / 2) / Sin(12 / gb)
   ha = radiusa * (1 - Cos(anglea))
   hb = radiusb * (1 - Cos(angleb))
   ka = ha / ga
   kb = hb / gb
    
    plineObj.SetBulge 1, kb * 2
    plineObj.SetBulge 2, -kb * 2
    plineObj.SetBulge 3, -kb * 2
    plineObj.SetBulge 4, kb * 2
    
    plineObj.Layer = "Help-line"
    plineObj.Update
    plineObj.Closed = True
    
    
plineObj.Layer = "Endmill-12"
plineObj.Update
  offsetObj = plineObj.Offset(4)
plineObj.Layer = "Ball-12"
plineObj.Update
  offsetObj = plineObj.Offset(10)
plineObj.Layer = "Ball-6"
plineObj.Update
  offsetObj = plineObj.Offset(16)
plineObj.Layer = "Help-line"
plineObj.Update

End If

  pointshelpb1(0) = points(0) + 66:      pointshelpb1(1) = points(3) - 86
  pointshelpb1(2) = points(0) + (b / 2): pointshelpb1(3) = points(3) - 62
  pointshelpb2(0) = points(0) + 66:      pointshelpb2(1) = points(3) - 90 + radiusb
  pointshelpb2(2) = pointswithin(4):    pointshelpb2(3) = pointswithin(5)
Set plineObjw1 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointshelpb1)
Set plineObjw2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointshelpb2)
intPointsb = plineObjw1.IntersectWith(plineObjw2, acExtendBoth)
Intersectionbx = intPointsb(0) - points(0)
Intersectionby = points(3) - intPointsb(1)
plineObjw1.Delete
plineObjw2.Delete
  
  pointswithin2(0) = points(0) + 62:                              pointswithin2(1) = points(1) + 62
  pointswithin2(2) = points(0) + 62:                              pointswithin2(3) = points(3) - 90
  pointswithin2(4) = points(0) + 66:                              pointswithin2(5) = points(3) - 86
  pointswithin2(6) = points(0) + Intersectionbx:                  pointswithin2(7) = points(3) - Intersectionby
  pointswithin2(8) = points(0) + (b / 2):                         pointswithin2(9) = points(3) - 62
  pointswithin2(10) = points(4) - Intersectionbx:                 pointswithin2(11) = points(3) - Intersectionby
  pointswithin2(12) = points(4) - 66:                             pointswithin2(13) = points(3) - 86
  pointswithin2(14) = points(4) - 62:                             pointswithin2(15) = points(3) - 90
  pointswithin2(16) = points(4) - 62:                             pointswithin2(17) = points(1) + 62
  

If a > 170 Then

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointswithin2)
  plineObj.Closed = True
    r = 4
    l = 2 * r * (Sqr(2) / 2)
    h = r * (1 - (Sqr(2) / 2))
    k = h / (l / 2)

    plineObj.SetBulge 1, -k
    plineObj.SetBulge 2, kb * 2
    plineObj.SetBulge 3, -kb * 2
    plineObj.SetBulge 4, -kb * 2
    plineObj.SetBulge 5, kb * 2
    plineObj.SetBulge 6, -k
  
    plineObj.Layer = "Endmill-12"
    plineObj.Update
    plineObj.Closed = True
    
    plineObj.Layer = "Ball-12"
plineObj.Update
  offsetObj = plineObj.Offset(-6)
plineObj.Layer = "Ball-6"
plineObj.Update
  offsetObj = plineObj.Offset(-12)
plineObj.Layer = "Endmill-12"
plineObj.Update
  
  l1(0) = points(0) + 50: l1(1) = points(1):     l1(2) = 0
  l2(0) = points(2) + 50: l2(1) = points(3):     l2(2) = 0
  l3(0) = points(4) - 50: l3(1) = points(5):     l3(2) = 0
  l4(0) = points(6) - 50: l4(1) = points(7):     l4(2) = 0
  
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(l1, l2)
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(l3, l4)
lineObj.Layer = "Ball-6"
lineObj.Update
  
  
End If
End If
End If
  
I = 100
  
 Dim plineObj2 As AcadLWPolyline
  Dim points2(0 To 7) As Double
  
  points2(0) = points(0):    points2(1) = a + I:
  points2(2) = points(0):    points2(3) = (a * 2) + I:
  points2(4) = points(6):    points2(5) = (a * 2) + I:
  points2(6) = points(6):    points2(7) = a + I:
  
'========================================================
'========================================================
'========================================================
  
  
 For d = 2 To Cells(c, 4)
If d > 1 Then

Set plineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(points2)
  plineObj2.Closed = True
  
If a > 170 Then
If b >= 170 Then
  
  pointswithin(0) = points2(0) + 66:                              pointswithin(1) = points2(1) + 66
  pointswithin(2) = points2(0) + 66:                              pointswithin(3) = points2(3) - 90
  pointswithin(4) = points2(0) + 66 + ((b - 132) / 4):            pointswithin(5) = points2(3) - 78
  pointswithin(6) = points2(0) + (b / 2):                         pointswithin(7) = points2(3) - 66
  pointswithin(8) = points2(4) - 66 - ((b - 132) / 4):            pointswithin(9) = points2(3) - 78
  pointswithin(10) = points2(4) - 66:                             pointswithin(11) = points2(3) - 90
  pointswithin(12) = points2(4) - 66:                             pointswithin(13) = points2(1) + 66
   
If a > 170 Then
  

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointswithin)
  plineObj.Closed = True
   
   pb = (b - 132) / 4
   pa = (a - 132) / 4
ga = Sqr((pa * pa) + 144)
gb = Sqr((pb * pb) + 144)
   anglea = Atn((12 / ga) / Sqr((-12 / ga) * (12 / ga) + 1))
   angleb = Atn((12 / gb) / Sqr((-12 / gb) * (12 / gb) + 1))
   radiusa = (ga / 2) / Sin(12 / ga)
   radiusb = (gb / 2) / Sin(12 / gb)
   ha = radiusa * (1 - Cos(anglea))
   hb = radiusb * (1 - Cos(angleb))
   ka = ha / ga
   kb = hb / gb
    
    plineObj.SetBulge 1, kb * 2
    plineObj.SetBulge 2, -kb * 2
    plineObj.SetBulge 3, -kb * 2
    plineObj.SetBulge 4, kb * 2
    
    plineObj.Layer = "Help-line"
    plineObj.Update
    plineObj.Closed = True
    
    
plineObj.Layer = "Endmill-12"
plineObj.Update
  offsetObj = plineObj.Offset(4)
plineObj.Layer = "Ball-12"
plineObj.Update
  offsetObj = plineObj.Offset(10)
plineObj.Layer = "Ball-6"
plineObj.Update
  offsetObj = plineObj.Offset(16)
plineObj.Layer = "Help-line"
plineObj.Update

End If

  pointshelpb1(0) = points2(0) + 66:      pointshelpb1(1) = points2(3) - 86
  pointshelpb1(2) = points2(0) + (b / 2): pointshelpb1(3) = points2(3) - 62
  pointshelpb2(0) = points2(0) + 66:      pointshelpb2(1) = points2(3) - 90 + radiusb
  pointshelpb2(2) = pointswithin(4):    pointshelpb2(3) = pointswithin(5)
Set plineObjw1 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointshelpb1)
Set plineObjw2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointshelpb2)
intPointsb = plineObjw1.IntersectWith(plineObjw2, acExtendBoth)
Intersectionbx = intPointsb(0) - points2(0)
Intersectionby = points2(3) - intPointsb(1)
plineObjw1.Delete
plineObjw2.Delete
  
  pointswithin2(0) = points2(0) + 62:                              pointswithin2(1) = points2(1) + 62
  pointswithin2(2) = points2(0) + 62:                              pointswithin2(3) = points2(3) - 90
  pointswithin2(4) = points2(0) + 66:                              pointswithin2(5) = points2(3) - 86
  pointswithin2(6) = points2(0) + Intersectionbx:                  pointswithin2(7) = points2(3) - Intersectionby
  pointswithin2(8) = points2(0) + (b / 2):                         pointswithin2(9) = points2(3) - 62
  pointswithin2(10) = points2(4) - Intersectionbx:                 pointswithin2(11) = points2(3) - Intersectionby
  pointswithin2(12) = points2(4) - 66:                             pointswithin2(13) = points2(3) - 86
  pointswithin2(14) = points2(4) - 62:                             pointswithin2(15) = points2(3) - 90
  pointswithin2(16) = points2(4) - 62:                             pointswithin2(17) = points2(1) + 62
  

If a > 170 Then

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointswithin2)
  plineObj.Closed = True
    r = 4
    l = 2 * r * (Sqr(2) / 2)
    h = r * (1 - (Sqr(2) / 2))
    k = h / (l / 2)

    plineObj.SetBulge 1, -k
    plineObj.SetBulge 2, kb * 2
    plineObj.SetBulge 3, -kb * 2
    plineObj.SetBulge 4, -kb * 2
    plineObj.SetBulge 5, kb * 2
    plineObj.SetBulge 6, -k
  
    plineObj.Layer = "Endmill-12"
    plineObj.Update
    plineObj.Closed = True
    
    plineObj.Layer = "Ball-12"
plineObj.Update
  offsetObj = plineObj.Offset(-6)
plineObj.Layer = "Ball-6"
plineObj.Update
  offsetObj = plineObj.Offset(-12)
plineObj.Layer = "Endmill-12"
plineObj.Update
  
  l1(0) = points2(0) + 50: l1(1) = points2(1):     l1(2) = 0
  l2(0) = points2(2) + 50: l2(1) = points2(3):     l2(2) = 0
  l3(0) = points2(4) - 50: l3(1) = points2(5):     l3(2) = 0
  l4(0) = points2(6) - 50: l4(1) = points2(7):     l4(2) = 0
  
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(l1, l2)
lineObj.Layer = "Ball-6"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(l3, l4)
lineObj.Layer = "Ball-6"
lineObj.Update
  
End If
End If
End If
  
  points2(1) = points2(3) + I:
  points2(3) = points2(3) + a + I:
  points2(5) = points2(3):
  points2(7) = points2(1):
  

Else: d = 1
d = d + 1

d = Cells(c, 4)
End If

Next d
c = c + 1

a = Cells(c, 2)
b = Cells(c, 3)

I = 100

Loop
              
ZoomAll
AP.Quit
End Sub

Sub MF177()
' Creating search objects
Dim AP As Excel.Application
Dim WB As Excel.Workbook
Dim WS As Excel.Worksheet

Set AP = Excel.Application
Set WB = AP.Workbooks.Open("C:\Users\������\Desktop\������\2.xlsx")
Set WS = WB.Worksheets("������ �˨������ ������")
c = 19
a = Cells(c, 2)
b = Cells(c, 3)
d = Cells(c, 4)
E = 1
I = 0
On Error Resume Next

' Description of all the points that will make up rectangles
 Dim plineObj As AcadLWPolyline
  Dim points(0 To 7) As Double
points(6) = 0

Do While a > 0

' Calculate the coordinates of each point
  points(0) = points(6) + I:    points(1) = 0
  points(2) = points(6) + I:    points(3) = a
  points(4) = points(0) + b:    points(5) = a
  points(6) = points(0) + b:    points(7) = 0

' Constructing four segments and connecting the corresponding points

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
  plineObj.Closed = True
  
Dim a1(0 To 2) As Double
Dim a2(0 To 2) As Double
Dim A3(0 To 2) As Double
Dim A4(0 To 2) As Double
Dim A5(0 To 2) As Double
Dim A6(0 To 2) As Double
Dim A7(0 To 2) As Double
Dim A8(0 To 2) As Double
Dim lineObj As AcadLine
  
  a1(0) = points(0):     a1(1) = 0:      a1(2) = 0
  a2(0) = points(0):     a2(1) = a:      a2(2) = 0
  A3(0) = points(6):     A3(1) = a:      A3(2) = 0
  A4(0) = points(6):     A4(1) = 0:      A4(2) = 0
  
  A5(0) = points(0) + 51: A5(1) = points(1) + 51:    A5(2) = 0
  A6(0) = points(2) + 51: A6(1) = points(3) - 51:    A6(2) = 0
  A7(0) = points(4) - 51: A7(1) = points(5) - 51:    A7(2) = 0
  A8(0) = points(6) - 51: A8(1) = points(7) + 51:    A8(2) = 0

   
If a > 100 Then
If b > 100 Then

lineObj.Layer = "K-Mill"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(a1, A5)
lineObj.Layer = "K-Mill"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(a2, A6)
lineObj.Layer = "K-Mill"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A3, A7)
lineObj.Layer = "K-Mill"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A4, A8)
 
  ' Offset the polyline
  Dim offsetObj As Variant
                                 'plineObj.color = acByLayer
                                 'Dim layerObj As AcadLayer
                                 'Set layerObj = ThisDrawing.Layers.Add("Ball-6.1")
                                 'layerObj.color = acRed

plineObj.Layer = "K-grav"
plineObj.Update
  offsetObj = plineObj.Offset(45)
plineObj.Layer = "K-Mill"
plineObj.Update
  offsetObj = plineObj.Offset(51)
plineObj.Layer = "K-grav"
plineObj.Update
  offsetObj = plineObj.Offset(53)
plineObj.Layer = "K-Mill"
plineObj.Update
  offsetObj = plineObj.Offset(53)
  offsetObj = plineObj.Offset(55)
  offsetObj = plineObj.Offset(57)
  offsetObj = plineObj.Offset(66)
  offsetObj = plineObj.Offset(68)
  offsetObj = plineObj.Offset(70)
plineObj.Layer = "D-Mill"
plineObj.Update
  offsetObj = plineObj.Offset(74)
plineObj.Layer = "0"
plineObj.Update

End If
End If

I = 100
  
 Dim plineObj2 As AcadLWPolyline
  Dim points2(0 To 7) As Double
  
  points2(0) = points(0):    points2(1) = a + I:
  points2(2) = points(0):    points2(3) = (a * 2) + I:
  points2(4) = points(6):    points2(5) = (a * 2) + I:
  points2(6) = points(6):    points2(7) = a + I:
  
 For d = 2 To Cells(c, 4)
If d > 1 Then

Set plineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(points2)
  plineObj2.Closed = True
  
  a1(0) = points2(0):     a1(1) = points2(1):      a1(2) = 0
  a2(0) = points2(2):     a2(1) = points2(3):      a2(2) = 0
  A3(0) = points2(4):     A3(1) = points2(5):      A3(2) = 0
  A4(0) = points2(6):     A4(1) = points2(7):      A4(2) = 0
  
  A5(0) = points2(0) + 51: A5(1) = points2(1) + 51:    A5(2) = 0
  A6(0) = points2(2) + 51: A6(1) = points2(3) - 51:    A6(2) = 0
  A7(0) = points2(4) - 51: A7(1) = points2(5) - 51:    A7(2) = 0
  A8(0) = points2(6) - 51: A8(1) = points2(7) + 51:    A8(2) = 0
   
If a > 100 Then
If b > 100 Then

lineObj.Layer = "K-Mill"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(a1, A5)
lineObj.Layer = "K-Mill"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(a2, A6)
lineObj.Layer = "K-Mill"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A3, A7)
lineObj.Layer = "K-Mill"
lineObj.Update
Set lineObj = ThisDrawing.ModelSpace.AddLine(A4, A8)
lineObj.Layer = "K-Mill"
lineObj.Update

  ' Offset the polyline

plineObj2.Layer = "K-grav"
plineObj2.Update
  offsetObj = plineObj2.Offset(45)
plineObj2.Layer = "K-Mill"
plineObj2.Update
  offsetObj = plineObj2.Offset(51)
plineObj2.Layer = "K-grav"
plineObj2.Update
  offsetObj = plineObj2.Offset(53)
plineObj2.Layer = "K-Mill"
plineObj2.Update
  offsetObj = plineObj2.Offset(53)
  offsetObj = plineObj2.Offset(55)
  offsetObj = plineObj2.Offset(57)
  offsetObj = plineObj2.Offset(66)
  offsetObj = plineObj2.Offset(68)
  offsetObj = plineObj2.Offset(70)
plineObj2.Layer = "D-Mill"
plineObj2.Update
  offsetObj = plineObj2.Offset(74)
plineObj2.Layer = "0"
plineObj2.Update

End If
End If

  points2(1) = points2(3) + I:
  points2(3) = points2(3) + a + I:
  points2(5) = points2(3):
  points2(7) = points2(1):
  
Else: d = 1
d = d + 1

d = Cells(c, 4)
End If

Next d
c = c + 1

a = Cells(c, 2)
b = Cells(c, 3)

I = 100

Loop
              
ZoomAll
AP.Quit
End Sub

Sub MF180()
' Creating search objects
Dim AP As Excel.Application
Dim WB As Excel.Workbook
Dim WS As Excel.Worksheet

Set AP = Excel.Application
Set WB = AP.Workbooks.Open("C:\Users\������\Desktop\������\2.xlsx")
Set WS = WB.Worksheets("������ �˨������ ������")
c = 19
a = Cells(c, 2)
b = Cells(c, 3)
d = Cells(c, 4)
E = 1
I = 0
On Error Resume Next

' Description of all the points that will make up rectangles
 Dim plineObj As AcadLWPolyline
  Dim points(0 To 7) As Double
points(6) = 0

Do While a > 0

' Calculate the coordinates of each point
  points(0) = points(6) + I:    points(1) = 0
  points(2) = points(6) + I:    points(3) = a
  points(4) = points(0) + b:    points(5) = a
  points(6) = points(0) + b:    points(7) = 0

' Constructing four segments and connecting the corresponding points

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
  plineObj.Closed = True
  
  ' Offset the polyline
Dim offsetObj As Variant
                                 'plineObj.color = acByLayer
                                 'Dim layerObj As AcadLayer
                                 'Set layerObj = ThisDrawing.Layers.Add("Ball-6.1")
                                 'layerObj.color = acRed

plineObj.Layer = "AreaClear"
plineObj.Update
  offsetObj = plineObj.Offset(51)
plineObj.Layer = "EndMill-12"
plineObj.Update
  offsetObj = plineObj.Offset(56)
plineObj.Layer = "0"
plineObj.Update

I = 100
  
 Dim plineObj2 As AcadLWPolyline
  Dim points2(0 To 7) As Double
  
  points2(0) = points(0):    points2(1) = a + I:
  points2(2) = points(0):    points2(3) = (a * 2) + I:
  points2(4) = points(6):    points2(5) = (a * 2) + I:
  points2(6) = points(6):    points2(7) = a + I:
  
 For d = 2 To Cells(c, 4)
If d > 1 Then

Set plineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(points2)
  plineObj2.Closed = True
   
  points2(1) = points2(3) + I:
  points2(3) = points2(3) + a + I:
  points2(5) = points2(3):
  points2(7) = points2(1):
  
  ' Offset the polyline

plineObj2.Layer = "AreaClear"
plineObj2.Update
  offsetObj = plineObj2.Offset(51)
plineObj2.Layer = "EndMill-12"
plineObj2.Update
  offsetObj = plineObj2.Offset(56)
plineObj2.Layer = "0"
plineObj2.Update

Else: d = 1
d = d + 1

d = Cells(c, 4)
End If

Next d
c = c + 1

a = Cells(c, 2)
b = Cells(c, 3)

I = 100

Loop
              
ZoomAll
AP.Quit
End Sub

Sub MF182()
' Creating search objects
Dim AP As Excel.Application
Dim WB As Excel.Workbook
Dim WS As Excel.Worksheet

Set AP = Excel.Application
Set WB = AP.Workbooks.Open("C:\Users\������\Desktop\������\2.xlsx")
Set WS = WB.Worksheets("������ �˨������ ������")
c = 19
a = Cells(c, 2)
b = Cells(c, 3)
d = Cells(c, 4)
E = 1
I = 0
On Error Resume Next

' Description of all the points that will make up rectangles
 Dim plineObj As AcadLWPolyline
  Dim points(0 To 7) As Double
points(6) = 0

Do While a > 0

' Calculate the coordinates of each point
  points(0) = points(6) + I:    points(1) = 0
  points(2) = points(6) + I:    points(3) = a
  points(4) = points(0) + b:    points(5) = a
  points(6) = points(0) + b:    points(7) = 0

' Constructing four segments and connecting the corresponding points

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
  plineObj.Closed = True
  
  ' Offset the polyline
Dim offsetObj As Variant
                                 'plineObj.color = acByLayer
                                 'Dim layerObj As AcadLayer
                                 'Set layerObj = ThisDrawing.Layers.Add("Ball-6.1")
                                 'layerObj.color = acRed

plineObj.Layer = "K-Mill"
plineObj.Update
  offsetObj = plineObj.Offset(52)
  offsetObj = plineObj.Offset(53)
  offsetObj = plineObj.Offset(54)
plineObj.Layer = "K-grav"
plineObj.Update
  offsetObj = plineObj.Offset(53)
plineObj.Layer = "AreaClear"
plineObj.Update
  offsetObj = plineObj.Offset(63)
plineObj.Layer = "0"
plineObj.Update

I = 100
  
 Dim plineObj2 As AcadLWPolyline
  Dim points2(0 To 7) As Double
  
  points2(0) = points(0):    points2(1) = a + I:
  points2(2) = points(0):    points2(3) = (a * 2) + I:
  points2(4) = points(6):    points2(5) = (a * 2) + I:
  points2(6) = points(6):    points2(7) = a + I:
  
 For d = 2 To Cells(c, 4)
If d > 1 Then

Set plineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(points2)
  plineObj2.Closed = True
  
  ' Offset the polyline

plineObj2.Layer = "K-Mill"
plineObj2.Update
  offsetObj = plineObj2.Offset(52)
  offsetObj = plineObj2.Offset(53)
  offsetObj = plineObj2.Offset(54)
plineObj2.Layer = "K-grav"
plineObj2.Update
  offsetObj = plineObj2.Offset(53)
plineObj2.Layer = "AreaClear"

plineObj2.Update
  offsetObj = plineObj2.Offset(63)
plineObj2.Update
plineObj2.Layer = "0"
plineObj2.Update

  points2(1) = points2(3) + I:
  points2(3) = points2(3) + a + I:
  points2(5) = points2(3):
  points2(7) = points2(1):

Else: d = 1
d = d + 1

d = Cells(c, 4)
End If

Next d
c = c + 1

a = Cells(c, 2)
b = Cells(c, 3)

I = 100

Loop
              
ZoomAll
AP.Quit
End Sub

Sub MF184()
' Creating search objects
Dim AP As Excel.Application
Dim WB As Excel.Workbook
Dim WS As Excel.Worksheet

Set AP = Excel.Application
Set WB = AP.Workbooks.Open("C:\Users\������\Desktop\������\2.xlsx")
Set WS = WB.Worksheets("������ �˨������ ������")
c = 19
a = Cells(c, 2)
b = Cells(c, 3)
d = Cells(c, 4)
E = 1
I = 0
On Error Resume Next

' Description of all the points that will make up rectangles
 Dim plineObj As AcadLWPolyline
  Dim points(0 To 7) As Double
points(6) = 0

Do While a > 0

' Calculate the coordinates of each point
  points(0) = points(6) + I:    points(1) = 0
  points(2) = points(6) + I:    points(3) = a
  points(4) = points(0) + b:    points(5) = a
  points(6) = points(0) + b:    points(7) = 0

' Constructing four segments and connecting the corresponding points

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
  plineObj.Closed = True
  
  ' Offset the polyline
Dim offsetObj As Variant
                                 'plineObj.color = acByLayer
                                 'Dim layerObj As AcadLayer
                                 'Set layerObj = ThisDrawing.Layers.Add("Ball-6.1")
                                 'layerObj.color = acRed

plineObj.Layer = "K-grav"
plineObj.Update
  offsetObj = plineObj.Offset(47)
plineObj.Layer = "K-Mill"
plineObj.Update
  offsetObj = plineObj.Offset(56)
  offsetObj = plineObj.Offset(57)
plineObj.Layer = "AreaClear"
plineObj.Update
  offsetObj = plineObj.Offset(60)
plineObj.Layer = "EndMill-12"
plineObj.Update
  offsetObj = plineObj.Offset(62)
plineObj.Layer = "0"
plineObj.Update

I = 100
  
 Dim plineObj2 As AcadLWPolyline
  Dim points2(0 To 7) As Double
  
  points2(0) = points(0):    points2(1) = a + I:
  points2(2) = points(0):    points2(3) = (a * 2) + I:
  points2(4) = points(6):    points2(5) = (a * 2) + I:
  points2(6) = points(6):    points2(7) = a + I:
  
 For d = 2 To Cells(c, 4)
If d > 1 Then

Set plineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(points2)
  plineObj2.Closed = True
   
  points2(1) = points2(3) + I:
  points2(3) = points2(3) + a + I:
  points2(5) = points2(3):
  points2(7) = points2(1):
  
 
  
  ' Offset the polyline

plineObj2.Layer = "K-grav"
plineObj2.Update
  offsetObj = plineObj2.Offset(47)
plineObj2.Layer = "K-Mill"
plineObj2.Update
  offsetObj = plineObj2.Offset(56)
  offsetObj = plineObj2.Offset(57)
plineObj2.Layer = "AreaClear"
plineObj2.Update
  offsetObj = plineObj2.Offset(60)
plineObj2.Layer = "EndMill-12"
plineObj2.Update
  offsetObj = plineObj2.Offset(62)
plineObj2.Layer = "0"
plineObj2.Update

Else: d = 1
d = d + 1

d = Cells(c, 4)
End If

Next d
c = c + 1

a = Cells(c, 2)
b = Cells(c, 3)

I = 100

Loop
              
ZoomAll
AP.Quit
End Sub

Sub MF185()
' Creating search objects
Dim AP As Excel.Application
Dim WB As Excel.Workbook
Dim WS As Excel.Worksheet

Set AP = Excel.Application
Set WB = AP.Workbooks.Open("C:\Users\������\Desktop\������\2.xlsx")
Set WS = WB.Worksheets("������ �˨������ ������")
c = 19
a = Cells(c, 2)
b = Cells(c, 3)
d = Cells(c, 4)
E = 1
I = 0
On Error Resume Next

' Description of all the points that will make up rectangles
 Dim plineObj As AcadLWPolyline
  Dim points(0 To 7) As Double
points(6) = 0

Do While a > 0

' Calculate the coordinates of each point
  points(0) = points(6) + I:    points(1) = 0
  points(2) = points(6) + I:    points(3) = a
  points(4) = points(0) + b:    points(5) = a
  points(6) = points(0) + b:    points(7) = 0

' Constructing four segments and connecting the corresponding points

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
  plineObj.Closed = True
  
  ' Offset the polyline
Dim offsetObj As Variant
                                 'plineObj.color = acByLayer
                                 'Dim layerObj As AcadLayer
                                 'Set layerObj = ThisDrawing.Layers.Add("Ball-6.1")
                                 'layerObj.color = acRed

plineObj.Layer = "K-grav"
plineObj.Update
  offsetObj = plineObj.Offset(50)
plineObj.Layer = "AreaClear"
plineObj.Update
  offsetObj = plineObj.Offset(57)
plineObj.Layer = "K-Mill"
plineObj.Update
  offsetObj = plineObj.Offset(58)
plineObj.Layer = "0"
plineObj.Update

I = 100
  
 Dim plineObj2 As AcadLWPolyline
  Dim points2(0 To 7) As Double
  
  points2(0) = points(0):    points2(1) = a + I:
  points2(2) = points(0):    points2(3) = (a * 2) + I:
  points2(4) = points(6):    points2(5) = (a * 2) + I:
  points2(6) = points(6):    points2(7) = a + I:
  
 For d = 2 To Cells(c, 4)
If d > 1 Then

Set plineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(points2)
  plineObj2.Closed = True
   
  points2(1) = points2(3) + I:
  points2(3) = points2(3) + a + I:
  points2(5) = points2(3):
  points2(7) = points2(1):
  
 
  
  ' Offset the polyline

plineObj2.Layer = "K-grav"
plineObj2.Update
  offsetObj = plineObj2.Offset(50)
plineObj2.Layer = "AreaClear"
plineObj2.Update
  offsetObj = plineObj2.Offset(57)
plineObj2.Layer = "K-Mill"
plineObj2.Update
  offsetObj = plineObj2.Offset(58)
plineObj2.Layer = "0"
plineObj2.Update

Else: d = 1
d = d + 1

d = Cells(c, 4)
End If

Next d
c = c + 1

a = Cells(c, 2)
b = Cells(c, 3)

I = 100

Loop
              
ZoomAll
AP.Quit
End Sub

Sub MF187()
' Creating search objects
Dim AP As Excel.Application
Dim WB As Excel.Workbook
Dim WS As Excel.Worksheet

Set AP = Excel.Application
Set WB = AP.Workbooks.Open("C:\Users\������\Desktop\������\2.xlsx")
Set WS = WB.Worksheets("������ �˨������ ������")
c = 19
a = Cells(c, 2)
b = Cells(c, 3)
d = Cells(c, 4)
E = 1
I = 0
On Error Resume Next

' Description of all the points that will make up rectangles
   Dim plineObj As AcadLWPolyline
   Dim points(0 To 7) As Double
   Dim pointswindow(0 To 7) As Double
   Dim pointmove1(0 To 2) As Double
   Dim pointmove2(0 To 2) As Double
   Dim offsetObj As Variant
points(6) = 0

Do While a > 0

' Calculate the coordinates of each point
  points(0) = points(6) + I:    points(1) = 0
  points(2) = points(6) + I:    points(3) = a
  points(4) = points(0) + b:    points(5) = a
  points(6) = points(0) + b:    points(7) = 0

' Constructing four segments and connecting the corresponding points

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
  plineObj.Closed = True
  
  If a >= 160 Then
  If b >= 140 Then
  
  
  a1 = (a - 108) / 3
  b1 = (b - 104) / 2
  
  pointswindow(0) = points(0) + 50:       pointswindow(1) = points(1) + 50
  pointswindow(2) = points(0) + 50:       pointswindow(3) = points(1) + 50 + a1
  pointswindow(4) = points(0) + 50 + b1:  pointswindow(5) = points(1) + 50 + a1
  pointswindow(6) = points(0) + 50 + b1:  pointswindow(7) = points(1) + 50

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointswindow)
  plineObj.Closed = True
  
plineObj.Layer = "K-grav"
plineObj.Update

plineObj.Layer = "EndMill-4"
plineObj.Update
  offsetObj = plineObj.Offset(7)
plineObj.Layer = "K-grav"
plineObj.Update
  
    RetVal = plineObj.Copy
    pointmove1(0) = 0: pointmove1(1) = 0:
    pointmove2(0) = 0: pointmove2(1) = a1 + 4:
    plineObj.Move pointmove1, pointmove2

plineObj.Layer = "EndMill-4"
plineObj.Update
  offsetObj = plineObj.Offset(7)
plineObj.Layer = "K-grav"
plineObj.Update
    
    RetVal = plineObj.Copy
    pointmove1(0) = 0: pointmove1(1) = 0:
    pointmove2(0) = 0: pointmove2(1) = a1 + 4:
    plineObj.Move pointmove1, pointmove2
    
plineObj.Layer = "EndMill-4"
plineObj.Update
  offsetObj = plineObj.Offset(7)
plineObj.Layer = "K-grav"
plineObj.Update
    
    RetVal = plineObj.Copy
    pointmove1(0) = 0: pointmove1(1) = 0:
    pointmove2(0) = b1 + 4: pointmove2(1) = 0:
    plineObj.Move pointmove1, pointmove2
    
plineObj.Layer = "EndMill-4"
plineObj.Update
  offsetObj = plineObj.Offset(7)
plineObj.Layer = "K-grav"
plineObj.Update

    RetVal = plineObj.Copy
    pointmove1(0) = 0: pointmove1(1) = 0:
    pointmove2(0) = 0: pointmove2(1) = -(a1 + 4):
    plineObj.Move pointmove1, pointmove2
    
plineObj.Layer = "EndMill-4"
plineObj.Update
  offsetObj = plineObj.Offset(7)
plineObj.Layer = "K-grav"
plineObj.Update
    
    RetVal = plineObj.Copy
    pointmove1(0) = 0: pointmove1(1) = 0:
    pointmove2(0) = 0: pointmove2(1) = -(a1 + 4):
    plineObj.Move pointmove1, pointmove2
    
plineObj.Layer = "EndMill-4"
plineObj.Update
  offsetObj = plineObj.Offset(7)
plineObj.Layer = "K-grav"
plineObj.Update
  
End If
End If

I = 100
  
 Dim plineObj2 As AcadLWPolyline
  Dim points2(0 To 7) As Double
  
  points2(0) = points(0):    points2(1) = a + I:
  points2(2) = points(0):    points2(3) = (a * 2) + I:
  points2(4) = points(6):    points2(5) = (a * 2) + I:
  points2(6) = points(6):    points2(7) = a + I:
  
 For d = 2 To Cells(c, 4)
If d > 1 Then

Set plineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(points2)
  plineObj2.Closed = True
  
If a >= 160 Then
If b >= 140 Then
   
  pointswindow(0) = points2(0) + 50:       pointswindow(1) = points2(1) + 50
  pointswindow(2) = points2(0) + 50:       pointswindow(3) = points2(1) + 50 + a1
  pointswindow(4) = points2(0) + 50 + b1:  pointswindow(5) = points2(1) + 50 + a1
  pointswindow(6) = points2(0) + 50 + b1:  pointswindow(7) = points2(1) + 50

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointswindow)
  plineObj.Closed = True
  
plineObj.Layer = "K-grav"
plineObj.Update

plineObj.Layer = "EndMill-4"
plineObj.Update
  offsetObj = plineObj.Offset(7)
plineObj.Layer = "K-grav"
plineObj.Update
  
    RetVal = plineObj.Copy
    pointmove1(0) = 0: pointmove1(1) = 0:
    pointmove2(0) = 0: pointmove2(1) = a1 + 4:
    plineObj.Move pointmove1, pointmove2

plineObj.Layer = "EndMill-4"
plineObj.Update
  offsetObj = plineObj.Offset(7)
plineObj.Layer = "K-grav"
plineObj.Update
    
    RetVal = plineObj.Copy
    pointmove1(0) = 0: pointmove1(1) = 0:
    pointmove2(0) = 0: pointmove2(1) = a1 + 4:
    plineObj.Move pointmove1, pointmove2
    
plineObj.Layer = "EndMill-4"
plineObj.Update
  offsetObj = plineObj.Offset(7)
plineObj.Layer = "K-grav"
plineObj.Update
    
    RetVal = plineObj.Copy
    pointmove1(0) = 0: pointmove1(1) = 0:
    pointmove2(0) = b1 + 4: pointmove2(1) = 0:
    plineObj.Move pointmove1, pointmove2
    
plineObj.Layer = "EndMill-4"
plineObj.Update
  offsetObj = plineObj.Offset(7)
plineObj.Layer = "K-grav"
plineObj.Update

    RetVal = plineObj.Copy
    pointmove1(0) = 0: pointmove1(1) = 0:
    pointmove2(0) = 0: pointmove2(1) = -(a1 + 4):
    plineObj.Move pointmove1, pointmove2
    
plineObj.Layer = "EndMill-4"
plineObj.Update
  offsetObj = plineObj.Offset(7)
plineObj.Layer = "K-grav"
plineObj.Update
    
    RetVal = plineObj.Copy
    pointmove1(0) = 0: pointmove1(1) = 0:
    pointmove2(0) = 0: pointmove2(1) = -(a1 + 4):
    plineObj.Move pointmove1, pointmove2
    
plineObj.Layer = "EndMill-4"
plineObj.Update
  offsetObj = plineObj.Offset(7)
plineObj.Layer = "K-grav"
plineObj.Update

End If
End If
  
  points2(1) = points2(3) + I:
  points2(3) = points2(3) + a + I:
  points2(5) = points2(3):
  points2(7) = points2(1):

Else: d = 1
d = d + 1

d = Cells(c, 4)
End If

Next d
c = c + 1

a = Cells(c, 2)
b = Cells(c, 3)

I = 100

Loop
              
ZoomAll
AP.Quit
End Sub

Sub MF191()
' Creating search objects
Dim AP As Excel.Application
Dim WB As Excel.Workbook
Dim WS As Excel.Worksheet

Set AP = Excel.Application
Set WB = AP.Workbooks.Open("C:\Users\������\Desktop\������\2.xlsx")
Set WS = WB.Worksheets("������ �˨������ ������")
c = 19
a = Cells(c, 2)
b = Cells(c, 3)
d = Cells(c, 4)
E = 1
I = 0
On Error Resume Next

' Description of all the points that will make up rectangles
 Dim plineObj As AcadLWPolyline
  Dim points(0 To 7) As Double
points(6) = 0

Do While a > 0

' Calculate the coordinates of each point
  points(0) = points(6) + I:    points(1) = 0
  points(2) = points(6) + I:    points(3) = a
  points(4) = points(0) + b:    points(5) = a
  points(6) = points(0) + b:    points(7) = 0

' Constructing four segments and connecting the corresponding points

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
  plineObj.Closed = True
  
  ' Offset the polyline
Dim offsetObj As Variant
                                 'plineObj.color = acByLayer
                                 'Dim layerObj As AcadLayer
                                 'Set layerObj = ThisDrawing.Layers.Add("Ball-6.1")
                                 'layerObj.color = acRed


plineObj.Layer = "EndMill-4_1mm"
plineObj.Update
  offsetObj = plineObj.Offset(49)
plineObj.Layer = "K-grav_8mm"
plineObj.Update
  offsetObj = plineObj.Offset(50)
plineObj.Layer = "EndMill-4_1mm"
plineObj.Update
  offsetObj = plineObj.Offset(51)
plineObj.Layer = "K-Mill"
plineObj.Update
  offsetObj = plineObj.Offset(64)
  offsetObj = plineObj.Offset(65)
plineObj.Layer = "EndMill-4_8mm"
plineObj.Update
  offsetObj = plineObj.Offset(66)
  offsetObj = plineObj.Offset(68)
plineObj.Layer = "AreaClear"
plineObj.Update
  offsetObj = plineObj.Offset(69)
plineObj.Layer = "EndMill-4_10mm"
plineObj.Update
  offsetObj = plineObj.Offset(71)
plineObj.Layer = "0"
plineObj.Update

I = 100
  
 Dim plineObj2 As AcadLWPolyline
  Dim points2(0 To 7) As Double
  
  points2(0) = points(0):    points2(1) = a + I:
  points2(2) = points(0):    points2(3) = (a * 2) + I:
  points2(4) = points(6):    points2(5) = (a * 2) + I:
  points2(6) = points(6):    points2(7) = a + I:
  
 For d = 2 To Cells(c, 4)
If d > 1 Then

Set plineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(points2)
  plineObj2.Closed = True
   
  points2(1) = points2(3) + I:
  points2(3) = points2(3) + a + I:
  points2(5) = points2(3):
  points2(7) = points2(1):
  
 
  
  ' Offset the polyline

plineObj2.Layer = "EndMill-4_1mm"
plineObj2.Update
  offsetObj = plineObj2.Offset(49)
plineObj2.Layer = "K-grav_8mm"
plineObj2.Update
  offsetObj = plineObj2.Offset(50)
plineObj2.Layer = "EndMill-4_1mm"
plineObj2.Update
  offsetObj = plineObj2.Offset(51)
plineObj2.Layer = "K-Mill"
plineObj2.Update
  offsetObj = plineObj2.Offset(64)
  offsetObj = plineObj2.Offset(65)
plineObj2.Layer = "EndMill-4_8mm"
plineObj2.Update
  offsetObj = plineObj2.Offset(66)
  offsetObj = plineObj2.Offset(68)
plineObj2.Layer = "AreaClear"
plineObj2.Update
  offsetObj = plineObj2.Offset(69)
plineObj2.Layer = "EndMill-4_10mm"
plineObj2.Update
  offsetObj = plineObj2.Offset(71)
plineObj2.Layer = "0"
plineObj2.Update

Else: d = 1
d = d + 1

d = Cells(c, 4)
End If

Next d
c = c + 1

a = Cells(c, 2)
b = Cells(c, 3)

I = 100

Loop
              
ZoomAll
AP.Quit
End Sub

Sub MF192()
' Creating search objects
Dim AP As Excel.Application
Dim WB As Excel.Workbook
Dim WS As Excel.Worksheet

Set AP = Excel.Application
Set WB = AP.Workbooks.Open("C:\Users\������\Desktop\������\2.xlsx")
Set WS = WB.Worksheets("������ �˨������ ������")
c = 19
a = Cells(c, 2)
b = Cells(c, 3)
d = Cells(c, 4)
E = 1
I = 0
On Error Resume Next

' Description of all the points that will make up rectangles
 Dim plineObj As AcadLWPolyline
  Dim points(0 To 7) As Double
  Dim pointsac(0 To 33) As Double
  Dim pointsac2(0 To 15) As Double
points(6) = 0

Do While a > 0

' Calculate the coordinates of each point
  points(0) = points(6) + I:    points(1) = 0
  points(2) = points(6) + I:    points(3) = a
  points(4) = points(0) + b:    points(5) = a
  points(6) = points(0) + b:    points(7) = 0

' Constructing four segments and connecting the corresponding points

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
  plineObj.Closed = True
  
  ' Offset the polyline
Dim offsetObj As Variant
                                 'plineObj.color = acByLayer
                                 'Dim layerObj As AcadLayer
                                 'Set layerObj = ThisDrawing.Layers.Add("Ball-6.1")
                                 'layerObj.color = acRed


plineObj.Layer = "K_grav_2mm"
plineObj.Update
  offsetObj = plineObj.Offset(50)
plineObj.Layer = "K_2mm"
plineObj.Update
  offsetObj = plineObj.Offset(54)
  offsetObj = plineObj.Offset(55)
  offsetObj = plineObj.Offset(56)
  offsetObj = plineObj.Offset(76)
plineObj.Layer = "0"
plineObj.Update

pointsac(0) = points(0) + 60:       pointsac(1) = points(1) + 60:
pointsac(2) = points(0) + 60:       pointsac(3) = points(3) - 60:
pointsac(4) = points(0) + 69.37:    pointsac(5) = points(3) - 66:
pointsac(6) = points(0) + 66:       pointsac(7) = points(3) - 69.37:

pointsac(8) = points(0) + 60:       pointsac(9) = points(3) - 60:
pointsac(10) = points(4) - 60:      pointsac(11) = points(3) - 60:
pointsac(12) = points(4) - 66:      pointsac(13) = points(3) - 69.37:
pointsac(14) = points(4) - 69.37:   pointsac(15) = points(3) - 66:

pointsac(16) = points(4) - 60:      pointsac(17) = points(3) - 60:
pointsac(18) = points(4) - 60:      pointsac(19) = points(1) + 60:
pointsac(20) = points(4) - 69.37:   pointsac(21) = points(1) + 66:
pointsac(22) = points(4) - 66:      pointsac(23) = points(1) + 69.37:
  
pointsac(24) = points(4) - 60:      pointsac(25) = points(1) + 60:
pointsac(26) = points(0) + 60:      pointsac(27) = points(1) + 60:
pointsac(28) = points(0) + 66:      pointsac(29) = points(1) + 69.37:
pointsac(30) = points(0) + 69.37:   pointsac(31) = points(1) + 66:
pointsac(32) = points(0) + 60:      pointsac(33) = points(1) + 60:

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsac)
  plineObj.Closed = True
plineObj.Layer = "AreaClear"
plineObj.Update


pointsac2(0) = points(0) + 70:       pointsac2(1) = points(1) + 76:
pointsac2(2) = points(0) + 70:       pointsac2(3) = points(3) - 76:
pointsac2(4) = points(0) + 76:       pointsac2(5) = points(3) - 70:
pointsac2(6) = points(4) - 76:       pointsac2(7) = points(3) - 70:

pointsac2(8) = points(4) - 70:       pointsac2(9) = points(3) - 76:
pointsac2(10) = points(4) - 70:      pointsac2(11) = points(1) + 76:
pointsac2(12) = points(4) - 76:      pointsac2(13) = points(1) + 70:
pointsac2(14) = points(0) + 76:      pointsac2(15) = points(1) + 70:

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsac2)
  plineObj.Closed = True
plineObj.Layer = "AreaClear"
plineObj.Update


Dim currentBulge As Double
    currentBulge = plineObj.GetBulge(2)
    ' Set the convexity of the 1st segment
    plineObj.SetBulge 1, -0.41421
    plineObj.Update
    plineObj.SetBulge 3, -0.41421
    plineObj.Update
    plineObj.SetBulge 5, -0.41421
    plineObj.Update
    plineObj.SetBulge 7, -0.41421
    plineObj.Update
    plineObj.Layer = "AreaClear"
    plineObj.Update
    plineObj.Closed = True

I = 100
  
 Dim plineObj2 As AcadLWPolyline
  Dim points2(0 To 7) As Double
  
  points2(0) = points(0):    points2(1) = a + I:
  points2(2) = points(0):    points2(3) = (a * 2) + I:
  points2(4) = points(6):    points2(5) = (a * 2) + I:
  points2(6) = points(6):    points2(7) = a + I:
  
 For d = 2 To Cells(c, 4)
If d > 1 Then

Set plineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(points2)
  plineObj2.Closed = True

  ' Offset the polyline

plineObj2.Layer = "K_grav_2mm"
plineObj2.Update
  offsetObj = plineObj2.Offset(50)
plineObj2.Layer = "K_2mm"
plineObj2.Update
  offsetObj = plineObj2.Offset(54)
  offsetObj = plineObj2.Offset(55)
  offsetObj = plineObj2.Offset(56)
  offsetObj = plineObj2.Offset(76)
plineObj2.Layer = "0"
plineObj2.Update

pointsac(0) = points2(0) + 60:       pointsac(1) = points2(1) + 60:
pointsac(2) = points2(0) + 60:       pointsac(3) = points2(3) - 60:
pointsac(4) = points2(0) + 69.37:    pointsac(5) = points2(3) - 66:
pointsac(6) = points2(0) + 66:       pointsac(7) = points2(3) - 69.37:

pointsac(8) = points2(0) + 60:       pointsac(9) = points2(3) - 60:
pointsac(10) = points2(4) - 60:      pointsac(11) = points2(3) - 60:
pointsac(12) = points2(4) - 66:      pointsac(13) = points2(3) - 69.37:
pointsac(14) = points2(4) - 69.37:   pointsac(15) = points2(3) - 66:

pointsac(16) = points2(4) - 60:      pointsac(17) = points2(3) - 60:
pointsac(18) = points2(4) - 60:      pointsac(19) = points2(1) + 60:
pointsac(20) = points2(4) - 69.37:   pointsac(21) = points2(1) + 66:
pointsac(22) = points2(4) - 66:      pointsac(23) = points2(1) + 69.37:
  
pointsac(24) = points2(4) - 60:      pointsac(25) = points2(1) + 60:
pointsac(26) = points2(0) + 60:      pointsac(27) = points2(1) + 60:
pointsac(28) = points2(0) + 66:      pointsac(29) = points2(1) + 69.37:
pointsac(30) = points2(0) + 69.37:   pointsac(31) = points2(1) + 66:
pointsac(32) = points2(0) + 60:      pointsac(33) = points2(1) + 60:

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsac)
  plineObj.Closed = True
plineObj.Layer = "AreaClear"
plineObj.Update

pointsac2(0) = points2(0) + 70:       pointsac2(1) = points2(1) + 76:
pointsac2(2) = points2(0) + 70:       pointsac2(3) = points2(3) - 76:
pointsac2(4) = points2(0) + 76:       pointsac2(5) = points2(3) - 70:
pointsac2(6) = points2(4) - 76:       pointsac2(7) = points2(3) - 70:

pointsac2(8) = points2(4) - 70:       pointsac2(9) = points2(3) - 76:
pointsac2(10) = points2(4) - 70:      pointsac2(11) = points2(1) + 76:
pointsac2(12) = points2(4) - 76:      pointsac2(13) = points2(1) + 70:
pointsac2(14) = points2(0) + 76:      pointsac2(15) = points2(1) + 70:

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pointsac2)
  plineObj.Closed = True
plineObj.Layer = "AreaClear"
plineObj.Update

    plineObj.SetBulge 1, -0.41421
    plineObj.Update
    plineObj.SetBulge 3, -0.41421
    plineObj.Update
    plineObj.SetBulge 5, -0.41421
    plineObj.Update
    plineObj.SetBulge 7, -0.41421
    plineObj.Update
    plineObj.Layer = "AreaClear"
    plineObj.Update
    plineObj.Closed = True
   
  points2(1) = points2(3) + I:
  points2(3) = points2(3) + a + I:
  points2(5) = points2(3):
  points2(7) = points2(1):

Else: d = 1
d = d + 1

d = Cells(c, 4)
End If

Next d
c = c + 1

a = Cells(c, 2)
b = Cells(c, 3)

I = 100

Loop
              
ZoomAll
AP.Quit
End Sub

Sub MF194()
' Creating search objects
Dim AP As Excel.Application
Dim WB As Excel.Workbook
Dim WS As Excel.Worksheet

Set AP = Excel.Application
Set WB = AP.Workbooks.Open("C:\Users\������\Desktop\������\2.xlsx")
Set WS = WB.Worksheets("������ �˨������ ������")
c = 19
a = Cells(c, 2)
b = Cells(c, 3)
d = Cells(c, 4)
E = 1
I = 0
On Error Resume Next

' Description of all the points that will make up rectangles
 Dim plineObj As AcadLWPolyline
  Dim points(0 To 7) As Double
points(6) = 0

Do While a > 0

' Calculate the coordinates of each point
  points(0) = points(6) + I:    points(1) = 0
  points(2) = points(6) + I:    points(3) = a
  points(4) = points(0) + b:    points(5) = a
  points(6) = points(0) + b:    points(7) = 0

' Constructing four segments and connecting the corresponding points

Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
  plineObj.Closed = True
  
  ' Offset the polyline
Dim offsetObj As Variant
                                 'plineObj.color = acByLayer
                                 'Dim layerObj As AcadLayer
                                 'Set layerObj = ThisDrawing.Layers.Add("Ball-6.1")
                                 'layerObj.color = acRed

plineObj.Layer = "K-grav"
plineObj.Update
  offsetObj = plineObj.Offset(60)
plineObj.Layer = "EndMill-4_7mm"
plineObj.Update
  offsetObj = plineObj.Offset(77.86)
plineObj.Layer = "K-Mill"
plineObj.Update
  offsetObj = plineObj.Offset(81.36)
plineObj.Layer = "Ball-35_7mm"
plineObj.Update
  offsetObj = plineObj.Offset(83)
plineObj.Layer = "0"
plineObj.Update

I = 100
  
 Dim plineObj2 As AcadLWPolyline
  Dim points2(0 To 7) As Double
  
  points2(0) = points(0):    points2(1) = a + I:
  points2(2) = points(0):    points2(3) = (a * 2) + I:
  points2(4) = points(6):    points2(5) = (a * 2) + I:
  points2(6) = points(6):    points2(7) = a + I:
  
 For d = 2 To Cells(c, 4)
If d > 1 Then

Set plineObj2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(points2)
  plineObj2.Closed = True
   
  points2(1) = points2(3) + I:
  points2(3) = points2(3) + a + I:
  points2(5) = points2(3):
  points2(7) = points2(1):
  
 
  
  ' Offset the polyline

plineObj2.Layer = "K-grav"
plineObj2.Update
  offsetObj = plineObj2.Offset(60)
plineObj2.Layer = "EndMill-4_7mm"
plineObj2.Update
  offsetObj = plineObj2.Offset(77.86)
plineObj2.Layer = "K-Mill"
plineObj2.Update
  offsetObj = plineObj2.Offset(81.36)
plineObj2.Layer = "Ball-35_7mm"
plineObj2.Update
  offsetObj = plineObj2.Offset(83)
plineObj2.Layer = "0"
plineObj2.Update

Else: d = 1
d = d + 1

d = Cells(c, 4)
End If

Next d
c = c + 1

a = Cells(c, 2)
b = Cells(c, 3)

I = 100

Loop
              
ZoomAll
AP.Quit
End Sub

