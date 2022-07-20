Private Function ReturnPointEvent(ByRef baseLine As MapObjects2.Line, _
                                  ByVal dist As Double) _
                                  As MapObjects2.Point

Dim lnSelectedPart As New MapObjects2.Line
Dim ptsPart As New MapObjects2.Points
Dim eventPoint As MapObjects2.Point
Dim startVertex As Long
Dim lineLength, runLength, segLength, fragLength As Double
Dim i As Long

For i = 0 To baseLine.Parts(0).Count - 1
  ptsPart.Add baseLine.Parts(0)(i)
Next
lnSelectedPart.Parts.Add ptsPart
lineLength = lnSelectedPart.Length

If dist < 0 Or dist > lineLength Then
  MsgBox "Invalid point event distance."
  Exit Function
End If
                                  
startVertex = 0
'Find start vertex
runLength = 0
For i = 0 To ptsPart.Count - 2
  segLength = ptsPart(i).DistanceTo(ptsPart(i + 1))
  runLength = runLength + segLength
  If runLength > dist Then
    startVertex = i
    fragLength = dist - (runLength - segLength)
    Set eventPoint = FindPtAlongSeg(ptsPart(i), ptsPart(i + 1), fragLength)
    Exit For
  End If
Next

Set ReturnPointEvent = eventPoint

End Function

Private Function ReturnLineEvent(ByRef baseLine As MapObjects2.Line, _
                                 ByVal startDist As Double, _
                                 ByVal stopDist As Double) _
                                 As MapObjects2.Line
                                 
Dim ptsNew As New MapObjects2.Points
Dim ptsPart As New MapObjects2.Points
Dim lnNew As New MapObjects2.Line
Dim lnSelectedPart As New MapObjects2.Line
Dim startPoint As MapObjects2.Point
Dim stopPoint As MapObjects2.Point
Dim startVertex, stopVertex As Long
Dim lineLength, runLength, segLength, fragLength As Double
Dim i, j As Long

For i = 0 To baseLine.Parts(0).Count - 1
  ptsPart.Add baseLine.Parts(0)(i)
Next
lnSelectedPart.Parts.Add ptsPart
lineLength = lnSelectedPart.Length

If startDist < 0 Then
  MsgBox "Invalid start distance."
  Exit Function
End If

If stopDist > lineLength Then
  MsgBox "Invalid stop distance."
  Exit Function
End If

startVertex = 0
stopVertex = 0
'Find start vertex
runLength = 0
For i = 0 To ptsPart.Count - 2
  segLength = ptsPart(i).DistanceTo(ptsPart(i + 1))
  runLength = runLength + segLength
  If runLength > startDist Then
    startVertex = i
    j = i
    fragLength = startDist - (runLength - segLength)
    Set startPoint = FindPtAlongSeg(ptsPart(i), ptsPart(i + 1), fragLength)
    Exit For
  End If
Next
'Find stop vertex
runLength = runLength - segLength
fragLength = 0
For i = j To ptsPart.Count - 2
  runLength = runLength + ptsPart(i).DistanceTo(ptsPart(i + 1))
  If runLength > stopDist Then
    stopVertex = i
    fragLength = stopDist - (runLength - segLength)
    Set stopPoint = FindPtAlongSeg(ptsPart(i), ptsPart(i + 1), fragLength)
    Exit For
  End If
Next
'Build line from startVertex to stopVertex and all
'vertices in between
ptsNew.Add startPoint
For i = (startVertex + 1) To stopVertex
  ptsNew.Add ptsPart(i)
Next
ptsNew.Add stopPoint
lnNew.Parts.Add ptsNew

Set ReturnLineEvent = lnNew
      
End Function

Private Function FindPtAlongSeg(ByRef ptVertex1 As MapObjects2.Point, _
                                ByRef ptVertex2 As MapObjects2.Point, _
                                ByVal dist As Double) _
                                As MapObjects2.Point
                                
Dim ptAlong As New MapObjects2.Point
Dim x1, y1, x2, y2 As Double
Dim xLeg, yLeg, hyp, ratio As Double
Dim xNew, yNew As Double

x1 = ptVertex1.x
y1 = ptVertex1.y
x2 = ptVertex2.x
y2 = ptVertex2.y
xLeg = Abs(x1 - x2)
yLeg = Abs(y1 - y2)
hyp = Sqr((xLeg ^ 2) + (yLeg ^ 2))
ratio = dist / hyp
Select Case True
  Case x1 < x2: xNew = x1 + (ratio * xLeg)
  Case x1 > x2: xNew = x1 - (ratio * xLeg)
  Case Else: xNew = x1
End Select
Select Case True
  Case y1 < y2: yNew = y1 + (ratio * yLeg)
  Case y1 > y2: yNew = y1 - (ratio * yLeg)
  Case Else: yNew = y1
End Select
ptAlong.x = xNew
ptAlong.y = yNew
Set FindPtAlongSeg = ptAlong

End Function
