Module General
    Public Function MeasureAString(ByRef Width0Height1Both2 As Byte, ByRef str As String, ByRef tFont As Font) As Object
        Using gx As Graphics = Graphics.FromImage(New Bitmap(1, 1))
            Dim SFormat As New System.Drawing.StringFormat
            Dim rect As New System.Drawing.RectangleF(0, 0, 6000, 6000)
            Dim range() As CharacterRange = New CharacterRange() {New CharacterRange(0, str.Length)}
            SFormat.SetMeasurableCharacterRanges(range)
            Dim regions() As Region = gx.MeasureCharacterRanges(str, tFont, rect, SFormat)
            rect = regions(0).GetBounds(gx)

            If Width0Height1Both2 = 0 Then Return rect.Right + 1 'gx.MeasureString(str, Font, 50000000, lolz).Width
            If Width0Height1Both2 = 1 Then Return rect.Bottom + 1 'gx.MeasureString(str, Font, 50000000, lolz).Height
            If Width0Height1Both2 = 2 Then Return New SizeF(rect.Right + 1, rect.Bottom + 1) 'gx.MeasureString(str, Font, 50000000, lolz)
        End Using
        Return -1
    End Function

    Public Sub FillArea(ByRef bmp As Bitmap, ByRef x As Integer, ByRef y As Integer, ByRef newColor As Color, ByRef oldColor As Color)
        Dim lolz As Color = bmp.GetPixel(x - 1, y)
        If x - 1 > 0 AndAlso bmp.GetPixel(x - 1, y) = oldColor Then
            bmp.SetPixel(x - 1, y, newColor)
            FillArea(bmp, x - 1, y, newColor, oldColor)
        End If
        If x + 1 < bmp.Width AndAlso bmp.GetPixel(x + 1, y) = oldColor Then
            bmp.SetPixel(x + 1, y, newColor)
            FillArea(bmp, x + 1, y, newColor, oldColor)
        End If

        If y - 1 > 0 AndAlso bmp.GetPixel(x, y - 1) = oldColor Then
            bmp.SetPixel(x, y - 1, newColor)
            FillArea(bmp, x, y - 1, newColor, oldColor)
        End If
        If y + 1 < bmp.Height AndAlso bmp.GetPixel(x, y + 1) = oldColor Then
            bmp.SetPixel(x, y + 1, newColor)
            FillArea(bmp, x, y + 1, newColor, oldColor)
        End If
    End Sub
End Module

#Region "Data Types"

Public Structure NumAndBool
    Public index As Integer, bool As Boolean
    Public Sub New(ByRef tIndex As Integer, ByVal tBoolean As Boolean)
        index = tIndex
        bool = tBoolean
    End Sub
End Structure

Public Class TwoStrings
    Public String1 As String
    Public String2 As String
    Public Sub New(ByRef tString1 As String, ByRef tString2 As String)
        String1 = tString1
        String2 = tString2
    End Sub
End Class

Public Class StringAndObject
    Public String1 As String
    Public Object1 As Object
    Public Sub New(ByRef tString1 As String, ByVal tObj As Object)
        String1 = tString1
        Object1 = tObj
    End Sub
End Class

Public Class CSize
    Public width As Short
    Public height As Short
    Public autosize As Boolean

    Public Sub New(ByRef sWidth As Short, ByRef sHeight As Short)
        width = sWidth
        height = sHeight
    End Sub

    Public Sub New(ByRef sAutosize As Boolean)
        autosize = True
    End Sub
End Class

Public Class CLocation
    Public left As Short
    Public top As Short
    Public leftCenter As Short 'Use to center the object at a specific X coord
    Public topCenter As Short

    Public Sub New(ByRef tLeft As Short, ByRef tTop As Short, Optional ByRef lCenter As Short = -1, Optional ByRef tCenter As Short = -1)
        left = tLeft
        top = tTop
        leftCenter = lCenter
        topCenter = tCenter
    End Sub
End Class

Public Class COptions
    Public PropertyList As List(Of StringAndObject)
    Public Sub New(ByRef tProperty As String, ByRef tValue As Object)
        PropertyList = New List(Of StringAndObject)
        PropertyList.Add(New StringAndObject(tProperty, tValue))
    End Sub

    Public Sub New(ByRef tPropertyList As List(Of StringAndObject))
        PropertyList = tPropertyList
    End Sub
End Class

#End Region
