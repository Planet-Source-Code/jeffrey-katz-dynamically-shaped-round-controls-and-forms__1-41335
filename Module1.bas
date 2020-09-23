Attribute VB_Name = "Module1"
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Const RGN_AND = 1
Public Const RGN_COPY = 5
Public Const RGN_DIFF = 4
Public Const RGN_OR = 2
Public Const RGN_XOR = 3
Public cSize As Integer

Sub ShapeObject(TheObject As Object, Optional UseTwips As Boolean = 0)                                                                                                                                                                                                         'This code and all parts of it Copyright (C) 2001 Jeff Katz
    
' In essesnce what this subroutine does is it creates first two copies of the rectangular size
' of the object. It then cuts somewhat of a cross pattern + leaving the corners out, and then
' it cuts a circle into the regeon of each corner, creating the inverse of the final shape. In
' the last step, it uses RGN_DIFF to invert the cutout onto the copy, leaving us with the final
' shape. This function can be used on any control or form with a hWnd.

With TheObject
'Const cSize = 80

Select Case UseTwips
    
    Case 0 ' Dont use twips in our calculations. The form's scalemode should be set to three
    
    thematrix = CreateRectRgn(0, 0, .Width, .Height)    'The Whole Object
    notthematrix = CreateRectRgn(0, 0, .Width, .Height) 'The Whole Object

    a = CreateRectRgn(cSize / 2, 0, .Width - (cSize / 2), .Height) '[] the object
    b = CreateRectRgn(0, cSize / 2, .Width, .Height - (cSize / 2)) ' = the object

    c = CreateEllipticRgn(0, 0, cSize, cSize) 'upper left corner
    d = CreateEllipticRgn(0, .Height, cSize, .Height - cSize)
    e = CreateEllipticRgn(.Width, 0, .Width - cSize, cSize)
    f = CreateEllipticRgn(.Width, .Height, .Width - cSize, .Height - cSize)

    Case 1 ' Use twips, using scalex and scaley

    thematrix = CreateRectRgn(0, 0, .ScaleX(.Width, 1, 3), .ScaleY(.Height, 1, 3)) 'The Whole Object
    notthematrix = CreateRectRgn(0, 0, .ScaleX(.Width, 1, 3), .Height) 'The Whole Object

    a = CreateRectRgn(cSize / 2, 0, .ScaleX(.Width, 1, 3) - (cSize / 2), .ScaleY(.Height, 1, 3)) '[] the object
    b = CreateRectRgn(0, cSize / 2, .ScaleX(.Width, 1, 3), .ScaleY(.Height, 1, 3) - (cSize / 2)) ' = the object

    c = CreateEllipticRgn(0, 0, cSize, cSize) 'upper left corner
    d = CreateEllipticRgn(0, .ScaleY(.Height, 1, 3), cSize, .ScaleY(.Height, 1, 3) - cSize)
    e = CreateEllipticRgn(.ScaleX(.Width, 1, 3), 0, .ScaleX(.Width, 1, 3) - cSize, cSize)
    f = CreateEllipticRgn(.ScaleX(.Width, 1, 3), .ScaleY(.Height, 1, 3), .ScaleX(.Width, 1, 3) - cSize, .ScaleY(.Height, 1, 3) - cSize)

End Select

    g = CombineRgn(thematrix, thematrix, a, 4) 'cut out pieces
    g = CombineRgn(thematrix, thematrix, b, 4) 'cut out pieces
    g = CombineRgn(thematrix, thematrix, c, 4) 'cut out pieces
    g = CombineRgn(thematrix, thematrix, d, 4) 'cut out pieces
    g = CombineRgn(thematrix, thematrix, e, 4) 'cut out pieces
    g = CombineRgn(thematrix, thematrix, f, 4) 'cut out pieces
    g = CombineRgn(thematrix, notthematrix, thematrix, 4) 'invert

    m = SetWindowRgn(.hwnd, thematrix, True) ' set the final rgn to the object's hwnd.
    DeleteObject thematrix
    DeleteObject notthematrix
    DeleteObject a
    DeleteObject b
    DeleteObject c
    DeleteObject d
    DeleteObject e
    DeleteObject f
    DeleteObject g
    DeleteObject m
    
    ' if we dont delete the objects, we end up with large memory leaks
End With
End Sub

