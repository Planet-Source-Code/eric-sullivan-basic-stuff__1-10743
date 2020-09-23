Attribute VB_Name = "Module1"

Public Sub ConvertImages(ImageToShow As Image, ImageToHide As Image)
    ImageToShow.Visible = True
    ImageToHide.Visible = False
End Sub

Public Sub SetButtonValue(Line1val As Variant, line2val As Variant, line3val As Variant, line4val As Variant)
    Form1.Line1.Visible = Line1val
    Form1.Line2.Visible = line2val
    Form1.Line3.Visible = line3val
    Form1.Line4.Visible = line4val
End Sub

Public Sub SetButton(Colour1 As ColorConstants, Colour2 As ColorConstants)
    Form1.Line1.BorderColor = Colour1
    Form1.Line2.BorderColor = Colour1
    Form1.Line3.BorderColor = Colour2
    Form1.Line4.BorderColor = Colour2
End Sub
