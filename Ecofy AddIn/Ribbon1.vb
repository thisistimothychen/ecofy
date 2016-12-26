Imports Microsoft.Office.Tools.Ribbon
Imports PowerPoint = Microsoft.Office.Interop.PowerPoint

Public Class Ribbon1
    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click

        Dim myDocument As PowerPoint.Slide
        Globals.ThisAddIn.Application.ActivePresentation.SlideMaster.Background.Fill.ForeColor.RGB = RGB(255, 255, 255)

        For Each sld In Globals.ThisAddIn.Application.ActivePresentation.Slides

            For Each sh In sld.Shapes
                With sh
                    On Error Resume Next

                    ' Shadows
                    .Shadow.Visible = False

                    ' Fill
                    If .Fill.ForeColor.RGB <> RGB(255, 255, 255) Then
                        .Fill.ForeColor.RGB = RGB(255, 255, 255)

                    End If

                    ' Line
                    If .Line.Visible Then
                        If .Line.ForeColor.RGB Then
                            .Line.ForeColor.RGB = RGB(255, 255, 255)
                        End If
                    End If

                    ' Text
                    If .HasTextFrame Then
                        If .TextFrame.HasText Then
                            For x = 1 To .TextFrame.TextRange.Runs.Count
                                .TextFrame.TextRange.Runs(x).Font.Color.RGB = RGB(0, 0, 0)
                            Next
                        End If
                    End If

                End With
            Next
        Next

    End Sub

End Class
