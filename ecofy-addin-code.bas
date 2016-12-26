Attribute VB_Name = "Module1"
Sub Auto_Open()
    Dim oToolbar As CommandBar
    Dim oButton As CommandBarButton
    Dim MyToolbar As String

    ' Give the toolbar a name
    MyToolbar = "Ecofy"

    On Error Resume Next
    ' so that it doesn't stop on the next line if the toolbar's already there

    ' Create the toolbar; PowerPoint will error if it already exists
    Set oToolbar = CommandBars.Add(Name:=MyToolbar, _
        Position:=msoBarFloating, Temporary:=True)
    If Err.Number <> 0 Then
          ' The toolbar's already there, so we have nothing to do
          Exit Sub
    End If

    On Error GoTo ErrorHandler

    ' Now add a button to the new toolbar
    Set oButton = oToolbar.Controls.Add(Type:=msoControlButton)

    ' And set some of the button's properties

    With oButton

         .DescriptionText = "Optimize presentation for printing"
          'Tooltip text when mouse if placed over button

         .Caption = "Do Button1 Stuff"
         'Text if Text in Icon is chosen

         .OnAction = "Ecofy"
          'Runs the Sub Button1() code when clicked

         .Style = msoButtonIcon
          ' Button displays as icon, not text or both

         .FaceId = 52
          ' chooses icon #52 from the available Office icons

    End With

    ' Repeat the above for as many more buttons as you need to add
    ' Be sure to change the .OnAction property at least for each new button

    ' You can set the toolbar position and visibility here if you like
    ' By default, it'll be visible when created. Position will be ignored in PPT 2007 and later
    oToolbar.Top = 150
    oToolbar.Left = 150
    oToolbar.Visible = True

NormalExit:
    Exit Sub   ' so it doesn't go on to run the errorhandler code

ErrorHandler:
     'Just in case there is an error
     MsgBox Err.Number & vbCrLf & Err.Description
     Resume NormalExit:
End Sub

Sub Ecofy()
    ActivePresentation.SlideMaster.Background.Fill.ForeColor.RGB = RGB(255, 255, 255)

    For Each sld In ActivePresentation.Slides
    
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
