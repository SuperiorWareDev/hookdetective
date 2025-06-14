Public Class HighLightForm
    Private Sub HighLightForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Remove the border and title bar to make it just a blank slate
        Me.FormBorderStyle = FormBorderStyle.None

        ' Hide the form from the Windows taskbar so it's extra stealthy
        Me.ShowInTaskbar = False

        ' Make the form as small as possible
        Me.Size = New Size(1, 1)

        ' This is the magic part for transparency. We set a background color...
        Me.BackColor = Color.Fuchsia

        ' ...and then we tell Windows that this specific color should be 100% transparent.
        Me.TransparencyKey = Color.Fuchsia

        ' Set the opacity to zero for good measure, making it fully invisible
        Me.Opacity = 0

        ' This ensures you can control its position later if you need to
        Me.StartPosition = FormStartPosition.Manual
    End Sub
End Class