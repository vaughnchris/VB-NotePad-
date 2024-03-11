Public Class Splash
    Private Sub Splash_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Dim tmr As New Timer
        tmr.Interval = 1000
        AddHandler tmr.Tick, AddressOf tmr_Tick
        tmr.Start()
    End Sub

    Private Sub tmr_Tick(sender As Object, e As EventArgs)
        CType(sender, Timer).Stop()
        Me.Close()
    End Sub
End Class