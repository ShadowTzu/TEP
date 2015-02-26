Imports TEP
Public Class Form1

    Private DoRender As Boolean
    Private Tempo As TimedExecution

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If DoRender = True Then
            DoRender = False
            e.Cancel = True
        End If
        Timer1.Enabled = False
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Show()
        Me.Focus()

        Tempo = New TimedExecution


        Tempo.Add(AddressOf Waiting, 0, 2000)
        Tempo.Add(AddressOf FadeOut, 0, 3000)
        Tempo.Add(AddressOf FadeIn, 0, 3000)
        Tempo.Add(AddressOf Waiting, 0, 5000)

        Tempo.Add(AddressOf SizeInc, 0, 3000)
        Tempo.Add(AddressOf SizeDec, 0, 3000)


        Tempo.Add(AddressOf ExitApp, 0, 3000)
        'methode with a timer
        Timer1.Enabled = True

        'methode with a loop
        'DoRender = True
        'Render()

    End Sub

    Private Sub Waiting(cTime As Long, MaxTime As Long)
        Me.Text = "Waiting " & (CInt(MaxTime) \ 1000I).ToString & " second: " & cTime.ToString & "/" & MaxTime.ToString
    End Sub

    Private Sub FadeIn(cTime As Long, MaxTime As Long)
        Me.Text = "Fade In " & cTime.ToString & "/" & MaxTime.ToString
        Me.Opacity = cTime / MaxTime
    End Sub

    Private Sub FadeOut(cTime As Long, MaxTime As Long)
        Me.Text = "Fade Out " & cTime.ToString & "/" & MaxTime.ToString
        Me.Opacity = 1 - cTime / MaxTime
    End Sub

    Private Sub ExitApp(cTime As Long, MaxTime As Long)
        Me.Text = "Quit in " & MaxTime - cTime & "ms"
        If cTime >= MaxTime Then Application.Exit()
    End Sub

    Private Sub SizeDec(cTime As Long, MaxTime As Long)
        Me.Text = "Size decrease " & cTime.ToString & "/" & MaxTime.ToString
        Me.Width = 300 + 400 - CInt(400 * (cTime / MaxTime))
        Me.Height = 300 + 400 - CInt(400 * (cTime / MaxTime))
    End Sub

    Private Sub SizeInc(cTime As Long, MaxTime As Long)
        Me.Text = "Size increase " & cTime.ToString & "/" & MaxTime.ToString
        Me.Width = 300 + CInt(400 * (cTime / MaxTime))
        Me.Height = 300 + CInt(400 * (cTime / MaxTime))
    End Sub


    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Tempo.Process()
    End Sub

    Private Sub Render()
        Do While DoRender

            Tempo.Process()
            Application.DoEvents()
        Loop
        Application.Exit()
    End Sub

End Class
