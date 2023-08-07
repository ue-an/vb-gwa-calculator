Public Class Frm_GWA

    Private Sub Frm_GWA_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        For subs = 1 To 9
            cmbSubjects.Items.Add(subs)
        Next
    End Sub

    Private Sub ComputeAVE()
        Dim grdSum = Val(txtGrd01.Text) + Val(txtGrd02.Text) + Val(txtGrd03.Text) + Val(txtGrd04.Text) +
                   Val(txtGrd05.Text) + Val(txtGrd06.Text) + Val(txtGrd07.Text) + Val(txtGrd08.Text) + Val(txtGrd09.Text)
        Dim ave = grdSum / Val(cmbSubjects.Text)

        lblGrdSum.Text = grdSum
        lblAverage.Text = Format(ave, "#,###,##0.000")
    End Sub

    Private Sub ComputeGWA()

        Dim gw1 = Val(txtGrd01.Text) * Val(txtUnit01.Text)
        Dim gw2 = Val(txtGrd02.Text) * Val(txtUnit02.Text)
        Dim gw3 = Val(txtGrd03.Text) * Val(txtUnit03.Text)
        Dim gw4 = Val(txtGrd04.Text) * Val(txtUnit04.Text)
        Dim gw5 = Val(txtGrd05.Text) * Val(txtUnit05.Text)
        Dim gw6 = Val(txtGrd06.Text) * Val(txtUnit06.Text)
        Dim gw7 = Val(txtGrd07.Text) * Val(txtUnit07.Text)
        Dim gw8 = Val(txtGrd08.Text) * Val(txtUnit08.Text)
        Dim gw9 = Val(txtGrd09.Text) * Val(txtUnit09.Text)

        Dim gwSum = gw1 + gw2 + gw3 + gw4 + gw5 + gw6 + gw7 + gw8 + gw9
        Dim unitSum = Val(txtUnit01.Text) + Val(txtUnit02.Text) + Val(txtUnit03.Text) + Val(txtUnit04.Text) +
                      Val(txtUnit05.Text) + Val(txtUnit06.Text) + Val(txtUnit07.Text) + Val(txtUnit08.Text) +
                      Val(txtUnit09.Text)

        Dim gwa = gwSum / unitSum

        lblUnitSum.Text = unitSum
        lblGwSum.Text = gwSum
        lblGWA.Text = Format(gwa, "#,###,##0.000")
    End Sub

    Private Sub EnableGradeUnit01()
        txtGrd01.Enabled = True
        txtUnit01.Enabled = True
        txtGrd02.Enabled = False
        txtUnit02.Enabled = False
        txtGrd03.Enabled = False
        txtUnit03.Enabled = False
        txtGrd04.Enabled = False
        txtUnit04.Enabled = False
        txtGrd05.Enabled = False
        txtUnit05.Enabled = False
        txtGrd06.Enabled = False
        txtUnit06.Enabled = False
        txtGrd07.Enabled = False
        txtUnit07.Enabled = False
        txtGrd08.Enabled = False
        txtUnit08.Enabled = False
        txtGrd09.Enabled = False
        txtUnit09.Enabled = False
    End Sub

    Private Sub EnableGradeUnit02()
        txtGrd01.Enabled = True
        txtUnit01.Enabled = True
        txtGrd02.Enabled = True
        txtUnit02.Enabled = True
        txtGrd03.Enabled = False
        txtUnit03.Enabled = False
        txtGrd04.Enabled = False
        txtUnit04.Enabled = False
        txtGrd05.Enabled = False
        txtUnit05.Enabled = False
        txtGrd06.Enabled = False
        txtUnit06.Enabled = False
        txtGrd07.Enabled = False
        txtUnit07.Enabled = False
        txtGrd08.Enabled = False
        txtUnit08.Enabled = False
        txtGrd09.Enabled = False
        txtUnit09.Enabled = False
    End Sub

    Private Sub EnableGradeUnit03()
        txtGrd01.Enabled = True
        txtUnit01.Enabled = True
        txtGrd02.Enabled = True
        txtUnit02.Enabled = True
        txtGrd03.Enabled = True
        txtUnit03.Enabled = True
        txtGrd04.Enabled = False
        txtUnit04.Enabled = False
        txtGrd05.Enabled = False
        txtUnit05.Enabled = False
        txtGrd06.Enabled = False
        txtUnit06.Enabled = False
        txtGrd07.Enabled = False
        txtUnit07.Enabled = False
        txtGrd08.Enabled = False
        txtUnit08.Enabled = False
        txtGrd09.Enabled = False
        txtUnit09.Enabled = False
    End Sub

    Private Sub EnableGradeUnit04()
        txtGrd01.Enabled = True
        txtUnit01.Enabled = True
        txtGrd02.Enabled = True
        txtUnit02.Enabled = True
        txtGrd03.Enabled = True
        txtUnit03.Enabled = True
        txtGrd04.Enabled = True
        txtUnit04.Enabled = True
        txtGrd05.Enabled = False
        txtUnit05.Enabled = False
        txtGrd06.Enabled = False
        txtUnit06.Enabled = False
        txtGrd07.Enabled = False
        txtUnit07.Enabled = False
        txtGrd08.Enabled = False
        txtUnit08.Enabled = False
        txtGrd09.Enabled = False
        txtUnit09.Enabled = False
    End Sub

    Private Sub EnableGradeUnit05()
        txtGrd01.Enabled = True
        txtUnit01.Enabled = True
        txtGrd02.Enabled = True
        txtUnit02.Enabled = True
        txtGrd03.Enabled = True
        txtUnit03.Enabled = True
        txtGrd04.Enabled = True
        txtUnit04.Enabled = True
        txtGrd05.Enabled = True
        txtUnit05.Enabled = True
        txtGrd06.Enabled = False
        txtUnit06.Enabled = False
        txtGrd07.Enabled = False
        txtUnit07.Enabled = False
        txtGrd08.Enabled = False
        txtUnit08.Enabled = False
        txtGrd09.Enabled = False
        txtUnit09.Enabled = False
    End Sub

    Private Sub EnableGradeUnit06()
        txtGrd01.Enabled = True
        txtUnit01.Enabled = True
        txtGrd02.Enabled = True
        txtUnit02.Enabled = True
        txtGrd03.Enabled = True
        txtUnit03.Enabled = True
        txtGrd04.Enabled = True
        txtUnit04.Enabled = True
        txtGrd05.Enabled = True
        txtUnit05.Enabled = True
        txtGrd06.Enabled = True
        txtUnit06.Enabled = True
        txtGrd07.Enabled = False
        txtUnit07.Enabled = False
        txtGrd08.Enabled = False
        txtUnit08.Enabled = False
        txtGrd09.Enabled = False
        txtUnit09.Enabled = False
    End Sub

    Private Sub EnableGradeUnit07()
        txtGrd01.Enabled = True
        txtUnit01.Enabled = True
        txtGrd02.Enabled = True
        txtUnit02.Enabled = True
        txtGrd03.Enabled = True
        txtUnit03.Enabled = True
        txtGrd04.Enabled = True
        txtUnit04.Enabled = True
        txtGrd05.Enabled = True
        txtUnit05.Enabled = True
        txtGrd06.Enabled = True
        txtUnit06.Enabled = True
        txtGrd07.Enabled = True
        txtUnit07.Enabled = True
        txtGrd08.Enabled = False
        txtUnit08.Enabled = False
        txtGrd09.Enabled = False
        txtUnit09.Enabled = False
    End Sub

    Private Sub EnableGradeUnit08()
        txtGrd01.Enabled = True
        txtUnit01.Enabled = True
        txtGrd02.Enabled = True
        txtUnit02.Enabled = True
        txtGrd03.Enabled = True
        txtUnit03.Enabled = True
        txtGrd04.Enabled = True
        txtUnit04.Enabled = True
        txtGrd05.Enabled = True
        txtUnit05.Enabled = True
        txtGrd06.Enabled = True
        txtUnit06.Enabled = True
        txtGrd07.Enabled = True
        txtUnit07.Enabled = True
        txtGrd08.Enabled = True
        txtUnit08.Enabled = True
        txtGrd09.Enabled = False
        txtUnit09.Enabled = False
    End Sub

    Private Sub EnableGradeUnit09()
        txtGrd01.Enabled = True
        txtUnit01.Enabled = True
        txtGrd02.Enabled = True
        txtUnit02.Enabled = True
        txtGrd03.Enabled = True
        txtUnit03.Enabled = True
        txtGrd04.Enabled = True
        txtUnit04.Enabled = True
        txtGrd05.Enabled = True
        txtUnit05.Enabled = True
        txtGrd06.Enabled = True
        txtUnit06.Enabled = True
        txtGrd07.Enabled = True
        txtUnit07.Enabled = True
        txtGrd08.Enabled = True
        txtUnit08.Enabled = True
        txtGrd09.Enabled = True
        txtUnit09.Enabled = True
    End Sub

    Private Sub cmbSubjects_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbSubjects.SelectedIndexChanged
        Select Case cmbSubjects.Text
            Case 1
                EnableGradeUnit01()
            Case 2
                EnableGradeUnit01()
                EnableGradeUnit02()
            Case 3
                EnableGradeUnit01()
                EnableGradeUnit02()
                EnableGradeUnit03()
            Case 4
                EnableGradeUnit01()
                EnableGradeUnit02()
                EnableGradeUnit03()
                EnableGradeUnit04()
            Case 5
                EnableGradeUnit01()
                EnableGradeUnit02()
                EnableGradeUnit03()
                EnableGradeUnit04()
                EnableGradeUnit05()
            Case 6
                EnableGradeUnit01()
                EnableGradeUnit02()
                EnableGradeUnit03()
                EnableGradeUnit04()
                EnableGradeUnit05()
                EnableGradeUnit06()
            Case 7
                EnableGradeUnit01()
                EnableGradeUnit02()
                EnableGradeUnit03()
                EnableGradeUnit04()
                EnableGradeUnit05()
                EnableGradeUnit06()
                EnableGradeUnit07()
            Case 8
                EnableGradeUnit01()
                EnableGradeUnit02()
                EnableGradeUnit03()
                EnableGradeUnit04()
                EnableGradeUnit05()
                EnableGradeUnit06()
                EnableGradeUnit07()
                EnableGradeUnit08()
            Case 9
                EnableGradeUnit01()
                EnableGradeUnit02()
                EnableGradeUnit03()
                EnableGradeUnit04()
                EnableGradeUnit05()
                EnableGradeUnit06()
                EnableGradeUnit07()
                EnableGradeUnit08()
                EnableGradeUnit09()
        End Select
    End Sub

    Private Sub txtGrd01_TextChanged(sender As Object, e As EventArgs) Handles txtGrd01.TextChanged
        ComputeAVE()
    End Sub

    Private Sub txtGrd02_TextChanged(sender As Object, e As EventArgs) Handles txtGrd02.TextChanged
        ComputeAVE()
    End Sub

    Private Sub txtGrd03_TextChanged(sender As Object, e As EventArgs) Handles txtGrd03.TextChanged
        ComputeAVE()
    End Sub

    Private Sub txtGrd04_TextChanged(sender As Object, e As EventArgs) Handles txtGrd04.TextChanged
        ComputeAVE()
    End Sub

    Private Sub txtGrd05_TextChanged(sender As Object, e As EventArgs) Handles txtGrd05.TextChanged
        ComputeAVE()
    End Sub

    Private Sub txtGrd06_TextChanged(sender As Object, e As EventArgs) Handles txtGrd06.TextChanged
        ComputeAVE()
    End Sub

    Private Sub txtGrd07_TextChanged(sender As Object, e As EventArgs) Handles txtGrd07.TextChanged
        ComputeAVE()
    End Sub

    Private Sub txtGrd08_TextChanged(sender As Object, e As EventArgs) Handles txtGrd08.TextChanged
        ComputeAVE()
    End Sub

    Private Sub txtGrd09_TextChanged(sender As Object, e As EventArgs) Handles txtGrd09.TextChanged
        ComputeAVE()
    End Sub


    Private Sub txtUnit01_TextChanged(sender As Object, e As EventArgs) Handles txtUnit01.TextChanged
        ComputeGWA()
    End Sub

    Private Sub txtUnit02_TextChanged(sender As Object, e As EventArgs) Handles txtUnit02.TextChanged
        ComputeGWA()
    End Sub

    Private Sub txtUnit03_TextChanged(sender As Object, e As EventArgs) Handles txtUnit03.TextChanged
        ComputeGWA()
    End Sub

    Private Sub txtUnit04_TextChanged(sender As Object, e As EventArgs) Handles txtUnit04.TextChanged
        ComputeGWA()
    End Sub

    Private Sub txtUnit05_TextChanged(sender As Object, e As EventArgs) Handles txtUnit05.TextChanged
        ComputeGWA()
    End Sub

    Private Sub txtUnit06_TextChanged(sender As Object, e As EventArgs) Handles txtUnit06.TextChanged
        ComputeGWA()
    End Sub

    Private Sub txtUnit07_TextChanged(sender As Object, e As EventArgs) Handles txtUnit07.TextChanged
        ComputeGWA()
    End Sub

    Private Sub txtUnit08_TextChanged(sender As Object, e As EventArgs) Handles txtUnit08.TextChanged
        ComputeGWA()
    End Sub

    Private Sub txtUnit09_TextChanged(sender As Object, e As EventArgs) Handles txtUnit09.TextChanged
        ComputeGWA()
    End Sub
End Class
