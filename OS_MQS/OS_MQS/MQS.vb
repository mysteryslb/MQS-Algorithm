Public Class MQS

    Private Sub MQS_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        pn1.Text = 0
        pn2.Text = 0
        pn3.Text = 0
        pn4.Text = 0
        pn5.Text = 0
        bt1.Text = 0
        bt2.Text = 0
        bt3.Text = 0
        bt4.Text = 0
        bt5.Text = 0
        btl1.Text = 0
        btl2.Text = 0
        btl3.Text = 0
        btl4.Text = 0
        btl5.Text = 0
        pp1.Text = 0
        pp2.Text = 0
        pp3.Text = 0
        pp4.Text = 0
        pp5.Text = 0
        dp1.Text = 0
        dp2.Text = 0
        dp3.Text = 0
        dp4.Text = 0
        dp5.Text = 0
        pl1.Text = 0
        pl2.Text = 0
        pl3.Text = 0
        pl4.Text = 0
        pl5.Text = 0
        pr1.Text = 0
        pr2.Text = 0
        pr3.Text = 0
        pr4.Text = 0
        pr5.Text = 0
        i1.Text = 0
        i2.Text = 0
        i3.Text = 0
        i4.Text = 0
        i5.Text = 0
        r1.Text = 0
        r2.Text = 0
        r3.Text = 0
        r4.Text = 0
        r5.Text = 0
    End Sub
    Private Sub refresh_Click(sender As Object, e As EventArgs) Handles refresh.Click

        s1.Text = p1.Text
        s2.Text = p2.Text
        s3.Text = p3.Text
        s4.Text = p4.Text
        s5.Text = p5.Text
        s6.Text = p6.Text
        s7.Text = p7.Text
        s8.Text = p8.Text
        s9.Text = p9.Text
        s10.Text = p10.Text
        s11.Text = p11.Text
        s12.Text = p12.Text
        s13.Text = p13.Text
        s14.Text = p14.Text
        s15.Text = p15.Text
        s16.Text = p16.Text
        s17.Text = p17.Text
        s18.Text = p18.Text
        s19.Text = p19.Text
        s20.Text = p20.Text
        s21.Text = p21.Text
        s22.Text = p22.Text
        s23.Text = p23.Text
        s24.Text = p24.Text
        s25.Text = p25.Text
        q1.Text = p1.Text
        q2.Text = p2.Text
        q3.Text = p3.Text
        q4.Text = p4.Text
        q5.Text = p5.Text
        q6.Text = p6.Text
        q7.Text = p7.Text
        q8.Text = p8.Text
        q9.Text = p9.Text
        q10.Text = p10.Text
        q11.Text = p11.Text
        q12.Text = p12.Text
        q13.Text = p13.Text
        q14.Text = p14.Text
        q15.Text = p15.Text
        q16.Text = p16.Text
        q17.Text = p17.Text
        q18.Text = p18.Text
        q19.Text = p19.Text
        q20.Text = p20.Text
        q21.Text = p21.Text
        q22.Text = p22.Text
        q23.Text = p23.Text
        q24.Text = p24.Text
        q25.Text = p25.Text
        l1.Text = mp1.Text
        l2.Text = mp2.Text
        l3.Text = mp3.Text
        l4.Text = mp4.Text
        l5.Text = mp5.Text
        k1.Text = pr1.Text
        k2.Text = pr2.Text
        k3.Text = pr3.Text
        k4.Text = pr4.Text
        k5.Text = pr5.Text
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        pl1.Text = pn1.Text
        pl2.Text = pn2.Text
        pl3.Text = pn3.Text
        pl4.Text = pn4.Text
        pl5.Text = pn5.Text
        btl1.Text = bt1.Text
        btl2.Text = bt2.Text
        btl3.Text = bt3.Text
        btl4.Text = bt4.Text
        btl5.Text = bt5.Text

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If ((pn1.Text = 1) Or (pn1.Text = 2) Or (pn1.Text = 3) Or (pn1.Text = 4) Or (pn1.Text = 5)) Then
            pp1.Text = pr1.Text
        ElseIf ((pn1.Text = 6) Or (pn1.Text = 7) Or (pn1.Text = 8) Or (pn1.Text = 9) Or (pn1.Text = 10)) Then
            pp1.Text = pr2.Text
        ElseIf ((pn1.Text = 11) Or (pn1.Text = 12) Or (pn1.Text = 13) Or (pn1.Text = 14) Or (pn1.Text = 15)) Then
            pp1.Text = pr3.Text
        ElseIf ((pn1.Text = 16) Or (pn1.Text = 17) Or (pn1.Text = 18) Or (pn1.Text = 19) Or (pn1.Text = 20)) Then
            pp1.Text = pr4.Text
        ElseIf ((pn1.Text = 21) Or (pn1.Text = 22) Or (pn1.Text = 23) Or (pn1.Text = 24) Or (pn1.Text = 25)) Then
            pp1.Text = pr5.Text
        Else
            pp1.Text = 0
        End If

        If ((pn2.Text = 1) Or (pn2.Text = 2) Or (pn2.Text = 3) Or (pn2.Text = 4) Or (pn2.Text = 5)) Then
            pp2.Text = pr1.Text
        ElseIf ((pn2.Text = 6) Or (pn2.Text = 7) Or (pn2.Text = 8) Or (pn2.Text = 9) Or (pn2.Text = 10)) Then
            pp2.Text = pr2.Text
        ElseIf ((pn2.Text = 11) Or (pn2.Text = 12) Or (pn2.Text = 13) Or (pn2.Text = 14) Or (pn2.Text = 15)) Then
            pp2.Text = pr3.Text
        ElseIf ((pn2.Text = 16) Or (pn2.Text = 17) Or (pn2.Text = 18) Or (pn2.Text = 19) Or (pn2.Text = 20)) Then
            pp2.Text = pr4.Text
        ElseIf ((pn2.Text = 21) Or (pn2.Text = 22) Or (pn2.Text = 23) Or (pn2.Text = 24) Or (pn2.Text = 25)) Then
            pp2.Text = pr5.Text
        Else
            pp2.Text = 0
        End If

        If ((pn3.Text = 1) Or (pn3.Text = 2) Or (pn3.Text = 3) Or (pn3.Text = 4) Or (pn3.Text = 5)) Then
            pp3.Text = pr1.Text
        ElseIf ((pn3.Text = 6) Or (pn3.Text = 7) Or (pn3.Text = 8) Or (pn3.Text = 9) Or (pn3.Text = 10)) Then
            pp3.Text = pr2.Text
        ElseIf ((pn3.Text = 11) Or (pn3.Text = 12) Or (pn3.Text = 13) Or (pn3.Text = 14) Or (pn3.Text = 15)) Then
            pp3.Text = pr3.Text
        ElseIf ((pn3.Text = 16) Or (pn3.Text = 17) Or (pn3.Text = 18) Or (pn3.Text = 19) Or (pn3.Text = 20)) Then
            pp3.Text = pr4.Text
        ElseIf ((pn3.Text = 21) Or (pn3.Text = 22) Or (pn3.Text = 23) Or (pn3.Text = 24) Or (pn3.Text = 25)) Then
            pp3.Text = pr5.Text
        Else
            pp3.Text = 0
        End If

        If ((pn4.Text = 1) Or (pn4.Text = 2) Or (pn4.Text = 3) Or (pn4.Text = 4) Or (pn4.Text = 5)) Then
            pp4.Text = pr1.Text
        ElseIf ((pn4.Text = 6) Or (pn4.Text = 7) Or (pn4.Text = 8) Or (pn4.Text = 9) Or (pn4.Text = 10)) Then
            pp4.Text = pr2.Text
        ElseIf ((pn4.Text = 11) Or (pn4.Text = 12) Or (pn4.Text = 13) Or (pn4.Text = 14) Or (pn4.Text = 15)) Then
            pp4.Text = pr3.Text
        ElseIf ((pn4.Text = 16) Or (pn4.Text = 17) Or (pn4.Text = 18) Or (pn4.Text = 19) Or (pn4.Text = 20)) Then
            pp4.Text = pr4.Text
        ElseIf ((pn4.Text = 21) Or (pn4.Text = 22) Or (pn4.Text = 23) Or (pn4.Text = 24) Or (pn4.Text = 25)) Then
            pp4.Text = pr5.Text
        Else
            pp4.Text = 0
        End If

        If ((pn5.Text = 1) Or (pn5.Text = 2) Or (pn5.Text = 3) Or (pn5.Text = 4) Or (pn5.Text = 5)) Then
            pp5.Text = pr1.Text
        ElseIf ((pn5.Text = 6) Or (pn5.Text = 7) Or (pn5.Text = 8) Or (pn5.Text = 9) Or (pn5.Text = 10)) Then
            pp5.Text = pr2.Text
        ElseIf ((pn5.Text = 11) Or (pn5.Text = 12) Or (pn5.Text = 13) Or (pn5.Text = 14) Or (pn5.Text = 15)) Then
            pp5.Text = pr3.Text
        ElseIf ((pn5.Text = 16) Or (pn5.Text = 17) Or (pn5.Text = 18) Or (pn5.Text = 19) Or (pn5.Text = 20)) Then
            pp5.Text = pr4.Text
        ElseIf ((pn5.Text = 21) Or (pn5.Text = 22) Or (pn5.Text = 23) Or (pn5.Text = 24) Or (pn5.Text = 25)) Then
            pp5.Text = pr5.Text
        Else
            pp5.Text = 0
        End If

        i1.Text = pp1.Text
        r1.Text = pp1.Text
        i2.Text = pp2.Text
        r2.Text = pp2.Text
        i3.Text = pp3.Text
        r3.Text = pp3.Text
        i4.Text = pp4.Text
        r4.Text = pp4.Text
        i5.Text = pp5.Text
        r5.Text = pp5.Text
        Dim arrA() As TextBox = {pl1, pl2, pl3, pl4, pl5}
        Dim arrB() As TextBox = {r1, r2, r3, r4, r5}
        Dim Changed As Boolean
        Do
            Changed = False
            For i = 0 To arrB.Count - 2 'Stop at the second to last array item because we check forward in the array
                If CInt(arrB(i).Text) > CInt(arrB(i + 1).Text) Then 'Next value is smaller than previous --> Switch values, also switch in arrA
                    Dim Temp As String = arrB(i + 1).Text
                    arrB(i + 1).Text = arrB(i).Text
                    arrB(i).Text = Temp
                    Temp = arrA(i + 1).Text
                    arrA(i + 1).Text = arrA(i).Text
                    arrA(i).Text = Temp
                    Changed = True

                End If
            Next
        Loop Until Changed = False 'Cancle the loop when everything is sorted
        Dim arrL() As TextBox = {dp1, dp2, dp3, dp4, dp5}
        For i = 0 To 4
            arrL(i).Text = arrA(i).Text
        Next

        Dim arrC() As TextBox = {i1, i2, i3, i4, i5}
        Dim arrD() As TextBox = {btl1, btl2, btl3, btl4, btl5}
        Dim Changeda As Boolean
        Do
            Changeda = False
            For i = 0 To arrC.Count - 2 'Stop at the second to last array item because we check forward in the array
                If CInt(arrC(i).Text) > CInt(arrC(i + 1).Text) Then 'Next value is smaller than previous --> Switch values, also switch in arrA
                    Dim Temp As String = arrC(i + 1).Text
                    arrC(i + 1).Text = arrC(i).Text
                    arrC(i).Text = Temp
                    Temp = arrD(i + 1).Text
                    arrD(i + 1).Text = arrD(i).Text
                    arrD(i).Text = Temp
                    Changeda = True
                End If
            Next
        Loop Until Changeda = False 'Cancle the loop when everything is sorted
        For i = 0 To 4
            arrL(i).Text = arrA(i).Text
        Next
        wt1.Text = 0
        wt2.Text = (Convert.ToInt32(btl1.Text))
        wt3.Text = (Convert.ToInt32(btl1.Text)) + (Convert.ToInt32(btl2.Text))
        wt4.Text = (Convert.ToInt32(btl1.Text)) + (Convert.ToInt32(btl2.Text)) + (Convert.ToInt32(btl3.Text))
        wt5.Text = (Convert.ToInt32(btl1.Text)) + (Convert.ToInt32(btl2.Text)) + (Convert.ToInt32(btl3.Text)) + (Convert.ToInt32(btl4.Text))

        tt1.Text = (Convert.ToInt32(btl1.Text))
        tt2.Text = (Convert.ToInt32(btl1.Text)) + (Convert.ToInt32(btl2.Text))
        tt3.Text = (Convert.ToInt32(btl1.Text)) + (Convert.ToInt32(btl2.Text)) + (Convert.ToInt32(btl3.Text))
        tt4.Text = (Convert.ToInt32(btl1.Text)) + (Convert.ToInt32(btl2.Text)) + (Convert.ToInt32(btl3.Text)) + (Convert.ToInt32(btl4.Text))
        tt5.Text = (Convert.ToInt32(btl1.Text)) + (Convert.ToInt32(btl2.Text)) + (Convert.ToInt32(btl3.Text)) + (Convert.ToInt32(btl4.Text)) + (Convert.ToInt32(btl5.Text))

        awt.Text = (((Convert.ToInt32(wt1.Text)) + (Convert.ToInt32(wt2.Text)) + (Convert.ToInt32(wt3.Text)) + (Convert.ToInt32(wt4.Text)) + (Convert.ToInt32(wt5.Text))) / 5)

        att.Text = (((Convert.ToInt32(tt1.Text)) + (Convert.ToInt32(tt2.Text)) + (Convert.ToInt32(tt3.Text)) + (Convert.ToInt32(tt4.Text)) + (Convert.ToInt32(tt5.Text))) / 5)
    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles g7.TextChanged

    End Sub

    Private Sub TextBox55_TextChanged(sender As Object, e As EventArgs) Handles g44.TextChanged

    End Sub

    Private Sub TextBox53_TextChanged(sender As Object, e As EventArgs) Handles g45.TextChanged

    End Sub
End Class
