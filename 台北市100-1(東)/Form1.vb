Public Class Form1
    Dim Box(6, 7) As Button
    Dim NowUser As Boolean
    Dim ck As Boolean
    'Dim SetNX() As Integer = {-1, -1, -1, 0, 0, 1, 1, 1} 'BOX(X,Y)
    'Dim SetNY() As Integer = {-1, 0, 1, -1, 1, -1, 0, 1}
    Dim SetNX() As Integer = {-1, 1, -1, 1, -1, 1, 0, 0} 'BOX(X,Y)
    Dim SetNY() As Integer = {-1, 1, 1, -1, 0, 0, -1, 1}
    'Dim fr As Integer
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        For i As Integer = 1 To 6
            For x As Integer = 1 To 7
                Box(i, x) = Me.Controls("Button" & (i - 1) * 7 + x)
                AddHandler Box(i, x).Click, AddressOf Tread
            Next
        Next
        Re()
    End Sub

    Sub Re() '重新
        NowUser = False : ck = False
        For i As Integer = 1 To 6
            For x As Integer = 1 To 7
                Box(i, x).BackColor = Color.SandyBrown
                Box(i, x).Text = ""
                Box(i, x).Tag = 0
                Box(i, x).Image = Nothing
            Next
        Next
    End Sub

    Sub Tread(ByVal sender As System.Object, ByVal e As System.EventArgs) '放棋子
        If ck = True Then Exit Sub
        Dim NowX, NowY, BlackFr, WhiteFr As Integer
        Dim Boxck(6, 7) As Boolean : ReDim Boxck(6, 7)
        NowX = (Val(Mid(CType(sender, Button).Name, 7)) - 1) \ 7 + 1
        NowY = (Val(Mid(CType(sender, Button).Name, 7)) - 1) Mod 7 + 1
        If Box(NowX, NowY).Tag = 0 Then
            If NowUser = False Then '白棋
                NowUser = True
                For Nx As Integer = 6 To 1 Step -1
                    If Box(Nx, NowY).Tag = 0 Then
                        Box(Nx, NowY).Tag = 1
                        Box(Nx, NowY).Image = New Bitmap("0.png")
                        WhiteFr = 0
                        For i As Integer = 0 To 7
                            Track(Nx, NowY, WhiteFr, 1, Boxck, i, False)
                            For x As Integer = 1 To 6
                                For y As Integer = 1 To 7
                                    Boxck(x, y) = False
                                Next
                            Next
                            If i Mod 2 <> 0 Then WhiteFr = 0
                        Next : Exit For
                    End If
                Next
            ElseIf NowUser = True Then '黑棋
                NowUser = False
                For Nx As Integer = 6 To 1 Step -1
                    If Box(Nx, NowY).Tag = 0 Then
                        Box(Nx, NowY).Tag = 2
                        Box(Nx, NowY).Image = New Bitmap("1.png")
                        BlackFr = 0
                        For i As Integer = 0 To 7
                            Track(Nx, NowY, BlackFr, 2, Boxck, i, False)
                            For x As Integer = 1 To 6
                                For y As Integer = 1 To 7
                                    Boxck(x, y) = False
                                Next
                            Next
                            If i Mod 2 <> 0 Then BlackFr = 0
                        Next
                        Exit For
                    End If
                Next
            End If
        End If
    End Sub

    Sub Track(ByVal NowX As Integer, ByVal NowY As Integer, ByRef Fr As Integer, ByVal Now As Integer, ByVal BoxCk(,) As Boolean, ByVal cnt As Integer, ByVal First As Boolean) '偵測 當前點選X,Y位子 FR=幾個棋子連載一起
        If ck = True Then Exit Sub
        If Fr = 3 Then '有連續的三顆(自己不算4-1=3)
            If Now = 1 Then
                MsgBox("白旗贏了") : ck = True : Exit Sub
            ElseIf Now = 2 Then
                MsgBox("黑旗贏了") : ck = True : Exit Sub
            End If
        End If
        If (NowX <= 0 Or NowX >= 7) Or (NowY <= 0 Or NowY >= 8) Then Exit Sub
        If Box(NowX, NowY).Tag = Now And BoxCk(NowX, NowY) = False Then
            If First = True Then Fr += 1 Else First = True '自己本身下的那顆不能算
            BoxCk(NowX, NowY) = True
            'Fr += 1
            Track(NowX + SetNX(cnt), NowY + SetNY(cnt), Fr, Now, BoxCk, cnt, First)
        End If
    End Sub
    Private Sub Button43_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button43.Click
        Re()
    End Sub
End Class
