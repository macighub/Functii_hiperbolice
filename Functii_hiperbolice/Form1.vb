Imports System.Drawing.Drawing2D
Imports System.Runtime.Intrinsics.X86

Public Class Form1
    Private b_Shown As Boolean = False
    Private val_Fct(12) As List(Of KeyValuePair(Of Double, Double))
    Private bmp_GRF(12) As Bitmap
    Private p_axes(,) As Integer = {{1, -1, 1, -1},
                                    {1, -1, 1, -1},
                                    {-1, -1, 1, 1},
                                    {1, -1, 1, -1},
                                    {1, -1, 1, -1},
                                    {1, -1, 1, -1},
                                    {1, -1, 1, -1},
                                    {1, -1, 1, -1},
                                    {-1, -1, 1, 1},
                                    {1, -1, 1, -1},
                                    {1, -1, 1, -1},
                                    {1, -1, 1, -1}}
    Private Colors() As Color = New Color() {Color.Lime,
                                             Color.Green,
                                             Color.Cyan,
                                             Color.RoyalBlue,
                                             Color.FromArgb(255, 128, 0),
                                             Color.Red,
                                             Color.FromArgb(255, 128, 255),
                                             Color.FromArgb(192, 0, 192),
                                             Color.Yellow,
                                             Color.Olive,
                                             Color.White,
                                             Color.DarkGray}

    Private Function CheckedFunctionCount() As Integer
        Dim l_CheckedFunctionCount As Integer = 0

        If chk_sh.Checked Then l_CheckedFunctionCount += 1
        If chk_arcsh.Checked Then l_CheckedFunctionCount += 1
        If chk_ch.Checked Then l_CheckedFunctionCount += 1
        If chk_arcch.Checked Then l_CheckedFunctionCount += 1
        If chk_th.Checked Then l_CheckedFunctionCount += 1
        If chk_argth.Checked Then l_CheckedFunctionCount += 1
        If chk_cth.Checked Then l_CheckedFunctionCount += 1
        If chk_argcth.Checked Then l_CheckedFunctionCount += 1
        If chk_sech.Checked Then l_CheckedFunctionCount += 1
        If chk_arsech.Checked Then l_CheckedFunctionCount += 1
        If chk_csch.Checked Then l_CheckedFunctionCount += 1
        If chk_acsch.Checked Then l_CheckedFunctionCount += 1

        Return l_CheckedFunctionCount
    End Function

    Private Function CheckedFunctions() As KeyValuePair(Of String, Boolean)()
        Dim l_CheckedFunctions(12) As KeyValuePair(Of String, Boolean)

        l_CheckedFunctions(0) = New KeyValuePair(Of String, Boolean)("sinh / sh", chk_sh.Checked)
        l_CheckedFunctions(1) = New KeyValuePair(Of String, Boolean)("arsinh / arcsh / argsinh / argsh", chk_arcsh.Checked)
        l_CheckedFunctions(2) = New KeyValuePair(Of String, Boolean)("cosh / ch", chk_ch.Checked)
        l_CheckedFunctions(3) = New KeyValuePair(Of String, Boolean)("arccosh / arcch / argcosh / argch", chk_arcch.Checked)
        l_CheckedFunctions(4) = New KeyValuePair(Of String, Boolean)("tanh/ th", chk_th.Checked)
        l_CheckedFunctions(5) = New KeyValuePair(Of String, Boolean)("arctanh / argth", chk_argth.Checked)
        l_CheckedFunctions(6) = New KeyValuePair(Of String, Boolean)("coth / cth", chk_cth.Checked)
        l_CheckedFunctions(7) = New KeyValuePair(Of String, Boolean)("arccoth / argcth", chk_argcth.Checked)
        l_CheckedFunctions(8) = New KeyValuePair(Of String, Boolean)("sech", chk_sech.Checked)
        l_CheckedFunctions(9) = New KeyValuePair(Of String, Boolean)("arcsech / arsech / asech", chk_arsech.Checked)
        l_CheckedFunctions(10) = New KeyValuePair(Of String, Boolean)("cosech / csch", chk_csch.Checked)
        l_CheckedFunctions(11) = New KeyValuePair(Of String, Boolean)("arccsch / arcsch / acsch", chk_acsch.Checked)

        Return l_CheckedFunctions
    End Function

    Private Function GridSize() As Size
        Dim l_GridSize As Size = New Size(0, 0)
        Dim l_CheckedFunctionCount As Integer = CheckedFunctionCount()

        l_GridSize.Width = Math.Ceiling(Math.Sqrt(l_CheckedFunctionCount))
        If l_GridSize.Width > 0 Then
            l_GridSize.Height = Math.Ceiling(l_CheckedFunctionCount / l_GridSize.Width)
        End If

        Return l_GridSize
    End Function

    Private Function GetLimits() As KeyValuePair(Of Double, Double)
        Dim l_Limits As KeyValuePair(Of Double, Double)

        If chk_Single.Checked Then
            If pic_Graphics.Width / pic_Graphics.Height >= 1.5 Then
                l_Limits = New KeyValuePair(Of Double, Double)(0 - Math.Round(((pic_Graphics.Width - 325) / 75), 2), Math.Round(((pic_Graphics.Width - 325) / 75), 2))
            Else
                l_Limits = New KeyValuePair(Of Double, Double)(0 - Math.Round(((pic_Graphics.Width - 20) / 75), 2), Math.Round(((pic_Graphics.Width - 20) / 75), 2))
            End If
        Else
            Dim l_GridSize As Size = GridSize()

            l_Limits = New KeyValuePair(Of Double, Double)(0 - Math.Round((((pic_Graphics.Width / l_GridSize.Width) - 20) / 75), 2), Math.Round((((pic_Graphics.Width / l_GridSize.Width) - 20) / 75), 2))
        End If

        Return l_Limits
    End Function

    Private Sub pic_Graphics_Paint(sender As Object, e As PaintEventArgs) Handles pic_Graphics.Paint
        Dim l_CheckedFunctionCount As Integer = CheckedFunctionCount()
        Dim l_CheckedFunctions() As KeyValuePair(Of String, Boolean) = CheckedFunctions()
        Dim l_GridSize As Size = GridSize()

        Dim sizeFactor As Double
        Dim sizeFactor_x As Double
        Dim sizeFactor_y As Double
        Dim v_VOffset As Integer
        Dim v_HOffset As Integer

        'CreateGraphs()

        Dim grfArea As Rectangle
        Dim lgdAtTop As Boolean

        If chk_Single.Checked Then
            Dim lgdSize As Size = New Size(CInt(Math.Ceiling(l_CheckedFunctionCount / 4)), CInt(Math.Ceiling(l_CheckedFunctionCount / 3)))
            lgdSize = New Size(l_GridSize.Height, l_GridSize.Width)

            If pic_Graphics.Width / pic_Graphics.Height >= 1.5 Then
                grfArea = New Rectangle(10, 10, pic_Graphics.Width - 325, pic_Graphics.Height - 20)
                lgdAtTop = False
            Else
                grfArea = New Rectangle(10, lgdSize.Height * 25 + 10, pic_Graphics.Width - 20, pic_Graphics.Height - lgdSize.Height * 25 - 20)
                lgdAtTop = True
            End If

            e.Graphics.FillRectangle(New SolidBrush(Color.FromArgb(25, 25, 25)),
                                     0,
                                     0,
                                     pic_Graphics.Width - 1,
                                     pic_Graphics.Height - 1)
            If l_CheckedFunctionCount > 0 Then
                If lgdAtTop Then
                    Dim grfIDX As Integer = 0
                    For grfCnt = 0 To 11
                        If l_CheckedFunctions(grfCnt).Value Then
                            If grfIDX Mod lgdSize.Height = 0 Then
                                e.Graphics.FillRectangle(New Drawing2D.LinearGradientBrush(New Rectangle(CInt(Math.Floor(grfIDX / lgdSize.Height) * (pic_Graphics.Width / lgdSize.Width)),
                                                                                                     0,
                                                                                                     CInt(pic_Graphics.Width / lgdSize.Width),
                                                                                                     lgdSize.Height * 25),
                                                                                       Color.FromArgb(64, 64, 64),
                                                                                       Color.Black,
                                                                                       Drawing2D.LinearGradientMode.Vertical),
                                                    CInt(Math.Floor(grfIDX / lgdSize.Height) * (pic_Graphics.Width / lgdSize.Width)),
                                                     0,
                                                     CInt(pic_Graphics.Width / lgdSize.Width),
                                                     lgdSize.Height * 25)
                            End If
                            e.Graphics.FillRectangle(New SolidBrush(Colors(grfCnt)),
                                                 CInt(Math.Floor(grfIDX / lgdSize.Height) * (pic_Graphics.Width / lgdSize.Width)) + 10,
                                                 CInt(grfIDX Mod lgdSize.Height) * 25 + 11,
                                                 25,
                                                 3)
                            e.Graphics.DrawString(l_CheckedFunctions(grfCnt).Key,
                                                  sender.font,
                                                  New SolidBrush(Color.White),
                                                  CInt(Math.Floor(grfIDX / lgdSize.Height) * (pic_Graphics.Width / lgdSize.Width)) + 40,
                                                  CInt(grfIDX Mod lgdSize.Height) * 25 + 2)
                            grfIDX += 1
                        End If
                    Next
                Else
                    e.Graphics.FillRectangle(New Drawing2D.LinearGradientBrush(New Rectangle(pic_Graphics.Width - 285,
                                                                                             CInt((pic_Graphics.Height / 2) - ((l_CheckedFunctionCount / 2) * 25)),
                                                                                             255,
                                                                                             l_CheckedFunctionCount * 25),
                                                                               Color.FromArgb(64, 64, 64),
                                                                               Color.Black,
                                                                               Drawing2D.LinearGradientMode.Vertical),
                                                 pic_Graphics.Width - 285,
                                                 CInt((pic_Graphics.Height / 2) - ((l_CheckedFunctionCount / 2) * 25)),
                                                 255,
                                                 l_CheckedFunctionCount * 25)
                    Dim grfIDX As Integer = 0
                    For grfCnt = 0 To 11
                        If l_CheckedFunctions(grfCnt).Value Then
                            e.Graphics.FillRectangle(New SolidBrush(Colors(grfCnt)),
                                                 pic_Graphics.Width - 275,
                                                 CInt((pic_Graphics.Height / 2) - ((l_CheckedFunctionCount / 2) * 25)) + grfIDX * 25 + 11,
                                                 25,
                                                 3)
                            e.Graphics.DrawString(l_CheckedFunctions(grfCnt).Key,
                                                      sender.font,
                                                      New SolidBrush(Color.White),
                                                      pic_Graphics.Width - 245,
                                                      CInt((pic_Graphics.Height / 2) - ((l_CheckedFunctionCount / 2) * 25)) + grfIDX * 25 + 2)
                            grfIDX += 1
                        End If
                    Next
                End If
            End If
            e.Graphics.DrawRectangle(New Pen(Color.Silver),
                                     0,
                                     0,
                                     pic_Graphics.Width - 1,
                                     pic_Graphics.Height - 1)
            e.Graphics.DrawLine(New Pen(Color.DimGray, 3),
                                grfArea.Left,
                                grfArea.Top + CInt(grfArea.Height / 2),
                                grfArea.Left + grfArea.Width,
                                grfArea.Top + CInt(grfArea.Height / 2))
            e.Graphics.DrawLine(New Pen(Color.DimGray, 3),
                                CInt(grfArea.Width / 2) + grfArea.Left,
                                grfArea.Top + grfArea.Height,
                                CInt(grfArea.Width / 2) + grfArea.Left,
                                grfArea.Top)
            For locCnt = 1 To 5
                e.Graphics.DrawLine(New Pen(Color.DimGray),
                                grfArea.Left + grfArea.Width - 10,
                                grfArea.Top + CInt(grfArea.Height / 2) - locCnt,
                                grfArea.Left + grfArea.Width,
                                grfArea.Top + CInt(grfArea.Height / 2))
                e.Graphics.DrawLine(New Pen(Color.DimGray),
                                grfArea.Left + grfArea.Width - 10,
                                grfArea.Top + CInt(grfArea.Height / 2) + locCnt,
                                grfArea.Left + grfArea.Width,
                                grfArea.Top + CInt(grfArea.Height / 2))
                e.Graphics.DrawLine(New Pen(Color.DimGray),
                                grfArea.Left + CInt(grfArea.Width / 2) - locCnt,
                                grfArea.Top + 10,
                                grfArea.Left + CInt(grfArea.Width / 2),
                                grfArea.Top)
                e.Graphics.DrawLine(New Pen(Color.DimGray),
                                    grfArea.Left + CInt(grfArea.Width / 2) + locCnt,
                                    grfArea.Top + 10,
                                    grfArea.Left + CInt(grfArea.Width / 2),
                                    grfArea.Top)
            Next

            Try
                Dim l_From As Double = Double.Parse(txt_From.Text, Globalization.NumberStyles.Any, Globalization.CultureInfo.CurrentCulture)
                Dim l_To As Double = Double.Parse(txt_To.Text, Globalization.NumberStyles.Any, Globalization.CultureInfo.CurrentCulture)

                For grfCnt = 0 To 11
                    If l_CheckedFunctions(grfCnt).Value Then
                        sizeFactor_x = ((grfArea.Width) / (l_To - l_From))
                        sizeFactor_y = ((grfArea.Height) / (val_Fct(grfCnt)(0).Value - val_Fct(grfCnt)(0).Key))
                        If Double.IsNaN(sizeFactor_y) Then
                            sizeFactor_y = sizeFactor_x
                        End If

                        If Double.IsNaN(sizeFactor) OrElse sizeFactor = 0 Then
                            sizeFactor = Math.Min(sizeFactor_x, sizeFactor_y)
                        Else
                            sizeFactor = Math.Min(sizeFactor, Math.Min(sizeFactor_x, sizeFactor_y))
                        End If
                    End If
                Next

                Dim v_step As Double = 0.0001
                While (v_step * sizeFactor_x) < 25
                    If Replace(Replace(Replace(v_step.ToString, "0", ""), ".", ""), ",", "") = "1" Then
                        v_step *= 5
                    Else
                        v_step *= 2
                    End If
                End While
                Dim lineCnt As Double = v_step
                While (lineCnt * sizeFactor_x) <= ((grfArea.Width / 2) - 10)
                    lineCnt = Math.Round(lineCnt, 2)
                    e.Graphics.DrawLine(New Pen(Color.FromArgb(10, 255, 255, 255), 1),
                                grfArea.Left + CInt((grfArea.Width / 2) - (lineCnt * sizeFactor_x)),
                                grfArea.Top,
                                grfArea.Left + CInt((grfArea.Width / 2) - (lineCnt * sizeFactor_x)),
                                grfArea.Top + grfArea.Height)
                    e.Graphics.DrawLine(New Pen(Color.FromArgb(10, 255, 255, 255), 1),
                                grfArea.Left + CInt((grfArea.Width / 2) + (lineCnt * sizeFactor_x)),
                                grfArea.Top,
                                grfArea.Left + CInt((grfArea.Width / 2) + (lineCnt * sizeFactor_x)),
                                grfArea.Top + grfArea.Height)
                    If CInt((grfArea.Width / 2) - (lineCnt * sizeFactor_x)) - 6 > 0 Then
                        v_HOffset = CInt(e.Graphics.MeasureString((-lineCnt).ToString, New Font(Me.Font.FontFamily, Math.Max(Math.Min(sizeFactor / 4, 10), 6))).Width / 2)
                        v_VOffset = e.Graphics.MeasureString((-lineCnt).ToString, New Font(Me.Font.FontFamily, Math.Max(Math.Min(sizeFactor / 4, 10), 6))).Height
                        e.Graphics.DrawString((-lineCnt).ToString,
                                        New Font(Me.Font.FontFamily, Math.Max(Math.Min(sizeFactor / 4, 10), 6)),
                                        New SolidBrush(Color.White),
                                        grfArea.Left + CInt((grfArea.Width / 2) - (lineCnt * sizeFactor_x)) - v_HOffset,
                                        grfArea.Top + CInt((grfArea.Height / 2) - v_VOffset) - 3)
                        v_HOffset = CInt(e.Graphics.MeasureString(lineCnt.ToString, New Font(Me.Font.FontFamily, Math.Max(Math.Min(sizeFactor / 4, 10), 6))).Width / 2)
                        v_VOffset = e.Graphics.MeasureString(lineCnt.ToString, New Font(Me.Font.FontFamily, Math.Max(Math.Min(sizeFactor / 4, 10), 6))).Height
                        e.Graphics.DrawString(lineCnt.ToString,
                                        New Font(Me.Font.FontFamily, Math.Max(Math.Min(sizeFactor / 4, 10), 6)),
                                        New SolidBrush(Color.White),
                                        grfArea.Left + CInt((grfArea.Width / 2) + (lineCnt * sizeFactor_x)) - v_HOffset,
                                        grfArea.Top + CInt((grfArea.Height / 2) + 6))
                    End If
                    lineCnt += v_step
                End While

                v_step = 0.001
                While (v_step * sizeFactor) < 15
                    If Replace(Replace(Replace(v_step.ToString, "0", ""), ".", ""), ",", "") = "1" Then
                        v_step *= 5
                    Else
                        v_step *= 2
                    End If
                End While
                lineCnt = v_step
                While (lineCnt * sizeFactor) <= ((grfArea.Height / 2) - 10)
                    e.Graphics.DrawLine(New Pen(Color.FromArgb(10, 255, 255, 255), 1),
                                grfArea.Left,
                                grfArea.Top + CInt((grfArea.Height / 2) + (lineCnt * sizeFactor)),
                                grfArea.Left + grfArea.Width,
                                grfArea.Top + CInt((grfArea.Height / 2) + (lineCnt * sizeFactor)))
                    e.Graphics.DrawLine(New Pen(Color.FromArgb(10, 255, 255, 255), 1),
                                grfArea.Left,
                                grfArea.Top + CInt((grfArea.Height / 2) - (lineCnt * sizeFactor)),
                                grfArea.Left + grfArea.Width,
                                grfArea.Top + CInt((grfArea.Height / 2) - (lineCnt * sizeFactor)))
                    If CInt((grfArea.Height / 2) - (lineCnt * sizeFactor)) - 6 > 0 Then
                        v_HOffset = e.Graphics.MeasureString((-lineCnt).ToString, New Font(Me.Font.FontFamily, Math.Max(Math.Min(sizeFactor / 4, 10), 6))).Width
                        v_VOffset = CInt(e.Graphics.MeasureString((-lineCnt).ToString, New Font(Me.Font.FontFamily, Math.Max(Math.Min(sizeFactor / 4, 10), 6))).Height / 2)
                        e.Graphics.DrawString((-lineCnt).ToString,
                                    New Font(Me.Font.FontFamily, Math.Max(Math.Min(sizeFactor / 4, 10), 6)),
                                    New SolidBrush(Color.White),
                                    grfArea.Left + CInt(grfArea.Width / 2) + 3,
                                    grfArea.Top + CInt((grfArea.Height / 2) + (lineCnt * sizeFactor)) - v_VOffset)
                        v_HOffset = e.Graphics.MeasureString(lineCnt.ToString, New Font(Me.Font.FontFamily, Math.Max(Math.Min(sizeFactor / 4, 10), 6))).Width
                        v_VOffset = CInt(e.Graphics.MeasureString(lineCnt.ToString, New Font(Me.Font.FontFamily, Math.Max(Math.Min(sizeFactor / 4, 10), 6))).Height / 2)
                        e.Graphics.DrawString(lineCnt.ToString,
                                    New Font(Me.Font.FontFamily, Math.Max(Math.Min(sizeFactor / 4, 10), 6)),
                                    New SolidBrush(Color.White),
                                    grfArea.Left + CInt(grfArea.Width / 2) - v_HOffset - 3,
                                    grfArea.Top + CInt((grfArea.Height / 2) - (lineCnt * sizeFactor)) - v_VOffset)
                    End If
                    lineCnt += v_step
                End While
            Catch

            End Try

            For grfCnt = 0 To 11
                If l_CheckedFunctions(grfCnt).Value Then
                    e.Graphics.DrawImage(bmp_GRF(grfCnt), grfArea.Left, grfArea.Top)
                End If
            Next
        Else
            Dim grfCnt As Integer = 0
            Dim grfRect As Rectangle
            For rowcnt = 0 To l_GridSize.Height - 1
                For colcnt = 0 To l_GridSize.Width - 1
                    grfRect = New Rectangle(CInt(colcnt * (pic_Graphics.Width / l_GridSize.Width)),
                                                    CInt(rowcnt * (pic_Graphics.Height / l_GridSize.Height)),
                                                    CInt(pic_Graphics.Width / l_GridSize.Width),
                                                    CInt(pic_Graphics.Height / l_GridSize.Height))
                    grfArea = New Rectangle(grfRect.Left + 10,
                                                    grfRect.Top + 35,
                                                    grfRect.Width - 20,
                                                    grfRect.Height - 45)
                    If (rowcnt * l_GridSize.Width) + colcnt + 1 <= l_CheckedFunctionCount Then
                        While Not l_CheckedFunctions(grfCnt).Value
                            grfCnt += 1
                        End While
                        'Graphic background
                        e.Graphics.FillRectangle(New SolidBrush(Color.FromArgb(25, 25, 25)),
                                                         grfRect.Left,
                                                         grfRect.Top,
                                                         grfRect.Width - 1,
                                                         grfRect.Height - 1)
                        'Graphic header
                        e.Graphics.FillRectangle(New Drawing2D.LinearGradientBrush(New Rectangle(grfRect.Left + 1,
                                                                                                         grfRect.Top,
                                                                                                         grfRect.Width - 2,
                                                                                                         25),
                                                                                            Color.FromArgb(64, 64, 64),
                                                                                            Color.Black,
                                                                                            Drawing2D.LinearGradientMode.Vertical),
                                                         grfRect.Left + 1,
                                                         grfRect.Top,
                                                         grfRect.Width - 2,
                                                         25)
                        'Graphic border
                        e.Graphics.DrawRectangle(New Pen(New SolidBrush(Color.Silver)),
                                                     grfRect.Left,
                                                     grfRect.Top,
                                                     grfRect.Width - 1,
                                                     grfRect.Height - 1)
                        'Graphic info
                        e.Graphics.FillRectangle(New SolidBrush(Colors(grfCnt)),
                                                     grfRect.Left + 10,
                                                     grfRect.Top + 11,
                                                     25,
                                                     3)
                        e.Graphics.DrawString(l_CheckedFunctions(grfCnt).Key,
                                                  sender.font,
                                                  New SolidBrush(Color.White),
                                                  grfRect.Left + 40,
                                                  grfRect.Top + 2)
                        'Graphic area background
                        e.Graphics.FillRectangle(New SolidBrush(Color.FromArgb(10, 255, 255, 255)), grfArea)
                        'Graphic axes
                        e.Graphics.DrawLine(New Pen(Color.DimGray, 3),
                                                grfArea.Left,
                                                grfArea.Top + CInt(grfArea.Height / 2),
                                                grfArea.Right,
                                                grfArea.Top + CInt(grfArea.Height / 2))
                        e.Graphics.DrawLine(New Pen(Color.DimGray, 3),
                                                grfArea.Left + CInt(grfArea.Width / 2),
                                                grfArea.Bottom,
                                                grfArea.Left + CInt(grfArea.Width / 2),
                                                grfArea.Top)
                        For locCnt = 1 To 5
                            e.Graphics.DrawLine(New Pen(Color.DimGray),
                                                grfArea.Right - 10,
                                                grfArea.Top + CInt(grfArea.Height / 2) - locCnt,
                                                grfArea.Right,
                                                grfArea.Top + CInt(grfArea.Height / 2))
                            e.Graphics.DrawLine(New Pen(Color.DimGray),
                                                grfArea.Right - 10,
                                                grfArea.Top + CInt(grfArea.Height / 2) + locCnt,
                                                grfArea.Right,
                                                grfArea.Top + CInt(grfArea.Height / 2))
                            e.Graphics.DrawLine(New Pen(Color.DimGray),
                                                grfArea.Left + CInt(grfArea.Width / 2) - locCnt,
                                                grfArea.Top + 10,
                                                grfArea.Left + CInt(grfArea.Width / 2),
                                                grfArea.Top)
                            e.Graphics.DrawLine(New Pen(Color.DimGray),
                                                grfArea.Left + CInt(grfArea.Width / 2) + locCnt,
                                                grfArea.Top + 10,
                                                grfArea.Left + CInt(grfArea.Width / 2),
                                                grfArea.Top)
                        Next
                        'Graphic shape
                        If bmp_GRF(grfCnt) IsNot Nothing Then
                            e.Graphics.DrawImage(bmp_GRF(grfCnt),
                                                 grfArea.Left,
                                                 grfArea.Top)
                        End If
                        grfCnt += 1
                    End If
                Next
            Next
        End If
    End Sub

    Private Sub chk_fct_CheckedChanged(sender As Object, e As EventArgs) Handles chk_sh.CheckedChanged, chk_th.CheckedChanged, chk_sech.CheckedChanged, chk_cth.CheckedChanged, chk_csch.CheckedChanged, chk_ch.CheckedChanged, chk_arsech.CheckedChanged, chk_argth.CheckedChanged, chk_argcth.CheckedChanged, chk_arcsh.CheckedChanged, chk_arcch.CheckedChanged, chk_acsch.CheckedChanged
        If sender.checked Then
            If CheckedFunctionCount() = 12 Then
                chk_all.Checked = True
            End If
        Else
            chk_all.Tag = False
            chk_all.Checked = False
        End If
        SetLimits()
        CalculateValues(False)
        pic_Graphics.Refresh()
    End Sub

    Private Sub pic_Graphics_Resize(sender As Object, e As EventArgs) Handles pic_Graphics.Resize
        tmr_Resize.Enabled = False
        tmr_Resize.Tag = pic_Graphics.Size
        tmr_Resize.Enabled = True
    End Sub

    Private Sub chk_Single_CheckedChanged(sender As Object, e As EventArgs) Handles chk_Single.CheckedChanged
        SetLimits()
        CalculateValues(False)
        pic_Graphics.Refresh()
    End Sub

    Private Sub SetLimits()
        Dim l_Limits As KeyValuePair(Of Double, Double) = GetLimits()

        If Double.IsInfinity(Math.Abs(l_Limits.Key)) Or Double.IsInfinity(Math.Abs(l_Limits.Value)) Then
            Exit Sub
        End If

        If txt_From.ForeColor = Color.Blue Or txt_To.ForeColor = Color.Blue Or (Trim(txt_From.Text) = vbNullString And Trim(txt_To.Text) = vbNullString) Then
            txt_From.Text = l_Limits.Key
            txt_To.Text = l_Limits.Value
        End If
    End Sub

    Private Sub CalculateValues(Optional useTimer As Boolean = True)
        If useTimer Then
            tmr_Calculate.Enabled = False
            tmr_Calculate.Enabled = True
            Exit Sub
        End If
        For locCnt = 0 To 11
            bmp_GRF(locCnt) = New Bitmap(1, 1)
            val_Fct(locCnt) = New List(Of KeyValuePair(Of Double, Double))
            val_Fct(locCnt).Add(New KeyValuePair(Of Double, Double)(Double.NaN, Double.NaN))
        Next
        If (txt_From.BackColor = Color.Cyan Or txt_From.BackColor = Color.Silver) And
                    (txt_To.BackColor = Color.Cyan Or txt_To.BackColor = Color.Silver) And
                    txt_Step.BackColor = Color.Silver And b_Shown Then
            Dim l_CheckedFunctions() As KeyValuePair(Of String, Boolean) = CheckedFunctions()
            Dim l_From As Double = Double.Parse(txt_From.Text, Globalization.NumberStyles.Any, Globalization.CultureInfo.CurrentCulture)
            Dim l_To As Double = Double.Parse(txt_To.Text, Globalization.NumberStyles.Any, Globalization.CultureInfo.CurrentCulture)
            Dim l_Step As Double = Double.Parse(txt_Step.Text, Globalization.NumberStyles.Any, Globalization.CultureInfo.CurrentCulture)
            Dim p_Step As Integer = 0
            While Math.Floor(l_Step * Math.Pow(10, p_Step)) / Math.Pow(10, p_Step) <> l_Step And p_Step < 4
                p_Step += 1
            End While

            Dim x As Double = l_From

            While x < l_To
                Dim val_x(11) As Double
                Dim val_y(11) As Double

                'Sinus hiperbolic
                If l_CheckedFunctions(0).Value Then
                    val_x(0) = x
                    val_y(0) = (Math.Pow(Math.E, val_x(0)) - Math.Pow(Math.E, -val_x(0))) / 2
                End If

                'Arcsinus hiperbolic
                If l_CheckedFunctions(1).Value Then
                    val_x(1) = x
                    val_y(1) = Math.Log(val_x(1) + Math.Sqrt(Math.Pow(val_x(1), 2) + 1))
                End If

                'Cosinus hiperbolic
                If l_CheckedFunctions(2).Value Then
                    val_x(2) = x
                    val_y(2) = (Math.Pow(Math.E, val_x(2)) + Math.Pow(Math.E, -val_x(2))) / 2
                End If

                'Arccosinus hiperbolic
                If l_CheckedFunctions(3).Value Then
                    If x >= 1 Then
                        val_x(3) = x
                        val_y(3) = (Math.Log(val_x(3) + Math.Sqrt(Math.Pow(val_x(3), 2) - 1)))
                    Else
                        val_x(3) = Double.NaN
                        val_y(3) = Double.NaN
                    End If
                End If

                'Tangenta hiperbolica
                If l_CheckedFunctions(4).Value Then
                    val_x(4) = x
                    val_y(4) = (Math.Exp(x) - Math.Exp(-x)) / (Math.Exp(x) + Math.Exp(-x))
                End If

                'Arctangenta hiperbolica
                If l_CheckedFunctions(5).Value Then
                    If Math.Abs(x) < 1 Then
                        val_x(5) = x
                        val_y(5) = Math.Log((1 + val_x(5)) / (1 - val_x(5))) / 2
                    Else
                        val_x(5) = Double.NaN
                        val_y(5) = Double.NaN
                    End If
                End If

                'Cotangenta hiperbolica
                If l_CheckedFunctions(6).Value Then
                    If x <> 0 Then
                        val_x(6) = x
                        val_y(6) = (Math.Exp(x) + Math.Exp(-x)) / (Math.Exp(x) - Math.Exp(-x))
                    Else
                        val_x(6) = 0
                        val_y(6) = Double.NaN
                    End If
                End If

                'Arccotangenta hipaerbolica
                If l_CheckedFunctions(7).Value Then
                    If Math.Abs(x) > 1 Then
                        val_x(7) = x
                        val_y(7) = Math.Log10((x + 1) / (x - 1)) / 2
                    ElseIf x = 0 Then
                        val_x(7) = 0
                        val_y(7) = Double.NaN
                    Else
                        val_x(7) = Double.NaN
                        val_y(7) = Double.NaN
                    End If
                End If

                'Secanta hiperbolica
                If l_CheckedFunctions(8).Value Then
                    val_x(8) = x
                    val_y(8) = 2 / (Math.Exp(x) + Math.Exp(-x))
                End If

                'Arcsecanta hiperbolica
                If l_CheckedFunctions(9).Value Then
                    If x > 0 And x <= 1 Then
                        val_x(9) = x
                        val_y(9) = Math.Log((1 + Math.Sqrt(1 - Math.Pow(x, 2))) / x)
                    Else
                        val_x(9) = Double.NaN
                        val_y(9) = Double.NaN
                    End If
                End If

                'Cosecanta hiperbolica
                If l_CheckedFunctions(10).Value Then
                    If x <> 0 Then
                        val_x(10) = x
                        val_y(10) = 2 / (Math.Exp(x) - Math.Exp(-x))
                    Else
                        val_x(10) = 0
                        val_y(10) = Double.NaN
                    End If
                End If

                'Arccosecanta hiperbolica
                If l_CheckedFunctions(11).Value Then
                    If x <> 0 Then
                        val_x(11) = x
                        'val_y(11) = Math.Log((1 + Math.Sqrt(1 + Math.Pow(x, 2))) / x)
                        val_y(11) = Math.Log((1 / x) + Math.Sqrt((1 / Math.Pow(x, 2)) + 1))
                    Else
                        val_x(11) = 0
                        val_y(11) = Double.NaN
                    End If
                End If

                For locCnt = 0 To 11
                    If l_CheckedFunctions(locCnt).Value Then
                        If Not Double.IsNaN(val_x(locCnt)) Then
                            val_Fct(locCnt).Add(New KeyValuePair(Of Double, Double)(val_x(locCnt), val_y(locCnt)))
                            If Not Double.IsNaN(val_y(locCnt)) AndAlso
                                       Math.Abs(val_y(locCnt)) <= Math.Max(Math.Abs(l_From), Math.Abs(l_To)) * 2.5 Then
                                If Double.IsNaN(val_Fct(locCnt).Item(0).Key) OrElse
                                         val_Fct(locCnt).Item(0).Key > val_y(locCnt) Then
                                    val_Fct(locCnt).Item(0) = New KeyValuePair(Of Double, Double)(val_y(locCnt), val_Fct(locCnt).Item(0).Value)
                                End If
                                If Double.IsNaN(val_Fct(locCnt).Item(0).Value) OrElse
                                         val_Fct(locCnt).Item(0).Value < val_y(locCnt) Then
                                    val_Fct(locCnt).Item(0) = New KeyValuePair(Of Double, Double)(val_Fct(locCnt).Item(0).Key, val_y(locCnt))
                                End If
                            End If
                            If Double.IsNaN(val_Fct(locCnt).Item(0).Key) OrElse val_Fct(locCnt).Item(0).Key > val_Fct(locCnt).Item(val_Fct(locCnt).Count - 1).Value Then
                                If Math.Abs(val_Fct(locCnt).Item(val_Fct(locCnt).Count - 1).Value) <= Math.Max(Math.Abs(l_From), Math.Abs(l_To)) * 2.5 Then
                                    val_Fct(locCnt).Item(0) = New KeyValuePair(Of Double, Double)(val_Fct(locCnt).Item(val_Fct(locCnt).Count - 1).Value, val_Fct(locCnt).Item(0).Value)
                                End If
                            End If
                            If Double.IsNaN(val_Fct(locCnt).Item(0).Value) OrElse val_Fct(locCnt).Item(0).Value < val_Fct(locCnt).Item(val_Fct(locCnt).Count - 1).Value Then
                                If Math.Abs(val_Fct(locCnt).Item(val_Fct(locCnt).Count - 1).Value) <= Math.Max(Math.Abs(l_From), Math.Abs(l_To)) * 2.5 Then
                                    val_Fct(locCnt).Item(0) = New KeyValuePair(Of Double, Double)(val_Fct(locCnt).Item(0).Key, val_Fct(locCnt).Item(val_Fct(locCnt).Count - 1).Value)
                                End If
                            End If
                        End If
                    End If
                Next
                If Math.Round((x + l_Step) * Math.Pow(10, p_Step)) / Math.Pow(10, p_Step) = -1 Then
                    x = -1 - Math.Pow(10, -15)
                ElseIf x = -1 - Math.Pow(10, -15) Then
                    x = -1
                ElseIf x = -1 Then
                    x = -1 + Math.Pow(10, -15)
                ElseIf x = -1 + Math.Pow(10, -15) Then
                    x = Math.Round((-1 + l_Step) * Math.Pow(10, p_Step)) / Math.Pow(10, p_Step)
                ElseIf Math.Round((x + l_Step) * Math.Pow(10, p_Step)) / Math.Pow(10, p_Step) = 0 Then
                    x = 0 - Math.Pow(10, -15)
                ElseIf x = 0 - Math.Pow(10, -15) Then
                    x = 0
                ElseIf x = 0 Then
                    x = 0 + Math.Pow(10, -15)
                ElseIf x = 0 + Math.Pow(10, -15) Then
                    x = l_Step
                ElseIf Math.Round((x + l_Step) * Math.Pow(10, p_Step)) / Math.Pow(10, p_Step) = 1 Then
                    x = 1 - Math.Pow(10, -15)
                ElseIf x = 1 - Math.Pow(10, -15) Then
                    x = 1
                ElseIf x = 1 Then
                    x = 1 + Math.Pow(10, -15)
                ElseIf x = 1 + Math.Pow(10, -15) Then
                    x = 1 + l_Step
                Else
                    x = Math.Round((x + l_Step) * Math.Pow(10, p_Step)) / Math.Pow(10, p_Step)
                End If
            End While
            CreateGraphs()
        End If
    End Sub

    Private Sub CreateGraphs()
        Dim l_CheckedFunctionCount As Integer = CheckedFunctionCount()
        Dim l_CheckedFunctions() As KeyValuePair(Of String, Boolean) = CheckedFunctions()
        If l_CheckedFunctionCount = 0 Then
            Exit Sub
        End If
        Dim l_GridSize As Size = GridSize()
        Dim bmp_Size As Size
        Dim sizeFactor As Double
        Dim sizeFactor_x As Double
        Dim sizeFactor_y As Double
        Dim l_From As Double = Double.Parse(txt_From.Text, Globalization.NumberStyles.Any, Globalization.CultureInfo.CurrentCulture)
        Dim l_To As Double = Double.Parse(txt_To.Text, Globalization.NumberStyles.Any, Globalization.CultureInfo.CurrentCulture)
        Dim l_Step As Double = Double.Parse(txt_Step.Text, Globalization.NumberStyles.Any, Globalization.CultureInfo.CurrentCulture)
        Dim v_VOffset As Integer
        Dim v_HOffset As Integer

        If chk_Single.Checked Then
            If pic_Graphics.Width / pic_Graphics.Height >= 1.5 Then
                bmp_Size = New Size(pic_Graphics.Width - 325, pic_Graphics.Height - 20)
            Else
                bmp_Size = New Size(pic_Graphics.Width - 20, pic_Graphics.Height - l_GridSize.Width * 25 - 20)
            End If

            For grfCnt = 0 To 11
                sizeFactor_x = ((bmp_Size.Width) / (l_To - l_From))
                sizeFactor_y = ((bmp_Size.Height) / (val_Fct(grfCnt)(0).Value - val_Fct(grfCnt)(0).Key))
                If Double.IsNaN(sizeFactor_y) Then
                    sizeFactor_y = sizeFactor_x
                End If

                If Double.IsNaN(sizeFactor) OrElse sizeFactor = 0 Then
                    sizeFactor = Math.Min(sizeFactor_x, sizeFactor_y)
                Else
                    sizeFactor = Math.Min(sizeFactor, Math.Min(sizeFactor_x, sizeFactor_y))
                End If
            Next
        Else
            bmp_Size = New Size(CInt(pic_Graphics.Width / l_GridSize.Width) - 20, CInt(pic_Graphics.Height / l_GridSize.Height) - 45)
        End If

        For bmpCnt = 0 To 11
            If l_CheckedFunctions(bmpCnt).Value Then
                If Not chk_Single.Checked Then
                    sizeFactor_x = ((bmp_Size.Width) / (l_To - l_From))
                    sizeFactor_y = ((bmp_Size.Height) / (val_Fct(bmpCnt)(0).Value - val_Fct(bmpCnt)(0).Key))
                    If Double.IsNaN(sizeFactor_y) Then
                        sizeFactor_y = sizeFactor_x
                    End If
                    sizeFactor = Math.Min(sizeFactor_x, sizeFactor_y)
                    If Double.IsNaN(sizeFactor) Then
                        sizeFactor = sizeFactor_x
                    End If
                End If
                bmp_GRF(bmpCnt) = New Bitmap(bmp_Size.Width, bmp_Size.Height)
                Dim g_BMP As Graphics = Graphics.FromImage(bmp_GRF(bmpCnt))
                g_BMP.FillRectangle(New SolidBrush(Color.Transparent), 0, 0, bmp_GRF(bmpCnt).Width, bmp_GRF(bmpCnt).Height)

                Dim v_step As Double
                Dim lineCnt As Double

                If Not chk_Single.Checked Then
                    v_step = 0.0001
                    While (v_step * sizeFactor_x) < 25
                        If Replace(Replace(Replace(v_step.ToString, "0", ""), ".", ""), ",", "") = "1" Then
                            v_step *= 5
                        Else
                            v_step *= 2
                        End If
                    End While
                    lineCnt = v_step
                    While (lineCnt * sizeFactor_x) < (bmp_Size.Width / 2)
                        lineCnt = Math.Round(lineCnt, 2)
                        g_BMP.DrawLine(New Pen(Color.FromArgb(10, 255, 255, 255), 1),
                                   CInt((bmp_Size.Width / 2) - (lineCnt * sizeFactor_x)),
                                   0,
                                   CInt((bmp_Size.Width / 2) - (lineCnt * sizeFactor_x)),
                                   bmp_Size.Height)
                        g_BMP.DrawLine(New Pen(Color.FromArgb(10, 255, 255, 255), 1),
                                   CInt((bmp_Size.Width / 2) + (lineCnt * sizeFactor_x)),
                                   0,
                                   CInt((bmp_Size.Width / 2) + (lineCnt * sizeFactor_x)),
                                   bmp_Size.Height)
                        If CInt((bmp_Size.Width / 2) - (lineCnt * sizeFactor_x)) - 6 > 0 Then
                            v_HOffset = CInt(g_BMP.MeasureString((-lineCnt).ToString,
                                                            New Font(Me.Font.FontFamily, Math.Max(Math.Min(sizeFactor / 4, 10), 6))).Width / 2)
                            v_VOffset = p_axes(bmpCnt, 0) *
                                    (CInt(g_BMP.MeasureString((-lineCnt).ToString,
                                                               New Font(Me.Font.FontFamily, Math.Max(Math.Min(sizeFactor / 4, 10), 6))).Height) / 2 + 3) +
                                    (CInt(g_BMP.MeasureString((-lineCnt).ToString,
                                                               New Font(Me.Font.FontFamily, Math.Max(Math.Min(sizeFactor / 4, 10), 6))).Height) / 2) - 2
                            g_BMP.DrawString((-lineCnt).ToString,
                                     New Font(Me.Font.FontFamily, Math.Max(Math.Min(sizeFactor / 4, 10), 6)),
                                     New SolidBrush(Color.White),
                                     CInt((bmp_Size.Width / 2) - (lineCnt * sizeFactor_x)) - v_HOffset,
                                     CInt((bmp_Size.Height) / 2 - v_VOffset))
                            v_HOffset = CInt(g_BMP.MeasureString(lineCnt.ToString,
                                                            New Font(Me.Font.FontFamily, Math.Max(Math.Min(sizeFactor / 4, 10), 6))).Width / 2)
                            v_VOffset = p_axes(bmpCnt, 1) *
                                    (CInt(g_BMP.MeasureString(lineCnt.ToString,
                                                               New Font(Me.Font.FontFamily, Math.Max(Math.Min(sizeFactor / 4, 10), 6))).Height) / 2 + 3) +
                                    (CInt(g_BMP.MeasureString(lineCnt.ToString,
                                                               New Font(Me.Font.FontFamily, Math.Max(Math.Min(sizeFactor / 4, 10), 6))).Height) / 2) - 2
                            g_BMP.DrawString(lineCnt.ToString,
                                     New Font(Me.Font.FontFamily, Math.Max(Math.Min(sizeFactor / 4, 10), 6)),
                                     New SolidBrush(Color.White),
                                     CInt((bmp_Size.Width / 2) + (lineCnt * sizeFactor_x)) - v_HOffset,
                                     CInt(bmp_Size.Height) / 2 - v_VOffset)
                        End If
                        lineCnt += v_step
                    End While
                End If

                If Not chk_Single.Checked Then
                    v_step = 0.001
                    While (v_step * sizeFactor) < 13
                        If Replace(Replace(Replace(v_step.ToString, "0", ""), ".", ""), ",", "") = "1" Then
                            v_step *= 5
                        Else
                            v_step *= 2
                        End If
                    End While
                    lineCnt = v_step
                    While (lineCnt * sizeFactor) < (bmp_Size.Height / 2)
                        g_BMP.DrawLine(New Pen(Color.FromArgb(10, 255, 255, 255), 1),
                               0,
                               CInt((bmp_Size.Height / 2) + (lineCnt * sizeFactor)),
                               bmp_Size.Width,
                               CInt((bmp_Size.Height / 2) + (lineCnt * sizeFactor)))
                        g_BMP.DrawLine(New Pen(Color.FromArgb(10, 255, 255, 255), 1),
                                   0,
                                   CInt((bmp_Size.Height / 2) - (lineCnt * sizeFactor)),
                                   bmp_Size.Width,
                                   CInt((bmp_Size.Height / 2) - (lineCnt * sizeFactor)))
                        If CInt((bmp_Size.Height / 2) - (lineCnt * sizeFactor)) - 6 > 0 Then
                            v_HOffset = -p_axes(bmpCnt, 2) *
                                    (CInt(g_BMP.MeasureString((-lineCnt).ToString,
                                                               New Font(Me.Font.FontFamily, Math.Max(Math.Min(sizeFactor / 4, 10), 6))).Width) / 2 + 3) +
                                    (CInt(g_BMP.MeasureString((-lineCnt).ToString,
                                                               New Font(Me.Font.FontFamily, Math.Max(Math.Min(sizeFactor / 4, 10), 6))).Width) / 2)
                            v_VOffset = CInt(g_BMP.MeasureString((-lineCnt).ToString,
                                                            New Font(Me.Font.FontFamily, Math.Max(Math.Min(sizeFactor / 4, 10), 6))).Height / 2)
                            g_BMP.DrawString((-lineCnt).ToString,
                                     New Font(Me.Font.FontFamily, Math.Max(Math.Min(sizeFactor / 4, 10), 6)),
                                     New SolidBrush(Color.White),
                                     CInt(bmp_Size.Width / 2) - v_HOffset,
                                     CInt((bmp_Size.Height / 2) + (lineCnt * sizeFactor)) - v_VOffset)
                            v_HOffset = -p_axes(bmpCnt, 3) *
                                    (CInt(g_BMP.MeasureString(lineCnt.ToString,
                                                               New Font(Me.Font.FontFamily, Math.Max(Math.Min(sizeFactor / 4, 10), 6))).Width) / 2 + 3) +
                                    (CInt(g_BMP.MeasureString((-lineCnt).ToString,
                                                               New Font(Me.Font.FontFamily, Math.Max(Math.Min(sizeFactor / 4, 10), 6))).Width) / 2)
                            v_VOffset = CInt(g_BMP.MeasureString(lineCnt.ToString,
                                                            New Font(Me.Font.FontFamily, Math.Max(Math.Min(sizeFactor / 4, 10), 6))).Height / 2)
                            g_BMP.DrawString(lineCnt.ToString,
                                     New Font(Me.Font.FontFamily, Math.Max(Math.Min(sizeFactor / 4, 10), 6)),
                                     New SolidBrush(Color.White),
                                     CInt(bmp_Size.Width / 2) - v_HOffset,
                                     CInt((bmp_Size.Height / 2) - (lineCnt * sizeFactor)) - v_VOffset)
                        End If
                        lineCnt += v_step
                    End While
                End If

                For valCnt = 2 To val_Fct(bmpCnt).Count - 1
                    'Graphic grid
                    If (bmp_Size.Height - ((val_Fct(bmpCnt)(valCnt - 1).Value * sizeFactor) + (bmp_Size.Height / 2)) > 0 Or
                                bmp_Size.Height - ((val_Fct(bmpCnt)(valCnt).Value * sizeFactor) + (bmp_Size.Height / 2)) > 0) And
                                (bmp_Size.Height - ((val_Fct(bmpCnt)(valCnt - 1).Value * sizeFactor) + (bmp_Size.Height / 2)) < bmp_Size.Height Or
                                bmp_Size.Height - ((val_Fct(bmpCnt)(valCnt).Value * sizeFactor) + (bmp_Size.Height / 2)) < bmp_Size.Height) And
                                Not Double.IsNaN(val_Fct(bmpCnt)(valCnt - 1).Value) And
                                Not Double.IsNaN(val_Fct(bmpCnt)(valCnt).Value) Then
                        Try
                            g_BMP.DrawLine(New Pen(Colors(bmpCnt), Math.Min(Math.Max(Math.Min(bmp_Size.Width, bmp_Size.Height) / 100, 1), 3)),
                                CInt((val_Fct(bmpCnt)(valCnt - 1).Key * sizeFactor_x) + (bmp_Size.Width / 2)),
                                bmp_Size.Height - (CInt((val_Fct(bmpCnt)(valCnt - 1).Value * sizeFactor) + (bmp_Size.Height / 2))),
                                CInt((val_Fct(bmpCnt)(valCnt).Key * sizeFactor_x) + (bmp_Size.Width / 2)),
                                bmp_Size.Height - (CInt((val_Fct(bmpCnt)(valCnt).Value * sizeFactor) + (bmp_Size.Height / 2))))
                        Catch

                        End Try
                    End If
                Next

                g_BMP.Flush()
            End If
        Next
    End Sub

    Private Sub txt_From_TextChanged(sender As Object, e As EventArgs) Handles txt_From.TextChanged
        SetValColors(sender)

        If Trim(txt_From.Text) = vbNullString Then
            txt_From.Text = GetLimits.Key.ToString
            txt_From.SelectionStart = 0
            txt_From.SelectionLength = txt_From.Text.Length
        End If

        CalculateValues()
    End Sub

    Private Sub txt_To_TextChanged(sender As Object, e As EventArgs) Handles txt_To.TextChanged
        SetValColors(sender)

        If Trim(txt_To.Text) = vbNullString Then
            txt_To.Text = GetLimits.Value.ToString
            txt_To.SelectionStart = 0
            txt_To.SelectionLength = txt_To.Text.Length
        End If

        CalculateValues()
    End Sub

    Private Sub txt_Step_TextChanged(sender As Object, e As EventArgs) Handles txt_Step.TextChanged
        SetValColors(sender)

        If Trim(txt_Step.Text) = vbNullString Then
            txt_Step.Text = (0.01).ToString
            txt_Step.SelectionStart = 0
            txt_Step.SelectionLength = txt_Step.Text.Length
        End If

        CalculateValues()
    End Sub

    Private Sub SetValColors(ByRef txt_Obj As TextBox)
        Dim l_limits As KeyValuePair(Of Double, Double) = GetLimits()
        Dim l_IsNumber() As Boolean = {True, True, True}
        Dim l_Compareable As Boolean = True
        Dim l_From As Decimal
        Dim l_To As Decimal
        Dim l_Step As Decimal

        l_IsNumber(1) = Decimal.TryParse(txt_From.Text, Globalization.NumberStyles.Any, Globalization.CultureInfo.CurrentCulture, l_From)
        l_IsNumber(2) = Decimal.TryParse(txt_To.Text, Globalization.NumberStyles.Any, Globalization.CultureInfo.CurrentCulture, l_To)
        l_IsNumber(0) = Decimal.TryParse(txt_Step.Text, Globalization.NumberStyles.Any, Globalization.CultureInfo.CurrentCulture, l_Step)

        If l_IsNumber(1) Then
            txt_From.BackColor = Color.Silver
            If l_From = l_limits.Key And l_To = l_limits.Value Then
                txt_From.ForeColor = Color.Blue
                txt_To.ForeColor = Color.Blue
            Else
                txt_From.ForeColor = Color.Black
                txt_To.ForeColor = Color.Black
            End If
        Else
            txt_From.BackColor = Color.Red
        End If

        If l_IsNumber(2) Then
            txt_To.BackColor = Color.Silver
            If l_From = l_limits.Key And l_To = l_limits.Value Then
                txt_From.ForeColor = Color.Blue
                txt_To.ForeColor = Color.Blue
            Else
                txt_From.ForeColor = Color.Black
                txt_To.ForeColor = Color.Black
            End If
        Else
            txt_To.BackColor = Color.Red
        End If

        txt_Step.BackColor = Color.Silver
        If l_IsNumber(1) And l_IsNumber(2) Then
            If l_From < l_To Then
                If l_IsNumber(0) Then
                    If l_Step > l_To - l_From Then
                        txt_Step.BackColor = Color.Yellow
                    End If
                End If
            Else
                txt_Obj.BackColor = Color.Yellow
                txt_Step.BackColor = Color.Yellow
            End If
        End If

        If Not (l_IsNumber(0) AndAlso l_Step >= 0.0001) Then
            txt_Step.BackColor = Color.Red
        End If
    End Sub

    Private Sub txt_Input_GotFocus(sender As TextBox, e As EventArgs) Handles txt_From.GotFocus, txt_To.GotFocus, txt_Step.GotFocus
        sender.SelectionStart = 0
        sender.SelectionLength = sender.Text.Length
    End Sub

    Private Sub txt_Input_MouseUp(sender As Object, e As MouseEventArgs) Handles txt_From.MouseUp, txt_To.MouseUp, txt_Step.MouseUp
        sender.SelectionStart = 0
        sender.SelectionLength = sender.Text.Length
    End Sub

    Private Sub chk_all_CheckedChanged(sender As Object, e As EventArgs) Handles chk_all.CheckedChanged
        If chk_all.Tag Is Nothing Or chk_all.Checked Then
            chk_sh.Checked = chk_all.Checked
            chk_arcsh.Checked = chk_all.Checked
            chk_ch.Checked = chk_all.Checked
            chk_arcch.Checked = chk_all.Checked
            chk_th.Checked = chk_all.Checked
            chk_argth.Checked = chk_all.Checked
            chk_cth.Checked = chk_all.Checked
            chk_argcth.Checked = chk_all.Checked
            chk_sech.Checked = chk_all.Checked
            chk_arsech.Checked = chk_all.Checked
            chk_csch.Checked = chk_all.Checked
            chk_acsch.Checked = chk_all.Checked
        End If
        chk_all.Tag = Nothing
    End Sub

    Private Sub Panel_Paint(sender As Object, e As PaintEventArgs) Handles Panel1.Paint, Panel2.Paint, Panel4.Paint, Panel5.Paint
        Dim tmpRect As Rectangle = sender.ClientRectangle()
        e.Graphics.FillRectangle(New Drawing2D.LinearGradientBrush(tmpRect,
                                                                   Color.FromArgb(64, 64, 64),
                                                                   Color.Black,
                                                                   Drawing2D.LinearGradientMode.Vertical),
                                 0, 0,
                                 sender.Width, sender.Height)
    End Sub

    Private Sub tmr_Resize_Tick(sender As Object, e As EventArgs) Handles tmr_Resize.Tick
        If pic_Graphics.Size.Width = tmr_Resize.Tag.Width And pic_Graphics.Height = tmr_Resize.Tag.Height And b_Shown Then
            tmr_Resize.Enabled = False
            SetLimits()
            CalculateValues(False)
            pic_Graphics.Refresh()
        End If
    End Sub

    Private Sub Panel2_Resize(sender As Object, e As EventArgs) Handles Panel2.Resize
        sender.refresh
    End Sub

    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        b_Shown = True
    End Sub

    Private Sub tmr_Calculate_Tick(sender As Object, e As EventArgs) Handles tmr_Calculate.Tick
        tmr_Calculate.Enabled = False
        CalculateValues(False)
        pic_Graphics.Refresh()
    End Sub
End Class
