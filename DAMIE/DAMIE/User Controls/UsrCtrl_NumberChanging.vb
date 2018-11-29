Imports System.Windows.Forms

Namespace AME_UserControl

    ''' <summary>
    ''' 用户控件，用来增加或减少指定的日期值。
    ''' </summary>
    ''' <remarks></remarks>
    Public Class UsrCtrl_NumberChanging

#Region "  ---  Types"

        Public Enum YearMonthDay
            Days = 0
            Months = 1
            Years = 2
        End Enum

#End Region

#Region "  ---  Properties"
        Private _unit As YearMonthDay
        Public Property unit As YearMonthDay
            Get
                Return _unit
            End Get
            Set(value As YearMonthDay)
                cbUnit.SelectedText = value.ToString
                _unit = value
            End Set
        End Property

        'Private P_CanGetFocus As Boolean
        'Public ReadOnly Property CanGetFocus As Boolean
        '    Get
        '        Return Me.P_CanGetFocus
        '    End Get
        'End Property
#End Region

#Region "  ---  Fields"
        Public Event ValueAdded()
        Public Event ValueMinused()
        Public Shadows Event TextChanged()
#End Region

        Private _Value_TimeSpan As Single
        ''' <summary>
        ''' 日期文本框上显示的用来进行日期值增减的数量
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Value_TimeSpan As Single
            Get
                Return _Value_TimeSpan
            End Get
        End Property

        Private Sub NumberChanging_Load(sender As Object, e As EventArgs) Handles Me.Load
            With cbUnit
                Dim names As String() = [Enum].GetNames(GetType(YearMonthDay))
                .Items.Clear()
                .Items.AddRange(names)
                .SelectedIndex = 0
            End With
            TextBoxNumber.Text = 1
            Me._Value_TimeSpan = 1
        End Sub

        Private Sub btnNext_Click(sender As Object, e As EventArgs) Handles btnNext.Click
            RaiseEvent ValueAdded()
        End Sub
        Private Sub btnPrevious_Click(sender As Object, e As EventArgs) Handles btnPrevious.Click
            RaiseEvent ValueMinused()
        End Sub

        Private Sub btnUnit_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbUnit.SelectedIndexChanged
            Me._unit = [Enum].Parse(GetType(YearMonthDay), cbUnit.SelectedItem, True)
        End Sub

        Private Sub TextBoxNumber_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBoxNumber.KeyUp
            Dim v As Single
            Try
                v = Single.Parse(TextBoxNumber.Text)
                Me._Value_TimeSpan = TextBoxNumber.Text
                RaiseEvent TextChanged()
            Catch ex As Exception
                TextBoxNumber.Text = ""
                Me._Value_TimeSpan = 0
            End Try
        End Sub
        '
    End Class

End Namespace
