Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub 'ETIQUETA DE REFERENCIA DE LAS TEXTBOX'

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click 'BOTÓN LIMPIAR'

        'LIMPIA EL FORMULARIO, EN ESTE CASO LAS DOS TEXTBOX.'
        TextBox1.Text = ""
        TextBox2.Text = ""

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click 'BOTÓN DEL CÁLCULO'

        If IsNumeric(TextBox1.Text) And IsNumeric(TextBox2.Text) Then
            Dim NumeroUno As Double                            'CREAMOS LA VARIABLE PARA ALBERGAR LA PRIMERA CAJA DEL TEXTBOX'
            Dim NumeroDos As Double                            'CREAMOS LA VARIABLE PARA ALBERGAR LA SEGUNDA CAJA DEL TEXTBOX'
            Dim Calc As Double                                 'CREAMOS LA VARIABLE DEL RESULTADO DEL CÁLCULO'


            NumeroUno = TextBox1.Text                          'LE DAMOS AL VALOR A LA VARIABLE DEL TEXTBOX'
            NumeroDos = TextBox2.Text                          'LE DAMOS AL VALOR A LA VARIABLE DEL TEXTBOX'
            Calc = (NumeroUno + NumeroDos) / 2                 'CÁLCULO DE LA MÉDIA ARITMÈTICA'

            MsgBox("El resultado del cáuclo es: " & Calc)      'MUESTRA EL RESULTADO POR PANTALLA EN UNA VENTANA'

        Else
            MsgBox("Introduce únicamente valores numéricos para su cálculo.", MessageBoxIcon.Error)    'MUESTRA UNA VENTANA DE ERROR.'

        End If

    End Sub
End Class
