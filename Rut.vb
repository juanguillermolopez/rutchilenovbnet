
'Uso de la clase:
'ya sea en formulario o consola
'instanciar como sigue:

'Dim r = New Rut(txtrut.Text)
'-----------------------------
'If r.validarRutCompleto Then
'.....MsgBox("Rut válido")
'Else
'.....MsgBox("Rut NO válido")
'End If

' Console.WriteLine(r.obtenerNumero)
' Console.WriteLine(r.obtenerDigito)
' Console.WriteLine(r.obtenerRutCompleto)

'Diseñado por: Guillermo.
'Año: 2023.
'Estudiente de segundo año de Ing. Inf. 
'Intituto IPP.
'Chile.
'
'Solo escribe para saber si te sirvió a:
'juan.guillermo.lopez.garcia@gmail.com


Public Class Rut

    Private _rut As String
    Private _segmento As String
    Private _digito As String
    Private _rutcompleto As String

    Public Property obtenerRutCompleto As String
        'obtiene el rut completo, número + '-' + dígito
        Get
            Return _rutcompleto
        End Get
        Set(ByVal value As String)
            _rutcompleto = value
        End Set
    End Property

    Public Property obtenerDigito() As String
        'solo obtiene el dígito verificador
        'como el resultado final del cálculo
        Get
            Return _digito
        End Get
        Set(ByVal value As String)
            _digito = value
        End Set
    End Property

    Public Property obtenerNumero() As String
        'obtiene al cadena anterior al guión '-'
        Get
            Return _segmento
        End Get
        Set(ByVal value As String)
            _segmento = value
        End Set
    End Property

    Public Sub New()
        ' constructor por defecto
    End Sub

    'rut de ejemplos para pruebas
    '12487521-8
    '10214564-K
    Public Sub New(r As String)
        'constructor con parámetros
        'entrega el valor a la varible local
        _rut = r.Trim.Length
    End Sub

    Public Function tieneLongitudCorrecta() As Boolean
        'funciona con rut mayores y menores que 10.000.000 (10 millones)
        Return _rut.Length >= 6
    End Function

    Public Function tienePrimerDigitoNoCero() As Boolean
        'comprueba que no empiece con 0
        Try
            Return _rut(0) <> "0"
        Catch ex As Exception

        End Try
    End Function

    Public Function tieneGuionEnPosicionCorrecta() As Boolean
        'compueba que contenga el guión donde debe ser
        Try
            Return _rut(_rut.Length - 2) = "-"
        Catch ex As Exception

        End Try
    End Function

    Public Function validarDigitoVerificador() As Boolean
        '
        Try
            '' variables locales a la función
            Dim factor As Integer = 2
            Dim suma As Integer


            'separa el rut en sus partes principales
            Dim segmento = _rut.Substring(0, _rut.Length - 2)
            Dim digito = _rut(_rut.Length - 1)

            'se multiplica cada dígito por el factor en forma inversa
            For i As Integer = segmento.Length - 1 To 0 Step -1
                Dim prod = Integer.Parse(segmento(i)) * factor
                suma += prod
                factor += 1
                If factor > 7 Then
                    factor = 2
                End If
            Next

            ''
            Dim cociente = suma \ 11 'división entera
            Dim resto = suma - (11 * cociente) 'resto entero

            ''
            Dim resultado As String 'variable para comparar los dígitos resultantes
            Dim dv = 11 - resto 'entero

            ''
            If dv = 10 Then
                resultado = "k"
            ElseIf dv = 11 Then
                resultado = "0"
            Else
                resultado = dv
            End If

            ''
            If resultado = digito Then
                'asignación de valores para luego obtenerlos mediante los métodos.
                _segmento = segmento
                _digito = resultado
                _rutcompleto = _segmento & "-" & _digito
                Return True 'si todo está correcto se devuelve true
            Else
                Return False 'rut con errores
            End If
        Catch ex As Exception

        End Try
    End Function

    Public Function validarRutCompleto() As Boolean
        'se comprueba que cada paso sea verdadero
        Return tieneLongitudCorrecta() And
                tienePrimerDigitoNoCero() And
                tieneGuionEnPosicionCorrecta() And
                validarDigitoVerificador()
    End Function

End Class
