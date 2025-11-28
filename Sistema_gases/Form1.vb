Public Class Form1
    ' ============================================
    ' DECLARACIÓN DE CONTROLES
    ' ============================================
    Private panelTitulo As Panel
    Private lblTitulo As Label
    Private lblGas As Label
    Private cboGas As ComboBox

    ' Controles del lado derecho
    Private lblTemperatura As Label
    Private txtTemperatura As TextBox
    Private cboUnidadTemperatura As ComboBox
    Private lblPresion As Label
    Private txtPresion As TextBox
    Private cboUnidadPresion As ComboBox

    ' Controles para constantes (cuando se selecciona "Otro")
    Private lblConstanteA As Label
    Private txtConstanteA As TextBox
    Private lblConstanteB As Label
    Private txtConstanteB As TextBox
    Private lblMasaMolar As Label
    Private txtMasaMolar As TextBox

    ' Botón de cálculo
    Private btnCalcular As Button

    ' Botones de conversión (se muestran después del cálculo)
    Private panelConversiones As Panel
    Private lblTituloConversiones As Label
    Private btnConvertirCm3g As Button
    Private btnConvertirFt3lb As Button
    Private btnConvertirIn3oz As Button

    ' Variables para almacenar las constantes del gas seleccionado
    Private constanteA As Double
    Private constanteB As Double
    Private masaMolar As Double

    ' Variable para almacenar el último resultado
    Private ultimoVolumenMolar As Double

    ' Constante universal de los gases (en kPa·m³/(mol·K))
    Private Const R As Double = 0.008314

    ' ============================================
    ' EVENTO: Cuando el formulario se carga
    ' ============================================
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ConfigurarFormulario()
        CrearPanelTitulo()
        CrearControlesEntrada()
        CrearCamposPersonalizados()
        CrearBotonCalcular()
        CrearPanelConversiones()

        ' Configurar evento cuando cambia la selección del gas
        AddHandler cboGas.SelectedIndexChanged, AddressOf cboGas_SelectedIndexChanged

        ' Inicializar constantes con el primer gas (Agua)
        ActualizarConstantesGas()
    End Sub

    ' ============================================
    ' MÉTODO: Configurar propiedades del formulario
    ' ============================================
    Private Sub ConfigurarFormulario()
        Me.Text = "Sistema de Cálculo para Gases"
        Me.Size = New Size(900, 750)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.BackColor = Color.WhiteSmoke
        Me.MaximizeBox = True
        Me.MinimumSize = New Size(800, 700)
    End Sub

    ' ============================================
    ' MÉTODO: Crear el panel superior con título CENTRADO
    ' ============================================
    Private Sub CrearPanelTitulo()
        panelTitulo = New Panel()
        panelTitulo.Dock = DockStyle.Top
        panelTitulo.Height = 80
        panelTitulo.BackColor = Color.FromArgb(41, 128, 185)
        Me.Controls.Add(panelTitulo)

        lblTitulo = New Label()
        lblTitulo.Text = "Sistema de cálculo para gases"
        lblTitulo.Font = New Font("Segoe UI", 18, FontStyle.Bold)
        lblTitulo.ForeColor = Color.White
        lblTitulo.Dock = DockStyle.Fill
        lblTitulo.TextAlign = ContentAlignment.MiddleCenter

        panelTitulo.Controls.Add(lblTitulo)
    End Sub

    ' ============================================
    ' MÉTODO: Crear controles de entrada
    ' ============================================
    Private Sub CrearControlesEntrada()
        lblGas = New Label()
        lblGas.Text = "Seleccionar Gas:"
        lblGas.Font = New Font("Segoe UI", 12, FontStyle.Bold)
        lblGas.Location = New Point(50, 120)
        lblGas.AutoSize = True
        Me.Controls.Add(lblGas)

        cboGas = New ComboBox()
        cboGas.Font = New Font("Segoe UI", 11)
        cboGas.Location = New Point(50, 150)
        cboGas.Size = New Size(200, 30)
        cboGas.DropDownStyle = ComboBoxStyle.DropDownList

        cboGas.Items.Add("Agua")
        cboGas.Items.Add("Amoniaco")
        cboGas.Items.Add("CO2")
        cboGas.Items.Add("CH4")
        cboGas.Items.Add("H2")
        cboGas.Items.Add("Otro")

        cboGas.SelectedIndex = 0

        Me.Controls.Add(cboGas)

        lblTemperatura = New Label()
        lblTemperatura.Text = "T (Temperatura):"
        lblTemperatura.Font = New Font("Segoe UI", 12, FontStyle.Bold)
        lblTemperatura.Location = New Point(450, 120)
        lblTemperatura.AutoSize = True
        Me.Controls.Add(lblTemperatura)

        txtTemperatura = New TextBox()
        txtTemperatura.Font = New Font("Segoe UI", 11)
        txtTemperatura.Location = New Point(450, 150)
        txtTemperatura.Size = New Size(150, 30)
        txtTemperatura.Text = "0"
        AddHandler txtTemperatura.KeyPress, AddressOf ValidarSoloNumeros
        Me.Controls.Add(txtTemperatura)

        cboUnidadTemperatura = New ComboBox()
        cboUnidadTemperatura.Font = New Font("Segoe UI", 11)
        cboUnidadTemperatura.Location = New Point(610, 150)
        cboUnidadTemperatura.Size = New Size(120, 30)
        cboUnidadTemperatura.DropDownStyle = ComboBoxStyle.DropDownList

        cboUnidadTemperatura.Items.Add("Kelvin")
        cboUnidadTemperatura.Items.Add("Celsius")
        cboUnidadTemperatura.Items.Add("Fahrenheit")
        cboUnidadTemperatura.SelectedIndex = 0

        Me.Controls.Add(cboUnidadTemperatura)

        lblPresion = New Label()
        lblPresion.Text = "P (Presión):"
        lblPresion.Font = New Font("Segoe UI", 12, FontStyle.Bold)
        lblPresion.Location = New Point(450, 200)
        lblPresion.AutoSize = True
        Me.Controls.Add(lblPresion)

        txtPresion = New TextBox()
        txtPresion.Font = New Font("Segoe UI", 11)
        txtPresion.Location = New Point(450, 230)
        txtPresion.Size = New Size(150, 30)
        txtPresion.Text = "0"
        AddHandler txtPresion.KeyPress, AddressOf ValidarSoloNumeros
        Me.Controls.Add(txtPresion)

        cboUnidadPresion = New ComboBox()
        cboUnidadPresion.Font = New Font("Segoe UI", 11)
        cboUnidadPresion.Location = New Point(610, 230)
        cboUnidadPresion.Size = New Size(120, 30)
        cboUnidadPresion.DropDownStyle = ComboBoxStyle.DropDownList

        cboUnidadPresion.Items.Add("KPA")
        cboUnidadPresion.Items.Add("PA")
        cboUnidadPresion.Items.Add("ATM")
        cboUnidadPresion.Items.Add("BAR")
        cboUnidadPresion.Items.Add("PSI")
        cboUnidadPresion.SelectedIndex = 0

        Me.Controls.Add(cboUnidadPresion)
    End Sub

    ' ============================================
    ' MÉTODO: Crear botón de cálculo
    ' ============================================
    Private Sub CrearBotonCalcular()
        btnCalcular = New Button()
        btnCalcular.Text = "CALCULAR"
        btnCalcular.Font = New Font("Segoe UI", 12, FontStyle.Bold)
        btnCalcular.Size = New Size(200, 45)
        btnCalcular.Location = New Point(450, 290)
        btnCalcular.BackColor = Color.FromArgb(46, 204, 113)
        btnCalcular.ForeColor = Color.White
        btnCalcular.FlatStyle = FlatStyle.Flat
        btnCalcular.FlatAppearance.BorderSize = 0
        btnCalcular.Cursor = Cursors.Hand

        AddHandler btnCalcular.Click, AddressOf btnCalcular_Click

        Me.Controls.Add(btnCalcular)
    End Sub

    ' ============================================
    ' MÉTODO: Crear panel de conversiones
    ' ============================================
    Private Sub CrearPanelConversiones()
        panelConversiones = New Panel()
        panelConversiones.Location = New Point(50, 400)
        panelConversiones.Size = New Size(800, 250)
        panelConversiones.BackColor = Color.FromArgb(236, 240, 241)
        panelConversiones.BorderStyle = BorderStyle.FixedSingle
        panelConversiones.Visible = False
        Me.Controls.Add(panelConversiones)

        lblTituloConversiones = New Label()
        lblTituloConversiones.Text = "Conversiones de Unidades"
        lblTituloConversiones.Font = New Font("Segoe UI", 14, FontStyle.Bold)
        lblTituloConversiones.ForeColor = Color.FromArgb(52, 73, 94)
        lblTituloConversiones.Location = New Point(20, 20)
        lblTituloConversiones.AutoSize = True
        panelConversiones.Controls.Add(lblTituloConversiones)

        ' Botón cm³/g
        btnConvertirCm3g = New Button()
        btnConvertirCm3g.Text = "Convertir a cm³/g" & vbCrLf & "(Volumen específico)"
        btnConvertirCm3g.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        btnConvertirCm3g.Size = New Size(220, 70)
        btnConvertirCm3g.Location = New Point(30, 80)
        btnConvertirCm3g.BackColor = Color.FromArgb(52, 152, 219)
        btnConvertirCm3g.ForeColor = Color.White
        btnConvertirCm3g.FlatStyle = FlatStyle.Flat
        btnConvertirCm3g.FlatAppearance.BorderSize = 0
        btnConvertirCm3g.Cursor = Cursors.Hand
        AddHandler btnConvertirCm3g.Click, AddressOf btnConvertirCm3g_Click
        panelConversiones.Controls.Add(btnConvertirCm3g)

        ' Botón ft³/lb
        btnConvertirFt3lb = New Button()
        btnConvertirFt3lb.Text = "Convertir a ft³/lb" & vbCrLf & "(Pies³ por libra)"
        btnConvertirFt3lb.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        btnConvertirFt3lb.Size = New Size(220, 70)
        btnConvertirFt3lb.Location = New Point(290, 80)
        btnConvertirFt3lb.BackColor = Color.FromArgb(155, 89, 182)
        btnConvertirFt3lb.ForeColor = Color.White
        btnConvertirFt3lb.FlatStyle = FlatStyle.Flat
        btnConvertirFt3lb.FlatAppearance.BorderSize = 0
        btnConvertirFt3lb.Cursor = Cursors.Hand
        AddHandler btnConvertirFt3lb.Click, AddressOf btnConvertirFt3lb_Click
        panelConversiones.Controls.Add(btnConvertirFt3lb)

        ' Botón in³/oz
        btnConvertirIn3oz = New Button()
        btnConvertirIn3oz.Text = "Convertir a in³/oz" & vbCrLf & "(Pulgadas³ por onza)"
        btnConvertirIn3oz.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        btnConvertirIn3oz.Size = New Size(220, 70)
        btnConvertirIn3oz.Location = New Point(550, 80)
        btnConvertirIn3oz.BackColor = Color.FromArgb(230, 126, 34)
        btnConvertirIn3oz.ForeColor = Color.White
        btnConvertirIn3oz.FlatStyle = FlatStyle.Flat
        btnConvertirIn3oz.FlatAppearance.BorderSize = 0
        btnConvertirIn3oz.Cursor = Cursors.Hand
        AddHandler btnConvertirIn3oz.Click, AddressOf btnConvertirIn3oz_Click
        panelConversiones.Controls.Add(btnConvertirIn3oz)
    End Sub

    ' ============================================
    ' EVENTO: Convertir a cm³/g
    ' ============================================
    Private Sub btnConvertirCm3g_Click(sender As Object, e As EventArgs)
        If ultimoVolumenMolar = 0 Then
            MessageBox.Show("Primero debe calcular el volumen molar.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Conversión: m³/mol → cm³/g
        Dim volumenEspecifico As Double = (ultimoVolumenMolar * 1000000) / (masaMolar * 1000)

        Dim mensaje As String = "=== CONVERSIÓN A cm³/g ===" & vbCrLf & vbCrLf
        mensaje = mensaje & "Volumen Molar: " & ultimoVolumenMolar.ToString("F8") & " m³/mol" & vbCrLf
        mensaje = mensaje & "Masa Molar: " & masaMolar.ToString("F6") & " kg/mol" & vbCrLf & vbCrLf
        mensaje = mensaje & "RESULTADO:" & vbCrLf
        mensaje = mensaje & "Volumen Específico = " & volumenEspecifico.ToString("F6") & " cm³/g"

        MessageBox.Show(mensaje, "Conversión a cm³/g", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    ' ============================================
    ' EVENTO: Convertir a ft³/lb
    ' ============================================
    Private Sub btnConvertirFt3lb_Click(sender As Object, e As EventArgs)
        If ultimoVolumenMolar = 0 Then
            MessageBox.Show("Primero debe calcular el volumen molar.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Conversión: m³/mol → ft³/lb
        Dim volumenEspecifico As Double = (ultimoVolumenMolar * 35.3147) / (masaMolar * 2.20462)

        Dim mensaje As String = "=== CONVERSIÓN A ft³/lb ===" & vbCrLf & vbCrLf
        mensaje = mensaje & "Volumen Molar: " & ultimoVolumenMolar.ToString("F8") & " m³/mol" & vbCrLf
        mensaje = mensaje & "Masa Molar: " & masaMolar.ToString("F6") & " kg/mol" & vbCrLf & vbCrLf
        mensaje = mensaje & "RESULTADO:" & vbCrLf
        mensaje = mensaje & "Volumen Específico = " & volumenEspecifico.ToString("F6") & " ft³/lb"

        MessageBox.Show(mensaje, "Conversión a ft³/lb", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    ' ============================================
    ' EVENTO: Convertir a in³/oz
    ' ============================================
    Private Sub btnConvertirIn3oz_Click(sender As Object, e As EventArgs)
        If ultimoVolumenMolar = 0 Then
            MessageBox.Show("Primero debe calcular el volumen molar.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Conversión: m³/mol → in³/oz
        Dim volumenEspecifico As Double = (ultimoVolumenMolar * 61023.7) / (masaMolar * 35.274)

        Dim mensaje As String = "=== CONVERSIÓN A in³/oz ===" & vbCrLf & vbCrLf
        mensaje = mensaje & "Volumen Molar: " & ultimoVolumenMolar.ToString("F8") & " m³/mol" & vbCrLf
        mensaje = mensaje & "Masa Molar: " & masaMolar.ToString("F6") & " kg/mol" & vbCrLf & vbCrLf
        mensaje = mensaje & "RESULTADO:" & vbCrLf
        mensaje = mensaje & "Volumen Específico = " & volumenEspecifico.ToString("F6") & " in³/oz"

        MessageBox.Show(mensaje, "Conversión a in³/oz", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    ' ============================================
    ' EVENTO: Cuando se hace clic en el botón Calcular
    ' ============================================
    Private Sub btnCalcular_Click(sender As Object, e As EventArgs)
        If cboGas.SelectedItem.ToString() = "Otro" Then
            ObtenerConstantesPersonalizadas()
        End If

        If String.IsNullOrWhiteSpace(txtTemperatura.Text) OrElse String.IsNullOrWhiteSpace(txtPresion.Text) Then
            MessageBox.Show("Por favor ingrese valores para Temperatura y Presión.", "Campos vacíos", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Try
            Dim temperatura As Double = Convert.ToDouble(txtTemperatura.Text)
            Dim presion As Double = Convert.ToDouble(txtPresion.Text)
            Dim unidadTemp As String = cboUnidadTemperatura.SelectedItem.ToString()
            Dim unidadPres As String = cboUnidadPresion.SelectedItem.ToString()
            Dim gasSeleccionado As String = cboGas.SelectedItem.ToString()

            ' Guardar valores originales para mostrar
            Dim temperaturaOriginal As Double = temperatura
            Dim presionOriginal As Double = presion

            ' ========== CONVERTIR TEMPERATURA A KELVIN ==========
            Dim temperaturaKelvin As Double = ConvertirAKelvin(temperatura, unidadTemp)

            ' ========== CONVERTIR PRESIÓN A KPA ==========
            Dim presionKPA As Double = ConvertirAKPA(presion, unidadPres)

            ' ========== CALCULAR VOLUMEN MOLAR CON VAN DER WAALS ==========
            Dim resultado As ResultadoVanDerWaals = CalcularVanDerWaals(presionKPA, temperaturaKelvin, constanteA, constanteB)

            ' Guardar el resultado para conversiones posteriores
            ultimoVolumenMolar = resultado.VolumenMolar

            ' Mostrar panel de conversiones
            panelConversiones.Visible = True

            ' ========== CONSTRUIR MENSAJE DE RESULTADOS ==========
            Dim mensaje As String = "=== RESULTADOS DEL CÁLCULO ===" & vbCrLf & vbCrLf

            mensaje = mensaje & "GAS SELECCIONADO: " & gasSeleccionado & vbCrLf & vbCrLf

            mensaje = mensaje & "--- Valores Ingresados ---" & vbCrLf
            mensaje = mensaje & "Temperatura: " & temperaturaOriginal.ToString("F2") & " " & unidadTemp & vbCrLf
            mensaje = mensaje & "Presión: " & presionOriginal.ToString("F2") & " " & unidadPres & vbCrLf & vbCrLf

            mensaje = mensaje & "--- Valores Convertidos ---" & vbCrLf
            mensaje = mensaje & "Temperatura: " & temperaturaKelvin.ToString("F4") & " K" & vbCrLf
            mensaje = mensaje & "Presión: " & presionKPA.ToString("F4") & " kPa" & vbCrLf & vbCrLf

            mensaje = mensaje & "--- Constantes del Gas ---" & vbCrLf
            mensaje = mensaje & "Constante a: " & constanteA.ToString("F10") & " kPa·m⁶/mol²" & vbCrLf
            mensaje = mensaje & "Constante b: " & constanteB.ToString("F8") & " m³/mol" & vbCrLf
            mensaje = mensaje & "Masa Molar: " & masaMolar.ToString("F6") & " kg/mol" & vbCrLf
            mensaje = mensaje & "Constante R: " & R.ToString("F3") & " kPa·m³/(mol·K)" & vbCrLf & vbCrLf

            mensaje = mensaje & "--- Resultados del Cálculo ---" & vbCrLf
            mensaje = mensaje & "Volumen Molar (V): " & resultado.VolumenMolar.ToString("F8") & " m³/mol" & vbCrLf
            mensaje = mensaje & "Volumen Molar (V): " & (resultado.VolumenMolar * 1000).ToString("F5") & " L/mol" & vbCrLf
            mensaje = mensaje & "Iteraciones: " & resultado.Iteraciones.ToString() & vbCrLf
            mensaje = mensaje & "Convergencia alcanzada: " & If(resultado.Convergio, "Sí", "No") & vbCrLf

            If resultado.Iteraciones > 0 Then
                mensaje = mensaje & vbCrLf & "--- Detalles de Iteraciones ---" & vbCrLf
                mensaje = mensaje & resultado.DetalleIteraciones
            End If

            MessageBox.Show(mensaje, "Resultados - Ecuación de Van der Waals", MessageBoxButtons.OK, MessageBoxIcon.Information)
            GuardarResultadosEnArchivo(mensaje, gasSeleccionado)

        Catch ex As Exception
            MessageBox.Show("Error al procesar los datos: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' ============================================
    ' ESTRUCTURA: Para almacenar resultados
    ' ============================================
    Private Structure ResultadoVanDerWaals
        Public VolumenMolar As Double
        Public Iteraciones As Integer
        Public Convergio As Boolean
        Public DetalleIteraciones As String
    End Structure

    ' ============================================
    ' MÉTODO: Calcular volumen molar con Van der Waals
    ' ============================================
    Private Function CalcularVanDerWaals(P As Double, T As Double, a As Double, b As Double) As ResultadoVanDerWaals
        Dim resultado As New ResultadoVanDerWaals
        Dim tolerancia As Double = 0.001
        Dim maxIteraciones As Integer = 100

        ' Valor inicial (Gas Ideal)
        Dim V As Double = (R * T) / P
        Dim V_anterior As Double
        Dim diferencia As Double

        Dim detalles As String = ""
        detalles = detalles & "V₀ (inicial - Gas Ideal) = " & V.ToString("F8") & " m³/mol" & vbCrLf & vbCrLf

        ' Iteraciones de Newton-Raphson
        For i As Integer = 1 To maxIteraciones
            V_anterior = V

            ' Calcular término común (P×b + R×T)
            Dim termino_PbRT As Double = P * b + R * T

            ' Calcular f(V) = P×V³ - (P×b + R×T)×V² + a×V - a×b
            Dim f_V As Double = P * V * V * V - termino_PbRT * V * V + a * V - a * b

            ' Calcular f'(V) = 3×P×V² - 2×(P×b + R×T)×V + a
            Dim f_prima_V As Double = 3 * P * V * V - 2 * termino_PbRT * V + a

            ' Evitar división por cero
            If Math.Abs(f_prima_V) < 0.0000000001 Then
                resultado.Convergio = False
                resultado.VolumenMolar = V
                resultado.Iteraciones = i
                resultado.DetalleIteraciones = detalles & vbCrLf & "Error: Derivada cercana a cero en iteración " & i.ToString()
                Return resultado
            End If

            ' Calcular nuevo V con Newton-Raphson
            V = V_anterior - (f_V / f_prima_V)

            ' Calcular diferencia
            diferencia = Math.Abs(V - V_anterior)

            ' Guardar detalles de esta iteración
            detalles = detalles & "Iteración " & i.ToString() & ":" & vbCrLf
            detalles = detalles & "  V" & (i - 1).ToString() & " = " & V_anterior.ToString("F8") & vbCrLf
            detalles = detalles & "  f(V) = " & f_V.ToString("F10") & vbCrLf
            detalles = detalles & "  f'(V) = " & f_prima_V.ToString("F10") & vbCrLf
            detalles = detalles & "  V" & i.ToString() & " = " & V.ToString("F8") & vbCrLf
            detalles = detalles & "  |V" & i.ToString() & " - V" & (i - 1).ToString() & "| = " & diferencia.ToString("F10") & vbCrLf & vbCrLf

            ' Verificar convergencia
            If diferencia < tolerancia Then
                resultado.Convergio = True
                resultado.VolumenMolar = V
                resultado.Iteraciones = i
                resultado.DetalleIteraciones = detalles & "¡Convergencia alcanzada!"
                Return resultado
            End If
        Next

        ' Si no convergió en el máximo de iteraciones
        resultado.Convergio = False
        resultado.VolumenMolar = V
        resultado.Iteraciones = maxIteraciones
        resultado.DetalleIteraciones = detalles & vbCrLf & "Advertencia: No convergió en " & maxIteraciones.ToString() & " iteraciones."

        Return resultado
    End Function

    ' ============================================
    ' MÉTODO: Guardar resultados en archivo
    ' ============================================
    Private Sub GuardarResultadosEnArchivo(mensaje As String, gasSeleccionado As String)
        Try
            ' Crear nombre de archivo con fecha y hora
            Dim fecha As String = DateTime.Now.ToString("yyyyMMdd_HHmmss")
            Dim nombreArchivo As String = "Resultado_" & gasSeleccionado & "_" & fecha & ".txt"

            ' Ruta completa (se guardará en el escritorio)
            Dim rutaEscritorio As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            Dim rutaCompleta As String = System.IO.Path.Combine(rutaEscritorio, nombreArchivo)

            ' Guardar el contenido en el archivo
            System.IO.File.WriteAllText(rutaCompleta, mensaje)

            ' Notificar al usuario
            MessageBox.Show("Resultados guardados exitosamente en:" & vbCrLf & vbCrLf & rutaCompleta, "Archivo Guardado", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show("Error al guardar el archivo: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' ============================================
    ' MÉTODO: Convertir temperatura a Kelvin
    ' ============================================
    Private Function ConvertirAKelvin(temperatura As Double, unidad As String) As Double
        Select Case unidad
            Case "Kelvin"
                Return temperatura
            Case "Celsius"
                Return temperatura + 273.15
            Case "Fahrenheit"
                Return (temperatura - 32) * 5 / 9 + 273.15
            Case Else
                Return temperatura
        End Select
    End Function

    ' ============================================
    ' MÉTODO: Convertir presión a KPA
    ' ============================================
    Private Function ConvertirAKPA(presion As Double, unidad As String) As Double
        Select Case unidad
            Case "KPA"
                Return presion
            Case "PA"
                Return presion / 1000
            Case "ATM"
                Return presion * 101.325
            Case "BAR"
                Return presion * 100
            Case "PSI"
                Return presion * 6.89476
            Case Else
                Return presion
        End Select
    End Function

    ' ============================================
    ' MÉTODO: Validar que solo se ingresen números
    ' ============================================
    Private Sub ValidarSoloNumeros(sender As Object, e As KeyPressEventArgs)
        If Not Char.IsDigit(e.KeyChar) AndAlso e.KeyChar <> "."c AndAlso e.KeyChar <> "-"c AndAlso e.KeyChar <> Convert.ToChar(Keys.Back) Then
            e.Handled = True
        End If

        Dim txt As TextBox = CType(sender, TextBox)
        If e.KeyChar = "."c AndAlso txt.Text.Contains(". ") Then
            e.Handled = True
        End If

        If e.KeyChar = "-"c AndAlso txt.Text.Length > 0 Then
            e.Handled = True
        End If
    End Sub

    ' ============================================
    ' MÉTODO: Obtener constantes según el gas seleccionado
    ' ============================================
    Private Sub ActualizarConstantesGas()
        Dim gasSeleccionado As String = cboGas.SelectedItem.ToString()

        OcultarCamposPersonalizados()

        Select Case gasSeleccionado
            Case "Agua"
                constanteA = 0.0005536
                constanteB = 0.00003049
                masaMolar = 0.018015

            Case "Amoniaco"
                constanteA = 0.000428
                constanteB = 0.0000371
                masaMolar = 0.017031

            Case "CO2"
                constanteA = 0.000364
                constanteB = 0.00004267
                masaMolar = 0.04401

            Case "CH4"
                constanteA = 0.000228
                constanteB = 0.00004278
                masaMolar = 0.01604

            Case "H2"
                constanteA = 0.0000247
                constanteB = 0.00002661
                masaMolar = 0.002016

            Case "Otro"
                MostrarCamposPersonalizados()

        End Select

        Debug.WriteLine("Gas: " & gasSeleccionado & ", a=" & constanteA.ToString() & ", b=" & constanteB.ToString() & ", MM=" & masaMolar.ToString())
    End Sub

    ' ============================================
    ' MÉTODO: Crear campos personalizados para "Otro"
    ' ============================================
    Private Sub CrearCamposPersonalizados()
        lblConstanteA = New Label()
        lblConstanteA.Text = "Constante a:"
        lblConstanteA.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        lblConstanteA.Location = New Point(50, 220)
        lblConstanteA.AutoSize = True
        lblConstanteA.Visible = False
        Me.Controls.Add(lblConstanteA)

        txtConstanteA = New TextBox()
        txtConstanteA.Font = New Font("Segoe UI", 10)
        txtConstanteA.Location = New Point(50, 245)
        txtConstanteA.Size = New Size(150, 25)
        txtConstanteA.Text = "0"
        txtConstanteA.Visible = False
        AddHandler txtConstanteA.KeyPress, AddressOf ValidarSoloNumeros
        Me.Controls.Add(txtConstanteA)

        lblConstanteB = New Label()
        lblConstanteB.Text = "Constante b:"
        lblConstanteB.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        lblConstanteB.Location = New Point(50, 280)
        lblConstanteB.AutoSize = True
        lblConstanteB.Visible = False
        Me.Controls.Add(lblConstanteB)

        txtConstanteB = New TextBox()
        txtConstanteB.Font = New Font("Segoe UI", 10)
        txtConstanteB.Location = New Point(50, 305)
        txtConstanteB.Size = New Size(150, 25)
        txtConstanteB.Text = "0"
        txtConstanteB.Visible = False
        AddHandler txtConstanteB.KeyPress, AddressOf ValidarSoloNumeros
        Me.Controls.Add(txtConstanteB)

        lblMasaMolar = New Label()
        lblMasaMolar.Text = "Masa Molar:"
        lblMasaMolar.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        lblMasaMolar.Location = New Point(50, 340)
        lblMasaMolar.AutoSize = True
        lblMasaMolar.Visible = False
        Me.Controls.Add(lblMasaMolar)

        txtMasaMolar = New TextBox()
        txtMasaMolar.Font = New Font("Segoe UI", 10)
        txtMasaMolar.Location = New Point(50, 365)
        txtMasaMolar.Size = New Size(150, 25)
        txtMasaMolar.Text = "0"
        txtMasaMolar.Visible = False
        AddHandler txtMasaMolar.KeyPress, AddressOf ValidarSoloNumeros
        Me.Controls.Add(txtMasaMolar)
    End Sub

    ' ============================================
    ' MÉTODO: Mostrar campos personalizados
    ' ============================================
    Private Sub MostrarCamposPersonalizados()
        If lblConstanteA IsNot Nothing Then
            lblConstanteA.Visible = True
            txtConstanteA.Visible = True
            lblConstanteB.Visible = True
            txtConstanteB.Visible = True
            lblMasaMolar.Visible = True
            txtMasaMolar.Visible = True
        End If
    End Sub

    ' ============================================
    ' MÉTODO: Ocultar campos personalizados
    ' ============================================
    Private Sub OcultarCamposPersonalizados()
        If lblConstanteA IsNot Nothing Then
            lblConstanteA.Visible = False
            txtConstanteA.Visible = False
            lblConstanteB.Visible = False
            txtConstanteB.Visible = False
            lblMasaMolar.Visible = False
            txtMasaMolar.Visible = False
        End If
    End Sub

    ' ============================================
    ' MÉTODO: Obtener constantes personalizadas (cuando es "Otro")
    ' ============================================
    Private Sub ObtenerConstantesPersonalizadas()
        If cboGas.SelectedItem.ToString() = "Otro" Then
            Try
                constanteA = Convert.ToDouble(txtConstanteA.Text)
                constanteB = Convert.ToDouble(txtConstanteB.Text)
                masaMolar = Convert.ToDouble(txtMasaMolar.Text)
            Catch ex As Exception
                MessageBox.Show("Por favor ingrese valores numéricos válidos para las constantes.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    ' ============================================
    ' EVENTO: Cuando cambia la selección del gas
    ' ============================================
    Private Sub cboGas_SelectedIndexChanged(sender As Object, e As EventArgs)
        ActualizarConstantesGas()
    End Sub

End Class