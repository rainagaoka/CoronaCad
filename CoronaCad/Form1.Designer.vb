<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Descartar substituições de formulário para limpar a lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Exigido pelo Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'OBSERVAÇÃO: o procedimento a seguir é exigido pelo Windows Form Designer
    'Pode ser modificado usando o Windows Form Designer.  
    'Não o modifique usando o editor de códigos.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.TextBox_Caminho = New System.Windows.Forms.TextBox()
        Me.Abrir = New System.Windows.Forms.Button()
        Me.Label_GPS = New System.Windows.Forms.Label()
        Me.RadioButton_Linear = New System.Windows.Forms.RadioButton()
        Me.RadioButton_Logarítmica = New System.Windows.Forms.RadioButton()
        Me.RadioButton_CasosNovos = New System.Windows.Forms.RadioButton()
        Me.RadioButton_CasosAcumulados = New System.Windows.Forms.RadioButton()
        Me.RadioButton_ObitosAcumulados = New System.Windows.Forms.RadioButton()
        Me.RadioButton_ObitosNovos = New System.Windows.Forms.RadioButton()
        Me.GroupBox_Dados = New System.Windows.Forms.GroupBox()
        Me.GroupBox_Escala = New System.Windows.Forms.GroupBox()
        Me.Button_Gerar = New System.Windows.Forms.Button()
        Me.Button_Cancelar = New System.Windows.Forms.Button()
        Me.Button_Limpar = New System.Windows.Forms.Button()
        Me.GroupBox_Dados.SuspendLayout()
        Me.GroupBox_Escala.SuspendLayout()
        Me.SuspendLayout()
        '
        'TextBox_Caminho
        '
        Me.TextBox_Caminho.Location = New System.Drawing.Point(157, 6)
        Me.TextBox_Caminho.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TextBox_Caminho.Name = "TextBox_Caminho"
        Me.TextBox_Caminho.Size = New System.Drawing.Size(381, 22)
        Me.TextBox_Caminho.TabIndex = 0
        '
        'Abrir
        '
        Me.Abrir.Location = New System.Drawing.Point(545, 6)
        Me.Abrir.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Abrir.Name = "Abrir"
        Me.Abrir.Size = New System.Drawing.Size(75, 23)
        Me.Abrir.TabIndex = 1
        Me.Abrir.Text = "Abrir"
        Me.Abrir.UseVisualStyleBackColor = True
        '
        'Label_GPS
        '
        Me.Label_GPS.AutoSize = True
        Me.Label_GPS.Location = New System.Drawing.Point(13, 9)
        Me.Label_GPS.Name = "Label_GPS"
        Me.Label_GPS.Size = New System.Drawing.Size(137, 17)
        Me.Label_GPS.TabIndex = 6
        Me.Label_GPS.Text = "Selecione o arquivo:"
        '
        'RadioButton_Linear
        '
        Me.RadioButton_Linear.AutoSize = True
        Me.RadioButton_Linear.Location = New System.Drawing.Point(8, 23)
        Me.RadioButton_Linear.Margin = New System.Windows.Forms.Padding(4)
        Me.RadioButton_Linear.Name = "RadioButton_Linear"
        Me.RadioButton_Linear.Size = New System.Drawing.Size(69, 21)
        Me.RadioButton_Linear.TabIndex = 8
        Me.RadioButton_Linear.TabStop = True
        Me.RadioButton_Linear.Text = "Linear"
        Me.RadioButton_Linear.UseVisualStyleBackColor = True
        '
        'RadioButton_Logarítmica
        '
        Me.RadioButton_Logarítmica.AutoSize = True
        Me.RadioButton_Logarítmica.Location = New System.Drawing.Point(8, 52)
        Me.RadioButton_Logarítmica.Margin = New System.Windows.Forms.Padding(4)
        Me.RadioButton_Logarítmica.Name = "RadioButton_Logarítmica"
        Me.RadioButton_Logarítmica.Size = New System.Drawing.Size(102, 21)
        Me.RadioButton_Logarítmica.TabIndex = 9
        Me.RadioButton_Logarítmica.TabStop = True
        Me.RadioButton_Logarítmica.Text = "Logarítmica"
        Me.RadioButton_Logarítmica.UseVisualStyleBackColor = True
        '
        'RadioButton_CasosNovos
        '
        Me.RadioButton_CasosNovos.AutoSize = True
        Me.RadioButton_CasosNovos.Location = New System.Drawing.Point(8, 23)
        Me.RadioButton_CasosNovos.Margin = New System.Windows.Forms.Padding(4)
        Me.RadioButton_CasosNovos.Name = "RadioButton_CasosNovos"
        Me.RadioButton_CasosNovos.Size = New System.Drawing.Size(112, 21)
        Me.RadioButton_CasosNovos.TabIndex = 11
        Me.RadioButton_CasosNovos.TabStop = True
        Me.RadioButton_CasosNovos.Text = "Casos Novos"
        Me.RadioButton_CasosNovos.UseVisualStyleBackColor = True
        '
        'RadioButton_CasosAcumulados
        '
        Me.RadioButton_CasosAcumulados.AutoSize = True
        Me.RadioButton_CasosAcumulados.Location = New System.Drawing.Point(8, 52)
        Me.RadioButton_CasosAcumulados.Margin = New System.Windows.Forms.Padding(4)
        Me.RadioButton_CasosAcumulados.Name = "RadioButton_CasosAcumulados"
        Me.RadioButton_CasosAcumulados.Size = New System.Drawing.Size(149, 21)
        Me.RadioButton_CasosAcumulados.TabIndex = 12
        Me.RadioButton_CasosAcumulados.TabStop = True
        Me.RadioButton_CasosAcumulados.Text = "Casos Acumulados"
        Me.RadioButton_CasosAcumulados.UseVisualStyleBackColor = True
        '
        'RadioButton_ObitosAcumulados
        '
        Me.RadioButton_ObitosAcumulados.AutoSize = True
        Me.RadioButton_ObitosAcumulados.Location = New System.Drawing.Point(8, 108)
        Me.RadioButton_ObitosAcumulados.Margin = New System.Windows.Forms.Padding(4)
        Me.RadioButton_ObitosAcumulados.Name = "RadioButton_ObitosAcumulados"
        Me.RadioButton_ObitosAcumulados.Size = New System.Drawing.Size(151, 21)
        Me.RadioButton_ObitosAcumulados.TabIndex = 13
        Me.RadioButton_ObitosAcumulados.TabStop = True
        Me.RadioButton_ObitosAcumulados.Text = "Óbitos Acumulados"
        Me.RadioButton_ObitosAcumulados.UseVisualStyleBackColor = True
        '
        'RadioButton_ObitosNovos
        '
        Me.RadioButton_ObitosNovos.AutoSize = True
        Me.RadioButton_ObitosNovos.Location = New System.Drawing.Point(8, 80)
        Me.RadioButton_ObitosNovos.Margin = New System.Windows.Forms.Padding(4)
        Me.RadioButton_ObitosNovos.Name = "RadioButton_ObitosNovos"
        Me.RadioButton_ObitosNovos.Size = New System.Drawing.Size(114, 21)
        Me.RadioButton_ObitosNovos.TabIndex = 14
        Me.RadioButton_ObitosNovos.TabStop = True
        Me.RadioButton_ObitosNovos.Text = "Óbitos Novos"
        Me.RadioButton_ObitosNovos.UseVisualStyleBackColor = True
        '
        'GroupBox_Dados
        '
        Me.GroupBox_Dados.Controls.Add(Me.RadioButton_CasosNovos)
        Me.GroupBox_Dados.Controls.Add(Me.RadioButton_ObitosAcumulados)
        Me.GroupBox_Dados.Controls.Add(Me.RadioButton_ObitosNovos)
        Me.GroupBox_Dados.Controls.Add(Me.RadioButton_CasosAcumulados)
        Me.GroupBox_Dados.Location = New System.Drawing.Point(16, 37)
        Me.GroupBox_Dados.Margin = New System.Windows.Forms.Padding(4)
        Me.GroupBox_Dados.Name = "GroupBox_Dados"
        Me.GroupBox_Dados.Padding = New System.Windows.Forms.Padding(4)
        Me.GroupBox_Dados.Size = New System.Drawing.Size(293, 140)
        Me.GroupBox_Dados.TabIndex = 16
        Me.GroupBox_Dados.TabStop = False
        Me.GroupBox_Dados.Text = "Dados:"
        '
        'GroupBox_Escala
        '
        Me.GroupBox_Escala.Controls.Add(Me.RadioButton_Logarítmica)
        Me.GroupBox_Escala.Controls.Add(Me.RadioButton_Linear)
        Me.GroupBox_Escala.Location = New System.Drawing.Point(319, 37)
        Me.GroupBox_Escala.Margin = New System.Windows.Forms.Padding(4)
        Me.GroupBox_Escala.Name = "GroupBox_Escala"
        Me.GroupBox_Escala.Padding = New System.Windows.Forms.Padding(4)
        Me.GroupBox_Escala.Size = New System.Drawing.Size(301, 140)
        Me.GroupBox_Escala.TabIndex = 17
        Me.GroupBox_Escala.TabStop = False
        Me.GroupBox_Escala.Text = "Escala:"
        '
        'Button_Gerar
        '
        Me.Button_Gerar.Location = New System.Drawing.Point(535, 183)
        Me.Button_Gerar.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Button_Gerar.Name = "Button_Gerar"
        Me.Button_Gerar.Size = New System.Drawing.Size(80, 27)
        Me.Button_Gerar.TabIndex = 18
        Me.Button_Gerar.Text = "Gerar"
        Me.Button_Gerar.UseVisualStyleBackColor = True
        '
        'Button_Cancelar
        '
        Me.Button_Cancelar.Location = New System.Drawing.Point(449, 183)
        Me.Button_Cancelar.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Button_Cancelar.Name = "Button_Cancelar"
        Me.Button_Cancelar.Size = New System.Drawing.Size(80, 27)
        Me.Button_Cancelar.TabIndex = 19
        Me.Button_Cancelar.Text = "Cancelar"
        Me.Button_Cancelar.UseVisualStyleBackColor = True
        '
        'Button_Limpar
        '
        Me.Button_Limpar.Location = New System.Drawing.Point(364, 183)
        Me.Button_Limpar.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Button_Limpar.Name = "Button_Limpar"
        Me.Button_Limpar.Size = New System.Drawing.Size(80, 27)
        Me.Button_Limpar.TabIndex = 20
        Me.Button_Limpar.Text = "Limpar"
        Me.Button_Limpar.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(629, 218)
        Me.Controls.Add(Me.Button_Limpar)
        Me.Controls.Add(Me.Button_Cancelar)
        Me.Controls.Add(Me.Button_Gerar)
        Me.Controls.Add(Me.GroupBox_Escala)
        Me.Controls.Add(Me.GroupBox_Dados)
        Me.Controls.Add(Me.Label_GPS)
        Me.Controls.Add(Me.Abrir)
        Me.Controls.Add(Me.TextBox_Caminho)
        Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.GroupBox_Dados.ResumeLayout(False)
        Me.GroupBox_Dados.PerformLayout()
        Me.GroupBox_Escala.ResumeLayout(False)
        Me.GroupBox_Escala.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents TextBox_Caminho As TextBox
    Friend WithEvents Abrir As Button
    Friend WithEvents Label_GPS As Label
    Friend WithEvents RadioButton_Linear As RadioButton
    Friend WithEvents RadioButton_Logarítmica As RadioButton
    Friend WithEvents RadioButton_CasosNovos As RadioButton
    Friend WithEvents RadioButton_CasosAcumulados As RadioButton
    Friend WithEvents RadioButton_ObitosAcumulados As RadioButton
    Friend WithEvents RadioButton_ObitosNovos As RadioButton
    Friend WithEvents GroupBox_Dados As GroupBox
    Friend WithEvents GroupBox_Escala As GroupBox
    Friend WithEvents Button_Gerar As Button
    Friend WithEvents Button_Cancelar As Button
    Friend WithEvents Button_Limpar As Button
End Class
