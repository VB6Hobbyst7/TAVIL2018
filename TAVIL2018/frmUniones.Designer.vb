<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmUniones
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
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

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmUniones))
        Me.btnCerrar = New System.Windows.Forms.Button()
        Me.tvUniones = New System.Windows.Forms.TreeView()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cbZoom = New System.Windows.Forms.CheckBox()
        Me.BtnInsertarUnion = New System.Windows.Forms.Button()
        Me.GUnion = New System.Windows.Forms.GroupBox()
        Me.LblInsertarUnion = New System.Windows.Forms.Label()
        Me.LblT2 = New System.Windows.Forms.Label()
        Me.LblT1 = New System.Windows.Forms.Label()
        Me.BtnT2 = New System.Windows.Forms.Button()
        Me.BtnT1 = New System.Windows.Forms.Button()
        Me.BtnCrearUnion = New System.Windows.Forms.Button()
        Me.BtnEditarUnion = New System.Windows.Forms.Button()
        Me.BtnBorrarUnion = New System.Windows.Forms.Button()
        Me.BtnCancelar = New System.Windows.Forms.Button()
        Me.GUnion.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnCerrar
        '
        Me.btnCerrar.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCerrar.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCerrar.Location = New System.Drawing.Point(659, 465)
        Me.btnCerrar.Margin = New System.Windows.Forms.Padding(4)
        Me.btnCerrar.Name = "btnCerrar"
        Me.btnCerrar.Size = New System.Drawing.Size(89, 28)
        Me.btnCerrar.TabIndex = 1
        Me.btnCerrar.Text = "Cerrar"
        '
        'tvUniones
        '
        Me.tvUniones.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.tvUniones.Location = New System.Drawing.Point(12, 29)
        Me.tvUniones.Name = "tvUniones"
        Me.tvUniones.Size = New System.Drawing.Size(145, 432)
        Me.tvUniones.TabIndex = 4
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(68, 17)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Uniones :"
        '
        'cbZoom
        '
        Me.cbZoom.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cbZoom.AutoSize = True
        Me.cbZoom.Checked = True
        Me.cbZoom.CheckState = System.Windows.Forms.CheckState.Checked
        Me.cbZoom.Location = New System.Drawing.Point(12, 470)
        Me.cbZoom.Name = "cbZoom"
        Me.cbZoom.Size = New System.Drawing.Size(122, 21)
        Me.cbZoom.TabIndex = 11
        Me.cbZoom.Text = "Zoom Uniones"
        Me.cbZoom.UseVisualStyleBackColor = True
        '
        'BtnInsertarUnion
        '
        Me.BtnInsertarUnion.BackColor = System.Drawing.SystemColors.Control
        Me.BtnInsertarUnion.Location = New System.Drawing.Point(13, 257)
        Me.BtnInsertarUnion.Name = "BtnInsertarUnion"
        Me.BtnInsertarUnion.Size = New System.Drawing.Size(132, 36)
        Me.BtnInsertarUnion.TabIndex = 19
        Me.BtnInsertarUnion.Text = "Insertar Unión"
        Me.BtnInsertarUnion.UseVisualStyleBackColor = False
        '
        'GUnion
        '
        Me.GUnion.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GUnion.Controls.Add(Me.BtnCancelar)
        Me.GUnion.Controls.Add(Me.LblInsertarUnion)
        Me.GUnion.Controls.Add(Me.LblT2)
        Me.GUnion.Controls.Add(Me.LblT1)
        Me.GUnion.Controls.Add(Me.BtnT2)
        Me.GUnion.Controls.Add(Me.BtnT1)
        Me.GUnion.Controls.Add(Me.BtnInsertarUnion)
        Me.GUnion.Location = New System.Drawing.Point(368, 29)
        Me.GUnion.Name = "GUnion"
        Me.GUnion.Size = New System.Drawing.Size(381, 357)
        Me.GUnion.TabIndex = 20
        Me.GUnion.TabStop = False
        Me.GUnion.Text = "Crear / Editar Unión"
        '
        'LblInsertarUnion
        '
        Me.LblInsertarUnion.Location = New System.Drawing.Point(154, 257)
        Me.LblInsertarUnion.Name = "LblInsertarUnion"
        Me.LblInsertarUnion.Size = New System.Drawing.Size(221, 80)
        Me.LblInsertarUnion.TabIndex = 24
        Me.LblInsertarUnion.Text = "Datos Unión:"
        '
        'LblT2
        '
        Me.LblT2.Location = New System.Drawing.Point(154, 140)
        Me.LblT2.Name = "LblT2"
        Me.LblT2.Size = New System.Drawing.Size(221, 80)
        Me.LblT2.TabIndex = 23
        Me.LblT2.Text = "Datos T2:"
        '
        'LblT1
        '
        Me.LblT1.Location = New System.Drawing.Point(154, 29)
        Me.LblT1.Name = "LblT1"
        Me.LblT1.Size = New System.Drawing.Size(221, 80)
        Me.LblT1.TabIndex = 22
        Me.LblT1.Text = "Datos T1:"
        '
        'BtnT2
        '
        Me.BtnT2.Location = New System.Drawing.Point(13, 140)
        Me.BtnT2.Name = "BtnT2"
        Me.BtnT2.Size = New System.Drawing.Size(132, 41)
        Me.BtnT2.TabIndex = 21
        Me.BtnT2.Text = "Transportador 2"
        Me.BtnT2.UseVisualStyleBackColor = True
        '
        'BtnT1
        '
        Me.BtnT1.Location = New System.Drawing.Point(13, 29)
        Me.BtnT1.Name = "BtnT1"
        Me.BtnT1.Size = New System.Drawing.Size(132, 41)
        Me.BtnT1.TabIndex = 20
        Me.BtnT1.Text = "Transportador 1"
        Me.BtnT1.UseVisualStyleBackColor = True
        '
        'BtnCrearUnion
        '
        Me.BtnCrearUnion.Location = New System.Drawing.Point(178, 29)
        Me.BtnCrearUnion.Name = "BtnCrearUnion"
        Me.BtnCrearUnion.Size = New System.Drawing.Size(171, 32)
        Me.BtnCrearUnion.TabIndex = 21
        Me.BtnCrearUnion.Text = "Crear Unión"
        Me.BtnCrearUnion.UseVisualStyleBackColor = True
        '
        'BtnEditarUnion
        '
        Me.BtnEditarUnion.Location = New System.Drawing.Point(178, 86)
        Me.BtnEditarUnion.Name = "BtnEditarUnion"
        Me.BtnEditarUnion.Size = New System.Drawing.Size(171, 32)
        Me.BtnEditarUnion.TabIndex = 22
        Me.BtnEditarUnion.Text = "Editar Unión"
        Me.BtnEditarUnion.UseVisualStyleBackColor = True
        '
        'BtnBorrarUnion
        '
        Me.BtnBorrarUnion.Location = New System.Drawing.Point(178, 143)
        Me.BtnBorrarUnion.Name = "BtnBorrarUnion"
        Me.BtnBorrarUnion.Size = New System.Drawing.Size(171, 32)
        Me.BtnBorrarUnion.TabIndex = 23
        Me.BtnBorrarUnion.Text = "Borrar Unión"
        Me.BtnBorrarUnion.UseVisualStyleBackColor = True
        '
        'BtnCancelar
        '
        Me.BtnCancelar.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.BtnCancelar.BackColor = System.Drawing.SystemColors.Control
        Me.BtnCancelar.Location = New System.Drawing.Point(8, 323)
        Me.BtnCancelar.Name = "BtnCancelar"
        Me.BtnCancelar.Size = New System.Drawing.Size(74, 27)
        Me.BtnCancelar.TabIndex = 25
        Me.BtnCancelar.Text = "Cancelar"
        Me.BtnCancelar.UseVisualStyleBackColor = False
        '
        'frmUniones
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCerrar
        Me.ClientSize = New System.Drawing.Size(761, 503)
        Me.Controls.Add(Me.BtnBorrarUnion)
        Me.Controls.Add(Me.BtnEditarUnion)
        Me.Controls.Add(Me.BtnCrearUnion)
        Me.Controls.Add(Me.GUnion)
        Me.Controls.Add(Me.btnCerrar)
        Me.Controls.Add(Me.cbZoom)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.tvUniones)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmUniones"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "frmUniones"
        Me.GUnion.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnCerrar As System.Windows.Forms.Button
    Friend WithEvents tvUniones As Windows.Forms.TreeView
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents cbZoom As Windows.Forms.CheckBox
    Friend WithEvents BtnInsertarUnion As Windows.Forms.Button
    Friend WithEvents GUnion As Windows.Forms.GroupBox
    Friend WithEvents BtnT2 As Windows.Forms.Button
    Friend WithEvents BtnT1 As Windows.Forms.Button
    Friend WithEvents LblInsertarUnion As Windows.Forms.Label
    Friend WithEvents LblT2 As Windows.Forms.Label
    Friend WithEvents LblT1 As Windows.Forms.Label
    Friend WithEvents BtnCrearUnion As Windows.Forms.Button
    Friend WithEvents BtnEditarUnion As Windows.Forms.Button
    Friend WithEvents BtnBorrarUnion As Windows.Forms.Button
    Friend WithEvents BtnCancelar As Windows.Forms.Button
End Class
