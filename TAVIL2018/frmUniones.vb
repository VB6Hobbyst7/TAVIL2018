Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports TAVIL2018.TAVIL2018
Imports System.Drawing
Imports System.Windows.Forms
Imports uau = UtilesAlberto.Utiles
Imports a2 = AutoCAD2acad.A2acad

Public Class frmUniones
    Private HighlightedPictureBox As PictureBox = Nothing
    Private oT1 As AcadBlockReference = Nothing
    Private oT2 As AcadBlockReference = Nothing
    Private oTT As ToolTip = Nothing

    Private Sub frmUniones_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "UNIONES - v" & cfg._appversion
        If clsA Is Nothing Then clsA = New a2.A2acad(Eventos.COMApp, cfg._appFullPath, regAPPCliente)
        ' Estado inicila de GUniones (Con todos los controles de selección y edición)
        GUnion.Enabled = False
        ' Controles fondo rojo claro. Para avisar que hay que pulsarlos. Solo activo Pb1, por defecto
        'pb1.Enabled = True : pb1.BackColor = Drawing.Color.FromArgb(255, 192, 192)
        'pb2.Enabled = False : pb2.BackColor = Drawing.Color.FromArgb(255, 192, 192)
        BtnInsertarUnion.Enabled = False : BtnInsertarUnion.BackColor = Drawing.Color.FromArgb(255, 192, 192)
        '
        PonToolTipControles()
        'CompruebaDatos()
        'app_procesointerno = True
        'app_procesointerno = True
        'gbAdministrar.Enabled = False
        'btnGrupoBorrar.Enabled = False
        'txtNombreGrupo.Text = ""
        'lblInf.Text = ""
        tvUniones_LlenaXDATA()
    End Sub
    Private Sub frmUniones_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        frmUn = Nothing
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCerrar.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub
    '
    Private Sub tvUniones_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles tvUniones.AfterSelect
        If tvUniones.SelectedNode Is Nothing Then Exit Sub
        '
        Me.Uniones_SeleccionarObjetos(tvUniones.SelectedNode.Text)
    End Sub
    '
    Private Sub BtnCrearUnion_Click(sender As Object, e As EventArgs) Handles BtnCrearUnion.Click

    End Sub

    Private Sub BtnEditarUnion_Click(sender As Object, e As EventArgs) Handles BtnEditarUnion.Click

    End Sub

    Private Sub BtnT1_Click(sender As Object, e As EventArgs) Handles BtnT1.Click
        Me.Visible = False
        clsA.ActivaAppAPI()
        Dim arrEntities As ArrayList = Nothing
        arrEntities = clsA.SeleccionaDameEntitiesONSCREEN(solouna:=True)
        If arrEntities Is Nothing OrElse arrEntities.Count = 0 Then
            Exit Sub
        End If
        '
        If TypeOf arrEntities(0) Is AcadBlockReference Then
            oT1 = arrEntities(0)
        End If
        '
        CompruebaDatos()
        Me.Visible = True
    End Sub

    Private Sub BtnT2_Click(sender As Object, e As EventArgs) Handles BtnT2.Click

    End Sub

    Private Sub BtnInsertarUnion_Click(sender As Object, e As EventArgs) Handles BtnInsertarUnion.Click

    End Sub

    Private Sub Pb1_Click(sender As Object, e As EventArgs)
    End Sub
    Private Sub Pb2_Click(sender As Object, e As EventArgs)
        'pb2.BackColor = Control.DefaultBackColor : pb2.Refresh()
        'Control_Borde(pb2, False)
        '
        Me.Visible = False
        'Ev.EvApp.ActiveDocument.Activate()
        Dim arrEntities As ArrayList = Nothing
REPITE:
        arrEntities = clsA.SeleccionaDameEntitiesONSCREEN(solouna:=True)
        If arrEntities Is Nothing OrElse arrEntities.Count = 0 Then
            Exit Sub
        End If
        '
        If TypeOf arrEntities(0) Is AcadBlockReference Then
            oT2 = arrEntities(0)
        End If
        '
        CompruebaDatos()
        Me.Visible = True
    End Sub
    '
    'Efecto Click en PictureBox. Click (Borde Rojo), otro click (Sin borde)  
    'Private Sub pb1_Click(sender As Object, e As EventArgs) Handles pb1.Click, pb2.Click 'etc  
    '    Dim pbX As PictureBox = DirectCast(sender, PictureBox)
    '    'Get the rectangle of the control and inflate it to represent the border area  
    '    Dim BorderBounds As Rectangle = pbX.ClientRectangle
    '    BorderBounds.Inflate(-1, -1)

    '    'Use ControlPaint to draw the border.  
    '    'Change the Color.Red parameter to your own colour below.  
    '    ControlPaint.DrawBorder(pbX.CreateGraphics,
    '                            BorderBounds,
    '                            Color.Red,
    '                            ButtonBorderStyle.Solid)

    '    If Not (HighlightedPictureBox Is Nothing) Then
    '        'Remove the border of the last PictureBox  
    '        HighlightedPictureBox.Invalidate()
    '    End If

    '    'Rememeber the last highlighted PictureBox  
    '    HighlightedPictureBox = CType(sender, PictureBox)
    '    pbX = Nothing
    'End Sub
#Region "FUNCIONES"
    Public Sub PonToolTipControles()
        oTT = New ToolTip()
        oTT.AutoPopDelay = 5000 ' Tiempo que estará visible
        oTT.InitialDelay = 500  ' Tiempo inicial para mostrarse
        oTT.ReshowDelay = 100   ' Tiempo de espera entre controles
        oTT.ShowAlways = True   ' Forzar a que se muestre el tooltip, aunque no este activo el Form.
        '
        Me.tvUniones.ShowNodeToolTips = True
        oTT.SetToolTip(Me.tvUniones, "Listado de uniones")
        oTT.SetToolTip(Me.cbZoom, "Zoom uniones")
        '
        oTT.SetToolTip(Me.BtnCrearUnion, "Crear nueva unión")
        oTT.SetToolTip(Me.BtnEditarUnion, "Editar unión seleccionada")
        '
        oTT.SetToolTip(Me.BtnT1, "Seleccionar Transportador 1")
        oTT.SetToolTip(Me.BtnT2, "Seleccionar Transportador 2")
        oTT.SetToolTip(Me.BtnInsertarUnion, "Insertar Unión")
        oTT.SetToolTip(Me.btnCerrar, "Cerrar Uniones")
    End Sub

    Private Sub CompruebaDatos()
        ' Deshabilitar controles, por defecto.
        'pb1.BackColor = Drawing.Color.FromArgb(255, 192, 192)
        'pb2.BackColor = Drawing.Color.FromArgb(255, 192, 192)
        'btnAgregarUnion.BackColor = Drawing.Color.FromArgb(255, 192, 192)
        '
        'pb1.Enabled = False
        'pb2.Enabled = False ': Control_Borde(pb2)
        'btnAgregarUnion.Enabled = False ': Control_Borde(btnAgregarUnion)
        'pb1.BackColor = Drawing.Color.FromArgb(255, 192, 192)
        '
        ' Comprobar objetos.
        'If oT1 IsNot Nothing Then
        '    lblInf.Text = "T1 = " & oT1.EffectiveName & vbCrLf
        '    pb2.Enabled = True
        'Else
        '    lblInf.Text = "T1 =" & vbCrLf
        '    lblInf.Text &= "T2 ="
        '    pb2.Enabled = False : pb2.BackColor = Drawing.Color.FromArgb(255, 192, 192)
        '    BtnInsertarUnion.Enabled = False : BtnInsertarUnion.BackColor = Drawing.Color.FromArgb(255, 192, 192)
        '    Exit Sub
        'End If
        ''
        'If oT2 IsNot Nothing Then
        '    lblInf.Text &= "T2 = " & oT2.EffectiveName & vbCrLf
        '    BtnInsertarUnion.Enabled = True
        'Else
        '    lblInf.Text &= "T2 ="
        'End If
    End Sub
    Public Sub tvUniones_LlenaXDATA()
        ' Rellenar tvGrupos con los grupos que haya ([nombre grupo]) Sacado de XData elementos (regAPPCliente, XData = "GRUPO")
        tvUniones.Nodes.Clear()
        If clsA Is Nothing Then clsA = New a2.A2acad(Eventos.COMApp, cfg._appFullPath, regAPPCliente)
        Dim arrTodos As List(Of Long) = clsA.SeleccionaTodosObjetos(,, True)
        If arrTodos Is Nothing OrElse arrTodos.Count = 0 Then
            Exit Sub
        End If
        ' Filtrar lista de grupo. Sacar nombres únicos.
        For Each queId As Long In arrTodos
            Dim acadObj As AcadObject = Eventos.COMDoc().ObjectIdToObject(queId)
            Dim union As String = clsA.XLeeDato(acadObj, cUNION)
            If union = "" Then Continue For
            '
            If tvUniones.Nodes.ContainsKey(union) Then Continue For
            '
            Dim oNode As New TreeNode
            oNode.Text = union
            oNode.Name = union
            oNode.Tag = union
            oNode.ToolTipText = oNode.Tag
            tvUniones.Nodes.Add(oNode)
            oNode = Nothing
            acadObj = Nothing
        Next
        tvUniones.Sort()
        '
        tvUniones.SelectedNode = Nothing
        tvUniones_AfterSelect(Nothing, Nothing)
    End Sub

    Public Sub Uniones_SeleccionarObjetos(union As String)
        Dim lGrupos As List(Of Long) = clsA.SeleccionaTodosObjetosXData(cUNION, union)
        If lGrupos IsNot Nothing AndAlso lGrupos.Count > 0 Then
            Dim arrSeleccion As New ArrayList
            For Each queId As Long In lGrupos
                arrSeleccion.Add(Eventos.COMDoc().ObjectIdToObject(queId))
            Next
            If arrSeleccion.Count > 0 Then
                'lblInf.Text = arrSeleccion.Count & " Elementos"
                clsA.SeleccionCreaResalta(arrSeleccion, 0, False)
                If cbZoom.Checked Then
                    'Zoom_Seleccion()
                End If
            End If
        Else
            'If tvGrupos.Nodes.ContainsKey(grupo) Then
            '    tvGrupos.Nodes.Item(grupo).Remove()
            'End If
            'lblInf.Text = "0 Elementos"
            'tvGrupos.SelectedNode = Nothing
            'tvGrupos_AfterSelect(Nothing, Nothing)
        End If
    End Sub

    ' PictureBox (U otros controles). Poner el borde de color rojo  
    Private Sub Control_Borde(oC As Control, Optional restaurar As Boolean = False)
        ' Restaurar botón.
        If restaurar = True Then
            oC.Invalidate()
            Exit Sub
        End If
        ' Poner borde Rojo
        If TypeOf oC Is PictureBox Then
            CType(oC, PictureBox).BorderStyle = BorderStyle.None
        End If
        'Rectangulo del control + Offset del rectangulo hacia fuera.  
        Dim BorderBounds As Rectangle = oC.ClientRectangle
        'BorderBounds.Inflate(-1, -1)

        'Use ControlPaint to draw the border.  
        'Change the Color.Red parameter to your own colour below.  
        ControlPaint.DrawBorder(oC.CreateGraphics,
                                        BorderBounds,
                                        Color.Red,
                                        ButtonBorderStyle.Outset)
    End Sub
#End Region
End Class
