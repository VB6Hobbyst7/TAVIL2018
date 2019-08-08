﻿Imports Autodesk.AutoCAD.Interop.Common
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
    Private oUnion As AcadBlockReference = Nothing
    Private oTT As ToolTip = Nothing
    Private capaUnionesVisible As Boolean = True
    Private capaUniones As String = "UNIONES"

    Private Sub frmUniones_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "UNIONES - v" & cfg._appversion
        If clsA Is Nothing Then clsA = New a2.A2acad(Eventos.COMApp, cfg._appFullPath, regAPPCliente)
        ' Cargar recursos
        clsA.ClonaTODODesdeDWG(BloqueRecursos)
        PonEstadoControlesInicial()
        tvUniones_LlenaXDATA()
        PonToolTipControles()
        Dim queCapa As AcadLayer = clsA.CapaDame(capaUniones)
        If queCapa IsNot Nothing Then
            queCapa.LayerOn = True
            queCapa.Lock = False
            queCapa.Freeze = False
            capaUnionesVisible = queCapa.LayerOn
            cbCapa.Checked = capaUnionesVisible
        End If
        clsA.SelectionSet_Delete()
    End Sub
    Public Sub PonEstadoControlesInicial()
        ' Estado inicila de GUniones (Con todos los controles de selección y edición)
        tvUniones.SelectedNode = Nothing
        GUnion.Enabled = False
        BtnCrearUnion.Enabled = True
        BtnEditarUnion.Enabled = False
        BtnBorrarUnion.Enabled = False
        GUnion.Enabled = False
        BtnT1.Enabled = False : BtnT1.BackColor = btnOn
        LblT1.Text = "Datos T1: "
        BtnT2.Enabled = False : BtnT2.BackColor = btnOn
        LblT2.Text = "Datos T2: "
        BtnInsertarUnion.Enabled = False : BtnInsertarUnion.BackColor = btnOn
        LblInsertarUnion.Text = "Datos Unión: "
    End Sub
    Private Sub frmUniones_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        If cbCapa.Enabled = True Then
            Dim queCapa As AcadLayer = clsA.CapaDame(capaUniones)
            If queCapa IsNot Nothing Then
                queCapa.LayerOn = cbCapa.Checked
                queCapa.Freeze = Not (cbCapa.Checked)
            End If
        End If
        frmUn = Nothing
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCerrar.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub
    '
    Private Sub tvUniones_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles tvUniones.AfterSelect
        clsA.Selection_Quitar()
        If tvUniones.SelectedNode Is Nothing Then
            PonEstadoControlesInicial()
            Exit Sub
        End If
        '
        BtnEditarUnion.Enabled = True
        BtnBorrarUnion.Enabled = True
        Me.Uniones_SeleccionarObjetos(tvUniones.SelectedNode.Tag)
    End Sub
    '
    Private Sub tvUniones_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles tvUniones.MouseDoubleClick
        Uniones_SeleccionarObjetos(tvUniones.SelectedNode.Tag, conZoom:=True)
    End Sub
    Private Sub BtnCrearUnion_Click(sender As Object, e As EventArgs) Handles BtnCrearUnion.Click
        oT1 = Nothing
        oT2 = Nothing
        oUnion = Nothing
        GUnion.Enabled = True
        CompruebaDatos()
    End Sub

    Private Sub BtnEditarUnion_Click(sender As Object, e As EventArgs) Handles BtnEditarUnion.Click
        If tvUniones.SelectedNode Is Nothing Then Exit Sub
        '
        GUnion.Enabled = True
        oUnion = Eventos.COMDoc().ObjectIdToObject(tvUniones.SelectedNode.Tag)
        Dim idT1 As String = clsA.XLeeDato(oUnion, "T1")
        Dim idT2 As String = clsA.XLeeDato(oUnion, "T2")
        If idT1 <> "" AndAlso IsNumeric(idT1) Then
            Try
                oT1 = Eventos.COMDoc().ObjectIdToObject(Convert.ToInt64(idT1))
            Catch ex As Exception
            End Try
        End If
        If idT2 <> "" AndAlso IsNumeric(idT2) Then
            Try
                oT2 = Eventos.COMDoc().ObjectIdToObject(Convert.ToInt64(idT2))
            Catch ex As Exception
            End Try
        End If
        CompruebaDatos()
    End Sub
    Private Sub BtnBorrarUnion_Click(sender As Object, e As EventArgs) Handles BtnBorrarUnion.Click
        clsA.Selection_Quitar()
        Dim oBl As AcadBlockReference = Eventos.COMDoc().ObjectIdToObject(tvUniones.SelectedNode.Tag)
        oBl.Delete()
        tvUniones.SelectedNode.Remove()
        tvUniones.SelectedNode = Nothing
        tvUniones_AfterSelect(Nothing, Nothing)
        BtnSeleccionar.Enabled = (tvUniones.Nodes.Count > 0)
    End Sub
    Private Sub BtnSeleccionar_Click(sender As Object, e As EventArgs) Handles BtnSeleccionar.Click
        Dim oBl As AcadBlockReference = clsA.Bloque_SeleccionaDame
        If oBl Is Nothing Then Exit Sub
        '
        tvUniones.SelectedNode = Nothing
        For Each oNode As TreeNode In tvUniones.Nodes
            If oNode.Tag = oBl.ObjectID Then
                tvUniones.SelectedNode = oNode
                tvUniones_MouseDoubleClick(tvUniones, Nothing)  ' Hacer Zoom en el elemento.
                Exit For
            End If
        Next
    End Sub
    Private Sub CbCapa_CheckedChanged(sender As Object, e As EventArgs) Handles cbCapa.CheckedChanged
        Dim queCapa As AcadLayer = clsA.CapaDame(capaUniones)
        If queCapa IsNot Nothing Then
            queCapa.LayerOn = cbCapa.Checked
            queCapa.Freeze = Not (cbCapa.Checked)
            Eventos.COMDoc().Regen(AcRegenType.acActiveViewport)
            PonEstadoControlesInicial()
            tvUniones.Enabled = cbCapa.Checked
            BtnCrearUnion.Enabled = cbCapa.Checked
            BtnSeleccionar.Enabled = cbCapa.Checked
        End If
    End Sub

    Private Sub BtnT1_Click(sender As Object, e As EventArgs) Handles BtnT1.Click
        Me.Visible = False
        clsA.ActivaAppAPI()
        oT1 = clsA.Bloque_SeleccionaDame
        If oT1 IsNot Nothing AndAlso oUnion IsNot Nothing Then
            clsA.XPonDato(oUnion, "T1", oT1.ObjectID)
        End If
        CompruebaDatos()
        Me.Visible = True
    End Sub

    'Private Sub BtnT1_Click_VIEJO(sender As Object, e As EventArgs) Handles BtnT1.Click
    '    Me.Visible = False
    '    clsA.ActivaAppAPI()
    '    Dim arrEntities As ArrayList = Nothing
    '    arrEntities = clsA.SeleccionaDameEntitiesONSCREEN(solouna:=True)
    '    If arrEntities Is Nothing OrElse arrEntities.Count = 0 Then
    '        Me.Visible = True
    '        Exit Sub
    '    End If
    '    '
    '    If TypeOf arrEntities(0) Is AcadBlockReference Then
    '        oT1 = arrEntities(0)
    '    End If
    '    '
    '    CompruebaDatos()
    '    Me.Visible = True
    'End Sub

    Private Sub BtnT2_Click(sender As Object, e As EventArgs) Handles BtnT2.Click
        Me.Visible = False
        clsA.ActivaAppAPI()
        oT2 = clsA.Bloque_SeleccionaDame
        If oT2 IsNot Nothing AndAlso oUnion IsNot Nothing Then
            clsA.XPonDato(oUnion, "T2", oT2.ObjectID)
        End If
        CompruebaDatos()
        Me.Visible = True
    End Sub

    Private Sub BtnInsertarUnion_Click(sender As Object, e As EventArgs) Handles BtnInsertarUnion.Click
        clsA.ActivaAppAPI()
        If oUnion Is Nothing Then
            clsA.Bloque_Inserta(, "UNION")
            If clsA.oBlult IsNot Nothing Then
                oUnion = Eventos.COMDoc.ObjectIdToObject(clsA.oBlult.ObjectID)
                If oT1 IsNot Nothing Then clsA.XPonDato(oUnion, "T1", oT1.ObjectID)
                If oT2 IsNot Nothing Then clsA.XPonDato(oUnion, "T2", oT2.ObjectID)
                tvUniones_PonNodo(oUnion.ObjectID)
                tvUniones.SelectedNode = Nothing
                tvUniones_AfterSelect(Nothing, Nothing)
            End If
        ElseIf oUnion IsNot Nothing Then
            clsA.Bloque_Inserta(, "UNION")
            If clsA.oBlult IsNot Nothing Then
                oUnion.Move(oUnion.InsertionPoint, clsA.oBlult.InsertionPoint)
                clsA.oBlult.Delete()
                clsA.oBlult = oUnion
                tvUniones.SelectedNode = Nothing
                tvUniones_AfterSelect(Nothing, Nothing)
            End If
        End If
        CompruebaDatos()
    End Sub

    Private Sub BtnCancelar_Click(sender As Object, e As EventArgs) Handles BtnCancelar.Click
        oT1 = Nothing
        oT2 = Nothing
        oUnion = Nothing
        tvUniones.SelectedNode = Nothing
        tvUniones_AfterSelect(Nothing, Nothing)
    End Sub
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
        oTT.SetToolTip(Me.BtnBorrarUnion, "Borrar unión seleccionada")
        oTT.SetToolTip(Me.BtnSeleccionar, "Seleccionar bloque UNION en Dibujo")
        oTT.SetToolTip(Me.BtnBorrarUnion, "Mostra/Ocultar capa UNIONES")
        '
        oTT.SetToolTip(Me.BtnT1, "Seleccionar Transportador 1")
        oTT.SetToolTip(Me.BtnT2, "Seleccionar Transportador 2")
        oTT.SetToolTip(Me.BtnInsertarUnion, "Insertar o Mover Unión")
        oTT.SetToolTip(Me.btnCerrar, "Cerrar Uniones")
    End Sub

    Private Sub CompruebaDatos()
        ' Deshabilitar controles, por defecto.
        'pb1.BackColor = btnOff 'Drawing.Color.FromArgb(255, 192, 192)
        'pb2.BackColor = Drawing.Color.FromArgb(255, 192, 192)
        'btnAgregarUnion.BackColor = Drawing.Color.FromArgb(255, 192, 192)
        'pb2.Enabled = False ': Control_Borde(pb2)
        'btnAgregarUnion.Enabled = False ': Control_Borde(btnAgregarUnion)
        'pb1.BackColor = Drawing.Color.FromArgb(255, 192, 192)
        '
        ' Salir, si no esta activo GUnion
        If GUnion.Enabled = False Then Exit Sub

        ' Comprobar objetos. BtnT1 y BtnT2 siempre activo
        BtnT1.Enabled = True
        BtnT2.Enabled = True
        ' *** BtnT1
        If oT1 IsNot Nothing Then
            LblT1.Text = "Datos T1:" & vbCrLf & oT1.EffectiveName
            BtnT1.BackColor = btnOn
        Else
            LblT1.Text = "Datos T1:"
            BtnT1.BackColor = btnOff
        End If
        ' *** BtnT2
        If oT2 IsNot Nothing Then
            LblT2.Text = "Datos T2:" & vbCrLf & oT2.EffectiveName
            BtnT2.BackColor = btnOn
        Else
            LblT2.Text = "Datos T2:"
            BtnT2.BackColor = btnOff
        End If
        ' *** BtnInsertarUnion
        If oUnion IsNot Nothing Then
            LblInsertarUnion.Text = "Datos Unión:" & vbCrLf & oUnion.EffectiveName
            BtnInsertarUnion.BackColor = btnOn
        Else
            LblInsertarUnion.Text = "Datos Unión:"
            BtnInsertarUnion.BackColor = btnOff
        End If
        ' Activación de BtnInsertarUnion
        BtnInsertarUnion.Enabled = (BtnT1.Enabled AndAlso BtnT2.Enabled)
    End Sub
    Public Sub tvUniones_LlenaXDATA()
        ' Rellenar tvGrupos con los grupos que haya ([nombre grupo]) Sacado de XData elementos (regAPPCliente, XData = "GRUPO")
        tvUniones.Nodes.Clear()
        If clsA Is Nothing Then clsA = New a2.A2acad(Eventos.COMApp(), cfg._appFullPath, regAPPCliente)
        'Dim arrTodos As List(Of Long) = clsA.SeleccionaTodosObjetos(,, True)
        Dim arrTodos As List(Of Long) = clsA.SeleccionaDameBloquesPorNombre(cUNION)
        If arrTodos Is Nothing OrElse arrTodos.Count = 0 Then
            Exit Sub
        End If
        ' Filtrar lista de grupo. Sacar nombres únicos.
        For Each queId As Long In arrTodos
            Dim acadObj As AcadObject = Eventos.COMDoc().ObjectIdToObject(queId)
            If TypeOf acadObj Is AcadBlockReference = False Then Continue For
            Dim oBl As AcadBlockReference = CType(acadObj, AcadBlockReference)
            If oBl.EffectiveName <> cUNION Then Continue For
            '
            'Dim union As String = clsA.XLeeDato(acadObj, cUNION)
            'If union = "" Then Continue For
            '
            tvUniones_PonNodo(oBl.ObjectID)
            'Dim nombre As String = oBl.ObjectID 'cUNION '& "·" & oBl.ObjectID
            'If tvUniones.Nodes.ContainsKey(nombre) Then Continue For
            ''
            'Dim oNode As New TreeNode
            'oNode.Text = nombre
            'oNode.Name = nombre
            'oNode.Tag = oBl.ObjectID
            'oNode.ToolTipText = "Unión = " & nombre & "·" & oBl.ObjectID
            'tvUniones.Nodes.Add(oNode)
            'oNode = Nothing
            'acadObj = Nothing
            oBl = Nothing
        Next
        tvUniones.Sort()
        '
        tvUniones.SelectedNode = Nothing
        tvUniones_AfterSelect(Nothing, Nothing)
        BtnSeleccionar.Enabled = (tvUniones.Nodes.Count > 0)
        cbCapa.Enabled = (tvUniones.Nodes.Count > 0)
    End Sub
    Public Sub tvUniones_PonNodo(queId As Long)
        Dim nombre As String = queId.ToString 'cUNION '& "·" & oBl.ObjectID
        If tvUniones.Nodes.ContainsKey(nombre) Then Exit Sub
        '
        Dim oNode As New TreeNode
        oNode.Text = nombre
        oNode.Name = nombre
        oNode.Tag = queId
        oNode.ToolTipText = "Unión = " & nombre & "·" & queId
        tvUniones.Nodes.Add(oNode)
        oNode = Nothing
    End Sub

    Public Sub Uniones_SeleccionarObjetos(IdUnion As Long, Optional conZoom As Boolean = False)
        Dim lGrupos As New List(Of Long)
        lGrupos.Add(IdUnion)
        oUnion = Eventos.COMDoc().ObjectIdToObject(IdUnion)
        Dim idT1 As String = clsA.XLeeDato(oUnion, "T1")
        Dim idT2 As String = clsA.XLeeDato(oUnion, "T2")
        If idT1 <> "" AndAlso IsNumeric(idT1) Then
            Try
                oT1 = Eventos.COMDoc().ObjectIdToObject(Convert.ToInt64(idT1))
                lGrupos.Add(oT1.ObjectID)
            Catch ex As Exception
            End Try
        End If
        If idT2 <> "" AndAlso IsNumeric(idT2) Then
            Try
                oT2 = Eventos.COMDoc().ObjectIdToObject(Convert.ToInt64(idT2))
                lGrupos.Add(oT2.ObjectID)
            Catch ex As Exception
            End Try
        End If

        If lGrupos.Count > 0 Then
            If cbZoom.Checked OrElse conZoom = True Then
                clsA.Selecciona_AcadID(lGrupos.ToArray)
                clsA.HazZoomObjeto(oUnion, 3, False)
            Else
                clsA.Selecciona_AcadID(lGrupos.ToArray)
            End If
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
