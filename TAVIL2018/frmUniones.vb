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
        ' Cargar recursos
        clsA.ClonaTODODesdeDWG(BloqueRecursos)
        PonEstadoControlesInicial()
        tvUniones_LlenaXDATA()
        PonToolTipControles()
    End Sub
    Public Sub PonEstadoControlesInicial()
        ' Estado inicila de GUniones (Con todos los controles de selección y edición)
        tvUniones.SelectedNode = Nothing
        GUnion.Enabled = False
        BtnCrearUnion.Enabled = True
        BtnEditarUnion.Enabled = False
        BtnBorrarUnion.Enabled = False
        BtnT1.Enabled = False : BtnT1.BackColor = Control.DefaultBackColor
        LblT1.Text = "Datos T1: "
        BtnT2.Enabled = False : BtnT2.BackColor = Control.DefaultBackColor
        LblT2.Text = "Datos T2: "
        BtnInsertarUnion.Enabled = False : BtnInsertarUnion.BackColor = Control.DefaultBackColor
        LblInsertarUnion.Text = "Datos Unión: "
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
        Eventos.COMDoc().ActiveSelectionSet.Clear()
        If tvUniones.SelectedNode Is Nothing Then
            PonEstadoControlesInicial()
            Exit Sub
        End If
        '
        Me.Uniones_SeleccionarObjetos(tvUniones.SelectedNode.Tag)
    End Sub
    '
    Private Sub tvUniones_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles tvUniones.MouseDoubleClick
        Dim oNode As TreeNode = CType(sender, TreeView).SelectedNode
        Dim oBlo As Autodesk.AutoCAD.Interop.Common.AcadBlockReference = Eventos.COMDoc().ObjectIdToObject(oNode.Tag)
        Dim oIPrt As New IntPtr(oBlo.ObjectID)
        Dim oId As New ObjectId(oIPrt)
        Dim arrIds() As ObjectId = {oId}
        Autodesk.AutoCAD.Internal.Utils.SelectObjects(arrIds)
        clsA.HazZoomObjeto(oBlo, 2)
        oBlo = Nothing
    End Sub
    Private Sub BtnCrearUnion_Click(sender As Object, e As EventArgs) Handles BtnCrearUnion.Click
        GUnion.Enabled = True
        BtnT1.Enabled = True : BtnT1.BackColor = Drawing.Color.FromArgb(255, 192, 192)
        BtnT2.Enabled = False : BtnT2.BackColor = Drawing.Color.FromArgb(255, 192, 192)
        BtnInsertarUnion.Enabled = False : BtnInsertarUnion.BackColor = Drawing.Color.FromArgb(255, 192, 192)
    End Sub

    Private Sub BtnEditarUnion_Click(sender As Object, e As EventArgs) Handles BtnEditarUnion.Click
        GUnion.Enabled = True
        BtnT1.Enabled = True : BtnT1.BackColor = Control.DefaultBackColor
        BtnT2.Enabled = False : BtnT2.BackColor = Control.DefaultBackColor
        BtnInsertarUnion.Enabled = False : BtnInsertarUnion.BackColor = Control.DefaultBackColor
    End Sub
    Private Sub BtnBorrarUnion_Click(sender As Object, e As EventArgs) Handles BtnBorrarUnion.Click
        Dim oSet As Autodesk.AutoCAD.Interop.AcadSelectionSet = Eventos.COMDoc().ActiveSelectionSet
        If oSet Is Nothing OrElse oSet.Count < 3 Then
            MsgBox("Esta unión está incompleta, borrarla a mano desde AutoCAD.")
            Exit Sub
        End If
        '
        For x As Integer = 0 To oSet.Count - 1
            Dim oBl As AcadBlockReference = CType(oSet.Item(x), AcadBlockReference)
            If oBl.EffectiveName = "UNION" Then
                oSet.Item(x).Delete()
                oSet.Clear()
                tvUniones.SelectedNode.Remove()
                tvUniones.SelectedNode = Nothing
                tvUniones_AfterSelect(Nothing, Nothing)
                Exit For
            End If
        Next
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
        Me.Visible = False
        clsA.ActivaAppAPI()
        Dim arrEntities As ArrayList = Nothing
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

    Private Sub BtnInsertarUnion_Click(sender As Object, e As EventArgs) Handles BtnInsertarUnion.Click

    End Sub

    Private Sub BtnCancelar_Click(sender As Object, e As EventArgs) Handles BtnCancelar.Click
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
            Dim nombre As String = cUNION & "·" & oBl.ObjectID
            If tvUniones.Nodes.ContainsKey(nombre) Then Continue For
            '
            Dim oNode As New TreeNode
            oNode.Text = nombre
            oNode.Name = nombre
            oNode.Tag = oBl.ObjectID
            oNode.ToolTipText = "Unión = " & nombre
            tvUniones.Nodes.Add(oNode)
            oNode = Nothing
            acadObj = Nothing
            oBl = Nothing
        Next
        tvUniones.Sort()
        '
        tvUniones.SelectedNode = Nothing
        tvUniones_AfterSelect(Nothing, Nothing)
    End Sub

    Public Sub Uniones_SeleccionarObjetos(IdUnion As Long)
        Dim lGrupos As List(Of Long) = clsA.SeleccionaTodosObjetosXData(cUNION, IdUnion.ToString, igual:=False)
        lGrupos.Add(IdUnion)
        If lGrupos IsNot Nothing AndAlso lGrupos.Count > 0 Then
            Dim arrSeleccion As New ArrayList
            For Each queId As Long In lGrupos
                arrSeleccion.Add(Eventos.COMDoc().ObjectIdToObject(queId))
            Next
            If arrSeleccion.Count > 0 Then
                'lblInf.Text = arrSeleccion.Count & " Elementos"
                If cbZoom.Checked Then
                    tvUniones_MouseDoubleClick(tvUniones, Nothing)
                End If
            End If
        Else
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
