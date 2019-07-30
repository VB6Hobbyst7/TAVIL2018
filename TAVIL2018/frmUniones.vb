Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports TAVIL2018.TAVIL2018
Imports System.Drawing
Imports System.Windows.Forms
Imports a2 = AutoCAD2acad.A2acad

Public Class frmUniones
    Private HighlightedPictureBox As PictureBox = Nothing
    Private oT1 As AcadBlockReference = Nothing
    Private oT2 As AcadBlockReference = Nothing
    Private oTT As ToolTip = Nothing

    Private Sub frmUniones_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "UNIONES - v" & cfg._appversion
        If clsA Is Nothing Then clsA = New a2.A2acad(oApp, cfg._appFullPath, regAPPCliente)
        oTT = New ToolTip()
        ' Set up the delays for the ToolTip.
        oTT.AutoPopDelay = 5000
        oTT.InitialDelay = 500
        oTT.ReshowDelay = 500
        ' Force the ToolTip text to be displayed whether or not the form is active.
        oTT.ShowAlways = True
        PonToolTipControles()

        pb1.BackColor = Drawing.Color.FromArgb(255, 192, 192)
        pb2.BackColor = Drawing.Color.FromArgb(255, 192, 192)
        'app_procesointerno = True
        'app_procesointerno = True
        'gbAdministrar.Enabled = False
        'btnGrupoBorrar.Enabled = False
        'txtNombreGrupo.Text = ""
        'lblInf.Text = ""
        tvUniones_LlenaXDATA()
    End Sub
    Public Sub PonToolTipControles()
        Me.tvUniones.ShowNodeToolTips = True
        oTT.SetToolTip(Me.tvUniones, "Listado de uniones")
        oTT.SetToolTip(Me.cbZoom, "Zoom uniones")
        oTT.SetToolTip(Me.pb1, "Seleccionar Transportador 1")
        oTT.SetToolTip(Me.pb2, "Seleccionar Transportador 2")
        oTT.SetToolTip(Me.lblInf, "Información selecciones")
        oTT.SetToolTip(Me.btnCerrar, "Cerrar Uniones")
    End Sub
    Private Sub frmUniones_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        frmUn = Nothing
    End Sub

    Private Sub btnAceptar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
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
    Private Sub Pb1_Click(sender As Object, e As EventArgs) Handles pb1.Click
        pb1.BackColor = Control.DefaultBackColor : pb1.Refresh()
        Control_Borde(pb1, False)
        '
        Me.Visible = False
        oApp.ActiveDocument.Activate()
        Dim arrEntities As ArrayList = Nothing
REPITE:
        arrEntities = clsA.SeleccionaDameEntitiesONSCREEN(solouna:=True)
        If arrEntities Is Nothing OrElse arrEntities.Count = 0 Then
            Exit Sub
        End If
        '
        If TypeOf arrEntities(0) Is AcadBlockReference Then
            oT1 = arrEntities(0)
        End If
        '
        PonDatos()
        Me.Visible = True
    End Sub
    Private Sub Pb2_Click(sender As Object, e As EventArgs) Handles pb2.Click
        pb2.BackColor = Control.DefaultBackColor : pb2.Refresh()
        Control_Borde(pb2, False)
        '
        Me.Visible = False
        oApp.ActiveDocument.Activate()
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
        PonDatos()
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
    Private Sub PonDatos()
        If oT1 IsNot Nothing Then
            lblInf.Text = "T1 = " & oT1.EffectiveName & vbCrLf
            Control_Borde(pb1, True)
        Else
            lblInf.Text = "T1 = " & vbCrLf
        End If
        '
        If oT2 IsNot Nothing Then
            lblInf.Text &= "T2 = " & oT2.EffectiveName & vbCrLf
            Control_Borde(pb2, True)
        Else
            lblInf.Text &= "T2 = "
        End If
    End Sub
    Public Sub tvUniones_LlenaXDATA()
        ' Rellenar tvGrupos con los grupos que haya ([nombre grupo]) Sacado de XData elementos (regAPPCliente, XData = "GRUPO")
        tvUniones.Nodes.Clear()
        If clsA Is Nothing Then clsA = New a2.A2acad(oApp, cfg._appFullPath, regAPPCliente)
        Dim arrTodos As List(Of Long) = clsA.SeleccionaTodosObjetos(,, True)
        If arrTodos Is Nothing OrElse arrTodos.Count = 0 Then
            Exit Sub
        End If
        ' Filtrar lista de grupo. Sacar nombres únicos.
        For Each queId As Long In arrTodos
            Dim acadObj As AcadObject = oApp.ActiveDocument.ObjectIdToObject(queId)
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
                arrSeleccion.Add(oApp.ActiveDocument.ObjectIdToObject(queId))
            Next
            If arrSeleccion.Count > 0 Then
                lblInf.Text = arrSeleccion.Count & " Elementos"
                clsA.SeleccionCreaResalta(arrSeleccion, 0, False)
                If cbZoom.Checked Then
                    Zoom_Seleccion()
                End If
            End If
        Else
            'If tvGrupos.Nodes.ContainsKey(grupo) Then
            '    tvGrupos.Nodes.Item(grupo).Remove()
            'End If
            lblInf.Text = "0 Elementos"
            'tvGrupos.SelectedNode = Nothing
            'tvGrupos_AfterSelect(Nothing, Nothing)
        End If
    End Sub

    ' PictureBox (U otros controles). Poner el borde de color rojo  
    Private Sub Control_Borde(oC As Control, Optional restaurar As Boolean = False)
        CType(oC, PictureBox).BorderStyle = BorderStyle.None
        If restaurar Then
            oC.Invalidate()
        Else
            'Rectangulo del control + Offset del rectangulo hacia fuera.  
            Dim BorderBounds As Rectangle = oC.ClientRectangle
            'BorderBounds.Inflate(-1, -1)

            'Use ControlPaint to draw the border.  
            'Change the Color.Red parameter to your own colour below.  
            ControlPaint.DrawBorder(oC.CreateGraphics,
                                    BorderBounds,
                                    Color.Red,
                                    ButtonBorderStyle.Outset)
        End If
    End Sub

#End Region
End Class
