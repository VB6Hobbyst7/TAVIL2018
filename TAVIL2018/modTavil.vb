Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports TAVIL2018.TAVIL2018
Imports System.Windows.Forms
Imports uau = UtilesAlberto.Utiles
Imports a2 = AutoCAD2acad.A2acad
Imports Ev = TAVIL2018.Eventos

Public Module modTavil
    Public LUniones As List(Of ClsUnion)
    Public Function tvUniones_LlenaXDATA(Optional queFiltro As EFiltro = EFiltro.TODOS) As List(Of TreeNode)
        Dim lUni As New List(Of TreeNode)
        LUniones = New List(Of ClsUnion)
        '
        If clsA Is Nothing Then clsA = New a2.A2acad(Eventos.COMApp(), cfg._appFullPath, regAPPCliente)
        Dim arrTodos As List(Of Long) = clsA.SeleccionaDameBloquesPorNombre(cUNION)
        If arrTodos Is Nothing OrElse arrTodos.Count = 0 Then
            Return lUni
            Exit Function
        End If
        '
        ' Filtrar lista de grupo. Sacar nombres únicos.
        For Each queId As Long In arrTodos
            Dim acadObj As AcadObject = Eventos.COMDoc().ObjectIdToObject(queId)
            If TypeOf acadObj Is AcadBlockReference = False Then Continue For
            Dim oBl As AcadBlockReference = CType(acadObj, AcadBlockReference)
            If oBl.EffectiveName <> cUNION Then Continue For
            '
            Dim union As String = clsA.XLeeDato(oBl, "UNION")
            Dim oNode As New TreeNode
            oNode.Text = queId.ToString
            oNode.Name = queId.ToString
            oNode.Tag = queId
            oNode.ToolTipText = "Unión = " & oBl.EffectiveName & "·" & queId
            oNode = Nothing

            '
            If queFiltro = "TODOS" Then
            ElseIf queFiltro = "XXX" Then
                Dim t1 As String = clsA.XLeeDato(oBl, "T1ID")
                Dim t2 As String = clsA.XLeeDato(oBl, "T2ID")
                If t1 <> "" AndAlso t2 <> "" Then
                    Continue For
                Else
                    lUni.Add(oNode)
                End If
                Dim Uni As String = clsA.XLeeDato(oBl, "UNION")
                If Uni = "" OrElse Uni = "XXX" Then
                    lUni.Add(oNode)
                    LUniones.Add(New ClsUnion(oBl.ObjectID))
                Else
                    Continue For
                End If
            End If
            oBl = Nothing
            oNode = Nothing
        Next
        '
        Return lUni
    End Function

    Public Enum EFiltro
        TODOS
        XXX
    End Enum
    Public Enum INCLINATION
        FLAT
        DOWN
        UP
    End Enum
End Module