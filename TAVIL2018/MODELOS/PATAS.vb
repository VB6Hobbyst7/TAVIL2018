Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports TAVIL2018.TAVIL2018
Imports System.Windows.Forms
Imports System.Linq
Imports a2 = AutoCAD2acad.A2acad
Public Class PATAS
    Public Shared DPatas As Dictionary(Of String, PATA)
End Class

Public Class PATA
    Public Property HANDLE As String
    Public Property CODE As String
    Public Property BLOCK As String
    Public Property HEIGHT As String
    Public Property WIDTH As String
    Public Property RADIUS As String
    Public Property KEY As String
    '
    Public Sub New(h As String) 'Handle
        If clsA Is Nothing Then clsA = New a2.A2acad(Eventos.COMApp(), cfg._appFullPath, regAPPCliente)
        Dim acadObj As AcadObject = Eventos.COMDoc.HandleToObject(h)
        If acadObj IsNot Nothing AndAlso TypeOf acadObj Is AcadBlockReference Then
            Dim atributos = CType(acadObj, AcadBlockReference).GetAttributes()
            Dim parametros As Object = Nothing
            If CType(acadObj, AcadBlockReference).IsDynamicBlock Then
                parametros = CType(acadObj, AcadBlockReference).GetDynamicBlockProperties
            End If
            Me.HANDLE = h
            Me.CODE = clsA.Bloque_DameDato_AttPropX(h, atributos, parametros, "CODE")
            Me.BLOCK = clsA.Bloque_DameDato_AttPropX(h, atributos, parametros, "BLOCK")
            Me.HEIGHT = clsA.Bloque_DameDato_AttPropX(h, atributos, parametros, "HEIGHT")
            Me.WIDTH = clsA.Bloque_DameDato_AttPropX(h, atributos, parametros, "WIDTH")
            Me.RADIUS = clsA.Bloque_DameDato_AttPropX(h, atributos, parametros, "RADIUS")
            KEY = CODE & BLOCK & HEIGHT & WIDTH & RADIUS
            atributos = Nothing
            parametros = Nothing
            acadObj = Nothing
        End If
    End Sub
End Class
