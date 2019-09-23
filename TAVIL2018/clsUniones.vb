Imports System.Linq
Imports ce = ClosedXML.Excel
Public Class clsUniones
    Public nHoja As String = HojaUniones
    Public Campos As List(Of String)
    Public filas As List(Of UNIONItem)        ' Key = ITEM_NUMBER, Value = clsPTItem
    Public Sub New()
        If cfg Is Nothing Then cfg = New UtilesAlberto.Conf(System.Reflection.Assembly.GetExecutingAssembly)
        'If cXML Is Nothing Then cXML = New ClosedXML2acad.ClosedXML2acad
        Campos = New List(Of String)
        filas = New List(Of UNIONItem)
        '
        For Each fila As ce.IXLRow In modClosedXMLTavil.Excel_LeeFilar(LAYOUTDB, nHoja, concabeceras:=True).AsParallel
            If fila.RowNumber = 1 Then
                For Each oCe As ce.IXLCell In fila.CellsUsed
                    Campos.Add(Convert.ToString(oCe.Value))
                Next
            Else
                filas.Add(New UNIONItem(fila))
            End If
            System.Windows.Forms.Application.DoEvents()
        Next
    End Sub
    Public Function Fila_BuscaDame(inC As String, inInc As String, outC As String, outInc As String, angle As String) As String()
        Dim resultado() As String = {"", ""}

        Dim fila = From x In filas
                   Where x.INFEED_CONVEYOR.Trim.StartsWith(inC.Trim) AndAlso
                                    x.INFEED_INCLINATION.Trim = inInc.Trim AndAlso
                                    x.OUTFEED_CONVEYOR.Trim.StartsWith(outC.Trim) AndAlso
                                    x.OUTFEED_INCLINATION.Trim = outInc.Trim AndAlso
                                    x.ANGLE.Trim = angle.Trim
        '
        If fila.Count > 0 Then
            resultado(0) = fila.FirstOrDefault.UNION
            resultado(1) = fila.FirstOrDefault.UNITS
        End If
        '
        Return resultado
    End Function
End Class


Public Class UNIONItem
    Public nFila As Integer = -1
    Public INFEED_CONVEYOR As String = ""
    Public INFEED_INCLINATION As String = ""
    Public UNION As String = ""
    Public UNITS As String = ""
    Public OUTFEED_CONVEYOR As String = ""
    Public OUTFEED_INCLINATION As String = ""
    Public ANGLE As String = ""
    Public INFORMACION As String = ""

    Public Sub New(fila As ce.IXLRow)
        'If cXML Is Nothing Then cXML = New ClosedXML2acad.ClosedXML2acad
        nFila = fila.RowNumber
        For Each oCell As ce.IXLCell In fila.Cells.AsParallel
            Dim cabecera As String = oCell.WorksheetColumn.FirstCell.Value.ToString
            Dim valor As String = oCell.Value
            Select Case cabecera
                Case "INFEED_CONVEYOR" : INFEED_CONVEYOR = Convert.ToString(valor)
                Case "INFEED_INCLINATION" : Me.INFEED_INCLINATION = Convert.ToString(valor)
                Case "UNION" : Me.UNION = Convert.ToString(valor)
                Case "UNITS" : Me.UNITS = Convert.ToString(valor)
                Case "OUTFEED_CONVEYOR" : Me.OUTFEED_CONVEYOR = Convert.ToString(valor)
                Case "OUTFEED_INCLINATION" : Me.OUTFEED_INCLINATION = Convert.ToString(valor)
                Case "ANGLE" : Me.ANGLE = Convert.ToString(valor)
                Case "INFORMACION" : Me.INFORMACION = Convert.ToString(valor)
            End Select
        Next
    End Sub
End Class

Public Enum nombreColumnaUNIONES
    INFEED_CONVEYOR
    INFEED_INCLINATION
    UNION
    UNITS
    OUTFEED_CONVEYOR
    OUTFEED_INCLINATION
    ANGLE
    INFORMACION
End Enum
