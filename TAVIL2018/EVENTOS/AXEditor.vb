Imports System
Imports System.Text
Imports System.Linq
Imports System.Xml
Imports System.Reflection
Imports System.ComponentModel
Imports System.Collections
Imports System.Collections.Generic
Imports System.Windows
Imports System.Windows.Media.Imaging
Imports System.Windows.Forms
Imports System.IO

Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.Geometry
Imports AXApp = Autodesk.AutoCAD.ApplicationServices.Application
Imports AXDoc = Autodesk.AutoCAD.ApplicationServices.Document
Imports AXWin = Autodesk.AutoCAD.Windows
Imports a2 = AutoCAD2acad.A2acad
'
Partial Public Class Eventos
    Public Shared Sub Subscribe_AXEditor()
        AddHandler AXEditor().SelectionRemoved, AddressOf AXEditor_SelectionRemoved_Handler
        AddHandler AXEditor().SelectionAdded, AddressOf AXEditor_SelectionAdded_Handler
        AddHandler AXEditor().Rollover, AddressOf AXEditor_Rollover_Handler
        AddHandler AXEditor().PromptingForString, AddressOf AXEditor_PromptingForString_Handler
        AddHandler AXEditor().PromptingForSelection, AddressOf AXEditor_PromptingForSelection_Handler
        AddHandler AXEditor().PromptingForPoint, AddressOf AXEditor_PromptingForPoint_Handler
        AddHandler AXEditor().PromptingForNestedEntity, AddressOf AXEditor_PromptingForNestedEntity_Handler
        AddHandler AXEditor().PromptingForKeyword, AddressOf AXEditor_PromptingForKeyword_Handler
        AddHandler AXEditor().PromptingForInteger, AddressOf AXEditor_PromptingForInteger_Handler
        AddHandler AXEditor().PromptingForEntity, AddressOf AXEditor_PromptingForEntity_Handler
        AddHandler AXEditor().PromptingForDouble, AddressOf AXEditor_PromptingForDouble_Handler
        AddHandler AXEditor().PromptingForDistance, AddressOf AXEditor_PromptingForDistance_Handler
        AddHandler AXEditor().PromptingForCorner, AddressOf AXEditor_PromptingForCorner_Handler
        AddHandler AXEditor().PromptingForAngle, AddressOf AXEditor_PromptingForAngle_Handler
        AddHandler AXEditor().PromptedForString, AddressOf AXEditor_PromptedForString_Handler
        AddHandler AXEditor().PromptedForSelection, AddressOf AXEditor_PromptedForSelection_Handler
        AddHandler AXEditor().PromptedForPoint, AddressOf AXEditor_PromptedForPoint_Handler
        AddHandler AXEditor().PromptedForNestedEntity, AddressOf AXEditor_PromptedForNestedEntity_Handler
        AddHandler AXEditor().PromptedForKeyword, AddressOf AXEditor_PromptedForKeyword_Handler
        AddHandler AXEditor().PromptedForInteger, AddressOf AXEditor_PromptedForInteger_Handler
        AddHandler AXEditor().PromptedForEntity, AddressOf AXEditor_PromptedForEntity_Handler
        AddHandler AXEditor().PromptedForDouble, AddressOf AXEditor_PromptedForDouble_Handler
        AddHandler AXEditor().PromptedForDistance, AddressOf AXEditor_PromptedForDistance_Handler
        AddHandler AXEditor().PromptedForCorner, AddressOf AXEditor_PromptedForCorner_Handler
        AddHandler AXEditor().PromptedForAngle, AddressOf AXEditor_PromptedForAngle_Handler
        AddHandler AXEditor().PromptForSelectionEnding, AddressOf AXEditor_PromptForSelectionEnding_Handler
        AddHandler AXEditor().PromptForEntityEnding, AddressOf AXEditor_PromptForEntityEnding_Handler
        AddHandler AXEditor().PointMonitor, AddressOf AXEditor_PointMonitor_Handler
        AddHandler AXEditor().PointFilter, AddressOf AXEditor_PointFilter_Handler
        AddHandler AXEditor().LeavingQuiescentState, AddressOf AXEditor_LeavingQuiescentState_Handler
        AddHandler AXEditor().EnteringQuiescentState, AddressOf AXEditor_EnteringQuiescentState_Handler
        AddHandler AXEditor().DraggingEnded, AddressOf AXEditor_DraggingEnded_Handler
        AddHandler AXEditor().Dragging, AddressOf AXEditor_Dragging_Handler
    End Sub

    Public Shared Sub Unsubscribe_AXEditor()
        RemoveHandler AXEditor().SelectionRemoved, AddressOf AXEditor_SelectionRemoved_Handler
        RemoveHandler AXEditor().SelectionAdded, AddressOf AXEditor_SelectionAdded_Handler
        RemoveHandler AXEditor().Rollover, AddressOf AXEditor_Rollover_Handler
        RemoveHandler AXEditor().PromptingForString, AddressOf AXEditor_PromptingForString_Handler
        RemoveHandler AXEditor().PromptingForSelection, AddressOf AXEditor_PromptingForSelection_Handler
        RemoveHandler AXEditor().PromptingForPoint, AddressOf AXEditor_PromptingForPoint_Handler
        RemoveHandler AXEditor().PromptingForNestedEntity, AddressOf AXEditor_PromptingForNestedEntity_Handler
        RemoveHandler AXEditor().PromptingForKeyword, AddressOf AXEditor_PromptingForKeyword_Handler
        RemoveHandler AXEditor().PromptingForInteger, AddressOf AXEditor_PromptingForInteger_Handler
        RemoveHandler AXEditor().PromptingForEntity, AddressOf AXEditor_PromptingForEntity_Handler
        RemoveHandler AXEditor().PromptingForDouble, AddressOf AXEditor_PromptingForDouble_Handler
        RemoveHandler AXEditor().PromptingForDistance, AddressOf AXEditor_PromptingForDistance_Handler
        RemoveHandler AXEditor().PromptingForCorner, AddressOf AXEditor_PromptingForCorner_Handler
        RemoveHandler AXEditor().PromptingForAngle, AddressOf AXEditor_PromptingForAngle_Handler
        RemoveHandler AXEditor().PromptedForString, AddressOf AXEditor_PromptedForString_Handler
        RemoveHandler AXEditor().PromptedForSelection, AddressOf AXEditor_PromptedForSelection_Handler
        RemoveHandler AXEditor().PromptedForPoint, AddressOf AXEditor_PromptedForPoint_Handler
        RemoveHandler AXEditor().PromptedForNestedEntity, AddressOf AXEditor_PromptedForNestedEntity_Handler
        RemoveHandler AXEditor().PromptedForKeyword, AddressOf AXEditor_PromptedForKeyword_Handler
        RemoveHandler AXEditor().PromptedForInteger, AddressOf AXEditor_PromptedForInteger_Handler
        RemoveHandler AXEditor().PromptedForEntity, AddressOf AXEditor_PromptedForEntity_Handler
        RemoveHandler AXEditor().PromptedForDouble, AddressOf AXEditor_PromptedForDouble_Handler
        RemoveHandler AXEditor().PromptedForDistance, AddressOf AXEditor_PromptedForDistance_Handler
        RemoveHandler AXEditor().PromptedForCorner, AddressOf AXEditor_PromptedForCorner_Handler
        RemoveHandler AXEditor().PromptedForAngle, AddressOf AXEditor_PromptedForAngle_Handler
        RemoveHandler AXEditor().PromptForSelectionEnding, AddressOf AXEditor_PromptForSelectionEnding_Handler
        RemoveHandler AXEditor().PromptForEntityEnding, AddressOf AXEditor_PromptForEntityEnding_Handler
        RemoveHandler AXEditor().PointMonitor, AddressOf AXEditor_PointMonitor_Handler
        RemoveHandler AXEditor().PointFilter, AddressOf AXEditor_PointFilter_Handler
        RemoveHandler AXEditor().LeavingQuiescentState, AddressOf AXEditor_LeavingQuiescentState_Handler
        RemoveHandler AXEditor().EnteringQuiescentState, AddressOf AXEditor_EnteringQuiescentState_Handler
        RemoveHandler AXEditor().DraggingEnded, AddressOf AXEditor_DraggingEnded_Handler
        RemoveHandler AXEditor().Dragging, AddressOf AXEditor_Dragging_Handler
    End Sub

    Public Shared Sub AXEditor_Dragging_Handler(ByVal sender As Object, ByVal e As DraggingEventArgs)
        'AXEditor().WriteMessage(vbLf & "Dragging" & vbLf)
    End Sub

    Public Shared Sub AXEditor_DraggingEnded_Handler(ByVal sender As Object, ByVal e As DraggingEndedEventArgs)
        'AXEditor().WriteMessage(vbLf & "DraggingEnded" & vbLf)
    End Sub

    Public Shared Sub AXEditor_EnteringQuiescentState_Handler(ByVal sender As Object, ByVal e As EventArgs)
        'AXEditor().WriteMessage(vbLf & "EnteringQuiescentState" & vbLf)
    End Sub

    Public Shared Sub AXEditor_LeavingQuiescentState_Handler(ByVal sender As Object, ByVal e As EventArgs)
        'AXEditor().WriteMessage(vbLf & "LeavingQuiescentState" & vbLf)
    End Sub

    Public Shared Sub AXEditor_PointFilter_Handler(ByVal sender As Object, ByVal e As PointFilterEventArgs)
        'AXEditor().WriteMessage(vbLf & "PointFilter" & vbLf)
    End Sub

    Public Shared Sub AXEditor_PointMonitor_Handler(ByVal sender As Object, ByVal e As PointMonitorEventArgs)
        'AXEditor().WriteMessage(vbLf & "PointMonitor" & vbLf)
    End Sub

    Public Shared Sub AXEditor_PromptForEntityEnding_Handler(ByVal sender As Object, ByVal e As PromptForEntityEndingEventArgs)
        'AXEditor().WriteMessage(vbLf & "PromptForEntityEnding" & vbLf)
    End Sub

    Public Shared Sub AXEditor_PromptForSelectionEnding_Handler(ByVal sender As Object, ByVal e As PromptForSelectionEndingEventArgs)
        'AXEditor().WriteMessage(vbLf & "PromptForSelectionEnding" & vbLf)
    End Sub

    Public Shared Sub AXEditor_PromptedForAngle_Handler(ByVal sender As Object, ByVal e As PromptDoubleResultEventArgs)
        'AXEditor().WriteMessage(vbLf & "PromptedForAngle" & vbLf)
    End Sub

    Public Shared Sub AXEditor_PromptedForCorner_Handler(ByVal sender As Object, ByVal e As PromptPointResultEventArgs)
        'AXEditor().WriteMessage(vbLf & "PromptedForCorner" & vbLf)
    End Sub

    Public Shared Sub AXEditor_PromptedForDistance_Handler(ByVal sender As Object, ByVal e As PromptDoubleResultEventArgs)
        'AXEditor().WriteMessage(vbLf & "PromptedForDistance" & vbLf)
    End Sub

    Public Shared Sub AXEditor_PromptedForDouble_Handler(ByVal sender As Object, ByVal e As PromptDoubleResultEventArgs)
        'AXEditor().WriteMessage(vbLf & "PromptedForDouble" & vbLf)
    End Sub

    Public Shared Sub AXEditor_PromptedForEntity_Handler(ByVal sender As Object, ByVal e As PromptEntityResultEventArgs)
        'AXEditor().WriteMessage(vbLf & "PromptedForEntity" & vbLf)
    End Sub

    Public Shared Sub AXEditor_PromptedForInteger_Handler(ByVal sender As Object, ByVal e As PromptIntegerResultEventArgs)
        'AXEditor().WriteMessage(vbLf & "PromptedForInteger" & vbLf)
    End Sub

    Public Shared Sub AXEditor_PromptedForKeyword_Handler(ByVal sender As Object, ByVal e As PromptStringResultEventArgs)
        'AXEditor().WriteMessage(vbLf & "PromptedForKeyword" & vbLf)
    End Sub

    Public Shared Sub AXEditor_PromptedForNestedEntity_Handler(ByVal sender As Object, ByVal e As PromptNestedEntityResultEventArgs)
        'AXEditor().WriteMessage(vbLf & "PromptedForNestedEntity" & vbLf)
    End Sub

    Public Shared Sub AXEditor_PromptedForPoint_Handler(ByVal sender As Object, ByVal e As PromptPointResultEventArgs)
        'AXEditor().WriteMessage(vbLf & "PromptedForPoint" & vbLf)
    End Sub

    Public Shared Sub AXEditor_PromptedForSelection_Handler(ByVal sender As Object, ByVal e As PromptSelectionResultEventArgs)
        'AXEditor().WriteMessage(vbLf & "PromptedForSelection" & vbLf)
    End Sub

    Public Shared Sub AXEditor_PromptedForString_Handler(ByVal sender As Object, ByVal e As PromptStringResultEventArgs)
        'AXEditor().WriteMessage(vbLf & "PromptedForString" & vbLf)
    End Sub

    Public Shared Sub AXEditor_PromptingForAngle_Handler(ByVal sender As Object, ByVal e As PromptAngleOptionsEventArgs)
        'AXEditor().WriteMessage(vbLf & "PromptingForAngle" & vbLf)
    End Sub

    Public Shared Sub AXEditor_PromptingForCorner_Handler(ByVal sender As Object, ByVal e As PromptPointOptionsEventArgs)
        'AXEditor().WriteMessage(vbLf & "PromptingForCorner" & vbLf)
    End Sub

    Public Shared Sub AXEditor_PromptingForDistance_Handler(ByVal sender As Object, ByVal e As PromptDistanceOptionsEventArgs)
        'AXEditor().WriteMessage(vbLf & "PromptingForDistance" & vbLf)
    End Sub

    Public Shared Sub AXEditor_PromptingForDouble_Handler(ByVal sender As Object, ByVal e As PromptDoubleOptionsEventArgs)
        'AXEditor().WriteMessage(vbLf & "PromptingForDouble" & vbLf)
    End Sub

    Public Shared Sub AXEditor_PromptingForEntity_Handler(ByVal sender As Object, ByVal e As PromptEntityOptionsEventArgs)
        'AXEditor().WriteMessage(vbLf & "PromptingForEntity" & vbLf)
    End Sub

    Public Shared Sub AXEditor_PromptingForInteger_Handler(ByVal sender As Object, ByVal e As PromptIntegerOptionsEventArgs)
        'AXEditor().WriteMessage(vbLf & "PromptingForInteger" & vbLf)
    End Sub

    Public Shared Sub AXEditor_PromptingForKeyword_Handler(ByVal sender As Object, ByVal e As PromptKeywordOptionsEventArgs)
        'AXEditor().WriteMessage(vbLf & "PromptingForKeyword" & vbLf)
    End Sub

    Public Shared Sub AXEditor_PromptingForNestedEntity_Handler(ByVal sender As Object, ByVal e As PromptNestedEntityOptionsEventArgs)
        'AXEditor().WriteMessage(vbLf & "PromptingForNestedEntity" & vbLf)
    End Sub

    Public Shared Sub AXEditor_PromptingForPoint_Handler(ByVal sender As Object, ByVal e As PromptPointOptionsEventArgs)
        'AXEditor().WriteMessage(vbLf & "PromptingForPoint" & vbLf)
    End Sub

    Public Shared Sub AXEditor_PromptingForSelection_Handler(ByVal sender As Object, ByVal e As PromptSelectionOptionsEventArgs)
        'AXEditor().WriteMessage(vbLf & "PromptingForSelection" & vbLf)
    End Sub

    Public Shared Sub AXEditor_PromptingForString_Handler(ByVal sender As Object, ByVal e As PromptStringOptionsEventArgs)
        'AXEditor().WriteMessage(vbLf & "PromptingForString" & vbLf)
    End Sub

    Public Shared Sub AXEditor_Rollover_Handler(ByVal sender As Object, ByVal e As RolloverEventArgs)
        'AXEditor().WriteMessage(vbLf & "Rollover" & vbLf)
    End Sub

    Public Shared Sub AXEditor_SelectionAdded_Handler(ByVal sender As Object, ByVal e As SelectionAddedEventArgs)
        'AXEditor().WriteMessage(vbLf & "SelectionAdded" & vbLf)
    End Sub

    Public Shared Sub AXEditor_SelectionRemoved_Handler(ByVal sender As Object, ByVal e As SelectionRemovedEventArgs)
        'AXEditor().WriteMessage(vbLf & "SelectionRemoved" & vbLf)
    End Sub
End Class

