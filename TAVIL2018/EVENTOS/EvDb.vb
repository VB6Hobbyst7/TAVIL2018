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
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Windows
Imports AcadApplication = Autodesk.AutoCAD.ApplicationServices.Application
Imports AcadDocument = Autodesk.AutoCAD.ApplicationServices.Document
Imports AcadWindows = Autodesk.AutoCAD.Windows
Imports a2 = AutoCAD2acad.A2acad
'
Namespace Eventos
    Partial Public Class AutoCADEventos

        Public Sub Database_ObjectModified(ByVal sender As Object, ByVal e As Autodesk.AutoCAD.DatabaseServices.ObjectEventArgs)
            'AutoEnumera_DBObjectModified(sender, e)
            ''oApp.ActiveDocument.Regen(AcRegenType.acActiveViewport)
            'AddHandler Autodesk.AutoCAD.ApplicationServices.Application.Idle, AddressOf Tavil_AppIdle
        End Sub

        Public Sub Database_ObjectAppended(ByVal sender As Object, ByVal e As Autodesk.AutoCAD.DatabaseServices.ObjectEventArgs)
            'Try
            '    If TypeOf e.DBObject.AcadObject Is AcadBlockReference AndAlso CType(e.DBObject.AcadObject, AcadBlockReference).EffectiveName.ToUpper.StartsWith("PT") Then
            '        Exit Sub
            '    End If
            'Catch ex As Exception
            '    Exit Sub
            'End Try
            'AutoEnumera_DBObjectAppended(sender, e)
            ''oApp.ActiveDocument.Regen(AcRegenType.acActiveViewport)
            'AddHandler Autodesk.AutoCAD.ApplicationServices.Application.Idle, AddressOf Tavil_AppIdle
        End Sub

        Public Sub Database_ObjectErased(ByVal sender As Object, ByVal e As Autodesk.AutoCAD.DatabaseServices.ObjectErasedEventArgs)
            'AutoEnumera_ObjectErased(sender, e)
            'AddHandler Autodesk.AutoCAD.ApplicationServices.Application.Idle, AddressOf Tavil_AppIdle
        End Sub
        '

        Public Shared Sub DbEvent_AbortDxfIn_Handler(ByVal sender As Object, ByVal e As EventArgs)
            AcadApplication.DocumentManager.GetDocument(TryCast(sender, Database)).Editor.WriteMessage("AbortDxfIn")
        End Sub

        Public Shared Sub DbEvent_AbortDxfOut_Handler(ByVal sender As Object, ByVal e As EventArgs)
            AcadApplication.DocumentManager.GetDocument(TryCast(sender, Database)).Editor.WriteMessage("AbortDxfOut")
        End Sub

        Public Shared Sub DbEvent_AbortSave_Handler(ByVal sender As Object, ByVal e As EventArgs)
            AcadApplication.DocumentManager.GetDocument(TryCast(sender, Database)).Editor.WriteMessage("AbortSave")
        End Sub

        Public Shared Sub DbEvent_BeginDeepClone_Handler(ByVal sender As Object, ByVal e As IdMappingEventArgs)
            AcadApplication.DocumentManager.GetDocument(TryCast(sender, Database)).Editor.WriteMessage("BeginDeepClone")
        End Sub

        Public Shared Sub DbEvent_BeginDeepCloneTranslation_Handler(ByVal sender As Object, ByVal e As IdMappingEventArgs)
            AcadApplication.DocumentManager.GetDocument(TryCast(sender, Database)).Editor.WriteMessage("BeginDeepCloneTranslation")
        End Sub

        Public Shared Sub DbEvent_BeginDxfIn_Handler(ByVal sender As Object, ByVal e As EventArgs)
            AcadApplication.DocumentManager.GetDocument(TryCast(sender, Database)).Editor.WriteMessage("BeginDxfIn")
        End Sub

        Public Shared Sub DbEvent_BeginDxfOut_Handler(ByVal sender As Object, ByVal e As EventArgs)
            AcadApplication.DocumentManager.GetDocument(TryCast(sender, Database)).Editor.WriteMessage("BeginDxfOut")
        End Sub

        Public Shared Sub DbEvent_BeginInsert_Handler(ByVal sender As Object, ByVal e As BeginInsertEventArgs)
            AcadApplication.DocumentManager.GetDocument(TryCast(sender, Database)).Editor.WriteMessage("BeginInsert")
        End Sub

        Public Shared Sub DbEvent_BeginSave_Handler(ByVal sender As Object, ByVal e As DatabaseIOEventArgs)
            AcadApplication.DocumentManager.GetDocument(TryCast(sender, Database)).Editor.WriteMessage("BeginSave")
        End Sub

        Public Shared Sub DbEvent_BeginWblockBlock_Handler(ByVal sender As Object, ByVal e As BeginWblockBlockEventArgs)
            AcadApplication.DocumentManager.GetDocument(TryCast(sender, Database)).Editor.WriteMessage("BeginWblockBlock")
        End Sub

        Public Shared Sub DbEvent_BeginWblockEntireDatabase_Handler(ByVal sender As Object, ByVal e As BeginWblockEntireDatabaseEventArgs)
            AcadApplication.DocumentManager.GetDocument(TryCast(sender, Database)).Editor.WriteMessage("BeginWblockEntireDatabase")
        End Sub

        Public Shared Sub DbEvent_BeginWblockObjects_Handler(ByVal sender As Object, ByVal e As BeginWblockObjectsEventArgs)
            AcadApplication.DocumentManager.GetDocument(TryCast(sender, Database)).Editor.WriteMessage("BeginWblockObjects")
        End Sub

        Public Shared Sub DbEvent_BeginWblockSelectedObjects_Handler(ByVal sender As Object, ByVal e As BeginWblockSelectedObjectsEventArgs)
            AcadApplication.DocumentManager.GetDocument(TryCast(sender, Database)).Editor.WriteMessage("BeginWblockSelectedObjects")
        End Sub

        Public Shared Sub DbEvent_DatabaseConstructed_Handler(ByVal sender As Object, ByVal e As EventArgs)
            AcadApplication.DocumentManager.GetDocument(TryCast(sender, Database)).Editor.WriteMessage("DatabaseConstructed")
        End Sub

        Public Shared Sub DbEvent_DatabaseToBeDestroyed_Handler(ByVal sender As Object, ByVal e As EventArgs)
            AcadApplication.DocumentManager.GetDocument(TryCast(sender, Database)).Editor.WriteMessage("DatabaseToBeDestroyed")
        End Sub

        Public Shared Sub DbEvent_DeepCloneAborted_Handler(ByVal sender As Object, ByVal e As EventArgs)
            AcadApplication.DocumentManager.GetDocument(TryCast(sender, Database)).Editor.WriteMessage("DeepCloneAborted")
        End Sub

        Public Shared Sub DbEvent_DeepCloneEnded_Handler(ByVal sender As Object, ByVal e As EventArgs)
            AcadApplication.DocumentManager.GetDocument(TryCast(sender, Database)).Editor.WriteMessage("DeepCloneEnded")
        End Sub

        Public Shared Sub DbEvent_Disposed_Handler(ByVal sender As Object, ByVal e As EventArgs)
            AcadApplication.DocumentManager.GetDocument(TryCast(sender, Database)).Editor.WriteMessage("Disposed")
        End Sub

        Public Shared Sub DbEvent_DwgFileOpened_Handler(ByVal sender As Object, ByVal e As DatabaseIOEventArgs)
            AcadApplication.DocumentManager.GetDocument(TryCast(sender, Database)).Editor.WriteMessage("DwgFileOpened")
        End Sub

        Public Shared Sub DbEvent_DxfInComplete_Handler(ByVal sender As Object, ByVal e As EventArgs)
            AcadApplication.DocumentManager.GetDocument(TryCast(sender, Database)).Editor.WriteMessage("DxfInComplete")
        End Sub

        Public Shared Sub DbEvent_DxfOutComplete_Handler(ByVal sender As Object, ByVal e As EventArgs)
            AcadApplication.DocumentManager.GetDocument(TryCast(sender, Database)).Editor.WriteMessage("DxfOutComplete")
        End Sub

        Public Shared Sub DbEvent_InitialDwgFileOpenComplete_Handler(ByVal sender As Object, ByVal e As EventArgs)
            AcadApplication.DocumentManager.GetDocument(TryCast(sender, Database)).Editor.WriteMessage("InitialDwgFileOpenComplete")
        End Sub

        Public Shared Sub DbEvent_InsertAborted_Handler(ByVal sender As Object, ByVal e As EventArgs)
            AcadApplication.DocumentManager.GetDocument(TryCast(sender, Database)).Editor.WriteMessage("InsertAborted")
        End Sub

        Public Shared Sub DbEvent_InsertEnded_Handler(ByVal sender As Object, ByVal e As EventArgs)
            AcadApplication.DocumentManager.GetDocument(TryCast(sender, Database)).Editor.WriteMessage("InsertEnded")
        End Sub

        Public Shared Sub DbEvent_InsertMappingAvailable_Handler(ByVal sender As Object, ByVal e As IdMappingEventArgs)
            AcadApplication.DocumentManager.GetDocument(TryCast(sender, Database)).Editor.WriteMessage("InsertMappingAvailable")
        End Sub

        Public Shared Sub DbEvent_ObjectAppended_Handler(ByVal sender As Object, ByVal e As ObjectEventArgs)
            AcadApplication.DocumentManager.GetDocument(TryCast(sender, Database)).Editor.WriteMessage("ObjectAppended")
        End Sub

        Public Shared Sub DbEvent_ObjectErased_Handler(ByVal sender As Object, ByVal e As ObjectErasedEventArgs)
            AcadApplication.DocumentManager.GetDocument(TryCast(sender, Database)).Editor.WriteMessage("ObjectErased")
        End Sub

        Public Shared Sub DbEvent_ObjectModified_Handler(ByVal sender As Object, ByVal e As ObjectEventArgs)
            AcadApplication.DocumentManager.GetDocument(TryCast(sender, Database)).Editor.WriteMessage("ObjectModified")
        End Sub

        Public Shared Sub DbEvent_ObjectOpenedForModify_Handler(ByVal sender As Object, ByVal e As ObjectEventArgs)
            AcadApplication.DocumentManager.GetDocument(TryCast(sender, Database)).Editor.WriteMessage("ObjectOpenedForModify")
        End Sub

        Public Shared Sub DbEvent_ObjectReappended_Handler(ByVal sender As Object, ByVal e As ObjectEventArgs)
            AcadApplication.DocumentManager.GetDocument(TryCast(sender, Database)).Editor.WriteMessage("ObjectReappended")
        End Sub

        Public Shared Sub DbEvent_ObjectUnappended_Handler(ByVal sender As Object, ByVal e As ObjectEventArgs)
            AcadApplication.DocumentManager.GetDocument(TryCast(sender, Database)).Editor.WriteMessage("ObjectUnappended")
        End Sub

        Public Shared Sub DbEvent_PartialOpenNotice_Handler(ByVal sender As Object, ByVal e As EventArgs)
            AcadApplication.DocumentManager.GetDocument(TryCast(sender, Database)).Editor.WriteMessage("PartialOpenNotice")
        End Sub

        Public Shared Sub DbEvent_ProxyResurrectionCompleted_Handler(ByVal sender As Object, ByVal e As ProxyResurrectionCompletedEventArgs)
            AcadApplication.DocumentManager.GetDocument(TryCast(sender, Database)).Editor.WriteMessage("ProxyResurrectionCompleted")
        End Sub

        Public Shared Sub DbEvent_SaveComplete_Handler(ByVal sender As Object, ByVal e As DatabaseIOEventArgs)
            AcadApplication.DocumentManager.GetDocument(TryCast(sender, Database)).Editor.WriteMessage("SaveComplete")
        End Sub

        Public Shared Sub DbEvent_SystemVariableChanged_Handler(ByVal sender As Object, ByVal e As Autodesk.AutoCAD.DatabaseServices.SystemVariableChangedEventArgs)
            AcadApplication.DocumentManager.GetDocument(TryCast(sender, Database)).Editor.WriteMessage("SystemVariableChanged")
        End Sub

        Public Shared Sub DbEvent_SystemVariableWillChange_Handler(ByVal sender As Object, ByVal e As Autodesk.AutoCAD.DatabaseServices.SystemVariableChangingEventArgs)
            AcadApplication.DocumentManager.GetDocument(TryCast(sender, Database)).Editor.WriteMessage("SystemVariableWillChange")
        End Sub

        Public Shared Sub DbEvent_WblockAborted_Handler(ByVal sender As Object, ByVal e As EventArgs)
            AcadApplication.DocumentManager.GetDocument(TryCast(sender, Database)).Editor.WriteMessage("WblockAborted")
        End Sub

        Public Shared Sub DbEvent_WblockEnded_Handler(ByVal sender As Object, ByVal e As EventArgs)
            AcadApplication.DocumentManager.GetDocument(TryCast(sender, Database)).Editor.WriteMessage("WblockEnded")
        End Sub

        Public Shared Sub DbEvent_WblockMappingAvailable_Handler(ByVal sender As Object, ByVal e As IdMappingEventArgs)
            AcadApplication.DocumentManager.GetDocument(TryCast(sender, Database)).Editor.WriteMessage("WblockMappingAvailable")
        End Sub

        Public Shared Sub DbEvent_WblockNotice_Handler(ByVal sender As Object, ByVal e As WblockNoticeEventArgs)
            AcadApplication.DocumentManager.GetDocument(TryCast(sender, Database)).Editor.WriteMessage("WblockNotice")
        End Sub

        Public Shared Sub DbEvent_XrefAttachAborted_Handler(ByVal sender As Object, ByVal e As EventArgs)
            AcadApplication.DocumentManager.GetDocument(TryCast(sender, Database)).Editor.WriteMessage("XrefAttachAborted")
        End Sub

        Public Shared Sub DbEvent_XrefAttachEnded_Handler(ByVal sender As Object, ByVal e As EventArgs)
            AcadApplication.DocumentManager.GetDocument(TryCast(sender, Database)).Editor.WriteMessage("XrefAttachEnded")
        End Sub

        Public Shared Sub DbEvent_XrefBeginAttached_Handler(ByVal sender As Object, ByVal e As XrefBeginOperationEventArgs)
            AcadApplication.DocumentManager.GetDocument(TryCast(sender, Database)).Editor.WriteMessage("XrefBeginAttached")
        End Sub

        Public Shared Sub DbEvent_XrefBeginOtherAttached_Handler(ByVal sender As Object, ByVal e As XrefBeginOperationEventArgs)
            AcadApplication.DocumentManager.GetDocument(TryCast(sender, Database)).Editor.WriteMessage("XrefBeginOtherAttached")
        End Sub

        Public Shared Sub DbEvent_XrefBeginRestore_Handler(ByVal sender As Object, ByVal e As XrefBeginOperationEventArgs)
            AcadApplication.DocumentManager.GetDocument(TryCast(sender, Database)).Editor.WriteMessage("XrefBeginRestore")
        End Sub

        Public Shared Sub DbEvent_XrefComandeered_Handler(ByVal sender As Object, ByVal e As XrefComandeeredEventArgs)
            AcadApplication.DocumentManager.GetDocument(TryCast(sender, Database)).Editor.WriteMessage("XrefComandeered")
        End Sub

        Public Shared Sub DbEvent_XrefPreXrefLockFile_Handler(ByVal sender As Object, ByVal e As XrefPreXrefLockFileEventArgs)
            AcadApplication.DocumentManager.GetDocument(TryCast(sender, Database)).Editor.WriteMessage("XrefPreXrefLockFile")
        End Sub

        Public Shared Sub DbEvent_XrefRedirected_Handler(ByVal sender As Object, ByVal e As XrefRedirectedEventArgs)
            AcadApplication.DocumentManager.GetDocument(TryCast(sender, Database)).Editor.WriteMessage("XrefRedirected")
        End Sub

        Public Shared Sub DbEvent_XrefRestoreAborted_Handler(ByVal sender As Object, ByVal e As EventArgs)
            AcadApplication.DocumentManager.GetDocument(TryCast(sender, Database)).Editor.WriteMessage("XrefRestoreAborted")
        End Sub

        Public Shared Sub DbEvent_XrefRestoreEnded_Handler(ByVal sender As Object, ByVal e As EventArgs)
            AcadApplication.DocumentManager.GetDocument(TryCast(sender, Database)).Editor.WriteMessage("XrefRestoreEnded")
        End Sub

        Public Shared Sub DbEvent_XrefSubCommandAborted_Handler(ByVal sender As Object, ByVal e As XrefSubCommandEventArgs)
            AcadApplication.DocumentManager.GetDocument(TryCast(sender, Database)).Editor.WriteMessage("XrefSubCommandAborted")
        End Sub

        Public Shared Sub DbEvent_XrefSubCommandStart_Handler(ByVal sender As Object, ByVal e As XrefVetoableSubCommandEventArgs)
            AcadApplication.DocumentManager.GetDocument(TryCast(sender, Database)).Editor.WriteMessage("XrefSubCommandStart")
        End Sub

        Public Sub New()
            SubscribeForApplication()
            AcadApplication.DocumentManager.DocumentCreated += New DocumentCollectionEventHandler(AddressOf DocumentManager_DocumentCreated)
        End Sub

        Public Shared Sub SubscribeForApplication()
            Dim enumerator As IEnumerator = AcadApplication.DocumentManager.GetEnumerator()

            While enumerator.MoveNext()
                SubscribeToDb((TryCast(enumerator.Current, Document)).Database)
            End While
        End Sub

        Public Shared Sub UnsubscribeForApplication()
            Dim enumerator As IEnumerator = AcadApplication.DocumentManager.GetEnumerator()

            While enumerator.MoveNext()
                SubscribeToDb((TryCast(enumerator.Current, Document)).Database)
            End While
        End Sub

        Public Shared Sub SubscribeToDb(ByVal db As Database)
            db.AbortDxfIn += AddressOf DbEvent_AbortDxfIn_Handler
            db.AbortDxfOut += AddressOf DbEvent_AbortDxfOut_Handler
            db.AbortSave += AddressOf DbEvent_AbortSave_Handler
            db.BeginDeepClone += AddressOf DbEvent_BeginDeepClone_Handler
            db.BeginDeepCloneTranslation += AddressOf DbEvent_BeginDeepCloneTranslation_Handler
            db.BeginDxfIn += AddressOf DbEvent_BeginDxfIn_Handler
            db.BeginDxfOut += AddressOf DbEvent_BeginDxfOut_Handler
            db.BeginInsert += AddressOf DbEvent_BeginInsert_Handler
            db.BeginSave += AddressOf DbEvent_BeginSave_Handler
            db.BeginWblockBlock += AddressOf DbEvent_BeginWblockBlock_Handler
            db.BeginWblockEntireDatabase += AddressOf DbEvent_BeginWblockEntireDatabase_Handler
            db.BeginWblockObjects += AddressOf DbEvent_BeginWblockObjects_Handler
            db.BeginWblockSelectedObjects += AddressOf DbEvent_BeginWblockSelectedObjects_Handler
            Database.DatabaseConstructed += AddressOf DbEvent_DatabaseConstructed_Handler
            db.DatabaseToBeDestroyed += AddressOf DbEvent_DatabaseToBeDestroyed_Handler
            db.DeepCloneAborted += AddressOf DbEvent_DeepCloneAborted_Handler
            db.DeepCloneEnded += AddressOf DbEvent_DeepCloneEnded_Handler
            db.Disposed += AddressOf DbEvent_Disposed_Handler
            db.DwgFileOpened += AddressOf DbEvent_DwgFileOpened_Handler
            db.DxfInComplete += AddressOf DbEvent_DxfInComplete_Handler
            db.DxfOutComplete += AddressOf DbEvent_DxfOutComplete_Handler
            db.InitialDwgFileOpenComplete += AddressOf DbEvent_InitialDwgFileOpenComplete_Handler
            db.InsertAborted += AddressOf DbEvent_InsertAborted_Handler
            db.InsertEnded += AddressOf DbEvent_InsertEnded_Handler
            db.InsertMappingAvailable += AddressOf DbEvent_InsertMappingAvailable_Handler
            db.ObjectAppended += AddressOf DbEvent_ObjectAppended_Handler
            db.ObjectErased += AddressOf DbEvent_ObjectErased_Handler
            db.ObjectModified += AddressOf DbEvent_ObjectModified_Handler
            db.ObjectOpenedForModify += AddressOf DbEvent_ObjectOpenedForModify_Handler
            db.ObjectReappended += AddressOf DbEvent_ObjectReappended_Handler
            db.ObjectUnappended += AddressOf DbEvent_ObjectUnappended_Handler
            db.PartialOpenNotice += AddressOf DbEvent_PartialOpenNotice_Handler
            db.ProxyResurrectionCompleted += AddressOf DbEvent_ProxyResurrectionCompleted_Handler
            db.SaveComplete += AddressOf DbEvent_SaveComplete_Handler
            db.SystemVariableChanged += AddressOf DbEvent_SystemVariableChanged_Handler
            db.SystemVariableWillChange += AddressOf DbEvent_SystemVariableWillChange_Handler
            db.WblockAborted += AddressOf DbEvent_WblockAborted_Handler
            db.WblockEnded += AddressOf DbEvent_WblockEnded_Handler
            db.WblockMappingAvailable += AddressOf DbEvent_WblockMappingAvailable_Handler
            db.WblockNotice += AddressOf DbEvent_WblockNotice_Handler
            Database.XrefAttachAborted += AddressOf DbEvent_XrefAttachAborted_Handler
            db.XrefAttachEnded += AddressOf DbEvent_XrefAttachEnded_Handler
            db.XrefBeginAttached += AddressOf DbEvent_XrefBeginAttached_Handler
            db.XrefBeginOtherAttached += AddressOf DbEvent_XrefBeginOtherAttached_Handler
            db.XrefBeginRestore += AddressOf DbEvent_XrefBeginRestore_Handler
            db.XrefComandeered += AddressOf DbEvent_XrefComandeered_Handler
            db.XrefPreXrefLockFile += AddressOf DbEvent_XrefPreXrefLockFile_Handler
            db.XrefRedirected += AddressOf DbEvent_XrefRedirected_Handler
            db.XrefRestoreAborted += AddressOf DbEvent_XrefRestoreAborted_Handler
            db.XrefRestoreEnded += AddressOf DbEvent_XrefRestoreEnded_Handler
            db.XrefSubCommandAborted += AddressOf DbEvent_XrefSubCommandAborted_Handler
            db.XrefSubCommandStart += AddressOf DbEvent_XrefSubCommandStart_Handler
        End Sub

        Public Shared Sub UnsubscribeFromDb(ByVal db As Database)
            db.AbortDxfIn -= AddressOf DbEvent_AbortDxfIn_Handler
            db.AbortDxfOut -= AddressOf DbEvent_AbortDxfOut_Handler
            db.AbortSave -= AddressOf DbEvent_AbortSave_Handler
            db.BeginDeepClone -= AddressOf DbEvent_BeginDeepClone_Handler
            db.BeginDeepCloneTranslation -= AddressOf DbEvent_BeginDeepCloneTranslation_Handler
            db.BeginDxfIn -= AddressOf DbEvent_BeginDxfIn_Handler
            db.BeginDxfOut -= AddressOf DbEvent_BeginDxfOut_Handler
            db.BeginInsert -= AddressOf DbEvent_BeginInsert_Handler
            db.BeginSave -= AddressOf DbEvent_BeginSave_Handler
            db.BeginWblockBlock -= AddressOf DbEvent_BeginWblockBlock_Handler
            db.BeginWblockEntireDatabase -= AddressOf DbEvent_BeginWblockEntireDatabase_Handler
            db.BeginWblockObjects -= AddressOf DbEvent_BeginWblockObjects_Handler
            db.BeginWblockSelectedObjects -= AddressOf DbEvent_BeginWblockSelectedObjects_Handler
            Database.DatabaseConstructed -= AddressOf DbEvent_DatabaseConstructed_Handler
            db.DatabaseToBeDestroyed -= AddressOf DbEvent_DatabaseToBeDestroyed_Handler
            db.DeepCloneAborted -= AddressOf DbEvent_DeepCloneAborted_Handler
            db.DeepCloneEnded -= AddressOf DbEvent_DeepCloneEnded_Handler
            db.Disposed -= AddressOf DbEvent_Disposed_Handler
            db.DwgFileOpened -= AddressOf DbEvent_DwgFileOpened_Handler
            db.DxfInComplete -= AddressOf DbEvent_DxfInComplete_Handler
            db.DxfOutComplete -= AddressOf DbEvent_DxfOutComplete_Handler
            db.InitialDwgFileOpenComplete -= AddressOf DbEvent_InitialDwgFileOpenComplete_Handler
            db.InsertAborted -= AddressOf DbEvent_InsertAborted_Handler
            db.InsertEnded -= AddressOf DbEvent_InsertEnded_Handler
            db.InsertMappingAvailable -= AddressOf DbEvent_InsertMappingAvailable_Handler
            db.ObjectAppended -= AddressOf DbEvent_ObjectAppended_Handler
            db.ObjectErased -= AddressOf DbEvent_ObjectErased_Handler
            db.ObjectModified -= AddressOf DbEvent_ObjectModified_Handler
            db.ObjectOpenedForModify -= AddressOf DbEvent_ObjectOpenedForModify_Handler
            db.ObjectReappended -= AddressOf DbEvent_ObjectReappended_Handler
            db.ObjectUnappended -= AddressOf DbEvent_ObjectUnappended_Handler
            db.PartialOpenNotice -= AddressOf DbEvent_PartialOpenNotice_Handler
            db.ProxyResurrectionCompleted -= AddressOf DbEvent_ProxyResurrectionCompleted_Handler
            db.SaveComplete -= AddressOf DbEvent_SaveComplete_Handler
            db.SystemVariableChanged -= AddressOf DbEvent_SystemVariableChanged_Handler
            db.SystemVariableWillChange -= AddressOf DbEvent_SystemVariableWillChange_Handler
            db.WblockAborted -= AddressOf DbEvent_WblockAborted_Handler
            db.WblockEnded -= AddressOf DbEvent_WblockEnded_Handler
            db.WblockMappingAvailable -= AddressOf DbEvent_WblockMappingAvailable_Handler
            db.WblockNotice -= AddressOf DbEvent_WblockNotice_Handler
            Database.XrefAttachAborted -= AddressOf DbEvent_XrefAttachAborted_Handler
            db.XrefAttachEnded -= AddressOf DbEvent_XrefAttachEnded_Handler
            db.XrefBeginAttached -= AddressOf DbEvent_XrefBeginAttached_Handler
            db.XrefBeginOtherAttached -= AddressOf DbEvent_XrefBeginOtherAttached_Handler
            db.XrefBeginRestore -= AddressOf DbEvent_XrefBeginRestore_Handler
            db.XrefComandeered -= AddressOf DbEvent_XrefComandeered_Handler
            db.XrefPreXrefLockFile -= AddressOf DbEvent_XrefPreXrefLockFile_Handler
            db.XrefRedirected -= AddressOf DbEvent_XrefRedirected_Handler
            db.XrefRestoreAborted -= AddressOf DbEvent_XrefRestoreAborted_Handler
            db.XrefRestoreEnded -= AddressOf DbEvent_XrefRestoreEnded_Handler
            db.XrefSubCommandAborted -= AddressOf DbEvent_XrefSubCommandAborted_Handler
            db.XrefSubCommandStart -= AddressOf DbEvent_XrefSubCommandStart_Handler
        End Sub

        Private Shared Sub DocumentManager_DocumentCreated(ByVal sender As Object, ByVal e As DocumentCollectionEventArgs)
            SubscribeToDb(e.Document.Database)
        End Sub
    End Class
End Namespace

