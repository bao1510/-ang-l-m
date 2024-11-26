' NX 10.0.3.5
' Journal created by user on Tue Nov 26 09:27:44 2024 東京 (標準時)
'
Option Strict Off
Imports System
Imports NXOpen
Imports NXOpen.UF

Module NXJournal
    Sub Main(ByVal args() As String)
        Dim ufs As UFSession = UFSession.GetUFSession()
        Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
        Dim workPart As NXOpen.Part = theSession.Parts.Work

        Dim displayPart As NXOpen.Part = theSession.Parts.Display


        Dim nullNXOpen_Features_BooleanFeature As NXOpen.Features.BooleanFeature = Nothing

        Dim booleanBuilder1 As NXOpen.Features.BooleanBuilder
        booleanBuilder1 = workPart.Features.CreateBooleanBuilderUsingCollector(nullNXOpen_Features_BooleanFeature)

        Dim scCollector1 As NXOpen.ScCollector
        scCollector1 = booleanBuilder1.ToolBodyCollector

        Dim booleanRegionSelect1 As NXOpen.GeometricUtilities.BooleanRegionSelect
        booleanRegionSelect1 = booleanBuilder1.BooleanRegionSelect

        booleanBuilder1.Tolerance = 0.01

        booleanBuilder1.Operation = NXOpen.Features.Feature.BooleanType.Subtract

        Dim a_body As Tag
        select_a_body(a_body)
        Dim body1 As NXOpen.Body = NXOpen.Utilities.NXObjectManager.Get(a_body)

        Dim added1 As Boolean
        added1 = booleanBuilder1.Targets.Add(body1)

        Dim targets1(0) As NXOpen.TaggedObject
        targets1(0) = body1
        booleanRegionSelect1.AssignTargets(targets1)

        Dim scCollector2 As NXOpen.ScCollector
        scCollector2 = workPart.ScCollectors.CreateCollector()

        Dim bodies1(0) As NXOpen.Body

        Dim a_body1 As Tag
        select_a_body(a_body1)

        Dim body2 As NXOpen.Body = NXOpen.Utilities.NXObjectManager.Get(a_body1)

        bodies1(0) = body2
        Dim bodyDumbRule1 As NXOpen.BodyDumbRule
        bodyDumbRule1 = workPart.ScRuleFactory.CreateRuleBodyDumb(bodies1, True)

        Dim rules1(0) As NXOpen.SelectionIntentRule
        rules1(0) = bodyDumbRule1
        scCollector2.ReplaceRules(rules1, False)

        booleanBuilder1.ToolBodyCollector = scCollector2

        Dim targets2(0) As NXOpen.TaggedObject
        targets2(0) = body1
        booleanRegionSelect1.AssignTargets(targets2)


        Dim nXObject1 As NXOpen.NXObject
        nXObject1 = booleanBuilder1.Commit()


        booleanBuilder1.Destroy()

    End Sub

    Function select_a_body(ByRef a_body As NXOpen.Tag) As Selection.Response
        Dim ufs As UFSession = UFSession.GetUFSession()
        Dim message As String = "Select a body"
        Dim title As String = "Select a body"
        Dim scope As Integer = UFConstants.UF_UI_SEL_SCOPE_ANY_IN_ASSEMBLY
        Dim response As Integer

        Dim view As NXOpen.Tag
        Dim cursor(2) As Double
        Dim ip As UFUi.SelInitFnT = AddressOf body_init_proc
        ufs.Ui.LockUgAccess(UFConstants.UF_UI_FROM_CUSTOM)

        Try
            ufs.Ui.SelectWithSingleDialog(message, title, scope, ip, Nothing, response, a_body, cursor, view)
        Finally
            ufs.Ui.UnlockUgAccess(UFConstants.UF_UI_FROM_CUSTOM)
        End Try

        If response <> UFConstants.UF_UI_OBJECT_SELECTED And response <> UFConstants.UF_UI_OBJECT_SELECTED_BY_NAME Then
            Return Selection.Response.Cancel
        Else
            ufs.Disp.SetHighlight(a_body, 0)
            Return Selection.Response.Ok
        End If

    End Function

    Function body_init_proc(ByVal select_ As IntPtr, ByVal userdata As IntPtr) As Integer
        Dim ufs As UFSession = UFSession.GetUFSession()
        Dim num_triples As Integer = 1
        Dim mask_triples(0) As UFUi.Mask
        mask_triples(0).object_type = UFConstants.UF_solid_type
        mask_triples(0).object_subtype = UFConstants.UF_solid_body_subtype
        mask_triples(0).solid_type = UFConstants.UF_UI_SEL_FEATURE_BODY

        ufs.Ui.SetSelMask(select_,
        UFUi.SelMaskAction.SelMaskClearAndEnableSpecific,
        num_triples, mask_triples)
        Return UFConstants.UF_UI_SEL_SUCCESS

    End Function

    Public Function GetUnloadOption(ByVal dummy As String) As Integer

        GetUnloadOption = UFConstants.UF_UNLOAD_IMMEDIATELY

    End Function

End Module
