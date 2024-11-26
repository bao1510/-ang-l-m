Imports NXOpen.CAE
Imports NXOpen
Imports NXOpen.UF
Imports NXOpen.VectorArithmetic
Imports System.Collections
Imports System.Collections.Generic
Imports System
Imports System.Linq
Imports System.String
Imports NXOpen.Features

Module abc

    Dim ufs As UFSession = UFSession.GetUFSession()
    Dim theSession As Session = Session.GetSession()
    Dim workPart As Part = theSession.Parts.Work
    Dim dispPart As NXOpen.Part = theSession.Parts.Display
    Dim lw As ListingWindow = theSession.ListingWindow


    Sub Main()

        lw.Open()
        Dim a_body As Tag
        select_a_body(a_body)
        Dim body1 As NXOpen.Body = NXOpen.Utilities.NXObjectManager.Get(a_body)
        Dim holes1 As List(Of HoleInfo) = MeasureTubeLength(body1)

        For Each tt_lo As HoleInfo In holes1
            lw.WriteLine("lỗ thứ :" & holes1.IndexOf(tt_lo))
            lw.WriteLine(tt_lo.Diameter)
            lw.WriteLine("diem thu nhat: x" & tt_lo.StartPoint.X & ",y: " & tt_lo.StartPoint.Y & "z:" & tt_lo.StartPoint.Z)
            lw.WriteLine("diem thu nhat: x" & tt_lo.EndPoint.X & ",y: " & tt_lo.EndPoint.Y & "z:" & tt_lo.EndPoint.Z)
        Next

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

    Public Class HoleInfo
        Public Property Diameter As Double
        Public Property StartPoint As Point3d
        Public Property EndPoint As Point3d
    End Class

    Function MeasureTubeLength(bodies As NXOpen.Body) As List(Of HoleInfo)
        Dim ufs As UFSession = UFSession.GetUFSession()
        Dim holeInfos As New List(Of HoleInfo)()


        For Each theFace As Face In bodies.GetFaces()
            ' Lấy loại mặt
            Dim f_type As Integer
            ufs.Modl.AskFaceType(theFace.Tag, f_type)

            ' Chỉ quan tâm đến các mặt trụ (f_type = 16)
            If f_type <> 16 Then
                Continue For
            End If

            ' Trích xuất bề mặt B-spline
            Dim extractedfeat1 As Feature = Nothing
            createExtractedBSurface(theFace, extractedfeat1)
            Dim extractedbodyfeat1 As BodyFeature = DirectCast(extractedfeat1, BodyFeature)
            Dim extractedbody1() As Body = extractedbodyfeat1.GetBodies
            Dim faces() As Face = extractedbody1(0).GetFaces


            ' Kiểm tra và lấy thông số bề mặt B-spline
            Dim bsurface1 As UFModl.Bsurface = Nothing
            ufs.Modl.AskBsurf(faces(0).Tag, bsurface1)
            Dim knotsU() As Double = bsurface1.knots_u
            Dim knotsV() As Double = bsurface1.knots_v
            Dim testface As Face = faces(0)

            ' Xử lý nếu mặt đóng theo hướng U hoặc V và lấy thông tin lỗ
            If knotsU(0) < 0.0 Then
                ' Đóng theo hướng U
                holeInfos.Add(CalculateHoleInfo(testface, uparm:=0.0, uparm1:=0.5, vparm:=0.0))
            ElseIf knotsV(0) < 0.0 Then
                ' Đóng theo hướng V
                holeInfos.Add(CalculateHoleInfo(testface, uparm:=0.0, uparm1:=0.5, vparm:=0.0))
            End If
        Next

        Return holeInfos
    End Function

    Function CalculateHoleInfo(testface As Face, uparm As Double, uparm1 As Double, vparm As Double) As HoleInfo
        Dim ufs As UFSession = UFSession.GetUFSession()
        Dim wp As Part = Session.GetSession().Parts.Work
        Dim params(1) As Double
        Dim pnt0(2) As Double
        Dim pnt1(2) As Double
        Dim junk3(2) As Double
        Dim junk2(1) As Double

        ' Lấy thông tin điểm đầu và điểm cuối
        params(0) = uparm
        params(1) = vparm
        ufs.Modl.AskFaceProps(testface.Tag, params, pnt0, junk3, junk3, junk3, junk3, junk3, junk2)

        params(0) = uparm1
        ufs.Modl.AskFaceProps(testface.Tag, params, pnt1, junk3, junk3, junk3, junk3, junk3, junk2)

        ' Tạo HoleInfo và lưu các giá trị
        Dim holeInfo As New HoleInfo()
        holeInfo.StartPoint = New Point3d(pnt0(0), pnt0(1), pnt0(2))
        holeInfo.EndPoint = New Point3d(pnt1(0), pnt1(1), pnt1(2))

        ' Tính bán kính và đường kính
        Dim radius As Double = Math.Sqrt((pnt1(0) - pnt0(0)) ^ 2 + (pnt1(1) - pnt0(1)) ^ 2 + (pnt1(2) - pnt0(2)) ^ 2) / 2.0
        ' holeInfo.Radius = radius
        holeInfo.Diameter = radius * 2

        Return holeInfo
    End Function

    Public Sub createExtractedBSurface(ByVal face1 As Face, ByRef extractedfeat1 As Feature)
        Dim nullFeatures_Feature As Features.Feature = Nothing
        Dim extractFaceBuilder1 As Features.ExtractFaceBuilder
        extractFaceBuilder1 = workPart.Features.CreateExtractFaceBuilder(nullFeatures_Feature)
        extractFaceBuilder1.ParentPart = Features.ExtractFaceBuilder.ParentPartType.WorkPart
        extractFaceBuilder1.Associative = True
        extractFaceBuilder1.FixAtCurrentTimestamp = True
        extractFaceBuilder1.HideOriginal = False
        extractFaceBuilder1.Type = Features.ExtractFaceBuilder.ExtractType.Face
        extractFaceBuilder1.InheritDisplayProperties = False
        extractFaceBuilder1.SurfaceType = Features.ExtractFaceBuilder.FaceSurfaceType.PolynomialCubic
        Dim added1 As Boolean
        added1 = extractFaceBuilder1.ObjectToExtract.Add(face1)
        extractedfeat1 = extractFaceBuilder1.Commit
    End Sub

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
End Module
