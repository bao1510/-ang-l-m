Imports NXOpen.CAE
Imports NXOpen
Imports NXOpen.UF
Imports NXOpen.VectorArithmetic
Imports System.Collections
Imports System.Collections.Generic
Imports System
Imports System.Linq
Imports NXOpen.Features

Module aaaa

    Dim ufs As UFSession = UFSession.GetUFSession()
    Dim theSession As Session = Session.GetSession()
    Dim workPart As Part = theSession.Parts.Work
    Dim dispPart As NXOpen.Part = theSession.Parts.Display
    Dim lw As ListingWindow = theSession.ListingWindow

    Sub Main()

        lw.Open()
        Dim list_body_Part As New List(Of NXOpen.Body)

        list_body_Part = AskAll_visible(workPart)

        For Each selectedbody2 As NXOpen.Body In list_body_Part
            lw.WriteLine("body thu: " & list_body_Part.IndexOf(selectedbody2))
            Get_Properties_of_body(selectedbody2)
        Next


    End Sub
    ' Hàm để tạo và hiển thị hệ tọa độ tạm thời từ Point3d và Matrix3x3
    Function CreateAndDisplayTemporaryCoordinateSystem(ByVal origin As NXOpen.Point3d, ByVal orientation As NXOpen.Matrix3x3) As NXOpen.CoordinateSystem

        ' Chuyển đổi Point3d thành mảng Double() để truyền vào CreateTempCsys
        Dim originArray(2) As Double
        originArray(0) = origin.X
        originArray(1) = origin.Y
        originArray(2) = origin.Z

        ' Tạo mảng chứa các giá trị vector cần thiết để truyền vào CreateTempCsys
        Dim csysMatrix(8) As Double
        csysMatrix(0) = orientation.Xx
        csysMatrix(1) = orientation.Xy
        csysMatrix(2) = orientation.Xz
        csysMatrix(3) = orientation.Yx
        csysMatrix(4) = orientation.Yy
        csysMatrix(5) = orientation.Yz
        csysMatrix(6) = orientation.Zx
        csysMatrix(7) = orientation.Zy
        csysMatrix(8) = orientation.Zz

        Dim mtx As NXOpen.Tag = NXOpen.Tag.Null
        'Dim csys As NXOpen.Tag = NXOpen.Tag.Null
        ufs.Csys.CreateMatrix(csysMatrix, mtx)
        ' Tạo hệ tọa độ tạm thời
        Dim csysTag As NXOpen.Tag
        Try
            ufs.Csys.CreateTempCsys(originArray, mtx, csysTag)
        Catch ex As Exception
            Throw New Exception("Không thể tạo hệ tọa độ tạm thời: " & ex.Message)
        End Try

        ' Chuyển từ Tag thành đối tượng CoordinateSystem
        ' Dim workPart As NXOpen.Part = NXOpen.Session.GetSession().Parts.Work
        Dim tempCsys As NXOpen.CoordinateSystem = CType(NXOpen.Utilities.NXObjectManager.Get(csysTag), NXOpen.CoordinateSystem)

        ' Thêm hệ tọa độ vào phần làm việc


        ' Trả về đối tượng CoordinateSystem để tham khảo và thao tác sau này
        Return tempCsys
    End Function

    Public Function Get_Properties_of_body(ByVal abody As NXOpen.Body) As BodyProperties

        Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
        Dim lw As ListingWindow = theSession.ListingWindow
        lw.Open()
        Dim body_volume As Double = AREA_BODY(abody)
        'MsgBox("thể tích của body là: " & body_volume)
        Dim selectedFace2 As NXOpen.Face = Largest_planar_Face(abody)
        Dim result As body_Coordinate = FindCentroidOfFace(selectedFace2, abody)
        ' Lấy giá trị Point3d (origin) và Matrix3x3 (orientation) từ đối tượng BodyCoordinate
        Dim origin As NXOpen.Point3d = result.Origin
        Dim matrix As NXOpen.Matrix3x3 = result.Orientation

        Dim list_point1 As List(Of Point3d) = GetVerticesOfFace(selectedFace2)

        For Each p3d As NXOpen.Point3d In list_point1
            lw.WriteLine("x:" & p3d.X & ",y:" & p3d.Y & ",z:" & p3d.Z)
        Next

        Dim Csys1 As CartesianCoordinateSystem
        Csys1 = CreateAndDisplayTemporaryCoordinateSystem(origin, matrix)

        Dim datumCsysBuilder1 As NXOpen.Features.DatumCsysBuilder
        Dim nullNXOpen_Features_Feature As NXOpen.Features.Feature = Nothing
        datumCsysBuilder1 = workPart.Features.CreateDatumCsysBuilder(nullNXOpen_Features_Feature)
        datumCsysBuilder1.Csys = Csys1

        datumCsysBuilder1.DisplayScaleFactor = 1.25

        Dim nXObject1 As NXOpen.NXObject
        nXObject1 = datumCsysBuilder1.Commit()

        datumCsysBuilder1.Destroy()

    End Function

    ' đo thể tích của body
    Function AREA_BODY(ByVal theBody As NXOpen.Body) As Double
        Dim theBodies(0) As NXOpen.Body
        theBodies(0) = theBody
        Dim myMeasure As MeasureManager = theSession.Parts.Display.MeasureManager()
        Dim massUnits(1) As Unit
        'massUnits(0) = theSession.Parts.Display.UnitCollection.GetBase("Area")
        massUnits(0) = theSession.Parts.Display.UnitCollection.GetBase("Volume")
        'massUnits(2) = theSession.Parts.Display.UnitCollection.GetBase("Mass")
        'massUnits(3) = theSession.Parts.Display.UnitCollection.GetBase("Length")
        Dim mb As MeasureBodies = Nothing
        mb = myMeasure.NewMassProperties(massUnits, 0.99, theBodies)
        mb.InformationUnit = MeasureBodies.AnalysisUnit.GramMillimeter

        AREA_BODY = mb.Volume

    End Function

    Public Class HoleInfo
        Public Property Diameter As Double
        Public Property StartPoint As Point3d
        Public Property EndPoint As Point3d
    End Class

    Public Function FindHoleFaces(theBody As NXOpen.Body) As List(Of HoleInfo)

        Dim holeInfos As New List(Of HoleInfo)()
        ' Iterate through all faces of the body (theFace is assumed to be a face object)
        For Each theFace As Face In theBody.GetFaces()

            Dim f_type As Integer
            ufs.Modl.AskFaceType(theFace.Tag, f_type)
            If f_type = 16 Then

                Dim faceType As Integer
                Dim facePt(2) As Double
                Dim faceDir(2) As Double
                Dim bbox(5) As Double
                Dim faceRadius As Double
                Dim faceRadData As Double
                Dim normDirection As Integer

                ' Retrieve face data using the NXOpen API
                ufs.Modl.AskFaceData(theFace.Tag, faceType, facePt, faceDir, bbox, faceRadius, faceRadData, normDirection)

                Dim basePoint(2) As Double
                Dim direction(2) As Double
                Dim radius As String
                Dim height As String

                ' Ask cylinder parameters to get radius and height of the hole
                ufs.Modl.AskCylinderParms(theFace.Tag, 0, radius, height)

                lw.WriteLine("thông số của lỗ là R:" & radius & "," & "độ dài: " & height)

                ' Create an instance of HoleInfo to store the details
                Dim holeInfo As New HoleInfo()
                holeInfo.Diameter = radius * 2
                ' Check if the faceType is 16 (cylinder face) and normDirection is -1
                If faceType <> 17 AndAlso normDirection = -1 Then
                    ' Create a new instance of HoleInfo
                    ' Set the diameter of the hole
                    holeInfo.Diameter = faceRadius * 2

                    ' Set the start and end points of the axis of the hole
                    Dim startPoint As New Point3d(facePt(0), facePt(1), facePt(2))
                    Dim endPoint As New Point3d(facePt(0) + faceDir(0) * faceRadius * 2,
                                            facePt(1) + faceDir(1) * faceRadius * 2,
                                            facePt(2) + faceDir(2) * faceRadius * 2)

                    holeInfo.StartPoint = startPoint
                    holeInfo.EndPoint = endPoint

                    ' Add the holeInfo instance to the list
                    holeInfos.Add(holeInfo)

                End If

            End If

        Next

        Return holeInfos

    End Function


    'Sub Main()
    '        ' Ví dụ sử dụng hàm MeasureTubeLength
    '        Dim s As Session = Session.GetSession()
    '        Dim workPart As Part = s.Parts.Work
    '        Dim tubeLengthTotal As Double = MeasureTubeLength(workPart.Bodies)
    '        Console.WriteLine("Total Tube Length: " & tubeLengthTotal.ToString())
    '    End Sub

    Function MeasureTubeLength(bodies As BodyCollection) As Double
            Dim ufs As UFSession = UFSession.GetUFSession()
            Dim totalTubeLength As Double = 0

            For Each theBody As Body In bodies
                For Each theFace As Face In theBody.GetFaces()
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

                    ' Xử lý nếu mặt đóng theo hướng U hoặc V
                    If knotsU(0) < 0.0 Then
                        ' Đóng theo hướng U
                        totalTubeLength += CalculateSplineLength(testface, knotsU, knotsV, uDirection:=True)
                    ElseIf knotsV(0) < 0.0 Then
                        ' Đóng theo hướng V
                        totalTubeLength += CalculateSplineLength(testface, knotsU, knotsV, uDirection:=False)
                    End If
                Next
            Next

            Return totalTubeLength
        End Function

    Function CalculateSplineLength(testface As Face, knotsU() As Double, knotsV() As Double, uDirection As Boolean) As Double
        Dim ufs As UFSession = UFSession.GetUFSession()
        Dim wp As Part = Session.GetSession().Parts.Work
        Dim ArrayOfPoints() As Point
        Dim params(1) As Double
        Dim pnt0(2) As Double
        Dim pnt1(2) As Double
        Dim junk3(2) As Double
        Dim junk2(1) As Double
        Dim uparm() As Double
        Dim vparm() As Double

        ' Tạo các giá trị uparm và vparm
        If uDirection Then
            ' Đóng theo hướng U
            uparm = {0.0, 0.5}
            ReDim vparm(knotsV.Length - 5)
            vparm(0) = 0.0
            For i As Integer = 1 To knotsV.Length - 7 Step 3
                vparm(i) = knotsV(i + 2) + (knotsV(i + 3) - knotsV(i + 2)) / 3.0
                vparm(i + 1) = knotsV(i + 2) + 2.0 * (knotsV(i + 3) - knotsV(i + 2)) / 3.0
                vparm(i + 2) = knotsV(i + 3)
            Next
            ReDim ArrayOfPoints(vparm.Length - 1)
        Else
            ' Đóng theo hướng V
            vparm = {0.0, 0.5}
            ReDim uparm(knotsU.Length - 5)
            uparm(0) = 0.0
            For i As Integer = 1 To knotsU.Length - 7 Step 3
                uparm(i) = knotsU(i + 2) + (knotsU(i + 3) - knotsU(i + 2)) / 3.0
                uparm(i + 1) = knotsU(i + 2) + 2.0 * (knotsU(i + 3) - knotsU(i + 2)) / 3.0
                uparm(i + 2) = knotsU(i + 3)
            Next
            ReDim ArrayOfPoints(uparm.Length - 1)
        End If

        ' Tạo các điểm trên mặt
        For i As Integer = 0 To ArrayOfPoints.Length - 1
            If uDirection Then
                params(0) = uparm(0)
                params(1) = vparm(i)
            Else
                params(0) = uparm(i)
                params(1) = vparm(0)
            End If
            ufs.Modl.AskFaceProps(testface.Tag, params, pnt0, junk3, junk3, junk3, junk3, junk3, junk2)
            If uDirection Then
                params(0) = uparm(1)
            Else
                params(1) = vparm(1)
            End If
            ufs.Modl.AskFaceProps(testface.Tag, params, pnt1, junk3, junk3, junk3, junk3, junk3, junk2)

            Dim coordinates1 As New Point3d((pnt0(0) + pnt1(0)) / 2.0, (pnt0(1) + pnt1(1)) / 2.0, (pnt0(2) + pnt1(2)) / 2.0)
            ArrayOfPoints(i) = wp.Points.CreatePoint(coordinates1)
        Next

        ' Tạo spline và tính chiều dài của nó
        Dim myStudioSpline As Features.StudioSpline = CreateStudioSplineThruPoints(ArrayOfPoints)
        Dim tubeLength As Double = 0
        For Each tempCurve As Curve In myStudioSpline.GetEntities
            tubeLength += tempCurve.GetLength
        Next

        Return tubeLength

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

    Public Function CreateStudioSplineThruPoints(ByRef points() As Point) As Features.StudioSpline
        Dim markId9 As Session.UndoMarkId
        markId9 = theSession.SetUndoMark(Session.MarkVisibility.Visible, "Studio Spline Thru Points")
        Dim Pcount As Integer = points.Length - 1
        Dim nullFeatures_StudioSpline As Features.StudioSpline = Nothing
        Dim studioSplineBuilderex1 As Features.StudioSplineBuilderEx
        studioSplineBuilderex1 = workPart.Features.CreateStudioSplineBuilderEx(nullFeatures_StudioSpline)
        studioSplineBuilderex1.OrientExpress.ReferenceOption = GeometricUtilities.OrientXpressBuilder.Reference.ProgramDefined
        studioSplineBuilderex1.Degree = 3
        studioSplineBuilderex1.OrientExpress.AxisOption = GeometricUtilities.OrientXpressBuilder.Axis.Passive
        studioSplineBuilderex1.OrientExpress.PlaneOption = GeometricUtilities.OrientXpressBuilder.Plane.Passive
        studioSplineBuilderex1.MatchKnotsType = Features.StudioSplineBuilderEx.MatchKnotsTypes.None
        Dim knots1(-1) As Double
        studioSplineBuilderex1.SetKnots(knots1)
        Dim parameters1(-1) As Double
        studioSplineBuilderex1.SetParameters(parameters1)
        Dim nullDirection As Direction = Nothing
        Dim nullScalar As Scalar = Nothing
        Dim nullOffset As Offset = Nothing
        Dim geometricConstraintData(Pcount) As Features.GeometricConstraintData
        For ii As Integer = 0 To Pcount
            geometricConstraintData(ii) = studioSplineBuilderex1.ConstraintManager.CreateGeometricConstraintData()
            geometricConstraintData(ii).Point = points(ii)
            geometricConstraintData(ii).AutomaticConstraintDirection = Features.GeometricConstraintData.ParameterDirection.Iso
            geometricConstraintData(ii).AutomaticConstraintType = Features.GeometricConstraintData.AutoConstraintType.Tangent
            geometricConstraintData(ii).TangentDirection = nullDirection
            geometricConstraintData(ii).TangentMagnitude = nullScalar
            geometricConstraintData(ii).Curvature = nullOffset
            geometricConstraintData(ii).CurvatureDerivative = nullOffset
            geometricConstraintData(ii).HasSymmetricModelingConstraint = False
        Next ii
        studioSplineBuilderex1.ConstraintManager.SetContents(geometricConstraintData)
        Dim feature1 As Features.StudioSpline
        feature1 = studioSplineBuilderex1.CommitFeature()
        studioSplineBuilderex1.Destroy()
        Return feature1
    End Function

    Public Function SelectFaces(ByVal propt As String) As Face()

        Dim theUI As UI = UI.GetUI
        Dim title As String = "Select faces"
        Dim includeFeatures As Boolean = False
        Dim keepHighlighted As Boolean = False
        Dim selAction As Selection.SelectionAction = Selection.SelectionAction.ClearAndEnableSpecific
        Dim scope As Selection.SelectionScope = Selection.SelectionScope.AnyInAssembly
        Dim selectionMask(0) As Selection.MaskTriple
        Dim selectedObjects() As TaggedObject = Nothing
        Dim selectedFaces As New List(Of Face)

        With selectionMask(0)
            .Type = UFConstants.UF_solid_type
            .Subtype = 0
            .SolidBodySubtype = UFConstants.UF_UI_SEL_FEATURE_ANY_FACE
        End With

        Dim responce1 As Selection.Response = theUI.SelectionManager.SelectTaggedObjects(
        propt, title, scope, selAction, includeFeatures, keepHighlighted, selectionMask, selectedObjects)

        If responce1 = Selection.Response.Ok Then
            For Each item As TaggedObject In selectedObjects
                selectedFaces.Add(item)
            Next
            Return selectedFaces.ToArray
        Else
            Return Nothing
        End If


    End Function



    Public Function FindCentroidOfFace(ByVal face As NXOpen.Face, temp_body As NXOpen.Body) As body_Coordinate
        ' Lấy pháp tuyến của mặt phẳng làm trục Z và tọa độ tâm của mặt phẳng
        Dim norm(2) As Double
        Dim centroid(2) As Double
        Dim u1(2) As Double, v1(2) As Double, u2(2) As Double, v2(2) As Double
        Dim radii(1) As Double
        Dim param(1) As Double

        ufs.Modl.AskFaceProps(face.Tag, param, centroid, u1, v1, u2, v2, norm, radii)

        ' Tìm cạnh dài nhất của mặt phẳng để xác định trục X
        Dim edges() As NXOpen.Edge = face.GetEdges()
        Dim longestEdge As NXOpen.Edge = Nothing
        Dim maxLength As Double = 0

        For Each edge As NXOpen.Edge In edges
            If edge.SolidEdgeType = Edge.EdgeType.Linear Then
                Dim length As Double = edge.GetLength()
                If length > maxLength Then
                    maxLength = length
                    longestEdge = edge
                End If
            End If
        Next

        ' Nếu không có cạnh dài nhất (trong trường hợp mặt phẳng hình tròn)
        Dim xVector(2) As Double
        If longestEdge IsNot Nothing Then
            ' Lấy điểm đầu và điểm cuối của cạnh dài nhất
            Dim startPoint(2) As Double
            Dim endPoint(2) As Double
            Dim vertexCount As Integer
            ufs.Modl.AskEdgeVerts(longestEdge.Tag, startPoint, endPoint, vertexCount)

            ' Xác định trục X từ cạnh dài nhất và chuẩn hóa
            xVector(0) = endPoint(0) - startPoint(0)
            xVector(1) = endPoint(1) - startPoint(1)
            xVector(2) = endPoint(2) - startPoint(2)

            Dim lengthX As Double = Math.Sqrt(xVector(0) ^ 2 + xVector(1) ^ 2 + xVector(2) ^ 2)
            xVector(0) /= lengthX
            xVector(1) /= lengthX
            xVector(2) /= lengthX

        Else
            ' Tạo trục X ngẫu nhiên nếu không có cạnh dài nhất
            ufs.Vec3.AskPerpendicular(norm, xVector)

            ' Chuẩn hóa vector X
            Dim lengthX As Double = Math.Sqrt(xVector(0) ^ 2 + xVector(1) ^ 2 + xVector(2) ^ 2)
            xVector(0) /= lengthX
            xVector(1) /= lengthX
            xVector(2) /= lengthX

        End If

        ' MsgBox("tọa độ vector x là:" & xVector(0) & "," & xVector(1) & "," & xVector(2))
        ' Chuẩn bị vector Z từ pháp tuyến của mặt phẳng
        Dim zVector() As Double = norm
        ' MsgBox("tọa độ vector z là:" & zVector(0) & "," & zVector(1) & "," & zVector(2))
        ' Tính toán trục Y bằng tích chéo của Z và X và chuẩn hóa (thứ tự đúng)
        Dim yVector(2) As Double
        yVector(0) = zVector(1) * xVector(2) - zVector(2) * xVector(1)
        yVector(1) = zVector(2) * xVector(0) - zVector(0) * xVector(2)
        yVector(2) = zVector(0) * xVector(1) - zVector(1) * xVector(0)

        ' Chuẩn hóa vector Y
        Dim lengthY As Double = Math.Sqrt(yVector(0) ^ 2 + yVector(1) ^ 2 + yVector(2) ^ 2)

        ' Kiểm tra nếu độ dài bằng 0 thì phải xử lý tránh chia cho 0
        If lengthY > 0 Then
            yVector(0) /= lengthY
            yVector(1) /= lengthY
            yVector(2) /= lengthY
        Else
            Throw New Exception("Không thể chuẩn hóa vector Y, độ dài của vector bằng 0")
        End If

        ' Khai báo mảng để lưu trữ ma trận
        Dim matrix_v_1(8) As Double

        ' Sử dụng ufs.Mtx3.Initialize để tạo ra ma trận từ vector X, Y và Z
        ufs.Mtx3.Initialize(xVector, yVector, matrix_v_1)

        ' Khởi tạo hệ tọa độ với các vector vừa tạo
        Dim mtx As NXOpen.Tag = NXOpen.Tag.Null
        Dim csys As NXOpen.Tag = NXOpen.Tag.Null

        ' Tạo ma trận từ các vector đã tính toán
        ufs.Csys.CreateMatrix(matrix_v_1, mtx)

        ' Tạo hệ tọa độ tạm thời tại centroid với ma trận đã tạo
        ufs.Csys.CreateTempCsys(centroid, mtx, csys)

        ' Tính hộp bao chính xác của mặt phẳng dựa trên hệ tọa độ vừa tạo
        Dim minCorner(2) As Double
        Dim directions(2, 2) As Double
        Dim distances(2) As Double

        ufs.Modl.AskBoundingBoxExact(temp_body.Tag, csys, minCorner, directions, distances)

        Dim origin_a As New NXOpen.Point3d(0, 0, 0)
        Dim absMatrix As NXOpen.Matrix3x3 = New NXOpen.Matrix3x3()
        absMatrix.Xx = 1.0
        absMatrix.Xy = 0.0
        absMatrix.Xz = 0.0
        absMatrix.Yx = 0.0
        absMatrix.Yy = 1.0
        absMatrix.Yz = 0.0
        absMatrix.Zx = 0.0
        absMatrix.Zy = 0.0
        absMatrix.Zz = 1.0

        ' Chuyển đổi minCorner từ hệ tọa độ tuyệt đối sang hệ tọa độ riêng của khối
        Dim localMinCorner As NXOpen.Point3d = point_Csys_to_Csys(New NXOpen.Point3d(minCorner(0), minCorner(1), minCorner(2)),
       convertToMatrix3x3(matrix_v_1), New NXOpen.Point3d(centroid(0), centroid(1), centroid(2)),
     absMatrix, origin_a)

        ' Tính toán tọa độ tâm của hộp bao trong hệ tọa độ tuyệt đối
        Dim temp_point1 As Point3d
        temp_point1 = New NXOpen.Point3d(localMinCorner.X + distances(0) / 2, localMinCorner.Y + distances(1) / 2, localMinCorner.Z + distances(2) / 2)

        ' Trả về tâm của mặt phẳng

        Dim absoluteCentroid As NXOpen.Point3d = point_Csys_to_Csys(temp_point1, absMatrix, origin_a,
       convertToMatrix3x3(matrix_v_1), New NXOpen.Point3d(centroid(0), centroid(1), centroid(2)))

        Return New body_Coordinate(absoluteCentroid, convertToMatrix3x3(matrix_v_1))

    End Function

    Public Class BodyProperties
        Public Property Volume As Double
        Public Property LargestFaceArea As Double
        Public Property body_coor As body_Coordinate
        Public Property Holes As List(Of HoleInfo) = New List(Of HoleInfo)()
        Public Property SpecialPoints As List(Of Point3d) = New List(Of Point3d)()
    End Class

    Public Class body_Coordinate
        Public Property Origin As NXOpen.Point3d
        Public Property Orientation As NXOpen.Matrix3x3

        Public Sub New(origin As NXOpen.Point3d, orientation As NXOpen.Matrix3x3)
            Me.Origin = origin
            Me.Orientation = orientation
        End Sub

    End Class
    '' Lớp lưu thông tin lỗ
    'Public Class HoleInfo
    '    Public Property Diameter As Double
    '    Public Property point3d() As Point3d
    'End Class
    ' Lớp lưu thông tọa độ face
    Public Class face_Coordinate
        Public Property Origin As NXOpen.Point3d
        Public Property Orientation As NXOpen.Matrix3x3

        Public Sub New(origin As NXOpen.Point3d, orientation As NXOpen.Matrix3x3)
            Me.Origin = origin
            Me.Orientation = orientation
        End Sub

    End Class
    ' Hàm kiểm tra xem hai điểm có giống nhau trong phạm vi sai số cho phép
    Function ArePointsEqual(p1 As Point3d, p2 As Point3d, epsilon As Double) As Boolean
        Return Math.Abs(p1.X - p2.X) < epsilon AndAlso
               Math.Abs(p1.Y - p2.Y) < epsilon AndAlso
               Math.Abs(p1.Z - p2.Z) < epsilon
    End Function
    ' Hàm so sánh hai danh sách Point3d và trả về danh sách đã sắp xếp
    Function SortPointListManual(list As List(Of Point3d)) As List(Of Point3d)

        list.Sort(Function(p1, p2)
                      ' So sánh theo trục X
                      If p1.X <> p2.X Then
                          Return p1.X.CompareTo(p2.X)
                          ' Nếu X bằng nhau, so sánh theo trục Y
                      ElseIf p1.Y <> p2.Y Then
                          Return p1.Y.CompareTo(p2.Y)
                          ' Nếu cả X và Y bằng nhau, so sánh theo Z
                      Else
                          Return p1.Z.CompareTo(p2.Z)
                      End If
                  End Function)

        Return list
    End Function

    ' Hàm so sánh hai danh sách Point3d
    Function ComparePointLists(list1 As List(Of Point3d), list2 As List(Of Point3d), epsilon As Double) As Boolean
        ' Sắp xếp cả hai danh sách
        Dim sortedList1 = SortPointListManual(list1)
        Dim sortedList2 = SortPointListManual(list2)

        ' Kiểm tra xem các danh sách có cùng số phần tử không
        If sortedList1.Count <> sortedList2.Count Then
            Return False
        End If

        ' Kiểm tra từng điểm trong hai danh sách đã sắp xếp
        For i As Integer = 0 To sortedList1.Count - 1
            If Not ArePointsEqual(sortedList1(i), sortedList2(i), epsilon) Then
                Return False
            End If
        Next

        ' Nếu không có sự khác biệt, danh sách giống nhau
        Return True
    End Function

    Function Largest_planar_Face(ByVal inputSolid As NXOpen.Body) As NXOpen.Face

        Dim workpart As NXOpen.Part = theSession.Parts.Work
        Dim dispPart As NXOpen.Part = theSession.Parts.Display

        Dim nullNXObject As NXObject = Nothing
        Dim measureFaceBuilder1 As MeasureFaceBuilder
        measureFaceBuilder1 = workpart.MeasureManager.CreateMeasureFaceBuilder(nullNXObject)

        Dim unit1 As Unit = CType(workpart.UnitCollection.FindObject("SquareMilliMeter"), Unit)
        Dim unit2 As Unit = CType(workpart.UnitCollection.FindObject("MilliMeter"), Unit)

        Dim objects1(0) As IParameterizedSurface
        Dim measureFaces1 As MeasureFaces

        Dim myFaces() As NXOpen.Face
        myFaces = inputSolid.GetFaces
        Dim largestFace As NXOpen.Face = myFaces(0)
        Dim largestArea As Double
        Dim added1 As Boolean
        Dim i As Integer = 0

        For Each tempFace As NXOpen.Face In myFaces
            Select Case tempFace.SolidFaceType
                Case 1  '2, 3, 5
                    i += 1
                    measureFaceBuilder1.FaceObjects.Clear()
                    added1 = measureFaceBuilder1.FaceObjects.Add(tempFace)
                    objects1(0) = tempFace
                    measureFaces1 = workpart.MeasureManager.NewFaceProperties(unit1, unit2, 0.999, objects1)
                    If i = 1 Then
                        largestFace = tempFace
                        largestArea = measureFaces1.Area
                    Else
                        If measureFaces1.Area > largestArea Then

                            largestFace = tempFace

                            largestArea = measureFaces1.Area


                        End If
                    End If

                Case Else

            End Select
        Next

        measureFaces1.Dispose()
        measureFaceBuilder1.FaceObjects.Clear()
        measureFaceBuilder1.Destroy()

        If i = 0 Then
            Return Nothing
        Else
            Return largestFace
        End If

    End Function
    Public Function point_Csys_to_Csys(ByVal point_1 As NXOpen.Point3d, ByVal matrix_1 As NXOpen.Matrix3x3,
                                       ByVal origin1 As NXOpen.Point3d, ByVal matrix_2 As NXOpen.Matrix3x3,
                                       ByVal origin2 As NXOpen.Point3d) As NXOpen.Point3d

        ' Tạo các vector trục từ ma trận định hướng
        Dim fromXAxis As Double() = {matrix_1.Xx, matrix_1.Xy, matrix_1.Xz}
        Dim fromYAxis As Double() = {matrix_1.Yx, matrix_1.Yy, matrix_1.Yz}
        Dim toXAxis As Double() = {matrix_2.Xx, matrix_2.Xy, matrix_2.Xz}
        Dim toYAxis As Double() = {matrix_2.Yx, matrix_2.Yy, matrix_2.Yz}

        ' Tạo ma trận chuyển đổi 4x4
        Dim mtx4Transform(15) As Double

        ' Sử dụng hàm CsysToCsys để tính toán ma trận chuyển đổi từ hệ tọa độ ban đầu sang hệ tọa độ mới
        ufs.Mtx4.CsysToCsys(New Double() {origin1.X, origin1.Y, origin1.Z}, fromXAxis, fromYAxis, New Double() {origin2.X, origin2.Y, origin2.Z}, toXAxis, toYAxis, mtx4Transform)

        ' Áp dụng ma trận chuyển đổi lên điểm ban đầu để tính toán điểm mới
        Dim newX As Double = mtx4Transform(0) * point_1.X + mtx4Transform(1) * point_1.Y + mtx4Transform(2) * point_1.Z + mtx4Transform(3)
        Dim newY As Double = mtx4Transform(4) * point_1.X + mtx4Transform(5) * point_1.Y + mtx4Transform(6) * point_1.Z + mtx4Transform(7)
        Dim newZ As Double = mtx4Transform(8) * point_1.X + mtx4Transform(9) * point_1.Y + mtx4Transform(10) * point_1.Z + mtx4Transform(11)

        ' Trả về điểm mới sau khi chuyển đổi
        Return New NXOpen.Point3d(newX, newY, newZ)

    End Function

    Public Function GetVerticesOfFace(ByVal face As NXOpen.Face) As List(Of NXOpen.Point3d)
        ' Tạo một danh sách để lưu trữ các đỉnh duy nhất
        Dim vertexList As New List(Of NXOpen.Point3d)()

        ' Lấy tất cả các cạnh của mặt phẳng
        Dim edges() As NXOpen.Edge = face.GetEdges()

        ' Duyệt qua tất cả các cạnh của mặt phẳng
        For Each edge As NXOpen.Edge In edges
            ' Lấy điểm đầu và điểm cuối của cạnh
            Dim startPoint(2) As Double
            Dim endPoint(2) As Double
            Dim vertexCount As Integer

            ' Sử dụng AskEdgeVerts để lấy các điểm đầu và cuối của cạnh
            ufs.Modl.AskEdgeVerts(edge.Tag, startPoint, endPoint, vertexCount)

            ' Chuyển đổi điểm đầu và điểm cuối thành Point3d
            Dim start As New NXOpen.Point3d(startPoint(0), startPoint(1), startPoint(2))
            Dim [end] As New NXOpen.Point3d(endPoint(0), endPoint(1), endPoint(2))

            ' Thêm điểm đầu vào danh sách nếu nó chưa tồn tại
            Dim isStartAdded As Boolean = False
            For Each vertex As NXOpen.Point3d In vertexList
                If IsSamePoint(vertex, start) Then
                    isStartAdded = True
                    Exit For
                End If
            Next
            If Not isStartAdded Then
                vertexList.Add(start)
            End If

            ' Thêm điểm cuối vào danh sách nếu nó chưa tồn tại
            Dim isEndAdded As Boolean = False
            For Each vertex As NXOpen.Point3d In vertexList
                If IsSamePoint(vertex, [end]) Then
                    isEndAdded = True
                    Exit For
                End If
            Next
            If Not isEndAdded Then
                vertexList.Add([end])
            End If
        Next

        ' Chuyển danh sách thành mảng và trả về

        Return SortPointListManual(vertexList) 'vertexList.ToArray()
    End Function

    ' Hàm để kiểm tra xem hai điểm có trùng nhau không
    Private Function IsSamePoint(ByVal p1 As NXOpen.Point3d, ByVal p2 As NXOpen.Point3d) As Boolean
        Const TOLERANCE As Double = 0.000001 ' Ngưỡng để kiểm tra độ chính xác
        Return Math.Abs(p1.X - p2.X) < TOLERANCE AndAlso
           Math.Abs(p1.Y - p2.Y) < TOLERANCE AndAlso
           Math.Abs(p1.Z - p2.Z) < TOLERANCE
    End Function

    Function AskAll_visible(ByVal the_part As NXOpen.Part) As List(Of NXOpen.Body)
        Dim visibleObjects() As DisplayableObject

        Dim mviews As ModelingViewCollection = theSession.Parts.Work.ModelingViews

        Dim topView As NXOpen.ModelingView

        For Each mv As ModelingView In mviews
            If mv.Name.Equals("Top") Then
                topView = mv
                theSession.Parts.Work.Layouts.Current.ReplaceView(theSession.Parts.Work.ModelingViews.WorkView, topView, True)
            End If
        Next

        visibleObjects = topView.AskVisibleObjects()
        Dim countBody As Integer = 0
        Dim icount As Integer = 0
        Dim tagCount As Integer
        Dim arrayBody As New List(Of NXOpen.Body)


        For Each obj As NXObject In visibleObjects
            If obj.GetType.Name = "Body" Then
                'lw.WriteLine(obj.GetType.Name.ToString)
                arrayBody.Add(obj)
                countBody = countBody + 1
            End If


        Next

        Return arrayBody
    End Function


    Function convertToMatrix3x3(ByVal mtx As Double()) As Matrix3x3

        Dim mx As Matrix3x3
        With mx
            .Xx = mtx(0)
            .Xy = mtx(1)
            .Xz = mtx(2)
            .Yx = mtx(3)
            .Yy = mtx(4)
            .Yz = mtx(5)
            .Zx = mtx(6)
            .Zy = mtx(7)
            .Zz = mtx(8)
        End With

        Return mx

    End Function

    '**********************************************************
    Public Function GetUnloadOption(ByVal dummy As String) As Integer
        Return CType(NXOpen.Session.LibraryUnloadOption.Immediately, Integer)
    End Function
    '**********************************************************

End Module