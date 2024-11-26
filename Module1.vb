Option Strict Off
Imports System
Imports NXOpen
Imports NXOpen.BlockStyler
Imports NXOpen.UF
Imports PLMComponents.Parasolid.PK_.Unsafe
Imports System.Collections
Imports System.Collections.Generic
Imports PLMComponents.Parasolid.PK_DEBUG_.Unsafe

Module FindCentroidOfFace
    Dim ufs As UFSession = UFSession.GetUFSession()
    Public theSession As Session = Session.GetSession()
    Dim workpart As NXOpen.Part = theSession.Parts.Work
    Dim dispPart As NXOpen.Part = theSession.Parts.Display

    Sub Main()

        Dim list_body_Part As New List(Of NXOpen.Body)
        list_body_Part = AskAll_visible(workpart)

        For Each selectedbody2 As NXOpen.Body In list_body_Part

            Dim selectedFace2 As NXOpen.Face = Largest_planar_Face(selectedbody2)

            Dim result As BodyCoordinate = FindCentroidOfFace(selectedFace2)

            ' Lấy giá trị Point3d (origin) và Matrix3x3 (orientation) từ đối tượng BodyCoordinate
            Dim origin As NXOpen.Point3d = result.Origin
            Dim matrix As NXOpen.Matrix3x3 = result.Orientation

            Dim csys1 As NXOpen.CartesianCoordinateSystem
            csys1 = workpart.CoordinateSystems.CreateCoordinateSystem(origin, matrix, True)

            Dim nullNXOpen_Features_Feature As NXOpen.Features.Feature = Nothing

            Dim datumCsysBuilder1 As NXOpen.Features.DatumCsysBuilder = Nothing
            datumCsysBuilder1 = workpart.Features.CreateDatumCsysBuilder(nullNXOpen_Features_Feature)
            datumCsysBuilder1.Csys = csys1

            datumCsysBuilder1.DisplayScaleFactor = 1.25

            Dim nXObject1 As NXOpen.NXObject = Nothing
            nXObject1 = datumCsysBuilder1.Commit()

            datumCsysBuilder1.Destroy()
        Next


    End Sub
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
    Public Class BodyCoordinate
        Public Property Origin As NXOpen.Point3d
        Public Property Orientation As NXOpen.Matrix3x3

        Public Sub New(origin As NXOpen.Point3d, orientation As NXOpen.Matrix3x3)
            Me.Origin = origin
            Me.Orientation = orientation
        End Sub
    End Class

    Public Function FindCentroidOfFace(ByVal face As NXOpen.Face) As BodyCoordinate

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

        ufs.Modl.AskBoundingBoxExact(face.Tag, csys, minCorner, directions, distances)

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

        Return New BodyCoordinate(absoluteCentroid, convertToMatrix3x3(matrix_v_1))

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
                Case 1
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