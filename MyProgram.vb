
Option Strict Off
Imports System
Imports NXOpen
Imports NXOpen.BlockStyler
Imports NXOpen.UF
Imports PLMComponents.Parasolid.PK_.Unsafe
Imports System.Collections
Imports System.Collections.Generic
Imports PLMComponents.Parasolid.PK_DEBUG_.Unsafe

'------------------------------------------------------------------------------
'Represents Block Styler application class
'------------------------------------------------------------------------------
Public Class BODY_MOVE
    'class members
    Private Shared theSession As NXOpen.Session
    Private Shared theUI As UI
    Private theDlxFileName As String
    Private theDialog As NXOpen.BlockStyler.BlockDialog
    Private group0 As NXOpen.BlockStyler.Group ' Block type: Group
    Private bodySelect0 As NXOpen.BlockStyler.BodyCollector ' Block type: Body Collector
    Private bodySelect01 As NXOpen.BlockStyler.BodyCollector ' Block type: Body Collector
    Private enum0 As NXOpen.BlockStyler.Enumeration ' Block type: Enumeration
    Private string0 As NXOpen.BlockStyler.StringBlock ' Block type: String
    '------------------------------------------------------------------------------
    'Bit Option for Property: EntityType
    '------------------------------------------------------------------------------
    Public Shared ReadOnly EntityType_AllowBodies As Integer = 64
    '------------------------------------------------------------------------------
    'Bit Option for Property: BodyRules
    '------------------------------------------------------------------------------
    Public Shared ReadOnly BodyRules_SingleBody As Integer = 1
    Public Shared ReadOnly BodyRules_FeatureBodies As Integer = 2
    Public Shared ReadOnly BodyRules_BodiesinGroup As Integer = 4

    '------------------------------------------------------------------------------
    '
    '------------------------------------------------------------------------------
    Dim ufs As UFSession = UFSession.GetUFSession()
    Dim ui As UI = UI.GetUI()

    Public Modl As UFModl

    Dim list_body As New List(Of NXOpen.Body)
    Dim a_body As NXOpen.Body
    Dim att_gt, att_tile, what_do As String
    Dim layer_old As Integer


#Region "Block Styler Dialog Designer generator code"
    '------------------------------------------------------------------------------
    'Constructor for NX Styler class
    '------------------------------------------------------------------------------
    Public Sub New()
        Try

            theSession = NXOpen.Session.GetSession()
            theUI = UI.GetUI()
            theDlxFileName = "E:\Code\BODY_MOVE.dlx"
            theDialog = theUI.CreateDialog(theDlxFileName)
            theDialog.AddApplyHandler(AddressOf apply_cb)
            theDialog.AddOkHandler(AddressOf ok_cb)
            theDialog.AddUpdateHandler(AddressOf update_cb)
            theDialog.AddInitializeHandler(AddressOf initialize_cb)
            theDialog.AddDialogShownHandler(AddressOf dialogShown_cb)


        Catch ex As Exception

            '---- Enter your exception handling code here -----
            Throw ex
        End Try
    End Sub
#End Region


    Public Shared Sub Main()
        Dim theBODY_MOVE As BODY_MOVE = Nothing
        Try

            theBODY_MOVE = New BODY_MOVE()
            ' The following method shows the dialog immediately
            theBODY_MOVE.Show()

        Catch ex As Exception

            '---- Enter your exception handling code here -----
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString)
        Finally
            If theBODY_MOVE IsNot Nothing Then
                theBODY_MOVE.Dispose()
                theBODY_MOVE = Nothing
            End If
        End Try
    End Sub

    Public Shared Sub UnloadLibrary(ByVal arg As String)
        Try


        Catch ex As Exception

            '---- Enter your exception handling code here -----
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString)
        End Try
    End Sub

    '------------------------------------------------------------------------------
    'This method shows the dialog on the screen
    '------------------------------------------------------------------------------
    Public Sub Show()
        Try

            theDialog.Show()

        Catch ex As Exception

            '---- Enter your exception handling code here -----
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString)
        End Try
    End Sub

    '------------------------------------------------------------------------------
    'Method Name: Dispose
    '------------------------------------------------------------------------------
    Public Sub Dispose()
        If theDialog IsNot Nothing Then
            theDialog.Dispose()
            theDialog = Nothing
        End If
    End Sub

    '------------------------------------------------------------------------------
    '---------------------Block UI Styler Callback Functions--------------------------
    '------------------------------------------------------------------------------

    '------------------------------------------------------------------------------
    'Callback Name: initialize_cb
    '------------------------------------------------------------------------------
    Public Sub initialize_cb()
        Try

            group0 = CType(theDialog.TopBlock.FindBlock("group0"), NXOpen.BlockStyler.Group)
            bodySelect0 = CType(theDialog.TopBlock.FindBlock("bodySelect0"), NXOpen.BlockStyler.BodyCollector)
            bodySelect01 = CType(theDialog.TopBlock.FindBlock("bodySelect01"), NXOpen.BlockStyler.BodyCollector)
            enum0 = CType(theDialog.TopBlock.FindBlock("enum0"), NXOpen.BlockStyler.Enumeration)
            string0 = CType(theDialog.TopBlock.FindBlock("string0"), NXOpen.BlockStyler.StringBlock)

        Catch ex As Exception

            '---- Enter your exception handling code here -----
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString)
        End Try


    End Sub

    '------------------------------------------------------------------------------
    'Callback Name: dialogShown_cb
    'This callback is executed just before the dialog launch. Thus any value set 
    'here will take precedence and dialog will be launched showing that value. 
    '------------------------------------------------------------------------------
    Public Sub dialogShown_cb()
        Try


            If enum0.ValueAsString = "画層移動" Then

                string0.Show = True

            Else

                string0.Show = False

            End If

        Catch ex As Exception

            '---- Enter your exception handling code here -----
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString)
        End Try
    End Sub

    '------------------------------------------------------------------------------
    'Callback Name: apply_cb
    '------------------------------------------------------------------------------
    Public Function apply_cb() As Integer
        Dim errorCode As Integer = 0
        Try

            what_do = enum0.ValueAsString

            Try
                layer_old = CDbl(string0.Value)
            Catch ex As Exception
                layer_old = 0
            End Try

            Dim lw As ListingWindow = theSession.ListingWindow

            Dim xx() As TaggedObject = bodySelect0.GetSelectedObjects

            a_body = xx(0)

            Dim xy() As TaggedObject = bodySelect01.GetSelectedObjects

            For Each x As TaggedObject In xy

                'lw.WriteLine("aa:   " & x.GetType.ToString)

                list_body.Add(x)

            Next

            find_same_body(a_body, list_body, what_do, layer_old)


        Catch ex As Exception

            '---- Enter your exception handling code here -----
            errorCode = 1
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString)
        End Try
        apply_cb = errorCode
    End Function

    '------------------------------------------------------------------------------
    'Callback Name: update_cb
    '------------------------------------------------------------------------------
    Public Function update_cb(ByVal block As NXOpen.BlockStyler.UIBlock) As Integer
        Try

            If block Is bodySelect0 Then


            ElseIf block Is bodySelect01 Then


            ElseIf block Is enum0 Then

                If enum0.ValueAsString = "画層移動" Then

                    string0.Show = True

                Else

                    string0.Show = False

                End If

                what_do = enum0.ValueAsString

            ElseIf block Is string0 Then

                Try
                    layer_old = CDbl(string0.Value)
                Catch ex As Exception
                    layer_old = 0
                End Try


            End If

        Catch ex As Exception

            '---- Enter your exception handling code here -----
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString)
        End Try
        update_cb = 0
    End Function

    '------------------------------------------------------------------------------
    'Callback Name: ok_cb
    '------------------------------------------------------------------------------
    Public Function ok_cb() As Integer
        Dim errorCode As Integer = 0
        Try

            '---- Enter your callback code here -----
            errorCode = apply_cb()

        Catch ex As Exception

            '---- Enter your exception handling code here -----
            errorCode = 1
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString)
        End Try
        ok_cb = errorCode
    End Function

    '------------------------------------------------------------------------------
    'Function Name: GetBlockProperties
    'Returns the propertylist of the specified BlockID
    '------------------------------------------------------------------------------
    Public Function GetBlockProperties(ByVal blockID As String) As PropertyList
        GetBlockProperties = Nothing
        Try

            GetBlockProperties = theDialog.GetBlockProperties(blockID)

        Catch ex As Exception

            '---- Enter your exception handling code here -----
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString)
        End Try
    End Function

    Function find_same_body(a_body_1 As NXOpen.Body, list_body_1 As List(Of NXOpen.Body), what_do_1 As String, layer_old_body As Integer)

        Dim workpart As NXOpen.Part = theSession.Parts.Work
        Dim dispPart As NXOpen.Part = theSession.Parts.Display

        Dim list_st_p As New List(Of NXOpen.Point3d)   '''' list point of face with Csys on center face
        Dim list_body As New List(Of NXOpen.Body)
        Dim list_face As New List(Of NXOpen.Face)
        Dim area_body_1 As Double

        workpart.WCS.Visibility = False

        Dim matrix_a As NXOpen.Matrix3x3
        Dim origin_a As NXOpen.Point3d
        'Dim matrix_a1 As NXOpen.NXMatrix
        origin_a = workpart.WCS.CoordinateSystem.Origin
        'matrix_a1 = workpart.WCS.CoordinateSystem.Orientation
        matrix_a = workpart.WCS.CoordinateSystem.Orientation.Element
        'NXMatrixCollection.Create(NXOpen.Matrix3x3)

        'maxtrix and point (CSYS) SLECTED Body 
        Dim matrix_slect As NXOpen.Matrix3x3
        Dim origin_slect As NXOpen.Point3d

        Dim lw As ListingWindow = theSession.ListingWindow

        Dim selectedbody1 As NXOpen.Body = a_body_1 'Nothing

        Dim the_bodies As List(Of NXOpen.Body) = list_body_1 'Nothing


        '''''''--------------------------------------------------------------------------------------------------------
        '''''''                     ADD OR SUB A SELECTED BODY OF LIST BODY START
        '''''''--------------------------------------------------------------------------------------------------------

        If the_bodies.Contains(selectedbody1) = True Then

            the_bodies.Remove(selectedbody1)

        End If

        '''''''--------------------------------------------------------------------------------------------------------
        '''''''                     ADD OR SUB A SELECTED BODY OF LIST BODY END
        '''''''--------------------------------------------------------------------------------------------------------


        area_body_1 = AREA_BODY(selectedbody1)

        list_body.Add(selectedbody1)

        Dim selectedFace1 As NXOpen.Face = Largest_planar_Face(selectedbody1)

        list_face.Add(selectedFace1)


        Dim faceFeat1 As Features.Feature
        Dim faceFeatTags1() As Tag = Nothing
        ufs.Modl.AskFaceFeats(selectedFace1.Tag, faceFeatTags1)
        faceFeat1 = Utilities.NXObjectManager.Get(faceFeatTags1(0))


        Dim list_body_Part As New List(Of NXOpen.Body)

        list_body_Part = AskAll_visible(workpart)

        list_body_Part.Remove(selectedbody1)

        For Each selectedbody2 As NXOpen.Body In list_body_Part

            Dim area_1 As Double = AREA_BODY(selectedbody2)
            'lw.WriteLine("      body_selected_Area: " & area_1)

            If Math.Abs(area_body_1 - area_1) <= 0.001 Then

                list_body.Add(selectedbody2)

                Dim selectedFace2 As NXOpen.Face = Largest_planar_Face(selectedbody2)

                list_face.Add(selectedFace2)

            End If

        Next


        Dim list_body_final As New List(Of NXOpen.Body)

        Dim list_origin_body As New List(Of NXOpen.Point3d)

        Dim list_matrix_body As New List(Of NXOpen.Matrix3x3)

        For j As Integer = 0 To list_body.Count - 1

            Dim s_body As NXOpen.Body = list_body(j)
            Dim s_face As NXOpen.Face = list_face(j)

            Dim cp(2) As Double
            Dim pt(2) As Double
            Dim u1(2) As Double
            Dim v1(2) As Double
            Dim u2(2) As Double
            Dim v2(2) As Double
            Dim norm(2) As Double
            Dim radii(1) As Double
            Dim param(1) As Double

            ufs.Modl.AskFaceProps(s_face.Tag, param, pt, u1, v1, u2, v2, norm, radii)

            Dim s_face_edges1() As NXOpen.Edge = s_face.GetEdges

            Dim d_edge_i As Double = 0
            Dim r_edgen_i As Double = 0

            Dim d_curve_1(2) As Double
            Dim origin1(2) As Double
            Dim origin1_p As NXOpen.Point3d

            Dim d_curve_c(2) As Double
            Dim origin1_c(2) As Double
            Dim origin1_p_c As NXOpen.Point3d

            Dim cylinder_TF As Boolean = True


            For Each e1 As NXOpen.Edge In s_face_edges1

                Dim parm As Double
                Dim curve_pnt_1(2) As Double
                Dim tangent(2) As Double
                Dim p_norm(2) As Double
                Dim b_norm(2) As Double
                Dim torsion As Double
                Dim radOfCurvature As Double

                ufs.Modl.AskCurveProps(e1.Tag, parm, curve_pnt_1, tangent, p_norm, b_norm, torsion, radOfCurvature)

                If e1.SolidEdgeType = NXOpen.Edge.EdgeType.Linear And d_edge_i < e1.GetLength() Then


                    d_edge_i = e1.GetLength()

                    d_curve_1(0) = tangent(0) / d_edge_i
                    d_curve_1(1) = tangent(1) / d_edge_i
                    d_curve_1(2) = tangent(2) / d_edge_i

                    origin1_p = New NXOpen.Point3d(curve_pnt_1(0), curve_pnt_1(1), curve_pnt_1(2))

                    origin1(0) = curve_pnt_1(0)
                    origin1(1) = curve_pnt_1(1)
                    origin1(2) = curve_pnt_1(2)

                    cylinder_TF = False

                ElseIf e1.SolidEdgeType = NXOpen.Edge.EdgeType.Circular Then


                    Dim edgeEvaluator As System.IntPtr
                    ufs.Eval.Initialize(e1.Tag, edgeEvaluator)
                    Dim arcInfor1 As UFEval.Arc

                    ufs.Eval.AskArc(edgeEvaluator, arcInfor1)

                    Dim cent_p(2), st_p(2), en_p(2) As Double

                    cent_p = arcInfor1.center


                    If r_edgen_i > 0 Then

                        Dim d As Double = Math.Sqrt((origin1_p_c.X - cent_p(0)) * (origin1_p_c.X - cent_p(0)) + (origin1_p_c.Y - cent_p(1)) * (origin1_p_c.Y - cent_p(1)) +
                                                    (origin1_p_c.Z - cent_p(2)) * (origin1_p_c.Z - cent_p(2)))

                        If d > 0.001 Then

                            d_curve_c(0) = (origin1_p_c.X - cent_p(0)) / d

                            d_curve_c(1) = (origin1_p_c.Y - cent_p(1)) / d

                            d_curve_c(2) = (origin1_p_c.Z - cent_p(2)) / d

                            cylinder_TF = False

                        End If

                    End If


                    If r_edgen_i < radOfCurvature Then

                        r_edgen_i = radOfCurvature

                        origin1_p_c.X = cent_p(0)

                        origin1_p_c.Y = cent_p(1)

                        origin1_p_c.Z = cent_p(2)

                    End If


                End If

            Next

            If cylinder_TF = True Then

                Dim s_body_edges2() As NXOpen.Edge = s_body.GetEdges

                Dim test As Boolean = False

                Dim d As Double = 0

                For Each e1 As NXOpen.Edge In s_body_edges2

                    If e1.SolidEdgeType = NXOpen.Edge.EdgeType.Linear Then

                        Dim point_1(2), point_2(2), vector_x(2) As Double
                        Dim p_count As Integer

                        ufs.Modl.AskEdgeVerts(e1.Tag, point_1, point_2, p_count)

                        Dim l As Double = e1.GetLength

                        vector_x(0) = (point_1(0) - point_2(0)) / l
                        vector_x(1) = (point_1(1) - point_2(1)) / l
                        vector_x(2) = (point_1(2) - point_2(2)) / l

                        If d < l And (vector_x(0) * norm(0) + vector_x(1) * norm(1) + vector_x(2) * norm(2)) < 0.001 Then

                            d = l

                            d_curve_c(0) = vector_x(0)
                            d_curve_c(1) = vector_x(1)
                            d_curve_c(2) = vector_x(2)

                        End If

                    End If

                    If e1.SolidEdgeType <> NXOpen.Edge.EdgeType.Circular Then

                        test = True

                    Else

                    End If

                Next


                If test = False Then

                    ufs.Vec3.AskPerpendicular(norm, d_curve_c)

                Else

                    If d = 0 Then

                        Dim myFaces() As NXOpen.Face
                        myFaces = s_body.GetFaces
                        Dim r As Double = 0

                        For Each tempFace As NXOpen.Face In myFaces

                            If tempFace.SolidFaceType = NXOpen.Face.FaceType.Cylindrical Then

                                Dim faceType As Integer
                                Dim facePt(2) As Double
                                Dim faceDir(2) As Double
                                Dim bbox(5) As Double
                                Dim faceRadius As Double
                                Dim faceRadData As Double
                                Dim normDirection As Integer

                                ufs.Modl.AskFaceData(tempFace.Tag, faceType, facePt, faceDir, bbox, faceRadius, faceRadData, normDirection)

                                If Math.Abs((faceDir(0) * norm(0) + faceDir(1) * norm(1) + faceDir(2) * norm(2))) < 0.01 And r < faceRadius Then

                                    r = faceRadius

                                    d_curve_c(0) = faceDir(0)
                                    d_curve_c(1) = faceDir(1)
                                    d_curve_c(2) = faceDir(2)

                                End If

                            End If

                        Next

                    End If

                End If

            End If

            If d_curve_1(0) = 0 And d_curve_1(1) = 0 And d_curve_1(2) = 0 Then

                d_curve_1(0) = d_curve_c(0)
                d_curve_1(1) = d_curve_c(1)
                d_curve_1(2) = d_curve_c(2)

            End If

            Dim mtx As NXOpen.Tag = NXOpen.Tag.Null
            Dim newCsys As NXOpen.Tag = NXOpen.Tag.Null
            Dim matrixValues(8) As Double

            matrixValues(0) = d_curve_1(1) * norm(2) - d_curve_1(2) * norm(1)
            matrixValues(1) = -d_curve_1(0) * norm(2) + d_curve_1(2) * norm(0)
            matrixValues(2) = d_curve_1(0) * norm(1) - d_curve_1(1) * norm(0)
            matrixValues(3) = d_curve_1(0)
            matrixValues(4) = d_curve_1(1)
            matrixValues(5) = d_curve_1(2)
            matrixValues(6) = norm(0)
            matrixValues(7) = norm(1)
            matrixValues(8) = norm(2)


            Dim matrix_v_1(8) As Double

            Dim xVector(2) As Double

            xVector(0) = matrixValues(0)
            xVector(1) = matrixValues(1)
            xVector(2) = matrixValues(2)

            Dim yVector(2) As Double

            yVector(0) = matrixValues(3)
            yVector(1) = matrixValues(4)
            yVector(2) = matrixValues(5)


            ufs.Mtx3.Initialize(xVector, yVector, matrix_v_1)


            Dim csys As NXOpen.Tag = NXOpen.Tag.Null
            Dim min_corner(2) As Double
            Dim directions(2, 2) As Double
            Dim distances(2) As Double
            Dim edge_len(2) As Double


            ufs.Csys.CreateMatrix(matrix_v_1, mtx)
            ufs.Csys.CreateTempCsys(origin1, mtx, csys)

            ufs.Modl.AskBoundingBoxExact(s_body.Tag, csys, min_corner, directions,
                        distances)


            edge_len(0) = distances(0).ToString()
            edge_len(1) = distances(1).ToString()
            edge_len(2) = distances(2).ToString()


            Dim origin2(2) As Double

            origin2(0) = min_corner(0)
            origin2(1) = min_corner(1)
            origin2(2) = min_corner(2)

            For k As Integer = 0 To 2


                origin2(0) = origin2(0) + edge_len(k) / 2 * directions(k, 0)
                origin2(1) = origin2(1) + edge_len(k) / 2 * directions(k, 1)
                origin2(2) = origin2(2) + edge_len(k) / 2 * directions(k, 2)

            Next


            Dim p_1 As NXOpen.Point3d

            p_1.X = origin2(0)
            p_1.Y = origin2(1)
            p_1.Z = origin2(2)

            Dim matrix1 As NXOpen.Matrix3x3
            matrix1.Xx = directions(0, 0)
            matrix1.Xy = directions(0, 1)
            matrix1.Xz = directions(0, 2)
            matrix1.Yx = directions(1, 0)
            matrix1.Yy = directions(1, 1)
            matrix1.Yz = directions(1, 2)
            matrix1.Zx = directions(2, 0)
            matrix1.Zy = directions(2, 1)
            matrix1.Zz = directions(2, 2)

            workpart.WCS.SetOriginAndMatrix(p_1, matrix1)

            Dim s_body_edges1() As NXOpen.Edge = s_body.GetEdges

            Dim list_point As New List(Of Point3d)

            Dim list_cp_p As New List(Of NXOpen.Point3d)

            For Each e1 As NXOpen.Edge In s_body_edges1

                Dim parm As Integer
                Dim curve_pnt_1(2) As Double
                Dim tangent(2) As Double
                Dim p_norm(2) As Double
                Dim b_norm(2) As Double
                Dim torsion As Double
                Dim radOfCurvature As Double

                For parm = 0 To 1

                    ufs.Modl.AskCurveProps(e1.Tag, parm, curve_pnt_1, tangent, p_norm, b_norm, torsion, radOfCurvature)

                    If e1.SolidEdgeType = NXOpen.Edge.EdgeType.Circular Then

                        Dim edgeEvaluator As System.IntPtr
                        ufs.Eval.Initialize(e1.Tag, edgeEvaluator)
                        Dim arcInfor1 As UFEval.Arc

                        ufs.Eval.AskArc(edgeEvaluator, arcInfor1)

                        Dim cent_p(2), st_p(2), en_p(2) As Double

                        cent_p = arcInfor1.center

                        Dim pl_1 As NXOpen.Point3d

                        pl_1.X = cent_p(0)
                        pl_1.Y = cent_p(1)
                        pl_1.Z = cent_p(2)

                        If contain_double(list_point, pl_1, 0.001) = False Then

                            list_point.Add(pl_1)
                            'Exit For
                        End If

                    Else

                        Dim pl_1 As NXOpen.Point3d

                        pl_1.X = curve_pnt_1(0)
                        pl_1.Y = curve_pnt_1(1)
                        pl_1.Z = curve_pnt_1(2)

                        If contain_double(list_point, pl_1, 0.001) = False Then

                            list_point.Add(pl_1)
                            'Exit For
                        End If

                    End If
                Next
            Next
            'lw.WriteLine("**********************************************")
            ' lw.WriteLine("so diem list :    " & list_point.count)
            For Each poit_i As NXOpen.Point3d In list_point

                Dim p_face As NXOpen.Point3d

                p_face = Abs2WCS(poit_i)

                If j = 0 Then

                    list_st_p.Add(p_face)
                    'lw.WriteLine("DIEM GOC:    " & list_point.IndexOf(poit_i) & Math.Round(p_face.X, 3) & "," & Math.Round(p_face.y, 3) & "," & Math.Round(p_face.z, 3))
                    'lw.WriteLine(Math.Round(p_face.X, 3) & "," & Math.Round(p_face.y, 3) & "," & Math.Round(p_face.z, 3))
                Else

                    list_cp_p.Add(p_face)

                    'lw.WriteLine("DIEM SO " & j & "thu: " & list_point.IndexOf(poit_i) & " :    " & Math.Round(p_face.X, 3) & "," & Math.Round(p_face.y, 3) & "," & Math.Round(p_face.z, 3))
                    ' lw.WriteLine(Math.Round(p_face.X, 3) & "," & Math.Round(p_face.y, 3) & "," & Math.Round(p_face.z, 3))

                End If

            Next

            Dim test1, test2 As New List(Of Boolean)
            Dim test1_v, test2_v As New List(Of Boolean)


            If j = 0 Then

                origin_slect = workpart.WCS.CoordinateSystem.Origin

                matrix_slect = workpart.WCS.CoordinateSystem.Orientation.Element

                GoTo boqua

            ElseIf list_st_p.Count = list_cp_p.Count And j > 0 Then

                For Each t_p As NXOpen.Point3d In list_st_p


                    For Each t_p_1 As NXOpen.Point3d In list_cp_p



                        Dim a, b, c, a1, b1, c1 As Double

                        a = Math.Abs(t_p.X - t_p_1.X)
                        b = Math.Abs(t_p.Y - t_p_1.Y)
                        c = Math.Abs(t_p.Z - t_p_1.Z)

                        a1 = Math.Abs(t_p.X + t_p_1.X)
                        b1 = Math.Abs(t_p.Y + t_p_1.Y)
                        c1 = Math.Abs(t_p.Z - t_p_1.Z)

                        If a < 0.001 And b < 0.001 And c < 0.001 Then

                            test1.Add(True)

                        ElseIf a1 < 0.001 And b1 < 0.001 And c1 < 0.001 Then

                            test2.Add(True)

                        End If

                        If Math.Abs(edge_len(0) - edge_len(1)) <= 0.001 Then

                            If Math.Abs(t_p.X - t_p_1.Y) <= 0.001 And Math.Abs(t_p.Y + t_p_1.X) <= 0.001 And Math.Abs(t_p.Z - t_p_1.Z) <= 0.001 Then

                                test1_v.Add(True)

                            ElseIf Math.Abs(t_p.X + t_p_1.Y) <= 0.001 And Math.Abs(t_p.Y - t_p_1.X) <= 0.001 And Math.Abs(t_p.Z - t_p_1.Z) <= 0.001 Then

                                test2_v.Add(True)

                            End If

                        End If

                    Next

                Next

            Else

                GoTo boqua

            End If

            '' SAVE WCS
            ' Dim cartesianCoordinateSystem1 As NXOpen.CartesianCoordinateSystem = Nothing
            ' cartesianCoordinateSystem1 = workpart.WCS.Save()

            If test1.Count = list_st_p.Count Then

                list_body_final.Add(s_body)

                list_origin_body.Add(workpart.WCS.CoordinateSystem.Origin)

                list_matrix_body.Add(workpart.WCS.CoordinateSystem.Orientation.Element)

            ElseIf test2.Count = list_st_p.Count Then

                list_body_final.Add(s_body)

                list_origin_body.Add(workpart.WCS.CoordinateSystem.Origin)

                workpart.WCS.Rotate(WCS.Axis.ZAxis, 180)

                list_matrix_body.Add(workpart.WCS.CoordinateSystem.Orientation.Element)

            Else

                If Math.Abs(edge_len(0) - edge_len(1)) <= 0.001 Then

                    If test1_v.Count = list_st_p.Count Then

                        list_body_final.Add(s_body)

                        list_origin_body.Add(workpart.WCS.CoordinateSystem.Origin)

                        workpart.WCS.Rotate(WCS.Axis.ZAxis, -90)

                        list_matrix_body.Add(workpart.WCS.CoordinateSystem.Orientation.Element)

                    ElseIf test2_v.Count = list_st_p.Count Then

                        list_body_final.Add(s_body)

                        list_origin_body.Add(workpart.WCS.CoordinateSystem.Origin)

                        workpart.WCS.Rotate(WCS.Axis.ZAxis, 90)

                        list_matrix_body.Add(workpart.WCS.CoordinateSystem.Orientation.Element)

                    Else

                    End If

                Else

                End If

            End If

            list_cp_p.Clear()

boqua:

        Next


        workpart.WCS.SetOriginAndMatrix(origin_slect, matrix_slect)

        For i As Integer = 0 To list_matrix_body.Count - 1

            Dim matrix_i As NXOpen.Matrix3x3 = list_matrix_body(i)

            Dim point_i As NXOpen.Point3d = list_origin_body(i)

            move_object(the_bodies, matrix_slect, origin_slect, matrix_i, point_i)

            If what_do_1 = "消す" Then

                Dim body1 As NXOpen.Body = list_body_final(i)
                ufs.Obj.DeleteObject(body1.Tag)

            ElseIf what_do_1 = "画層移動" Then

                Dim displayModification1 As NXOpen.DisplayModification
                displayModification1 = theSession.DisplayManager.NewDisplayModification()

                displayModification1.ApplyToAllFaces = True

                displayModification1.ApplyToOwningParts = False

                displayModification1.NewTranslucency = 100

                Dim objects1(0) As NXOpen.DisplayableObject
                Dim body1 As NXOpen.Body = list_body_final(i)

                objects1(0) = body1
                displayModification1.Apply(objects1)

                displayModification1.Dispose()

                body1.Color = 69
                body1.RedisplayObject()

                workpart.Layers.MoveDisplayableObjects(layer_old, objects1)

            ElseIf what_do_1 = "そのまま" Then

            End If


        Next

        Dim stateArray1(0) As NXOpen.Layer.StateInfo
        stateArray1(0) = New NXOpen.Layer.StateInfo(250, NXOpen.Layer.State.Hidden)
        workpart.Layers.ChangeStates(stateArray1, False)

ket_thuc:

        workpart.WCS.SetOriginAndMatrix(origin_a, matrix_a)
        workpart.WCS.Visibility = True

    End Function

    Function AREA_BODY(ByVal theBody As NXOpen.Body) As Double
        Const undoMarkName As String = "NXJ journal"
        Dim markId1 As NXOpen.Session.UndoMarkId
        markId1 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, undoMarkName)

        Dim theBodies(0) As NXOpen.Body
        theBodies(0) = theBody

        Dim myMeasure As MeasureManager = theSession.Parts.Display.MeasureManager()
        Dim massUnits(4) As Unit
        massUnits(0) = theSession.Parts.Display.UnitCollection.GetBase("Area")
        massUnits(1) = theSession.Parts.Display.UnitCollection.GetBase("Volume")
        massUnits(2) = theSession.Parts.Display.UnitCollection.GetBase("Mass")
        massUnits(3) = theSession.Parts.Display.UnitCollection.GetBase("Length")

        Dim mb As MeasureBodies = Nothing
        mb = myMeasure.NewMassProperties(massUnits, 0.99, theBodies)
        mb.InformationUnit = MeasureBodies.AnalysisUnit.GramMillimeter

        AREA_BODY = mb.Area

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

    Function AskAllBodies_visible(ByVal the_part As NXOpen.Part) As List(Of NXOpen.Body)

        Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
        'Dim workpart As NXOpen.Part = theSession.Parts.Work
        'Dim dispPart As NXOpen.Part = theSession.Parts.Display

        'Dim lw As ListingWindow = theSession.ListingWindow
        'lw.Open()

        Dim visibleObjects() As DisplayableObject

        visibleObjects = the_part.Views.WorkView.AskVisibleObjects()

        Dim list_body As New List(Of NXOpen.Body)

        For Each displayableObject As DisplayableObject In visibleObjects
            If TypeOf displayableObject Is Body Then

                Dim body_1 As Body = displayableObject

                list_body.Add(body_1)

                ' lw.WriteLine("Body: " & body_1.JournalIdentifier)
                ' body_1.Highlight()
                'Do other processing
            End If
        Next

        AskAllBodies_visible = list_body

    End Function


    Function AskAllBodies(ByVal thePart As NXOpen.Part) As List(Of NXOpen.Body)
        Dim thebodies As New List(Of NXOpen.Body)
        Try
            Dim BodyTag As Tag = NXOpen.Tag.Null
            Do
                ufs.Obj.CycleObjsInPart(thePart.Tag, UFConstants.UF_solid_type, BodyTag)
                If BodyTag = NXOpen.Tag.Null Then
                    Exit Do
                End If
                Dim theType As Integer, theSubtype As Integer
                ufs.Obj.AskTypeAndSubtype(BodyTag, theType, theSubtype)
                If theSubtype = UFConstants.UF_solid_body_subtype Then
                    thebodies.Add(Utilities.NXObjectManager.Get(BodyTag))
                End If
            Loop While True
        Catch ex As NXException
            'lw.WriteLine(ex.ErrorCode & ex.Message)
        End Try
        Return thebodies
    End Function

    Public Function Abs2WCS(ByVal inPt As Point3d) As Point3d
        Dim pt1(2), pt2(2) As Double

        pt1(0) = inPt.X
        pt1(1) = inPt.Y
        pt1(2) = inPt.Z

        ufs.Csys.MapPoint(UFConstants.UF_CSYS_ROOT_COORDS, pt1, UFConstants.UF_CSYS_ROOT_WCS_COORDS, pt2)

        Abs2WCS.X = pt2(0)
        Abs2WCS.Y = pt2(1)
        Abs2WCS.Z = pt2(2)

    End Function

    Function WCS2Abs(ByVal inPt As Point3d) As Point3d
        Dim pt1(2), pt2(2) As Double

        pt1(0) = inPt.X
        pt1(1) = inPt.Y
        pt1(2) = inPt.Z

        ufs.Csys.MapPoint(UFConstants.UF_CSYS_ROOT_WCS_COORDS, pt1,
            UFConstants.UF_CSYS_ROOT_COORDS, pt2)

        WCS2Abs.X = pt2(0)
        WCS2Abs.Y = pt2(1)
        WCS2Abs.Z = pt2(2)

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

    Function move_object(ByVal selected_body As List(Of NXOpen.Body), ByVal maxtrix_body As NXOpen.Matrix3x3, point_body As NXOpen.Point3d, ByVal maxtrix_body_1 As NXOpen.Matrix3x3, point_body_1 As NXOpen.Point3d)

        Dim workpart As NXOpen.Part = theSession.Parts.Work
        Dim dispPart As NXOpen.Part = theSession.Parts.Display

        Dim direction1, direction2, direction3, direction4 As NXOpen.Vector3d

        direction1.X = maxtrix_body.Xx
        direction1.Y = maxtrix_body.Xy
        direction1.Z = maxtrix_body.Xz

        direction2.X = maxtrix_body.Yx
        direction2.Y = maxtrix_body.Yy
        direction2.Z = maxtrix_body.Yz

        direction3.X = maxtrix_body_1.Xx
        direction3.Y = maxtrix_body_1.Xy
        direction3.Z = maxtrix_body_1.Xz

        direction4.X = maxtrix_body_1.Yx
        direction4.Y = maxtrix_body_1.Yy
        direction4.Z = maxtrix_body_1.Yz


        Dim xform1 As NXOpen.Xform
        xform1 = workpart.Xforms.CreateXform(point_body, direction1, direction2, NXOpen.SmartObject.UpdateOption.WithinModeling, 1.0)

        Dim cartesianCoordinateSystem1 As NXOpen.CartesianCoordinateSystem
        cartesianCoordinateSystem1 = workpart.CoordinateSystems.CreateCoordinateSystem(xform1, NXOpen.SmartObject.UpdateOption.WithinModeling)

        Dim xform2 As NXOpen.Xform
        xform2 = workpart.Xforms.CreateXform(point_body_1, direction3, direction4, NXOpen.SmartObject.UpdateOption.WithinModeling, 1.0)

        Dim cartesianCoordinateSystem2 As NXOpen.CartesianCoordinateSystem
        cartesianCoordinateSystem2 = workpart.CoordinateSystems.CreateCoordinateSystem(xform2, NXOpen.SmartObject.UpdateOption.WithinModeling)

        Dim nullNXOpen_Features_MoveObject As NXOpen.Features.MoveObject = Nothing

        Dim moveObjectBuilder1 As NXOpen.Features.MoveObjectBuilder
        moveObjectBuilder1 = workpart.BaseFeatures.CreateMoveObjectBuilder(nullNXOpen_Features_MoveObject)

        Dim added1 As Boolean
        added1 = moveObjectBuilder1.ObjectToMoveObject.Add(selected_body.ToArray)

        moveObjectBuilder1.TransformMotion.DeltaEnum = NXOpen.GeometricUtilities.ModlMotion.Delta.ReferenceWcsWorkPart

        moveObjectBuilder1.TransformMotion.Option = NXOpen.GeometricUtilities.ModlMotion.Options.CsysToCsys

        moveObjectBuilder1.MoveObjectResult = NXOpen.Features.MoveObjectBuilder.MoveObjectResultOptions.CopyOriginal

        moveObjectBuilder1.TransformMotion.FromCsys = cartesianCoordinateSystem1

        moveObjectBuilder1.TransformMotion.ToCsys = cartesianCoordinateSystem2

        Dim nXObject1 As NXOpen.NXObject
        nXObject1 = moveObjectBuilder1.Commit()

        Dim objects2() As NXOpen.NXObject
        objects2 = moveObjectBuilder1.GetCommittedObjects()

        moveObjectBuilder1.Destroy()


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
    Function contain_double(list_point As List(Of NXOpen.Point3d), test_point As NXOpen.Point3d, a As Double) As Boolean
        For Each p As NXOpen.Point3d In list_point
            If Math.Abs(p.X - test_point.X) < a AndAlso
           Math.Abs(p.Y - test_point.Y) < a AndAlso
           Math.Abs(p.Z - test_point.Z) < a Then
                Return True
            End If
        Next
        Return False
    End Function


    '**********************************************************
    Public Function GetUnloadOption(ByVal dummy As String) As Integer
        'Return Session.LibraryUnloadOption.Immediately
        Return CType(NXOpen.Session.LibraryUnloadOption.Immediately, Integer)
    End Function
    '**********************************************************

End Class

