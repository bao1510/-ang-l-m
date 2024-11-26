Option Strict Off
Imports System
Imports NXOpen
Imports NXOpen.Features
Imports NXOpen.UF

Module Start_End_Hole_Position

    Sub Main()

        Dim theSession As Session = Session.GetSession()
        Dim workPart As Part = theSession.Parts.Work
        Dim displayPart As Part = theSession.Parts.Display
        Dim theUFSession As UFSession = UFSession.GetUFSession
        Dim lw As ListingWindow = theSession.ListingWindow

        Dim featArray() As Feature = workPart.Features.GetFeatures()

        Dim markId1 As Session.UndoMarkId
        markId1 = theSession.SetUndoMark(Session.MarkVisibility.Visible, "Start / End Hole Position")

        theUFSession.Ui.ExitListingWindow()
        lw.Open()

        For Each myFeature As Feature In featArray

            If myFeature.FeatureType = "HOLE PACKAGE" Then

                Dim holeFeature As HolePackage = myFeature
                'lw.writeline(myFeature.GetFeatureName)

                Dim holeDir() As Vector3d
                holeFeature.GetDirections(holeDir)

                Dim holeBuilder As HolePackageBuilder
                holeBuilder = workPart.Features.CreateHolePackageBuilder(holeFeature)

                'lw.WriteLine("hole type: " & holeBuilder.Type.ToString)

                Dim FaceType As Integer
                Dim FacePoint(2) As Double
                Dim FaceDir(2) As Double
                Dim FaceBox(5) As Double
                Dim FaceRadius As Double
                Dim FaceRad_data As Double
                Dim FaceNorm_dir As Integer

                Dim csys As NXOpen.Tag = NXOpen.Tag.Null
                Dim min_corner(2) As Double
                Dim directions(2, 2) As Double
                Dim distances(2) As Double

                If holeBuilder.Type = HolePackageBuilder.Types.GeneralHole Then
                    f

                End If
                lw.WriteLine("")
                holeBuilder.Destroy()

            End If

        Next



    End Sub

End Module
