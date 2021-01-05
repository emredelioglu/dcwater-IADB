Imports ESRI.ArcGIS.Geometry

Public Class RotatingCalipers


    Private m_CurrentArea As Double = 0
    Private m_BestArea As Double = Double.MaxValue
    Private m_BestRectangle() As IPoint

    ' Find the slope of the edge from point i to point i+1.
    Private Sub FindDxDy(ByRef dx As Single, ByRef dy As Single, ByVal i As Integer, ByVal m_NumPoints As Integer, ByRef pointCol As IPointCollection)
        Dim i2 As Integer = (i + 1) Mod m_NumPoints
        dx = pointCol.Point(i2).X - pointCol.Point(i).X
        dy = pointCol.Point(i2).Y - pointCol.Point(i).Y
    End Sub

    ' Find the point of intersection between two lines.
    Private Function FindIntersection(ByVal X1 As Double, ByVal Y1 As Double, ByVal X2 As Double, ByVal Y2 As Double, ByVal A1 As Double, ByVal B1 As Double, ByVal A2 As Double, ByVal B2 As Double, ByRef pt As IPoint) As Boolean
        Dim dx As Double = X2 - X1
        Dim dy As Double = Y2 - Y1
        Dim da As Double = A2 - A1
        Dim db As Double = B2 - B1
        Dim s, t As Double

        ' If the segments are parallel, return False.
        If Math.Abs(da * dy - db * dx) <= 0 Then Return False

        ' Find the point of intersection.
        s = (dx * (B1 - Y1) + dy * (X1 - A1)) / (da * dy - db * dx)
        t = (da * (Y1 - B1) + db * (A1 - X1)) / (db * dx - da * dy)
        pt.PutCoords(X1 + t * dx, Y1 + t * dy)
        Return True
    End Function

    Private Sub ComputeBoundingRectangle(ByVal px0 As Double, ByVal py0 As Double, ByVal dx0 As Double, ByVal dy0 As Double, _
        ByVal px1 As Double, ByVal py1 As Double, ByVal dx1 As Double, ByVal dy1 As Double, _
        ByVal px2 As Double, ByVal py2 As Double, ByVal dx2 As Double, ByVal dy2 As Double, _
        ByVal px3 As Double, ByVal py3 As Double, ByVal dx3 As Double, ByVal dy3 As Double)

        ' Find the points of intersection.
        Dim pts(3) As IPoint
        Dim i As Integer
        For i = 0 To 3
            pts(i) = New Point
        Next
        FindIntersection(px0, py0, px0 + dx0, py0 + dy0, px1, py1, px1 + dx1, py1 + dy1, pts(0))
        FindIntersection(px1, py1, px1 + dx1, py1 + dy1, px2, py2, px2 + dx2, py2 + dy2, pts(1))
        FindIntersection(px2, py2, px2 + dx2, py2 + dy2, px3, py3, px3 + dx3, py3 + dy3, pts(2))
        FindIntersection(px3, py3, px3 + dx3, py3 + dy3, px0, py0, px0 + dx0, py0 + dy0, pts(3))

        ' See if this is the best bounding rectangle do far.
        ' Get the area of the bounding rectangle.
        Dim vx0 As Double = pts(0).X - pts(1).X
        Dim vy0 As Double = pts(0).Y - pts(1).Y
        Dim len0 As Double = Math.Sqrt(vx0 * vx0 + vy0 * vy0)

        Dim vx1 As Double = pts(1).X - pts(2).X
        Dim vy1 As Double = pts(1).Y - pts(2).Y
        Dim len1 As Double = Math.Sqrt(vx1 * vx1 + vy1 * vy1)

        m_CurrentArea = len0 * len1
        If m_CurrentArea > 0 AndAlso m_CurrentArea < m_BestArea Then
            m_BestArea = m_CurrentArea
            m_BestRectangle = pts
        End If

    End Sub



    Public Function GetMER(ByVal pGeomCol As IGeometryCollection) As IPolygon

        Dim p As IPolygon = New Polygon
        Dim pCol As IPointCollection = p

        Dim pTopo As ITopologicalOperator
        'pTopo = pGeomCol
        'Dim pGeomBag As IGeometryBag = pGeomCol
        pTopo = pGeomCol.Geometry(0)
        Dim iGeo As Integer
        For iGeo = 1 To pGeomCol.GeometryCount - 1
            pTopo = pTopo.Union(pGeomCol.Geometry(iGeo))
        Next

        Dim pConvexHull As IGeometry = pTopo.ConvexHull
        pTopo = pConvexHull
        pTopo.Simplify()

        Dim pPointCol As IPointCollection = pConvexHull


        Dim pSegCol As ISegmentCollection = pConvexHull
        If pSegCol.SegmentCount = 3 Then
            Dim iSeg As Integer = 0
            Dim iSegMaxLen As Integer
            Dim segMaxLen As Double = 0
            For iSeg = 0 To pSegCol.SegmentCount - 1
                If segMaxLen < pSegCol.Segment(iSeg).Length Then
                    segMaxLen = pSegCol.Segment(iSeg).Length
                    iSegMaxLen = iSeg
                End If
            Next

            'Now we have the longest segment for rotation

            pCol.AddPoint(pSegCol.Segment(iSegMaxLen).FromPoint)
            pCol.AddPoint(pSegCol.Segment(iSegMaxLen).ToPoint)

            Dim intersectPoint As IPoint = New Point

            Dim px1, py1, px2, py2, px3, py3 As Double
            px1 = pSegCol.Segment(iSegMaxLen).ToPoint.X
            py1 = pSegCol.Segment(iSegMaxLen).ToPoint.Y

            If iSegMaxLen = 2 Then
                px2 = pSegCol.Segment(0).ToPoint.X
                py2 = pSegCol.Segment(0).ToPoint.Y
            Else
                px2 = pSegCol.Segment(iSegMaxLen + 1).ToPoint.X
                py2 = pSegCol.Segment(iSegMaxLen + 1).ToPoint.Y
            End If

            px3 = pSegCol.Segment(iSegMaxLen).FromPoint.X
            py3 = pSegCol.Segment(iSegMaxLen).FromPoint.Y

            Dim dx As Double = px1 - px3
            Dim dy As Double = py1 - py3
            'dy, -dx

            FindIntersection(px1, py1, px1 + dy, py1 - dx, px2, py2, px2 - dx, py2 - dy, intersectPoint)
            pCol.AddPoint(intersectPoint)
            FindIntersection(px2, py2, px2 - dx, py2 + -dy, px3, py3, px3 - dy, py3 + dx, intersectPoint)
            pCol.AddPoint(intersectPoint)

            Return p

        End If

        m_CurrentArea = Double.MaxValue
        m_BestArea = Double.MaxValue


        Dim m_ControlPoints(3) As Integer

        Dim minx As Double = pPointCol.Point(0).X
        Dim maxx As Double = minx
        Dim miny As Double = pPointCol.Point(0).Y
        Dim maxy As Double = miny
        Dim minxi As Integer = 0
        Dim maxxi As Integer = 0
        Dim minyi As Integer = 0
        Dim maxyi As Integer = 0

        Dim i As Integer
        For i = 1 To pPointCol.PointCount - 1
            If minx > pPointCol.Point(i).X Then
                minx = pPointCol.Point(i).X
                minxi = i
            End If
            If maxx < pPointCol.Point(i).X Then
                maxx = pPointCol.Point(i).X
                maxxi = i
            End If
            If miny > pPointCol.Point(i).Y Then
                miny = pPointCol.Point(i).Y
                minyi = i
            End If
            If maxy < pPointCol.Point(i).Y Then
                maxy = pPointCol.Point(i).Y
                maxyi = i
            End If
        Next i

        m_ControlPoints(0) = minxi
        m_ControlPoints(1) = maxyi
        m_ControlPoints(2) = maxxi
        m_ControlPoints(3) = minyi

        Dim m_NumPoints As Integer = pPointCol.PointCount
        Dim m_RectanglesExamined As Integer = 0
        Dim m_CurrentControlPoint As Integer = -1
        Dim m_EdgeChecked() As Boolean

        ReDim m_EdgeChecked(m_NumPoints - 1)
        For i = 0 To m_NumPoints - 1
            m_EdgeChecked(i) = False
        Next i

        While m_RectanglesExamined < pPointCol.PointCount

            If m_CurrentControlPoint >= 0 Then
                m_ControlPoints(m_CurrentControlPoint) = (m_ControlPoints(m_CurrentControlPoint) + 1) Mod m_NumPoints
            End If

            ' Find the next point on an edge.
            Dim xmindx, xmindy As Double
            Dim ymaxdx, ymaxdy As Double
            Dim xmaxdx, xmaxdy As Double
            Dim ymindx, ymindy As Double
            FindDxDy(xmindx, xmindy, m_ControlPoints(0), m_NumPoints, pPointCol)
            FindDxDy(ymaxdx, ymaxdy, m_ControlPoints(1), m_NumPoints, pPointCol)
            FindDxDy(xmaxdx, xmaxdy, m_ControlPoints(2), m_NumPoints, pPointCol)
            FindDxDy(ymindx, ymindy, m_ControlPoints(3), m_NumPoints, pPointCol)

            ' Switch so we can look for the smallest opposite/adjacent ratio.
            Dim xminopp As Double = xmindx
            Dim xminadj As Double = xmindy
            Dim ymaxopp As Double = -ymaxdy
            Dim ymaxadj As Double = ymaxdx
            Dim xmaxopp As Double = -xmaxdx
            Dim xmaxadj As Double = -xmaxdy
            Dim yminopp As Double = ymindy
            Dim yminadj As Double = -ymindx

            ' Pick initial values that will make every point an improvement.
            Dim bestopp As Double = 1
            Dim bestadj As Double = 0
            Dim best_control_point As Integer = -1

            If xminopp >= 0 AndAlso xminadj >= 0 Then
                If xminopp * bestadj < bestopp * xminadj Then
                    bestopp = xminopp
                    bestadj = xminadj
                    best_control_point = 0
                End If
            End If
            If ymaxopp >= 0 AndAlso ymaxadj >= 0 Then
                If ymaxopp * bestadj < bestopp * ymaxadj Then
                    bestopp = ymaxopp
                    bestadj = ymaxadj
                    best_control_point = 1
                End If
            End If
            If xmaxopp >= 0 AndAlso xmaxadj >= 0 Then
                If xmaxopp * bestadj < bestopp * xmaxadj Then
                    bestopp = xmaxopp
                    bestadj = xmaxadj
                    best_control_point = 2
                End If
            End If
            If yminopp >= 0 AndAlso yminadj >= 0 Then
                If yminopp * bestadj < bestopp * yminadj Then
                    bestopp = yminopp
                    bestadj = yminadj
                    best_control_point = 3
                End If
            End If

            ' Make sure we found a usable edge.
            If best_control_point >= 0 Then

                ' Use the new best control point.
                m_CurrentControlPoint = best_control_point
                ' Remember that we have checked this edge.
                m_EdgeChecked(m_ControlPoints(m_CurrentControlPoint)) = True
                ' See if we have checked all possible bounding rectangles.


                Dim i1 As Integer = m_ControlPoints(m_CurrentControlPoint)
                Dim i2 As Integer = (i1 + 1) Mod m_NumPoints
                Dim dx As Double = pPointCol.Point(i2).X - pPointCol.Point(i1).X
                Dim dy As Double = pPointCol.Point(i2).Y - pPointCol.Point(i1).Y

                ' Make dx and dy work for the first line.
                Select Case m_CurrentControlPoint
                    Case 0  ' Nothing to do.
                    Case 1  ' dx = -dy, dy = dx
                        Dim temp As Single = dx
                        dx = -dy
                        dy = temp
                    Case 2  ' dx = -dx, dy = -dy
                        dx = -dx
                        dy = -dy
                    Case 3  ' dx = dy, dy = -dx
                        Dim temp As Single = dx
                        dx = dy
                        dy = -temp
                End Select

                ComputeBoundingRectangle(pPointCol.Point(m_ControlPoints(0)).X, pPointCol.Point(m_ControlPoints(0)).Y, dx, dy, _
                                                    pPointCol.Point(m_ControlPoints(1)).X, pPointCol.Point(m_ControlPoints(1)).Y, dy, -dx, _
                                                    pPointCol.Point(m_ControlPoints(2)).X, pPointCol.Point(m_ControlPoints(2)).Y, -dx, -dy, _
                                                    pPointCol.Point(m_ControlPoints(3)).X, pPointCol.Point(m_ControlPoints(3)).Y, -dy, dx)

            End If
            m_RectanglesExamined += 1
        End While


        For i = 0 To 3
            pCol.AddPoint(m_BestRectangle(i))
        Next

        Return p

    End Function

End Class
