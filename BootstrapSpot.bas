Attribute VB_Name = "BootstrapSpot"
Option Explicit

' =============================================================================
' 회사채 YTM으로부터 스팟 레이트(spot rate) 커브를 부트스트래핑하는 매크로.
' 관행: 3개월 복리(분기 복리), Actual/365 근사.
'
' 사용법:
'   1) 시트에 아래처럼 데이터를 입력한다.
'
'      A열: 만기(년)       B열: 표면금리(소수)  C열: YTM(소수)   D열(옵션): 액면가
'      예)  1                   0.035                 0.0358             100
'
'      헤더는 1행, 데이터는 2행부터.
'   2) 개발도구 > 매크로 > BootstrapSpotCurve 실행 (단축키 Alt+F8)
'   3) E열에 해당 만기의 스팟 레이트가 기록된다.
'
' 중간 쿠폰 시점의 스팟은 이미 부트스트랩된 점들을 선형보간해 사용한다.
' =============================================================================

Public Sub BootstrapSpotCurve()
    Const FREQ As Long = 4               ' 분기 복리
    Const DEFAULT_FACE As Double = 100#
    Const MAX_ITER As Long = 400
    Const TOL As Double = 0.00000000001  ' 1e-11

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "A열 2행부터 만기(년), B열 표면금리, C열 YTM 을 입력하세요.", vbExclamation
        Exit Sub
    End If

    Dim n As Long
    n = lastRow - 1

    Dim mats() As Double, coups() As Double, ytms() As Double, faces() As Double
    ReDim mats(1 To n)
    ReDim coups(1 To n)
    ReDim ytms(1 To n)
    ReDim faces(1 To n)

    Dim i As Long
    For i = 1 To n
        mats(i) = CDbl(ws.Cells(i + 1, 1).Value)
        coups(i) = CDbl(ws.Cells(i + 1, 2).Value)
        ytms(i) = CDbl(ws.Cells(i + 1, 3).Value)
        If IsNumeric(ws.Cells(i + 1, 4).Value) And ws.Cells(i + 1, 4).Value <> "" Then
            faces(i) = CDbl(ws.Cells(i + 1, 4).Value)
        Else
            faces(i) = DEFAULT_FACE
        End If
    Next i

    SortByMaturity mats, coups, ytms, faces, n

    Dim spots() As Double, knownMats() As Double
    ReDim spots(1 To n)
    ReDim knownMats(1 To n)

    ws.Cells(1, 5).Value = "Spot Rate"
    ws.Cells(1, 5).Font.Bold = True

    Dim k As Long
    For k = 1 To n
        Dim T As Double, cRate As Double, y As Double, face As Double
        T = mats(k): cRate = coups(k): y = ytms(k): face = faces(k)

        Dim target As Double
        target = PriceFromYtm(T, cRate, y, face, FREQ)

        Dim lo As Double, hi As Double, mid As Double, pv As Double
        lo = -0.5: hi = 1#
        Dim iter As Long
        For iter = 1 To MAX_ITER
            mid = (lo + hi) / 2
            pv = PriceGivenSpots(T, cRate, face, FREQ, knownMats, spots, k - 1, mid)
            If pv > target Then
                lo = mid
            Else
                hi = mid
            End If
            If (hi - lo) < TOL Then Exit For
        Next iter

        spots(k) = mid
        knownMats(k) = T

        ws.Cells(k + 1, 5).Value = mid
        ws.Cells(k + 1, 5).NumberFormat = "0.0000%"
    Next k

    ws.Columns("A:E").AutoFit
    MsgBox "부트스트래핑 완료: " & n & "개 만기", vbInformation
End Sub

Private Sub SortByMaturity(ByRef mats() As Double, ByRef coups() As Double, _
                           ByRef ytms() As Double, ByRef faces() As Double, ByVal n As Long)
    Dim i As Long, j As Long
    Dim tm As Double, tc As Double, ty As Double, tf As Double
    For i = 2 To n
        tm = mats(i): tc = coups(i): ty = ytms(i): tf = faces(i)
        j = i - 1
        Do While j >= 1
            If mats(j) <= tm Then Exit Do
            mats(j + 1) = mats(j)
            coups(j + 1) = coups(j)
            ytms(j + 1) = ytms(j)
            faces(j + 1) = faces(j)
            j = j - 1
        Loop
        mats(j + 1) = tm
        coups(j + 1) = tc
        ytms(j + 1) = ty
        faces(j + 1) = tf
    Next i
End Sub

Private Function PriceFromYtm(ByVal T As Double, ByVal cRate As Double, ByVal y As Double, _
                              ByVal face As Double, ByVal freq As Long) As Double
    Dim nCpns As Long
    nCpns = CLng(Int(T * freq + 0.5))
    If nCpns < 1 Then nCpns = 1

    Dim coupon As Double
    coupon = cRate * face / freq

    Dim pv As Double, i As Long, t As Double, cf As Double
    pv = 0#
    For i = 1 To nCpns
        If i = nCpns Then
            t = T
            cf = coupon + face
        Else
            t = i / freq
            cf = coupon
        End If
        pv = pv + cf / (1 + y / freq) ^ (freq * t)
    Next i
    PriceFromYtm = pv
End Function

Private Function PriceGivenSpots(ByVal T As Double, ByVal cRate As Double, ByVal face As Double, _
                                 ByVal freq As Long, ByRef knownMats() As Double, _
                                 ByRef knownSpots() As Double, ByVal knownN As Long, _
                                 ByVal sT As Double) As Double
    Dim nCpns As Long
    nCpns = CLng(Int(T * freq + 0.5))
    If nCpns < 1 Then nCpns = 1

    Dim coupon As Double
    coupon = cRate * face / freq

    Dim pv As Double, i As Long, t As Double, cf As Double, s As Double
    pv = 0#
    For i = 1 To nCpns
        If i = nCpns Then
            t = T
            cf = coupon + face
        Else
            t = i / freq
            cf = coupon
        End If
        s = InterpSpot(t, knownMats, knownSpots, knownN, T, sT)
        pv = pv + cf / (1 + s / freq) ^ (freq * t)
    Next i
    PriceGivenSpots = pv
End Function

Private Function InterpSpot(ByVal t As Double, ByRef knownMats() As Double, _
                            ByRef knownSpots() As Double, ByVal knownN As Long, _
                            ByVal newT As Double, ByVal newSpot As Double) As Double
    If knownN = 0 Then
        InterpSpot = newSpot
        Exit Function
    End If

    Dim totalN As Long
    totalN = knownN + 1

    Dim allMats() As Double, allSpots() As Double
    ReDim allMats(1 To totalN)
    ReDim allSpots(1 To totalN)

    Dim i As Long
    For i = 1 To knownN
        allMats(i) = knownMats(i)
        allSpots(i) = knownSpots(i)
    Next i
    allMats(totalN) = newT
    allSpots(totalN) = newSpot

    If t <= allMats(1) Then
        InterpSpot = allSpots(1)
        Exit Function
    End If
    If t >= allMats(totalN) Then
        InterpSpot = allSpots(totalN)
        Exit Function
    End If

    Dim k As Long, w As Double
    For k = 1 To totalN - 1
        If t >= allMats(k) And t <= allMats(k + 1) Then
            w = (t - allMats(k)) / (allMats(k + 1) - allMats(k))
            InterpSpot = allSpots(k) * (1 - w) + allSpots(k + 1) * w
            Exit Function
        End If
    Next k

    InterpSpot = allSpots(totalN)
End Function
