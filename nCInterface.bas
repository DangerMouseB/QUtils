Attribute VB_Name = "nCInterface"
'*************************************************************************************************************************************************************************************************************************************************
'            COPYRIGHT NOTICE
'
' Copyright (C) David Briant 2009 - All rights reserved
'
'*************************************************************************************************************************************************************************************************************************************************

Option Explicit


'*************************************************************************************************************************************************************************************************************************************************
' repository access
'*************************************************************************************************************************************************************************************************************************************************

Function uDBRepository() As Dictionary
    If Not g_IsInitialised Then initDLL
    Set uDBRepository = Repository()
End Function

Sub uDBReleaseRepository()
    If Not g_IsInitialised Then initDLL
    ReleaseRepository
End Sub


'*************************************************************************************************************************************************************************************************************************************************
' random number generators
'*************************************************************************************************************************************************************************************************************************************************

Function uNRRan0(seed As Long) As Double
    If Not g_IsInitialised Then initDLL
    uNRRan0 = ran0(seed)
End Function

Function uNRRan1(seed As Long) As Double
    If Not g_IsInitialised Then initDLL
    uNRRan1 = ran1(seed)
End Function

Function uNRGaussianRan1(seed As Long) As Double
    If Not g_IsInitialised Then initDLL
    uNRGaussianRan1 = gaussianRan1(seed)
End Function

Function uNRLogNormalFromNormal(zeroOneGaussian As Double, mean As Double, sd As Double) As Double
    If Not g_IsInitialised Then initDLL
    uNRLogNormalFromNormal = logNormalFromNormal(zeroOneGaussian, mean, sd)
End Function


'*************************************************************************************************************************************************************************************************************************************************
' GUID generation
'*************************************************************************************************************************************************************************************************************************************************

Function uDBNewGUID() As String
    If Not g_IsInitialised Then initDLL
    uDBNewGUID = DBNewGUID()
End Function


'*************************************************************************************************************************************************************************************************************************************************
' Stats
'*************************************************************************************************************************************************************************************************************************************************

Function uDBMoments1D(data2D() As Double) As Double()
    If Not g_IsInitialised Then initDLL
    uDBMoments1D = DBMoments1D(data2D)
End Function


'*************************************************************************************************************************************************************************************************************************************************
' SafeArray utilities
'*************************************************************************************************************************************************************************************************************************************************

Function uDBGetSafeArrayDetails(anArray As Variant, oSA As SAFEARRAY) As HRESULT
    If Not g_IsInitialised Then initDLL
    uDBGetSafeArrayDetails = getSafeArrayDetails(anArray, oSA)
End Function

Function uDBGetSafeArrayPointer(anArray As Variant) As Long
    If Not g_IsInitialised Then initDLL
    uDBGetSafeArrayPointer = getSafeArrayPointer(anArray)
End Function

Function uDBReDimPreserve(a As Variant, nDimensions As Long, x1 As Long, x2 As Long, y1 As Long, y2 As Long, z1 As Long, z2 As Long) As HRESULT
    If Not g_IsInitialised Then initDLL
    uDBReDimPreserve = redimPreserve(a, nDimensions, x1, x2, y1, y2, z1, z2)
End Function

Function uDBCreateDoubleArrayMap(oMap() As Double, ptr As Long, nDimensions As Long, i1 As Long, i2 As Long, j1 As Long, j2 As Long, k1 As Long, k2 As Long) As HRESULT
    If Not g_IsInitialised Then initDLL
    uDBCreateDoubleArrayMap = createDoubleArrayMap(oMap, ptr, nDimensions, i1, i2, j1, j2, k1, k2)
End Function

Function uDBCreateDateArrayMap(oMap() As Date, ptr As Long, nDimensions As Long, i1 As Long, i2 As Long, j1 As Long, j2 As Long, k1 As Long, k2 As Long) As HRESULT
    If Not g_IsInitialised Then initDLL
    uDBCreateDateArrayMap = createDateArrayMap(oMap, ptr, nDimensions, i1, i2, j1, j2, k1, k2)
End Function

Function uDBCreateSingleArrayMap(oMap() As Single, ptr As Long, nDimensions As Long, i1 As Long, i2 As Long, j1 As Long, j2 As Long, k1 As Long, k2 As Long) As HRESULT
    If Not g_IsInitialised Then initDLL
    uDBCreateSingleArrayMap = createSingleArrayMap(oMap, ptr, nDimensions, i1, i2, j1, j2, k1, k2)
End Function

Function uDBCreateLongArrayMap(oMap() As Long, ptr As Long, nDimensions As Long, i1 As Long, i2 As Long, j1 As Long, j2 As Long, k1 As Long, k2 As Long) As HRESULT
    If Not g_IsInitialised Then initDLL
    uDBCreateLongArrayMap = createLongArrayMap(oMap, ptr, nDimensions, i1, i2, j1, j2, k1, k2)
End Function

Function uDBCreateIntegerArrayMap(oMap() As Integer, ptr As Long, nDimensions As Long, i1 As Long, i2 As Long, j1 As Long, j2 As Long, k1 As Long, k2 As Long) As HRESULT
    If Not g_IsInitialised Then initDLL
    uDBCreateIntegerArrayMap = createIntegerArrayMap(oMap, ptr, nDimensions, i1, i2, j1, j2, k1, k2)
End Function

Function uDBCreateByteArrayMap(oMap() As Byte, ptr As Long, nDimensions As Long, i1 As Long, i2 As Long, j1 As Long, j2 As Long, k1 As Long, k2 As Long) As HRESULT
    If Not g_IsInitialised Then initDLL
    uDBCreateByteArrayMap = createByteArrayMap(oMap, ptr, nDimensions, i1, i2, j1, j2, k1, k2)
End Function

Function uDBReleaseArrayMap(map As Variant) As HRESULT
    If Not g_IsInitialised Then initDLL
    uDBReleaseArrayMap = releaseArrayMap(map)
End Function


'*************************************************************************************************************************************************************************************************************************************************
' Brent Root Finder
'*************************************************************************************************************************************************************************************************************************************************

Function uNRBRF_newToken(tolerance As Double, lower As Double, fLower As Double, UPPER As Double, fUpper As Double) As Double()
    If Not g_IsInitialised Then initDLL
    uNRBRF_newToken = NRBRF_newToken(tolerance, lower, fLower, UPPER, fUpper)
End Function

Function uNRBRF_reEstimateX(BRFToken() As Double) As HRESULT
    If Not g_IsInitialised Then initDLL
    NRBRF_reEstimateX BRFToken
End Function

Function uNRBRF_setFx(BRFToken() As Double, fx As Double) As HRESULT
    If Not g_IsInitialised Then initDLL
    NRBRF_fx(BRFToken) = fx
End Function

Function uNRBRF_x(BRFToken() As Double) As Double
    If Not g_IsInitialised Then initDLL
    uNRBRF_x = NRBRF_x(BRFToken)
End Function

Function uNRBRF_setMaxIterations(BRFToken() As Double, maxIterations As Double) As HRESULT
    If Not g_IsInitialised Then initDLL
    NRBRF_maxIterations(BRFToken) = maxIterations
End Function

Function uNRBRF_isWithinTolerance(BRFToken() As Double) As Boolean
    If Not g_IsInitialised Then initDLL
    uNRBRF_isWithinTolerance = NRBRF_isWithinTolerance(BRFToken)
End Function


'*************************************************************************************************************************************************************************************************************************************************
' Brent Minimiser
'*************************************************************************************************************************************************************************************************************************************************

Function uNRBM_newToken(tolerance As Double, lower As Double, middle As Double, fMiddle As Double, UPPER As Double) As Double()
    If Not g_IsInitialised Then initDLL
    uNRBM_newToken = NRBM_newToken(tolerance, lower, middle, fMiddle, UPPER)
End Function

Function uNRBM_reEstimateX(BMToken() As Double) As HRESULT
    If Not g_IsInitialised Then initDLL
    NRBM_reEstimateX BMToken
End Function

Function uNRBM_fx(BMToken() As Double) As Double
    If Not g_IsInitialised Then initDLL
    uNRBM_fx = NRBM_fx(BMToken)
End Function

Function uNRBM_setFx(BMToken() As Double, fx As Double) As HRESULT
    If Not g_IsInitialised Then initDLL
    NRBM_fx(BMToken) = fx
End Function

Function uNRBM_x(BMToken() As Double) As Double
    If Not g_IsInitialised Then initDLL
    uNRBM_x = NRBM_x(BMToken)
End Function

Function uNRBM_setMaxIterations(BMToken() As Double, maxIterations As Double) As HRESULT
    If Not g_IsInitialised Then initDLL
    NRBM_maxIterations(BMToken) = maxIterations
End Function

Function uNRBM_isWithinTolerance(BMToken() As Double) As Boolean
    If Not g_IsInitialised Then initDLL
    uNRBM_isWithinTolerance = NRBM_isWithinTolerance(BMToken)
End Function

Function uNRBM_numberOfIterations(BMToken() As Double) As Long
    If Not g_IsInitialised Then initDLL
    uNRBM_numberOfIterations = NRBM_numberOfIterations(BMToken)
End Function


'*************************************************************************************************************************************************************************************************************************************************
' Cholesky Decomposition
'*************************************************************************************************************************************************************************************************************************************************

Function uNRCholeskyDecomposition(io_aMatrix() As Double, n As Long, o_pVector2D() As Double) As HRESULT
    If Not g_IsInitialised Then initDLL
    uNRCholeskyDecomposition = NRCholeskyDecomposition(io_aMatrix, n, o_pVector2D)
End Function

Function uNRCholeskySolve(aMatrix() As Double, n As Long, pVector2D() As Double, bVector2D() As Double, ioXVector2D() As Double) As HRESULT
    If Not g_IsInitialised Then initDLL
    NRCholeskySolve aMatrix, n, pVector2D, bVector2D, ioXVector2D
    uNRCholeskySolve.HRESULT = S_OK
End Function

