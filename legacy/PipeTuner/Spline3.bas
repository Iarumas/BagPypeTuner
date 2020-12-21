Attribute VB_Name = "Spline"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Copyright (c) 2007, Sergey Bochkanov (ALGLIB project).
'
'Redistribution and use in source and binary forms, with or without
'modification, are permitted provided that the following conditions are
'met:
'
'- Redistributions of source code must retain the above copyright
'  notice, this list of conditions and the following disclaimer.
'
'- Redistributions in binary form must reproduce the above copyright
'  notice, this list of conditions and the following disclaimer listed
'  in this license in the documentation and/or other materials
'  provided with the distribution.
'
'- Neither the name of the copyright holders nor the names of its
'  contributors may be used to endorse or promote products derived from
'  this software without specific prior written permission.
'
'THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
'"AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
'LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
'A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT
'OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
'SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT
'LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,
'DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
'THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
'(INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
'OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Routines
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This subroutine builds linear spline coefficients table.
'
'Input parameters:
'    X   -   spline nodes, array[0..N-1]
'    Y   -   function values, array[0..N-1]
'    N   -   points count, N>=2
'
'Output parameters:
'    C   -   coefficients table.  Used  by  SplineInterpolation  and  other
'            subroutines from this file.
'
'  -- ALGLIB PROJECT --
'     Copyright 24.06.2007 by Bochkanov Sergey
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub BuildLinearSpline(ByRef x_() As Double, _
         ByRef y_() As Double, _
         ByVal N As Long, _
         ByRef C() As Double)
    Dim X() As Double
    Dim Y() As Double
    Dim i As Long
    Dim TblSize As Long
    X = x_
    Y = y_

    
    '
    ' Sort points
    '
    Call HeapSortPoints(X, Y, N)
    
    '
    ' Fill C:
    '  C[0]            -   length(C)
    '  C[1]            -   type(C):
    '                      3 - general cubic spline
    '  C[2]            -   N
    '  C[3]...C[3+N-1] -   x[i], i = 0...N-1
    '  C[3+N]...C[3+N+(N-1)*4-1] - coefficients table
    '
    TblSize = 3# + N + (N - 1#) * 4#
    ReDim C(0# To TblSize - 1#)
    C(0#) = TblSize
    C(1#) = 3#
    C(2#) = N
    For i = 0# To N - 1# Step 1
        C(3# + i) = X(i)
    Next i
    For i = 0# To N - 2# Step 1
        C(3# + N + 4# * i + 0#) = Y(i)
        C(3# + N + 4# * i + 1#) = (Y(i + 1#) - Y(i)) / (X(i + 1#) - X(i))
        C(3# + N + 4# * i + 2#) = 0#
        C(3# + N + 4# * i + 3#) = 0#
    Next i
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This subroutine builds cubic spline coefficients table.
'
'Input parameters:
'    X           -   spline nodes, array[0..N-1]
'    Y           -   function values, array[0..N-1]
'    N           -   points count, N>=2
'    BoundLType  -   boundary condition type for the left boundary
'    BoundL      -   left boundary condition (first or second derivative,
'                    depending on the BoundLType)
'    BoundRType  -   boundary condition type for the right boundary
'    BoundR      -   right boundary condition (first or second derivative,
'                    depending on the BoundRType)
'
'Output parameters:
'    C           -   coefficients table.  Used  by  SplineInterpolation and
'                    other subroutines from this file.
'
'The BoundLType/BoundRType parameters can have the following values:
'    * 0,   which  corresponds  to  the  parabolically   terminated  spline
'      (BoundL/BoundR are ignored).
'    * 1, which corresponds to the first derivative boundary condition
'    * 2, which corresponds to the second derivative boundary condition
'
'  -- ALGLIB PROJECT --
'     Copyright 23.06.2007 by Bochkanov Sergey
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub BuildCubicSpline(ByRef x_() As Double, _
         ByRef y_() As Double, _
         ByVal N As Long, _
         ByVal BoundLType As Long, _
         ByVal BoundL As Double, _
         ByVal BoundRType As Long, _
         ByVal BoundR As Double, _
         ByRef C() As Double)
    Dim X() As Double
    Dim Y() As Double
    Dim A1() As Double
    Dim A2() As Double
    Dim A3() As Double
    Dim B() As Double
    Dim D() As Double
    Dim i As Long
    Dim TblSize As Long
    Dim Delta As Double
    Dim Delta2 As Double
    Dim Delta3 As Double
    X = x_
    Y = y_

    ReDim A1(0# To N - 1#)
    ReDim A2(0# To N - 1#)
    ReDim A3(0# To N - 1#)
    ReDim B(0# To N - 1#)
    
    '
    ' Special case:
    ' * N=2
    ' * parabolic terminated boundary condition on both ends
    '
    If N = 2# And BoundLType = 0# And BoundRType = 0# Then
        
        '
        ' Change task type
        '
        BoundLType = 2#
        BoundL = 0#
        BoundRType = 2#
        BoundR = 0#
    End If
    
    '
    '
    ' Sort points
    '
    Call HeapSortPoints(X, Y, N)
    
    '
    ' Left boundary conditions
    '
    If BoundLType = 0# Then
        A1(0#) = 0#
        A2(0#) = 1#
        A3(0#) = 1#
        B(0#) = 2# * (Y(1#) - Y(0#)) / (X(1#) - X(0#))
    End If
    If BoundLType = 1# Then
        A1(0#) = 0#
        A2(0#) = 1#
        A3(0#) = 0#
        B(0#) = BoundL
    End If
    If BoundLType = 2# Then
        A1(0#) = 0#
        A2(0#) = 2#
        A3(0#) = 1#
        B(0#) = 3# * (Y(1#) - Y(0#)) / (X(1#) - X(0#)) - 0.5 * BoundL * (X(1#) - X(0#))
    End If
    
    '
    ' Central conditions
    '
    For i = 1# To N - 2# Step 1
        A1(i) = X(i + 1#) - X(i)
        A2(i) = 2# * (X(i + 1#) - X(i - 1#))
        A3(i) = X(i) - X(i - 1#)
        B(i) = 3# * (Y(i) - Y(i - 1#)) / (X(i) - X(i - 1#)) * (X(i + 1#) - X(i)) + 3# * (Y(i + 1#) - Y(i)) / (X(i + 1#) - X(i)) * (X(i) - X(i - 1#))
    Next i
    
    '
    ' Right boundary conditions
    '
    If BoundRType = 0# Then
        A1(N - 1#) = 1#
        A2(N - 1#) = 1#
        A3(N - 1#) = 0#
        B(N - 1#) = 2# * (Y(N - 1#) - Y(N - 2#)) / (X(N - 1#) - X(N - 2#))
    End If
    If BoundRType = 1# Then
        A1(N - 1#) = 0#
        A2(N - 1#) = 1#
        A3(N - 1#) = 0#
        B(N - 1#) = BoundR
    End If
    If BoundRType = 2# Then
        A1(N - 1#) = 1#
        A2(N - 1#) = 2#
        A3(N - 1#) = 0#
        B(N - 1#) = 3# * (Y(N - 1#) - Y(N - 2#)) / (X(N - 1#) - X(N - 2#)) + 0.5 * BoundR * (X(N - 1#) - X(N - 2#))
    End If
    
    '
    ' Solve
    '
    Call SolveTridiagonal(A1, A2, A3, B, N, D)
    
    '
    ' Now problem is reduced to the cubic Hermite spline
    '
    Call BuildHermiteSpline(X, Y, D, N, C)
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This subroutine builds cubic Hermite spline coefficients table.
'
'Input parameters:
'    X           -   spline nodes, array[0..N-1]
'    Y           -   function values, array[0..N-1]
'    D           -   derivatives, array[0..N-1]
'    N           -   points count, N>=2
'
'Output parameters:
'    C           -   coefficients table.  Used  by  SplineInterpolation and
'                    other subroutines from this file.
'
'  -- ALGLIB PROJECT --
'     Copyright 23.06.2007 by Bochkanov Sergey
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub BuildHermiteSpline(ByRef x_() As Double, _
         ByRef y_() As Double, _
         ByRef D_() As Double, _
         ByVal N As Long, _
         ByRef C() As Double)
    Dim X() As Double
    Dim Y() As Double
    Dim D() As Double
    Dim i As Long
    Dim TblSize As Long
    Dim Delta As Double
    Dim Delta2 As Double
    Dim Delta3 As Double
    X = x_
    Y = y_
    D = D_

    
    '
    ' Sort points
    '
    Call HeapSortDPoints(X, Y, D, N)
    
    '
    ' Fill C:
    '  C[0]            -   length(C)
    '  C[1]            -   type(C):
    '                      3 - general cubic spline
    '  C[2]            -   N
    '  C[3]...C[3+N-1] -   x[i], i = 0...N-1
    '  C[3+N]...C[3+N+(N-1)*4-1] - coefficients table
    '
    TblSize = 3# + N + (N - 1#) * 4#
    ReDim C(0# To TblSize - 1#)
    C(0#) = TblSize
    C(1#) = 3#
    C(2#) = N
    For i = 0# To N - 1# Step 1
        C(3# + i) = X(i)
    Next i
    For i = 0# To N - 2# Step 1
        Delta = X(i + 1#) - X(i)
        Delta2 = Square(Delta)
        Delta3 = Delta * Delta2
        C(3# + N + 4# * i + 0#) = Y(i)
        C(3# + N + 4# * i + 1#) = D(i)
        C(3# + N + 4# * i + 2#) = (3# * (Y(i + 1#) - Y(i)) - 2# * D(i) * Delta - D(i + 1#) * Delta) / Delta2
        C(3# + N + 4# * i + 3#) = (2# * (Y(i) - Y(i + 1#)) + D(i) * Delta + D(i + 1#) * Delta) / Delta3
    Next i
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This subroutine builds Akima spline coefficients table.
'
'Input parameters:
'    X           -   spline nodes, array[0..N-1]
'    Y           -   function values, array[0..N-1]
'    N           -   points count, N>=5
'
'Output parameters:
'    C           -   coefficients table.  Used  by  SplineInterpolation and
'                    other subroutines from this file.
'
'  -- ALGLIB PROJECT --
'     Copyright 24.06.2007 by Bochkanov Sergey
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub BuildAkimaSpline(ByRef x_() As Double, _
         ByRef y_() As Double, _
         ByVal N As Long, _
         ByRef C() As Double)
    Dim X() As Double
    Dim Y() As Double
    Dim i As Long
    Dim D() As Double
    Dim W() As Double
    Dim Diff() As Double
    X = x_
    Y = y_

    
    '
    ' Sort points
    '
    Call HeapSortPoints(X, Y, N)
    
    '
    ' Prepare W (weights), Diff (divided differences)
    '
    ReDim W(1# To N - 2#)
    ReDim Diff(0# To N - 2#)
    For i = 0# To N - 2# Step 1
        Diff(i) = (Y(i + 1#) - Y(i)) / (X(i + 1#) - X(i))
    Next i
    For i = 1# To N - 2# Step 1
        W(i) = Abs(Diff(i) - Diff(i - 1#))
    Next i
    
    '
    ' Prepare Hermite interpolation scheme
    '
    ReDim D(0# To N - 1#)
    For i = 2# To N - 3# Step 1
        If Abs(W(i - 1#)) + Abs(W(i + 1#)) <> 0# Then
            D(i) = (W(i + 1#) * Diff(i - 1#) + W(i - 1#) * Diff(i)) / (W(i + 1#) + W(i - 1#))
        Else
            D(i) = ((X(i + 1#) - X(i)) * Diff(i - 1#) + (X(i) - X(i - 1#)) * Diff(i)) / (X(i + 1#) - X(i - 1#))
        End If
    Next i
    D(0#) = DiffThreePoint(X(0#), X(0#), Y(0#), X(1#), Y(1#), X(2#), Y(2#))
    D(1#) = DiffThreePoint(X(1#), X(0#), Y(0#), X(1#), Y(1#), X(2#), Y(2#))
    D(N - 2#) = DiffThreePoint(X(N - 2#), X(N - 3#), Y(N - 3#), X(N - 2#), Y(N - 2#), X(N - 1#), Y(N - 1#))
    D(N - 1#) = DiffThreePoint(X(N - 1#), X(N - 3#), Y(N - 3#), X(N - 2#), Y(N - 2#), X(N - 1#), Y(N - 1#))
    
    '
    ' Build Akima spline using Hermite interpolation scheme
    '
    Call BuildHermiteSpline(X, Y, D, N, C)
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This subroutine calculates the value of the spline at the given point X.
'
'Input parameters:
'    C           -   coefficients table. Built by BuildLinearSpline,
'                    BuildHermiteSpline, BuildCubicSpline, BuildAkimaSpline.
'    X           -   point
'
'Result:
'    S(x)
'
'  -- ALGLIB PROJECT --
'     Copyright 23.06.2007 by Bochkanov Sergey
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SplineInterpolation(ByRef C() As Double, _
         ByVal X As Double) As Double
    Dim Result As Double
    Dim N As Long
    Dim L As Long
    Dim R As Long
    Dim M As Long

    N = Round(C(2#))
    
    '
    ' Binary search in the [ x[0], ..., x[n-2] ] (x[n-1] is not included)
    '
    L = 3#
    R = 3# + N - 2# + 1#
    Do While L <> R - 1#
        M = (L + R) \ 2#
        If C(M) >= X Then
            R = M
        Else
            L = M
        End If
    Loop
    
    '
    ' Interpolation
    '
    X = X - C(L)
    M = 3# + N + 4# * (L - 3#)
    Result = C(M) + X * (C(M + 1#) + X * (C(M + 2#) + X * C(M + 3#)))

    SplineInterpolation = Result
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This subroutine differentiates the spline.
'
'Input parameters:
'    C   -   coefficients table. Built by BuildLinearSpline,
'            BuildHermiteSpline, BuildCubicSpline, BuildAkimaSpline.
'    X   -   point
'
'Result:
'    S   -   S(x)
'    DS  -   S'(x)
'    D2S -   S''(x)
'
'  -- ALGLIB PROJECT --
'     Copyright 24.06.2007 by Bochkanov Sergey
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SplineDifferentiation(ByRef C() As Double, _
         ByVal X As Double, _
         ByRef S As Double, _
         ByRef DS As Double, _
         ByRef D2S As Double)
    Dim N As Long
    Dim L As Long
    Dim R As Long
    Dim M As Long

    N = Round(C(2#))
    
    '
    ' Binary search
    '
    L = 3#
    R = 3# + N - 2# + 1#
    Do While L <> R - 1#
        M = (L + R) \ 2#
        If C(M) >= X Then
            R = M
        Else
            L = M
        End If
    Loop
    
    '
    ' Differentiation
    '
    X = X - C(L)
    M = 3# + N + 4# * (L - 3#)
    S = C(M) + X * (C(M + 1#) + X * (C(M + 2#) + X * C(M + 3#)))
    DS = C(M + 1#) + 2# * X * C(M + 2#) + 3# * Square(X) * C(M + 3#)
    D2S = 2# * C(M + 2#) + 6# * X * C(M + 3#)
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This subroutine makes the copy of the spline.
'
'Input parameters:
'    C   -   coefficients table. Built by BuildLinearSpline,
'            BuildHermiteSpline, BuildCubicSpline, BuildAkimaSpline.
'
'Result:
'    CC  -   spline copy
'
'  -- ALGLIB PROJECT --
'     Copyright 29.06.2007 by Bochkanov Sergey
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SplineCopy(ByRef C() As Double, ByRef CC() As Double)
    Dim S As Long
    Dim i_ As Long

    S = Round(C(0#))
    ReDim CC(0# To S - 1#)
    For i_ = 0# To S - 1# Step 1
        CC(i_) = C(i_)
    Next i_
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This subroutine unpacks the spline into the coefficients table.
'
'Input parameters:
'    C   -   coefficients table. Built by BuildLinearSpline,
'            BuildHermiteSpline, BuildCubicSpline, BuildAkimaSpline.
'    X   -   point
'
'Result:
'    Tbl -   coefficients table, unpacked format, array[0..N-2, 0..5].
'            For I = 0...N-2:
'                Tbl[I,0] = X[i]
'                Tbl[I,1] = X[i+1]
'                Tbl[I,2] = C0
'                Tbl[I,3] = C1
'                Tbl[I,4] = C2
'                Tbl[I,5] = C3
'            On [x[i], x[i+1]] spline is equals to:
'                S(x) = C0 + C1*t + C2*t^2 + C3*t^3
'                t = x-x[i]
'
'  -- ALGLIB PROJECT --
'     Copyright 29.06.2007 by Bochkanov Sergey
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SplineUnpack(ByRef C() As Double, _
         ByRef N As Long, _
         ByRef Tbl() As Double)
    Dim i As Long

    N = Round(C(2#))
    ReDim Tbl(0# To N - 2#, 0# To 5#)
    
    '
    ' Fill
    '
    For i = 0# To N - 2# Step 1
        Tbl(i, 0#) = C(3# + i)
        Tbl(i, 1#) = C(3# + i + 1#)
        Tbl(i, 2#) = C(3# + N + 4# * i)
        Tbl(i, 3#) = C(3# + N + 4# * i + 1#)
        Tbl(i, 4#) = C(3# + N + 4# * i + 2#)
        Tbl(i, 5#) = C(3# + N + 4# * i + 3#)
    Next i
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This subroutine performs linear transformation of the spline argument.
'
'Input parameters:
'    C   -   coefficients table. Built by BuildLinearSpline,
'            BuildHermiteSpline, BuildCubicSpline, BuildAkimaSpline.
'    A, B-   transformation coefficients: x = A*t + B
'Result:
'    C   -   transformed spline
'
'  -- ALGLIB PROJECT --
'     Copyright 30.06.2007 by Bochkanov Sergey
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SplineLinTransX(ByRef C() As Double, _
         ByVal A As Double, _
         ByVal B As Double)
    Dim i As Long
    Dim N As Long
    Dim V As Double
    Dim DV As Double
    Dim D2V As Double
    Dim X() As Double
    Dim Y() As Double
    Dim D() As Double

    N = Round(C(2#))
    
    '
    ' Special case: A=0
    '
    If A = 0# Then
        V = SplineInterpolation(C, B)
        For i = 0# To N - 2# Step 1
            C(3# + N + 4# * i) = V
            C(3# + N + 4# * i + 1#) = 0#
            C(3# + N + 4# * i + 2#) = 0#
            C(3# + N + 4# * i + 3#) = 0#
        Next i
        Exit Sub
    End If
    
    '
    ' General case: A<>0.
    ' Unpack, X, Y, dY/dX.
    ' Scale and pack again.
    '
    ReDim X(0# To N - 1#)
    ReDim Y(0# To N - 1#)
    ReDim D(0# To N - 1#)
    For i = 0# To N - 1# Step 1
        X(i) = C(3# + i)
        Call SplineDifferentiation(C, X(i), V, DV, D2V)
        X(i) = (X(i) - B) / A
        Y(i) = V
        D(i) = A * DV
    Next i
    Call BuildHermiteSpline(X, Y, D, N, C)
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This subroutine performs linear transformation of the spline.
'
'Input parameters:
'    C   -   coefficients table. Built by BuildLinearSpline,
'            BuildHermiteSpline, BuildCubicSpline, BuildAkimaSpline.
'    A, B-   transformation coefficients: S2(x) = A*S(x) + B
'Result:
'    C   -   transformed spline
'
'  -- ALGLIB PROJECT --
'     Copyright 30.06.2007 by Bochkanov Sergey
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SplineLinTransY(ByRef C() As Double, _
         ByVal A As Double, _
         ByVal B As Double)
    Dim i As Long
    Dim N As Long
    Dim V As Double
    Dim DV As Double
    Dim D2V As Double
    Dim X() As Double
    Dim Y() As Double
    Dim D() As Double

    N = Round(C(2#))
    
    '
    ' Special case: A=0
    '
    For i = 0# To N - 2# Step 1
        C(3# + N + 4# * i) = A * C(3# + N + 4# * i) + B
        C(3# + N + 4# * i + 1#) = A * C(3# + N + 4# * i + 1#)
        C(3# + N + 4# * i + 2#) = A * C(3# + N + 4# * i + 2#)
        C(3# + N + 4# * i + 3#) = A * C(3# + N + 4# * i + 3#)
    Next i
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This subroutine integrates the spline.
'
'Input parameters:
'    C   -   coefficients table. Built by BuildLinearSpline,
'            BuildHermiteSpline, BuildCubicSpline, BuildAkimaSpline.
'    X   -   right bound of the integration interval [a, x]
'Result:
'    integral(S(t)dt,a,x)
'
'  -- ALGLIB PROJECT --
'     Copyright 23.06.2007 by Bochkanov Sergey
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SplineIntegration(ByRef C() As Double, _
         ByVal X As Double) As Double
    Dim Result As Double
    Dim N As Long
    Dim i As Long
    Dim L As Long
    Dim R As Long
    Dim M As Long
    Dim W As Double

    N = Round(C(2#))
    
    '
    ' Binary search in the [ x[0], ..., x[n-2] ] (x[n-1] is not included)
    '
    L = 3#
    R = 3# + N - 2# + 1#
    Do While L <> R - 1#
        M = (L + R) \ 2#
        If C(M) >= X Then
            R = M
        Else
            L = M
        End If
    Loop
    
    '
    ' Integration
    '
    Result = 0#
    For i = 3# To L - 1# Step 1
        W = C(i + 1#) - C(i)
        M = 3# + N + 4# * (i - 3#)
        Result = Result + C(M) * W
        Result = Result + C(M + 1#) * Square(W) / 2#
        Result = Result + C(M + 2#) * Square(W) * W / 3#
        Result = Result + C(M + 3#) * Square(Square(W)) / 4#
    Next i
    W = X - C(L)
    M = 3# + N + 4# * (L - 3#)
    Result = Result + C(M) * W
    Result = Result + C(M + 1#) * Square(W) / 2#
    Result = Result + C(M + 2#) * Square(W) * W / 3#
    Result = Result + C(M + 3#) * Square(Square(W)) / 4#

    SplineIntegration = Result
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Obsolete subroutine, left for backward compatibility.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Spline3BuildTable(ByVal N As Long, _
         ByRef DiffN As Long, _
         ByRef x_() As Double, _
         ByRef y_() As Double, _
         ByRef BoundL As Double, _
         ByRef BoundR As Double, _
         ByRef ctbl() As Double)
    Dim X() As Double
    Dim Y() As Double
    Dim C As Boolean
    Dim E As Long
    Dim G As Long
    Dim Tmp As Double
    Dim nxm1 As Long
    Dim i As Long
    Dim j As Long
    Dim DX As Double
    Dim DXJ As Double
    Dim DYJ As Double
    Dim DXJP1 As Double
    Dim DYJP1 As Double
    Dim DXP As Double
    Dim DYP As Double
    Dim YPPA As Double
    Dim YPPB As Double
    Dim PJ As Double
    Dim b1 As Double
    Dim b2 As Double
    Dim b3 As Double
    Dim b4 As Double
    X = x_
    Y = y_

    N = N - 1#
    G = (N + 1#) \ 2#
    Do
        i = G
        Do
            j = i - G
            C = True
            Do
                If X(j) <= X(j + G) Then
                    C = False
                Else
                    Tmp = X(j)
                    X(j) = X(j + G)
                    X(j + G) = Tmp
                    Tmp = Y(j)
                    Y(j) = Y(j + G)
                    Y(j + G) = Tmp
                End If
                j = j - 1#
            Loop Until Not (j >= 0# And C)
            i = i + 1#
        Loop Until Not i <= N
        G = G \ 2#
    Loop Until Not G > 0#
    ReDim ctbl(0# To 4#, 0# To N)
    N = N + 1#
    If DiffN = 1# Then
        b1 = 1#
        b2 = 6# / (X(1#) - X(0#)) * ((Y(1#) - Y(0#)) / (X(1#) - X(0#)) - BoundL)
        b3 = 1#
        b4 = 6# / (X(N - 1#) - X(N - 2#)) * (BoundR - (Y(N - 1#) - Y(N - 2#)) / (X(N - 1#) - X(N - 2#)))
    Else
        b1 = 0#
        b2 = 2# * BoundL
        b3 = 0#
        b4 = 2# * BoundR
    End If
    nxm1 = N - 1#
    If N >= 2# Then
        If N > 2# Then
            DXJ = X(1#) - X(0#)
            DYJ = Y(1#) - Y(0#)
            j = 2#
            Do While j <= nxm1
                DXJP1 = X(j) - X(j - 1#)
                DYJP1 = Y(j) - Y(j - 1#)
                DXP = DXJ + DXJP1
                ctbl(1#, j - 1#) = DXJP1 / DXP
                ctbl(2#, j - 1#) = 1# - ctbl(1#, j - 1#)
                ctbl(3#, j - 1#) = 6# * (DYJP1 / DXJP1 - DYJ / DXJ) / DXP
                DXJ = DXJP1
                DYJ = DYJP1
                j = j + 1#
            Loop
        End If
        ctbl(1#, 0#) = -(b1 / 2#)
        ctbl(2#, 0#) = b2 / 2#
        If N <> 2# Then
            j = 2#
            Do While j <= nxm1
                PJ = ctbl(2#, j - 1#) * ctbl(1#, j - 2#) + 2#
                ctbl(1#, j - 1#) = -(ctbl(1#, j - 1#) / PJ)
                ctbl(2#, j - 1#) = (ctbl(3#, j - 1#) - ctbl(2#, j - 1#) * ctbl(2#, j - 2#)) / PJ
                j = j + 1#
            Loop
        End If
        YPPB = (b4 - b3 * ctbl(2#, nxm1 - 1#)) / (b3 * ctbl(1#, nxm1 - 1#) + 2#)
        i = 1#
        Do While i <= nxm1
            j = N - i
            YPPA = ctbl(1#, j - 1#) * YPPB + ctbl(2#, j - 1#)
            DX = X(j) - X(j - 1#)
            ctbl(3#, j - 1#) = (YPPB - YPPA) / DX / 6#
            ctbl(2#, j - 1#) = YPPA / 2#
            ctbl(1#, j - 1#) = (Y(j) - Y(j - 1#)) / DX - (ctbl(2#, j - 1#) + ctbl(3#, j - 1#) * DX) * DX
            YPPB = YPPA
            i = i + 1#
        Loop
        For i = 1# To N Step 1
            ctbl(0#, i - 1#) = Y(i - 1#)
            ctbl(4#, i - 1#) = X(i - 1#)
        Next i
    End If
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Obsolete subroutine, left for backward compatibility.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Spline3Interpolate(ByVal N As Long, _
         ByRef C() As Double, _
         ByRef X As Double) As Double
    Dim Result As Double
    Dim i As Long
    Dim L As Long
    Dim Half As Long
    Dim First As Long
    Dim Middle As Long

    N = N - 1#
    L = N
    First = 0#
    Do While L > 0#
        Half = L \ 2#
        Middle = First + Half
        If C(4#, Middle) < X Then
            First = Middle + 1#
            L = L - Half - 1#
        Else
            L = Half
        End If
    Loop
    i = First - 1#
    If i < 0# Then
        i = 0#
    End If
    Result = C(0#, i) + (X - C(4#, i)) * (C(1#, i) + (X - C(4#, i)) * (C(2#, i) + C(3#, i) * (X - C(4#, i))))

    Spline3Interpolate = Result
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Internal subroutine. Heap sort.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HeapSortPoints(ByRef X() As Double, _
         ByRef Y() As Double, _
         ByVal N As Long)
    Dim i As Long
    Dim j As Long
    Dim K As Long
    Dim T As Long
    Dim Tmp As Double
    Dim IsAscending As Boolean
    Dim IsDescending As Boolean

    
    '
    ' Test for already sorted set
    '
    IsAscending = True
    IsDescending = True
    For i = 1# To N - 1# Step 1
        IsAscending = IsAscending And X(i) > X(i - 1#)
        IsDescending = IsDescending And X(i) < X(i - 1#)
    Next i
    If IsAscending Then
        Exit Sub
    End If
    If IsDescending Then
        For i = 0# To N - 1# Step 1
            j = N - 1# - i
            If j <= i Then
                Exit For
            End If
            Tmp = X(i)
            X(i) = X(j)
            X(j) = Tmp
            Tmp = Y(i)
            Y(i) = Y(j)
            Y(j) = Tmp
        Next i
        Exit Sub
    End If
    
    '
    ' Special case: N=1
    '
    If N = 1# Then
        Exit Sub
    End If
    
    '
    ' General case
    '
    i = 2#
    Do
        T = i
        Do While T <> 1#
            K = T \ 2#
            If X(K - 1#) >= X(T - 1#) Then
                T = 1#
            Else
                Tmp = X(K - 1#)
                X(K - 1#) = X(T - 1#)
                X(T - 1#) = Tmp
                Tmp = Y(K - 1#)
                Y(K - 1#) = Y(T - 1#)
                Y(T - 1#) = Tmp
                T = K
            End If
        Loop
        i = i + 1#
    Loop Until Not i <= N
    i = N - 1#
    Do
        Tmp = X(i)
        X(i) = X(0#)
        X(0#) = Tmp
        Tmp = Y(i)
        Y(i) = Y(0#)
        Y(0#) = Tmp
        T = 1#
        Do While T <> 0#
            K = 2# * T
            If K > i Then
                T = 0#
            Else
                If K < i Then
                    If X(K) > X(K - 1#) Then
                        K = K + 1#
                    End If
                End If
                If X(T - 1#) >= X(K - 1#) Then
                    T = 0#
                Else
                    Tmp = X(K - 1#)
                    X(K - 1#) = X(T - 1#)
                    X(T - 1#) = Tmp
                    Tmp = Y(K - 1#)
                    Y(K - 1#) = Y(T - 1#)
                    Y(T - 1#) = Tmp
                    T = K
                End If
            End If
        Loop
        i = i - 1#
    Loop Until Not i >= 1#
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Internal subroutine. Heap sort.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HeapSortDPoints(ByRef X() As Double, _
         ByRef Y() As Double, _
         ByRef D() As Double, _
         ByVal N As Long)
    Dim i As Long
    Dim j As Long
    Dim K As Long
    Dim T As Long
    Dim Tmp As Double
    Dim IsAscending As Boolean
    Dim IsDescending As Boolean

    
    '
    ' Test for already sorted set
    '
    IsAscending = True
    IsDescending = True
    For i = 1# To N - 1# Step 1
        IsAscending = IsAscending And X(i) > X(i - 1#)
        IsDescending = IsDescending And X(i) < X(i - 1#)
    Next i
    If IsAscending Then
        Exit Sub
    End If
    If IsDescending Then
        For i = 0# To N - 1# Step 1
            j = N - 1# - i
            If j <= i Then
                Exit For
            End If
            Tmp = X(i)
            X(i) = X(j)
            X(j) = Tmp
            Tmp = Y(i)
            Y(i) = Y(j)
            Y(j) = Tmp
            Tmp = D(i)
            D(i) = D(j)
            D(j) = Tmp
        Next i
        Exit Sub
    End If
    
    '
    ' Special case: N=1
    '
    If N = 1# Then
        Exit Sub
    End If
    
    '
    ' General case
    '
    i = 2#
    Do
        T = i
        Do While T <> 1#
            K = T \ 2#
            If X(K - 1#) >= X(T - 1#) Then
                T = 1#
            Else
                Tmp = X(K - 1#)
                X(K - 1#) = X(T - 1#)
                X(T - 1#) = Tmp
                Tmp = Y(K - 1#)
                Y(K - 1#) = Y(T - 1#)
                Y(T - 1#) = Tmp
                Tmp = D(K - 1#)
                D(K - 1#) = D(T - 1#)
                D(T - 1#) = Tmp
                T = K
            End If
        Loop
        i = i + 1#
    Loop Until Not i <= N
    i = N - 1#
    Do
        Tmp = X(i)
        X(i) = X(0#)
        X(0#) = Tmp
        Tmp = Y(i)
        Y(i) = Y(0#)
        Y(0#) = Tmp
        Tmp = D(i)
        D(i) = D(0#)
        D(0#) = Tmp
        T = 1#
        Do While T <> 0#
            K = 2# * T
            If K > i Then
                T = 0#
            Else
                If K < i Then
                    If X(K) > X(K - 1#) Then
                        K = K + 1#
                    End If
                End If
                If X(T - 1#) >= X(K - 1#) Then
                    T = 0#
                Else
                    Tmp = X(K - 1#)
                    X(K - 1#) = X(T - 1#)
                    X(T - 1#) = Tmp
                    Tmp = Y(K - 1#)
                    Y(K - 1#) = Y(T - 1#)
                    Y(T - 1#) = Tmp
                    Tmp = D(K - 1#)
                    D(K - 1#) = D(T - 1#)
                    D(T - 1#) = Tmp
                    T = K
                End If
            End If
        Loop
        i = i - 1#
    Loop Until Not i >= 1#
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Internal subroutine. Tridiagonal solver.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SolveTridiagonal(ByRef A_() As Double, _
         ByRef B_() As Double, _
         ByRef C_() As Double, _
         ByRef D_() As Double, _
         ByVal N As Long, _
         ByRef X() As Double)
    Dim A() As Double
    Dim B() As Double
    Dim C() As Double
    Dim D() As Double
    Dim K As Long
    Dim T As Double
    A = A_
    B = B_
    C = C_
    D = D_

    ReDim X(0# To N - 1#)
    A(0#) = 0#
    C(N - 1#) = 0#
    For K = 1# To N - 1# Step 1
        T = A(K) / B(K - 1#)
        B(K) = B(K) - T * C(K - 1#)
        D(K) = D(K) - T * D(K - 1#)
    Next K
    X(N - 1#) = D(N - 1#) / B(N - 1#)
    For K = N - 2# To 0# Step -1
        X(K) = (D(K) - C(K) * X(K + 1#)) / B(K)
    Next K
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Internal subroutine. Three-point differentiation
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DiffThreePoint(ByVal T As Double, _
         ByVal X0 As Double, _
         ByVal F0 As Double, _
         ByVal X1 As Double, _
         ByVal F1 As Double, _
         ByVal X2 As Double, _
         ByVal F2 As Double) As Double
    Dim Result As Double
    Dim A As Double
    Dim B As Double

    T = T - X0
    X1 = X1 - X0
    X2 = X2 - X0
    A = (F2 - F0 - X2 / X1 * (F1 - F0)) / (Square(X2) - X1 * X2)
    B = (F1 - F0 - A * Square(X1)) / X1
    Result = 2# * A * T + B

    DiffThreePoint = Result
End Function




