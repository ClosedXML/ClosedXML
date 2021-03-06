using System;
using System.Text.RegularExpressions;

namespace ClosedXML.Excel.CalcEngine.Functions
{
    internal class XLMatrix
    {
        public XLMatrix L;
        public XLMatrix U;
        public int cols;
        private double detOfP = 1;
        public double[,] mat;
        private int[] pi;
        public int rows;

        public XLMatrix(int iRows, int iCols) // XLMatrix Class constructor
        {
            rows = iRows;
            cols = iCols;
            mat = new double[rows, cols];
        }

        public XLMatrix(Double[,] arr)
            : this(arr.GetLength(0), arr.GetLength(1))
        {
            var roCount = arr.GetLength(0);
            var coCount = arr.GetLength(1);
            for (int ro = 0; ro < roCount; ro++)
            {
                for (int co = 0; co < coCount; co++)
                {
                    mat[ro, co] = arr[ro, co];
                }
            }
        }

        public double this[int iRow, int iCol] // Access this matrix as a 2D array
        {
            get { return mat[iRow, iCol]; }
            set { mat[iRow, iCol] = value; }
        }

        public Boolean IsSquare()
        {
            return (rows == cols);
        }

        public XLMatrix GetCol(int k)
        {
            var m = new XLMatrix(rows, 1);
            for (var i = 0; i < rows; i++) m[i, 0] = mat[i, k];
            return m;
        }

        public void SetCol(XLMatrix v, int k)
        {
            for (var i = 0; i < rows; i++) mat[i, k] = v[i, 0];
        }

        public void MakeLU() // Function for LU decomposition
        {
            if (!IsSquare()) throw new InvalidOperationException("The matrix is not square!");
            L = IdentityMatrix(rows, cols);
            U = Duplicate();

            pi = new int[rows];
            for (var i = 0; i < rows; i++) pi[i] = i;

            var k0 = 0;

            for (var k = 0; k < cols - 1; k++)
            {
                double p = 0;
                for (var i = k; i < rows; i++) // find the row with the biggest pivot
                {
                    if (Math.Abs(U[i, k]) > p)
                    {
                        p = Math.Abs(U[i, k]);
                        k0 = i;
                    }
                }
                if (p == 0)
                    throw new InvalidOperationException("The matrix is singular!");

                var pom1 = pi[k];
                pi[k] = pi[k0];
                pi[k0] = pom1; // switch two rows in permutation matrix

                double pom2;
                for (var i = 0; i < k; i++)
                {
                    pom2 = L[k, i];
                    L[k, i] = L[k0, i];
                    L[k0, i] = pom2;
                }

                if (k != k0) detOfP *= -1;

                for (var i = 0; i < cols; i++) // Switch rows in U
                {
                    pom2 = U[k, i];
                    U[k, i] = U[k0, i];
                    U[k0, i] = pom2;
                }

                for (var i = k + 1; i < rows; i++)
                {
                    L[i, k] = U[i, k] / U[k, k];
                    for (var j = k; j < cols; j++)
                        U[i, j] = U[i, j] - L[i, k] * U[k, j];
                }
            }
        }

        public XLMatrix SolveWith(XLMatrix v) // Function solves Ax = v in confirmity with solution vector "v"
        {
            if (rows != cols) throw new InvalidOperationException("The matrix is not square!");
            if (rows != v.rows) throw new ArgumentException("Wrong number of results in solution vector!");
            if (L == null) MakeLU();

            var b = new XLMatrix(rows, 1);
            for (var i = 0; i < rows; i++) b[i, 0] = v[pi[i], 0]; // switch two items in "v" due to permutation matrix

            var z = SubsForth(L, b);
            var x = SubsBack(U, z);

            return x;
        }

        public XLMatrix Invert() // Function returns the inverted matrix
        {
            if (L == null) MakeLU();

            var inv = new XLMatrix(rows, cols);

            for (var i = 0; i < rows; i++)
            {
                var Ei = ZeroMatrix(rows, 1);
                Ei[i, 0] = 1;
                var col = SolveWith(Ei);
                inv.SetCol(col, i);
            }
            return inv;
        }

        public double Determinant() // Function for determinant
        {
            if (L == null) MakeLU();
            var det = detOfP;
            for (var i = 0; i < rows; i++) det *= U[i, i];
            return det;
        }

        public XLMatrix GetP() // Function returns permutation matrix "P" due to permutation vector "pi"
        {
            if (L == null) MakeLU();

            var matrix = ZeroMatrix(rows, cols);
            for (var i = 0; i < rows; i++) matrix[pi[i], i] = 1;
            return matrix;
        }

        public XLMatrix Duplicate() // Function returns the copy of this matrix
        {
            var matrix = new XLMatrix(rows, cols);
            for (var i = 0; i < rows; i++)
                for (var j = 0; j < cols; j++)
                    matrix[i, j] = mat[i, j];
            return matrix;
        }

        public static XLMatrix SubsForth(XLMatrix A, XLMatrix b) // Function solves Ax = b for A as a lower triangular matrix
        {
            if (A.L == null) A.MakeLU();
            var n = A.rows;
            var x = new XLMatrix(n, 1);

            for (var i = 0; i < n; i++)
            {
                x[i, 0] = b[i, 0];
                for (var j = 0; j < i; j++) x[i, 0] -= A[i, j] * x[j, 0];
                x[i, 0] = x[i, 0] / A[i, i];
            }
            return x;
        }

        public static XLMatrix SubsBack(XLMatrix A, XLMatrix b) // Function solves Ax = b for A as an upper triangular matrix
        {
            if (A.L == null) A.MakeLU();
            var n = A.rows;
            var x = new XLMatrix(n, 1);

            for (var i = n - 1; i > -1; i--)
            {
                x[i, 0] = b[i, 0];
                for (var j = n - 1; j > i; j--) x[i, 0] -= A[i, j] * x[j, 0];
                x[i, 0] = x[i, 0] / A[i, i];
            }
            return x;
        }

        public static XLMatrix ZeroMatrix(int iRows, int iCols) // Function generates the zero matrix
        {
            var matrix = new XLMatrix(iRows, iCols);
            for (var i = 0; i < iRows; i++)
                for (var j = 0; j < iCols; j++)
                    matrix[i, j] = 0;
            return matrix;
        }

        public static XLMatrix IdentityMatrix(int iRows, int iCols) // Function generates the identity matrix
        {
            var matrix = ZeroMatrix(iRows, iCols);
            for (var i = 0; i < Math.Min(iRows, iCols); i++)
                matrix[i, i] = 1;
            return matrix;
        }

        public static XLMatrix RandomMatrix(int iRows, int iCols, int dispersion) // Function generates the zero matrix
        {
            var random = new Random();
            var matrix = new XLMatrix(iRows, iCols);
            for (var i = 0; i < iRows; i++)
                for (var j = 0; j < iCols; j++)
                    matrix[i, j] = random.Next(-dispersion, dispersion);
            return matrix;
        }

        public static XLMatrix Parse(string ps) // Function parses the matrix from string
        {
            var s = NormalizeMatrixString(ps);
            var rows = Regex.Split(s, "\r\n");
            var nums = rows[0].Split(' ');
            var matrix = new XLMatrix(rows.Length, nums.Length);
            try
            {
                for (var i = 0; i < rows.Length; i++)
                {
                    nums = rows[i].Split(' ');
                    for (var j = 0; j < nums.Length; j++) matrix[i, j] = double.Parse(nums[j]);
                }
            }
            catch (FormatException fe)
            {
                throw new FormatException("Wrong input format!", fe);
            }
            return matrix;
        }

        public override string ToString() // Function returns matrix as a string
        {
            var s = "";
            for (var i = 0; i < rows; i++)
            {
                for (var j = 0; j < cols; j++) s += String.Format("{0,5:0.00}", mat[i, j]) + " ";
                s += "\r\n";
            }
            return s;
        }

        public static XLMatrix Transpose(XLMatrix m) // XLMatrix transpose, for any rectangular matrix
        {
            var t = new XLMatrix(m.cols, m.rows);
            for (var i = 0; i < m.rows; i++)
                for (var j = 0; j < m.cols; j++)
                    t[j, i] = m[i, j];
            return t;
        }

        public static XLMatrix Power(XLMatrix m, int pow) // Power matrix to exponent
        {
            if (pow == 0) return IdentityMatrix(m.rows, m.cols);
            if (pow == 1) return m.Duplicate();
            if (pow == -1) return m.Invert();

            XLMatrix x;
            if (pow < 0)
            {
                x = m.Invert();
                pow *= -1;
            }
            else x = m.Duplicate();

            var ret = IdentityMatrix(m.rows, m.cols);
            while (pow != 0)
            {
                if ((pow & 1) == 1) ret *= x;
                x *= x;
                pow >>= 1;
            }
            return ret;
        }

        private static void SafeAplusBintoC(XLMatrix A, int xa, int ya, XLMatrix B, int xb, int yb, XLMatrix C, int size)
        {
            for (var i = 0; i < size; i++) // rows
                for (var j = 0; j < size; j++) // cols
                {
                    C[i, j] = 0;
                    if (xa + j < A.cols && ya + i < A.rows) C[i, j] += A[ya + i, xa + j];
                    if (xb + j < B.cols && yb + i < B.rows) C[i, j] += B[yb + i, xb + j];
                }
        }

        private static void SafeAminusBintoC(XLMatrix A, int xa, int ya, XLMatrix B, int xb, int yb, XLMatrix C, int size)
        {
            for (var i = 0; i < size; i++) // rows
                for (var j = 0; j < size; j++) // cols
                {
                    C[i, j] = 0;
                    if (xa + j < A.cols && ya + i < A.rows) C[i, j] += A[ya + i, xa + j];
                    if (xb + j < B.cols && yb + i < B.rows) C[i, j] -= B[yb + i, xb + j];
                }
        }

        private static void SafeACopytoC(XLMatrix A, int xa, int ya, XLMatrix C, int size)
        {
            for (var i = 0; i < size; i++) // rows
                for (var j = 0; j < size; j++) // cols
                {
                    C[i, j] = 0;
                    if (xa + j < A.cols && ya + i < A.rows) C[i, j] += A[ya + i, xa + j];
                }
        }

        private static void AplusBintoC(XLMatrix A, int xa, int ya, XLMatrix B, int xb, int yb, XLMatrix C, int size)
        {
            for (var i = 0; i < size; i++) // rows
                for (var j = 0; j < size; j++) C[i, j] = A[ya + i, xa + j] + B[yb + i, xb + j];
        }

        private static void AminusBintoC(XLMatrix A, int xa, int ya, XLMatrix B, int xb, int yb, XLMatrix C, int size)
        {
            for (var i = 0; i < size; i++) // rows
                for (var j = 0; j < size; j++) C[i, j] = A[ya + i, xa + j] - B[yb + i, xb + j];
        }

        private static void ACopytoC(XLMatrix A, int xa, int ya, XLMatrix C, int size)
        {
            for (var i = 0; i < size; i++) // rows
                for (var j = 0; j < size; j++) C[i, j] = A[ya + i, xa + j];
        }

        private static XLMatrix StrassenMultiply(XLMatrix A, XLMatrix B) // Smart matrix multiplication
        {
            if (A.cols != B.rows) throw new ArgumentException("Wrong dimension of matrix!");

            XLMatrix R;

            var msize = Math.Max(Math.Max(A.rows, A.cols), Math.Max(B.rows, B.cols));

            if (msize < 32)
            {
                R = ZeroMatrix(A.rows, B.cols);
                for (var i = 0; i < R.rows; i++)
                    for (var j = 0; j < R.cols; j++)
                        for (var k = 0; k < A.cols; k++)
                            R[i, j] += A[i, k] * B[k, j];
                return R;
            }

            var size = 1;
            var n = 0;
            while (msize > size)
            {
                size *= 2;
                n++;
            }

            var h = size / 2;

            var mField = new XLMatrix[n, 9];

            /*
             *  8x8, 8x8, 8x8, ...
             *  4x4, 4x4, 4x4, ...
             *  2x2, 2x2, 2x2, ...
             *  . . .
             */

            for (var i = 0; i < n - 4; i++) // rows
            {
                var z = (int)Math.Pow(2, n - i - 1);
                for (var j = 0; j < 9; j++) mField[i, j] = new XLMatrix(z, z);
            }

            SafeAplusBintoC(A, 0, 0, A, h, h, mField[0, 0], h);
            SafeAplusBintoC(B, 0, 0, B, h, h, mField[0, 1], h);
            StrassenMultiplyRun(mField[0, 0], mField[0, 1], mField[0, 1 + 1], 1, mField); // (A11 + A22) * (B11 + B22);

            SafeAplusBintoC(A, 0, h, A, h, h, mField[0, 0], h);
            SafeACopytoC(B, 0, 0, mField[0, 1], h);
            StrassenMultiplyRun(mField[0, 0], mField[0, 1], mField[0, 1 + 2], 1, mField); // (A21 + A22) * B11;

            SafeACopytoC(A, 0, 0, mField[0, 0], h);
            SafeAminusBintoC(B, h, 0, B, h, h, mField[0, 1], h);
            StrassenMultiplyRun(mField[0, 0], mField[0, 1], mField[0, 1 + 3], 1, mField); //A11 * (B12 - B22);

            SafeACopytoC(A, h, h, mField[0, 0], h);
            SafeAminusBintoC(B, 0, h, B, 0, 0, mField[0, 1], h);
            StrassenMultiplyRun(mField[0, 0], mField[0, 1], mField[0, 1 + 4], 1, mField); //A22 * (B21 - B11);

            SafeAplusBintoC(A, 0, 0, A, h, 0, mField[0, 0], h);
            SafeACopytoC(B, h, h, mField[0, 1], h);
            StrassenMultiplyRun(mField[0, 0], mField[0, 1], mField[0, 1 + 5], 1, mField); //(A11 + A12) * B22;

            SafeAminusBintoC(A, 0, h, A, 0, 0, mField[0, 0], h);
            SafeAplusBintoC(B, 0, 0, B, h, 0, mField[0, 1], h);
            StrassenMultiplyRun(mField[0, 0], mField[0, 1], mField[0, 1 + 6], 1, mField); //(A21 - A11) * (B11 + B12);

            SafeAminusBintoC(A, h, 0, A, h, h, mField[0, 0], h);
            SafeAplusBintoC(B, 0, h, B, h, h, mField[0, 1], h);
            StrassenMultiplyRun(mField[0, 0], mField[0, 1], mField[0, 1 + 7], 1, mField); // (A12 - A22) * (B21 + B22);

            R = new XLMatrix(A.rows, B.cols); // result

            // C11
            for (var i = 0; i < Math.Min(h, R.rows); i++) // rows
                for (var j = 0; j < Math.Min(h, R.cols); j++) // cols
                    R[i, j] = mField[0, 1 + 1][i, j] + mField[0, 1 + 4][i, j] - mField[0, 1 + 5][i, j] +
                              mField[0, 1 + 7][i, j];

            // C12
            for (var i = 0; i < Math.Min(h, R.rows); i++) // rows
                for (var j = h; j < Math.Min(2 * h, R.cols); j++) // cols
                    R[i, j] = mField[0, 1 + 3][i, j - h] + mField[0, 1 + 5][i, j - h];

            // C21
            for (var i = h; i < Math.Min(2 * h, R.rows); i++) // rows
                for (var j = 0; j < Math.Min(h, R.cols); j++) // cols
                    R[i, j] = mField[0, 1 + 2][i - h, j] + mField[0, 1 + 4][i - h, j];

            // C22
            for (var i = h; i < Math.Min(2 * h, R.rows); i++) // rows
                for (var j = h; j < Math.Min(2 * h, R.cols); j++) // cols
                    R[i, j] = mField[0, 1 + 1][i - h, j - h] - mField[0, 1 + 2][i - h, j - h] +
                              mField[0, 1 + 3][i - h, j - h] + mField[0, 1 + 6][i - h, j - h];

            return R;
        }

        // function for square matrix 2^N x 2^N

        private static void StrassenMultiplyRun(XLMatrix A, XLMatrix B, XLMatrix C, int l, XLMatrix[,] f)
        // A * B into C, level of recursion, matrix field
        {
            var size = A.rows;
            var h = size / 2;

            if (size < 32)
            {
                for (var i = 0; i < C.rows; i++)
                    for (var j = 0; j < C.cols; j++)
                    {
                        C[i, j] = 0;
                        for (var k = 0; k < A.cols; k++) C[i, j] += A[i, k] * B[k, j];
                    }
                return;
            }

            AplusBintoC(A, 0, 0, A, h, h, f[l, 0], h);
            AplusBintoC(B, 0, 0, B, h, h, f[l, 1], h);
            StrassenMultiplyRun(f[l, 0], f[l, 1], f[l, 1 + 1], l + 1, f); // (A11 + A22) * (B11 + B22);

            AplusBintoC(A, 0, h, A, h, h, f[l, 0], h);
            ACopytoC(B, 0, 0, f[l, 1], h);
            StrassenMultiplyRun(f[l, 0], f[l, 1], f[l, 1 + 2], l + 1, f); // (A21 + A22) * B11;

            ACopytoC(A, 0, 0, f[l, 0], h);
            AminusBintoC(B, h, 0, B, h, h, f[l, 1], h);
            StrassenMultiplyRun(f[l, 0], f[l, 1], f[l, 1 + 3], l + 1, f); //A11 * (B12 - B22);

            ACopytoC(A, h, h, f[l, 0], h);
            AminusBintoC(B, 0, h, B, 0, 0, f[l, 1], h);
            StrassenMultiplyRun(f[l, 0], f[l, 1], f[l, 1 + 4], l + 1, f); //A22 * (B21 - B11);

            AplusBintoC(A, 0, 0, A, h, 0, f[l, 0], h);
            ACopytoC(B, h, h, f[l, 1], h);
            StrassenMultiplyRun(f[l, 0], f[l, 1], f[l, 1 + 5], l + 1, f); //(A11 + A12) * B22;

            AminusBintoC(A, 0, h, A, 0, 0, f[l, 0], h);
            AplusBintoC(B, 0, 0, B, h, 0, f[l, 1], h);
            StrassenMultiplyRun(f[l, 0], f[l, 1], f[l, 1 + 6], l + 1, f); //(A21 - A11) * (B11 + B12);

            AminusBintoC(A, h, 0, A, h, h, f[l, 0], h);
            AplusBintoC(B, 0, h, B, h, h, f[l, 1], h);
            StrassenMultiplyRun(f[l, 0], f[l, 1], f[l, 1 + 7], l + 1, f); // (A12 - A22) * (B21 + B22);

            // C11
            for (var i = 0; i < h; i++) // rows
                for (var j = 0; j < h; j++) // cols
                    C[i, j] = f[l, 1 + 1][i, j] + f[l, 1 + 4][i, j] - f[l, 1 + 5][i, j] + f[l, 1 + 7][i, j];

            // C12
            for (var i = 0; i < h; i++) // rows
                for (var j = h; j < size; j++) // cols
                    C[i, j] = f[l, 1 + 3][i, j - h] + f[l, 1 + 5][i, j - h];

            // C21
            for (var i = h; i < size; i++) // rows
                for (var j = 0; j < h; j++) // cols
                    C[i, j] = f[l, 1 + 2][i - h, j] + f[l, 1 + 4][i - h, j];

            // C22
            for (var i = h; i < size; i++) // rows
                for (var j = h; j < size; j++) // cols
                    C[i, j] = f[l, 1 + 1][i - h, j - h] - f[l, 1 + 2][i - h, j - h] + f[l, 1 + 3][i - h, j - h] +
                              f[l, 1 + 6][i - h, j - h];
        }

        public static XLMatrix StupidMultiply(XLMatrix m1, XLMatrix m2) // Stupid matrix multiplication
        {
            if (m1.cols != m2.rows) throw new ArgumentException("Wrong dimensions of matrix!");

            var result = ZeroMatrix(m1.rows, m2.cols);
            for (var i = 0; i < result.rows; i++)
                for (var j = 0; j < result.cols; j++)
                    for (var k = 0; k < m1.cols; k++)
                        result[i, j] += m1[i, k] * m2[k, j];
            return result;
        }

        private static XLMatrix Multiply(double n, XLMatrix m) // Multiplication by constant n
        {
            var r = new XLMatrix(m.rows, m.cols);
            for (var i = 0; i < m.rows; i++)
                for (var j = 0; j < m.cols; j++)
                    r[i, j] = m[i, j] * n;
            return r;
        }

        private static XLMatrix Add(XLMatrix m1, XLMatrix m2)
        {
            if (m1.rows != m2.rows || m1.cols != m2.cols)
                throw new ArgumentException("Matrices must have the same dimensions!");
            var r = new XLMatrix(m1.rows, m1.cols);
            for (var i = 0; i < r.rows; i++)
                for (var j = 0; j < r.cols; j++)
                    r[i, j] = m1[i, j] + m2[i, j];
            return r;
        }

        public static string NormalizeMatrixString(string matStr) // From Andy - thank you! :)
        {
            // Remove any multiple spaces
            while (matStr.IndexOf("  ") != -1)
                matStr = matStr.Replace("  ", " ");

            // Remove any spaces before or after newlines
            matStr = matStr.Replace(" \r\n", "\r\n");
            matStr = matStr.Replace("\r\n ", "\r\n");

            // If the data ends in a newline, remove the trailing newline.
            // Make it easier by first replacing \r\n’s with |’s then
            // restore the |’s with \r\n’s
            matStr = matStr.Replace("\r\n", "|");
            while (matStr.LastIndexOf("|") == (matStr.Length - 1))
                matStr = matStr.Substring(0, matStr.Length - 1);

            matStr = matStr.Replace("|", "\r\n");
            return matStr;
        }

        //   O P E R A T O R S

        public static XLMatrix operator -(XLMatrix m)
        {
            return Multiply(-1, m);
        }

        public static XLMatrix operator +(XLMatrix m1, XLMatrix m2)
        {
            return Add(m1, m2);
        }

        public static XLMatrix operator -(XLMatrix m1, XLMatrix m2)
        {
            return Add(m1, -m2);
        }

        public static XLMatrix operator *(XLMatrix m1, XLMatrix m2)
        {
            return StrassenMultiply(m1, m2);
        }

        public static XLMatrix operator *(double n, XLMatrix m)
        {
            return Multiply(n, m);
        }
    }
}
