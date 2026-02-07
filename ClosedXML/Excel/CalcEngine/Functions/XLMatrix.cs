#nullable disable

using System;

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

        public Boolean IsSingular()
        {
            for (var row = 0; row < rows; row++)
            {
                for (var col = 0; col < cols; col++)
                {
                    var element = mat[row, col];
                    if (double.IsNaN(element) || double.IsInfinity(element))
                        return true;
                }
            }

            return false;
        }

        public Boolean IsSquare()
        {
            return (rows == cols);
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

        public XLMatrix SolveWith(XLMatrix v) // Function solves Ax = v in conformity with solution vector "v"
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

        public XLMatrix ZeroMatrix(int iRows, int iCols) // Function generates the zero matrix
        {
            var matrix = new XLMatrix(iRows, iCols);
            for (var i = 0; i < iRows; i++)
            {
                for (var j = 0; j < iCols; j++)
                {
                    matrix[i, j] = 0;
                }
            }

            return matrix;
        }

        public XLMatrix IdentityMatrix(int iRows, int iCols) // Function generates the identity matrix
        {
            var matrix = ZeroMatrix(iRows, iCols);
            for (var i = 0; i < Math.Min(iRows, iCols); i++)
            {
                matrix[i, i] = 1;
            }

            return matrix;
        }
    }
}
