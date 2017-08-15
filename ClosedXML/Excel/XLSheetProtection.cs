using System;

namespace ClosedXML.Excel
{
    internal class XLSheetProtection: IXLSheetProtection
    {
        public XLSheetProtection()
        {
            SelectLockedCells = true;
            SelectUnlockedCells = true;
        }

        public Boolean Protected { get; set; }
        internal String PasswordHash { get; set; }
        public String Password 
        {
            set
            {
                PasswordHash = GetPasswordHash(value);
            }
        }

        public Boolean AutoFilter { get; set; }
        public Boolean DeleteColumns { get; set; }
        public Boolean DeleteRows { get; set; }
        public Boolean FormatCells { get; set; }
        public Boolean FormatColumns { get; set; }
        public Boolean FormatRows { get; set; }
        public Boolean InsertColumns { get; set; }
        public Boolean InsertHyperlinks { get; set; }
        public Boolean InsertRows { get; set; }
        public Boolean Objects { get; set; }
        public Boolean PivotTables { get; set; }
        public Boolean Scenarios { get; set; }
        public Boolean SelectLockedCells { get; set; }
        public Boolean SelectUnlockedCells { get; set; }
        public Boolean Sort { get; set; }

        public IXLSheetProtection SetAutoFilter() { AutoFilter = true; return this; }	public IXLSheetProtection SetAutoFilter(Boolean value) { AutoFilter = value; return this; }
        public IXLSheetProtection SetDeleteColumns() { DeleteColumns = true; return this; }	public IXLSheetProtection SetDeleteColumns(Boolean value) { DeleteColumns = value; return this; }
        public IXLSheetProtection SetDeleteRows() { DeleteRows = true; return this; }	public IXLSheetProtection SetDeleteRows(Boolean value) { DeleteRows = value; return this; }
        public IXLSheetProtection SetFormatCells() { FormatCells = true; return this; }	public IXLSheetProtection SetFormatCells(Boolean value) { FormatCells = value; return this; }
        public IXLSheetProtection SetFormatColumns() { FormatColumns = true; return this; }	public IXLSheetProtection SetFormatColumns(Boolean value) { FormatColumns = value; return this; }
        public IXLSheetProtection SetFormatRows() { FormatRows = true; return this; }	public IXLSheetProtection SetFormatRows(Boolean value) { FormatRows = value; return this; }
        public IXLSheetProtection SetInsertColumns() { InsertColumns = true; return this; }	public IXLSheetProtection SetInsertColumns(Boolean value) { InsertColumns = value; return this; }
        public IXLSheetProtection SetInsertHyperlinks() { InsertHyperlinks = true; return this; }	public IXLSheetProtection SetInsertHyperlinks(Boolean value) { InsertHyperlinks = value; return this; }
        public IXLSheetProtection SetInsertRows() { InsertRows = true; return this; }	public IXLSheetProtection SetInsertRows(Boolean value) { InsertRows = value; return this; }
        public IXLSheetProtection SetObjects() { Objects = true; return this; }	public IXLSheetProtection SetObjects(Boolean value) { Objects = value; return this; }
        public IXLSheetProtection SetPivotTables() { PivotTables = true; return this; }	public IXLSheetProtection SetPivotTables(Boolean value) { PivotTables = value; return this; }
        public IXLSheetProtection SetScenarios() { Scenarios = true; return this; }	public IXLSheetProtection SetScenarios(Boolean value) { Scenarios = value; return this; }
        public IXLSheetProtection SetSelectLockedCells() { SelectLockedCells = true; return this; }	public IXLSheetProtection SetSelectLockedCells(Boolean value) { SelectLockedCells = value; return this; }
        public IXLSheetProtection SetSelectUnlockedCells() { SelectUnlockedCells = true; return this; }	public IXLSheetProtection SetSelectUnlockedCells(Boolean value) { SelectUnlockedCells = value; return this; }
        public IXLSheetProtection SetSort() { Sort = true; return this; }	public IXLSheetProtection SetSort(Boolean value) { Sort = value; return this; }

        public IXLSheetProtection Protect()
        {
            return Protect(String.Empty);
        }

        public IXLSheetProtection Protect(String password)
        {
            if (Protected)
            {
                throw new InvalidOperationException("The worksheet is already protected");
            }
            else
            {
                Protected = true;
                PasswordHash = GetPasswordHash(password);
            }
            return this;
        }

        public IXLSheetProtection Unprotect()
        {
            return Unprotect(String.Empty);
        }

        public IXLSheetProtection Unprotect(String password)
        {
            if (Protected)
            {
                String hash = GetPasswordHash(password);
                if (hash != PasswordHash)
                    throw new ArgumentException("Invalid password");
                else
                {
                    Protected = false;
                    PasswordHash = String.Empty;
                }
            }

            return this;
        }

        private String GetPasswordHash(String password)
        {
            Int32 pLength = password.Length;
            Int32 hash = 0;
            if (pLength == 0) return String.Empty;

            for (Int32 i = pLength - 1; i >= 0; i--)
            {
                hash ^= password[i];
                hash = hash >> 14 & 0x01 | hash << 1 & 0x7fff;
            }
            hash ^= 0x8000 | 'N' << 8 | 'K';
            hash ^= pLength;
            return hash.ToString("X");
        }
    }
}
