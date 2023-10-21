*********
Functions
*********

ClosedXML can evaluate formula functions. 

.. note::
   Excel has a a list of functions that are defined in ECMA-376 and a newer
   ones that are added in some subsequent version (future functions).
   The future functions generally have a prefix ``_xlfn``. The prefix is hidden
   in the GUI, but is present in the file (e.g. ``_xlfn.CONCAT(A1:A2)`` is
   displayed as a ``=CONCAT(A1:B1)`` in the Excel).
   
   The cell formula that uses a future functions that were added in later
   version of Excel must use a correct name of a function, including the prefix.

   .. code-block:: csharp

      ws.Cell(1,1).FormulaA1 = "_xlfn.CONCAT(A1:A2)";

   **Excel won't recognize future functions without a prefix!** It will try
   to match the function, but won't find anything and it will display
   a ``#NAME?`` error.

   See the `list of future functions <https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/5d1b6d44-6fc1-4ecd-8fef-0b27406cc2bf>`_


.. note::
   ClosedXML doesn't calculate and save values of a formula cells by default.
   The saved cell contains only formula and when the file is opened in the Excel,
   it recalculates values of formulas.
   
   You can save values by setting ``SaveOptions.EvaluateFormulasBeforeSaving``
   to ``true`` and passing the options to the ``XLWorkbook.SaveAs`` or
   ``XLWorkbook.Save`` method.
   
   Workbook without formula values can exhibit slighly odd behavior in some cases:
   
   * ``IXLCell.Style.Alignment.WrapText`` doesn't correctly auto-size cell height
     when opened in Excel (#1833).

Standard functions
##################

.. flat-table:: Standard function implementation status
   :header-rows: 1

   * - Category
     - Function
     - Implemented

   * - :rspan:`6` Cube
     - CUBEKPIMEMBER
     - No

   * - CUBEMEMBER
     - No

   * - CUBEMEMBERPROPERTY
     - No

   * - CUBERANKEDMEMBER
     - No

   * - CUBESET
     - No

   * - CUBESETCOUNT
     - No

   * - CUBEVALUE
     - No

   * - :rspan:`11` Database
     - DAVERAGE
     - No

   * - DCOUNT
     - No

   * - DCOUNTA
     - No

   * - DGET
     - No

   * - DMAX
     - No

   * - DMIN
     - No

   * - DPRODUCT
     - No

   * - DSTDEV
     - No

   * - DSTDEVP
     - No

   * - DSUM
     - No

   * - DVAR
     - No

   * - DVARP
     - No

   * - :rspan:`22` Date and Time
     - DATE
     - **Yes**

   * - DATEDIF
     - **Yes**

   * - DATEVALUE
     - **Yes**

   * - DAY
     - **Yes**

   * - DAYS360
     - **Yes**

   * - EDATE
     - **Yes**

   * - EOMONTH
     - **Yes**

   * - HOUR
     - **Yes**

   * - MINUTE
     - **Yes**

   * - MONTH
     - **Yes**

   * - NETWORKDAYS
     - **Yes**

   * - NETWORKDAYS.INTL
     - No

   * - NOW
     - **Yes**

   * - SECOND
     - **Yes**

   * - TIME
     - **Yes**

   * - TIMEVALUE
     - **Yes**

   * - TODAY
     - **Yes**

   * - WEEKDAY
     - **Yes**

   * - WEEKNUM
     - **Yes**

   * - WORKDAY
     - **Yes**

   * - WORKDAY.INTL
     - No

   * - YEAR
     - **Yes**

   * - YEARFRAC
     - **Yes**

   * - :rspan:`38` Engineering
     - BESSELI
     - No

   * - BESSELJ
     - No

   * - BESSELK
     - No

   * - BESSELY
     - No

   * - BIN2DEC
     - No

   * - BIN2HEX
     - No

   * - BIN2OCT
     - No

   * - COMPLEX
     - No

   * - CONVERT
     - No

   * - DEC2BIN
     - No

   * - DEC2HEX
     - No

   * - DEC2OCT
     - No

   * - DELTA
     - No

   * - ERF
     - No

   * - ERFC
     - No

   * - GESTEP
     - No

   * - HEX2BIN
     - No

   * - HEX2DEC
     - No

   * - HEX2OCT
     - No

   * - IMABS
     - No

   * - IMAGINARY
     - No

   * - IMARGUMENT
     - No

   * - IMCONJUGATE
     - No

   * - IMCOS
     - No

   * - IMDIV
     - No

   * - IMEXP
     - No

   * - IMLN
     - No

   * - IMLOG10
     - No

   * - IMLOG2
     - No

   * - IMPOWER
     - No

   * - IMPRODUCT
     - No

   * - IMREAL
     - No

   * - IMSIN
     - No

   * - IMSQRT
     - No

   * - IMSUB
     - No

   * - IMSUM
     - No

   * - OCT2BIN
     - No

   * - OCT2DEC
     - No

   * - OCT2HEX
     - No

   * - :rspan:`52` Financial
     - ACCRINT
     - No

   * - ACCRINTM
     - No

   * - AMORDEGRC
     - No

   * - AMORLINC
     - No

   * - COUPDAYBS
     - No

   * - COUPDAYS
     - No

   * - COUPDAYSNC
     - No

   * - COUPNCD
     - No

   * - COUPNUM
     - No

   * - COUPPCD
     - No

   * - CUMIPMT
     - No

   * - CUMPRINC
     - No

   * - DB
     - No

   * - DDB
     - No

   * - DISC
     - No

   * - DOLLARDE
     - No

   * - DOLLARFR
     - No

   * - DURATION
     - No

   * - EFFECT
     - No

   * - FV
     - YES

   * - FVSCHEDULE
     - No

   * - INTRATE
     - No

   * - IPMT
     - YES

   * - IRR
     - No

   * - ISPMT
     - No

   * - MDURATION
     - No

   * - MIRR
     - No

   * - NOMINAL
     - No

   * - NPER
     - No

   * - NPV
     - No

   * - ODDFPRICE
     - No

   * - ODDFYIELD
     - No

   * - ODDLPRICE
     - No

   * - ODDLYIELD
     - No

   * - PMT
     - **Yes**

   * - PPMT
     - No

   * - PRICE
     - No

   * - PRICEDISC
     - No

   * - PRICEMAT
     - No

   * - PV
     - No

   * - RATE
     - No

   * - RECEIVED
     - No

   * - SLN
     - No

   * - SYD
     - No

   * - TBILLEQ
     - No

   * - TBILLPRICE
     - No

   * - TBILLYIELD
     - No

   * - VDB
     - No

   * - XIRR
     - No

   * - XNPV
     - No

   * - YIELD
     - No

   * - YIELDDISC
     - No

   * - YIELDMAT
     - No

   * - :rspan:`16` Information
     - CELL
     - No

   * - ERROR.TYPE
     - **YES**

   * - INFO
     - No

   * - ISBLANK
     - **YES**

   * - ISERR
     - **YES**

   * - ISERROR
     - **YES**

   * - ISEVEN
     - **YES**

   * - ISLOGICAL
     - **YES**

   * - ISNA
     - **YES**

   * - ISNONTEXT
     - **YES**

   * - ISNUMBER
     - **YES**

   * - ISODD
     - **YES**

   * - ISREF
     - **YES**

   * - ISTEXT
     - **YES**

   * - N
     - **YES**

   * - NA
     - **YES**

   * - TYPE
     - **YES**

   * - :rspan:`6` Logical
     - AND
     - **YES**

   * - FALSE
     - **YES**

   * - IF
     - **YES**

   * - IFERROR
     - **YES**

   * - NOT
     - **YES**

   * - OR
     - **YES**

   * - TRUE
     - **YES**

   * - :rspan:`17` Lookup and Reference
     - ADDRESS
     - No

   * - AREAS
     - No

   * - CHOOSE
     - No

   * - COLUMN
     - **YES**

   * - COLUMNS
     - **YES**

   * - GETPIVOTDATA
     - No

   * - HLOOKUP
     - **YES**

   * - HYPERLINK
     - **YES**

   * - INDEX
     - **YES**

   * - INDIRECT
     - No

   * - LOOKUP
     - No

   * - MATCH
     - **YES**

   * - OFFSET
     - No

   * - ROW
     - **YES**

   * - ROWS
     - **YES**

   * - RTD
     - No

   * - TRANSPOSE
     - No

   * - VLOOKUP
     - **YES**

   * - :rspan:`61` Math and Trig
     - ABS
     - **YES**

   * - ACOS
     - **YES**

   * - ACOSH
     - **YES**

   * - ASIN
     - **YES**

   * - ASINH
     - **YES**

   * - ATAN
     - **YES**

   * - ATAN2
     - **YES**

   * - ATANH
     - **YES**

   * - CEILING
     - **YES**

   * - COMBIN
     - **YES**

   * - COS
     - **YES**

   * - COSH
     - **YES**

   * - DEGREES
     - **YES**

   * - ECMA.CEILING
     - No

   * - EVEN
     - **YES**

   * - EXP
     - **YES**

   * - FACT
     - **YES**

   * - FACTDOUBLE
     - **YES**

   * - FLOOR
     - **YES**

   * - GCD
     - **YES**

   * - INT
     - **YES**

   * - ISO.CEILING
     - No

   * - LCM
     - **YES**

   * - LN
     - **YES**

   * - LOG
     - **YES**

   * - LOG10
     - **YES**

   * - MDETERM
     - **YES**

   * - MINVERSE
     - **YES**

   * - MMULT
     - **YES**

   * - MOD
     - **YES**

   * - MROUND
     - **YES**

   * - MULTINOMIAL
     - **YES**

   * - ODD
     - **YES**

   * - PI
     - **YES**

   * - POWER
     - **YES**

   * - PRODUCT
     - **YES**

   * - QUOTIENT
     - **YES**

   * - RADIANS
     - **YES**

   * - RAND
     - **YES**

   * - RANDBETWEEN
     - **YES**

   * - ROMAN
     - **YES**

   * - ROUND
     - **YES**

   * - ROUNDDOWN
     - **YES**

   * - ROUNDUP
     - **YES**

   * - SERIESSUM
     - **YES**

   * - SIGN
     - **YES**

   * - SIN
     - **YES**

   * - SINH
     - **YES**

   * - SQRT
     - **YES**

   * - SQRTPI
     - **YES**

   * - SUBTOTAL
     - **YES**

   * - SUM
     - **YES**

   * - SUMIF
     - **YES**

   * - SUMIFS
     - **YES**

   * - SUMPRODUCT
     - **YES**

   * - SUMSQ
     - **YES**

   * - SUMX2MY2
     - No

   * - SUMX2PY2
     - No

   * - SUMXMY2
     - No

   * - TAN
     - **YES**

   * - TANH
     - **YES**

   * - TRUNC
     - **YES**

   * - :rspan:`82` Statistical
     - AVEDEV
     - No

   * - AVERAGE
     - **YES**

   * - AVERAGEA
     - **YES**

   * - AVERAGEIF
     - No

   * - AVERAGEIFS
     - No

   * - BETADIST
     - No

   * - BETAINV
     - No

   * - BINOMDIST
     - No

   * - CHIDIST
     - No

   * - CHIINV
     - No

   * - CHITEST
     - No

   * - CONFIDENCE
     - No

   * - CORREL
     - No

   * - COUNT
     - **YES**

   * - COUNTA
     - **YES**

   * - COUNTBLANK
     - **YES**

   * - COUNTIF
     - **YES**

   * - COUNTIFS
     - **YES**

   * - COVAR
     - No

   * - CRITBINOM
     - No

   * - DEVSQ
     - **YES**

   * - EXPONDIST
     - No

   * - FDIST
     - No

   * - FINV
     - No

   * - FISHER
     - **YES**

   * - FISHERINV
     - No

   * - FORECAST
     - No

   * - FREQUENCY
     - No

   * - FTEST
     - No

   * - GAMMADIST
     - No

   * - GAMMAINV
     - No

   * - GAMMALN
     - No

   * - GEOMEAN
     - **YES**

   * - GROWTH
     - No

   * - HARMEAN
     - No

   * - HYPGEOMDIST
     - No

   * - INTERCEPT
     - No

   * - KURT
     - No

   * - LARGE
     - No

   * - LINEST
     - No

   * - LOGEST
     - No

   * - LOGINV
     - No

   * - LOGNORMDIST
     - No

   * - MAX
     - **YES**

   * - MAXA
     - **YES**

   * - MEDIAN
     - **YES**

   * - MIN
     - **YES**

   * - MINA
     - **YES**

   * - MODE
     - No

   * - NEGBINOMDIST
     - No

   * - NORMDIST
     - No

   * - NORMINV
     - No

   * - NORMSDIST
     - No

   * - NORMSINV
     - No

   * - PEARSON
     - No

   * - PERCENTILE
     - No

   * - PERCENTRANK
     - No

   * - PERMUT
     - No

   * - POISSON
     - No

   * - PROB
     - No

   * - QUARTILE
     - No

   * - RANK
     - No

   * - RSQ
     - No

   * - SKEW
     - No

   * - SLOPE
     - No

   * - SMALL
     - No

   * - STANDARDIZE
     - No

   * - STDEV
     - **YES**

   * - STDEVA
     - **YES**

   * - STDEVP
     - **YES**

   * - STDEVPA
     - **YES**

   * - STEYX
     - No

   * - TDIST
     - No

   * - TINV
     - No

   * - TREND
     - No

   * - TRIMMEAN
     - No

   * - TTEST
     - No

   * - VAR
     - **YES**

   * - VARA
     - **YES**

   * - VARP
     - **YES**

   * - VARPA
     - **YES**

   * - WEIBULL
     - No

   * - ZTEST
     - No

   * - :rspan:`33` Text and Data
     - ASC
     - **YES**

   * - BAHTTEXT
     - No

   * - CHAR
     - **YES**

   * - CLEAN
     - **YES**

   * - CODE
     - **YES**

   * - CONCATENATE
     - **YES**

   * - DOLLAR
     - **YES**

   * - EXACT
     - **YES**

   * - FIND
     - **YES**

   * - FINDB
     - No

   * - FIXED
     - **YES**

   * - JIS
     - No

   * - LEFT
     - **YES**

   * - LEFTB
     - No

   * - LEN
     - **YES**

   * - LENB
     - No

   * - LOWER
     - **YES**

   * - MID
     - **YES**

   * - MIDB
     - No

   * - PHONETIC
     - No

   * - PROPER
     - **YES**

   * - REPLACE
     - **YES**

   * - REPLACEB
     - No

   * - REPT
     - **YES**

   * - RIGHT
     - **YES**

   * - RIGHTB
     - No

   * - SEARCH
     - **YES**

   * - SEARCHB
     - No

   * - SUBSTITUTE
     - **YES**

   * - T
     - **YES**

   * - TEXT
     - **YES**

   * - TRIM
     - **YES**

   * - UPPER
     - **YES**

   * - VALUE
     - **YES**


Future functions
################

.. flat-table:: Future functions
   :header-rows: 1

   * - Category
     - Function
     - Implemented

   * - Date and Time
     - _xlfn.ISOWEEKNUM
     - **Yes**

   * - :rspan:`13` Math and Trig
     - _xlfn.ACOT
     - **Yes**

   * - _xlfn.ACOTH
     - **Yes**

   * - _xlfn.ARABIC
     - **Yes**

   * - _xlfn.BASE
     - **Yes**

   * - _xlfn.CEILING.MATH
     - **Yes**

   * - _xlfn.COMBINA
     - **Yes**

   * - _xlfn.COT
     - **Yes**

   * - _xlfn.COTH
     - **Yes**

   * - _xlfn.CSC
     - **Yes**

   * - _xlfn.CSCH
     - **Yes**

   * - _xlfn.DECIMAL
     - **Yes**

   * - _xlfn.FLOOR.MATH
     - **Yes**

   * - _xlfn.SEC
     - **Yes**

   * - _xlfn.SECH
     - **Yes**

   * - :rspan:`1` Statistical
     - _xlfn.STDEV.S
     - **Yes**

   * - _xlfn.STDEV.P
     - **Yes**

   * - :rspan:`6` Text and Data
     - _xlfn.CONCAT
     - **Yes**

   * - _xlfn.NUMBERVALUE
     - **Yes**

   * - _xlfn.TEXTJOIN
     - **Yes**
