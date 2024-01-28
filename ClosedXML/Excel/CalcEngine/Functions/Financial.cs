using System;
using static ClosedXML.Excel.CalcEngine.Functions.SignatureAdapter;

namespace ClosedXML.Excel.CalcEngine
{
    internal static class Financial
    {
        public static void Register(FunctionRegistry ce)
        {
            // ACCRINT Returns the accrued interest for a security that pays periodic interest
            // ACCRINTM Returns the accrued interest for a security that pays interest at maturity
            // AMORDEGRC Returns the depreciation for each accounting period by using a depreciation coefficient
            // AMORLINC Returns the depreciation for each accounting period
            // COUPDAYBS Returns the number of days from the beginning of the coupon period to the settlement date
            // COUPDAYS Returns the number of days in the coupon period that contains the settlement date
            // COUPDAYSNC Returns the number of days from the settlement date to the next coupon date
            // COUPNCD Returns the next coupon date after the settlement date
            // COUPNUM Returns the number of coupons payable between the settlement date and maturity date
            // COUPPCD Returns the previous coupon date before the settlement date
            // CUMIPMT Returns the cumulative interest paid between two periods
            // CUMPRINC Returns the cumulative principal paid on a loan between two periods
            // DB Returns the depreciation of an asset for a specified period by using the fixed-declining balance method
            // DDB Returns the depreciation of an asset for a specified period by using the double-declining balance method or some other method that you specify
            // DISC Returns the discount rate for a security
            // DOLLARDE Converts a dollar price, expressed as a fraction, into a dollar price, expressed as a decimal number
            // DOLLARFR Converts a dollar price, expressed as a decimal number, into a dollar price, expressed as a fraction
            // DURATION Returns the annual duration of a security with periodic interest payments
            // EFFECT Returns the effective annual interest rate
            ce.RegisterFunction("FV", 3, 5, AdaptLastTwoOptional(Fv, 0, 0), FunctionFlags.Scalar); // Returns the future value of an investment
            // FVSCHEDULE Returns the future value of an initial principal after applying a series of compound interest rates
            // INTRATE Returns the interest rate for a fully invested security
            ce.RegisterFunction("IPMT", 4, 6, AdaptLastTwoOptional(Ipmt, 0, 0), FunctionFlags.Scalar); // Returns the interest payment for an investment for a given period
            // IRR Returns the internal rate of return for a series of cash flows
            // ISPMT Calculates the interest paid during a specific period of an investment
            // MDURATION Returns the Macauley modified duration for a security with an assumed par value of $100
            // MIRR Returns the internal rate of return where positive and negative cash flows are financed at different rates
            // NOMINAL Returns the annual nominal interest rate
            // NPER Returns the number of periods for an investment
            // NPV Returns the net present value of an investment based on a series of periodic cash flows and a discount rate
            // ODDFPRICE Returns the price per $100 face value of a security with an odd first period
            // ODDFYIELD Returns the yield of a security with an odd first period
            // ODDLPRICE Returns the price per $100 face value of a security with an odd last period
            // ODDLYIELD Returns the yield of a security with an odd last period
            // PDURATION Returns the number of periods required by an investment to reach a specified value
            ce.RegisterFunction("PMT", 3, 5, AdaptLastTwoOptional(Pmt, 0, 0), FunctionFlags.Scalar); // Returns the periodic payment for an annuity
            // PPMT Returns the payment on the principal for an investment for a given period
            // PRICE Returns the price per $100 face value of a security that pays periodic interest
            // PRICEDISC Returns the price per $100 face value of a discounted security
            // PRICEMAT Returns the price per $100 face value of a security that pays interest at maturity
            // PV Returns the present value of an investment
            // RATE Returns the interest rate per period of an annuity
            // RECEIVED Returns the amount received at maturity for a fully invested security
            // RRI Returns an equivalent interest rate for the growth of an investment
            // SLN Returns the straight-line depreciation of an asset for one period
            // SYD Returns the sum-of-years' digits depreciation of an asset for a specified period
            // TBILLEQ Returns the bond-equivalent yield for a Treasury bill
            // TBILLPRICE Returns the price per $100 face value for a Treasury bill
            // TBILLYIELD Returns the yield for a Treasury bill
            // VDB Returns the depreciation of an asset for a specified or partial period by using a declining balance method
            // XIRR Returns the internal rate of return for a schedule of cash flows that is not necessarily periodic
            // XNPV Returns the net present value for a schedule of cash flows that is not necessarily periodic
            // YIELD Returns the yield on a security that pays periodic interest
            // YIELDDISC Returns the annual yield for a discounted security; for example, a Treasury bill
            // YIELDMAT Returns the annual yield of a security that pays interest at maturity
        }

        private static AnyValue Fv(double rate, double numberOfPayments, double pmt, double presentValue, double type)
        {
            if (numberOfPayments == 0)
                return -presentValue;

            return FvInternal(rate, numberOfPayments, pmt, presentValue, type);
        }

        private static double FvInternal(double rate, double numberOfPayments, double pmt, double presentValue, double type)
        {
            if (rate == 0.0)
                return -(pmt * numberOfPayments + presentValue);

            if (type != 0.0)
                pmt *= (1 + rate);

            return -(pmt * (Math.Pow(1 + rate, numberOfPayments) - 1) / rate + presentValue * Math.Pow(1 + rate, numberOfPayments));
        }

        private static AnyValue Ipmt(double rate, double period, double numberOfPayments, double presentValue, double futureValue, double type)
        {
            if (numberOfPayments <= 0 || rate <= -1)
                return XLError.NumberInvalid;

            numberOfPayments = Math.Ceiling(numberOfPayments);

            if (period < 1 || period > numberOfPayments)
                return XLError.NumberInvalid;

            double ipmt = FvInternal(rate, period - 1, PmtInternal(rate, numberOfPayments, presentValue, futureValue, type), presentValue, type) * rate;

            if (type != 0.0)
                ipmt /= (1 + rate);

            return ipmt;
        }

        private static AnyValue Pmt(double rate, double numberOfPayments, double presentValue, double futureValue, double type)
        {
            if (numberOfPayments == 0 || rate <= -1)
                return XLError.NumberInvalid;

            return PmtInternal(rate, numberOfPayments, presentValue, futureValue, type);
        }

        private static double PmtInternal(double rate, double numberOfPayments, double presentValue, double futureValue, double type)
        {
            if (rate == 0.0)
                return -(presentValue + futureValue) / numberOfPayments;

            const int paymentAtTheEndOfPeriod = 0;
            const int paymentAtTheBeginningOfPeriod = 1;
            var timingOffset = type != 0.0 ? paymentAtTheBeginningOfPeriod : paymentAtTheEndOfPeriod;

            return (-futureValue - presentValue * Math.Pow(1.0 + rate, numberOfPayments)) /
               (1 + rate * timingOffset) / ((Math.Pow(1.0 + rate, numberOfPayments) - 1) / rate);
        }
    }
}
