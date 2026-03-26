package oleXL

import (
	"fmt"
	"log"

	"github.com/go-ole/go-ole"
)

type workFunction struct {
	app    *Excel
	parent any
	num    int
}

func (Q *workApp) WorksheetFunction() *workFunction {
	var body workFunction
	xl := Q.app

	name := "WorksheetFunction"
	core, num := xl.cores.FindAdd(name, xl.num)
	if core.disp == nil {
		cmd := "Get"

		ans, err := xl.cores.SendNum(cmd, name, xl.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}

		switch x := ans.(type) {
		case *ole.IDispatch:
			core.disp = x
			core.lock = 1 //Lock.on
		}
	}
	body.app = xl
	body.num = num
	body.parent = Q
	return &body
}

func (Q *workFunction) function(funcName string, args ...any) any {
	xl := Q.app

	var opt []any
	if len(args) > 0 {
		for i := range args {
			switch x := args[i].(type) {
			case string:
				opt = append(opt, args[i])
			case float64:
				opt = append(opt, args[i])
			case *workRange:
				core := xl.cores.getCore(x.num)
				opt = append(opt, core.disp)
			}
			if i > 30 {
				break
			}
		}
	}

	cmd := "Method"
	ans, err := xl.cores.SendNum(cmd, funcName, Q.num, opt)
	if err != nil {
		return fmt.Errorf("#ERROR %v", err)
	}

	return ans
}

func (Q *workFunction) AccrInt(args ...any) any    { return Q.function("AccrInt", args...) }
func (Q *workFunction) AccrIntM(args ...any) any   { return Q.function("AccrIntM", args...) }
func (Q *workFunction) Acos(args ...any) any       { return Q.function("Acos", args...) }
func (Q *workFunction) Acosh(args ...any) any      { return Q.function("Acosh", args...) }
func (Q *workFunction) Acot(args ...any) any       { return Q.function("Acot", args...) }
func (Q *workFunction) Acoth(args ...any) any      { return Q.function("Acoth", args...) }
func (Q *workFunction) Aggregate(args ...any) any  { return Q.function("Aggregate", args...) }
func (Q *workFunction) AmorDegrc(args ...any) any  { return Q.function("AmorDegrc", args...) }
func (Q *workFunction) AmorLinc(args ...any) any   { return Q.function("AmorLinc", args...) }
func (Q *workFunction) And(args ...any) any        { return Q.function("And", args...) }
func (Q *workFunction) Arabic(args ...any) any     { return Q.function("Arabic", args...) }
func (Q *workFunction) Asc(args ...any) any        { return Q.function("Asc", args...) }
func (Q *workFunction) Asin(args ...any) any       { return Q.function("Asin", args...) }
func (Q *workFunction) Asinh(args ...any) any      { return Q.function("Asinh", args...) }
func (Q *workFunction) Atan2(args ...any) any      { return Q.function("Atan2", args...) }
func (Q *workFunction) Atanh(args ...any) any      { return Q.function("Atanh", args...) }
func (Q *workFunction) AveDev(args ...any) any     { return Q.function("AveDev", args...) }
func (Q *workFunction) Average(args ...any) any    { return Q.function("Average", args...) }
func (Q *workFunction) AverageIf(args ...any) any  { return Q.function("AverageIf", args...) }
func (Q *workFunction) AverageIfs(args ...any) any { return Q.function("AverageIfs", args...) }
func (Q *workFunction) BahtText(args ...any) any   { return Q.function("BahtText", args...) }
func (Q *workFunction) Base(args ...any) any       { return Q.function("Base", args...) }
func (Q *workFunction) BesselI(args ...any) any    { return Q.function("BesselI", args...) }
func (Q *workFunction) BesselJ(args ...any) any    { return Q.function("BesselJ", args...) }
func (Q *workFunction) BesselK(args ...any) any    { return Q.function("BesselK", args...) }
func (Q *workFunction) BesselY(args ...any) any    { return Q.function("BesselY", args...) }
func (Q *workFunction) Beta_Dist(args ...any) any  { return Q.function("Beta_Dist", args...) }
func (Q *workFunction) Beta_Inv(args ...any) any   { return Q.function("Beta_Inv", args...) }
func (Q *workFunction) BetaDist(args ...any) any   { return Q.function("BetaDist", args...) }
func (Q *workFunction) BetaInv(args ...any) any    { return Q.function("BetaInv", args...) }
func (Q *workFunction) Bin2Dec(args ...any) any    { return Q.function("Bin2Dec", args...) }
func (Q *workFunction) Bin2Hex(args ...any) any    { return Q.function("Bin2Hex", args...) }
func (Q *workFunction) Bin2Oct(args ...any) any    { return Q.function("Bin2Oct", args...) }
func (Q *workFunction) Binom_Dist(args ...any) any { return Q.function("Binom_Dist", args...) }
func (Q *workFunction) Binom_Dist_Range(args ...any) any {
	return Q.function("Binom_Dist_Range", args...)
}
func (Q *workFunction) Binom_Inv(args ...any) any    { return Q.function("Binom_Inv", args...) }
func (Q *workFunction) BinomDist(args ...any) any    { return Q.function("BinomDist", args...) }
func (Q *workFunction) Bitand(args ...any) any       { return Q.function("Bitand", args...) }
func (Q *workFunction) Bitlshift(args ...any) any    { return Q.function("Bitlshift", args...) }
func (Q *workFunction) Bitor(args ...any) any        { return Q.function("Bitor", args...) }
func (Q *workFunction) Bitrshift(args ...any) any    { return Q.function("Bitrshift", args...) }
func (Q *workFunction) Bitxor(args ...any) any       { return Q.function("Bitxor", args...) }
func (Q *workFunction) Ceiling(args ...any) any      { return Q.function("Ceiling", args...) }
func (Q *workFunction) Ceiling_Math(args ...any) any { return Q.function("Ceiling_Math", args...) }
func (Q *workFunction) Ceiling_Precise(args ...any) any {
	return Q.function("Ceiling_Precise", args...)
}
func (Q *workFunction) ChiDist(args ...any) any       { return Q.function("ChiDist", args...) }
func (Q *workFunction) ChiInv(args ...any) any        { return Q.function("ChiInv", args...) }
func (Q *workFunction) ChiSq_Dist(args ...any) any    { return Q.function("ChiSq_Dist", args...) }
func (Q *workFunction) ChiSq_Dist_RT(args ...any) any { return Q.function("ChiSq_Dist_RT", args...) }
func (Q *workFunction) ChiSq_Inv(args ...any) any     { return Q.function("ChiSq_Inv", args...) }
func (Q *workFunction) ChiSq_Inv_RT(args ...any) any  { return Q.function("ChiSq_Inv_RT", args...) }
func (Q *workFunction) ChiSq_Test(args ...any) any    { return Q.function("ChiSq_Test", args...) }
func (Q *workFunction) ChiTest(args ...any) any       { return Q.function("ChiTest", args...) }
func (Q *workFunction) Choose(args ...any) any        { return Q.function("Choose", args...) }
func (Q *workFunction) Clean(args ...any) any         { return Q.function("Clean", args...) }
func (Q *workFunction) Combin(args ...any) any        { return Q.function("Combin", args...) }
func (Q *workFunction) Combina(args ...any) any       { return Q.function("Combina", args...) }
func (Q *workFunction) Complex(args ...any) any       { return Q.function("Complex", args...) }
func (Q *workFunction) Confidence(args ...any) any    { return Q.function("Confidence", args...) }
func (Q *workFunction) Confidence_Norm(args ...any) any {
	return Q.function("Confidence_Norm", args...)
}
func (Q *workFunction) Confidence_T(args ...any) any  { return Q.function("Confidence_T", args...) }
func (Q *workFunction) Convert(args ...any) any       { return Q.function("Convert", args...) }
func (Q *workFunction) Correl(args ...any) any        { return Q.function("Correl", args...) }
func (Q *workFunction) Cosh(args ...any) any          { return Q.function("Cosh", args...) }
func (Q *workFunction) Cot(args ...any) any           { return Q.function("Cot", args...) }
func (Q *workFunction) Coth(args ...any) any          { return Q.function("Coth", args...) }
func (Q *workFunction) Count(args ...any) any         { return Q.function("Count", args...) }
func (Q *workFunction) CountA(args ...any) any        { return Q.function("CountA", args...) }
func (Q *workFunction) CountBlank(args ...any) any    { return Q.function("CountBlank", args...) }
func (Q *workFunction) CountIf(args ...any) any       { return Q.function("CountIf", args...) }
func (Q *workFunction) CountIfs(args ...any) any      { return Q.function("CountIfs", args...) }
func (Q *workFunction) CoupDayBs(args ...any) any     { return Q.function("CoupDayBs", args...) }
func (Q *workFunction) CoupDays(args ...any) any      { return Q.function("CoupDays", args...) }
func (Q *workFunction) CoupDaysNc(args ...any) any    { return Q.function("CoupDaysNc", args...) }
func (Q *workFunction) CoupNcd(args ...any) any       { return Q.function("CoupNcd", args...) }
func (Q *workFunction) CoupNum(args ...any) any       { return Q.function("CoupNum", args...) }
func (Q *workFunction) CoupPcd(args ...any) any       { return Q.function("CoupPcd", args...) }
func (Q *workFunction) Covar(args ...any) any         { return Q.function("Covar", args...) }
func (Q *workFunction) Covariance_P(args ...any) any  { return Q.function("Covariance_P", args...) }
func (Q *workFunction) Covariance_S(args ...any) any  { return Q.function("Covariance_S", args...) }
func (Q *workFunction) CritBinom(args ...any) any     { return Q.function("CritBinom", args...) }
func (Q *workFunction) Csc(args ...any) any           { return Q.function("Csc", args...) }
func (Q *workFunction) Csch(args ...any) any          { return Q.function("Csch", args...) }
func (Q *workFunction) CumIPmt(args ...any) any       { return Q.function("CumIPmt", args...) }
func (Q *workFunction) CumPrinc(args ...any) any      { return Q.function("CumPrinc", args...) }
func (Q *workFunction) DAverage(args ...any) any      { return Q.function("DAverage", args...) }
func (Q *workFunction) Day(args ...any) any           { return Q.function("Day", args...) }
func (Q *workFunction) Days360(args ...any) any       { return Q.function("Days360", args...) }
func (Q *workFunction) Db(args ...any) any            { return Q.function("Db", args...) }
func (Q *workFunction) Dbcs(args ...any) any          { return Q.function("Dbcs", args...) }
func (Q *workFunction) DCount(args ...any) any        { return Q.function("DCount", args...) }
func (Q *workFunction) DCountA(args ...any) any       { return Q.function("DCountA", args...) }
func (Q *workFunction) Ddb(args ...any) any           { return Q.function("Ddb", args...) }
func (Q *workFunction) Dec2Bin(args ...any) any       { return Q.function("Dec2Bin", args...) }
func (Q *workFunction) Dec2Hex(args ...any) any       { return Q.function("Dec2Hex", args...) }
func (Q *workFunction) Dec2Oct(args ...any) any       { return Q.function("Dec2Oct", args...) }
func (Q *workFunction) Decimal(args ...any) any       { return Q.function("Decimal", args...) }
func (Q *workFunction) Degrees(args ...any) any       { return Q.function("Degrees", args...) }
func (Q *workFunction) Delta(args ...any) any         { return Q.function("Delta", args...) }
func (Q *workFunction) DevSq(args ...any) any         { return Q.function("DevSq", args...) }
func (Q *workFunction) DGet(args ...any) any          { return Q.function("DGet", args...) }
func (Q *workFunction) Disc(args ...any) any          { return Q.function("Disc", args...) }
func (Q *workFunction) DMax(args ...any) any          { return Q.function("DMax", args...) }
func (Q *workFunction) DMin(args ...any) any          { return Q.function("DMin", args...) }
func (Q *workFunction) Dollar(args ...any) any        { return Q.function("Dollar", args...) }
func (Q *workFunction) DollarDe(args ...any) any      { return Q.function("DollarDe", args...) }
func (Q *workFunction) DollarFr(args ...any) any      { return Q.function("DollarFr", args...) }
func (Q *workFunction) DProduct(args ...any) any      { return Q.function("DProduct", args...) }
func (Q *workFunction) DStDev(args ...any) any        { return Q.function("DStDev", args...) }
func (Q *workFunction) DStDevP(args ...any) any       { return Q.function("DStDevP", args...) }
func (Q *workFunction) DSum(args ...any) any          { return Q.function("DSum", args...) }
func (Q *workFunction) Duration(args ...any) any      { return Q.function("Duration", args...) }
func (Q *workFunction) DVar(args ...any) any          { return Q.function("DVar", args...) }
func (Q *workFunction) DVarP(args ...any) any         { return Q.function("DVarP", args...) }
func (Q *workFunction) EDate(args ...any) any         { return Q.function("EDate", args...) }
func (Q *workFunction) Effect(args ...any) any        { return Q.function("Effect", args...) }
func (Q *workFunction) EncodeUrl(args ...any) any     { return Q.function("EncodeUrl", args...) }
func (Q *workFunction) EoMonth(args ...any) any       { return Q.function("EoMonth", args...) }
func (Q *workFunction) Erf(args ...any) any           { return Q.function("Erf", args...) }
func (Q *workFunction) Erf_Precise(args ...any) any   { return Q.function("Erf_Precise", args...) }
func (Q *workFunction) ErfC(args ...any) any          { return Q.function("ErfC", args...) }
func (Q *workFunction) ErfC_Precise(args ...any) any  { return Q.function("ErfC_Precise", args...) }
func (Q *workFunction) Even(args ...any) any          { return Q.function("Even", args...) }
func (Q *workFunction) Expon_Dist(args ...any) any    { return Q.function("Expon_Dist", args...) }
func (Q *workFunction) ExponDist(args ...any) any     { return Q.function("ExponDist", args...) }
func (Q *workFunction) F_Dist(args ...any) any        { return Q.function("F_Dist", args...) }
func (Q *workFunction) F_Dist_RT(args ...any) any     { return Q.function("F_Dist_RT", args...) }
func (Q *workFunction) F_Inv(args ...any) any         { return Q.function("F_Inv", args...) }
func (Q *workFunction) F_Inv_RT(args ...any) any      { return Q.function("F_Inv_RT", args...) }
func (Q *workFunction) F_Test(args ...any) any        { return Q.function("F_Test", args...) }
func (Q *workFunction) Fact(args ...any) any          { return Q.function("Fact", args...) }
func (Q *workFunction) FactDouble(args ...any) any    { return Q.function("FactDouble", args...) }
func (Q *workFunction) FDist(args ...any) any         { return Q.function("FDist", args...) }
func (Q *workFunction) FilterXML(args ...any) any     { return Q.function("FilterXML", args...) }
func (Q *workFunction) Find(args ...any) any          { return Q.function("Find", args...) }
func (Q *workFunction) FindB(args ...any) any         { return Q.function("FindB", args...) }
func (Q *workFunction) FInv(args ...any) any          { return Q.function("FInv", args...) }
func (Q *workFunction) Fisher(args ...any) any        { return Q.function("Fisher", args...) }
func (Q *workFunction) FisherInv(args ...any) any     { return Q.function("FisherInv", args...) }
func (Q *workFunction) Fixed(args ...any) any         { return Q.function("Fixed", args...) }
func (Q *workFunction) Floor(args ...any) any         { return Q.function("Floor", args...) }
func (Q *workFunction) Floor_Math(args ...any) any    { return Q.function("Floor_Math", args...) }
func (Q *workFunction) Floor_Precise(args ...any) any { return Q.function("Floor_Precise", args...) }
func (Q *workFunction) Forecast(args ...any) any      { return Q.function("Forecast", args...) }
func (Q *workFunction) Forecast_ETS(args ...any) any  { return Q.function("Forecast_ETS", args...) }
func (Q *workFunction) Forecast_ETS_ConfInt(args ...any) any {
	return Q.function("Forecast_ETS_ConfInt", args...)
}
func (Q *workFunction) Forecast_ETS_Seasonality(args ...any) any {
	return Q.function("Forecast_ETS_Seasonality", args...)
}
func (Q *workFunction) Forecast_ETS_STAT(args ...any) any {
	return Q.function("Forecast_ETS_STAT", args...)
}
func (Q *workFunction) Forecast_Linear(args ...any) any {
	return Q.function("Forecast_Linear", args...)
}
func (Q *workFunction) Frequency(args ...any) any  { return Q.function("Frequency", args...) }
func (Q *workFunction) FTest(args ...any) any      { return Q.function("FTest", args...) }
func (Q *workFunction) Fv(args ...any) any         { return Q.function("Fv", args...) }
func (Q *workFunction) FVSchedule(args ...any) any { return Q.function("FVSchedule", args...) }
func (Q *workFunction) Gamma(args ...any) any      { return Q.function("Gamma", args...) }
func (Q *workFunction) Gamma_Dist(args ...any) any { return Q.function("Gamma_Dist", args...) }
func (Q *workFunction) Gamma_Inv(args ...any) any  { return Q.function("Gamma_Inv", args...) }
func (Q *workFunction) GammaDist(args ...any) any  { return Q.function("GammaDist", args...) }
func (Q *workFunction) GammaInv(args ...any) any   { return Q.function("GammaInv", args...) }
func (Q *workFunction) GammaLn(args ...any) any    { return Q.function("GammaLn", args...) }
func (Q *workFunction) GammaLn_Precise(args ...any) any {
	return Q.function("GammaLn_Precise", args...)
}
func (Q *workFunction) Gauss(args ...any) any        { return Q.function("Gauss", args...) }
func (Q *workFunction) Gcd(args ...any) any          { return Q.function("Gcd", args...) }
func (Q *workFunction) GeoMean(args ...any) any      { return Q.function("GeoMean", args...) }
func (Q *workFunction) GeStep(args ...any) any       { return Q.function("GeStep", args...) }
func (Q *workFunction) Growth(args ...any) any       { return Q.function("Growth", args...) }
func (Q *workFunction) HarMean(args ...any) any      { return Q.function("HarMean", args...) }
func (Q *workFunction) Hex2Bin(args ...any) any      { return Q.function("Hex2Bin", args...) }
func (Q *workFunction) Hex2Dec(args ...any) any      { return Q.function("Hex2Dec", args...) }
func (Q *workFunction) Hex2Oct(args ...any) any      { return Q.function("Hex2Oct", args...) }
func (Q *workFunction) HLookup(args ...any) any      { return Q.function("HLookup", args...) }
func (Q *workFunction) HypGeom_Dist(args ...any) any { return Q.function("HypGeom_Dist", args...) }
func (Q *workFunction) HypGeomDist(args ...any) any  { return Q.function("HypGeomDist", args...) }
func (Q *workFunction) IfError(args ...any) any      { return Q.function("IfError", args...) }
func (Q *workFunction) IfNa(args ...any) any         { return Q.function("IfNa", args...) }
func (Q *workFunction) ImAbs(args ...any) any        { return Q.function("ImAbs", args...) }
func (Q *workFunction) Imaginary(args ...any) any    { return Q.function("Imaginary", args...) }
func (Q *workFunction) ImArgument(args ...any) any   { return Q.function("ImArgument", args...) }
func (Q *workFunction) ImConjugate(args ...any) any  { return Q.function("ImConjugate", args...) }
func (Q *workFunction) ImCos(args ...any) any        { return Q.function("ImCos", args...) }
func (Q *workFunction) ImCosh(args ...any) any       { return Q.function("ImCosh", args...) }
func (Q *workFunction) ImCot(args ...any) any        { return Q.function("ImCot", args...) }
func (Q *workFunction) ImCsc(args ...any) any        { return Q.function("ImCsc", args...) }
func (Q *workFunction) ImCsch(args ...any) any       { return Q.function("ImCsch", args...) }
func (Q *workFunction) ImDiv(args ...any) any        { return Q.function("ImDiv", args...) }
func (Q *workFunction) ImExp(args ...any) any        { return Q.function("ImExp", args...) }
func (Q *workFunction) ImLn(args ...any) any         { return Q.function("ImLn", args...) }
func (Q *workFunction) ImLog10(args ...any) any      { return Q.function("ImLog10", args...) }
func (Q *workFunction) ImLog2(args ...any) any       { return Q.function("ImLog2", args...) }
func (Q *workFunction) ImPower(args ...any) any      { return Q.function("ImPower", args...) }
func (Q *workFunction) ImProduct(args ...any) any    { return Q.function("ImProduct", args...) }
func (Q *workFunction) ImReal(args ...any) any       { return Q.function("ImReal", args...) }
func (Q *workFunction) ImSec(args ...any) any        { return Q.function("ImSec", args...) }
func (Q *workFunction) ImSech(args ...any) any       { return Q.function("ImSech", args...) }
func (Q *workFunction) ImSin(args ...any) any        { return Q.function("ImSin", args...) }
func (Q *workFunction) ImSinh(args ...any) any       { return Q.function("ImSinh", args...) }
func (Q *workFunction) ImSqrt(args ...any) any       { return Q.function("ImSqrt", args...) }
func (Q *workFunction) ImSub(args ...any) any        { return Q.function("ImSub", args...) }
func (Q *workFunction) ImSum(args ...any) any        { return Q.function("ImSum", args...) }
func (Q *workFunction) ImTan(args ...any) any        { return Q.function("ImTan", args...) }
func (Q *workFunction) インデックス(args ...any) any {
	return Q.function("インデックス", args...)
}
func (Q *workFunction) Intercept(args ...any) any     { return Q.function("Intercept", args...) }
func (Q *workFunction) IntRate(args ...any) any       { return Q.function("IntRate", args...) }
func (Q *workFunction) Ipmt(args ...any) any          { return Q.function("Ipmt", args...) }
func (Q *workFunction) Irr(args ...any) any           { return Q.function("Irr", args...) }
func (Q *workFunction) IsErr(args ...any) any         { return Q.function("IsErr", args...) }
func (Q *workFunction) IsError(args ...any) any       { return Q.function("IsError", args...) }
func (Q *workFunction) IsEven(args ...any) any        { return Q.function("IsEven", args...) }
func (Q *workFunction) IsFormula(args ...any) any     { return Q.function("IsFormula", args...) }
func (Q *workFunction) IsLogical(args ...any) any     { return Q.function("IsLogical", args...) }
func (Q *workFunction) IsNA(args ...any) any          { return Q.function("IsNA", args...) }
func (Q *workFunction) IsNonText(args ...any) any     { return Q.function("IsNonText", args...) }
func (Q *workFunction) IsNumber(args ...any) any      { return Q.function("IsNumber", args...) }
func (Q *workFunction) ISO_Ceiling(args ...any) any   { return Q.function("ISO_Ceiling", args...) }
func (Q *workFunction) IsOdd(args ...any) any         { return Q.function("IsOdd", args...) }
func (Q *workFunction) IsoWeekNum(args ...any) any    { return Q.function("IsoWeekNum", args...) }
func (Q *workFunction) Ispmt(args ...any) any         { return Q.function("Ispmt", args...) }
func (Q *workFunction) IsText(args ...any) any        { return Q.function("IsText", args...) }
func (Q *workFunction) Kurt(args ...any) any          { return Q.function("Kurt", args...) }
func (Q *workFunction) Large(args ...any) any         { return Q.function("Large", args...) }
func (Q *workFunction) Lcm(args ...any) any           { return Q.function("Lcm", args...) }
func (Q *workFunction) LinEst(args ...any) any        { return Q.function("LinEst", args...) }
func (Q *workFunction) Ln(args ...any) any            { return Q.function("Ln", args...) }
func (Q *workFunction) Log(args ...any) any           { return Q.function("Log", args...) }
func (Q *workFunction) Log10(args ...any) any         { return Q.function("Log10", args...) }
func (Q *workFunction) LogEst(args ...any) any        { return Q.function("LogEst", args...) }
func (Q *workFunction) LogInv(args ...any) any        { return Q.function("LogInv", args...) }
func (Q *workFunction) LogNorm_Dist(args ...any) any  { return Q.function("LogNorm_Dist", args...) }
func (Q *workFunction) LogNorm_Inv(args ...any) any   { return Q.function("LogNorm_Inv", args...) }
func (Q *workFunction) LogNormDist(args ...any) any   { return Q.function("LogNormDist", args...) }
func (Q *workFunction) Lookup(args ...any) any        { return Q.function("Lookup", args...) }
func (Q *workFunction) Match(args ...any) any         { return Q.function("Match", args...) }
func (Q *workFunction) Max(args ...any) any           { return Q.function("Max", args...) }
func (Q *workFunction) MDeterm(args ...any) any       { return Q.function("MDeterm", args...) }
func (Q *workFunction) MDuration(args ...any) any     { return Q.function("MDuration", args...) }
func (Q *workFunction) Median(args ...any) any        { return Q.function("Median", args...) }
func (Q *workFunction) Min(args ...any) any           { return Q.function("Min", args...) }
func (Q *workFunction) MInverse(args ...any) any      { return Q.function("MInverse", args...) }
func (Q *workFunction) MIrr(args ...any) any          { return Q.function("MIrr", args...) }
func (Q *workFunction) MMult(args ...any) any         { return Q.function("MMult", args...) }
func (Q *workFunction) Mode(args ...any) any          { return Q.function("Mode", args...) }
func (Q *workFunction) Mode_Mult(args ...any) any     { return Q.function("Mode_Mult", args...) }
func (Q *workFunction) Mode_Sngl(args ...any) any     { return Q.function("Mode_Sngl", args...) }
func (Q *workFunction) MRound(args ...any) any        { return Q.function("MRound", args...) }
func (Q *workFunction) MultiNomial(args ...any) any   { return Q.function("MultiNomial", args...) }
func (Q *workFunction) Munit(args ...any) any         { return Q.function("Munit", args...) }
func (Q *workFunction) NegBinom_Dist(args ...any) any { return Q.function("NegBinom_Dist", args...) }
func (Q *workFunction) NegBinomDist(args ...any) any  { return Q.function("NegBinomDist", args...) }
func (Q *workFunction) NetworkDays(args ...any) any   { return Q.function("NetworkDays", args...) }
func (Q *workFunction) NetworkDays_Intl(args ...any) any {
	return Q.function("NetworkDays_Intl", args...)
}
func (Q *workFunction) Nominal(args ...any) any        { return Q.function("Nominal", args...) }
func (Q *workFunction) Norm_Dist(args ...any) any      { return Q.function("Norm_Dist", args...) }
func (Q *workFunction) Norm_Inv(args ...any) any       { return Q.function("Norm_Inv", args...) }
func (Q *workFunction) Norm_S_Dist(args ...any) any    { return Q.function("Norm_S_Dist", args...) }
func (Q *workFunction) Norm_S_Inv(args ...any) any     { return Q.function("Norm_S_Inv", args...) }
func (Q *workFunction) NormDist(args ...any) any       { return Q.function("NormDist", args...) }
func (Q *workFunction) NormInv(args ...any) any        { return Q.function("NormInv", args...) }
func (Q *workFunction) NormSDist(args ...any) any      { return Q.function("NormSDist", args...) }
func (Q *workFunction) NormSInv(args ...any) any       { return Q.function("NormSInv", args...) }
func (Q *workFunction) NPer(args ...any) any           { return Q.function("NPer", args...) }
func (Q *workFunction) Npv(args ...any) any            { return Q.function("Npv", args...) }
func (Q *workFunction) NumberValue(args ...any) any    { return Q.function("NumberValue", args...) }
func (Q *workFunction) Oct2Bin(args ...any) any        { return Q.function("Oct2Bin", args...) }
func (Q *workFunction) Oct2Dec(args ...any) any        { return Q.function("Oct2Dec", args...) }
func (Q *workFunction) Oct2Hex(args ...any) any        { return Q.function("Oct2Hex", args...) }
func (Q *workFunction) Odd(args ...any) any            { return Q.function("Odd", args...) }
func (Q *workFunction) OddFPrice(args ...any) any      { return Q.function("OddFPrice", args...) }
func (Q *workFunction) OddFYield(args ...any) any      { return Q.function("OddFYield", args...) }
func (Q *workFunction) OddLPrice(args ...any) any      { return Q.function("OddLPrice", args...) }
func (Q *workFunction) OddLYield(args ...any) any      { return Q.function("OddLYield", args...) }
func (Q *workFunction) Or(args ...any) any             { return Q.function("Or", args...) }
func (Q *workFunction) PDuration(args ...any) any      { return Q.function("PDuration", args...) }
func (Q *workFunction) Pearson(args ...any) any        { return Q.function("Pearson", args...) }
func (Q *workFunction) Percentile(args ...any) any     { return Q.function("Percentile", args...) }
func (Q *workFunction) Percentile_Exc(args ...any) any { return Q.function("Percentile_Exc", args...) }
func (Q *workFunction) Percentile_Inc(args ...any) any { return Q.function("Percentile_Inc", args...) }
func (Q *workFunction) PercentRank(args ...any) any    { return Q.function("PercentRank", args...) }
func (Q *workFunction) PercentRank_Exc(args ...any) any {
	return Q.function("PercentRank_Exc", args...)
}
func (Q *workFunction) PercentRank_Inc(args ...any) any {
	return Q.function("PercentRank_Inc", args...)
}
func (Q *workFunction) Permut(args ...any) any       { return Q.function("Permut", args...) }
func (Q *workFunction) Permutationa(args ...any) any { return Q.function("Permutationa", args...) }
func (Q *workFunction) Phi(args ...any) any          { return Q.function("Phi", args...) }
func (Q *workFunction) Phonetic(args ...any) any     { return Q.function("Phonetic", args...) }
func (Q *workFunction) Pi(args ...any) any           { return Q.function("Pi", args...) }
func (Q *workFunction) Pmt(args ...any) any          { return Q.function("Pmt", args...) }
func (Q *workFunction) Poisson(args ...any) any      { return Q.function("Poisson", args...) }
func (Q *workFunction) Poisson_Dist(args ...any) any { return Q.function("Poisson_Dist", args...) }
func (Q *workFunction) Power(args ...any) any        { return Q.function("Power", args...) }
func (Q *workFunction) Ppmt(args ...any) any         { return Q.function("Ppmt", args...) }
func (Q *workFunction) Price(args ...any) any        { return Q.function("Price", args...) }
func (Q *workFunction) PriceDisc(args ...any) any    { return Q.function("PriceDisc", args...) }
func (Q *workFunction) PriceMat(args ...any) any     { return Q.function("PriceMat", args...) }
func (Q *workFunction) Prob(args ...any) any         { return Q.function("Prob", args...) }
func (Q *workFunction) Product(args ...any) any      { return Q.function("Product", args...) }
func (Q *workFunction) Proper(args ...any) any       { return Q.function("Proper", args...) }
func (Q *workFunction) Pv(args ...any) any           { return Q.function("Pv", args...) }
func (Q *workFunction) Quartile(args ...any) any     { return Q.function("Quartile", args...) }
func (Q *workFunction) Quartile_Exc(args ...any) any { return Q.function("Quartile_Exc", args...) }
func (Q *workFunction) Quartile_Inc(args ...any) any { return Q.function("Quartile_Inc", args...) }
func (Q *workFunction) Quotient(args ...any) any     { return Q.function("Quotient", args...) }
func (Q *workFunction) Radians(args ...any) any      { return Q.function("Radians", args...) }
func (Q *workFunction) RandBetween(args ...any) any  { return Q.function("RandBetween", args...) }
func (Q *workFunction) Rank(args ...any) any         { return Q.function("Rank", args...) }
func (Q *workFunction) Rank_Avg(args ...any) any     { return Q.function("Rank_Avg", args...) }
func (Q *workFunction) Rank_Eq(args ...any) any      { return Q.function("Rank_Eq", args...) }
func (Q *workFunction) Rate(args ...any) any         { return Q.function("Rate", args...) }
func (Q *workFunction) Received(args ...any) any     { return Q.function("Received", args...) }
func (Q *workFunction) Replace(args ...any) any      { return Q.function("Replace", args...) }
func (Q *workFunction) ReplaceB(args ...any) any     { return Q.function("ReplaceB", args...) }
func (Q *workFunction) Rept(args ...any) any         { return Q.function("Rept", args...) }
func (Q *workFunction) Roman(args ...any) any        { return Q.function("Roman", args...) }
func (Q *workFunction) Round(args ...any) any        { return Q.function("Round", args...) }
func (Q *workFunction) RoundDown(args ...any) any    { return Q.function("RoundDown", args...) }
func (Q *workFunction) RoundUp(args ...any) any      { return Q.function("RoundUp", args...) }
func (Q *workFunction) Rri(args ...any) any          { return Q.function("Rri", args...) }
func (Q *workFunction) RSq(args ...any) any          { return Q.function("RSq", args...) }
func (Q *workFunction) RTD(args ...any) any          { return Q.function("RTD", args...) }
func (Q *workFunction) Search(args ...any) any       { return Q.function("Search", args...) }
func (Q *workFunction) SearchB(args ...any) any      { return Q.function("SearchB", args...) }
func (Q *workFunction) Sec(args ...any) any          { return Q.function("Sec", args...) }
func (Q *workFunction) Sech(args ...any) any         { return Q.function("Sech", args...) }
func (Q *workFunction) SeriesSum(args ...any) any    { return Q.function("SeriesSum", args...) }
func (Q *workFunction) Sinh(args ...any) any         { return Q.function("Sinh", args...) }
func (Q *workFunction) Skew(args ...any) any         { return Q.function("Skew", args...) }
func (Q *workFunction) Skew_p(args ...any) any       { return Q.function("Skew_p", args...) }
func (Q *workFunction) Sln(args ...any) any          { return Q.function("Sln", args...) }
func (Q *workFunction) Slope(args ...any) any        { return Q.function("Slope", args...) }
func (Q *workFunction) Small(args ...any) any        { return Q.function("Small", args...) }
func (Q *workFunction) SqrtPi(args ...any) any       { return Q.function("SqrtPi", args...) }
func (Q *workFunction) Standardize(args ...any) any  { return Q.function("Standardize", args...) }
func (Q *workFunction) StDev(args ...any) any        { return Q.function("StDev", args...) }
func (Q *workFunction) StDev_P(args ...any) any      { return Q.function("StDev_P", args...) }
func (Q *workFunction) StDev_S(args ...any) any      { return Q.function("StDev_S", args...) }
func (Q *workFunction) StDevP(args ...any) any       { return Q.function("StDevP", args...) }
func (Q *workFunction) StEyx(args ...any) any        { return Q.function("StEyx", args...) }
func (Q *workFunction) Substitute(args ...any) any   { return Q.function("Substitute", args...) }
func (Q *workFunction) Subtotal(args ...any) any     { return Q.function("Subtotal", args...) }
func (Q *workFunction) Sum(args ...any) any          { return Q.function("Sum", args...) }
func (Q *workFunction) SumIf(args ...any) any        { return Q.function("SumIf", args...) }
func (Q *workFunction) SumIfs(args ...any) any       { return Q.function("SumIfs", args...) }
func (Q *workFunction) SumProduct(args ...any) any   { return Q.function("SumProduct", args...) }
func (Q *workFunction) SumSq(args ...any) any        { return Q.function("SumSq", args...) }
func (Q *workFunction) SumX2MY2(args ...any) any     { return Q.function("SumX2MY2", args...) }
func (Q *workFunction) SumX2PY2(args ...any) any     { return Q.function("SumX2PY2", args...) }
func (Q *workFunction) SumXMY2(args ...any) any      { return Q.function("SumXMY2", args...) }
func (Q *workFunction) Syd(args ...any) any          { return Q.function("Syd", args...) }
func (Q *workFunction) T_Dist(args ...any) any       { return Q.function("T_Dist", args...) }
func (Q *workFunction) T_Dist_2T(args ...any) any    { return Q.function("T_Dist_2T", args...) }
func (Q *workFunction) T_Dist_RT(args ...any) any    { return Q.function("T_Dist_RT", args...) }
func (Q *workFunction) T_Inv(args ...any) any        { return Q.function("T_Inv", args...) }
func (Q *workFunction) T_Inv_2T(args ...any) any     { return Q.function("T_Inv_2T", args...) }
func (Q *workFunction) T_Test(args ...any) any       { return Q.function("T_Test", args...) }
func (Q *workFunction) Tanh(args ...any) any         { return Q.function("Tanh", args...) }
func (Q *workFunction) TBillEq(args ...any) any      { return Q.function("TBillEq", args...) }
func (Q *workFunction) TBillPrice(args ...any) any   { return Q.function("TBillPrice", args...) }
func (Q *workFunction) TBillYield(args ...any) any   { return Q.function("TBillYield", args...) }
func (Q *workFunction) TDist(args ...any) any        { return Q.function("TDist", args...) }
func (Q *workFunction) Text(args ...any) any         { return Q.function("Text", args...) }
func (Q *workFunction) TInv(args ...any) any         { return Q.function("TInv", args...) }
func (Q *workFunction) Transpose(args ...any) any    { return Q.function("Transpose", args...) }
func (Q *workFunction) Trend(args ...any) any        { return Q.function("Trend", args...) }
func (Q *workFunction) Trim(args ...any) any         { return Q.function("Trim", args...) }
func (Q *workFunction) TrimMean(args ...any) any     { return Q.function("TrimMean", args...) }
func (Q *workFunction) TTest(args ...any) any        { return Q.function("TTest", args...) }
func (Q *workFunction) Unichar(args ...any) any      { return Q.function("Unichar", args...) }
func (Q *workFunction) Unicode(args ...any) any      { return Q.function("Unicode", args...) }
func (Q *workFunction) USDollar(args ...any) any     { return Q.function("USDollar", args...) }
func (Q *workFunction) Var(args ...any) any          { return Q.function("Var", args...) }
func (Q *workFunction) Var_P(args ...any) any        { return Q.function("Var_P", args...) }
func (Q *workFunction) Var_S(args ...any) any        { return Q.function("Var_S", args...) }
func (Q *workFunction) VarP(args ...any) any         { return Q.function("VarP", args...) }
func (Q *workFunction) Vdb(args ...any) any          { return Q.function("Vdb", args...) }
func (Q *workFunction) VLookup(args ...any) any      { return Q.function("VLookup", args...) }
func (Q *workFunction) WebService(args ...any) any   { return Q.function("WebService", args...) }
func (Q *workFunction) Weekday(args ...any) any      { return Q.function("Weekday", args...) }
func (Q *workFunction) WeekNum(args ...any) any      { return Q.function("WeekNum", args...) }
func (Q *workFunction) Weibull(args ...any) any      { return Q.function("Weibull", args...) }
func (Q *workFunction) Weibull_Dist(args ...any) any { return Q.function("Weibull_Dist", args...) }
func (Q *workFunction) WorkDay(args ...any) any      { return Q.function("WorkDay", args...) }
func (Q *workFunction) WorkDay_Intl(args ...any) any { return Q.function("WorkDay_Intl", args...) }
func (Q *workFunction) Xirr(args ...any) any         { return Q.function("Xirr", args...) }
func (Q *workFunction) Xnpv(args ...any) any         { return Q.function("Xnpv", args...) }
func (Q *workFunction) Xor(args ...any) any          { return Q.function("Xor", args...) }
func (Q *workFunction) YearFrac(args ...any) any     { return Q.function("YearFrac", args...) }
func (Q *workFunction) YieldDisc(args ...any) any    { return Q.function("YieldDisc", args...) }
func (Q *workFunction) YieldMat(args ...any) any     { return Q.function("YieldMat", args...) }
func (Q *workFunction) Z_Test(args ...any) any       { return Q.function("Z_Test", args...) }
func (Q *workFunction) ZTest(args ...any) any        { return Q.function("ZTest", args...) }
