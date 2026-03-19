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

func (wa *workApp) WorksheetFunction() *workFunction {
	var wf workFunction
	xl := wa.app

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
	wf.app = xl
	wf.num = num
	wf.parent = wa
	return &wf
}

func (wf *workFunction) function(funcName string, args ...any) any {
	xl := wf.app

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
	ans, err := xl.cores.SendNum(cmd, funcName, wf.num, opt)
	if err != nil {
		return fmt.Errorf("#ERROR %v", err)
	}

	return ans
}

func (wf *workFunction) AccrInt(args ...any) any    { return wf.function("AccrInt", args...) }
func (wf *workFunction) AccrIntM(args ...any) any   { return wf.function("AccrIntM", args...) }
func (wf *workFunction) Acos(args ...any) any       { return wf.function("Acos", args...) }
func (wf *workFunction) Acosh(args ...any) any      { return wf.function("Acosh", args...) }
func (wf *workFunction) Acot(args ...any) any       { return wf.function("Acot", args...) }
func (wf *workFunction) Acoth(args ...any) any      { return wf.function("Acoth", args...) }
func (wf *workFunction) Aggregate(args ...any) any  { return wf.function("Aggregate", args...) }
func (wf *workFunction) AmorDegrc(args ...any) any  { return wf.function("AmorDegrc", args...) }
func (wf *workFunction) AmorLinc(args ...any) any   { return wf.function("AmorLinc", args...) }
func (wf *workFunction) And(args ...any) any        { return wf.function("And", args...) }
func (wf *workFunction) Arabic(args ...any) any     { return wf.function("Arabic", args...) }
func (wf *workFunction) Asc(args ...any) any        { return wf.function("Asc", args...) }
func (wf *workFunction) Asin(args ...any) any       { return wf.function("Asin", args...) }
func (wf *workFunction) Asinh(args ...any) any      { return wf.function("Asinh", args...) }
func (wf *workFunction) Atan2(args ...any) any      { return wf.function("Atan2", args...) }
func (wf *workFunction) Atanh(args ...any) any      { return wf.function("Atanh", args...) }
func (wf *workFunction) AveDev(args ...any) any     { return wf.function("AveDev", args...) }
func (wf *workFunction) Average(args ...any) any    { return wf.function("Average", args...) }
func (wf *workFunction) AverageIf(args ...any) any  { return wf.function("AverageIf", args...) }
func (wf *workFunction) AverageIfs(args ...any) any { return wf.function("AverageIfs", args...) }
func (wf *workFunction) BahtText(args ...any) any   { return wf.function("BahtText", args...) }
func (wf *workFunction) Base(args ...any) any       { return wf.function("Base", args...) }
func (wf *workFunction) BesselI(args ...any) any    { return wf.function("BesselI", args...) }
func (wf *workFunction) BesselJ(args ...any) any    { return wf.function("BesselJ", args...) }
func (wf *workFunction) BesselK(args ...any) any    { return wf.function("BesselK", args...) }
func (wf *workFunction) BesselY(args ...any) any    { return wf.function("BesselY", args...) }
func (wf *workFunction) Beta_Dist(args ...any) any  { return wf.function("Beta_Dist", args...) }
func (wf *workFunction) Beta_Inv(args ...any) any   { return wf.function("Beta_Inv", args...) }
func (wf *workFunction) BetaDist(args ...any) any   { return wf.function("BetaDist", args...) }
func (wf *workFunction) BetaInv(args ...any) any    { return wf.function("BetaInv", args...) }
func (wf *workFunction) Bin2Dec(args ...any) any    { return wf.function("Bin2Dec", args...) }
func (wf *workFunction) Bin2Hex(args ...any) any    { return wf.function("Bin2Hex", args...) }
func (wf *workFunction) Bin2Oct(args ...any) any    { return wf.function("Bin2Oct", args...) }
func (wf *workFunction) Binom_Dist(args ...any) any { return wf.function("Binom_Dist", args...) }
func (wf *workFunction) Binom_Dist_Range(args ...any) any {
	return wf.function("Binom_Dist_Range", args...)
}
func (wf *workFunction) Binom_Inv(args ...any) any    { return wf.function("Binom_Inv", args...) }
func (wf *workFunction) BinomDist(args ...any) any    { return wf.function("BinomDist", args...) }
func (wf *workFunction) Bitand(args ...any) any       { return wf.function("Bitand", args...) }
func (wf *workFunction) Bitlshift(args ...any) any    { return wf.function("Bitlshift", args...) }
func (wf *workFunction) Bitor(args ...any) any        { return wf.function("Bitor", args...) }
func (wf *workFunction) Bitrshift(args ...any) any    { return wf.function("Bitrshift", args...) }
func (wf *workFunction) Bitxor(args ...any) any       { return wf.function("Bitxor", args...) }
func (wf *workFunction) Ceiling(args ...any) any      { return wf.function("Ceiling", args...) }
func (wf *workFunction) Ceiling_Math(args ...any) any { return wf.function("Ceiling_Math", args...) }
func (wf *workFunction) Ceiling_Precise(args ...any) any {
	return wf.function("Ceiling_Precise", args...)
}
func (wf *workFunction) ChiDist(args ...any) any       { return wf.function("ChiDist", args...) }
func (wf *workFunction) ChiInv(args ...any) any        { return wf.function("ChiInv", args...) }
func (wf *workFunction) ChiSq_Dist(args ...any) any    { return wf.function("ChiSq_Dist", args...) }
func (wf *workFunction) ChiSq_Dist_RT(args ...any) any { return wf.function("ChiSq_Dist_RT", args...) }
func (wf *workFunction) ChiSq_Inv(args ...any) any     { return wf.function("ChiSq_Inv", args...) }
func (wf *workFunction) ChiSq_Inv_RT(args ...any) any  { return wf.function("ChiSq_Inv_RT", args...) }
func (wf *workFunction) ChiSq_Test(args ...any) any    { return wf.function("ChiSq_Test", args...) }
func (wf *workFunction) ChiTest(args ...any) any       { return wf.function("ChiTest", args...) }
func (wf *workFunction) Choose(args ...any) any        { return wf.function("Choose", args...) }
func (wf *workFunction) Clean(args ...any) any         { return wf.function("Clean", args...) }
func (wf *workFunction) Combin(args ...any) any        { return wf.function("Combin", args...) }
func (wf *workFunction) Combina(args ...any) any       { return wf.function("Combina", args...) }
func (wf *workFunction) Complex(args ...any) any       { return wf.function("Complex", args...) }
func (wf *workFunction) Confidence(args ...any) any    { return wf.function("Confidence", args...) }
func (wf *workFunction) Confidence_Norm(args ...any) any {
	return wf.function("Confidence_Norm", args...)
}
func (wf *workFunction) Confidence_T(args ...any) any  { return wf.function("Confidence_T", args...) }
func (wf *workFunction) Convert(args ...any) any       { return wf.function("Convert", args...) }
func (wf *workFunction) Correl(args ...any) any        { return wf.function("Correl", args...) }
func (wf *workFunction) Cosh(args ...any) any          { return wf.function("Cosh", args...) }
func (wf *workFunction) Cot(args ...any) any           { return wf.function("Cot", args...) }
func (wf *workFunction) Coth(args ...any) any          { return wf.function("Coth", args...) }
func (wf *workFunction) Count(args ...any) any         { return wf.function("Count", args...) }
func (wf *workFunction) CountA(args ...any) any        { return wf.function("CountA", args...) }
func (wf *workFunction) CountBlank(args ...any) any    { return wf.function("CountBlank", args...) }
func (wf *workFunction) CountIf(args ...any) any       { return wf.function("CountIf", args...) }
func (wf *workFunction) CountIfs(args ...any) any      { return wf.function("CountIfs", args...) }
func (wf *workFunction) CoupDayBs(args ...any) any     { return wf.function("CoupDayBs", args...) }
func (wf *workFunction) CoupDays(args ...any) any      { return wf.function("CoupDays", args...) }
func (wf *workFunction) CoupDaysNc(args ...any) any    { return wf.function("CoupDaysNc", args...) }
func (wf *workFunction) CoupNcd(args ...any) any       { return wf.function("CoupNcd", args...) }
func (wf *workFunction) CoupNum(args ...any) any       { return wf.function("CoupNum", args...) }
func (wf *workFunction) CoupPcd(args ...any) any       { return wf.function("CoupPcd", args...) }
func (wf *workFunction) Covar(args ...any) any         { return wf.function("Covar", args...) }
func (wf *workFunction) Covariance_P(args ...any) any  { return wf.function("Covariance_P", args...) }
func (wf *workFunction) Covariance_S(args ...any) any  { return wf.function("Covariance_S", args...) }
func (wf *workFunction) CritBinom(args ...any) any     { return wf.function("CritBinom", args...) }
func (wf *workFunction) Csc(args ...any) any           { return wf.function("Csc", args...) }
func (wf *workFunction) Csch(args ...any) any          { return wf.function("Csch", args...) }
func (wf *workFunction) CumIPmt(args ...any) any       { return wf.function("CumIPmt", args...) }
func (wf *workFunction) CumPrinc(args ...any) any      { return wf.function("CumPrinc", args...) }
func (wf *workFunction) DAverage(args ...any) any      { return wf.function("DAverage", args...) }
func (wf *workFunction) 日(args ...any) any             { return wf.function("日", args...) }
func (wf *workFunction) Days360(args ...any) any       { return wf.function("Days360", args...) }
func (wf *workFunction) Db(args ...any) any            { return wf.function("Db", args...) }
func (wf *workFunction) Dbcs(args ...any) any          { return wf.function("Dbcs", args...) }
func (wf *workFunction) DCount(args ...any) any        { return wf.function("DCount", args...) }
func (wf *workFunction) DCountA(args ...any) any       { return wf.function("DCountA", args...) }
func (wf *workFunction) Ddb(args ...any) any           { return wf.function("Ddb", args...) }
func (wf *workFunction) Dec2Bin(args ...any) any       { return wf.function("Dec2Bin", args...) }
func (wf *workFunction) Dec2Hex(args ...any) any       { return wf.function("Dec2Hex", args...) }
func (wf *workFunction) Dec2Oct(args ...any) any       { return wf.function("Dec2Oct", args...) }
func (wf *workFunction) Decimal(args ...any) any       { return wf.function("Decimal", args...) }
func (wf *workFunction) Degrees(args ...any) any       { return wf.function("Degrees", args...) }
func (wf *workFunction) 差分(args ...any) any            { return wf.function("差分", args...) }
func (wf *workFunction) DevSq(args ...any) any         { return wf.function("DevSq", args...) }
func (wf *workFunction) DGet(args ...any) any          { return wf.function("DGet", args...) }
func (wf *workFunction) Disc(args ...any) any          { return wf.function("Disc", args...) }
func (wf *workFunction) DMax(args ...any) any          { return wf.function("DMax", args...) }
func (wf *workFunction) DMin(args ...any) any          { return wf.function("DMin", args...) }
func (wf *workFunction) Dollar(args ...any) any        { return wf.function("Dollar", args...) }
func (wf *workFunction) DollarDe(args ...any) any      { return wf.function("DollarDe", args...) }
func (wf *workFunction) DollarFr(args ...any) any      { return wf.function("DollarFr", args...) }
func (wf *workFunction) DProduct(args ...any) any      { return wf.function("DProduct", args...) }
func (wf *workFunction) DStDev(args ...any) any        { return wf.function("DStDev", args...) }
func (wf *workFunction) DStDevP(args ...any) any       { return wf.function("DStDevP", args...) }
func (wf *workFunction) DSum(args ...any) any          { return wf.function("DSum", args...) }
func (wf *workFunction) 期間(args ...any) any            { return wf.function("期間", args...) }
func (wf *workFunction) DVar(args ...any) any          { return wf.function("DVar", args...) }
func (wf *workFunction) DVarP(args ...any) any         { return wf.function("DVarP", args...) }
func (wf *workFunction) EDate(args ...any) any         { return wf.function("EDate", args...) }
func (wf *workFunction) Effect(args ...any) any        { return wf.function("Effect", args...) }
func (wf *workFunction) EncodeUrl(args ...any) any     { return wf.function("EncodeUrl", args...) }
func (wf *workFunction) EoMonth(args ...any) any       { return wf.function("EoMonth", args...) }
func (wf *workFunction) Erf(args ...any) any           { return wf.function("Erf", args...) }
func (wf *workFunction) Erf_Precise(args ...any) any   { return wf.function("Erf_Precise", args...) }
func (wf *workFunction) ErfC(args ...any) any          { return wf.function("ErfC", args...) }
func (wf *workFunction) ErfC_Precise(args ...any) any  { return wf.function("ErfC_Precise", args...) }
func (wf *workFunction) Even(args ...any) any          { return wf.function("Even", args...) }
func (wf *workFunction) Expon_Dist(args ...any) any    { return wf.function("Expon_Dist", args...) }
func (wf *workFunction) ExponDist(args ...any) any     { return wf.function("ExponDist", args...) }
func (wf *workFunction) F_Dist(args ...any) any        { return wf.function("F_Dist", args...) }
func (wf *workFunction) F_Dist_RT(args ...any) any     { return wf.function("F_Dist_RT", args...) }
func (wf *workFunction) F_Inv(args ...any) any         { return wf.function("F_Inv", args...) }
func (wf *workFunction) F_Inv_RT(args ...any) any      { return wf.function("F_Inv_RT", args...) }
func (wf *workFunction) F_Test(args ...any) any        { return wf.function("F_Test", args...) }
func (wf *workFunction) Fact(args ...any) any          { return wf.function("Fact", args...) }
func (wf *workFunction) FactDouble(args ...any) any    { return wf.function("FactDouble", args...) }
func (wf *workFunction) FDist(args ...any) any         { return wf.function("FDist", args...) }
func (wf *workFunction) FilterXML(args ...any) any     { return wf.function("FilterXML", args...) }
func (wf *workFunction) Find(args ...any) any          { return wf.function("Find", args...) }
func (wf *workFunction) FindB(args ...any) any         { return wf.function("FindB", args...) }
func (wf *workFunction) FInv(args ...any) any          { return wf.function("FInv", args...) }
func (wf *workFunction) Fisher(args ...any) any        { return wf.function("Fisher", args...) }
func (wf *workFunction) FisherInv(args ...any) any     { return wf.function("FisherInv", args...) }
func (wf *workFunction) Fixed(args ...any) any         { return wf.function("Fixed", args...) }
func (wf *workFunction) Floor(args ...any) any         { return wf.function("Floor", args...) }
func (wf *workFunction) Floor_Math(args ...any) any    { return wf.function("Floor_Math", args...) }
func (wf *workFunction) Floor_Precise(args ...any) any { return wf.function("Floor_Precise", args...) }
func (wf *workFunction) Forecast(args ...any) any      { return wf.function("Forecast", args...) }
func (wf *workFunction) Forecast_ETS(args ...any) any  { return wf.function("Forecast_ETS", args...) }
func (wf *workFunction) Forecast_ETS_ConfInt(args ...any) any {
	return wf.function("Forecast_ETS_ConfInt", args...)
}
func (wf *workFunction) Forecast_ETS_Seasonality(args ...any) any {
	return wf.function("Forecast_ETS_Seasonality", args...)
}
func (wf *workFunction) Forecast_ETS_STAT(args ...any) any {
	return wf.function("Forecast_ETS_STAT", args...)
}
func (wf *workFunction) Forecast_Linear(args ...any) any {
	return wf.function("Forecast_Linear", args...)
}
func (wf *workFunction) Frequency(args ...any) any  { return wf.function("Frequency", args...) }
func (wf *workFunction) FTest(args ...any) any      { return wf.function("FTest", args...) }
func (wf *workFunction) Fv(args ...any) any         { return wf.function("Fv", args...) }
func (wf *workFunction) FVSchedule(args ...any) any { return wf.function("FVSchedule", args...) }
func (wf *workFunction) Gamma(args ...any) any      { return wf.function("Gamma", args...) }
func (wf *workFunction) Gamma_Dist(args ...any) any { return wf.function("Gamma_Dist", args...) }
func (wf *workFunction) Gamma_Inv(args ...any) any  { return wf.function("Gamma_Inv", args...) }
func (wf *workFunction) GammaDist(args ...any) any  { return wf.function("GammaDist", args...) }
func (wf *workFunction) GammaInv(args ...any) any   { return wf.function("GammaInv", args...) }
func (wf *workFunction) GammaLn(args ...any) any    { return wf.function("GammaLn", args...) }
func (wf *workFunction) GammaLn_Precise(args ...any) any {
	return wf.function("GammaLn_Precise", args...)
}
func (wf *workFunction) Gauss(args ...any) any        { return wf.function("Gauss", args...) }
func (wf *workFunction) Gcd(args ...any) any          { return wf.function("Gcd", args...) }
func (wf *workFunction) GeoMean(args ...any) any      { return wf.function("GeoMean", args...) }
func (wf *workFunction) GeStep(args ...any) any       { return wf.function("GeStep", args...) }
func (wf *workFunction) Growth(args ...any) any       { return wf.function("Growth", args...) }
func (wf *workFunction) HarMean(args ...any) any      { return wf.function("HarMean", args...) }
func (wf *workFunction) Hex2Bin(args ...any) any      { return wf.function("Hex2Bin", args...) }
func (wf *workFunction) Hex2Dec(args ...any) any      { return wf.function("Hex2Dec", args...) }
func (wf *workFunction) Hex2Oct(args ...any) any      { return wf.function("Hex2Oct", args...) }
func (wf *workFunction) HLookup(args ...any) any      { return wf.function("HLookup", args...) }
func (wf *workFunction) HypGeom_Dist(args ...any) any { return wf.function("HypGeom_Dist", args...) }
func (wf *workFunction) HypGeomDist(args ...any) any  { return wf.function("HypGeomDist", args...) }
func (wf *workFunction) IfError(args ...any) any      { return wf.function("IfError", args...) }
func (wf *workFunction) IfNa(args ...any) any         { return wf.function("IfNa", args...) }
func (wf *workFunction) ImAbs(args ...any) any        { return wf.function("ImAbs", args...) }
func (wf *workFunction) Imaginary(args ...any) any    { return wf.function("Imaginary", args...) }
func (wf *workFunction) ImArgument(args ...any) any   { return wf.function("ImArgument", args...) }
func (wf *workFunction) ImConjugate(args ...any) any  { return wf.function("ImConjugate", args...) }
func (wf *workFunction) ImCos(args ...any) any        { return wf.function("ImCos", args...) }
func (wf *workFunction) ImCosh(args ...any) any       { return wf.function("ImCosh", args...) }
func (wf *workFunction) ImCot(args ...any) any        { return wf.function("ImCot", args...) }
func (wf *workFunction) ImCsc(args ...any) any        { return wf.function("ImCsc", args...) }
func (wf *workFunction) ImCsch(args ...any) any       { return wf.function("ImCsch", args...) }
func (wf *workFunction) ImDiv(args ...any) any        { return wf.function("ImDiv", args...) }
func (wf *workFunction) ImExp(args ...any) any        { return wf.function("ImExp", args...) }
func (wf *workFunction) ImLn(args ...any) any         { return wf.function("ImLn", args...) }
func (wf *workFunction) ImLog10(args ...any) any      { return wf.function("ImLog10", args...) }
func (wf *workFunction) ImLog2(args ...any) any       { return wf.function("ImLog2", args...) }
func (wf *workFunction) ImPower(args ...any) any      { return wf.function("ImPower", args...) }
func (wf *workFunction) ImProduct(args ...any) any    { return wf.function("ImProduct", args...) }
func (wf *workFunction) ImReal(args ...any) any       { return wf.function("ImReal", args...) }
func (wf *workFunction) ImSec(args ...any) any        { return wf.function("ImSec", args...) }
func (wf *workFunction) ImSech(args ...any) any       { return wf.function("ImSech", args...) }
func (wf *workFunction) ImSin(args ...any) any        { return wf.function("ImSin", args...) }
func (wf *workFunction) ImSinh(args ...any) any       { return wf.function("ImSinh", args...) }
func (wf *workFunction) ImSqrt(args ...any) any       { return wf.function("ImSqrt", args...) }
func (wf *workFunction) ImSub(args ...any) any        { return wf.function("ImSub", args...) }
func (wf *workFunction) ImSum(args ...any) any        { return wf.function("ImSum", args...) }
func (wf *workFunction) ImTan(args ...any) any        { return wf.function("ImTan", args...) }
func (wf *workFunction) インデックス(args ...any) any {
	return wf.function("インデックス", args...)
}
func (wf *workFunction) Intercept(args ...any) any     { return wf.function("Intercept", args...) }
func (wf *workFunction) IntRate(args ...any) any       { return wf.function("IntRate", args...) }
func (wf *workFunction) Ipmt(args ...any) any          { return wf.function("Ipmt", args...) }
func (wf *workFunction) Irr(args ...any) any           { return wf.function("Irr", args...) }
func (wf *workFunction) IsErr(args ...any) any         { return wf.function("IsErr", args...) }
func (wf *workFunction) IsError(args ...any) any       { return wf.function("IsError", args...) }
func (wf *workFunction) IsEven(args ...any) any        { return wf.function("IsEven", args...) }
func (wf *workFunction) IsFormula(args ...any) any     { return wf.function("IsFormula", args...) }
func (wf *workFunction) IsLogical(args ...any) any     { return wf.function("IsLogical", args...) }
func (wf *workFunction) IsNA(args ...any) any          { return wf.function("IsNA", args...) }
func (wf *workFunction) IsNonText(args ...any) any     { return wf.function("IsNonText", args...) }
func (wf *workFunction) IsNumber(args ...any) any      { return wf.function("IsNumber", args...) }
func (wf *workFunction) ISO_Ceiling(args ...any) any   { return wf.function("ISO_Ceiling", args...) }
func (wf *workFunction) IsOdd(args ...any) any         { return wf.function("IsOdd", args...) }
func (wf *workFunction) IsoWeekNum(args ...any) any    { return wf.function("IsoWeekNum", args...) }
func (wf *workFunction) Ispmt(args ...any) any         { return wf.function("Ispmt", args...) }
func (wf *workFunction) IsText(args ...any) any        { return wf.function("IsText", args...) }
func (wf *workFunction) Kurt(args ...any) any          { return wf.function("Kurt", args...) }
func (wf *workFunction) Large(args ...any) any         { return wf.function("Large", args...) }
func (wf *workFunction) Lcm(args ...any) any           { return wf.function("Lcm", args...) }
func (wf *workFunction) LinEst(args ...any) any        { return wf.function("LinEst", args...) }
func (wf *workFunction) Ln(args ...any) any            { return wf.function("Ln", args...) }
func (wf *workFunction) Log(args ...any) any           { return wf.function("Log", args...) }
func (wf *workFunction) Log10(args ...any) any         { return wf.function("Log10", args...) }
func (wf *workFunction) LogEst(args ...any) any        { return wf.function("LogEst", args...) }
func (wf *workFunction) LogInv(args ...any) any        { return wf.function("LogInv", args...) }
func (wf *workFunction) LogNorm_Dist(args ...any) any  { return wf.function("LogNorm_Dist", args...) }
func (wf *workFunction) LogNorm_Inv(args ...any) any   { return wf.function("LogNorm_Inv", args...) }
func (wf *workFunction) LogNormDist(args ...any) any   { return wf.function("LogNormDist", args...) }
func (wf *workFunction) Lookup(args ...any) any        { return wf.function("Lookup", args...) }
func (wf *workFunction) Match(args ...any) any         { return wf.function("Match", args...) }
func (wf *workFunction) Max(args ...any) any           { return wf.function("Max", args...) }
func (wf *workFunction) MDeterm(args ...any) any       { return wf.function("MDeterm", args...) }
func (wf *workFunction) MDuration(args ...any) any     { return wf.function("MDuration", args...) }
func (wf *workFunction) Median(args ...any) any        { return wf.function("Median", args...) }
func (wf *workFunction) Min(args ...any) any           { return wf.function("Min", args...) }
func (wf *workFunction) MInverse(args ...any) any      { return wf.function("MInverse", args...) }
func (wf *workFunction) MIrr(args ...any) any          { return wf.function("MIrr", args...) }
func (wf *workFunction) MMult(args ...any) any         { return wf.function("MMult", args...) }
func (wf *workFunction) Mode(args ...any) any          { return wf.function("Mode", args...) }
func (wf *workFunction) Mode_Mult(args ...any) any     { return wf.function("Mode_Mult", args...) }
func (wf *workFunction) Mode_Sngl(args ...any) any     { return wf.function("Mode_Sngl", args...) }
func (wf *workFunction) MRound(args ...any) any        { return wf.function("MRound", args...) }
func (wf *workFunction) MultiNomial(args ...any) any   { return wf.function("MultiNomial", args...) }
func (wf *workFunction) Munit(args ...any) any         { return wf.function("Munit", args...) }
func (wf *workFunction) NegBinom_Dist(args ...any) any { return wf.function("NegBinom_Dist", args...) }
func (wf *workFunction) NegBinomDist(args ...any) any  { return wf.function("NegBinomDist", args...) }
func (wf *workFunction) NetworkDays(args ...any) any   { return wf.function("NetworkDays", args...) }
func (wf *workFunction) NetworkDays_Intl(args ...any) any {
	return wf.function("NetworkDays_Intl", args...)
}
func (wf *workFunction) Nominal(args ...any) any     { return wf.function("Nominal", args...) }
func (wf *workFunction) Norm_Dist(args ...any) any   { return wf.function("Norm_Dist", args...) }
func (wf *workFunction) Norm_Inv(args ...any) any    { return wf.function("Norm_Inv", args...) }
func (wf *workFunction) Norm_S_Dist(args ...any) any { return wf.function("Norm_S_Dist", args...) }
func (wf *workFunction) Norm_S_Inv(args ...any) any  { return wf.function("Norm_S_Inv", args...) }
func (wf *workFunction) NormDist(args ...any) any    { return wf.function("NormDist", args...) }
func (wf *workFunction) NormInv(args ...any) any     { return wf.function("NormInv", args...) }
func (wf *workFunction) NormSDist(args ...any) any   { return wf.function("NormSDist", args...) }
func (wf *workFunction) NormSInv(args ...any) any    { return wf.function("NormSInv", args...) }
func (wf *workFunction) NPer(args ...any) any        { return wf.function("NPer", args...) }
func (wf *workFunction) Npv(args ...any) any         { return wf.function("Npv", args...) }
func (wf *workFunction) NumberValue(args ...any) any { return wf.function("NumberValue", args...) }
func (wf *workFunction) Oct2Bin(args ...any) any     { return wf.function("Oct2Bin", args...) }
func (wf *workFunction) Oct2Dec(args ...any) any     { return wf.function("Oct2Dec", args...) }
func (wf *workFunction) Oct2Hex(args ...any) any     { return wf.function("Oct2Hex", args...) }
func (wf *workFunction) Odd(args ...any) any         { return wf.function("Odd", args...) }
func (wf *workFunction) OddFPrice(args ...any) any   { return wf.function("OddFPrice", args...) }
func (wf *workFunction) OddFYield(args ...any) any   { return wf.function("OddFYield", args...) }
func (wf *workFunction) OddLPrice(args ...any) any   { return wf.function("OddLPrice", args...) }
func (wf *workFunction) OddLYield(args ...any) any   { return wf.function("OddLYield", args...) }
func (wf *workFunction) Or(args ...any) any          { return wf.function("Or", args...) }
func (wf *workFunction) PDuration(args ...any) any   { return wf.function("PDuration", args...) }
func (wf *workFunction) Pearson(args ...any) any     { return wf.function("Pearson", args...) }
func (wf *workFunction) Percentile(args ...any) any  { return wf.function("Percentile", args...) }
func (wf *workFunction) Percentile_Exc(args ...any) any {
	return wf.function("Percentile_Exc", args...)
}
func (wf *workFunction) Percentile_Inc(args ...any) any {
	return wf.function("Percentile_Inc", args...)
}
func (wf *workFunction) PercentRank(args ...any) any { return wf.function("PercentRank", args...) }
func (wf *workFunction) PercentRank_Exc(args ...any) any {
	return wf.function("PercentRank_Exc", args...)
}
func (wf *workFunction) PercentRank_Inc(args ...any) any {
	return wf.function("PercentRank_Inc", args...)
}
func (wf *workFunction) Permut(args ...any) any       { return wf.function("Permut", args...) }
func (wf *workFunction) Permutationa(args ...any) any { return wf.function("Permutationa", args...) }
func (wf *workFunction) Phi(args ...any) any          { return wf.function("Phi", args...) }
func (wf *workFunction) Phonetic(args ...any) any     { return wf.function("Phonetic", args...) }
func (wf *workFunction) Pi(args ...any) any           { return wf.function("Pi", args...) }
func (wf *workFunction) Pmt(args ...any) any          { return wf.function("Pmt", args...) }
func (wf *workFunction) Poisson(args ...any) any      { return wf.function("Poisson", args...) }
func (wf *workFunction) Poisson_Dist(args ...any) any { return wf.function("Poisson_Dist", args...) }
func (wf *workFunction) 電源(args ...any) any           { return wf.function("電源", args...) }
func (wf *workFunction) Ppmt(args ...any) any         { return wf.function("Ppmt", args...) }
func (wf *workFunction) Price(args ...any) any        { return wf.function("Price", args...) }
func (wf *workFunction) PriceDisc(args ...any) any    { return wf.function("PriceDisc", args...) }
func (wf *workFunction) PriceMat(args ...any) any     { return wf.function("PriceMat", args...) }
func (wf *workFunction) Prob(args ...any) any         { return wf.function("Prob", args...) }
func (wf *workFunction) Product(args ...any) any      { return wf.function("Product", args...) }
func (wf *workFunction) Proper(args ...any) any       { return wf.function("Proper", args...) }
func (wf *workFunction) Pv(args ...any) any           { return wf.function("Pv", args...) }
func (wf *workFunction) Quartile(args ...any) any     { return wf.function("Quartile", args...) }
func (wf *workFunction) Quartile_Exc(args ...any) any { return wf.function("Quartile_Exc", args...) }
func (wf *workFunction) Quartile_Inc(args ...any) any { return wf.function("Quartile_Inc", args...) }
func (wf *workFunction) Quotient(args ...any) any     { return wf.function("Quotient", args...) }
func (wf *workFunction) Radians(args ...any) any      { return wf.function("Radians", args...) }
func (wf *workFunction) RandBetween(args ...any) any  { return wf.function("RandBetween", args...) }
func (wf *workFunction) Rank(args ...any) any         { return wf.function("Rank", args...) }
func (wf *workFunction) Rank_Avg(args ...any) any     { return wf.function("Rank_Avg", args...) }
func (wf *workFunction) Rank_Eq(args ...any) any      { return wf.function("Rank_Eq", args...) }
func (wf *workFunction) Rate(args ...any) any         { return wf.function("Rate", args...) }
func (wf *workFunction) Received(args ...any) any     { return wf.function("Received", args...) }
func (wf *workFunction) Replace(args ...any) any      { return wf.function("Replace", args...) }
func (wf *workFunction) ReplaceB(args ...any) any     { return wf.function("ReplaceB", args...) }
func (wf *workFunction) Rept(args ...any) any         { return wf.function("Rept", args...) }
func (wf *workFunction) Roman(args ...any) any        { return wf.function("Roman", args...) }
func (wf *workFunction) Round(args ...any) any        { return wf.function("Round", args...) }
func (wf *workFunction) RoundDown(args ...any) any    { return wf.function("RoundDown", args...) }
func (wf *workFunction) RoundUp(args ...any) any      { return wf.function("RoundUp", args...) }
func (wf *workFunction) Rri(args ...any) any          { return wf.function("Rri", args...) }
func (wf *workFunction) RSq(args ...any) any          { return wf.function("RSq", args...) }
func (wf *workFunction) RTD(args ...any) any          { return wf.function("RTD", args...) }
func (wf *workFunction) 検索(args ...any) any           { return wf.function("検索", args...) }
func (wf *workFunction) SearchB(args ...any) any      { return wf.function("SearchB", args...) }
func (wf *workFunction) Sec(args ...any) any          { return wf.function("Sec", args...) }
func (wf *workFunction) Sech(args ...any) any         { return wf.function("Sech", args...) }
func (wf *workFunction) SeriesSum(args ...any) any    { return wf.function("SeriesSum", args...) }
func (wf *workFunction) Sinh(args ...any) any         { return wf.function("Sinh", args...) }
func (wf *workFunction) Skew(args ...any) any         { return wf.function("Skew", args...) }
func (wf *workFunction) Skew_p(args ...any) any       { return wf.function("Skew_p", args...) }
func (wf *workFunction) Sln(args ...any) any          { return wf.function("Sln", args...) }
func (wf *workFunction) Slope(args ...any) any        { return wf.function("Slope", args...) }
func (wf *workFunction) Small(args ...any) any        { return wf.function("Small", args...) }
func (wf *workFunction) SqrtPi(args ...any) any       { return wf.function("SqrtPi", args...) }
func (wf *workFunction) Standardize(args ...any) any  { return wf.function("Standardize", args...) }
func (wf *workFunction) StDev(args ...any) any        { return wf.function("StDev", args...) }
func (wf *workFunction) StDev_P(args ...any) any      { return wf.function("StDev_P", args...) }
func (wf *workFunction) StDev_S(args ...any) any      { return wf.function("StDev_S", args...) }
func (wf *workFunction) StDevP(args ...any) any       { return wf.function("StDevP", args...) }
func (wf *workFunction) StEyx(args ...any) any        { return wf.function("StEyx", args...) }
func (wf *workFunction) Substitute(args ...any) any   { return wf.function("Substitute", args...) }
func (wf *workFunction) Subtotal(args ...any) any     { return wf.function("Subtotal", args...) }
func (wf *workFunction) Sum(args ...any) any          { return wf.function("Sum", args...) }
func (wf *workFunction) SumIf(args ...any) any        { return wf.function("SumIf", args...) }
func (wf *workFunction) SumIfs(args ...any) any       { return wf.function("SumIfs", args...) }
func (wf *workFunction) SumProduct(args ...any) any   { return wf.function("SumProduct", args...) }
func (wf *workFunction) SumSq(args ...any) any        { return wf.function("SumSq", args...) }
func (wf *workFunction) SumX2MY2(args ...any) any     { return wf.function("SumX2MY2", args...) }
func (wf *workFunction) SumX2PY2(args ...any) any     { return wf.function("SumX2PY2", args...) }
func (wf *workFunction) SumXMY2(args ...any) any      { return wf.function("SumXMY2", args...) }
func (wf *workFunction) Syd(args ...any) any          { return wf.function("Syd", args...) }
func (wf *workFunction) T_Dist(args ...any) any       { return wf.function("T_Dist", args...) }
func (wf *workFunction) T_Dist_2T(args ...any) any    { return wf.function("T_Dist_2T", args...) }
func (wf *workFunction) T_Dist_RT(args ...any) any    { return wf.function("T_Dist_RT", args...) }
func (wf *workFunction) T_Inv(args ...any) any        { return wf.function("T_Inv", args...) }
func (wf *workFunction) T_Inv_2T(args ...any) any     { return wf.function("T_Inv_2T", args...) }
func (wf *workFunction) T_Test(args ...any) any       { return wf.function("T_Test", args...) }
func (wf *workFunction) Tanh(args ...any) any         { return wf.function("Tanh", args...) }
func (wf *workFunction) TBillEq(args ...any) any      { return wf.function("TBillEq", args...) }
func (wf *workFunction) TBillPrice(args ...any) any   { return wf.function("TBillPrice", args...) }
func (wf *workFunction) TBillYield(args ...any) any   { return wf.function("TBillYield", args...) }
func (wf *workFunction) TDist(args ...any) any        { return wf.function("TDist", args...) }
func (wf *workFunction) テキスト(args ...any) any         { return wf.function("テキスト", args...) }
func (wf *workFunction) TInv(args ...any) any         { return wf.function("TInv", args...) }
func (wf *workFunction) Transpose(args ...any) any    { return wf.function("Transpose", args...) }
func (wf *workFunction) Trend(args ...any) any        { return wf.function("Trend", args...) }
func (wf *workFunction) Trim(args ...any) any         { return wf.function("Trim", args...) }
func (wf *workFunction) TrimMean(args ...any) any     { return wf.function("TrimMean", args...) }
func (wf *workFunction) TTest(args ...any) any        { return wf.function("TTest", args...) }
func (wf *workFunction) Unichar(args ...any) any      { return wf.function("Unichar", args...) }
func (wf *workFunction) Unicode(args ...any) any      { return wf.function("Unicode", args...) }
func (wf *workFunction) USDollar(args ...any) any     { return wf.function("USDollar", args...) }
func (wf *workFunction) Var(args ...any) any          { return wf.function("Var", args...) }
func (wf *workFunction) Var_P(args ...any) any        { return wf.function("Var_P", args...) }
func (wf *workFunction) Var_S(args ...any) any        { return wf.function("Var_S", args...) }
func (wf *workFunction) VarP(args ...any) any         { return wf.function("VarP", args...) }
func (wf *workFunction) Vdb(args ...any) any          { return wf.function("Vdb", args...) }
func (wf *workFunction) VLookup(args ...any) any      { return wf.function("VLookup", args...) }
func (wf *workFunction) WebService(args ...any) any   { return wf.function("WebService", args...) }
func (wf *workFunction) Weekday(args ...any) any      { return wf.function("Weekday", args...) }
func (wf *workFunction) WeekNum(args ...any) any      { return wf.function("WeekNum", args...) }
func (wf *workFunction) Weibull(args ...any) any      { return wf.function("Weibull", args...) }
func (wf *workFunction) Weibull_Dist(args ...any) any { return wf.function("Weibull_Dist", args...) }
func (wf *workFunction) WorkDay(args ...any) any      { return wf.function("WorkDay", args...) }
func (wf *workFunction) WorkDay_Intl(args ...any) any { return wf.function("WorkDay_Intl", args...) }
func (wf *workFunction) Xirr(args ...any) any         { return wf.function("Xirr", args...) }
func (wf *workFunction) Xnpv(args ...any) any         { return wf.function("Xnpv", args...) }
func (wf *workFunction) Xor(args ...any) any          { return wf.function("Xor", args...) }
func (wf *workFunction) YearFrac(args ...any) any     { return wf.function("YearFrac", args...) }
func (wf *workFunction) YieldDisc(args ...any) any    { return wf.function("YieldDisc", args...) }
func (wf *workFunction) YieldMat(args ...any) any     { return wf.function("YieldMat", args...) }
func (wf *workFunction) Z_Test(args ...any) any       { return wf.function("Z_Test", args...) }
func (wf *workFunction) ZTest(args ...any) any        { return wf.function("ZTest", args...) }
