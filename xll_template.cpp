// xll_template.cpp - Sample xll project.
#include <cmath> // for double tgamma(double)
#include "xll_template.h"

using namespace xll;

// Throw on float
AddIn xai_tgamma(
	// Return double, C++ name of function, Excel name.
	Function(XLL_DOUBLE, "xll_tgamma", "TGAMMA")
	// Array of function arguments.
	.Arguments({
		Arg(XLL_DOUBLE, "x", "is the value for which you want to calculate Gamma.")
	})
	// Function Wizard help.
	.FunctionHelp("Return the Gamma function value.")
	// Function Wizard category.
	.Category("MATH")
	// URL linked to `Help on this function`.
	.HelpTopic("https://docs.microsoft.com/en-us/cpp/c-runtime-library/reference/tgamma-tgammaf-tgammal")
	.Documentation(R"xyzyx(
The <i>Gamma</i> function is \(\Gamma(x) = \int_0^\infty t^{x - 1} e^{-t}\,dt\), \(x \ge 0\).
If \(n\) is a natural number then \(\Gamma(n + 1) = n! = n(n - 1)\cdots 1\).
<p>
Any valid HTML using <a href="https://katex.org/" target="_blank">KaTeX</a> can 
be used for documentation.
)xyzyx")
);
// WINAPI calling convention must be specified
double WINAPI xll_tgamma(double x)
{
#pragma XLLEXPORT // must be specified to export function
	double result = std::numeric_limits<double>::quiet_NaN();

	try {
		result = tgamma(x);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}
	catch (...) {
		XLL_ERROR("Unknown error in xll_tgamma");
	}

	return result;
}

// Press Alt-F8 then type 'XLL.MACRO' to call 'xll_macro'
// See https://xlladdins.github.io/Excel4Macros/
AddIn xai_macro(
	// C++ function, Excel name of macro
	Macro("xll_macro", "XLL.MACRO")
);
// Macros must have `int WINAPI (*)(void)` signature.
int WINAPI xll_macro(void)
{
#pragma XLLEXPORT
	// https://xlladdins.github.io/Excel4Macros/reftext.html
	// A1 style instead of default R1C1.
	OPER reftext = Excel(xlfReftext, Excel(xlfActiveCell), OPER(true));
	// UTF-8 strings can be used.
	Excel(xlcAlert, OPER("XLL.MACRO called with активный 细胞: ") & reftext);

	return TRUE;
}

#ifdef _DEBUG

// Create XML documentation and index.html in `$(TargetPath)` folder.
// Use `xsltproc file.xml -o file.html` to create HTML documentation.
Auto<Open> xao_template_docs([]() {

	return Documentation("MATH", "Documentation for the xll_template add-in.");

});

#endif // _DEBUG
