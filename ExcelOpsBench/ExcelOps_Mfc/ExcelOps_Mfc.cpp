#include "pch.h"
#include "framework.h"
#include "CApplication.h"
#include "CWorkbooks.h"
#include "CWorkbook.h"
#include "CWorksheets.h"
#include "CWorksheet.h"
#include "CRanges.h"

#pragma push_macro("DialogBox")
#undef DialogBox
#include "CRange.h"
#pragma pop_macro("DialogBox")

int main()
{
	auto mainRoutine = [](const int count) -> int
	{
		const long xlCalculationAutomatic = -4105;	// Excel が再計算を制御します。
		const long xlCalculationManual = -4135;		// ユーザーが要求すると、計算が完了します。
		const long xlCalculationSemiautomatic = 2;	// Excel が再計算を制御しますが、テーブル内の変更は無視します。

		// Initialize a NULL variant
		COleVariant vtNull;
		vtNull.vt = VT_NULL;

		// Get path to Documents folder
		wchar_t szPath[MAX_PATH]{};
		SHGetFolderPathW(nullptr, CSIDL_MYDOCUMENTS, nullptr, SHGFP_TYPE_CURRENT, szPath);
		PathAppendW(szPath, L"Sample.xlsx");

		// Start Excel and get Application object
		CApplication excel;
		if (const auto hr = excel.CreateDispatch(_T("Excel.Application")); FAILED(hr))
			return -2;

		// この設定をすると高速化する
		excel.put_Visible(VARIANT_FALSE);
		excel.put_ScreenUpdating(VARIANT_FALSE);
		//excel.put_Calculation(xlCalculationManual);
		excel.put_EnableEvents(VARIANT_FALSE);

		// Get Workbooks collection
		CWorkbooks workbooks(excel.get_Workbooks());

		// Open the workbook
		CWorkbook workbook;
		workbook.AttachDispatch(workbooks.Open(CString(szPath), 
			vtNull, // UpdateLinks
			vtNull, // ReadOnly
			vtNull, // Format
			vtNull, // Password
			vtNull, // WriteResPassword
			vtNull, // IgnoreReadOnlyRecommended
			vtNull, // Origin
			vtNull, // Delimiter
			vtNull, // Editable
			vtNull, // Notify
			vtNull, // Converter
			vtNull, // AddToMru
			vtNull, // Local
			vtNull  // CorruptLoad
		));

		// Get the first worksheet
		CWorksheets worksheets(workbook.get_Worksheets());
		COleVariant vtIndex((short)1);
		CWorksheet worksheet(worksheets.get_Item(vtIndex));

		// Get the cells collection
		CRanges cells(worksheet.get_Cells());

		// Set A1 to 0
		COleVariant vtNewValue((long)0);
		CRange cellA1(cells.get_Item(vtIndex));
		cellA1.put_Value2(vtNewValue);

		for (int i = 0; i < count; i++)
		{
			// Increment A1
			COleVariant vtValue = cellA1.get_Value2();
			vtValue.dblVal++;
			cellA1.put_Value2(vtValue);
		}

		// Quit Excel without saving
		COleVariant vtFalse(VARIANT_FALSE);
		workbook.Close(vtFalse, vtNull, vtNull);
		excel.Quit();

		return 0;
	};

	if (const auto hr = CoInitialize(nullptr); FAILED(hr))
		return -1;

	const auto start = GetTickCount64();
	const auto result = mainRoutine(1024);
	const auto end = GetTickCount64();
	const auto elapsed = end - start;
	printf_s("ExcelOps_Mfc: %llu ms\n", elapsed);
	CoUninitialize();
	return result;
}

