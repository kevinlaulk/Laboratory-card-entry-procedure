#pragma once
#include "CRange.h"
#include "CApplication.h"
#include "CWorkbook.h"
#include "CWorkbooks.h"
#include "CWorksheet.h"
#include "CWorksheets.h"
#include "CValidation.h"
#include "Cnterior.h"
#include "CFont0.h"
#include "CNames.h"

class ExcelOperation
{
private:
	ExcelOperation();
	~ExcelOperation();
	static ExcelOperation* excel;
public:
	static ExcelOperation* getInstance();
	static void destroyInstance();

	BOOL ExcelOperation::judgeExcelVer(int Ver);
	BOOL ExcelOperation::createServer(CString officeVer);
	BOOL init();

	void setView(bool show);
	void saveExcelAs(CString savePathCSt);
	BOOL openExcelFile(CString strBookPath, const char* excleTemplate = nullptr);
	void OpenSheet(const char* sheetName);
	void OpenSheetwithId(long tableId);
	int getSheetCount();
	int getColumnCount();
	int getRowCount();
	void setCellsFormat(const char* cellBeginChar, const char* cellBEndChar, const char* cellFormat);
	int getCellcolor(long iRow, long iColumn);
	void setCellFormat(const char* ccellIndexChar, const char* valueChar);
	void setCellValue(const char* ccellIndexChar, const char* valueChar);
	CString getCellCString(long iRow, long iColumn);
	void setCellCStringValue(CString cellIndex, CString value);
	double getCellDouble(long iRow, long iColumn);
	int getCellInt(long iRow, long iColumn);

	void InitializeUI();
	void UnInitializeUI();
	void saveExcel();
	void Close();
private:
	CApplication ExcelApp;
	CWorkbooks books;
	CWorkbook book;
	CWorksheets sheets;
	CWorksheet sheet;
	CRange range;
	CRange currentRange;
	CValidation validation;
	Cnterior interior;
	CFont0 font;
	CNames names;
	//CString filePath;
	LPDISPATCH lpDisp;
	//Î±ÔìµÄExcelApp
	CApplication ExcelApp_fake;
	CWorkbooks books_fake;
	COleVariant m_filePath;
	int excelVer;
};