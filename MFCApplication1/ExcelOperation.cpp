#include "stdafx.h"
#include "ExcelOperation.h"
#include <comutil.h>
#include <string>

using namespace std;

namespace
{
	void ConstCharConver(const char* pFileName, CString &pWideChar)
	{
		//计算char *数组大小，以字节为单位，一个汉字占两个字节
		int charLen = strlen(pFileName);

		//计算多字节字符的大小，按字符计算。
		int len = MultiByteToWideChar(CP_ACP, 0, pFileName, charLen, NULL, 0);

		//为宽字节字符数组申请空间，数组大小为按字节计算的多字节字符大小
		TCHAR *buf = new TCHAR[len + 1];

		//多字节编码转换成宽字节编码
		MultiByteToWideChar(CP_ACP, 0, pFileName, charLen, buf, len);

		buf[len] = '\0';  //添加字符串结尾，注意不是len+1

		//将TCHAR数组转换为CString

		pWideChar.Append(buf);
	}
}

ExcelOperation* ExcelOperation::excel = nullptr;

ExcelOperation::ExcelOperation()
{
	lpDisp = NULL;
	excelVer = 0;
}

ExcelOperation::~ExcelOperation()
{
	try
	{
		sheet.ReleaseDispatch();
		sheets.ReleaseDispatch();
		book.ReleaseDispatch();
		books.ReleaseDispatch();
		ExcelApp.ReleaseDispatch();
		ExcelApp.Quit();
		//退出伪装的app
		if (!ExcelApp_fake.get_ActiveSheet())
		{
			books_fake.ReleaseDispatch();
			ExcelApp_fake.ReleaseDispatch();
			ExcelApp_fake.Quit();
		}
	}
	catch (COleDispatchException*)
	{
		//AfxMessageBox(Notice_get_by_id(IDS_POW_OFF_EXCEL_FAIL));
		AfxMessageBox(_T("关闭Excel服务出错。"));
	}
}

//获取对象
ExcelOperation* ExcelOperation::getInstance()
{
	if (excel == NULL)
	{
		excel = new ExcelOperation();
	}
	return excel;
}
//销毁对象
void ExcelOperation::destroyInstance()
{
	if (excel != NULL)
	{
		excel->~ExcelOperation();
		delete excel;
		excel = NULL;
	}
}

/************************************************************************/
/* 判断Excel版本号                                                      */
/************************************************************************/
BOOL ExcelOperation::judgeExcelVer(int Ver)
{
	HKEY hkey;
	int ret;
	CString str;
	LONG len;
	str.Format(_T("Excel.Application.%d"), Ver);
	str += _T("\\CLSID");
	ret = RegCreateKey(HKEY_CLASSES_ROOT, str, &hkey);
	if (ret == ERROR_SUCCESS)
	{
		RegQueryValue(HKEY_CLASSES_ROOT, str, NULL, &len);
		//如果注册表中 HKEY_CLASSES_ROOT\Excel.Application.x\CPLSID中的值为空，则读取到'\0'，长度为2
		return len == 2 ? FALSE : TRUE;
	}
	else
	{
		return FALSE;
	}
}

/************************************************************************/
/* 启动Excel服务，传入的字符串参数格式为 office ****                    */
/************************************************************************/
BOOL ExcelOperation::createServer(CString officeVer)
{
	//去除前后空格
	officeVer.Trim();
	//获取版本号字符
	CString verNum = officeVer.Right(4);
	int ver = _ttoi(verNum);
	switch (ver)
	{
	case 2003:
		if (judgeExcelVer(11))
		{
			if (ExcelApp.CreateDispatch(_T("Excel.Application.11"), NULL))
			{
				excelVer = 2003;
			}
			else
			{
				return FALSE;
			}
		}
		else
		{
			return FALSE;
		}
		break;
	case 2007:
		if (judgeExcelVer(12))
		{
			if (ExcelApp.CreateDispatch(_T("Excel.Application.12"), NULL))
			{
				excelVer = 2007;
			}
			else
			{
				return FALSE;
			}
		}
		else
		{
			return FALSE;
		}
		break;
	case 2010:
		if (judgeExcelVer(14))
		{
			if (ExcelApp.CreateDispatch(_T("Excel.Application.14"), NULL))
			{
				excelVer = 2010;
			}
			else
			{
				return FALSE;
			}
		}
		else
		{
			return FALSE;
		}
		break;
	case 2013:
		if (judgeExcelVer(15))
		{
			if (ExcelApp.CreateDispatch(_T("Excel.Application.15"), NULL))
			{
				excelVer = 2013;
			}
			else
			{
				return FALSE;
			}
		}
		else
		{
			return FALSE;
		}
		break;
	case 2016:
		if (judgeExcelVer(16))
		{
			if (ExcelApp.CreateDispatch(_T("Excel.Application.16"), NULL))
			{
				excelVer = 2016;
			}
			else
			{
				return FALSE;
			}
		}
		else
		{
			return FALSE;
		}
		break;
	}
	return TRUE;
}

BOOL ExcelOperation::init()
{
	CString strOfficeVer[5] = { _T("office 2003"), _T("office 2007"), _T("office 2010"), _T("office 2013"), _T("office 2016") };
	BOOL result = FALSE;
	for (int i = 4; i >= 0; i--)
	{
		if (!createServer(strOfficeVer[i]))
			continue;
		else
		{
			result = TRUE;
		}
	}
	if (excelVer == 0)
	{
		result = FALSE;
	}
	return result;
}

//显示隐藏excel
void ExcelOperation::setView(bool show)
{
	ExcelApp.put_Visible(show);
	ExcelApp.put_UserControl(show);
}

//打开excel
BOOL ExcelOperation::openExcelFile(CString strBookPath, const char* excleTemplate)
{
	lpDisp = NULL;

	/*判断当前Excel的版本*/
	CString strExcelVersion = ExcelApp.get_Version();
	int iStart = 0;
	strExcelVersion = strExcelVersion.Tokenize(_T("."), iStart);

	setView(false);

	/*得到工作簿容器*/
	books.AttachDispatch(ExcelApp.get_Workbooks());

	/*打开一个工作簿，如不存在，则新增一个工作簿*/
	//CString strBookPath;
	//ConstCharConver(path, strBookPath);
	try
	{
		/*打开一个工作簿*/
		lpDisp = books.Open(strBookPath,
			vtMissing, vtMissing, vtMissing, vtMissing, vtMissing,
			vtMissing, vtMissing, vtMissing, vtMissing, vtMissing,
			vtMissing, vtMissing, vtMissing, vtMissing);
		book.AttachDispatch(lpDisp);
	}
	catch (...)
	{
		/*增加一个新的工作簿*/
		if (excleTemplate == nullptr)
		{
			lpDisp = books.Add(vtMissing);
		}
		else
		{
			lpDisp = books.Add(_variant_t(excleTemplate));
		}
		book.AttachDispatch(lpDisp);
		saveExcelAs(strBookPath);
	}
	return true;
}

/************************************************************************/
/* 另存为                                                               */
/************************************************************************/
void ExcelOperation::saveExcelAs(CString savePathCSt)
{
	
	//ConstCharConver(savePath, savePathCSt);
	savePathCSt.Trim();
	book.SaveAs(_variant_t(savePathCSt),
		vtMissing, vtMissing, vtMissing, vtMissing, vtMissing, 0,
		vtMissing, vtMissing, vtMissing, vtMissing, vtMissing);
}

//打开Sheet
void ExcelOperation::OpenSheet(const char* sheetName)
{
	LPDISPATCH lpDisp = NULL;
	/*得到工作簿中的Sheet的容器*/
	sheets.AttachDispatch(book.get_Sheets());
	/*打开一个Sheet，如不存在，就新增一个Sheet*/
	CString strSheetName(sheetName);
	try
	{
		/*打开一个已有的Sheet*/
		lpDisp = sheets.get_Item(_variant_t(strSheetName));
		sheet.AttachDispatch(lpDisp);
		currentRange.AttachDispatch(sheet.get_Cells());
	}
	catch (...)
	{
		/*创建一个新的Sheet*/
		lpDisp = sheets.Add(vtMissing, vtMissing, _variant_t((long)1), vtMissing);
		sheet.AttachDispatch(lpDisp);
		sheet.put_Name(strSheetName);
	}
}



int ExcelOperation::getSheetCount()
{
	return sheets.get_Count();
}

void ExcelOperation::OpenSheetwithId(long tableId)
{
	LPDISPATCH lpDis = nullptr;
	currentRange.ReleaseDispatch();
	lpDis = sheets.get_Item(COleVariant((long)tableId));
	sheet.AttachDispatch(lpDis, true);
	currentRange.AttachDispatch(sheet.get_Cells(), true);

}

int ExcelOperation::getColumnCount()
{
	CRange range;
	CRange usedRange;

	usedRange.AttachDispatch(sheet.get_UsedRange(), true);
	range.AttachDispatch(usedRange.get_Columns(), true);
	int count = range.get_Count();

	usedRange.ReleaseDispatch();
	range.ReleaseDispatch();

	return count;
}

int ExcelOperation::getRowCount()
{
	CRange range;
	CRange usedRange;

	usedRange.AttachDispatch(sheet.get_UsedRange(), true);
	range.AttachDispatch(usedRange.get_Rows(), true);

	int count = range.get_Count();

	usedRange.ReleaseDispatch();
	range.ReleaseDispatch();

	return count;
}

CString ExcelOperation::getCellCString(long iRow, long iColumn)
{
	COleVariant vResult;
	CString str;
	//字符串  
	CRange range;
	range.AttachDispatch(currentRange.get_Item(COleVariant((long)iRow), COleVariant((long)iColumn)).pdispVal, true);
	vResult = range.get_Value2();
	range.ReleaseDispatch();



	if (vResult.vt == VT_BSTR)
	{
		str = vResult.bstrVal;
	}
	//整数  
	else if (vResult.vt == VT_INT)
	{
		str.Format(_T("%d"), vResult.pintVal);
	}
	//8字节的数字   
	else if (vResult.vt == VT_R8)
	{
		str.Format(_T("%0.0f"), vResult.dblVal);
	}
	//时间格式  
	else if (vResult.vt == VT_DATE)
	{
		SYSTEMTIME st;
		VariantTimeToSystemTime(vResult.date, &st);
		CTime tm(st);
		str = tm.Format(_T("%Y-%m-%d"));

	}
	//单元格空的  
	else if (vResult.vt == VT_EMPTY)
	{
		str = "";
	}

	return str;
}



double ExcelOperation::getCellDouble(long iRow, long iColumn)
{
	double rtn_value = 0;
	COleVariant vresult;
	//字符串  
	CRange range;
	range.AttachDispatch(currentRange.get_Item(COleVariant((long)iRow), COleVariant((long)iColumn)).pdispVal, true);
	vresult = range.get_Value2();
	range.ReleaseDispatch();
	

	if (vresult.vt == VT_R8)
	{
		rtn_value = vresult.dblVal;
	}

	return rtn_value;
}

int ExcelOperation::getCellInt(long iRow, long iColumn)
{
	int num;
	COleVariant vresult;

	CRange range;
	range.AttachDispatch(currentRange.get_Item(COleVariant((long)iRow), COleVariant((long)iColumn)).pdispVal, true);
	vresult = range.get_Value2();
	range.ReleaseDispatch();

	num = static_cast<int>(vresult.dblVal);

	return num;
}

////读取sheet数据
//int ExcelOperation::ReadSheetgetRowCol(const char* sheetName)
//{
//	CRange  usedRange;
//	CRange  cRange;
//	LPDISPATCH lpDisp = NULL;
//	int carray[2];
//	/*得到工作簿中的Sheet的容器*/
//	sheets.AttachDispatch(book.get_Sheets());
//	/*打开一个Sheet，如不存在，就新增一个Sheet*/
//	CString strSheetName(sheetName);
//	try
//	{
//		/*打开一个已有的Sheet*/
//		lpDisp = sheets.get_Item(_variant_t(strSheetName));
//		sheet.AttachDispatch(lpDisp);
//		// 获得使用的区域Range( 区域 )
//		usedRange.AttachDispatch(sheet.get_UsedRange(), true);
//		//获取已经使用的行数
//		cRange.AttachDispatch(usedRange.get_Rows());
//		carray[0] = cRange.get_Count();
//		// 获得使用的列数
//		cRange.ReleaseDispatch();
//		cRange.AttachDispatch(usedRange.get_Columns(), TRUE);
//		carray[1] = cRange.get_Count();
//		return carray;
//	}
//	catch (...)
//	{
//		/*提示错误*/
//		AfxMessageBox(_T("SheetName not exists..."));
//
//	}
//}



//设置单元格格式
void ExcelOperation::setCellsFormat(const char* cellBeginChar, const char* cellBEndChar, const char* cellFormat)
{
	CString cellBegin, cellEnd, format;
	ConstCharConver(cellBeginChar, cellBegin);
	ConstCharConver(cellBEndChar, cellEnd);
	ConstCharConver(cellFormat, format);
	lpDisp = sheet.get_Range(_variant_t(cellBegin), _variant_t(cellEnd));
	range.AttachDispatch(lpDisp);
	range.put_NumberFormatLocal(_variant_t(format));
}

//设置单个单元格格式
void ExcelOperation::setCellFormat(const char* ccellIndexChar, const char* cellFormat)
{
	CString cellIndex, cellFormatChar;
	ConstCharConver(ccellIndexChar, cellIndex);
	ConstCharConver(cellFormat, cellFormatChar);
	range = sheet.get_Range(_variant_t(cellIndex), _variant_t(cellIndex));
	range.put_NumberFormat(_variant_t(cellFormatChar));
}

int ExcelOperation::getCellcolor(long iRow, long iColumn)
{
	CFont0 font;
	CRange range;
	range.AttachDispatch(currentRange.get_Item(COleVariant((long)iRow), COleVariant((long)iColumn)).pdispVal, true);
	font.AttachDispatch(range.get_Font());
	COLORREF color = (long)font.get_Color().dblVal; //获取颜色
	return (int)color;
}


//设置单个单元格的值
void ExcelOperation::setCellValue(const char* ccellIndexChar, const char* valueChar)
{
	CString cellIndex, value;
	ConstCharConver(ccellIndexChar, cellIndex);
	ConstCharConver(valueChar, value);
	range = sheet.get_Range(_variant_t(cellIndex), _variant_t(cellIndex));
	range.put_Value2(_variant_t(value));
}

//设置单个单元格的值
void ExcelOperation::setCellCStringValue(CString cellIndex, CString value)
{
	range = sheet.get_Range(_variant_t(cellIndex), _variant_t(cellIndex));
	range.put_Value2(COleVariant(value));
}

/************************************************************************/
/* 初始化界面
非MFC项目使用*/
/************************************************************************/
void ExcelOperation::InitializeUI()
{
	if (S_OK != CoInitialize(NULL)){
		AfxMessageBox(_T("Initialize com failed..."));
		return;
	}
}

/************************************************************************/
/* 释放资源                                                               */
/************************************************************************/
void ExcelOperation::UnInitializeUI()
{
	CoUninitialize();
}

/************************************************************************/
/* 保存                                                                 */
/************************************************************************/
void ExcelOperation::saveExcel()
{
	ExcelApp.put_DisplayAlerts(FALSE);
	//book.Close(vtMissing, vtMissing, vtMissing);
	book.Save();
}

/************************************************************************/
/* 关闭*/
/************************************************************************/
void ExcelOperation::Close()
{
	sheets.ReleaseDispatch();
	sheet.ReleaseDispatch();
	currentRange.ReleaseDispatch();
	range.ReleaseDispatch();
	book.ReleaseDispatch();
	books.ReleaseDispatch();
	books.Close();
}
