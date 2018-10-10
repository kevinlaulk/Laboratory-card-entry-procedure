#include "ExcelOperation.h"
#include <comutil.h>

namespace
{
	void ConstCharConver(const char* pFileName, CString &pWideChar)
	{
		//����char *�����С�����ֽ�Ϊ��λ��һ������ռ�����ֽ�
		int charLen = strlen(pFileName);

		//������ֽ��ַ��Ĵ�С�����ַ����㡣
		int len = MultiByteToWideChar(CP_ACP, 0, pFileName, charLen, NULL, 0);

		//Ϊ���ֽ��ַ���������ռ䣬�����СΪ���ֽڼ���Ķ��ֽ��ַ���С
		TCHAR *buf = new TCHAR[len + 1];

		//���ֽڱ���ת���ɿ��ֽڱ���
		MultiByteToWideChar(CP_ACP, 0, pFileName, charLen, buf, len);

		buf[len] = '\0';  //����ַ�����β��ע�ⲻ��len+1

		//��TCHAR����ת��ΪCString

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
		//�˳�αװ��app
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
		AfxMessageBox(_T("�ر�Excel�������"));
	}
}

//��ȡ����
ExcelOperation* ExcelOperation::getInstance()
{
	if (excel == NULL)
	{
		excel = new ExcelOperation();
	}
	return excel;
}
//���ٶ���
void ExcelOperation::destroyInstance()
{
	if (excel != NULL)
	{
		delete excel;
		excel = NULL;
	}
}

/************************************************************************/
/* �ж�Excel�汾��                                                      */
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
		//���ע����� HKEY_CLASSES_ROOT\Excel.Application.x\CPLSID�е�ֵΪ�գ����ȡ��'\0'������Ϊ2
		return len == 2 ? FALSE : TRUE;
	}
	else
	{
		return FALSE;
	}
}

/************************************************************************/
/* ����Excel���񣬴�����ַ���������ʽΪ office ****                    */
/************************************************************************/
BOOL ExcelOperation::createServer(CString officeVer)
{
	//ȥ��ǰ��ո�
	officeVer.Trim();
	//��ȡ�汾���ַ�
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

//��ʾ����excel
void ExcelOperation::setView(bool show)
{
	ExcelApp.put_Visible(TRUE);
	ExcelApp.put_UserControl(FALSE);
}

//��excel
BOOL ExcelOperation::openExcelFile(const char*path, const char* excleTemplate)
{
	lpDisp = NULL;

	/*�жϵ�ǰExcel�İ汾*/
	CString strExcelVersion = ExcelApp.get_Version();
	int iStart = 0;
	strExcelVersion = strExcelVersion.Tokenize(_T("."), iStart);

	setView(false);

	/*�õ�����������*/
	books.AttachDispatch(ExcelApp.get_Workbooks());

	/*��һ�����������粻���ڣ�������һ��������*/
	CString strBookPath;
	ConstCharConver(path, strBookPath);
	try
	{
		/*��һ��������*/
		lpDisp = books.Open(strBookPath,
			vtMissing, vtMissing, vtMissing, vtMissing, vtMissing,
			vtMissing, vtMissing, vtMissing, vtMissing, vtMissing,
			vtMissing, vtMissing, vtMissing, vtMissing);
		book.AttachDispatch(lpDisp);
	}
	catch (...)
	{
		/*����һ���µĹ�����*/
		if (excleTemplate == nullptr)
		{
			lpDisp = books.Add(vtMissing);
		}
		else
		{
			lpDisp = books.Add(_variant_t(excleTemplate));
		}
		book.AttachDispatch(lpDisp);
		saveExcelAs(path);
	}
	return true;
}

/************************************************************************/
/* ���Ϊ                                                               */
/************************************************************************/
void ExcelOperation::saveExcelAs(const char* savePath)
{
	CString savePathCSt;
	ConstCharConver(savePath, savePathCSt);
	savePathCSt.Trim();
	book.SaveAs(_variant_t(savePathCSt),
		vtMissing, vtMissing, vtMissing, vtMissing, vtMissing, 0,
		vtMissing, vtMissing, vtMissing, vtMissing, vtMissing);
}

//��Sheet
void ExcelOperation::OpenSheet(const char* sheetName)
{
	LPDISPATCH lpDisp = NULL;
	/*�õ��������е�Sheet������*/
	sheets.AttachDispatch(book.get_Sheets());
	/*��һ��Sheet���粻���ڣ�������һ��Sheet*/
	CString strSheetName(sheetName);
	try
	{
		/*��һ�����е�Sheet*/
		lpDisp = sheets.get_Item(_variant_t(strSheetName));
		sheet.AttachDispatch(lpDisp);
	}
	catch (...)
	{
		/*����һ���µ�Sheet*/
		lpDisp = sheets.Add(vtMissing, vtMissing, _variant_t((long)1), vtMissing);
		sheet.AttachDispatch(lpDisp);
		sheet.put_Name(strSheetName);
	}
}

//���õ�Ԫ���ʽ
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

//���õ�����Ԫ���ʽ
void ExcelOperation::setCellFormat(const char* ccellIndexChar, const char* cellFormat)
{
	CString cellIndex, cellFormatChar;
	ConstCharConver(ccellIndexChar, cellIndex);
	ConstCharConver(cellFormat, cellFormatChar);
	range = sheet.get_Range(_variant_t(cellIndex), _variant_t(cellIndex));
	range.put_NumberFormat(_variant_t(cellFormatChar));
}


//���õ�����Ԫ���ֵ
void ExcelOperation::setCellValue(const char* ccellIndexChar, const char* valueChar)
{
	CString cellIndex , value;
	ConstCharConver(ccellIndexChar, cellIndex);
	ConstCharConver(valueChar, value);
	range = sheet.get_Range(_variant_t(cellIndex), _variant_t(cellIndex));
	range.put_Value2(_variant_t(value));
}

/************************************************************************/
/* ��ʼ������  
��MFC��Ŀʹ��*/
/************************************************************************/
void ExcelOperation::InitializeUI()
{
	if (S_OK != CoInitialize(NULL)){
		AfxMessageBox(_T("Initialize com failed..."));
		return;
	}
}

/************************************************************************/
/* �ͷ���Դ                                                               */
/************************************************************************/
void ExcelOperation::UnInitializeUI()
{
	CoUninitialize();
}

/************************************************************************/
/* ����                                                                 */
/************************************************************************/
void ExcelOperation::saveExcel()
{
	ExcelApp.put_DisplayAlerts(FALSE);
	//book.Close(vtMissing, vtMissing, vtMissing);
	book.Save();
}

/************************************************************************/
/* �ر�*/
/************************************************************************/
void ExcelOperation::Close()
{
	books.Close();
}
