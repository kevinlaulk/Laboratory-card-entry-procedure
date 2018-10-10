#pragma once
#include "afxwin.h"


// subDlg 对话框

class subDlg : public CDialogEx
{
	DECLARE_DYNAMIC(subDlg)

public:
	subDlg(CWnd* pParent = NULL);   // 标准构造函数
	virtual ~subDlg();

// 对话框数据
	enum { IDD = IDD_DIALOG1 };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

	DECLARE_MESSAGE_MAP()
public:
	CStatic Help_Static;
	virtual BOOL OnInitDialog();
};
