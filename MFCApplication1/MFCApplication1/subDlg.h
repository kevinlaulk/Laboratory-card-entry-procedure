#pragma once
#include "afxwin.h"


// subDlg �Ի���

class subDlg : public CDialogEx
{
	DECLARE_DYNAMIC(subDlg)

public:
	subDlg(CWnd* pParent = NULL);   // ��׼���캯��
	virtual ~subDlg();

// �Ի�������
	enum { IDD = IDD_DIALOG1 };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

	DECLARE_MESSAGE_MAP()
public:
	CStatic Help_Static;
	virtual BOOL OnInitDialog();
};
