
// MFCApplication1Dlg.h : ͷ�ļ�
//

#pragma once
#include "afxcmn.h"
#include "afxwin.h"


// CMFCApplication1Dlg �Ի���
class CMFCApplication1Dlg : public CDialogEx
{
// ����
public:
	CMFCApplication1Dlg(CWnd* pParent = NULL);	// ��׼���캯��

// �Ի�������
	enum { IDD = IDD_MFCAPPLICATION1_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV ֧��


// ʵ��
protected:
	HICON m_hIcon;

	// ���ɵ���Ϣӳ�亯��
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedButton1();
	CListCtrl m_List1;
	CListCtrl m_List2;
	CListCtrl m_List3;
	afx_msg void OnLvnItemchangedList1(NMHDR *pNMHDR, LRESULT *pResult);
	CProgressCtrl m_proGress;
	CStatic m_STATIC;
	CEdit m_EDIT1;
	CEdit m_EDIT2;

	CString XLSPath; //����.xls·
	CString XLSPath1;//����.xls·��
	
	afx_msg void OnBnClickedButton3();
	afx_msg void OnBnClickedButton4();
	afx_msg void OnBnClickedButton5();
};
