// subDlg.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "MFCApplication1.h"
#include "subDlg.h"
#include "afxdialogex.h"


// subDlg �Ի���

IMPLEMENT_DYNAMIC(subDlg, CDialogEx)

subDlg::subDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(subDlg::IDD, pParent)
{

}

subDlg::~subDlg()
{
}

void subDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_STATIC11, Help_Static);
}


BEGIN_MESSAGE_MAP(subDlg, CDialogEx)
END_MESSAGE_MAP()


// subDlg ��Ϣ�������


BOOL subDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// TODO:  �ڴ���Ӷ���ĳ�ʼ��
	CString str;
	str.Format(_T("1���� \"���ڱ���.xls\" �� \"�γ�ͳ����.xls\" ��������λ�û�ó����Ŀ¼�£��Ƽ�)\r\n\
2����� \"���ڱ���.xls �� �γ�ͳ����.xls\" λ�ڷǳ����Ŀ¼�£��ֱ�ѡ���Ӧ�������·��\r\n\
3����� \"���ڱ���.xls �� �γ�ͳ����.xls\" λ�ڸ�Ŀ¼�£����ֱ�����в���Ҫ�޸�·��\r\n\
4������ \"��ʼ\" ���Զ���ȡĿ��·���µ�������񲢼��㣬�����������Ի�����Զ��� \"���ڱ���.xls\"\n\r\
5������Ὣ�� ���ڱ���.xls ���½�sheet����д������� \r\n\
6�����飺�� �γ�ͳ����.xls ���ڸó����Ŀ¼�£��� ���ڱ���.xls ����ĳ�ļ����£����� ��x��/���ڱ���.xls��\
   ��ֱ���޸����� ������ �ھ���.xls���Ա�����"));

	CFont *m_Font;
	m_Font = new CFont;
	m_Font->CreateFont(18, 7, 0, 0, 15,
		FALSE, FALSE, 0, ANSI_CHARSET, OUT_DEFAULT_PRECIS,
		CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, FF_SWISS, _T("Arial"));
	CEdit *m_Edit = (CEdit *)GetDlgItem(IDC_STATIC11);
	m_Edit->SetFont(m_Font, FALSE);
	GetDlgItem(IDC_STATIC11)->SetFont(m_Font);
	GetDlgItem(IDC_STATIC11)->SetWindowText(str);//��̬�ı���ʾ

	return TRUE;  // return TRUE unless you set the focus to a control
	// �쳣:  OCX ����ҳӦ���� FALSE
}
