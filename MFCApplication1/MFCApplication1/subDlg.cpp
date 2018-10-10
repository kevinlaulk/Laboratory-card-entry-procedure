// subDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "MFCApplication1.h"
#include "subDlg.h"
#include "afxdialogex.h"


// subDlg 对话框

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


// subDlg 消息处理程序


BOOL subDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// TODO:  在此添加额外的初始化
	CString str;
	str.Format(_T("1、将 \"考勤报表.xls\" 和 \"课程统计用.xls\" 放于任意位置或该程序根目录下（推荐)\r\n\
2、如果 \"考勤报表.xls 和 课程统计用.xls\" 位于非程序根目录下，分别选择对应表格所在路径\r\n\
3、如果 \"考勤报表.xls 和 课程统计用.xls\" 位于根目录下，则可直接运行不需要修改路径\r\n\
4、单击 \"开始\" 会自动读取目标路径下的两个表格并计算，待弹出结束对话框后，自动打开 \"考勤报表.xls\"\n\r\
5、程序会将在 考勤报表.xls 中新建sheet，并写入计算结果 \r\n\
6、建议：将 课程统计用.xls 放在该程序根目录下，将 考勤报表.xls 放于某文件夹下（例如 第x周/考勤报表.xls）\
   或直接修改名字 （例如 第九周.xls）以备保存"));

	CFont *m_Font;
	m_Font = new CFont;
	m_Font->CreateFont(18, 7, 0, 0, 15,
		FALSE, FALSE, 0, ANSI_CHARSET, OUT_DEFAULT_PRECIS,
		CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, FF_SWISS, _T("Arial"));
	CEdit *m_Edit = (CEdit *)GetDlgItem(IDC_STATIC11);
	m_Edit->SetFont(m_Font, FALSE);
	GetDlgItem(IDC_STATIC11)->SetFont(m_Font);
	GetDlgItem(IDC_STATIC11)->SetWindowText(str);//静态文本显示

	return TRUE;  // return TRUE unless you set the focus to a control
	// 异常:  OCX 属性页应返回 FALSE
}
