
// MFCApplication1Dlg.cpp : 实现文件
//

#include "stdafx.h"
#include "MFCApplication1.h"
#include "MFCApplication1Dlg.h"
#include "afxdialogex.h"
#include "ExcelOperation.h"
#include "subDlg.h"
#include <string>

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

using namespace std;
// 用于应用程序“关于”菜单项的 CAboutDlg 对话框

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// 对话框数据
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

// 实现
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CMFCApplication1Dlg 对话框



CMFCApplication1Dlg::CMFCApplication1Dlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CMFCApplication1Dlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDI_ICON1);
}

void CMFCApplication1Dlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LIST2, m_List1);
	DDX_Control(pDX, IDC_LIST1, m_List2);
	DDX_Control(pDX, IDC_LIST3, m_List3);
	DDX_Control(pDX, IDC_PROGRESS1, m_proGress);
	DDX_Control(pDX, IDC_STATIC3, m_STATIC);
	DDX_Control(pDX, IDC_EDIT2, m_EDIT1);
	DDX_Control(pDX, IDC_EDIT3, m_EDIT2);
}

BEGIN_MESSAGE_MAP(CMFCApplication1Dlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BUTTON1, &CMFCApplication1Dlg::OnBnClickedButton1)
	ON_BN_CLICKED(IDC_BUTTON3, &CMFCApplication1Dlg::OnBnClickedButton3)
	ON_BN_CLICKED(IDC_BUTTON4, &CMFCApplication1Dlg::OnBnClickedButton4)
	ON_BN_CLICKED(IDC_BUTTON5, &CMFCApplication1Dlg::OnBnClickedButton5)
END_MESSAGE_MAP()


// CMFCApplication1Dlg 消息处理程序

BOOL CMFCApplication1Dlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 将“关于...”菜单项添加到系统菜单中。

	// IDM_ABOUTBOX 必须在系统命令范围内。
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// 设置此对话框的图标。  当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	ShowWindow(SW_MINIMIZE);

	// TODO:  在此添加额外的初始化代码
	//表格
	m_List1.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES); //整行选择、网格线
	m_List1.InsertColumn(0, _T(""), LVCFMT_LEFT, 0);
	m_List1.InsertColumn(1, _T("姓名"), LVCFMT_LEFT, 100); // 插入第二列列名
	m_List1.InsertColumn(2, _T("正常工作时长"), LVCFMT_LEFT, 100); // 插入第三列列名
	m_List1.InsertColumn(3, _T("加班时长"), LVCFMT_LEFT, 100); // 插入第四列列名
	m_List1.InsertColumn(4, _T("总工作时长"), LVCFMT_LEFT, 100); // 插入第四列列名
	m_List1.InsertColumn(5, _T("迟到次数"), LVCFMT_LEFT, 100); // 插入第三列列名
	m_List1.InsertColumn(6, _T("修正次数"), LVCFMT_LEFT, 100); // 插入第四列列名

	m_List2.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES); //整行选择、网格线
	m_List2.InsertColumn(0, _T(""), LVCFMT_LEFT, 0);
	m_List2.InsertColumn(1, _T("姓名"), LVCFMT_LEFT, 100); // 插入第二列列名
	m_List2.InsertColumn(2, _T("原因"), LVCFMT_LEFT, 100); // 插入第三列列名

	m_List3.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES); //整行选择、网格线
	m_List3.InsertColumn(0, _T(""), LVCFMT_LEFT, 0);
	m_List3.InsertColumn(1, _T("姓名"), LVCFMT_LEFT, 100); // 插入第二列列名
	m_List3.InsertColumn(2, _T("统计时长"), LVCFMT_LEFT, 100); // 插入第二列列名
	m_List3.InsertColumn(3, _T("原因"), LVCFMT_LEFT, 100); // 插入第二列列名

	//进度条
	m_proGress.SetRange(0, 100);
	m_proGress.SetStep(1);
	m_proGress.SetPos(0);

	//初始路径
	char exeFullPath[MAX_PATH], exeFullPath1[MAX_PATH];;

	//GetModuleFileName(NULL, exeFullPath, 320);
	GetCurrentDirectoryA(MAX_PATH, exeFullPath);//获取当前工作目录
	strcat_s(exeFullPath, "\\..\\考勤报表.xls");//设置要打开文件的完整路径
	XLSPath = CStringW(exeFullPath);
	m_EDIT1.SetWindowText(XLSPath);
	GetCurrentDirectoryA(MAX_PATH, exeFullPath1);//获取当前工作目录
	strcat_s(exeFullPath1, "\\..\\课程统计用.xls");//设置要打开文件的完整路径
	XLSPath1 = CStringW(exeFullPath1);
	m_EDIT2.SetWindowText(XLSPath1);

	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

void CMFCApplication1Dlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。  对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CMFCApplication1Dlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 用于绘制的设备上下文

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 使图标在工作区矩形中居中
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 绘制图标
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标
//显示。
HCURSOR CMFCApplication1Dlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}



void CMFCApplication1Dlg::OnBnClickedButton1()
{
	/*********************************************
	**********  读取考勤报表.xls     
	**********************************************/
	CString str;
	// TODO:  在此添加控件通知处理程序代码
	// MessageBoxA(0,"Hello, Windows!", "hello", MB_OK);

	

	//设置.xls路径
	str.Format(_T("打开:%s"), CStringW(XLSPath));
	GetDlgItem(IDC_STATIC3)->SetWindowText(str);//静态文本显示
	//strcat(exeFullPath, "\\data.xlsx");//设置要打开文件的完整路径
    m_proGress.SetPos(10);//进度条


	// Open Excel
	ExcelOperation* excel = ExcelOperation::getInstance();
	excel->init();
	excel->setView(FALSE);
	excel->openExcelFile(XLSPath);
	m_proGress.SetPos(15);//进度条

// 1、open “sheet 2”
	excel->OpenSheet("考勤汇总表");
	int nRow = excel->getRowCount();//获取sheet中行数
	int nCol = excel->getColumnCount();//获取sheet中行数
	str.Format(_T("行数：%d， 列数：%d"), nRow, nCol);
	//MessageBox(str);
	GetDlgItem(IDC_STATIC3)->SetWindowText(str);//静态文本显示
	m_proGress.SetPos(18);//进度条

    //Read Excel
	double d_Workhours, d_Addhours, d_Sumhours, array_SumHours[200];
	CString s_Name, s_Workhours, s_Addhours, s_Sumhours, s_fix_latatimes;
	CStringList list_Name;
	m_proGress.SetPos(20);//进度条

	for (int i = 5; i <= nRow; ++i)
	{
		//读取数据
		s_Name = excel->getCellCString(i, 2);
		list_Name.AddTail(s_Name);
		s_Workhours = excel->getCellCString(i, 5);
		s_Addhours = excel->getCellCString(i, 10);
		d_Workhours = _wtof(s_Workhours.GetBuffer());
		d_Addhours = _wtof(s_Addhours.GetBuffer());
		d_Sumhours = d_Workhours + d_Addhours;
		array_SumHours[i - 5] = d_Sumhours;
		//str.Format(_T("d_Workhours：%.1f， d_Sumhours：%.1f, array_SumHours:%.1f"), d_Workhours, d_Sumhours, array_SumHours[i-5]);
		//MessageBox(str);
		s_Workhours.Format(_T("%.1f"), d_Workhours);
		s_Addhours.Format(_T("%.1f"), d_Addhours);
		s_Sumhours.Format(_T("%.1f"), d_Sumhours);
		//插入列表
		m_List1.InsertItem(i-5, _T(""));
		m_List1.SetItemText(i - 5, 1, s_Name);
		m_List1.SetItemText(i - 5, 2, s_Workhours);
		m_List1.SetItemText(i - 5, 3, s_Addhours);
		m_List1.SetItemText(i - 5, 4, s_Sumhours);
	}
	m_proGress.SetPos(30);//进度条

//2、open sheet else
	int SheetCount = excel->getSheetCount();
	str.Format(_T("Sheet数：%d"), SheetCount);
	//MessageBox(str);
	GetDlgItem(IDC_STATIC3)->SetWindowText(str);//静态文本显示

	double d_Latetimes, array_Latetimes[200], array_FixLatetimes[200], colnum[3] = { 10, 26, 42 }, colnum1[3] = { 3, 8, 12 };
	// 限制：实验室人数同时签到人数不能超过200人
	POSITION rPos = list_Name.GetHeadPosition();
	for (int i = 3; i <= SheetCount; ++i)
	{
		excel->OpenSheetwithId((long)i);
		for (int num = 0; num <= 2; ++num)
		{
			int stu_num = (i - 3) * 3 + num;
			if (stu_num<=nRow-1)
			{
				// read late times in everysheet and three students persheet.
				CString s_Latetimes = excel->getCellCString(8, colnum[num]);
				d_Latetimes = _wtof(s_Latetimes.GetBuffer());
				array_Latetimes[stu_num] = d_Latetimes;
				//m_List1.InsertItem(stu_num, _T(""));
				s_Name = list_Name.GetNext(rPos);
				//m_List1.SetItemText(stu_num, 5, s_Name);
				m_List1.SetItemText(stu_num, 5, s_Latetimes);
				//MessageBox(s_Latetimes);
				// read weekend
				int fix_latatimes = 0;
				for (int weekendi = 18; weekendi <= 19; ++weekendi)
				{
					for (int weekend_col = 0; weekend_col < 3; ++weekend_col)
					{
						CString s_starttime = excel->getCellCString(weekendi, colnum1[weekend_col] + num * 16);
						int color = excel->getCellcolor(weekendi, colnum1[weekend_col] + num*16);
						/*str.Format(_T("number: %s; color:%d"), s_starttime, color);
						MessageBox(str);*/
						if ((!s_starttime.IsEmpty()) && color == 255)
						{
							// color=16711680 -> means ”矿工“，”正常“，”空白“
							// color=255 -> means "加班空白"，迟到
							// 严谨： s_starttime.IsEmpty() && color==255
							fix_latatimes++;
						}

						
					}
				}
				array_FixLatetimes[stu_num] = fix_latatimes;
				s_fix_latatimes.Format(_T("%d"), fix_latatimes);
				m_List1.SetItemText(stu_num, 6, s_fix_latatimes);
			}
			
		}
		
	}
	excel->Close();
	excel->destroyInstance();
	m_proGress.SetPos(40);//进度条
	GetDlgItem(IDC_STATIC3)->SetWindowText(_T("退出当前EXCEL"));//静态文本显示
	


	/*********************************************
	**********  课程统计用.xls
	**********************************************/

	str.Format(_T("打开：%s"), CStringW(XLSPath1));
	GetDlgItem(IDC_STATIC3)->SetWindowText(str);//静态文本显示

	// Open Excel
	excel = ExcelOperation::getInstance();
	excel->init();
	excel->setView(FALSE);
	excel->openExcelFile(XLSPath1);
	excel->OpenSheet("sheet1");
	int nRow2 = excel->getRowCount();//获取sheet中行数
	int nCol2 = excel->getColumnCount();//获取sheet中行数
	str.Format(_T("行数：%d， 列数：%d"), nRow2, nCol2);
	//MessageBox(str);
	GetDlgItem(IDC_STATIC3)->SetWindowText(str);//静态文本显示
	m_proGress.SetPos(50);//进度条

	//Read Excel
	CString s_Name1, s_reason;
	CStringList list_reason;
	for (int i = 2; i <= nRow2; ++i)
	{
		//读取数据
		s_Name1 = excel->getCellCString(i, 1);
		list_reason.AddTail(s_Name1);
		s_reason = excel->getCellCString(i, 2);
		list_reason.AddTail(s_reason);
		//插入列表
		
		m_List2.InsertItem(i - 2, _T(""));
		m_List2.SetItemText(i - 2, 1, s_Name1);
		m_List2.SetItemText(i - 2, 2, s_reason);
	}
	excel->Close();
	excel->destroyInstance();
	m_proGress.SetPos(70);//进度条
	GetDlgItem(IDC_STATIC3)->SetWindowText(_T("退出EXCEL"));//静态文本显示

	/*********************************************
	**********  计算并写入.xls
	**********************************************/
	double ResultSumhours[200];
	CString CellTime, CellPotion, CellName, CellReason;

	GetDlgItem(IDC_STATIC3)->SetWindowText(_T("计算并写入数据"));//静态文本显示
	excel = ExcelOperation::getInstance();
	excel->init();
	excel->setView(FALSE);
	excel->openExcelFile(XLSPath);
	excel->OpenSheet("统计结果");
	// 两个sheet的大小和顺序完全相同，只需处理"统计课程用"即可， 从这个sheet找"统计课程用".
	m_proGress.SetPos(80);//进度条
	GetDlgItem(IDC_STATIC3)->SetWindowText(_T("Sheet不存在，重新创建"));//静态文本显示

	POSITION pos, pos1;
	pos = list_Name.GetHeadPosition();
	//CString StudentName = list_Name.GetHead();
	for (int i = 1; i <= nRow-4; ++i)
	{
		// write time
		ResultSumhours[i - 1] = array_SumHours[i - 1] - (array_Latetimes[i-1]  - array_FixLatetimes[i - 1]) * 3;//point! そこわばぐのこぉ
		CellPotion.Format(_T("B%d"),i);
		CellTime.Format(_T("%.1f"), ResultSumhours[i-1]);
		//MessageBox(CellPotion);
		//s_Name1.Format(_T("ResultSumhours:%.1f;array_SumHours:%.1f;array_Latetimes:%.1f"), ResultSumhours[i - 5], array_SumHours[i - 5], array_Latetimes[i - 5]);
		//MessageBox(s_Name1);
		//MessageBox(CellTime);
		excel->setCellCStringValue(CellPotion, CellTime);
		m_List3.InsertItem(i - 1, _T(""));
		m_List3.SetItemText(i - 1, 2, CellTime);

		// write name
		CellPotion.Format(_T("A%d"), i);
		CellName = list_Name.GetNext(pos);
		excel->setCellCStringValue(CellPotion, CellName);
		m_List3.SetItemText(i - 1, 1, CellName);

		// write reason
		
		pos1 = list_reason.Find(CellName);
		if (pos1 != NULL)
		{
			CellPotion.Format(_T("C%d"), i);
			//MessageBox(list_reason.GetAt(pos1));
			CellReason = list_reason.GetNext(pos1);
			CellReason = list_reason.GetAt(pos1);
			//MessageBox(CellReason);
			excel->setCellCStringValue(CellPotion, CellReason);
			m_List3.SetItemText(i - 1, 3, CellReason);
		}
		
	}
	m_proGress.SetPos(95);//进度条

	MessageBox(_T("Finish"));
	excel->setView(true);
	excel->saveExcel();
	excel->Close();
	excel->destroyInstance();
	m_proGress.SetPos(100);//进度条
	GetDlgItem(IDC_STATIC3)->SetWindowText(_T("退出写入表格"));//静态文本显示

	/*excel->setCellFormat("A1", "@");
	excel->setCellValue("A1", "3342");
	excel->setCellFormat("A2", "0.00");
	excel->setCellValue("A2", "3342.5684");
	excel->setView(true);
	excel->saveExcel();
	excel->saveExcelAs("C:\\saveTest.xlsx");
	excel->Close();
	~ExcelOperation()*/


}

//char* setxlspath(const char* xlsname)
//{
//	char exefullpath[max_path];
//	//getmodulefilename(null, exefullpath, 320);
//	getcurrentdirectorya(max_path, exefullpath);//获取当前工作目录
//	char xlsnamepath[max_path] = "\\..\\..\\第九周\\";
//	strcat_s(xlsnamepath, xlsname);
//	strcat_s(exefullpath, xlsnamepath);//设置要打开文件的完整路径
//	cstring m_exefullpath(exefullpath);
//	messagebox(0, m_exefullpath, _t("加载文件"), mb_ok);
//	//strcat(exefullpath, "\\data.xlsx");//设置要打开文件的完整路径
//	return exefullpath;
//}

void CMFCApplication1Dlg::OnBnClickedButton3()
{
	// TODO:  在此添加控件通知处理程序代码
	
	CFileDialog dlg(TRUE, //TRUE为OPEN对话框，FALSE为SAVE AS对话框
		NULL,
		NULL,
		OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
		(LPCTSTR)_TEXT("JPG Files (*.xls)|*.xls|All Files (*.*)|*.*||"),
		NULL);
	if (dlg.DoModal() == IDOK)
	{
		XLSPath = dlg.GetPathName(); //文件名保存在了FilePathName里
		m_EDIT1.SetWindowText(XLSPath);
	}
	else
	{
		return;
	}
}


void CMFCApplication1Dlg::OnBnClickedButton4()
{
	// TODO:  在此添加控件通知处理程序代码

	CFileDialog dlg(TRUE, //TRUE为OPEN对话框，FALSE为SAVE AS对话框
		NULL,
		NULL,
		OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
		(LPCTSTR)_TEXT("JPG Files (*.xls)|*.xls|All Files (*.*)|*.*||"),
		NULL);
	if (dlg.DoModal() == IDOK)
	{
		XLSPath1 = dlg.GetPathName(); //文件名保存在了FilePathName里
		m_EDIT2.SetWindowText(XLSPath);
	}
	else
	{
		return;
	}
}


void CMFCApplication1Dlg::OnBnClickedButton5()
{
	// TODO:  在此添加控件通知处理程序代码
	subDlg *m_subDlg = new subDlg;
	m_subDlg->Create(IDD_DIALOG1, this);
	m_subDlg->ShowWindow(SW_SHOW);
}
