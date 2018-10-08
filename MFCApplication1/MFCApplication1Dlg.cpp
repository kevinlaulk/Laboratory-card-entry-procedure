
// MFCApplication1Dlg.cpp : ʵ���ļ�
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
// ����Ӧ�ó��򡰹��ڡ��˵���� CAboutDlg �Ի���

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// �Ի�������
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

// ʵ��
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


// CMFCApplication1Dlg �Ի���



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


// CMFCApplication1Dlg ��Ϣ�������

BOOL CMFCApplication1Dlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// ��������...���˵�����ӵ�ϵͳ�˵��С�

	// IDM_ABOUTBOX ������ϵͳ���Χ�ڡ�
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

	// ���ô˶Ի����ͼ�ꡣ  ��Ӧ�ó��������ڲ��ǶԻ���ʱ����ܽ��Զ�
	//  ִ�д˲���
	SetIcon(m_hIcon, TRUE);			// ���ô�ͼ��
	SetIcon(m_hIcon, FALSE);		// ����Сͼ��

	ShowWindow(SW_MINIMIZE);

	// TODO:  �ڴ���Ӷ���ĳ�ʼ������
	//���
	m_List1.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES); //����ѡ��������
	m_List1.InsertColumn(0, _T(""), LVCFMT_LEFT, 0);
	m_List1.InsertColumn(1, _T("����"), LVCFMT_LEFT, 100); // ����ڶ�������
	m_List1.InsertColumn(2, _T("��������ʱ��"), LVCFMT_LEFT, 100); // �������������
	m_List1.InsertColumn(3, _T("�Ӱ�ʱ��"), LVCFMT_LEFT, 100); // �������������
	m_List1.InsertColumn(4, _T("�ܹ���ʱ��"), LVCFMT_LEFT, 100); // �������������
	m_List1.InsertColumn(5, _T("�ٵ�����"), LVCFMT_LEFT, 100); // �������������
	m_List1.InsertColumn(6, _T("��������"), LVCFMT_LEFT, 100); // �������������

	m_List2.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES); //����ѡ��������
	m_List2.InsertColumn(0, _T(""), LVCFMT_LEFT, 0);
	m_List2.InsertColumn(1, _T("����"), LVCFMT_LEFT, 100); // ����ڶ�������
	m_List2.InsertColumn(2, _T("ԭ��"), LVCFMT_LEFT, 100); // �������������

	m_List3.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES); //����ѡ��������
	m_List3.InsertColumn(0, _T(""), LVCFMT_LEFT, 0);
	m_List3.InsertColumn(1, _T("����"), LVCFMT_LEFT, 100); // ����ڶ�������
	m_List3.InsertColumn(2, _T("ͳ��ʱ��"), LVCFMT_LEFT, 100); // ����ڶ�������
	m_List3.InsertColumn(3, _T("ԭ��"), LVCFMT_LEFT, 100); // ����ڶ�������

	//������
	m_proGress.SetRange(0, 100);
	m_proGress.SetStep(1);
	m_proGress.SetPos(0);

	//��ʼ·��
	char exeFullPath[MAX_PATH], exeFullPath1[MAX_PATH];;

	//GetModuleFileName(NULL, exeFullPath, 320);
	GetCurrentDirectoryA(MAX_PATH, exeFullPath);//��ȡ��ǰ����Ŀ¼
	strcat_s(exeFullPath, "\\..\\���ڱ���.xls");//����Ҫ���ļ�������·��
	XLSPath = CStringW(exeFullPath);
	m_EDIT1.SetWindowText(XLSPath);
	GetCurrentDirectoryA(MAX_PATH, exeFullPath1);//��ȡ��ǰ����Ŀ¼
	strcat_s(exeFullPath1, "\\..\\�γ�ͳ����.xls");//����Ҫ���ļ�������·��
	XLSPath1 = CStringW(exeFullPath1);
	m_EDIT2.SetWindowText(XLSPath1);

	return TRUE;  // ���ǽ��������õ��ؼ������򷵻� TRUE
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

// �����Ի��������С����ť������Ҫ����Ĵ���
//  �����Ƹ�ͼ�ꡣ  ����ʹ���ĵ�/��ͼģ�͵� MFC Ӧ�ó���
//  �⽫�ɿ���Զ���ɡ�

void CMFCApplication1Dlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // ���ڻ��Ƶ��豸������

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// ʹͼ���ڹ����������о���
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// ����ͼ��
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//���û��϶���С������ʱϵͳ���ô˺���ȡ�ù��
//��ʾ��
HCURSOR CMFCApplication1Dlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}



void CMFCApplication1Dlg::OnBnClickedButton1()
{
	/*********************************************
	**********  ��ȡ���ڱ���.xls     
	**********************************************/
	CString str;
	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	// MessageBoxA(0,"Hello, Windows!", "hello", MB_OK);

	

	//����.xls·��
	str.Format(_T("��:%s"), CStringW(XLSPath));
	GetDlgItem(IDC_STATIC3)->SetWindowText(str);//��̬�ı���ʾ
	//strcat(exeFullPath, "\\data.xlsx");//����Ҫ���ļ�������·��
    m_proGress.SetPos(10);//������


	// Open Excel
	ExcelOperation* excel = ExcelOperation::getInstance();
	excel->init();
	excel->setView(FALSE);
	excel->openExcelFile(XLSPath);
	m_proGress.SetPos(15);//������

// 1��open ��sheet 2��
	excel->OpenSheet("���ڻ��ܱ�");
	int nRow = excel->getRowCount();//��ȡsheet������
	int nCol = excel->getColumnCount();//��ȡsheet������
	str.Format(_T("������%d�� ������%d"), nRow, nCol);
	//MessageBox(str);
	GetDlgItem(IDC_STATIC3)->SetWindowText(str);//��̬�ı���ʾ
	m_proGress.SetPos(18);//������

    //Read Excel
	double d_Workhours, d_Addhours, d_Sumhours, array_SumHours[200];
	CString s_Name, s_Workhours, s_Addhours, s_Sumhours, s_fix_latatimes;
	CStringList list_Name;
	m_proGress.SetPos(20);//������

	for (int i = 5; i <= nRow; ++i)
	{
		//��ȡ����
		s_Name = excel->getCellCString(i, 2);
		list_Name.AddTail(s_Name);
		s_Workhours = excel->getCellCString(i, 5);
		s_Addhours = excel->getCellCString(i, 10);
		d_Workhours = _wtof(s_Workhours.GetBuffer());
		d_Addhours = _wtof(s_Addhours.GetBuffer());
		d_Sumhours = d_Workhours + d_Addhours;
		array_SumHours[i - 5] = d_Sumhours;
		//str.Format(_T("d_Workhours��%.1f�� d_Sumhours��%.1f, array_SumHours:%.1f"), d_Workhours, d_Sumhours, array_SumHours[i-5]);
		//MessageBox(str);
		s_Workhours.Format(_T("%.1f"), d_Workhours);
		s_Addhours.Format(_T("%.1f"), d_Addhours);
		s_Sumhours.Format(_T("%.1f"), d_Sumhours);
		//�����б�
		m_List1.InsertItem(i-5, _T(""));
		m_List1.SetItemText(i - 5, 1, s_Name);
		m_List1.SetItemText(i - 5, 2, s_Workhours);
		m_List1.SetItemText(i - 5, 3, s_Addhours);
		m_List1.SetItemText(i - 5, 4, s_Sumhours);
	}
	m_proGress.SetPos(30);//������

//2��open sheet else
	int SheetCount = excel->getSheetCount();
	str.Format(_T("Sheet����%d"), SheetCount);
	//MessageBox(str);
	GetDlgItem(IDC_STATIC3)->SetWindowText(str);//��̬�ı���ʾ

	double d_Latetimes, array_Latetimes[200], array_FixLatetimes[200], colnum[3] = { 10, 26, 42 }, colnum1[3] = { 3, 8, 12 };
	// ���ƣ�ʵ��������ͬʱǩ���������ܳ���200��
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
							// color=16711680 -> means ���󹤡����������������հס�
							// color=255 -> means "�Ӱ�հ�"���ٵ�
							// �Ͻ��� s_starttime.IsEmpty() && color==255
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
	m_proGress.SetPos(40);//������
	GetDlgItem(IDC_STATIC3)->SetWindowText(_T("�˳���ǰEXCEL"));//��̬�ı���ʾ
	


	/*********************************************
	**********  �γ�ͳ����.xls
	**********************************************/

	str.Format(_T("�򿪣�%s"), CStringW(XLSPath1));
	GetDlgItem(IDC_STATIC3)->SetWindowText(str);//��̬�ı���ʾ

	// Open Excel
	excel = ExcelOperation::getInstance();
	excel->init();
	excel->setView(FALSE);
	excel->openExcelFile(XLSPath1);
	excel->OpenSheet("sheet1");
	int nRow2 = excel->getRowCount();//��ȡsheet������
	int nCol2 = excel->getColumnCount();//��ȡsheet������
	str.Format(_T("������%d�� ������%d"), nRow2, nCol2);
	//MessageBox(str);
	GetDlgItem(IDC_STATIC3)->SetWindowText(str);//��̬�ı���ʾ
	m_proGress.SetPos(50);//������

	//Read Excel
	CString s_Name1, s_reason;
	CStringList list_reason;
	for (int i = 2; i <= nRow2; ++i)
	{
		//��ȡ����
		s_Name1 = excel->getCellCString(i, 1);
		list_reason.AddTail(s_Name1);
		s_reason = excel->getCellCString(i, 2);
		list_reason.AddTail(s_reason);
		//�����б�
		
		m_List2.InsertItem(i - 2, _T(""));
		m_List2.SetItemText(i - 2, 1, s_Name1);
		m_List2.SetItemText(i - 2, 2, s_reason);
	}
	excel->Close();
	excel->destroyInstance();
	m_proGress.SetPos(70);//������
	GetDlgItem(IDC_STATIC3)->SetWindowText(_T("�˳�EXCEL"));//��̬�ı���ʾ

	/*********************************************
	**********  ���㲢д��.xls
	**********************************************/
	double ResultSumhours[200];
	CString CellTime, CellPotion, CellName, CellReason;

	GetDlgItem(IDC_STATIC3)->SetWindowText(_T("���㲢д������"));//��̬�ı���ʾ
	excel = ExcelOperation::getInstance();
	excel->init();
	excel->setView(FALSE);
	excel->openExcelFile(XLSPath);
	excel->OpenSheet("ͳ�ƽ��");
	// ����sheet�Ĵ�С��˳����ȫ��ͬ��ֻ�账��"ͳ�ƿγ���"���ɣ� �����sheet��"ͳ�ƿγ���".
	m_proGress.SetPos(80);//������
	GetDlgItem(IDC_STATIC3)->SetWindowText(_T("Sheet�����ڣ����´���"));//��̬�ı���ʾ

	POSITION pos, pos1;
	pos = list_Name.GetHeadPosition();
	//CString StudentName = list_Name.GetHead();
	for (int i = 1; i <= nRow-4; ++i)
	{
		// write time
		ResultSumhours[i - 1] = array_SumHours[i - 1] - (array_Latetimes[i-1]  - array_FixLatetimes[i - 1]) * 3;//point! ������Ф��Τ���
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
	m_proGress.SetPos(95);//������

	MessageBox(_T("Finish"));
	excel->setView(true);
	excel->saveExcel();
	excel->Close();
	excel->destroyInstance();
	m_proGress.SetPos(100);//������
	GetDlgItem(IDC_STATIC3)->SetWindowText(_T("�˳�д����"));//��̬�ı���ʾ

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
//	getcurrentdirectorya(max_path, exefullpath);//��ȡ��ǰ����Ŀ¼
//	char xlsnamepath[max_path] = "\\..\\..\\�ھ���\\";
//	strcat_s(xlsnamepath, xlsname);
//	strcat_s(exefullpath, xlsnamepath);//����Ҫ���ļ�������·��
//	cstring m_exefullpath(exefullpath);
//	messagebox(0, m_exefullpath, _t("�����ļ�"), mb_ok);
//	//strcat(exefullpath, "\\data.xlsx");//����Ҫ���ļ�������·��
//	return exefullpath;
//}

void CMFCApplication1Dlg::OnBnClickedButton3()
{
	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	
	CFileDialog dlg(TRUE, //TRUEΪOPEN�Ի���FALSEΪSAVE AS�Ի���
		NULL,
		NULL,
		OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
		(LPCTSTR)_TEXT("JPG Files (*.xls)|*.xls|All Files (*.*)|*.*||"),
		NULL);
	if (dlg.DoModal() == IDOK)
	{
		XLSPath = dlg.GetPathName(); //�ļ�����������FilePathName��
		m_EDIT1.SetWindowText(XLSPath);
	}
	else
	{
		return;
	}
}


void CMFCApplication1Dlg::OnBnClickedButton4()
{
	// TODO:  �ڴ���ӿؼ�֪ͨ����������

	CFileDialog dlg(TRUE, //TRUEΪOPEN�Ի���FALSEΪSAVE AS�Ի���
		NULL,
		NULL,
		OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
		(LPCTSTR)_TEXT("JPG Files (*.xls)|*.xls|All Files (*.*)|*.*||"),
		NULL);
	if (dlg.DoModal() == IDOK)
	{
		XLSPath1 = dlg.GetPathName(); //�ļ�����������FilePathName��
		m_EDIT2.SetWindowText(XLSPath);
	}
	else
	{
		return;
	}
}


void CMFCApplication1Dlg::OnBnClickedButton5()
{
	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	subDlg *m_subDlg = new subDlg;
	m_subDlg->Create(IDD_DIALOG1, this);
	m_subDlg->ShowWindow(SW_SHOW);
}
