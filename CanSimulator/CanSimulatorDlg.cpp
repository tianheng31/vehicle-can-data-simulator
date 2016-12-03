
// CanSimulatorDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "CanSimulator.h"
#include "CanSimulatorDlg.h"
#include "afxdialogex.h"
#include "resource.h"

#include "CApplication.h"
#include "CRange.h"
#include "CWorkbook.h"
#include "CWorkbooks.h"
#include "CWorksheet.h"
#include "CWorksheets.h"


#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CCanSimulatorDlg 对话框

CApplication app;
CWorkbook book;
CWorkbooks books;
CWorksheet sheet;
CWorksheets sheets;
CRange range;
CFont font;
CRange cols;
LPDISPATCH lpDisp;

CCanSimulatorDlg::CCanSimulatorDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(IDD_CANSIMULATOR_DIALOG, pParent)
{
	EnableActiveAccessibility();
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CCanSimulatorDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_COMBO1, m_SelectProtocols);
	DDX_Control(pDX, IDC_EDIT1, m_TransferInterval);
	DDX_Control(pDX, IDC_EDIT2, m_InputDeviceId);
	DDX_Control(pDX, IDC_LIST1, m_AddedList);
	DDX_Control(pDX, IDC_DATETIMEPICKER1, m_SystemTime);
}

BEGIN_MESSAGE_MAP(CCanSimulatorDlg, CDialogEx)
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BUTTON1, &CCanSimulatorDlg::OnBnClickedAddDeviceId)
	ON_BN_CLICKED(IDC_BUTTON2, &CCanSimulatorDlg::OnBnClickedRemoveDeviceId)
	ON_BN_CLICKED(IDC_BUTTON4, &CCanSimulatorDlg::OnBnClickedExportDeviceIdTemplet)
	ON_BN_CLICKED(IDC_BUTTON3, &CCanSimulatorDlg::OnBnClickedImportDeviceId)
END_MESSAGE_MAP()


// CCanSimulatorDlg 消息处理程序

BOOL CCanSimulatorDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 设置此对话框的图标。  当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	ShowWindow(SW_DENORMAL);

	// TODO: 在此添加额外的初始化代码
	// 向“选择协议”组合框中添加数据
	m_SelectProtocols.AddString(_T("v0"));
	m_SelectProtocols.AddString(_T("v1"));

	// 设置“传输间隔”编辑框的默认值
	m_TransferInterval.SetWindowTextW(_T("1"));

	// 初始化“车载设备号”编辑框
	m_InputDeviceId.SetWindowTextW(_T(""));

	// 初始化“系统时间”控件
	CTime tm;
	tm = CTime::GetCurrentTime();
	m_SystemTime.SetTime(&tm);
	m_SystemTime.SetFormat(_T("yyyy-MM-dd HH:mm:ss"));

	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。  对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CCanSimulatorDlg::OnPaint()
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
HCURSOR CCanSimulatorDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}



void CCanSimulatorDlg::OnBnClickedAddDeviceId()
{
	// TODO: 在此添加控件通知处理程序代码
	CString deviceid;
	m_InputDeviceId.GetWindowTextW(deviceid);

	if (deviceid == "")
		return;

	m_AddedList.AddString(deviceid);

	// 重置编辑框
	m_InputDeviceId.SetWindowTextW(_T(""));
}


void CCanSimulatorDlg::OnBnClickedRemoveDeviceId()
{
	// TODO: 在此添加控件通知处理程序代码
	int index = m_AddedList.GetCurSel();
	if (index < 0) {
		MessageBox(_T("请选择要移除的设备号！"));
		return;
	}

	m_AddedList.DeleteString(index);
	if (index > 0) {
		m_AddedList.SetCurSel(index - 1);
	}
}


void CCanSimulatorDlg::OnBnClickedExportDeviceIdTemplet()
{
	// TODO: 在此添加控件通知处理程序代码
	// 导出Excel
	CFileDialog dlg(FALSE, _T("(*.xls)"), NULL, OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT, _T("(*.xls)|*.xls||"), NULL);
	if (dlg.DoModal() == IDOK)
	{
		CString strFileName = dlg.GetPathName();
		COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
		if (!app.CreateDispatch(_T("Excel.Application")))
		{
			this->MessageBox(_T("创建Excel服务失败！"));
			return;
		}

		books = app.get_Workbooks();
		book = books.Add(covOptional);
		sheets = book.get_Worksheets(); // 得到Worksheets
		sheet = sheets.get_Item(COleVariant((short)1));
		range = sheet.get_Range(COleVariant(_T("A1")), COleVariant(_T("A1")));
		range.put_Value2(COleVariant(_T("车载设备号")));
		cols = range.get_EntireColumn();
		cols.put_NumberFormat(COleVariant(L"@")); // 将整列设置为文本格式
		cols.AutoFit();
		range = sheet.get_Range(COleVariant(_T("B1")), COleVariant(_T("B1")));
		range.put_Value2(COleVariant(_T("本次导入设备号数量")));

		book.SaveCopyAs(COleVariant(strFileName));
		book.put_Saved(true);

		// 释放对象
		book.ReleaseDispatch();
		books.ReleaseDispatch();
		app.Quit();
		app.ReleaseDispatch();
	}

}


void CCanSimulatorDlg::OnBnClickedImportDeviceId()
{
	// TODO: 在此添加控件通知处理程序代码
	HRESULT hr;
	hr = CoInitialize(NULL);
	if (FAILED(hr)) {
		AfxMessageBox(_T("Failed to call CoInitialize()."));
	}

	CFileDialog  filedlg(TRUE, L"*.xls", NULL, OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT, L"xls文件 (*.xls)|*.xls");
	filedlg.m_ofn.lpstrTitle = L"打开文件";
	CString strFilePath;
	if (IDOK == filedlg.DoModal())
	{
		strFilePath = filedlg.GetPathName();
	}
	else
	{
		return;
	}

	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	if (!app.CreateDispatch(_T("Excel.Application")))
	{
		this->MessageBox(_T("创建Excel服务失败！"));
		return;
	}

	books = app.get_Workbooks();
	lpDisp = books.Open(strFilePath, covOptional
		, covOptional, covOptional, covOptional, covOptional, covOptional, covOptional
		, covOptional, covOptional, covOptional
		, covOptional, covOptional, covOptional
		, covOptional);
	book.AttachDispatch(lpDisp);
	sheets = book.get_Worksheets();
	sheet = sheets.get_Item(COleVariant((short)1));

	// 读取Excel表格中若干单元格的值，并在m_AddedList中显示
	COleVariant vResult;
	// 读取已经使用区域的信息，包括已经使用的行数、列数、起始行、起始列
	range.AttachDispatch(sheet.get_UsedRange());
	range.AttachDispatch(range.get_Rows());
	// 获得已经使用的行数
	long iRowNum = range.get_Count();
	range.AttachDispatch(range.get_Columns());
	// 获得已经使用的列数
	long iColNum = range.get_Count();
	// 获得已使用区域的起始行，从1开始
	long iStartRow = range.get_Row();
	// 获得已使用区域的起始列，从1开始
	long iStartCol = range.get_Column();
	// 获得本次导入设备号数量
	range.AttachDispatch(sheet.get_Cells());
	range.AttachDispatch(range.get_Item(COleVariant((long)2), COleVariant((long)2)).pdispVal);
	vResult = range.get_Value2();
	vResult.ChangeType(VT_I4);
	int total = vResult.intVal;

	for (int i = 2; i <= total + 1; i++)
	{
		// 读取单元格的值
		range.AttachDispatch(sheet.get_Cells());
		range.AttachDispatch(range.get_Item(COleVariant((long)i), COleVariant((long)1)).pdispVal);
		vResult = range.get_Value2();
		CString str, stry, strm, strd;
		SYSTEMTIME st;
		if (vResult.vt == VT_BSTR) // 字符串
		{
			str = vResult.bstrVal;
			m_AddedList.AddString(str);
		}
		else if (vResult.vt == VT_R8) // 8-byte real
		{
			str.Format(L"%f", vResult.dblVal);
			m_AddedList.AddString(str);
		}
		else if (vResult.vt == VT_DATE) // 时间
		{
			VariantTimeToSystemTime(vResult.date, &st);
			stry.Format(L"%d", st.wYear);
			strm.Format(L"%d", st.wMonth);
			strd.Format(L"%d", st.wDay);
			str = stry + L"-" + strm + L"-" + strd;
			m_AddedList.AddString(str);
		}
		else if (vResult.vt == VT_EMPTY) // 单元格为空
		{
			str = L"";
			m_AddedList.AddString(str);
		}
		else if (vResult.vt == VT_I4) // 4-byte integer
		{
			str.Format(_T("%ld"), (int)vResult.lVal);
			m_AddedList.AddString(str);
		}
	}

	// 释放对象
	range.ReleaseDispatch();
	sheet.ReleaseDispatch();
	sheets.ReleaseDispatch();
	book.ReleaseDispatch();
	books.ReleaseDispatch();
	app.Quit();
	app.ReleaseDispatch();
}
