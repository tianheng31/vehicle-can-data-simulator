
// CanSimulatorDlg.h : 头文件
//

#pragma once
#include "afxwin.h"
#include "afxdtctl.h"


// CCanSimulatorDlg 对话框
class CCanSimulatorDlg : public CDialogEx
{
// 构造
public:
	CCanSimulatorDlg(CWnd* pParent = NULL);	// 标准构造函数

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_CANSIMULATOR_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持


// 实现
protected:
	HICON m_hIcon;

	// 生成的消息映射函数
	virtual BOOL OnInitDialog();
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	CComboBox m_SelectProtocols;
	CEdit m_TransferInterval;
	CEdit m_InputDeviceId;
	afx_msg void OnBnClickedAddDeviceId();
	afx_msg void OnBnClickedRemoveDeviceId();
	afx_msg void OnBnClickedExportDeviceIdTemplet();
	CListBox m_AddedList;
	afx_msg void OnBnClickedImportDeviceId();
	CDateTimeCtrl m_SystemTime;
};
