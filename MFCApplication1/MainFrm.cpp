// この MFC サンプル ソース コードでは、MFC Microsoft Office Fluent ユーザー インターフェイス
// ("Fluent UI") の使用方法を示します。このコードは、MFC C++ ライブラリ ソフトウェアに
// 同梱されている Microsoft Foundation Class リファレンスおよび関連電子ドキュメントを
// 補完するための参考資料として提供されます。
// Fluent UI を複製、使用、または配布するためのライセンス条項は個別に用意されています。
// Fluent UI ライセンス プログラムの詳細については、Web サイト
// https://go.microsoft.com/fwlink/?LinkId=238214.
//
// Copyright (C) Microsoft Corporation
// All rights reserved.

// MainFrm.cpp : CMainFrame クラスの実装
//

#include "pch.h"
#include "framework.h"
#include "MFCApplication1.h"

#include "MainFrm.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif
#include "MFCApplication1Set.h"
#include "MFCApplication1Doc.h"
#include "MFCApplication1View.h"

// CMainFrame

IMPLEMENT_DYNCREATE(CMainFrame, CFrameWndEx)

BEGIN_MESSAGE_MAP(CMainFrame, CFrameWndEx)
	ON_WM_CREATE()
	ON_COMMAND_RANGE(ID_VIEW_APPLOOK_WIN_2000, ID_VIEW_APPLOOK_WINDOWS_7, &CMainFrame::OnApplicationLook)
	ON_UPDATE_COMMAND_UI_RANGE(ID_VIEW_APPLOOK_WIN_2000, ID_VIEW_APPLOOK_WINDOWS_7, &CMainFrame::OnUpdateApplicationLook)
	ON_COMMAND(ID_FILE_PRINT, &CMainFrame::OnFilePrint)
	ON_COMMAND(ID_FILE_PRINT_DIRECT, &CMainFrame::OnFilePrint)
	ON_COMMAND(ID_FILE_PRINT_PREVIEW, &CMainFrame::OnFilePrintPreview)
	ON_UPDATE_COMMAND_UI(ID_FILE_PRINT_PREVIEW, &CMainFrame::OnUpdateFilePrintPreview)
	ON_UPDATE_COMMAND_UI(ID_FILE_OPEN, &CMainFrame::OnUpdateFileOpen)
	ON_UPDATE_COMMAND_UI(ID_FILE_NEW, &CMainFrame::OnUpdateFileNew)
	ON_UPDATE_COMMAND_UI_RANGE(ID_FILE_CLOSE, ID_FILE_PAGE_SETUP, &CMainFrame::OnUpdateFileNew)
	ON_UPDATE_COMMAND_UI_RANGE(ID_FILE_PRINT, ID_FILE_SEND_MAIL, &CMainFrame::OnUpdateFileNew)
	ON_UPDATE_COMMAND_UI_RANGE(ID_EDIT_CLEAR, ID_EDIT_CLEAR_ALL, &CMainFrame::OnUpdateFileNew)
	ON_UPDATE_COMMAND_UI_RANGE(ID_EDIT_FIND, ID_EDIT_FIND, &CMainFrame::OnUpdateFileNew)
	ON_UPDATE_COMMAND_UI_RANGE(ID_EDIT_REPEAT, ID_EDIT_REPLACE, &CMainFrame::OnUpdateFileNew)
	ON_UPDATE_COMMAND_UI_RANGE(ID_EDIT_UNDO, ID_EDIT_REDO, &CMainFrame::OnUpdateFileNew)
	ON_UPDATE_COMMAND_UI_RANGE(ID_OLE_INSERT_NEW, ID_OLE_VERB_LAST, &CMainFrame::OnUpdateFileNew)
	ON_UPDATE_COMMAND_UI(ID_INDICATOR_ROW, &CMainFrame::OnUpdateIndicatorRow)
	ON_UPDATE_COMMAND_UI(ID_INDICATOR_COL, &CMainFrame::OnUpdateIndicatorCol)
	ON_UPDATE_COMMAND_UI(ID_INDICATOR_MODIFY, &CMainFrame::OnUpdateIndicatorModify)
	ON_UPDATE_COMMAND_UI(ID_INDICATOR_DATE, &CMainFrame::OnUpdateIndicatorDate)
	ON_UPDATE_COMMAND_UI(ID_INDICATOR_TIME, &CMainFrame::OnUpdateIndicatorTime)
	ON_WM_TIMER()
	ON_WM_CLOSE()
	ON_MESSAGE(WM_SETMESSAGESTRING, &CMainFrame::OnSetmessagestring)
	ON_WM_GETMINMAXINFO()
END_MESSAGE_MAP()

static UINT indicators[] =
{
	ID_SEPARATOR,           // ステータス ライン インジケーター
	ID_INDICATOR_ROW,
	ID_INDICATOR_COL,
	ID_INDICATOR_CAPS,
	ID_INDICATOR_NUM,
	ID_INDICATOR_SCRL,
	ID_INDICATOR_MODIFY,
	ID_INDICATOR_DATE,
	ID_INDICATOR_TIME,
};

// CMainFrame コンストラクション/デストラクション

CMainFrame::CMainFrame() noexcept
{
	// TODO: メンバー初期化コードをここに追加してください。
	theApp.m_nAppLook = theApp.GetInt(_T("ApplicationLook"), ID_VIEW_APPLOOK_OFF_2007_BLUE);
	::SecureZeroMemory(&m_systemTime, sizeof(SYSTEMTIME));
	UINT
		source[] = {106,103,102,101,104},
		value = sizeof(source) / sizeof(UINT);
	BOOL result = m_imgSmall.Create(16, 15, ILC_COLOR24 | ILC_MASK, value, 0);
	if (result)
	{
		result = m_imgLarge.Create(32, 32, ILC_COLOR24 | ILC_MASK, value, 0);
		if (result)
		{
			for (UINT index = 0; result && index < value; index++)
			{
				result = FALSE;
				HICON value = ::LoadIcon(nullptr, MAKEINTRESOURCE(source[index]));
				if (nullptr != value && INVALID_HANDLE_VALUE != value)
				{
					result = 0 <= m_imgSmall.Add(value);
					if (result)
					{
						result = 0 <= m_imgLarge.Add(value);
					}
				}
			};
		}
	}
}

CMainFrame::~CMainFrame()
{
}

int CMainFrame::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
	if (CFrameWndEx::OnCreate(lpCreateStruct) == -1)
		return -1;

	m_wndRibbonBar.Create(this);
	m_wndRibbonBar.LoadFromResource(IDR_RIBBON);

	if (!m_wndStatusBar.Create(this))
	{
		TRACE0("ステータス バーの作成に失敗しました。\n");
		return -1;      // 作成できない場合
	}

	UINT count = sizeof(indicators) / sizeof(UINT);
	for (UINT index = 0; index < count; index++)
	{
		UINT nID = indicators[index];
		CMFCRibbonStatusBarPane* pane = nullptr;
		CString value;
		switch (nID)
		{
		case ID_SEPARATOR:
			if (value.LoadString(AFX_IDS_IDLEMESSAGE))
			{
				pane = new CMFCRibbonStatusBarPane(nID, value, TRUE);
				if (nullptr != pane)
				{
					m_wndStatusBar.AddElement(pane, value);
				}
			}
			break;
		default:
			if (value.LoadString(nID))
			{
				pane = new CMFCRibbonStatusBarPane(nID, value, TRUE);
				if (nullptr != pane)
				{
					m_wndStatusBar.AddExtendedElement(pane, value);
				}
			}
			break;
		}
	};

	// Visual Studio 2005 スタイルのドッキング ウィンドウ動作を有効にします
	CDockingManager::SetDockingMode(DT_SMART);
	// Visual Studio 2005 スタイルのドッキング ウィンドウの自動非表示動作を有効にします
	EnableAutoHidePanes(CBRS_ALIGN_ANY);
	// 固定値に基づいてビジュアル マネージャーと visual スタイルを設定します
	OnApplicationLook(theApp.m_nAppLook);

	//SetWindowPos(nullptr, 0, 0, 838, 720, SWP_NOMOVE | SWP_NOZORDER);// ←効いてない
	MessageBox(AFX_IDS_IDLEMESSAGE, MB_ICONINFORMATION);
	SetTimer(ID_TIMER_INTERVAL, 128, nullptr);
	return 0;
}

BOOL CMainFrame::PreCreateWindow(CREATESTRUCT& cs)
{
	if( !CFrameWndEx::PreCreateWindow(cs) )
		return FALSE;
	// TODO: この位置で CREATESTRUCT cs を修正して Window クラスまたはスタイルを
	//  修正してください。

	return TRUE;
}

// CMainFrame の診断

#ifdef _DEBUG
void CMainFrame::AssertValid() const
{
	CFrameWndEx::AssertValid();
}

void CMainFrame::Dump(CDumpContext& dc) const
{
	CFrameWndEx::Dump(dc);
}
#endif //_DEBUG


// CMainFrame メッセージ ハンドラー

void CMainFrame::OnApplicationLook(UINT id)
{
	CWaitCursor wait;

	theApp.m_nAppLook = id;

	switch (theApp.m_nAppLook)
	{
	case ID_VIEW_APPLOOK_WIN_2000:
		CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CMFCVisualManager));
		m_wndRibbonBar.SetWindows7Look(FALSE);
		break;

	case ID_VIEW_APPLOOK_OFF_XP:
		CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CMFCVisualManagerOfficeXP));
		m_wndRibbonBar.SetWindows7Look(FALSE);
		break;

	case ID_VIEW_APPLOOK_WIN_XP:
		CMFCVisualManagerWindows::m_b3DTabsXPTheme = TRUE;
		CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CMFCVisualManagerWindows));
		m_wndRibbonBar.SetWindows7Look(FALSE);
		break;

	case ID_VIEW_APPLOOK_OFF_2003:
		CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CMFCVisualManagerOffice2003));
		CDockingManager::SetDockingMode(DT_SMART);
		m_wndRibbonBar.SetWindows7Look(FALSE);
		break;

	case ID_VIEW_APPLOOK_VS_2005:
		CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CMFCVisualManagerVS2005));
		CDockingManager::SetDockingMode(DT_SMART);
		m_wndRibbonBar.SetWindows7Look(FALSE);
		break;

	case ID_VIEW_APPLOOK_VS_2008:
		CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CMFCVisualManagerVS2008));
		CDockingManager::SetDockingMode(DT_SMART);
		m_wndRibbonBar.SetWindows7Look(FALSE);
		break;

	case ID_VIEW_APPLOOK_WINDOWS_7:
		CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CMFCVisualManagerWindows7));
		CDockingManager::SetDockingMode(DT_SMART);
		m_wndRibbonBar.SetWindows7Look(TRUE);
		break;

	default:
		switch (theApp.m_nAppLook)
		{
		case ID_VIEW_APPLOOK_OFF_2007_BLUE:
			CMFCVisualManagerOffice2007::SetStyle(CMFCVisualManagerOffice2007::Office2007_LunaBlue);
			break;

		case ID_VIEW_APPLOOK_OFF_2007_BLACK:
			CMFCVisualManagerOffice2007::SetStyle(CMFCVisualManagerOffice2007::Office2007_ObsidianBlack);
			break;

		case ID_VIEW_APPLOOK_OFF_2007_SILVER:
			CMFCVisualManagerOffice2007::SetStyle(CMFCVisualManagerOffice2007::Office2007_Silver);
			break;

		case ID_VIEW_APPLOOK_OFF_2007_AQUA:
			CMFCVisualManagerOffice2007::SetStyle(CMFCVisualManagerOffice2007::Office2007_Aqua);
			break;
		}

		CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CMFCVisualManagerOffice2007));
		CDockingManager::SetDockingMode(DT_SMART);
		m_wndRibbonBar.SetWindows7Look(FALSE);
	}

	RedrawWindow(nullptr, nullptr, RDW_ALLCHILDREN | RDW_INVALIDATE | RDW_UPDATENOW | RDW_FRAME | RDW_ERASE);

	theApp.WriteInt(_T("ApplicationLook"), theApp.m_nAppLook);
}

void CMainFrame::OnUpdateApplicationLook(CCmdUI* pCmdUI)
{
	pCmdUI->SetRadio(theApp.m_nAppLook == pCmdUI->m_nID);
}


void CMainFrame::OnFilePrint()
{
	if (IsPrintPreview())
	{
		PostMessage(WM_COMMAND, AFX_ID_PREVIEW_PRINT);
	}
}

void CMainFrame::OnFilePrintPreview()
{
	if (IsPrintPreview())
	{
		PostMessage(WM_COMMAND, AFX_ID_PREVIEW_CLOSE);  // 印刷プレビュー モードを強制的に終了する
	}
}

void CMainFrame::OnUpdateFilePrintPreview(CCmdUI* pCmdUI)
{
	pCmdUI->SetCheck(IsPrintPreview());
}



void CMainFrame::OnUpdateFileOpen(CCmdUI* pCmdUI)
{
	if (nullptr != pCmdUI)
	{
		BOOL result = FALSE;
		CDocument* pDoc = GetActiveDocument();
		if (nullptr != pDoc)
		{
			result = pDoc->GetPathName().IsEmpty();
		}
		pCmdUI->Enable(result);
	}
}


void CMainFrame::OnUpdateFileNew(CCmdUI* pCmdUI)
{
	if (nullptr != pCmdUI)
	{
		BOOL result = FALSE;
		CDocument* pDoc = GetActiveDocument();
		if (nullptr != pDoc)
		{
			result = !pDoc->GetPathName().IsEmpty();
		}
		pCmdUI->Enable(result);
	}
}


void CMainFrame::OnUpdateIndicatorRow(CCmdUI* pCmdUI)
{
	if (nullptr != pCmdUI)
	{
		BOOL result = FALSE;
		int index = 0;
		CDocument* pDoc = GetActiveDocument();
		if (nullptr != pDoc)
		{
			result = !pDoc->GetPathName().IsEmpty();
			if (result)
			{
				CMFCApplication1View* pView =
					(CMFCApplication1View*)GetActiveView();
				if (nullptr != pView)
				{
					index = 1 + pView->GetRow();
				}
			}
		}
		CString value;
		value.Format(ID_INDICATOR_ROW, index);
		pCmdUI->SetText(value);
		pCmdUI->Enable(result);
	}
}


void CMainFrame::OnUpdateIndicatorCol(CCmdUI* pCmdUI)
{
	BOOL result = FALSE;
	int index = 0;
	CDocument* pDoc = GetActiveDocument();
	if (nullptr != pDoc)
	{
		result = !pDoc->GetPathName().IsEmpty();
		if (result)
		{
			CMFCApplication1View* pView =
				(CMFCApplication1View*)GetActiveView();
			if (nullptr != pView)
			{
				index = 1 + pView->GetCol();
			}
		}
	}
	CString value;
	value.Format(ID_INDICATOR_COL, index);
	pCmdUI->SetText(value);
	pCmdUI->Enable(result);
}


void CMainFrame::OnUpdateIndicatorModify(CCmdUI* pCmdUI)
{
	if (nullptr != pCmdUI)
	{
		BOOL result = FALSE;
		CDocument* pDoc = GetActiveDocument();
		if (nullptr != pDoc)
		{
			result = pDoc->IsModified();
		}
		pCmdUI->Enable(result);
	}
}


void CMainFrame::OnUpdateIndicatorDate(CCmdUI* pCmdUI)
{
	if (nullptr != pCmdUI)
	{
		CString value;
		value.Format(_T("%04d/%02d/%02d"),
			m_systemTime.wYear,
			m_systemTime.wMonth,
			m_systemTime.wDay);
		pCmdUI->SetText(value);
		pCmdUI->Enable(TRUE);
	}
}


void CMainFrame::OnUpdateIndicatorTime(CCmdUI* pCmdUI)
{
	if (nullptr != pCmdUI)
	{
		CString value;
		value.Format(
			m_systemTime.wMilliseconds < 500 ?
			_T("%02d:%02d") : _T("%02d %02d"),
			m_systemTime.wHour, m_systemTime.wMinute);
		pCmdUI->SetText(value);
		pCmdUI->Enable(TRUE);
	}
}


void CMainFrame::OnTimer(UINT_PTR nIDEvent)
{
	switch (nIDEvent)
	{
	case ID_TIMER_INTERVAL:
		::GetLocalTime(&m_systemTime);
		break;
	}
	CFrameWndEx::OnTimer(nIDEvent);
}


void CMainFrame::OnClose()
{
	BOOL result = TRUE;
	if (!m_pPrintPreviewFrame)
	{
		CDocument* pDoc = GetActiveDocument();
		if (nullptr != pDoc)
		{
			if (pDoc->IsModified())
			{
				result = pDoc->SaveModified();
				if (result)
				{
					pDoc->SetModifiedFlag(FALSE);
				}
			}
			else
			{
				result = IDYES ==
					MessageBox(AFX_IDP_ASK_TO_EXIT,
						MB_YESNO | MB_ICONQUESTION | MB_DEFBUTTON2);
			}
			if (result)
			{
				KillTimer(ID_TIMER_INTERVAL);
			}
		}
	}
	if (result)
	{
		CFrameWndEx::OnClose();
	}
}


UINT CMainFrame::MessageBox(UINT nID, UINT style)
{
	BOOL result = IDOK;
	CString value;
	if (value.LoadString(nID))
	{
		result = MessageBox(value, style);
	}
	return result;
}


UINT CMainFrame::MessageBox(LPCTSTR caption, UINT style)
{
	UINT index = MB_ICONEXCLAMATION | MB_ICONINFORMATION;
	index &= style;
	index >>= 4;
	{
		CMFCRibbonStatusBarPane* source = 
			(CMFCRibbonStatusBarPane*)m_wndStatusBar.FindElement(ID_SEPARATOR);
		if (nullptr != source)
		{
			m_wndStatusBar.RemoveElement(ID_SEPARATOR);
			HICON value = nullptr;
			if ((int)index < m_imgSmall.GetImageCount())
			{
				value = m_imgSmall.ExtractIcon(index);
			}
			if (nullptr == value)
			{
				value = m_imgSmall.ExtractIcon(4);
			}
			if (nullptr != value)
			{
				source = new CMFCRibbonStatusBarPane(ID_SEPARATOR, caption, TRUE, value);
			}
			else
			{
				source = new CMFCRibbonStatusBarPane(ID_SEPARATOR, caption, TRUE);
			}
			if (nullptr != source)
			{
				m_wndStatusBar.AddElement(source, caption);
				m_wndStatusBar.RecalcLayout();
				m_wndStatusBar.Invalidate();
			}
		}
	}
	CString value;
	if (value.LoadString(IDS_EDIT_TITLE))
	{
		::AfxExtractSubString(value, value, index);
	}
	SendMessage(WM_SETMESSAGESTRING, 0, (LPARAM)caption);
	BOOL result = IDOK;
	switch (style)
	{
	case MB_ICONINFORMATION:
		break;
	default:
		result = CWnd::MessageBox(caption, value, style);
		break;
	}
	switch (result)
	{
	case IDCANCEL:
	case IDNO:
		MessageBox(IDS_AFXBARRES_CANCEL, MB_ICONINFORMATION);
		break;
	case IDABORT:
		MessageBox(IDS_EDIT_ABORT, MB_ICONINFORMATION);
		break;
	}
	return result;
}


LRESULT CMainFrame::OnSetmessagestring(WPARAM wParam, LPARAM lParam)
{
	CString value;
	if (0 != wParam)
	{
		if (value.LoadString((UINT)wParam))
		{
			OnSetmessagestring(0, (LPARAM)(LPCTSTR)value);
		}
	}
	else if (0 != lParam)
	{
		value = (LPCTSTR)lParam;
		CMFCRibbonStatusBarPane* source =
			(CMFCRibbonStatusBarPane*)m_wndStatusBar.FindElement(ID_SEPARATOR);
		if (nullptr != source)
		{
			source->SetText(value);
			m_wndStatusBar.RecalcLayout();
			m_wndStatusBar.Invalidate();
		}
	}
	return 0;
}


void CMainFrame::OnGetMinMaxInfo(MINMAXINFO* lpMMI)
{
	lpMMI->ptMinTrackSize.x = 560;
	lpMMI->ptMinTrackSize.y = 223;
	CFrameWnd::OnGetMinMaxInfo(lpMMI);
}

