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

// MainFrm.h : CMainFrame クラスのインターフェイス
//

#pragma once

class CMainFrame : public CFrameWndEx
{
	
protected: // シリアル化からのみ作成します。
	CMainFrame() noexcept;
	DECLARE_DYNCREATE(CMainFrame)

// 属性
public:

// 操作
public:

// オーバーライド
public:
	virtual BOOL PreCreateWindow(CREATESTRUCT& cs);

	/// <summary>
	/// メッセージボックスの処理
	/// </summary>
	/// <param name="nID">メッセージ文字列のID</param>
	/// <param name="style">メッセージボックスのスタイル</param>
	/// <returns>処理結果</returns>
	UINT MessageBox(UINT nID, UINT style = MB_ICONEXCLAMATION);
	/// <summary>
	/// メッセージボックスの処理
	/// </summary>
	/// <param name="nID">メッセージ文字列</param>
	/// <param name="style">メッセージボックスのスタイル</param>
	/// <returns>処理結果</returns>
	UINT MessageBox(LPCTSTR caption, UINT style = MB_ICONEXCLAMATION);
// 実装
public:
	virtual ~CMainFrame();
#ifdef _DEBUG
	virtual void AssertValid() const;
	virtual void Dump(CDumpContext& dc) const;
#endif
	/// <summary>
	/// 小さいアイコンのイメージリスト
	/// </summary>
	CImageList m_imgSmall;
	/// <summary>
	/// 大きいアイコンのイメージリスト
	/// </summary>
	CImageList m_imgLarge;

protected:  // コントロール バー用メンバー
	CMFCRibbonBar     m_wndRibbonBar;
	CMFCRibbonApplicationButton m_MainButton;
	CMFCToolBarImages m_PanelImages;
	CMFCRibbonStatusBar  m_wndStatusBar;
	/// <summary>
	/// 現在日時の保持先
	/// </summary>
	SYSTEMTIME        m_systemTime;

// 生成された、メッセージ割り当て関数
protected:
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg void OnApplicationLook(UINT id);
	afx_msg void OnUpdateApplicationLook(CCmdUI* pCmdUI);
	afx_msg void OnFilePrint();
	afx_msg void OnFilePrintPreview();
	afx_msg void OnUpdateFilePrintPreview(CCmdUI* pCmdUI);
	DECLARE_MESSAGE_MAP()

	/// <summary>
	/// ファイルを「開く」処理の許可禁止処理
	/// </summary>
	/// <param name="pCmdUI"></param>
	afx_msg void OnUpdateFileOpen(CCmdUI* pCmdUI);
	/// <summary>
	/// 「新規作成」「ファイルを保存」「名前をつけて保存」などの許可/禁止処理
	/// </summary>
	/// <param name="pCmdUI"></param>
	afx_msg void OnUpdateFileNew(CCmdUI* pCmdUI);
	/// <summary>
	/// ステータスバーの行位置表示処理
	/// </summary>
	/// <param name="pCmdUI"></param>
	afx_msg void OnUpdateIndicatorRow(CCmdUI* pCmdUI);
	/// <summary>
	/// ステータスバーの列位置表示処理
	/// </summary>
	/// <param name="pCmdUI"></param>
	afx_msg void OnUpdateIndicatorCol(CCmdUI* pCmdUI);
	/// <summary>
	/// ステータスバーの編集中表示処理
	/// </summary>
	/// <param name="pCmdUI"></param>
	afx_msg void OnUpdateIndicatorModify(CCmdUI* pCmdUI);
	/// <summary>
	/// ステータスバーの日付表示処理
	/// </summary>
	/// <param name="pCmdUI"></param>
	afx_msg void OnUpdateIndicatorDate(CCmdUI* pCmdUI);
	/// <summary>
	/// ステータスバーの時刻表示処理
	/// </summary>
	/// <param name="pCmdUI"></param>
	afx_msg void OnUpdateIndicatorTime(CCmdUI* pCmdUI);
	/// <summary>
	/// インターバルタイマー処理
	/// </summary>
	/// <param name="nIDEvent"></param>
	afx_msg void OnTimer(UINT_PTR nIDEvent);
	/// <summary>
	/// 終了処理
	/// </summary>
	afx_msg void OnClose();
	/// <summary>
	/// ステータスバーのステータス文字表示処理
	/// </summary>
	/// <param name="wParam"></param>
	/// <param name="lParam"></param>
	/// <returns></returns>
	afx_msg LRESULT OnSetmessagestring(WPARAM wParam, LPARAM lParam);
	/// <summary>
	/// メインフレームの最小の大きさの処理
	/// </summary>
	/// <param name="lpMMI"></param>
	afx_msg void OnGetMinMaxInfo(MINMAXINFO* lpMMI);
};


