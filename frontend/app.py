"""
eHojo BudgetChecker - 사용자 인터페이스(GUI) 모듈

이 모듈은 customtkinter를 사용하여 데스크톱 GUI 애플리케이션을 구성합니다.

주요 기능:
- 파일 선택 UI (예산서, 지출집행내역, 출력 폴더)
- 결과 파일명 자동 생성
- 실행 및 초기화 버튼
- 상태 표시

사용자 흐름:
1. 예산서 파일 선택 (찾아보기)
2. 지출집행내역 파일 선택 (찾아보기)
3. 출력 폴더 선택 (폴더 선택)
4. 실행 버튼 클릭
5. 결과 확인 (완료 메시지)

"""

import platform
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

import xlrd

from budget_checker.checker import BudgetChecker
from budget_checker.config import Constant
from budget_checker.excel_reader import validate_columns

# 아이콘 파일 경로
# PyInstaller로 빌드된 exe에서는 sys._MEIPASS 임시 폴더에서 assets를 찾음
# 일반 실행 시에는 프로젝트 루트의 assets 폴더에서 찾음
import sys
if getattr(sys, 'frozen', False):
    _BASE_DIR = Path(sys._MEIPASS)
else:
    _BASE_DIR = Path(__file__).parent.parent

ICON_PATH        = _BASE_DIR / 'assets' / 'icon.png'         # 창 아이콘 (검은색, macOS용)
ICON_ICO_PATH    = _BASE_DIR / 'assets' / 'icon.ico'         # 창 아이콘 (Windows용 ico)
ICON_HEADER_PATH = _BASE_DIR / 'assets' / 'icon_header.png'  # 헤더 아이콘 (#D8E4F4)
GUIDE_IMG_PATH   = _BASE_DIR / 'assets' / 'images' / 'Final_Combined_Budget.png'  # 안내 이미지

# ========== 애플리케이션 테마 설정 ==========
# 라이트 모드 사용 (밝은 화면)
ctk.set_appearance_mode('light')
# 기본 강조 색상을 파란색으로 설정
ctk.set_default_color_theme('blue')

# ========== 플랫폼별 폰트 설정 ==========
# macOS와 Windows에서 다른 폰트를 사용하여 최적의 가독성 제공
if platform.system() == 'Darwin':
    # macOS 환경
    FONT_KO   = 'Apple SD Gothic Neo'  # 한글 폰트
    FONT_MONO = 'Menlo'                # 고정폭 폰트 (코드, 경로 표시용)
else:
    # Windows 및 기타 환경
    FONT_KO   = 'Malgun Gothic'  # 한글 폰트
    FONT_MONO = 'Consolas'       # 고정폭 폰트

# ========== 색상 팔레트 ==========
# 주 강조색 (버튼, 링크 등에 사용)
ACCENT     = '#2E5FA3'  # 파란색
ACCENT_DK  = '#1E4A8A'  # 어두운 파란색 (호버 상태)
ACCENT_LT  = '#EBF1FA'  # 밝은 파란색 (배경)

# 텍스트 색상
TEXT       = '#1A2332'  # 기본 텍스트 (검은색)
TEXT_SEC   = '#5A6A80'  # 보조 텍스트 (진회색)
TEXT_MUTED = '#8A9AB0'  # 옅은 텍스트 (밝은 회색)

# 배경 색상
HEADER_BG  = '#2A3A52'  # 헤더 배경 (어두운 파란색)
BG         = '#F0F2F5'  # 기본 배경 (밝은 회색)
SURFACE    = '#FFFFFF'  # 표면 (흰색)

# ========== 폰트 크기 (픽셀) ==========
# 스케일링이 필요할 경우 여기서 일괄 조정 가능
FS_HEADER_TITLE = 20  # 헤더 제목 ("예산·집행 현황 생성기")
FS_HEADER_SUB   = 14  # 헤더 부제목 ("예산서와 지출집행내역을 병합하여...")
FS_SECTION      = 14  # 섹션 레이블 ("입력 파일", "출력 설정")
FS_FIELD_LABEL  = 16  # 필드 레이블 ("예산서 파일", "지출집행내역 파일")
FS_ENTRY        = 12  # 입력창 텍스트
FS_HINT         = 13  # 힌트 텍스트 (회색 설명)
FS_BROWSE       = 14  # 버튼 ("찾아보기", "폴더 선택")
FS_PREVIEW      = 14  # 파일명 미리보기
FS_RUN          = 18  # 실행 버튼 (크고 눈에 띄게)
FS_RESET        = 18  # 초기화 버튼
FS_STATUS       = 12  # 상태 표시줄


class App(ctk.CTk):
    """
    eHojo BudgetChecker 메인 GUI 애플리케이션 클래스.

    이 클래스는 customtkinter의 CTk를 상속하여 메인 윈도우를 구성합니다.
    - 파일 선택 인터페이스 제공
    - 사용자 입력 처리
    - 백엔드 처리 실행

    Attributes:
        budget_var (StringVar): 선택된 예산서 파일 경로
        exec_var (StringVar): 선택된 지출집행내역 파일 경로
        output_dir_var (StringVar): 선택된 출력 폴더 경로
        status_var (StringVar): 상태 표시줄에 표시할 메시지
        run_btn (CTkButton): 실행 버튼
        filename_label (CTkLabel): 결과 파일명 미리보기
    """

    def __init__(self):
        """
        애플리케이션 윈도우를 초기화합니다.

        - 윈도우 설정 (크기, 제목, 아이콘)
        - StringVar 변수 초기화
        - UI 구성 요소 생성
        """
        super().__init__()
        self.title('eHojo BudgetChecker')
        self.geometry('800x720')
        self.resizable(False, False)
        self.configure(fg_color=BG)

        self.budget_var     = ctk.StringVar()
        self.exec_var       = ctk.StringVar()
        self.output_dir_var = ctk.StringVar()

        # 창 아이콘 설정
        # Windows: iconbitmap (ico 파일)으로 설정해야 customtkinter 기본 아이콘을 덮어씀
        # macOS: iconphoto (png 파일)로 설정
        if platform.system() == 'Windows' and ICON_ICO_PATH.exists():
            self.after(200, lambda: self.iconbitmap(str(ICON_ICO_PATH)))
        elif ICON_PATH.exists():
            self._app_icon = tk.PhotoImage(file=str(ICON_PATH))
            self.iconphoto(True, self._app_icon)

        self._build_ui()

    # ================================================================
    # UI 구성
    # ================================================================

    def _build_ui(self):
        self._build_header()
        self._build_body()
        self._build_statusbar()
        self._check_date_message()

    def _build_header(self):
        header = ctk.CTkFrame(self, fg_color=HEADER_BG, corner_radius=0, height=90)
        header.pack(fill='x')
        header.pack_propagate(False)

        inner = ctk.CTkFrame(header, fg_color='transparent')
        inner.pack(expand=True, fill='both', padx=32)

        self._icon_click_times = []
        icon_canvas = tk.Canvas(inner, width=50, height=50, bg=HEADER_BG,
                                highlightthickness=0, bd=0)
        icon_canvas.pack(side='left', pady=20, padx=(0, 16))
        if ICON_HEADER_PATH.exists():
            self._header_icon = tk.PhotoImage(file=str(ICON_HEADER_PATH)).subsample(5)
            icon_canvas.create_image(25, 25, image=self._header_icon, anchor='center')
        icon_canvas.bind('<Button-1>', self._on_icon_click)

        text_frame = ctk.CTkFrame(inner, fg_color='transparent')
        text_frame.pack(side='left', pady=20)
        ctk.CTkLabel(text_frame,
                     text='예산·집행 현황 생성기',
                     font=(FONT_KO, FS_HEADER_TITLE, 'bold'),
                     text_color='#D8E4F4').pack(anchor='w')
        ctk.CTkLabel(text_frame,
                     text='예산서와 지출집행내역을 병합하여 결과 엑셀을 생성합니다',
                     font=(FONT_KO, FS_HEADER_SUB),
                     text_color='#7A9AC0').pack(anchor='w', pady=(5, 0))

    def _build_body(self):
        body = ctk.CTkFrame(self, fg_color=BG, corner_radius=0)
        body.pack(fill='x')

        inner = ctk.CTkFrame(body, fg_color='transparent')
        inner.pack(fill='x', padx=32, pady=20)

        self._section_label(inner, '입력 파일')

        self._file_field(inner,
                         label='예산서 파일',
                         var=self.budget_var,
                         hint='e호조 합본예산서(15100) 엑셀 자료를 입력하세요',
                         command=lambda: self._browse_file(self.budget_var, '예산서 파일'))

        self._file_field(inner,
                         label='지출집행내역 파일',
                         var=self.exec_var,
                         hint='e호조 지출집행현황(21126)에서 내려받은 엑셀 파일을 선택하세요',
                         command=lambda: self._browse_file(self.exec_var, '집행내역 파일'))

        self._section_label(inner, '출력 설정')

        self._folder_field(inner, '출력 폴더', self.output_dir_var)

        preview = ctk.CTkFrame(inner, fg_color=ACCENT_LT, corner_radius=8)
        preview.pack(fill='x', pady=(10, 0))

        preview_inner = ctk.CTkFrame(preview, fg_color='transparent')
        preview_inner.pack(fill='x', padx=16, pady=12)

        ctk.CTkLabel(preview_inner, text='📄',
                     font=(FONT_KO, 16)).pack(side='left', padx=(0, 10))

        self.filename_label = ctk.CTkLabel(preview_inner,
                                            text=self._make_filename(),
                                            font=(FONT_MONO, FS_PREVIEW, 'bold'),
                                            text_color=ACCENT)
        self.filename_label.pack(side='left')

        btn_row = ctk.CTkFrame(inner, fg_color='transparent')
        btn_row.pack(fill='x', pady=(20, 0))

        self.run_btn = ctk.CTkButton(btn_row,
                                      text='▶  실행',
                                      font=(FONT_KO, FS_RUN, 'bold'),
                                      fg_color=ACCENT,
                                      hover_color=ACCENT_DK,
                                      text_color='white',
                                      corner_radius=8,
                                      height=50,
                                      width=140,
                                      command=self._on_run)
        self.run_btn.pack(side='right')

        ctk.CTkButton(btn_row,
                      text='초기화',
                      font=(FONT_KO, FS_RESET),
                      fg_color='transparent',
                      hover_color='#E8ECF4',
                      text_color=TEXT_SEC,
                      border_color='#D0D5DD',
                      border_width=1,
                      corner_radius=8,
                      height=50,
                      width=96,
                      command=self._on_reset).pack(side='right', padx=(0, 10))

    def _build_statusbar(self):
        bar = ctk.CTkFrame(self, fg_color=HEADER_BG, corner_radius=0, height=36)
        bar.pack(fill='x', side='bottom')
        bar.pack_propagate(False)

        self.status_var = ctk.StringVar(value='대기 중')
        ctk.CTkLabel(bar, textvariable=self.status_var,
                     font=(FONT_MONO, FS_STATUS),
                     text_color='#8AA4C8').pack(side='left', padx=18)
        ctk.CTkLabel(bar, text='eHojo BudgetChecker v0.1.0',
                     font=(FONT_MONO, FS_STATUS),
                     text_color='#607A9C').pack(side='right', padx=18)

    # ================================================================
    # UI 헬퍼
    # ================================================================

    def _section_label(self, parent, text):
        frame = ctk.CTkFrame(parent, fg_color='transparent')
        frame.pack(fill='x', pady=(12, 10))

        ctk.CTkLabel(frame, text=text.upper(),
                     font=(FONT_KO, FS_SECTION, 'bold'),
                     text_color=TEXT_MUTED).pack(side='left')

        ctk.CTkFrame(frame, fg_color='#D0D5DD',
                     height=1, corner_radius=0).pack(
            side='left', fill='x', expand=True, padx=(10, 0), pady=1)

    def _file_field(self, parent, label, var, hint, command):
        ctk.CTkLabel(parent, text=label,
                     font=(FONT_KO, FS_FIELD_LABEL, 'bold'),
                     text_color=TEXT).pack(anchor='w', pady=(0, 6))

        row = ctk.CTkFrame(parent, fg_color='transparent')
        row.pack(fill='x', pady=(0, 4))

        ctk.CTkEntry(row,
                     textvariable=var,
                     font=(FONT_MONO, FS_ENTRY),
                     fg_color=SURFACE,
                     text_color=TEXT_SEC,
                     border_color='#D0D5DD',
                     border_width=1,
                     corner_radius=6,
                     height=46,
                     state='readonly').pack(side='left', fill='x', expand=True)

        ctk.CTkButton(row,
                      text='찾아보기',
                      font=(FONT_KO, FS_BROWSE, 'bold'),
                      fg_color=SURFACE,
                      hover_color='#F0F4FA',
                      text_color=TEXT,
                      border_color='#D0D5DD',
                      border_width=1,
                      corner_radius=6,
                      height=46,
                      width=108,
                      command=command).pack(side='left', padx=(8, 0))

        ctk.CTkLabel(parent, text=hint,
                     font=(FONT_KO, FS_HINT),
                     text_color=TEXT_MUTED).pack(anchor='w', pady=(2, 10))

    def _folder_field(self, parent, label, var):
        ctk.CTkLabel(parent, text=label,
                     font=(FONT_KO, FS_FIELD_LABEL, 'bold'),
                     text_color=TEXT).pack(anchor='w', pady=(0, 6))

        row = ctk.CTkFrame(parent, fg_color='transparent')
        row.pack(fill='x')

        ctk.CTkEntry(row,
                     textvariable=var,
                     font=(FONT_MONO, FS_ENTRY),
                     fg_color=SURFACE,
                     text_color=TEXT_SEC,
                     border_color='#D0D5DD',
                     border_width=1,
                     corner_radius=6,
                     height=46,
                     state='readonly').pack(side='left', fill='x', expand=True)

        ctk.CTkButton(row,
                      text='폴더 선택',
                      font=(FONT_KO, FS_BROWSE, 'bold'),
                      fg_color=SURFACE,
                      hover_color='#F0F4FA',
                      text_color=TEXT,
                      border_color='#D0D5DD',
                      border_width=1,
                      corner_radius=6,
                      height=46,
                      width=108,
                      command=self._browse_folder).pack(side='left', padx=(8, 0))

    # ================================================================
    # 파일/폴더 선택
    # ================================================================

    def _browse_file(self, var, label):
        path = filedialog.askopenfilename(
            title=f'{label} 선택',
            filetypes=[('Excel 파일', '*.xlsx'), ('모든 파일', '*.*')]
        )
        if path:
            var.set(path)

    def _browse_folder(self):
        path = filedialog.askdirectory(title='출력 폴더 선택')
        if path:
            self.output_dir_var.set(path)

    # ================================================================
    # 실행 / 초기화
    # ================================================================

    def _on_run(self):
        budget_path = self.budget_var.get()
        exec_path   = self.exec_var.get()
        output_dir  = self.output_dir_var.get()

        if not budget_path:
            messagebox.showwarning('입력 오류', '예산서 파일을 선택해주세요.')
            return
        if not exec_path:
            messagebox.showwarning('입력 오류', '지출집행내역 파일을 선택해주세요.')
            return
        if not output_dir:
            messagebox.showwarning('입력 오류', '출력 폴더를 선택해주세요.')
            return

        # 컬럼 검증
        try:
            xl_b = xlrd.open_workbook(budget_path)
            st_b = xl_b.sheets()[0]
        except Exception:
            messagebox.showerror('파일 오류',
                                 '예산서 파일을 열 수 없습니다.\n'
                                 '올바른 Excel 파일(.xlsx)인지 확인해주세요.')
            return

        try:
            xl_e = xlrd.open_workbook(exec_path)
            st_e = xl_e.sheets()[0]
        except Exception:
            messagebox.showerror('파일 오류',
                                 '지출집행내역 파일을 열 수 없습니다.\n'
                                 '올바른 Excel 파일(.xlsx)인지 확인해주세요.')
            return

        budget_missing = validate_columns(st_b, Constant.BUDGET_COLUMNS)
        exec_missing = validate_columns(st_e, Constant.EXECUTION_COLUMNS)

        if budget_missing or exec_missing:
            self._show_file_guide(budget_missing, exec_missing)
            return

        filename    = self._make_filename()
        output_path = str(Path(output_dir) / filename)

        self.filename_label.configure(text=filename)
        self._set_running(True)
        self.status_var.set('처리 중…')

        threading.Thread(
            target=self._run_checker,
            args=(budget_path, exec_path, output_path),
            daemon=True
        ).start()

    def _run_checker(self, budget_path, exec_path, output_path):
        try:
            BudgetChecker(
                budget_path=budget_path,
                execution_path=exec_path,
                output_path=output_path,
            )
            self.after(0, self._on_done, output_path)
        except Exception as e:
            self.after(0, self._on_error, str(e))

    def _on_done(self, output_path):
        self.status_var.set('완료')
        self._set_running(False)
        messagebox.showinfo('완료', f'파일이 생성되었습니다.\n\n{output_path}')

    def _on_error(self, msg):
        self.status_var.set('오류')
        self._set_running(False)
        messagebox.showerror('오류', msg)

    def _on_reset(self):
        self.budget_var.set('')
        self.exec_var.set('')
        self.output_dir_var.set('')
        self.filename_label.configure(text=self._make_filename())
        self.status_var.set('대기 중')

    # ================================================================
    # 파일 검증 안내 다이얼로그
    # ================================================================

    def _show_file_guide(self, budget_missing, exec_missing):
        win = ctk.CTkToplevel(self)
        win.title('파일 확인 필요')
        win.resizable(False, False)
        win.grab_set()
        win.configure(fg_color=SURFACE)

        # 다이얼로그 아이콘 설정
        if platform.system() == 'Windows' and ICON_ICO_PATH.exists():
            win.after(200, lambda: win.iconbitmap(str(ICON_ICO_PATH)))
        elif ICON_PATH.exists():
            self._guide_icon = tk.PhotoImage(file=str(ICON_PATH))
            win.iconphoto(True, self._guide_icon)

        # 헤더
        header = ctk.CTkFrame(win, fg_color=HEADER_BG, corner_radius=0, height=60)
        header.pack(fill='x')
        header.pack_propagate(False)
        ctk.CTkLabel(header, text='파일을 확인해주세요',
                     font=(FONT_KO, 18, 'bold'),
                     text_color='#D8E4F4').pack(expand=True)

        # 본문
        body = ctk.CTkFrame(win, fg_color='transparent')
        body.pack(fill='both', expand=True, padx=24, pady=16)

        ctk.CTkLabel(body,
                     text='선택한 파일에 필요한 컬럼이 없습니다.\n'
                          'e호조에서 아래 방법으로 다시 다운로드해주세요.',
                     font=(FONT_KO, 13),
                     text_color=TEXT_SEC,
                     justify='left').pack(anchor='w', pady=(0, 12))

        # 예산서 안내
        if budget_missing:
            self._add_guide_section(body, '합본예산서 다운로드', Constant.BUDGET_GUIDE)
            # 안내 이미지
            if GUIDE_IMG_PATH.exists():
                photo = tk.PhotoImage(file=str(GUIDE_IMG_PATH))
                factor = max(1, photo.width() // 440)
                if factor > 1:
                    photo = photo.subsample(factor, factor)
                self._guide_img_ref = photo
                tk.Label(body, image=photo, bg=SURFACE).pack(pady=(0, 10))

        # 집행내역 안내
        if exec_missing:
            self._add_guide_section(body, '지출집행현황(21126) 다운로드', Constant.EXECUTION_GUIDE)

        # 닫기 버튼
        btn_frame = ctk.CTkFrame(win, fg_color='transparent')
        btn_frame.pack(fill='x', padx=24, pady=(0, 16))
        ctk.CTkButton(btn_frame, text='확인',
                      font=(FONT_KO, 14, 'bold'),
                      fg_color=ACCENT,
                      hover_color=ACCENT_DK,
                      text_color='white',
                      corner_radius=6,
                      height=40,
                      width=100,
                      command=win.destroy).pack()

        # 내용에 맞춰 창 크기 자동 설정
        win.update_idletasks()
        req_h = min(win.winfo_reqheight(), 750)
        win.geometry(f'520x{req_h}')

    def _add_guide_section(self, parent, title, guide_text):
        frame = ctk.CTkFrame(parent, fg_color=ACCENT_LT, corner_radius=8)
        frame.pack(fill='x', pady=(0, 10))

        inner = ctk.CTkFrame(frame, fg_color='transparent')
        inner.pack(fill='x', padx=16, pady=12)

        ctk.CTkLabel(inner, text=title,
                     font=(FONT_KO, 14, 'bold'),
                     text_color=ACCENT).pack(anchor='w')

        ctk.CTkLabel(inner, text=guide_text,
                     font=(FONT_KO, 12),
                     text_color=TEXT,
                     justify='left',
                     anchor='w').pack(anchor='w', pady=(8, 0))

    # ================================================================
    # 유틸
    # ================================================================

    def _make_filename(self):
        return datetime.now().strftime('예산집행현황_%Y%m%d_%H%M%S.xlsx')

    def _set_running(self, running: bool):
        if running:
            self.run_btn.configure(state='disabled', fg_color='#A0B8D8', text='처리 중…')
        else:
            self.run_btn.configure(state='normal', fg_color=ACCENT, text='▶  실행')

    # ================================================================
    # 이스터에그
    # ================================================================

    def _on_icon_click(self, event=None):
        now = datetime.now().timestamp()
        self._icon_click_times = [t for t in self._icon_click_times if now - t < 1.5]
        self._icon_click_times.append(now)
        if len(self._icon_click_times) >= 5:
            self._icon_click_times = []
            self._show_easter_egg()

    def _show_easter_egg(self):
        win = ctk.CTkToplevel(self)
        win.title('정보')
        win.geometry('380x600')
        win.resizable(False, False)
        win.grab_set()
        win.configure(fg_color=HEADER_BG)

        # 창 아이콘 설정
        if platform.system() == 'Windows' and ICON_ICO_PATH.exists():
            win.after(200, lambda: win.iconbitmap(str(ICON_ICO_PATH)))
        elif ICON_PATH.exists():
            self._egg_icon = tk.PhotoImage(file=str(ICON_PATH))
            win.iconphoto(True, self._egg_icon)

        frame = ctk.CTkFrame(win, fg_color='transparent')
        frame.pack(expand=True, fill='both', padx=30, pady=16)

        ctk.CTkLabel(frame, text='eHojo BudgetChecker',
                     font=(FONT_KO, 18, 'bold'),
                     text_color='#D8E4F4').pack()
        ctk.CTkLabel(frame, text='v0.1.0',
                     font=(FONT_MONO, 13),
                     text_color='#7A9AC0').pack(pady=(2, 8))

        ctk.CTkFrame(frame, fg_color='#3A5070', height=1, corner_radius=0).pack(fill='x')

        ctk.CTkLabel(frame, text='"예산은 거짓말을 하지 않는다"',
                     font=(FONT_KO, 13, 'italic'),
                     text_color='#A8C4E0').pack(pady=(8, 2))
        ctk.CTkLabel(frame, text='"하지만 엑셀은 거짓말을 한다"',
                     font=(FONT_KO, 13, 'italic'),
                     text_color='#A8C4E0').pack(pady=(0, 8))

        ctk.CTkFrame(frame, fg_color='#3A5070', height=1, corner_radius=0).pack(fill='x')

        ctk.CTkLabel(frame, text='Made with ☕ and 야근',
                     font=(FONT_KO, 13),
                     text_color='#8AA4C8').pack(pady=(8, 2))
        ctk.CTkLabel(frame, text='Powered by 권혁수 팀장의 내리갈굼',
                     font=(FONT_KO, 13),
                     text_color='#8AA4C8').pack(pady=(0, 8))

        ctk.CTkFrame(frame, fg_color='#3A5070', height=1, corner_radius=0).pack(fill='x')

        for quote in ('"내가 왜? 널 시키면 되는데"',
                      '"너는 다 할 수 있잖아"',
                      '"다 너 능력 키워주려고 하는 거야"'):
            ctk.CTkLabel(frame, text=quote,
                         font=(FONT_KO, 12, 'italic'),
                         text_color='#7A9AC0').pack(pady=1)
        ctk.CTkLabel(frame, text='— 권혁수 팀장',
                     font=(FONT_KO, 12),
                     text_color='#607A9C').pack(pady=(2, 8))

        ctk.CTkFrame(frame, fg_color='#3A5070', height=1, corner_radius=0).pack(fill='x')

        thanks = ctk.CTkFrame(frame, fg_color='transparent')
        thanks.pack(pady=(8, 0))
        ctk.CTkLabel(thanks, text='Special Thanks',
                     font=(FONT_KO, 12, 'bold'),
                     text_color='#8AA4C8').pack()
        for item in ('☕ 자판기 아메리카노', '🌙 야근수당 (없음)', '🤖 Claude AI (진짜 일한 놈)'):
            ctk.CTkLabel(thanks, text=item,
                         font=(FONT_KO, 12),
                         text_color='#607A9C').pack()

        ctk.CTkLabel(frame, text='© 2026 갈굼당한 개발자',
                     font=(FONT_KO, 11),
                     text_color='#4A6080').pack(pady=(8, 0))

        ctk.CTkButton(frame, text='닫기',
                      font=(FONT_KO, 13, 'bold'),
                      fg_color=ACCENT,
                      hover_color=ACCENT_DK,
                      text_color='white',
                      corner_radius=6,
                      height=40,
                      width=100,
                      command=win.destroy).pack(pady=(10, 0))

    def _check_date_message(self):
        now = datetime.now()
        md = (now.month, now.day)
        date_msgs = {
            (1, 1):   '새해 복 많이 받으세요 🎊',
            (3, 1):   '대한독립 만세 🇰🇷',
            (5, 5):   '어린이날인데 왜 일하고 계세요?',
            (10, 9):  '한글날입니다. 보고서 맞춤법 확인하셨나요?',
            (12, 25): '메리 크리스마스 🎄',
        }
        msg = date_msgs.get(md)
        if msg is None and now.weekday() == 4 and now.hour >= 17:
            msg = '불금입니다. 이걸 보고 있다는 건… 퇴근 못하시는 거죠?'
        if msg:
            self.status_var.set(msg)
            self.after(5000, lambda: self.status_var.set('대기 중'))
