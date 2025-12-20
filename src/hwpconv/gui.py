import os
import sys
import threading
from pathlib import Path
from typing import Dict, Optional
from dataclasses import dataclass

try:
    import customtkinter as ctk
    import tkinterDnD  # python-tkdnd 패키지
except ImportError:
    pass 

from tkinter import filedialog, messagebox, simpledialog
import tkinter.font as tkfont
import tkinter as tk

from . import config as app_config

# 프리미엄 컨러 팔레트 (파란색 계열 통일)
COLORS = {
    'primary': '#2563eb',
    'primary_light': '#93c5fd',
    'primary_lighter': '#dbeafe',
    'primary_dark': '#1e40af',
    'text': '#1f2937',
    'text_secondary': '#6b7280',
    'text_muted': '#9ca3af',
    'border': '#e5e7eb',
    'bg': '#ffffff',
    'bg_subtle': '#f8fafc',    
    'success': '#059669',
    'error': '#ef4444',
    'surface': '#ffffff'
}

@dataclass
class FileItem:
    path: Path
    status: str = "pending"
    output_content: Optional[str] = None
    output_path: Optional[Path] = None

class HwpConverterApp(tkinterDnD.Tk):
    """python-tkdnd 기반 HWP 변환기 (드래그 피드백 완벽 지원)"""
    
    def __init__(self):
        super().__init__()
        
        ctk.set_appearance_mode("light")
        self.title("HWP 변환기")
        self.geometry("640x720")
        self.minsize(500, 600)
        self.configure(bg=COLORS['bg'])
        
        self.files: Dict[str, FileItem] = {}
        self.output_format = tk.StringVar(value='마크다운')
        self.enable_image_analysis = tk.BooleanVar(value=False)  # 이미지 분석 옵션 (기본 비활성)
        self._dot_count = 0  # 점 애니메이션 카운터

        self._setup_fonts()
        self._create_ui()
        self._start_animation()  # 점 애니메이션 시작

    def _setup_fonts(self):
        fonts = ['맑은 고딕', 'Malgun Gothic', 'Pretendard', 'Segoe UI', 'Arial']
        available = tkfont.families()
        self.font_family = next((f for f in fonts if f in available), 'Segoe UI')

    def _create_ui(self):
        # 메인 프레임 (CustomTkinter 사용)
        main = ctk.CTkFrame(self, fg_color=COLORS['bg'])
        main.pack(fill="both", expand=True, padx=0, pady=0)
        
        # 1. 헤더
        header = ctk.CTkFrame(main, fg_color="transparent")
        header.pack(fill="x", padx=30, pady=(30, 20))
        
        ctk.CTkLabel(
            header, text="HWP 변환기", 
            font=ctk.CTkFont(family=self.font_family, size=18, weight="bold"),
            text_color=COLORS['primary']
        ).pack(side="left")
        
        ctk.CTkLabel(
            header, text=" Pro", 
            font=ctk.CTkFont(family=self.font_family, size=18),
            text_color=COLORS['text_muted']
        ).pack(side="left")

        # 형식 선택 + 설정 버튼
        right_frame = ctk.CTkFrame(header, fg_color="transparent")
        right_frame.pack(side="right")
        
        # 설정 버튼
        ctk.CTkButton(
            right_frame, text="⚙", width=28, height=28,
            fg_color="transparent", text_color=COLORS['text_secondary'],
            hover_color=COLORS['bg_subtle'],
            font=ctk.CTkFont(size=16),
            command=self._show_settings
        ).pack(side="right", padx=(10, 0))
        
        # 형식 선택 버튼들
        fmt_frame = ctk.CTkFrame(right_frame, fg_color="transparent")
        fmt_frame.pack(side="right")
        
        self.btn_md = ctk.CTkButton(
            fmt_frame, text="마크다운", width=70, height=28,
            font=ctk.CTkFont(family=self.font_family, size=11, weight="bold"),
            command=lambda: self._set_format("마크다운")
        )
        self.btn_md.pack(side="left", padx=(0, 2))
        
        self.btn_html = ctk.CTkButton(
            fmt_frame, text="HTML", width=60, height=28,
            font=ctk.CTkFont(family=self.font_family, size=11, weight="bold"),
            command=lambda: self._set_format("HTML")
        )
        self.btn_html.pack(side="left")

        self._set_format("마크다운")

        # 이미지 분석 옵션 (형식 선택 아래)
        img_option_frame = ctk.CTkFrame(main, fg_color="transparent")
        img_option_frame.pack(fill="x", padx=30, pady=(5, 10))

        self.img_analysis_checkbox = ctk.CTkCheckBox(
            img_option_frame,
            text="🖼️ 이미지 분석 (Gemini API)",
            variable=self.enable_image_analysis,
            font=ctk.CTkFont(family=self.font_family, size=11),
            text_color=COLORS['text'],
            command=self._on_image_analysis_toggle
        )
        self.img_analysis_checkbox.pack(side="left")

        self.img_analysis_warning = ctk.CTkLabel(
            img_option_frame,
            text="⚠️ 분석 시 변환 시간이 크게 증가합니다",
            font=ctk.CTkFont(family=self.font_family, size=10),
            text_color="#b45309"
        )
        # 초기에는 숨김 (체크박스 활성화 시 표시)

        # API 키 없으면 비활성화
        self._update_image_analysis_state()

        # 2. 드롭 영역 (tk.Frame 사용 - python-tkdnd 호환)
        self.drop_frame = tk.Frame(
            main, bg=COLORS['bg_subtle'], 
            highlightbackground=COLORS['primary_light'], highlightthickness=2,
            height=160
        )
        self.drop_frame.pack(fill="x", padx=30, pady=(0, 20))
        self.drop_frame.pack_propagate(False)
        
        # 드래그 앤 드롭 등록
        self.drop_frame.register_drop_target("*")
        self.drop_frame.bind("<<DropEnter>>", self._on_drag_enter)
        self.drop_frame.bind("<<DropLeave>>", self._on_drag_leave)
        self.drop_frame.bind("<<Drop>>", self._on_drop)
        self.drop_frame.bind("<Button-1>", lambda e: self._select_files())
        
        # 드롭 영역 내부
        self.drop_inner = tk.Frame(self.drop_frame, bg=COLORS['bg_subtle'])
        self.drop_inner.place(relx=0.5, rely=0.5, anchor="center")
        
        self.drop_icon = tk.Label(
            self.drop_inner, text="📁", font=(self.font_family, 24), bg=COLORS['bg_subtle']
        )
        self.drop_icon.pack()
        
        self.drop_title = tk.Label(
            self.drop_inner, text="HWP/HWPX 파일을 여기에 드래그하세요",
            font=(self.font_family, 12, "bold"), bg=COLORS['bg_subtle'], fg=COLORS['text']
        )
        self.drop_title.pack(pady=(6, 2))
        
        self.drop_subtitle = tk.Label(
            self.drop_inner, text="또는 클릭하여 파일 선택",
            font=(self.font_family, 10), bg=COLORS['bg_subtle'], fg=COLORS['text_secondary']
        )
        self.drop_subtitle.pack()

        # 3. 작업 버튼
        actions = ctk.CTkFrame(main, fg_color="transparent")
        actions.pack(fill="x", padx=30, pady=(0, 10))
        
        self.status_label = ctk.CTkLabel(
            actions, text="준비됨", text_color=COLORS['text_secondary'],
            font=ctk.CTkFont(family=self.font_family, size=12)
        )
        self.status_label.pack(side="left")
        
        ctk.CTkButton(
            actions, text="목록 지우기", width=90, fg_color="transparent", border_width=1,
            border_color=COLORS['border'], text_color=COLORS['text_secondary'],
            font=ctk.CTkFont(family=self.font_family, size=12), command=self._clear_files
        ).pack(side="right", padx=(10, 0))
        
        ctk.CTkButton(
            actions, text="전체 저장", width=100, fg_color=COLORS['success'], hover_color="#047857",
            font=ctk.CTkFont(family=self.font_family, size=12, weight="bold"), command=self._download_all
        ).pack(side="right")

        # 4. 파일 목록
        self.list_container = ctk.CTkScrollableFrame(
            main, fg_color="transparent", corner_radius=0
        )
        self.list_container.pack(fill="both", expand=True, padx=30, pady=10)
        
        # 5. 하단 안내
        footer = ctk.CTkFrame(main, fg_color="#eff6ff", corner_radius=8)
        footer.pack(fill="x", padx=30, pady=(10, 30))
        
        self.footer_label = ctk.CTkLabel(
            footer, 
            text="",
            font=ctk.CTkFont(family=self.font_family, size=11),
            text_color=COLORS['primary_dark'], justify="left", anchor="w"
        )
        self.footer_label.pack(padx=15, pady=8, fill="x")
        self._update_footer_status()
    
    def _update_footer_status(self):
        """하단 상태 메시지 업데이트"""
        if app_config.has_api_key():
            self.footer_label.configure(
                text="💡 HWP/HWPX 파일의 표, 텍스트, 이미지를 완벽하게 변환합니다. (이미지 분석 활성화 가능)",
                text_color=COLORS['primary_dark']
            )
        else:
            self.footer_label.configure(
                text="⚠️ API 키 미설정 - 이미지 분석이 비활성화됩니다. (⚙ 설정에서 Gemini API 키 입력)",
                text_color="#b45309"  # 주황색
            )
    
    def _show_settings(self):
        """설정 다이얼로그 표시"""
        current_key = app_config.get_api_key()
        masked = current_key[:8] + "..." if current_key else "(미설정)"
        
        new_key = simpledialog.askstring(
            "API 설정",
            f"Gemini API 키를 입력하세요.\n현재: {masked}\n\n(이미지 분석에 사용됩니다)",
            parent=self
        )
        
        if new_key is not None:  # 취소가 아닌 경우
            if new_key.strip():
                app_config.save_api_key(new_key.strip())
                messagebox.showinfo("저장 완료", "API 키가 저장되었습니다.")
            else:
                app_config.save_api_key("")
                messagebox.showinfo("초기화", "API 키가 삭제되었습니다.")
            self._update_footer_status()
            self._update_image_analysis_state()
    
    def _set_format(self, fmt):
        """형식 선택 및 버튼 상태 업데이트"""
        self.output_format.set(fmt)
        if fmt == "마크다운":
            self.btn_md.configure(fg_color=COLORS['primary'], text_color="white", hover_color=COLORS['primary_dark'])
            self.btn_html.configure(fg_color=COLORS['border'], text_color=COLORS['text'], hover_color=COLORS['bg_subtle'])
        else:
            self.btn_md.configure(fg_color=COLORS['border'], text_color=COLORS['text'], hover_color=COLORS['bg_subtle'])
            self.btn_html.configure(fg_color=COLORS['primary'], text_color="white", hover_color=COLORS['primary_dark'])

    def _on_image_analysis_toggle(self):
        """이미지 분석 옵션 토글 시 경고 표시/숨김"""
        if self.enable_image_analysis.get():
            self.img_analysis_warning.pack(side="left", padx=(15, 0))
        else:
            self.img_analysis_warning.pack_forget()

    def _update_image_analysis_state(self):
        """API 키 상태에 따라 이미지 분석 옵션 활성화/비활성화"""
        if app_config.has_api_key():
            self.img_analysis_checkbox.configure(state="normal")
            self.img_analysis_checkbox.configure(text="🖼️ 이미지 분석 (Gemini API)")
        else:
            self.img_analysis_checkbox.configure(state="disabled")
            self.enable_image_analysis.set(False)
            self.img_analysis_checkbox.configure(text="🖼️ 이미지 분석 (API 키 필요)")
            self.img_analysis_warning.pack_forget()
    
    def _start_animation(self):
        """애니메이션 타이머 (현재 비활성화)"""
        pass  # 깜박거림 방지를 위해 비활성화

    def _on_drag_enter(self, event):
        """드래그 진입 시 시각적 피드백 (파란색 계열)"""
        self.drop_frame.configure(bg=COLORS['primary_lighter'], highlightbackground=COLORS['primary'], highlightthickness=3)
        self.drop_inner.configure(bg=COLORS['primary_lighter'])
        self.drop_icon.configure(bg=COLORS['primary_lighter'], text="⬇")
        self.drop_title.configure(bg=COLORS['primary_lighter'], text="여기에 놓으세요!", fg=COLORS['primary'])
        self.drop_subtitle.configure(bg=COLORS['primary_lighter'], text="")
        return event.action
        
    def _on_drag_leave(self, event):
        """드래그 이탈 시 원래대로"""
        self._reset_drop_zone()
        return event.action

    def _on_drop(self, event):
        """드롭 성공"""
        self._reset_drop_zone()
        files = self.tk.splitlist(event.data)
        self._add_files(files)
        return event.action
    
    def _reset_drop_zone(self):
        """드롭 영역 초기화"""
        self.drop_frame.configure(bg=COLORS['bg_subtle'], highlightbackground=COLORS['primary_light'], highlightthickness=2)
        self.drop_inner.configure(bg=COLORS['bg_subtle'])
        self.drop_icon.configure(bg=COLORS['bg_subtle'], text="📁")
        self.drop_title.configure(bg=COLORS['bg_subtle'], text="HWP/HWPX 파일을 여기에 드래그하세요", fg=COLORS['text'])
        self.drop_subtitle.configure(bg=COLORS['bg_subtle'], text="또는 클릭하여 파일 선택")

    def _select_files(self):
        files = filedialog.askopenfilenames(
            title="변환할 파일 선택",
            filetypes=[("HWP/HWPX 파일", "*.hwp *.hwpx"), ("모든 파일", "*.*")]
        )
        if files:
            self._add_files(files)

    def _add_files(self, files):
        new_files = []
        for f in files:
            path = Path(f)
            if path.suffix.lower() in ['.hwp', '.hwpx']:
                key = str(path)
                if key not in self.files:
                    item = FileItem(path=path)
                    self.files[key] = item
                    new_files.append(key)
        
        self._update_list()
        if new_files:
            threading.Thread(target=self._process_queue, args=(new_files,), daemon=True).start()

    def _process_queue(self, keys):
        import time
        fmt = self.output_format.get()
        analyze_images = self.enable_image_analysis.get()  # 이미지 분석 옵션

        try:
            from .parsers.hwp import HwpParser
            from .parsers.hwpx import HwpxParser
            from .converters.markdown import MarkdownConverter
            from .converters.html import HtmlConverter

            for key in keys:
                item = self.files.get(key)
                if not item: continue

                item.status = 'converting'
                self.after(0, self._update_list)

                start_time = time.time()
                try:
                    ext = item.path.suffix.lower()
                    if ext == '.hwpx':
                        doc = HwpxParser().parse(str(item.path), analyze_images=analyze_images)
                    else:
                        doc = HwpParser().parse(str(item.path), analyze_images=analyze_images)
                    
                    if fmt == 'HTML':
                        item.output_content = HtmlConverter(include_images=True).convert(doc)
                    else:
                        item.output_content = MarkdownConverter(include_images=True).convert(doc)
                        
                    item.status = 'success'
                    elapsed = time.time() - start_time
                    
                    # 로그에 변환 시간 기록
                    log_msg = f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] ✅ 변환 완료: {item.path.name} ({elapsed:.2f}초, 이미지 {len(doc.images)}개)\n"
                    print(log_msg.strip())
                    log_path = app_config.get_config_dir() / 'gemini_debug.log'
                    with open(log_path, 'a', encoding='utf-8') as f:
                        f.write(log_msg)
                        
                except Exception as e:
                    elapsed = time.time() - start_time
                    log_msg = f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] ❌ 변환 실패: {item.path.name} ({elapsed:.2f}초) - {str(e)}\n"
                    print(log_msg.strip())
                    log_path = app_config.get_config_dir() / 'gemini_debug.log'
                    with open(log_path, 'a', encoding='utf-8') as f:
                        f.write(log_msg)
                    item.status = 'error'
                
                self.after(0, self._update_list)
        except Exception as e:
            print(f"Import 오류: {e}")

    def _clear_files(self):
        self.files.clear()
        self._update_list()

    def _remove_file(self, key):
        if key in self.files:
            del self.files[key]
            self._update_list()

    def _update_list(self):
        for widget in self.list_container.winfo_children():
            widget.destroy()
            
        success_count = sum(1 for f in self.files.values() if f.status == 'success')
        self.status_label.configure(text=f"{len(self.files)}개 파일 ({success_count}개 완료)")
        
        for key, item in self.files.items():
            row = ctk.CTkFrame(self.list_container, fg_color=COLORS['surface'], height=40)
            row.pack(fill="x", pady=2)
            
            # 상태별 아이콘 및 색상
            if item.status == 'success':
                icon, color = "✓", COLORS['success']
            elif item.status == 'converting':
                icon, color = "⟳", COLORS['primary']
            elif item.status == 'error':
                icon, color = "✗", COLORS['error']
            else:
                icon, color = "○", COLORS['text_muted']
            
            ctk.CTkLabel(row, text=icon, text_color=color, font=ctk.CTkFont(size=14)).pack(side="left", padx=(10, 5))
            
            # 파일명 (최대 35자, 초과시 말줄임표)
            name = item.path.name
            if len(name) > 35:
                name = name[:32] + "..."
            
            ctk.CTkLabel(
                row, text=name, text_color=COLORS['text'],
                font=ctk.CTkFont(family=self.font_family, size=11)
            ).pack(side="left")
            
            # 버튼들 (오른쪽 고정)
            ctk.CTkButton(
                row, text="✕", width=24, height=24, fg_color="transparent", text_color=COLORS['text_secondary'],
                hover_color=COLORS['bg_subtle'], command=lambda k=key: self._remove_file(k)
            ).pack(side="right", padx=2)
            
            if item.status == 'success':
                ctk.CTkButton(
                    row, text="저장", width=45, height=24, fg_color=COLORS['primary'],
                    font=ctk.CTkFont(size=10), command=lambda k=key: self._save_file(k)
                ).pack(side="right", padx=2)
            elif item.status == 'converting':
                # 점 애니메이션 (., .., ...)
                dots = "." * (self._dot_count % 3 + 1)
                ctk.CTkLabel(
                    row, text=dots, text_color=COLORS['primary'], width=25,
                    font=ctk.CTkFont(size=14, weight="bold")
                ).pack(side="right", padx=5)

    def _save_file(self, key):
        item = self.files.get(key)
        if not item or not item.output_content: return
        
        ext = '.md' if self.output_format.get() == '마크다운' else '.html'
        default_folder = self._get_output_folder()
        
        path = filedialog.asksaveasfilename(
            title="파일 저장",
            initialdir=str(default_folder),
            defaultextension=ext, initialfile=item.path.stem + ext
        )
        if path:
            Path(path).write_text(item.output_content, encoding='utf-8')
            messagebox.showinfo("완료", f"저장되었습니다:\n{path}")

    def _get_output_folder(self) -> Path:
        """출력 폴더 반환 (EXE 실행 폴더/HWPCONV)"""
        if getattr(sys, 'frozen', False):
            base = Path(sys.executable).parent
        else:
            base = Path.cwd()
        
        output_dir = base / "HWPCONV_Output"
        output_dir.mkdir(exist_ok=True)
        return output_dir

    def _download_all(self):
        ready = [f for f in self.files.values() if f.status == 'success']
        if not ready:
            messagebox.showwarning("알림", "저장할 파일이 없습니다.")
            return
        
        # 자동으로 HWPCONV_Output 폴더에 저장
        folder = self._get_output_folder()
        ext = '.md' if self.output_format.get() == '마크다운' else '.html'
        
        for item in ready:
            output_path = folder / (item.path.stem + ext)
            output_path.write_text(item.output_content, encoding='utf-8')
        
        messagebox.showinfo("완료", f"{len(ready)}개 파일이 저장되었습니다.\n\n📁 {folder}")

if __name__ == "__main__":
    app = HwpConverterApp()
    app.mainloop()
