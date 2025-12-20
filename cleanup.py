import os
import glob

# 제거할 파일 패턴
patterns = [
    # 기존 임시 파일들
    "analyze_*.py", "debug_*.py", "check_*.py", "verify_result.py", "simple_compare.py", "compare_all.py",
    "struct_log.txt", "struct_log_head.txt", "compare_result.txt",
    "hwp_analysis.txt", "kanji_list.txt", "short_paragraphs.txt",
    "*.log",
    
    # 임시 결과물 MD 파일들
    "test_*.md", "test_*.html",
    "clean_*.md", "final_*.md", "perfect_*.md",
    
    # 원본 파일명 기반 생성된 구버전 결과물 (필요시 주석 처리)
    "비영리법인 설립신청 및 업무서식 안내-공개용.md", 
    "정관 작성기준 및 정관 예시.md"
]

# 유지할 파일 (혹시 패턴에 걸리더라도 보호)
keep_files = [
    "README.md", "hwp-converter-prd.md",
    "run_gui.py", "cleanup.py",
    "solution_hwpx.md", "solution_hwp.md" # 최종 결과물은 유지
]

for p in patterns:
    for f in glob.glob(p):
        if f not in keep_files:
            try:
                os.remove(f)
                print(f"Deleted: {f}")
            except Exception as e:
                print(f"Error deleting {f}: {e}")
