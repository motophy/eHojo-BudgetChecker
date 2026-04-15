"""
budget_checker 패키지

예산서와 지출집행내역 파일을 처리하는 백엔드 모듈입니다.

모듈 구성:
- config: 상수 및 설정 정의
- excel_reader: 엑셀 파일 읽기 유틸리티
- excel_writer: 엑셀 파일 쓰기 및 서식 적용
- checker: 메인 처리 클래스

사용 예:
    from budget_checker.checker import BudgetChecker
    checker = BudgetChecker(
        budget_path='예산서.xlsx',
        execution_path='집행내역.xlsx',
        output_path='결과.xlsx'
    )
"""
