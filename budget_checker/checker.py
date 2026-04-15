"""
예산·집행 현황 병합 처리 모듈

이 모듈은 e호조에서 내려받은 예산서와 지출집행내역 파일을 읽고,
이를 병합하여 최종 엑셀 파일을 생성합니다.

주요 처리 단계:
1. 예산서 파일 로드
2. 지출집행내역 파일 로드
3. 예산 항목별 지출 매칭
4. 최종 엑셀 파일 생성

작성자: eHojo BudgetChecker Team
버전: 0.1.0
"""

import xlrd

from budget_checker.config import Constant
from budget_checker.excel_reader import (
    get_unique_items,
    get_joined_text,
    get_row_values,
    get_sum_value,
    get_rows_sorted,
)
from budget_checker.excel_writer import XlWriter


class BudgetChecker:
    """
    예산서와 지출집행내역을 병합하는 메인 처리 클래스.

    이 클래스는 두 개의 엑셀 파일을 읽어서 병합한 후
    최종 결과를 새로운 엑셀 파일로 생성합니다.

    Attributes:
        xl_budget: 예산서 엑셀 파일 객체
        st_budget: 예산서 엑셀의 첫 번째 시트
        xl_execution: 지출집행내역 엑셀 파일 객체
        st_execution: 지출집행내역 엑셀의 첫 번째 시트
        writer: 결과 엑셀 파일을 생성하는 XlWriter 객체
    """
    def __init__(self,
                 budget_path='합본예산서최종(세출) (2).xlsx',
                 execution_path='지출집행현황지출결의_20260414092954.xlsx',
                 output_path='result.xlsx'):
        """
        BudgetChecker 객체를 초기화합니다.

        생성자에서 예산서와 지출집행내역 파일을 로드한 후
        자동으로 병합 처리(run 메소드)를 실행합니다.

        Args:
            budget_path (str): 예산서 엑셀 파일 경로
                기본값: '합본예산서최종(세출) (2).xlsx'
            execution_path (str): 지출집행내역 엑셀 파일 경로
                기본값: '지출집행현황지출결의_20260414092954.xlsx'
            output_path (str): 결과 엑셀 파일을 저장할 경로
                기본값: 'result.xlsx'

        Raises:
            FileNotFoundError: 입력 파일이 없을 때 발생
            xlrd.XLRDError: 엑셀 파일 형식이 잘못되었을 때 발생
        """
        # 예산서 파일 로드
        self.xl_budget = xlrd.open_workbook(budget_path)
        self.st_budget = self.xl_budget.sheets()[0]  # 첫 번째 시트만 사용

        # 지출집행내역 파일 로드
        self.xl_execution = xlrd.open_workbook(execution_path)
        self.st_execution = self.xl_execution.sheets()[0]  # 첫 번째 시트만 사용

        # 결과 파일 생성 도구 초기화
        self.writer = XlWriter(output_path)

        # 병합 처리 실행 (생성자에서 자동으로 처리)
        self.run()

    def run(self):
        """
        예산서와 지출집행내역을 병합하여 최종 엑셀 파일을 생성합니다.

        처리 단계:
        1. 예산서에서 고유한 예산 항목 추출
        2. 각 예산 항목별로:
           - 예산 정보 수집
           - 여러 줄의 내용 병합
           - 예산액 합계 계산
           - 매칭되는 지출 내역 검색 및 정렬
           - 엑셀 행 생성
        3. 최종 엑셀 파일 저장

        Returns:
            None
        """
        # 1단계: 예산서에서 고유한 예산 항목 추출
        # (부서코드, 정책사업코드 등의 조합으로 구분된 항목)
        unique_budget_items = get_unique_items(self.st_budget, Constant.BUDGET_ITEM_COLUMNS)
        total = len(unique_budget_items)

        # 2단계: 각 예산 항목을 처리
        for i, budget_items in enumerate(unique_budget_items):
            # 예산 정보를 딕셔너리로 변환
            budget_write_data = get_row_values(self.st_budget, Constant.BUDGET_ITEM_COLUMNS, budget_items,
                                               Constant.BUDGET_COLUMNS)

            # 같은 예산 항목이 여러 행인 경우, 해당 컬럼들의 값을 줄바꿈으로 병합
            for budget_join_column in Constant.BUDGET_JOIN_COLUMNS:
                budget_write_data[budget_join_column] = get_joined_text(self.st_budget, Constant.BUDGET_ITEM_COLUMNS,
                                                                        budget_items, budget_join_column)

            # 예산액 합계 계산 (예산서는 천단위이므로 1000을 곱해 원단위로 변환)
            budget_write_data['사업별예산액'] = get_sum_value(self.st_budget, Constant.BUDGET_ITEM_COLUMNS,
                                                         budget_items, '예산액') * 1000

            # 이 예산 항목과 매칭되는 지출 내역을 찾아서 지급일자순으로 정렬
            executions = get_rows_sorted(self.st_execution, Constant.EXECUTION_ITEM_COLUMNS,
                                         budget_items, Constant.EXECUTION_COLUMNS)
            budget_write_data['지출집행내역'] = executions

            # 마지막 행인지 판단 (엑셀 서식 적용 시 필요)
            last_element = (i == total - 1)
            # 엑셀 행 생성
            self.writer.create_xl(budget_write_data, last_element)

        # 3단계: 최종 엑셀 파일 저장
        self.writer.close()
