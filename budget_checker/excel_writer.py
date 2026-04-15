import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell, xl_range

from budget_checker.config import Constant


class XlWriter:
    """
    예산·지출 데이터를 엑셀 파일로 작성하는 클래스.

    ──────────────────────────────────────────────
    서식 수정 가이드 (아래 상수만 변경하면 전체 반영)
    ──────────────────────────────────────────────
    WRAP_COLUMNS              : 자동줄바꿈 적용 컬럼
    NUMBER_COLUMNS            : 천단위 콤마 숫자 서식 컬럼
    LEFT_THICK_BORDER_COLUMNS : 왼쪽 굵은 테두리 컬럼
    RIGHT_THICK_BORDER_COLUMNS: 오른쪽 굵은 테두리 컬럼
    ──────────────────────────────────────────────
    """

    # ================================================================
    # 서식 카테고리 상수 (수정 시 여기만 변경)
    # ================================================================

    # 헤더 배경색
    HEADER_BG_COLOR = '#E2EFDA'  # 연한 연두색

    # 왼쪽 정렬 컬럼
    LEFT_ALIGN_COLUMNS = frozenset({'산출근거명', '산출근거식'})

    # 오른쪽 정렬 컬럼 (NUMBER_COLUMNS는 자동으로 오른쪽 정렬되므로 여기엔 그 외 컬럼만)
    RIGHT_ALIGN_COLUMNS = frozenset({'예산액'})

    # shrink 대신 자동줄바꿈을 적용할 컬럼
    WRAP_COLUMNS = frozenset({
        '의무/재량구분', '산출근거명', '산출근거식', '예산구분', '예산액',
    })

    # 천단위 콤마 숫자 서식 (소수점 없음)
    NUMBER_COLUMNS = frozenset({
        '사업별예산액', '잔액', '총지출금액', '결의금액',
    })

    # 왼쪽 테두리를 굵게 할 컬럼 (맨왼쪽은 자동 처리)
    LEFT_THICK_BORDER_COLUMNS = frozenset({'사업별예산액'})

    # 오른쪽 테두리를 굵게 할 컬럼 (맨오른쪽은 자동 처리)
    RIGHT_THICK_BORDER_COLUMNS = frozenset({'총지출금액'})

    # ================================================================
    # 초기화
    # ================================================================

    def __init__(self, output_path='result.xlsx'):
        self.workbook = xlsxwriter.Workbook(output_path)
        self.worksheet = self.workbook.add_worksheet()
        self.write_idx = 1  # 데이터 시작 행 (0행은 헤더)

        # 컬럼 구조 정의
        self.BUDGET_COLUMNS = (
            Constant.BUDGET_COLUMNS
            + Constant.BUDGET_JOIN_COLUMNS
            + ('사업별예산액',)
        )
        self.EXECUTION_COLUMNS = Constant.EXECUTION_COLUMNS

        # 잔액·총지출금액은 예산과 집행 사이에 위치
        self.CALC_COLUMNS = ('잔액', '총지출금액')

        # 전체 컬럼 순서 (헤더 작성 및 인덱스 계산용)
        self.ALL_COLUMNS = (
            self.BUDGET_COLUMNS
            + self.CALC_COLUMNS
            + self.EXECUTION_COLUMNS
        )

        # 자주 쓰는 컬럼 인덱스를 미리 계산 (매번 len() 호출 방지)
        self.COL_IDX_BALANCE = len(self.BUDGET_COLUMNS)      # '잔액' 위치
        self.COL_IDX_TOTAL = len(self.BUDGET_COLUMNS) + 1    # '총지출금액' 위치
        self.COL_IDX_EXEC_START = len(self.BUDGET_COLUMNS) + 2  # 집행 데이터 시작
        self.COL_IDX_LAST = len(self.ALL_COLUMNS) - 1        # 맨 오른쪽

        # 포맷 캐시 (동일 조합의 포맷 객체 재사용)
        self._format_cache = {}

        # 헤더 작성
        self._write_header()
        self._apply_column_widths()

    # ================================================================
    # 서식 생성
    # ================================================================

    def _apply_column_widths(self):
        """ALL_COLUMNS 순서대로 너비 지정 또는 숨김 처리한다."""

        for col_index, col_name in enumerate(self.ALL_COLUMNS):
            # 설정에 없으면 기본 너비 사용
            width = Constant.COLUMN_WIDTH.get(col_name, Constant.DEFAULT_WIDTH)

            if not width:
                # 0 또는 None → 숨김
                self.worksheet.set_column(col_index, col_index, None, None, {'hidden': True})
            else:
                # 지정 너비 적용
                self.worksheet.set_column(col_index, col_index, width)

    def _build_format_props(self, row_type, content_type, is_first_col, is_last_col, column_name):
        """
        조건에 맞는 서식 속성 딕셔너리를 조립한다.

        Args:
            row_type    : 'header' | 'data' | 'last'
            content_type: 'shrink' | 'wrap' | 'number'
            is_first_col: 맨 왼쪽 컬럼 여부
            is_last_col : 맨 오른쪽 컬럼 여부
            column_name : 컬럼 이름 (테두리 판단용)

        Returns:
            dict: xlsxwriter add_format에 전달할 속성 딕셔너리
        """

        if row_type == 'header':
            align = 'center'
        elif content_type == 'number' or column_name in self.RIGHT_ALIGN_COLUMNS:
            align = 'right'
        elif column_name in self.LEFT_ALIGN_COLUMNS:
            align = 'left'
        else:
            align = 'center'

        props = {
            'align': align,
            'valign': 'vcenter',
        }

        # ── 오른쪽 정렬 텍스트 여백 (숫자는 num_format의 _ 로 처리) ──
        if row_type != 'header' and column_name in self.RIGHT_ALIGN_COLUMNS:
            props['indent'] = 1

        # ── 헤더 배경색 ──
        if row_type == 'header':
            props['bg_color'] = self.HEADER_BG_COLOR

        # ── 내용 표시 방식 (기존과 동일) ──
        if content_type == 'wrap':
            props['text_wrap'] = True
        elif content_type == 'number':
            props['num_format'] = '#,##0_ '
            props['shrink'] = True
        else:
            props['shrink'] = True

        # ── 테두리 (기존과 동일) ──
        border = {'top': 1, 'bottom': 1, 'left': 1, 'right': 1}

        if row_type == 'header':
            border['top'] = 2
        elif row_type == 'last':
            border['bottom'] = 2

        if is_first_col or column_name in self.LEFT_THICK_BORDER_COLUMNS:
            border['left'] = 2

        if is_last_col or column_name in self.RIGHT_THICK_BORDER_COLUMNS:
            border['right'] = 2

        props.update(border)
        return props

    def _get_content_type(self, column_name, row_type):
        """컬럼 이름과 행 타입으로 내용 표시 방식을 결정한다."""

        # 헤더행은 무조건 shrink
        if row_type == 'header':
            return 'shrink'

        # 줄바꿈 대상 컬럼
        if column_name in self.WRAP_COLUMNS:
            return 'wrap'

        # 숫자 서식 대상 컬럼
        if column_name in self.NUMBER_COLUMNS:
            return 'number'

        # 기본
        return 'shrink'

    def get_cell_format(self, column_name, col_index, row_type='data'):
        """
        컬럼 이름 + 행 타입으로 적절한 서식 객체를 반환한다.
        동일한 조합은 캐시에서 재사용.

        Args:
            column_name: 컬럼 이름
            col_index  : 전체 엑셀 기준 컬럼 인덱스
            row_type   : 'header' | 'data' | 'last'

        Returns:
            xlsxwriter.Format 객체
        """

        content_type = self._get_content_type(column_name, row_type)
        is_first = (col_index == 0)
        is_last = (col_index == self.COL_IDX_LAST)

        # 캐시 키 (같은 조건이면 같은 포맷 재사용)
        cache_key = (row_type, content_type, is_first, is_last, column_name)

        if cache_key not in self._format_cache:
            props = self._build_format_props(
                row_type, content_type, is_first, is_last, column_name
            )
            self._format_cache[cache_key] = self.workbook.add_format(props)

        return self._format_cache[cache_key]

    # ================================================================
    # 셀 쓰기 헬퍼
    # ================================================================

    def _write_cell(self, row, col_index, value, fmt, merge_bottom=None):
        """
        단일 셀 쓰기 또는 세로 병합 쓰기를 자동 판별한다.

        Args:
            row         : 시작 행 인덱스
            col_index   : 컬럼 인덱스
            value       : 셀에 넣을 값 (문자열, 숫자, 수식 등)
            fmt         : xlsxwriter.Format 객체
            merge_bottom: 병합 시 마지막 행 인덱스 (None이면 단일 셀)
        """

        if merge_bottom is not None and merge_bottom > row:
            # 2행 이상 병합
            self.worksheet.merge_range(
                row, col_index,
                merge_bottom, col_index,
                value, fmt
            )
        else:
            # 단일 셀
            self.worksheet.write(row, col_index, value, fmt)

    def _make_balance_formula(self, row):
        """잔액 수식 생성: 사업별예산액 - 총지출금액"""

        budget_cell = xl_rowcol_to_cell(row, self.COL_IDX_BALANCE - 1)  # 사업별예산액
        total_cell = xl_rowcol_to_cell(row, self.COL_IDX_TOTAL)         # 총지출금액
        return f'={budget_cell}-{total_cell}'

    def _make_total_formula(self, row, merge_bottom=None):
        """총지출금액 수식 생성: 집행 결의금액들의 합계"""

        # 집행 시작 컬럼 (결의금액이 첫 번째 집행 컬럼이라고 가정)
        end_row = merge_bottom if merge_bottom else row
        return f'=SUM({xl_range(row, self.COL_IDX_EXEC_START, end_row, self.COL_IDX_EXEC_START)})'

    # ================================================================
    # 헤더 작성
    # ================================================================

    def _write_header(self):
        """0행에 컬럼 헤더를 작성한다."""

        for col_index, col_name in enumerate(self.ALL_COLUMNS):
            fmt = self.get_cell_format(col_name, col_index, row_type='header')
            self.worksheet.write(0, col_index, col_name, fmt)

    # ================================================================
    # 데이터 작성 (핵심 메서드)
    # ================================================================

    def create_xl(self, data, last_element=False):
        """
        한 건의 예산·지출 데이터를 엑셀에 작성한다.

        Args:
            data        : {'컬럼명': 값, ..., '지출집행내역': [{'컬럼명': 값}, ...]}
            last_element: True이면 마지막 데이터 (아래 테두리 굵게)
        """

        executions = data.get('지출집행내역', [])
        exec_count = len(executions)

        # ── 병합 여부 결정 ──
        # 지출이 2건 이상이면 예산 컬럼들을 세로 병합
        merge_bottom = (
            self.write_idx + exec_count - 1
            if exec_count >= 2
            else None
        )

        # ── 행 타입 결정 ──
        row_type = 'last' if last_element else 'data'

        # ── 1) 예산 컬럼 작성 ──
        for col_index, col_name in enumerate(self.BUDGET_COLUMNS):
            fmt = self.get_cell_format(col_name, col_index, row_type)
            self._write_cell(
                self.write_idx, col_index,
                data.get(col_name), fmt, merge_bottom
            )

        # ── 2) 잔액 작성 (수식) ──
        fmt_balance = self.get_cell_format('잔액', self.COL_IDX_BALANCE, row_type)
        self._write_cell(
            self.write_idx, self.COL_IDX_BALANCE,
            self._make_balance_formula(self.write_idx),
            fmt_balance, merge_bottom
        )

        # ── 3) 총지출금액 작성 (수식) ──
        fmt_total = self.get_cell_format('총지출금액', self.COL_IDX_TOTAL, row_type)
        self._write_cell(
            self.write_idx, self.COL_IDX_TOTAL,
            self._make_total_formula(self.write_idx, merge_bottom),
            fmt_total, merge_bottom
        )

        # ── 4) 지출 집행 내역 작성 ──
        if executions:
            for exec_idx, execution_data in enumerate(executions):
                is_last_row = last_element and (exec_idx == exec_count - 1)
                current_row_type = 'last' if is_last_row else 'data'

                for col_offset, col_name in enumerate(self.EXECUTION_COLUMNS):
                    abs_col = self.COL_IDX_EXEC_START + col_offset
                    fmt = self.get_cell_format(col_name, abs_col, current_row_type)
                    self.worksheet.write(
                        self.write_idx, abs_col,
                        execution_data.get(col_name), fmt
                    )

                self.write_idx += 1
        else:
            # 지출 0건이어도 빈 셀에 서식 적용 (테두리·정렬 유지)
            for col_offset, col_name in enumerate(self.EXECUTION_COLUMNS):
                abs_col = self.COL_IDX_EXEC_START + col_offset
                fmt = self.get_cell_format(col_name, abs_col, row_type)
                self.worksheet.write(self.write_idx, abs_col, '', fmt)

            self.write_idx += 1

    # ================================================================
    # 파일 닫기
    # ================================================================

    def close(self):
        """엑셀 파일을 저장하고 닫는다."""
        # 숫자가 텍스트로 저장된 셀의 경고 삼각형 숨김
        self.worksheet.ignore_errors({
            'number_stored_as_text': f'A1:XFD{self.write_idx}'
        })
        self.workbook.close()
