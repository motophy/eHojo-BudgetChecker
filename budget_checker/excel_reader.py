def get_unique_items(sheet, column_names: tuple, start_row: int = 3):
    """
    엑셀 파일에서 지정한 여러 열(컬럼)의 값 조합을
    파일에 등장한 순서 그대로 중복 없이 추출하는 함수

    [매개변수]
    sheet        : 읽을 엑셀 시트 객체
    column_names : 조회할 열 이름 튜플           예) ("city", "category")
    start_row    : 데이터 읽기 시작할 엑셀 행 번호 (기본값 2 = 엑셀 기준 2번째 행)
                   예) 1행 헤더 + 2행부터 데이터면 → 2
                       3행까지 헤더/설명이고 4행부터 데이터면 → 4

    [반환값]
    파일 위→아래 등장 순서가 유지된 중복 제거 튜플
    예) (("Seoul", "A"), ("Busan", "B"), ("Seoul", "B"))
    """

    # 중복 체크용 집합(set) — 순서 보장 안 됨, 빠른 조회 전용
    seen = set()

    # 순서 유지용 리스트 — 처음 등장한 순서대로 값을 보관
    ordered = []

    # 시트 객체를 ws 변수에 할당
    ws = sheet

    # 0번 행(헤더행)에서 모든 열 이름을 리스트로 수집
    # ws.ncols = 총 열 개수
    # 결과 예: ["name", "city", "category", "value"]
    headers = [ws.cell_value(0, c) for c in range(ws.ncols)]

    # column_names의 각 열 이름이 headers에서 몇 번째인지 인덱스로 변환
    # 결과 예: ("city", "category") → (1, 2)
    col_indexes = tuple(headers.index(name) for name in column_names)

    # start_row는 엑셀 기준 행 번호 (1부터 시작)
    # 파이썬 인덱스는 0부터 시작하므로 -1 변환 필요
    # 예) start_row=2 → range(1, ws.nrows) → 엑셀 2행부터 읽음
    # 예) start_row=4 → range(3, ws.nrows) → 엑셀 4행부터 읽음
    for r in range(start_row - 1, ws.nrows):

        # 지정된 열들의 값만 골라서 하나의 튜플로 묶음
        # 예: city="Seoul", category="A" 이면 → ("Seoul", "A")
        val = tuple(ws.cell_value(r, c) for c in col_indexes)

        # 이 조합이 처음 등장했을 때만 처리
        if val not in seen:
            seen.add(val)        # 중복 체크용 set에 추가
            ordered.append(val)  # 순서 유지용 list에 추가

    # 등장 순서가 유지된 리스트를 tuple로 변환하여 반환
    return tuple(ordered)


def get_joined_text(sheet, column_names: tuple, match_values: tuple, target_column: str, start_row: int = 3):
    """
    특정 열 조합이 일치하는 행들을 찾아서
    지정한 열의 값을 줄바꿈(\n)으로 합쳐 반환하는 함수

    [매개변수]
    sheet        : 읽을 엑셀 시트 객체
    column_names : 조건으로 사용할 열 이름 튜플       예) ("city", "category")
    match_values : 찾을 값 조합 튜플                  예) ("Seoul", "A")
    target_column: 값을 합칠 대상 열 이름             예) "name"
    start_row    : 데이터 읽기 시작 엑셀 행 번호 (기본값 2)

    [반환값]
    조건에 맞는 행들의 target_column 값을 \n 으로 합친 문자열
    예) "홍길동\n김철수\n이영희"
    """

    # 시트 객체 할당
    ws = sheet

    # 0번 행(헤더행)에서 모든 열 이름을 리스트로 수집
    # 결과 예: ["name", "city", "category", "value"]
    headers = [ws.cell_value(0, c) for c in range(ws.ncols)]

    # 조건 열들의 인덱스 변환
    # 결과 예: ("city", "category") → (1, 2)
    col_indexes = tuple(headers.index(name) for name in column_names)

    # 값을 합칠 대상 열의 인덱스 찾기
    # 결과 예: "name" → 0
    target_index = headers.index(target_column)

    # 조건에 맞는 행의 target_column 값을 담을 리스트
    result = []

    # start_row 기준으로 엑셀 행 순회
    for r in range(start_row - 1, ws.nrows):

        # 조건 열들의 값을 튜플로 묶음
        # 예: ("Seoul", "A")
        val = tuple(ws.cell_value(r, c) for c in col_indexes)

        # match_values와 일치하는 행일 때만 처리
        if val == match_values:

            # 대상 열의 값을 문자열로 변환하여 리스트에 추가
            # str() 로 감싸는 이유 : 숫자, None 등 문자열이 아닌 값도 대비
            raw_data = ws.cell_value(r, target_index)
            if type(raw_data) == float:
                formatted_value = f"{raw_data:,.0f}"
                result.append(formatted_value)

            else:
                result.append(str(ws.cell_value(r, target_index)))

    # 리스트의 항목들을 \n 으로 연결하여 하나의 문자열로 반환
    return "\n".join(result)


def get_row_values(sheet, column_names: tuple, match_values: tuple, target_columns: tuple, start_row: int = 3):
    """
    조건에 일치하는 첫 번째 행에서 지정한 열들의 값을 딕셔너리로 반환하는 함수

    [매개변수]
    sheet          : 읽을 엑셀 시트 객체
    column_names   : 조건으로 사용할 열 이름 튜플       예) ("city", "category")
    match_values   : 찾을 값 조합 튜플                  예) ("Seoul", "A")
    target_columns : 가져올 열 이름 튜플                예) ("name", "value")
    start_row      : 데이터 읽기 시작 엑셀 행 번호 (기본값 2)

    [반환값]
    조건에 맞는 첫 번째 행의 {열이름: 값} 딕셔너리
    예) {"name": "홍길동", "value": 42.0}
    일치하는 행이 없으면 None 반환
    """

    # 시트 객체 할당
    ws = sheet

    # 0번 행(헤더행)에서 모든 열 이름을 리스트로 수집
    # 결과 예: ["name", "city", "category", "value"]
    headers = [ws.cell_value(0, c) for c in range(ws.ncols)]

    # 조건 열들의 인덱스 변환
    # 결과 예: ("city", "category") → (1, 2)
    col_indexes = tuple(headers.index(name) for name in column_names)

    # 가져올 대상 열들의 인덱스 변환
    # 결과 예: ("name", "value") → (0, 3)
    target_indexes = tuple(headers.index(name) for name in target_columns)

    # start_row 기준으로 엑셀 행 순회
    for r in range(start_row - 1, ws.nrows):

        # 조건 열들의 값을 튜플로 묶음
        # 예: ("Seoul", "A")
        val = tuple(ws.cell_value(r, c) for c in col_indexes)

        # match_values와 일치하는 첫 번째 행 발견 시
        if val == match_values:

            # {열이름: 셀값} 형태의 딕셔너리로 반환
            # 결과 예: {"name": "홍길동", "value": 42.0}
            return {target_columns[i]: ws.cell_value(r, target_indexes[i]) for i in range(len(target_columns))}

    # 일치하는 행이 하나도 없으면 None 반환
    return None


def get_sum_value(sheet, column_names: tuple, match_values: tuple, target_column: str, start_row: int = 3):
    """
    조건에 일치하는 행들의 특정 열 값(천단위 , 가 포함된 텍스트 숫자)을 합산하는 함수

    [매개변수]
    sheet        : 읽을 엑셀 시트 객체
    column_names : 조건으로 사용할 열 이름 튜플       예) ("city", "category")
    match_values : 찾을 값 조합 튜플                  예) ("Seoul", "A")
    target_column: 합산할 열 이름                     예) "sales"
    start_row    : 데이터 읽기 시작 엑셀 행 번호 (기본값 3)

    [반환값]
    조건에 맞는 행들의 target_column 합산 결과 (int)
    예) 1,000 + 2,500 + 3,000 → 6500
    일치하는 행이 없으면 0 반환
    """

    # 시트 객체 할당
    ws = sheet

    # 0번 행(헤더행)에서 모든 열 이름을 리스트로 수집
    # 결과 예: ["name", "city", "category", "sales"]
    headers = [ws.cell_value(0, c) for c in range(ws.ncols)]

    # 조건 열들의 인덱스 변환
    # 결과 예: ("city", "category") → (1, 2)
    col_indexes = tuple(headers.index(name) for name in column_names)

    # 합산할 대상 열의 인덱스 찾기
    # 결과 예: "sales" → 3
    target_index = headers.index(target_column)

    # 합산 결과를 누적할 변수
    total = 0

    # start_row 기준으로 엑셀 행 순회
    for r in range(start_row - 1, ws.nrows):

        # 조건 열들의 값을 튜플로 묶음
        # 예: ("Seoul", "A")
        val = tuple(ws.cell_value(r, c) for c in col_indexes)

        # match_values와 일치하는 행일 때만 처리
        if val == match_values:

            # 대상 열의 셀 값을 문자열로 변환
            # 예: "1,234,567"
            raw = str(ws.cell_value(r, target_index))

            # 천단위 구분자 , 를 제거하고 숫자로 변환
            # "1,234,567" → "1234567" → 1234567
            # strip() : 앞뒤 공백 제거 (셀에 공백이 있을 수 있으므로)
            cleaned = raw.replace(",", "").strip()

            # 빈 문자열이거나 숫자로 변환 불가한 값은 건너뜀
            if cleaned == "" or not cleaned.replace(".", "").lstrip("-").isdigit():
                continue

            # 소수점이 있으면 float, 없으면 int로 변환 후 누적
            total += float(cleaned) if "." in cleaned else int(cleaned)

    return total


def get_rows_sorted(sheet, column_names: tuple, match_values: tuple, target_columns: tuple, start_row: int = 2):
    """
    조건에 일치하는 모든 행에서 지정한 열들의 값을
    딕셔너리 튜플로 반환하는 함수 ('지급일자' 기준 오름차순 정렬)

    [매개변수]
    sheet          : 읽을 엑셀 시트 객체
    column_names   : 조건으로 사용할 열 이름 튜플       예) ("거래처명",)
    match_values   : 찾을 값 조합 튜플                  예) ("홍길동상사",)
    target_columns : 가져올 열 이름 튜플                예) ("거래처명", "지급일자", "금액")
    start_row      : 데이터 읽기 시작 엑셀 행 번호 (기본값 2)

    [반환값]
    조건에 맞는 모든 행의 딕셔너리를 '지급일자' 기준 오름차순 정렬한 튜플
    예) (
            {"거래처명": "홍길동상사", "지급일자": "2024-01-01", "금액": 1000},
            {"거래처명": "홍길동상사", "지급일자": "2024-02-01", "금액": 2000},
        )
    일치하는 행이 없으면 빈 튜플 () 반환
    """

    # 시트 객체 할당
    ws = sheet

    # 0번 행(헤더행)에서 모든 열 이름을 리스트로 수집
    # 결과 예: ["거래처명", "지급일자", "금액"]
    headers = [ws.cell_value(0, c) for c in range(ws.ncols)]

    # 조건 열들의 인덱스 변환
    # 결과 예: ("거래처명",) → (0,)
    col_indexes = tuple(headers.index(name) for name in column_names)

    # 가져올 대상 열들의 인덱스 변환
    # 결과 예: ("거래처명", "지급일자", "금액") → (0, 1, 2)
    target_indexes = tuple(headers.index(name) for name in target_columns)

    # 조건에 맞는 행들의 딕셔너리를 담을 리스트
    # → 마지막 하나만 남기지 않고 전부 담음
    result = []

    # start_row 기준으로 엑셀 행 순회
    for r in range(start_row - 1, ws.nrows):

        # 조건 열들의 값을 튜플로 묶음
        # 예: ("홍길동상사",)
        val = tuple(ws.cell_value(r, c) for c in col_indexes)

        # match_values와 일치하는 행일 때마다 리스트에 추가
        # (덮어쓰지 않고 append → 전체 수집)
        if val == match_values:
            row_dict = {target_columns[i]: ws.cell_value(r, target_indexes[i]) for i in range(len(target_columns))}
            result.append(row_dict)

    # '지급일자' 열 기준 오름차순 정렬
    # target_columns 에 '지급일자' 가 없으면 KeyError 발생 → 반드시 포함해야 함
    result.sort(key=lambda row: row["지급일자"])

    # 정렬된 리스트를 튜플로 변환하여 반환
    return tuple(result)
