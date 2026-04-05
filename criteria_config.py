# criteria_config.py
# 모든 유형·계정·연산자 정의의 단일 출처(single source of truth)

NUMERIC_OPERATORS = {"초과": "gt", "이상": "gte", "미만": "lt", "이하": "lte", "같음": "eq"}
TEXT_OPERATORS = {
    "텍스트 일치": "equals",
    "All equals": "all_equals",
    "텍스트 포함": "contains",
    "공란": "blank",
    "공란아님": "not_blank",
}

COUNT_REQUIREMENT_OPTIONS = ["모든연도", "1개년이라도", "N개년이상"]
COUNT_REQUIREMENT_MAPPING = {"모든연도": "all", "1개년이라도": "any"}  # "N개년이상" → int(n)

INCLUDE_OPTIONS = ["포함", "제외"]
INCLUDE_MAPPING = {"포함": True, "제외": False}

CRITERIA_TYPES = {
    "텍스트": {
        "accounts": {
            "감사의견": "Audit status\n",
            "상장여부": "Listing status",
            "회사상태": "Status",
            "주요활동": "Main activity",
            "주요사업": "Primary business line",
            "웹사이트": "Website address",
            "SIC코드": "US SIC, primary code(s)",
        },
        "operators": list(TEXT_OPERATORS.keys()),
        "has_value": True,
        "has_year_cond": False,
        "value_hint": "예: Unqualified",
    },
    "숫자-개별연도": {
        "accounts": {
            "매출액(Turnover)": "Operating revenue (Turnover)\nth USD ",
            "매출액(Sales)": "Sales\nth USD ",
            "매출원가": "Costs of goods sold\nth USD ",
            "매출총이익": "Gross profit\nth USD ",
            "영업비용": "Other operating expense (income)\nth USD ",
            "영업이익(EBIT)": "Operating profit (loss) [EBIT]\nth USD ",
            "연구개발비": "Research & Development expenses\nth USD ",
            "직원수": "Number of employees\n",
            "총자산": "Total assets\nth USD ",
            "매출채권": "Debtors\nth USD ",
            "매입채무": "Creditors\nth USD ",
            "재고자산": "Stock\nth USD ",
            "무형자산": "Intangible assets\nth USD ",
            "유형자산": "Tangible fixed assets\nth USD ",
        },
        "operators": list(NUMERIC_OPERATORS.keys()),
        "has_value": True,
        "has_year_cond": True,
        "value_hint": "예: 0",
    },
    "숫자-WA3평균": {
        "accounts": {
            "매출액": "매출액",
            "영업이익": "영업이익",
            "영업비용": "영업비용",
            "재고자산": "재고자산",
            "연구개발비": "연구개발비",
            "무형자산": "무형자산",
            "유형자산": "유형자산",
            "총자산": "총자산",
            "매출원가": "매출원가",
            "종업원수": "종업원수",
        },
        "operators": list(NUMERIC_OPERATORS.keys()),
        "has_value": True,
        "has_year_cond": False,
        "value_hint": "예: 0",
    },
    "비율": {
        "accounts": {
            "연구개발비/매출액": "연구개발비/매출액",
            "영업비용/매출액": "영업비용/매출액",
            "무형자산/총자산": "무형자산/총자산",
            "유형자산/총자산": "유형자산/총자산",
            "재고자산/총자산": "재고자산/총자산",
            "재고자산보유일수\n(365/재고자산회전율)": "재고자산보유일수",
        },
        "operators": list(NUMERIC_OPERATORS.keys()),
        "has_value": True,
        "has_year_cond": False,
        "value_hint": "소수 (예: 0.01 = 1%)",
    },
    "데이터가용성": {
        "accounts": {
            "재무정보가용성": [
                "Operating revenue (Turnover)\nth USD ",
                "Gross profit\nth USD ",
                "Operating profit (loss) [EBIT]\nth USD ",
            ]
        },
        "operators": ["존재함"],
        "has_value": False,
        "has_year_cond": False,
        "value_hint": "",
    },
}

TYPE_DISPLAY_NAMES = list(CRITERIA_TYPES.keys())

# preset명 → (type_key, account_korean) 매핑 (apply_preset용)
PRESET_TO_TYPE_ACCOUNT = {
    "감사의견":                                   ("텍스트",        "감사의견"),
    "상장여부":                                   ("텍스트",        "상장여부"),
    "재무정보가용성":                             ("데이터가용성",  "재무정보가용성"),
    "매출액(금액, 평균)":                         ("숫자-WA3평균",  "매출액"),
    "영업이익(금액, 평균)":                       ("숫자-WA3평균",  "영업이익"),
    "영업비용(금액, 평균)":                       ("숫자-WA3평균",  "영업비용"),
    "재고자산(금액, 평균)":                       ("숫자-WA3평균",  "재고자산"),
    "연구개발비(금액, 평균)":                     ("숫자-WA3평균",  "연구개발비"),
    "무형자산(금액, 평균)":                       ("숫자-WA3평균",  "무형자산"),
    "유형자산(금액, 평균)":                       ("숫자-WA3평균",  "유형자산"),
    "총자산(금액, 평균)":                         ("숫자-WA3평균",  "총자산"),
    "매출원가(금액, 평균)":                       ("숫자-WA3평균",  "매출원가"),
    "영업이익(금액, 1개년이라도)":                ("숫자-개별연도", "영업이익(EBIT)"),
    "영업이익(금액, 연속)":                       ("숫자-개별연도", "영업이익(EBIT)"),
    "연구개발비/매출액(비율, 평균)":              ("비율",          "연구개발비/매출액"),
    "영업비용/매출액(비율, 평균)":                ("비율",          "영업비용/매출액"),
    "무형자산/총자산(비율, 평균)":                ("비율",          "무형자산/총자산"),
    "유형자산/총자산(비율, 평균)":                ("비율",          "유형자산/총자산"),
    "재고자산/총자산(비율, 평균)":                ("비율",          "재고자산/총자산"),
    "재고자산보유일수(365/재고자산회전율)(평균)": ("비율",          "재고자산보유일수\n(365/재고자산회전율)"),
}
