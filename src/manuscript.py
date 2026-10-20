"""Fill the JRM Word template with the comparison manuscript.
usage: python manuscript.py <unpacked_template_dir>
"""
import re, sys
from xml.sax.saxutils import escape

D = sys.argv[1]
doc_path = f"{D}/word/document.xml"
fn_path = f"{D}/word/footnotes.xml"
x = open(doc_path, encoding="utf8").read()

# ---------------------------------------------------------------- helpers
def run(text, b=False, sz=None, sup=False, color=None, font=None):
    rpr = ""
    if font: rpr += f'<w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}"/>'
    if b: rpr += "<w:b/>"
    if color: rpr += f'<w:color w:val="{color}"/>'
    rpr += '<w:position w:val="0"/>'
    if sz: rpr += f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>'
    if sup: rpr += '<w:vertAlign w:val="superscript"/>'
    rpr = f"<w:rPr>{rpr}</w:rPr>" if rpr else ""
    return f'<w:r>{rpr}<w:t xml:space="preserve">{escape(text)}</w:t></w:r>'

FN = {}  # id -> text
def fnref(i):
    return f'<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="{i}"/></w:r>'

def rich(s, sz=20):
    """'text [[fn:2]] more' -> runs; **bold** supported."""
    out = []
    for part in re.split(r"(\[\[fn:\d+\]\]|\*\*[^*]+\*\*)", s):
        if not part: continue
        m = re.match(r"\[\[fn:(\d+)\]\]", part)
        if m: out.append(fnref(m.group(1))); continue
        if part.startswith("**"): out.append(run(part[2:-2], b=True, sz=sz)); continue
        out.append(run(part, sz=sz))
    return "".join(out)

def para(content, style="1", extra_ppr=""):
    return f'<w:p><w:pPr><w:pStyle w:val="{style}"/>{extra_ppr}</w:pPr>{content}</w:p>'

def body(s): return para(rich(s), "1", '<w:ind w:firstLine="200"/><w:rPr><w:position w:val="0"/></w:rPr>')
def blank(): return para("", "a3")
KN = '<w:keepNext/>'
def kblank(): return para("", "a3", KN)
def h1(s): return para(run(s), "a4", KN) + kblank()
def h2(s): return kblank() + para(run(s), "11", KN + '<w:ind w:left="200" w:firstLine="0"/>') + kblank()
def h3(s): return para(run(s), "110")
def caption(s):
    return para(run(s, b=True, sz=18), "1", '<w:keepNext/><w:spacing w:before="160" w:after="80"/><w:ind w:firstLine="0"/><w:jc w:val="center"/>')
def note(s):
    return para(run(s, sz=15), "1", '<w:spacing w:before="40" w:after="160" w:line="260" w:lineRule="auto"/><w:ind w:firstLine="0"/>')

TOTAL = 7031  # text width in twips
def tcell(text, w, header=False, first=False, fill=None, keep=True):
    if header: fill = "E5E5E5"
    shd = f'<w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>' if fill else ""
    lines = text.split("\n")
    ps = "".join(
        f'<w:p><w:pPr>{"<w:keepNext/>" if keep else ""}<w:spacing w:line="250" w:lineRule="auto"/><w:jc w:val="{"center" if (header or first) else "left"}"/></w:pPr>'
        f'{run(l, b=(header or first), sz=15)}</w:p>' for l in lines)
    return (f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/>{shd}'
            f'<w:tcMar><w:top w:w="40" w:type="dxa"/><w:left w:w="70" w:type="dxa"/><w:bottom w:w="40" w:type="dxa"/><w:right w:w="70" w:type="dxa"/></w:tcMar>'
            f'<w:vAlign w:val="center"/></w:tcPr>{ps}</w:tc>')

def table(widths, header, rows, fills=None):
    assert sum(widths) == TOTAL, sum(widths)
    grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in widths)
    b = ('<w:tblBorders><w:top w:val="single" w:sz="12" w:space="0" w:color="000000"/>'
         '<w:bottom w:val="single" w:sz="12" w:space="0" w:color="000000"/>'
         '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="808080"/>'
         '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="808080"/></w:tblBorders>')
    out = (f'<w:tbl><w:tblPr><w:tblW w:w="{TOTAL}" w:type="dxa"/><w:jc w:val="center"/>{b}'
           f'<w:tblLayout w:type="fixed"/><w:tblLook w:val="0000"/></w:tblPr><w:tblGrid>{grid}</w:tblGrid>')
    out += '<w:tr><w:trPr><w:tblHeader/><w:cantSplit/></w:trPr>' + "".join(tcell(h, widths[i], header=True) for i, h in enumerate(header)) + "</w:tr>"
    for r_i, r in enumerate(rows):
        out += "<w:tr><w:trPr><w:cantSplit/></w:trPr>"
        last = r_i == len(rows) - 1
        for i, c in enumerate(r):
            f = fills.get((r_i, i)) if fills else None
            out += tcell(c, widths[i], first=(i == 0), fill=f, keep=not last)
        out += "</w:tr>"
    return out + "</w:tbl>"

def pagebreak(): return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'

# ---------------------------------------------------------------- front matter
TITLE_KO = "다변화하는 위험 환경과 보험계리사 자격시험의 데이터사이언스 역량 평가: 한국·미국·캐나다·영국 시험제도 비교"
TITLE_EN = "Assessing Data Science Competence in Actuarial Qualification: A Comparison of Korea, the United States, Canada and the United Kingdom"

ABSTRACT_KO = ("본 연구는 다변화하는 위험과 고도화하는 위험관리 기법에 대응하기 위하여 한국 보험계리사 시험에 통계 및 데이터사이언스 관련 "
    "평가를 확대할 필요가 있는지를 검토한다. 미국계리사협회(SOA)의 자격 경로를 기준으로 데이터사이언스 역량을 기초이론, 기초실습, "
    "과제형 심화실습의 세 단계로 구분하고, 한국 시험과 SOA, 미국손해보험계리사협회(CAS), 캐나다계리사협회(CIA), 영국계리사협회(IFoA)의 "
    "시험제도를 공개 자료에 근거하여 비교하였다. 해외 4개 기관은 평가 방식은 다르지만 모두 통계적 학습 이론을 자격 요건에 포함하고 "
    "실습형 평가를 운영하는 반면, 한국 시험에는 이에 대응하는 평가 요소가 확인되지 않았다. 허용 계산기 정책에서도 해외 기관은 "
    "공학용·재무용 계산기나 스프레드시트를 허용하는 데 비해 한국 시험은 단순계산기만 허용하였다. 이를 바탕으로 출제범위의 단계적 확대, "
    "실습형 평가의 도입, 계산 도구 정책의 정비 방안을 논의한다.")
KW_KO = "보험계리사 시험, 데이터사이언스, 예측분석, 계리사 교육, 자격제도 비교, 계산기 정책"

ABSTRACT_EN = ("This study examines whether the Korean actuarial examination should place greater weight on statistics and data science "
    "in response to diversifying risks and increasingly sophisticated risk management techniques. Using the current Associateship "
    "pathway of the Society of Actuaries (SOA) as a reference, we classify data science competence into three tiers: foundational "
    "theory, foundational practice, and advanced project-based practice. We then compare the Korean examination with the "
    "qualification systems of the SOA, the Casualty Actuarial Society (CAS), the Canadian Institute of Actuaries (CIA) and the "
    "Institute and Faculty of Actuaries (IFoA), based on the published materials of each body. All four overseas bodies include "
    "statistical learning theory in their requirements and operate practical assessments that use software and data, although their "
    "assessment methods differ. No explicit counterpart was found in the Korean examination. A comparison of calculator policies shows a "
    "similar gap: the overseas bodies permit scientific and financial calculators or spreadsheets, whereas the Korean examination allows only "
    "simple calculators without data storage. Based on these findings, we discuss a phased expansion of the syllabus, the introduction of "
    "practical assessments and a revision of the calculator policy.")
KW_EN = "Actuarial Examination, Data Science, Predictive Analytics, Actuarial Education, Comparative Study of Qualifications, Calculator Policy"

# title cell
x = re.sub(r'(<w:pStyle w:val="ad"/>.*?</w:pPr>).*?(</w:p>)', lambda m: m.group(1) + run(TITLE_KO) + m.group(2), x, count=1, flags=re.S)
x = re.sub(r'(<w:pStyle w:val="ae"/><w:spacing w:line="240" w:lineRule="auto"/></w:pPr>).*?(</w:p>)', lambda m: m.group(1) + run(TITLE_EN, sz=22) + m.group(2), x, count=1, flags=re.S)
# author rows (placeholders, to be filled by the author on the cover page)
x = re.sub(r'(<w:pStyle w:val="ab"/>.*?</w:pPr>)(</w:p>)', lambda m: m.group(1) + run("[저자명]*") + m.group(2), x, count=1, flags=re.S)
x = re.sub(r'(<w:pStyle w:val="ac"/>.*?</w:pPr>)(</w:p>)', lambda m: m.group(1) + run("([Author Name]*)") + m.group(2), x, count=1, flags=re.S)
# abstract box (Korean abstract at the front)
x = re.sub(r'(<w:pStyle w:val="a8"/>.*?</w:pPr>).*?(</w:p>)', lambda m: m.group(1) + run(ABSTRACT_KO) + m.group(2), x, count=1, flags=re.S)
# keywords table
i = x.find('<w:t>K</w:t>')
seg_start = x.rfind('<w:p ', 0, i); seg_end = x.find('</w:p>', i) + 6
p_kw = x[seg_start:seg_end]
p_kw_new = re.sub(r'(</w:pPr>).*', r'\1' + run("주제어:", font="KoPub돋움체 Medium") + '</w:p>', p_kw, flags=re.S)
x = x[:seg_start] + p_kw_new + x[seg_end:]
x = x.replace('<w:rPr><w:color w:val="FFFFFF"/></w:rPr></w:pPr><w:r><w:rPr><w:color w:val="FFFFFF"/></w:rPr><w:footnoteReference w:id="1"/></w:r>',
              '</w:pPr>' + run(KW_KO) + '<w:r><w:rPr><w:color w:val="FFFFFF"/></w:rPr><w:footnoteReference w:id="1"/></w:r>', 1)

# ---------------------------------------------------------------- body
B = []
B.append(h1("Ⅰ. 서론"))
B.append(body("보험회사가 관리해야 하는 위험은 기후 변화, 인구 고령화, 사이버 위험 등으로 다변화하고 있으며, 이를 측정하고 관리하는 기법 역시 일반화선형모형(GLM)을 넘어 트리 기반 앙상블, 비지도학습 등 통계적 학습 방법으로 확장되고 있다. 이에 따라 주요국 계리사 단체는 최근 수년간 자격시험에서 통계 및 데이터사이언스 역량을 평가하는 비중을 꾸준히 늘려 왔다. 미국계리사협회(Society of Actuaries, 이하 SOA)는 위험모형 통계(SRM), 예측분석(PA), 예측분석 심화(ATPA)로 이어지는 평가 체계를 갖추었고, 미국손해보험계리사협회(Casualty Actuarial Society, 이하 CAS)는 2026년부터 예측분석 시험과 과제형 프로젝트를 준회원 자격의 필수 요건으로 지정하였다(CAS, 2024)."))
B.append(body("국내에서는 IFRS 17과 신지급여력제도(K-ICS)의 시행 이후 보험계리사 수요가 크게 늘었으나, 보험연구원은 현행 보험계리사 시험이 실무와 충분히 연계되지 못하고 있다고 지적한 바 있다(CEO스코어데일리, 2026). 그러나 국내 시험이 데이터사이언스 역량을 어느 정도 평가하고 있는지를 해외 제도와 체계적으로 비교한 자료는 많지 않다."))
B.append(body("본 연구는 SOA의 평가 체계를 기준으로 데이터사이언스 역량을 기초이론, 기초실습, 과제형 심화실습의 세 단계로 구분하고, 한국 보험계리사 시험과 SOA, CAS, 캐나다계리사협회(Canadian Institute of Actuaries, 이하 CIA), 영국계리사협회(Institute and Faculty of Actuaries, 이하 IFoA)의 시험제도를 이 틀에 따라 비교한다. 이를 통해 한국 시험제도의 공백을 확인하고 개선 방향을 제시하는 것이 본 연구의 목적이다. 아울러 수험자가 사용할 수 있는 계산 도구는 출제 가능한 문항의 범위와 직결되므로, 시험 운영 측면에서 각 기관의 허용 계산기 정책도 함께 비교한다. 분석은 2026년 10월 현재 각 기관이 공개한 시험 안내, 출제범위, 자격 경로 자료에 근거하였다."))
B.append(body("본 논문의 구성은 다음과 같다. Ⅱ장에서는 비교 틀을 제시하고, Ⅲ장에서는 기관별 시험제도를 검토한다. Ⅳ장에서는 비교 결과를 종합하고 시사점을 도출하며, Ⅴ장에서는 허용 계산기 정책을 비교한다. Ⅵ장에서 결론과 연구의 한계를 서술한다."))

B.append(blank())
B.append(h1("Ⅱ. 비교 틀"))
B.append(body("SOA의 준회원(ASA) 경로는 데이터사이언스 역량을 서로 다른 형식의 세 가지 평가로 나누어 검증한다. 위험모형 통계(Statistics for Risk Modeling, SRM) 시험은 회귀 및 시계열 모형, 주성분분석, 의사결정나무, 군집분석과 모형 선택 및 검증 방법을 평가한다(SOA, 2026c). 예측분석(Predictive Analytics, PA) 시험은 e-러닝 모듈과 시험으로 구성되며, 주어진 데이터셋을 이용하여 비즈니스 문제를 해결하는 능력을 평가한다(SOA, 2026b). 예측분석 심화(Advanced Topics in Predictive Analytics, ATPA) 평가는 데이터사이언스와 예측분석 기법으로 문제를 해결하고 그 결과를 효과적으로 소통하는 능력을 다루며, 수리통계, 확률, SRM, PA의 내용을 선수지식으로 전제한다(SOA, 2026a)."))
B.append(body("본 연구는 이 구조를 <표 1>과 같이 일반화하여 비교의 기준으로 삼는다. 각 단계는 평가 형식으로 구분되며, 이론 시험만으로는 기초실습이나 과제형 심화실습 단계를 충족한 것으로 보지 않는다."))
B.append(caption("<표 1> 데이터사이언스 역량 평가의 비교 틀"))
B.append(table([1500, 1500, 4031], ["단계", "SOA 대표 평가", "평가 대상 역량"], [
    ["기초이론", "SRM", "통계적 학습 방법론(회귀·GLM, 트리, 주성분분석, 군집분석 등)과 모형 선택·검증에 대한 이론적 이해"],
    ["기초실습", "PA", "주어진 데이터셋과 소프트웨어를 활용하여 문제를 정의·분석하고 결과를 해석하는 능력"],
    ["과제형\n심화실습", "ATPA", "일정 기간 동안 실제와 유사한 과제를 수행하고 분석 결과를 보고서로 소통하는 능력"],
]))
B.append(note("자료: SOA(2026a; 2026b; 2026c)를 바탕으로 저자 작성."))

B.append(blank())
B.append(h1("Ⅲ. 기관별 시험제도 검토"))
B.append(h2("1. 한국 보험계리사 시험"))
B.append(body("한국 보험계리사 시험은 금융감독원이 시행하고 보험개발원이 실무를 담당하는 국가자격시험이다[[fn:2]]. 시험은 1차와 2차로 나뉘며, 1차 시험은 보험계약법·보험업법·근로자퇴직급여보장법, 경제학원론, 보험수학, 회계원리의 객관식 과목과 공인영어성적으로 대체하는 영어로 구성된다(금융감독원, 2021). 2차 시험은 계리리스크관리, 보험수리학, 연금수리학, 계리모형론, 재무관리 및 금융공학의 5과목이며, 1차 시험 합격 후 5년 이내에 각 과목에서 60점 이상을 얻으면 최종 합격하는 과목별 합격제로 운영된다(조세금융신문, 2025)."))
B.append(body("통계 관련 내용은 1차 보험수학의 확률·통계 기초와 2차 계리모형론에 포함되어 있다. 그러나 통계적 학습이나 예측모델링을 별도의 출제영역으로 명시한 과목은 확인되지 않으며, 모든 시험이 지필 형식으로 치러져 데이터와 소프트웨어를 직접 다루는 실습형 평가도 존재하지 않는다."))
B.append(caption("<표 2> 한국 보험계리사 시험의 과목 구성"))
B.append(table([1300, 3700, 2031], ["구분", "과목", "형식 및 합격 기준"], [
    ["1차 시험", "보험계약법·보험업법·근로자퇴직급여보장법, 경제학원론, 보험수학, 회계원리, 영어(공인영어성적 대체)", "객관식\n매 과목 40점 이상, 평균 60점 이상"],
    ["2차 시험", "계리리스크관리, 보험수리학, 연금수리학, 계리모형론, 재무관리 및 금융공학", "주관식\n과목별 60점 이상(과목별 합격제)"],
]))
B.append(note("자료: 금융감독원(2021), 조세금융신문(2025)을 바탕으로 저자 작성."))

B.append(h2("2. 미국계리사협회(SOA)"))
B.append(body("SOA는 세 단계를 모두 독립된 평가로 운영하는 유일한 기관이다. 기초이론에 해당하는 SRM은 객관식 컴퓨터 기반 시험으로, 회귀·시계열 모형과 주요 통계적 학습 기법을 다룬다(SOA, 2026c). 기초실습에 해당하는 PA는 e-러닝 모듈과 데이터셋 기반 시험을 결합한 형태이며(SOA, 2026b), 과제형 심화실습에 해당하는 ATPA는 e-러닝 모듈과 과제형 평가로 구성된다(SOA, 2026a). ATPA가 SRM과 PA의 내용을 선수지식으로 명시하고 있어, 세 평가는 이론에서 실습, 다시 과제 수행으로 이어지는 위계를 이룬다."))

B.append(h2("3. 미국손해보험계리사협회(CAS)"))
B.append(body("CAS에서 기초이론에 해당하는 시험은 현대계리통계(Modern Actuarial Statistics) I과 II(이하 MAS-I, MAS-II)이다. 두 시험은 2018년에 도입되었으며, MAS-I은 통계적 추론과 GLM을 포함한 응용모델링을, MAS-II는 신뢰도 이론과 함께 선형혼합모형, 베이지안 분석과 마르코프 연쇄 몬테카를로(MCMC), 통계적 학습을 다룬다(CAS, 2017). MAS-II 출제범위에서 통계적 학습의 비중은 40~50%이며, 교재로 James et al.(2021), Goldburd et al.(2020) 등을 사용한다(CAS, 2023)."))
B.append(body("실습 단계에는 손해보험 예측분석(Property and Casualty Predictive Analytics, 이하 PCPA)이 대응한다. PCPA는 시험과 과제형 프로젝트의 두 부분으로 이루어진 준회원(ACAS) 요건으로, 2026년 1월 1일부터 적용되었다(CAS, 2024; CAS, 2025). PCPA 시험은 연중 상시 응시할 수 있는 2시간 객관식 시험이고, 프로젝트는 정해진 기간 동안 수행하는 과제형 평가이며, 시험에 합격해야 프로젝트에 등록할 수 있다(actuary.info, 2026). 따라서 PCPA 시험은 형식상 SOA의 PA와 부분적으로만 대응하며, PCPA 프로젝트는 ATPA와 성격이 가깝다."))

B.append(h2("4. 캐나다계리사협회(CIA)"))
B.append(body("CIA는 2023년부터 자체 교육체계를 도입하면서 데이터사이언스를 독립 시험으로 두지 않고 대학 교육과 CIA의 교육 모듈에 포함하는 방식을 택하였다. CIA는 예측분석 교육에 데이터셋과 프로그래밍·시각화 도구가 필요하므로 지필 시험보다 대학 교육이 적합하다는 입장을 밝힌 바 있다(Insurance Portal, 2021)."))
B.append(body("기초이론은 CIA 인증 대학의 학위(Pathway 1)로 인정하거나, 다른 계리사 단체의 시험 크레딧으로 인정한다. Pathway 2에서 SOA 응시자는 SRM을, CAS 응시자는 MAS-I과 MAS-II를 보유해야 하며, SOA·CAS·IFoA 등 인정 단체의 준회원 자격 보유자는 그 자격 자체로 인정받는다(CIA, 2026a). 실습 단계는 모든 경로에 필수인 ACIA Module 2가 담당하며, 이 모듈에는 분량이 큰 예측분석 채점 과제가 포함된다(CIA, 2026b). SOA 응시자는 Exam PA를 제외한 준회원 요건을 인정받고 Module 2를 별도로 이수하므로(CIA, 2026c), Module 2는 사실상 PA를 대신하는 역할을 한다. 별도의 과제형 심화실습 과목은 없다[[fn:3]]."))

B.append(h2("5. 영국계리사협회(IFoA)"))
B.append(body("IFoA의 핵심원리(Core Principles) 과목 가운데 계리통계(CS1)와 위험모형 및 생존분석(CS2)이 기초이론에 해당한다. CS1은 통계적 추론, GLM을 포함한 회귀모형, 베이지안 통계의 기초를 다루고, CS2는 위험모형, 시계열 분석, 확률과정, 생존분석과 함께 머신러닝에 대한 개괄적 소개를 포함한다(IFoA, 2026a). 다만 SOA의 SRM이나 CAS의 MAS-II와 비교하면 통계적 학습 기법의 비중은 상대적으로 작다."))
B.append(body("IFoA의 특징은 이론 과목 안에 실습형 평가를 결합하였다는 점이다. CS1과 CS2는 각각 이론 중심의 A 시험(3시간 20분)과 R을 사용하는 문제해결형 B 시험(1시간 50분)으로 구성되며, 두 시험은 같은 회차에 응시해야 한다(IFoA, 2026a). 두 시험은 70대 30의 비율로 하나의 점수로 합산된다(IFoA, 2026b). 이와 별도로 핵심실무(Core Practices) 과목인 모델링 실무(CP2)는 주어진 시나리오와 데이터로 데이터 분석과 스프레드시트 모형 구축을 수행하고 감사 추적(audit trail)과 요약 보고서를 작성하게 하며, 문서화에 45%의 비중을 둔다(IFoA, 2025a). CP2는 시간제한이 있는 시험이지만 모형 구축, 문서화, 결과 소통을 함께 평가한다는 점에서 과제형 평가의 요소를 일부 갖고 있다. 한편 IFoA는 2020년 사우샘프턴 대학교와 함께 데이터사이언스 인증과정(Certificate in Data Science)을 개설하였으나, 이는 자격 요건과 별개인 독립 과정이다(University of Southampton, 2020). 2026년 출제범위 변경에서도 CS1, CS2, CP2의 데이터사이언스 관련 내용에는 변화가 없었다(IFoA, 2026c)."))

B.append(blank())
B.append(h1("Ⅳ. 비교 결과와 시사점"))
B.append(h2("1. 비교 결과"))
B.append(body("<표 3>은 Ⅲ장의 검토 결과를 비교 틀에 따라 정리한 것이다. 해외 4개 기관은 기초이론 단계를 모두 자격 요건에 포함하고 있으며, 그 방식은 독립 시험(SOA, CAS), 대학 교육 및 타 단체 크레딧 인정(CIA), 이론 과목 내 포함(IFoA)으로 나뉜다. 기초실습 단계 역시 4개 기관이 모두 운영하고 있으나 형식은 독립 시험(SOA의 PA, CAS의 PCPA 시험), 교육 모듈 내 과제(CIA의 Module 2), 이론 과목과 결합한 소프트웨어 기반 시험(IFoA의 CS1B·CS2B)으로 다양하다. 과제형 심화실습은 SOA와 CAS만 독립 평가로 운영하며, IFoA의 CP2는 이에 부분적으로 대응한다."))
B.append(caption("<표 3> 기관별 데이터사이언스 역량 평가 체계"))
GAP = "FBE5D6"
B.append(table([1000, 2011, 2010, 2010], ["기관", "기초이론", "기초실습", "과제형 심화실습"], [
    ["한국", "전용 과목 없음\n(보험수학·계리모형론에 전통적 계리통계 포함)", "없음\n(전 과목 지필시험)", "없음"],
    ["SOA", "SRM\n(객관식 CBT)", "PA\n(e-러닝 + 데이터셋 기반 시험)", "ATPA\n(e-러닝 + 과제형 평가)"],
    ["CAS", "MAS-I, MAS-II\n(통계적 학습 40~50%)", "PCPA 시험\n(2시간 객관식, 부분 대응)", "PCPA 프로젝트\n(기간형 과제)"],
    ["CIA", "인증 대학 학위 또는\nSRM·MAS 크레딧 인정", "ACIA Module 2\n(예측분석 채점 과제)", "독립 과목 없음"],
    ["IFoA", "CS1, CS2 A 시험\n(머신러닝은 개괄 수준)", "CS1B, CS2B\n(R 사용 시험, 30% 반영)", "CP2 모델링 실무\n(시간제 시험, 부분 대응)"],
], fills={(0, 1): GAP, (0, 2): GAP, (0, 3): GAP}))
B.append(note("주: 음영은 해당 단계에 대응하는 평가 요소가 확인되지 않은 부분임.\n자료: Ⅲ장에서 인용한 각 기관 자료를 바탕으로 저자 작성."))

B.append(h2("2. 시사점"))
B.append(body("첫째, 해외 4개 기관이 모두 통계적 학습 이론을 자격 요건에 포함하고 있다는 점에서, 한국 시험에도 이에 대응하는 출제영역을 명시할 필요가 있다. 가장 부담이 적은 방법은 2차 계리모형론 또는 1차 보험수학의 출제범위에 GLM, 트리 모형, 모형 검증 등을 추가하는 것이며, 이는 CAS가 기존 시험을 개편하여 MAS-I과 MAS-II에 통계적 학습을 반영한 방식과 유사하다."))
B.append(body("둘째, 실습형 평가는 형식이 다양하다는 점에서 국내 여건에 맞는 방식을 선택할 수 있다. 독립된 실습 시험을 신설하는 SOA·CAS 방식 외에도, 기존 과목에 소프트웨어 기반 문항을 결합하는 IFoA 방식이나 자격 취득 과정에 교육 모듈을 두는 CIA 방식이 있다. 특히 IFoA의 방식은 과목 수를 늘리지 않으면서 실습 역량을 평가할 수 있어 현행 5과목 체계를 유지하는 한국 2차 시험에 참고할 만하다."))
B.append(body("셋째, 해외 기관들은 개편을 단계적으로 진행하였다. CAS는 2018년 MAS 시험으로 이론을 먼저 반영하고 2026년에 PCPA를 필수화하였으며, CIA는 기존 시험 크레딧을 인정하면서 자체 모듈로 실습을 보완하였다. 한국 역시 출제범위 확대를 우선 추진하고, 실습형 평가는 시험 운영 여건을 갖춘 뒤 도입하는 단계적 접근을 검토할 수 있다."))

B.append(blank())
B.append(h1("Ⅴ. 시험 운영 측면의 비교: 허용 계산기 정책"))
B.append(body("데이터사이언스 역량 평가를 확대하려면 출제범위뿐 아니라 수험자가 사용할 수 있는 계산 도구도 함께 고려해야 한다. 계산 도구가 단순계산기로 제한되면 반복 계산이나 행렬 연산이 필요한 문항을 출제하기 어렵고, 실습형 평가의 도입도 제약을 받는다. 이 장에서는 허용 계산기의 종류, 메모리 삭제 확인, 화면 표시 기능, 복수 지참 허용 여부의 네 가지 항목으로 5개 기관의 정책을 비교한다. 화면 표시 기능은 여러 줄의 식과 결과를 함께 보여 주는 다행 표시(multi-line display)와 그래프 표시 기능을 기준으로 보았다."))
B.append(h2("1. 기관별 정책"))
B.append(body("**한국.** 2026년 제49회 시험 시행계획 공고에 따르면 보험계리사 시험의 모든 과목에서 자료를 저장할 수 없는 단순계산기만 개별 지참하여 사용할 수 있으며, 공학용·재무용 등 데이터 저장 기능이 있는 전자계산기는 사용할 수 없다(금융감독원, 2026). 저장 기능 자체를 금지하므로 메모리 삭제 확인 절차는 규정되어 있지 않고, 화면 표시 방식과 지참 개수에 대한 별도 규정도 없다."))
B.append(body("**SOA.** 텍사스인스트루먼트(TI)의 6개 모델만 허용한다. 재무용인 BA-35, BA II Plus, BA II Plus Professional과 공학용인 TI-30Xa, TI-30X II, TI-30X MultiView가 이에 해당한다. TI-30X II, TI-30X MultiView, BA II Plus 계열은 응시자가 입실할 때 감독관이 메모리를 삭제하며, BA II Plus 계열은 이 과정에서 공장 초기값으로 재설정된다. 다행 표시 모델인 TI-30X MultiView는 허용되지만 그래프 계산기는 목록에 없다. 승인된 계산기는 여러 대 사용할 수 있으나, 목록에 없는 계산기를 사용하면 실격 처리되어 채점하지 않는다(SOA, 2026d)."))
B.append(body("**CAS.** SOA와 같은 6개 모델을 허용한다. 다만 메모리 삭제는 응시자가 시험 시작 전에 감독관에게 삭제된 상태를 보여 주는 방식이며, 시험실에는 1인당 1대만 반입할 수 있고 나머지는 사물함에 보관해야 한다. 또한 컴퓨터 기반 시험 화면에 TI-30XS MultiView 가상 계산기가 내장되어 있어, 계산기를 지참하지 않아도 응시할 수 있다(CAS, 2025)."))
B.append(body("**CIA.** ACIA 종합시험(Capstone Exam)에서는 표준 계산기(standard calculator) 사용을 허용하고 스마트 기기의 계산 기능은 금지한다. 시험에는 Excel과 RStudio가 필수이며, 답안을 R이나 Excel 파일로 제출할 수 있고 다중 모니터 사용도 허용된다(CIA, 2026d). '표준 계산기'의 범위와 메모리 삭제, 지참 개수에 대한 규정은 공개 자료에서 확인되지 않았다[[fn:4]]."))
B.append(body("**IFoA.** 2026년 9월 기술·환경 정책은 실물 계산기라면 제조사나 모델을 제한하지 않고, 온라인 계산기는 어떠한 경우에도 금지하며, 모든 시험에서 계산 보조 수단으로 Excel을 사용할 수 있도록 한다(IFoA, 2026d). 원격 감독 시험 안내서도 휴대전화나 온라인 계산기만 금지하고 그 밖의 계산기에는 제한이 없다고 밝히고 있다(IFoA, 2025b). 메모리 삭제와 지참 개수에 대한 규정은 공개 자료에서 확인되지 않았다."))
B.append(caption("<표 4> 기관별 허용 계산기 정책"))
B.append(table([1000, 2011, 1340, 1340, 1340], ["기관", "허용 종류", "메모리 삭제 확인", "화면 표시 기능", "복수 지참"], [
    ["한국", "단순계산기만 허용\n(공학용·재무용 금지)", "규정 없음\n(저장 기능 자체 금지)", "규정 없음", "규정 없음"],
    ["SOA", "TI 지정 6개 모델\n(재무용·공학용)", "감독관이\n입실 시 삭제", "다행 표시 허용\n(그래프 계산기 불가)", "허용"],
    ["CAS", "TI 지정 6개 모델\n(재무용·공학용)", "응시자가\n감독관에게 확인", "다행 표시 허용\n(화면 내 가상 계산기)", "1대만 반입"],
    ["CIA", "표준 계산기\n(스마트 기기 불가)\nExcel·R 필수", "확인되지 않음", "규정 없음\n(다중 모니터 허용)", "확인되지 않음"],
    ["IFoA", "실물 계산기 제한 없음\n(온라인 계산기 불가)\nExcel 허용", "확인되지 않음", "제한 없음", "확인되지 않음"],
], fills={(0, 1): GAP}))
B.append(note("주: '확인되지 않음'은 검토한 공개 자료에 해당 규정이 없음을 뜻함.\n자료: 금융감독원(2026), SOA(2026d), CAS(2025), CIA(2026d), IFoA(2025b; 2026d)를 바탕으로 저자 작성."))
B.append(h2("2. 시사점"))
B.append(body("한국 시험은 5개 기관 가운데 유일하게 공학용·재무용 계산기를 금지한다. SOA와 CAS는 재무용과 공학용 계산기를 허용하되 지정 모델 목록과 메모리 삭제 확인으로 부정행위 위험을 관리하고, CIA와 IFoA는 컴퓨터 기반 시험에서 Excel과 R 같은 분석 도구까지 허용한다. 계산 도구의 제한은 출제 가능한 문항의 범위를 좁힌다. 현행 정책에서는 연금 현가나 회귀계수처럼 반복 계산이 필요한 값을 단순계산기로 구해야 하므로, 시험 시간이 개념의 이해와 해석보다 계산 절차에 쓰이게 된다."))
B.append(body("따라서 데이터사이언스 평가의 확대와 함께 계산기 정책의 정비를 검토할 필요가 있다. 단기적으로는 SOA·CAS와 같이 지정 모델 목록을 두고 시험 시작 전에 메모리 삭제를 확인하는 방식으로 공학용·재무용 계산기를 허용할 수 있다. 장기적으로 실습형 평가를 도입한다면 CIA·IFoA와 같이 컴퓨터 기반 시험 환경에서 스프레드시트나 통계 소프트웨어를 사용하게 하는 방안을 함께 고려할 수 있다. 복수 지참은 SOA가 허용하고 CAS가 1대로 제한하는 등 기관마다 판단이 다르므로, 국내 시험의 감독 여건에 맞추어 정할 수 있을 것이다."))

B.append(blank())
B.append(h1("Ⅵ. 결론"))
B.append(body("본 연구는 데이터사이언스 역량을 기초이론, 기초실습, 과제형 심화실습의 세 단계로 구분하고, 한국 보험계리사 시험과 SOA, CAS, CIA, IFoA의 시험제도를 비교하였다. 해외 4개 기관은 평가 방식이 서로 다르지만 모두 이론과 실습 단계를 자격 요건에 포함하고 있었고, 한국 시험에서는 이에 대응하는 명시적 평가 요소가 확인되지 않았다. 다변화하는 위험과 고도화하는 위험관리 기법에 대응할 수 있는 계리 인력을 양성하기 위해서는 한국 보험계리사 시험에도 통계 및 데이터사이언스 관련 평가를 확대할 필요가 있다. 또한 허용 계산기 정책을 비교한 결과 한국 시험만이 공학용·재무용 계산기를 금지하고 있어, 평가 범위의 확대와 함께 계산 도구 정책도 정비할 필요가 있다."))
B.append(body("본 연구는 다음과 같은 한계를 가진다. 첫째, 비교는 각 기관이 공개한 안내 자료와 출제범위에 근거하였으므로 실제 문항의 난이도나 출제 비중까지는 반영하지 못하였다. 둘째, 한국 시험의 출제범위는 공식 출제기준을 통해 과목별 세부 내용을 확인할 필요가 있다. 셋째, CIA는 비인증 대학 출신자를 위한 새로운 경로의 도입을 예고하는 등 각국 제도가 계속 변화하고 있으므로, 후속 연구에서 이를 반영할 필요가 있다. 넷째, CIA와 IFoA의 계산기 정책 가운데 메모리 삭제와 지참 개수에 관한 사항은 공개 자료에서 확인하지 못하였다."))

B.append(pagebreak())
# ---------------------------------------------------------------- references
ACC = "(접속일: 2026.10.20.)"
REF_KO = [
    f"금융감독원. 「2021년도 제44회 보험계리사 및 손해사정사 시험 시행계획 공고」. 2021. https://support.epasskorea.com/hc/ko/article_attachments/900007490683 {ACC}",
    f"금융감독원. 「2026년도 제49회 보험계리사 및 손해사정사 시험 시행계획 공고」. 2026.1.9. 큐스터디 게재본. https://www.qstudy.kr/web/212 {ACC}",
    f"조세금융신문. 「금감원 “올해 보험계리사 120명, 손해사정사 506명 최종합격”」. 2025.9.26. https://www.tfmedia.co.kr/mobile/article.html?no=195820 {ACC}",
    f"CEO스코어데일리. 「보험부채 계산할 사람이 부족하다…‘계리사 모시기’ 승자는 삼성화재 179명」. 2026.6. https://m.ceoscoredaily.com/page/view/2026061820453141810 {ACC}",
]
REF_EN = [
    "Goldburd, Mark, Anand Khare, Dan Tevet, and Dmitriy Guller. Generalized Linear Models for Insurance Rating. 2nd ed., CAS Monograph No. 5, Casualty Actuarial Society, 2020.",
    "James, Gareth, Daniela Witten, Trevor Hastie, and Robert Tibshirani. An Introduction to Statistical Learning: with Applications in R. 2nd ed., Springer, 2021.",
]
REF_WEB = [
    f"actuary.info. \"The CAS at a Glance: CAS Exam Pathway 2026.\" 2026. https://actuary.info/insights/cas-exam-pathway-2026-complete-guide {ACC}",
    f"Canadian Institute of Actuaries (CIA). \"Pathway 2.\" 2026a. https://www.cia-ica.ca/becoming-a-member/qualification-pathways/pathway-2/ {ACC}",
    f"Canadian Institute of Actuaries (CIA). \"ACIA Syllabus.\" 2026b. https://cia-ica.ca/becoming-a-member/becoming-an-associate/acia-syllabus/ {ACC}",
    f"Canadian Institute of Actuaries (CIA). \"Education FAQ.\" 2026c. https://cia-ica.ca/becoming-a-member/education-faq/ {ACC}",
    f"Canadian Institute of Actuaries (CIA). \"ACIA Modules and Capstone Exam.\" 2026d. https://www.cia-ica.ca/becoming-a-member/becoming-an-associate/acia-modules-and-capstone-exam/ {ACC}",
    f"Casualty Actuarial Society (CAS). \"Modern Actuarial Statistics I and II: New Exams for an Evolving Profession.\" CAS Student Central, 2017. https://www.casstudentcentral.org/?p=3040 {ACC}",
    f"Casualty Actuarial Society (CAS). \"Exam MAS-II Content Outline.\" 2023. https://www.casact.org/sites/default/files/2022-12/MASII_Content_Outline_Fall-2023.pdf {ACC}",
    f"Casualty Actuarial Society (CAS). \"Property and Casualty Predictive Analytics (PCPA) FAQ.\" 2024. https://www.casact.org/sites/default/files/2024-05/CAS_PCPA_FAQ.pdf {ACC}",
    f"Casualty Actuarial Society (CAS). \"2025 Syllabus of Basic Education.\" 2025. https://casact.org/sites/default/files/2025-01/2025_SOBE.pdf {ACC}",
    f"Institute and Faculty of Actuaries (IFoA). \"CP2 Modelling Practice Syllabus 2025.\" 2025a. https://actuaries.org.uk/media/scjb3a3s/cp2-modelling-practice-syllabus-2025.pdf {ACC}",
    f"Institute and Faculty of Actuaries (IFoA). \"Examination Handbook for Remotely Invigilated Candidates.\" Sept. 2025b. https://actuaries.org.uk/media/5qmd1rw5/remotely-invigilated-examination-handbook-september-2025.pdf {ACC}",
    f"Institute and Faculty of Actuaries (IFoA). \"Actuarial Statistics.\" 2026a. https://actuaries.org.uk/qualify/curriculum/actuarial-statistics/ {ACC}",
    f"Institute and Faculty of Actuaries (IFoA). \"Frequently Asked Questions.\" 2026b. https://actuaries.org.uk/qualify/curriculum/frequently-asked-questions/ {ACC}",
    f"Institute and Faculty of Actuaries (IFoA). \"2026 Syllabus Changes.\" 2026c. https://actuaries.org.uk/media/444hp2jy/2026-syllabus-changes.pdf {ACC}",
    f"Institute and Faculty of Actuaries (IFoA). \"Technology and Environment Policy.\" Sept. 2026d. https://actuaries.org.uk/media/agdpgyw0/technology-and-environment-policy-final-sept-2026.pdf {ACC}",
    f"Insurance Portal. \"New Canadian Institute of Actuaries President Concerned about Climate.\" 2021. https://insurance-portal.ca/article/new-canadian-institute-of-actuaries-president-concerned-about-climate/ {ACC}",
    f"Society of Actuaries (SOA). \"Advanced Topics in Predictive Analytics (ATPA) Assessment.\" 2026a. https://www.soa.org/education/exam-req/edu-exam-atpa/ {ACC}",
    f"Society of Actuaries (SOA). \"Predictive Analytics (PA) Exam.\" 2026b. https://www.soa.org/education/exam-req/edu-exam-pa-detail/ {ACC}",
    f"Society of Actuaries (SOA). \"Statistics for Risk Modeling (SRM) Exam.\" 2026c. https://www.soa.org/education/exam-req/edu-exam-srm-detail/ {ACC}",
    f"Society of Actuaries (SOA). \"Use of Calculators and Other Assistance in Writing Examinations.\" 2026d. https://www.soa.org/education/exam-req/exam-day-info/edu-calculators.aspx {ACC}",
    f"University of Southampton. \"Southampton Data Science Academy to Deliver IFoA Certificate in Data Science.\" 2020. https://www.southampton.ac.uk/wsi/news/2020/02/ifoa-certification.page {ACC}",
]
def ref(s): return para(run(s, sz=20), "aa", '<w:jc w:val="left"/>')
R = [para(run("<참고문헌>"), "a9")]
R += [ref(s) for s in REF_KO + REF_EN + REF_WEB]
R.append(pagebreak())

# English abstract section (reuse template paragraphs, replace text)
EN = [para(run("<Abstract>"), "a9"),
      para(run(ABSTRACT_EN, sz=20), "1", '<w:spacing w:line="320" w:lineRule="auto"/><w:rPr><w:position w:val="0"/></w:rPr>'),
      para("", "1"),
      para(run("Keywords", font="KoPub돋움체 Bold", sz=20) + run("：" + KW_EN, sz=20), "af1", '<w:ind w:left="1041" w:hanging="1041"/>')]

# replace everything from the page break after keywords table up to sectPr
start = x.find('<w:br w:type="page"/>')
start = x.rfind('<w:p ', 0, start)
end = x.rfind('<w:sectPr')
x = x[:start] + pagebreak() + "".join(B) + "".join(R) + "".join(EN) + x[end:]
open(doc_path, "w", encoding="utf8").write(x)

# ---------------------------------------------------------------- footnotes
FN["2"] = "한국보험계리사회(Institute of Actuaries of Korea)는 계리사의 직업 단체로서 자격시험을 주관하지 않는다. 본 연구에서 비교하는 대상은 국가자격시험인 보험계리사 시험이다."
FN["4"] = "CIA의 FCIA 시험 응시자 지침(Participant Guidelines)에도 허용 도구에 관한 규정이 있을 수 있으나, 본 연구에서는 해당 문서의 원문을 확인하지 못하였다."
FN["3"] = "ACIA 자격의 종합시험(Capstone Exam)은 오픈북 형식의 응용·통합형 시험으로, 예측분석 내용을 일부 포함할 수 있으나 데이터사이언스 과제를 독립적으로 평가하는 시험은 아니다(CIA, 2026b)."
f = open(fn_path, encoding="utf8").read()
add = "".join(
    f'<w:footnote w:id="{i}"><w:p><w:pPr><w:pStyle w:val="af"/></w:pPr>'
    f'<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r>{run(" " + t, sz=18)}</w:p></w:footnote>'
    for i, t in sorted(FN.items(), key=lambda kv: int(kv[0])))
f = f.replace("</w:footnotes>", add + "</w:footnotes>")
open(fn_path, "w", encoding="utf8").write(f)
print("ok")

# body style 本文1 carries a 0.5pt baseline raise (w:position=1) that some
# renderers turn into a shrunken superscript; drop it so body text is 10pt.
sp = f"{D}/word/styles.xml"
st = open(sp, encoding="utf8").read()
st = st.replace('<w:name w:val="본문1"/><w:pPr><w:tabs><w:tab w:val="center" w:pos="3704"/><w:tab w:val="right" w:pos="7271"/></w:tabs><w:ind w:firstLine="200"/></w:pPr><w:rPr><w:position w:val="1"/></w:rPr>',
                '<w:name w:val="본문1"/><w:pPr><w:tabs><w:tab w:val="center" w:pos="3704"/><w:tab w:val="right" w:pos="7271"/></w:tabs><w:ind w:firstLine="200"/></w:pPr><w:rPr><w:sz w:val="20"/></w:rPr>', 1)
open(sp, "w", encoding="utf8").write(st)
