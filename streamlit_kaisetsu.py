"""
薬剤師国家試験 解説生成システム（Streamlit版）
ブラウザで完結。ドラッグ＆ドロップ→生成→ダウンロード。
"""

import streamlit as st
import json
import re
import copy
import io
import urllib.request
import urllib.error
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ──────────────────────────────────────────────
# フォーマットエンジン
# ──────────────────────────────────────────────
TNR      = 'Times New Roman'
CENTURY  = 'Century'
GOTHIC   = 'ＭＳ ゴシック'
HIRAGINO = 'ヒラギノ角ゴ ProN W3'
MARKUP   = re.compile(r'\{\{([^}]+)\}\}')
NUMBERS  = ['１', '２', '３', '４', '５']


_DASH_CHARS = ('-', '−', '‐', '‑', '–', '—',
               '－')  # 半角ハイフン/数学マイナス/各種ハイフン・ダッシュ/全角


def normalize_dashes(text):
    """連結記号・マイナスを全角ハイフンマイナス「－」（U+FF0D）に統一する。
    半角ハイフン・数学マイナス(U+2212)・各種ダッシュを吸収し、
    フォントが周囲と揃うようにする。"""
    for ch in _DASH_CHARS:
        text = text.replace(ch, '－')
    return text


def parse_markup(text):
    runs, pos = [], 0
    for m in MARKUP.finditer(text):
        if m.start() > pos:
            runs.append({'text': normalize_dashes(text[pos:m.start()])})
        c = m.group(1)
        if   c == '-':           runs.append({'text': '－'})
        elif c == 'mu':          runs.append({'text': 'µ',  'font': CENTURY, 'italic': True})
        elif c.startswith('sup:'): runs.append({'text': c[4:], 'sup': True})
        elif c.startswith('CL:'): runs += [{'text':'CL','font':TNR,'italic':True},{'text':c[3:],'sub':True}]
        elif c == 'CL':          runs.append({'text':'CL','font':TNR,'italic':True})
        elif c.startswith('f:'): runs += [{'text':'f','font':CENTURY,'italic':True},{'text':c[2:],'sub':True}]
        elif c == 'f':           runs.append({'text':'f','font':CENTURY,'italic':True})
        elif c.startswith('K:'): runs += [{'text':'K','font':CENTURY,'italic':True},{'text':c[2:],'sub':True}]
        elif c.startswith('t:'): runs += [{'text':'t','font':CENTURY,'italic':True},{'text':c[2:],'sub':True}]
        elif 'Vd' in c:
            runs.append({'text':'Vd','font':CENTURY,'italic':True})
            if ':' in c: runs.append({'text':c.split(':',1)[1],'sub':True})
        else: runs.append({'text': c})
        pos = m.end()
    if pos < len(text): runs.append({'text': normalize_dashes(text[pos:])})
    return runs


def make_run(text, font=None, italic=False, bold=None, sub=False, sup=False):
    r = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    if font:
        rf = OxmlElement('w:rFonts')
        for a in ('w:ascii','w:hAnsi','w:eastAsia','w:cs'): rf.set(qn(a), font)
        rPr.append(rf)
    if bold is True:  rPr.append(OxmlElement('w:b'))
    elif bold is False:
        b = OxmlElement('w:b'); b.set(qn('w:val'),'0'); rPr.append(b)
    if italic: rPr.append(OxmlElement('w:i')); rPr.append(OxmlElement('w:iCs'))
    if sub or sup:
        va = OxmlElement('w:vertAlign')
        va.set(qn('w:val'), 'subscript' if sub else 'superscript')
        rPr.append(va)
    if len(rPr): r.append(rPr)
    t = OxmlElement('w:t'); t.text = text
    if text and (text[0]==' ' or text[-1]==' '):
        t.set('{http://www.w3.org/XML/1998/namespace}space','preserve')
    r.append(t)
    return r


def add_runs(p_elem, specs):
    for s in specs:
        p_elem.append(make_run(s.get('text',''), font=s.get('font'),
            italic=s.get('italic',False), bold=s.get('bold'),
            sub=s.get('sub',False), sup=s.get('sup',False)))


def clear_runs(para):
    for r in para._p.findall(qn('w:r')): para._p.remove(r)


EXP_PAT = re.compile(r'^[１２３４５]　(誤|正)：')
HEAD_PAT = re.compile(r'^問\s*(\d+)\s*(?:（([^）]*)）)?')


def _is_answer_line(text):
    t = text.strip()
    return t.startswith('問') and '解答' in t


def segment_questions(doc):
    """テンプレートを小問ブロックに分割する。
    連問（問196以降）は複数の解答行を持つため、解答行を区切りとしてブロック化する。
    各ブロック: 問番号・科目・前文プレースホルダ・選択肢解説プレースホルダ・解答行。
    単問の場合は要素1つのリストを返す。"""
    paras = doc.paragraphs
    ans_idxs = [i for i, p in enumerate(paras) if _is_answer_line(p.text)]
    blocks = []
    prev = -1
    for ai in ans_idxs:
        exps = [i for i in range(prev + 1, ai) if EXP_PAT.match(paras[i].text)]
        if not exps:
            prev = ai
            continue
        first = exps[0]
        # ブロック内の問番号ヘッダー（問NNN（科目））
        header_idx, qnum, subject = None, None, None
        for i in range(first - 1, prev, -1):
            t = paras[i].text.strip()
            if t.startswith('問') and '解答' not in t:
                m = HEAD_PAT.match(t)
                if m:
                    header_idx, qnum, subject = i, m.group(1), m.group(2)
                    break
        # 解答行からも問番号を補完
        if qnum is None:
            m = HEAD_PAT.match(paras[ai].text.strip())
            if m:
                qnum = m.group(1)
        # 前文プレースホルダ（選択肢解説の直前・非空・番号始まりでない段落）
        pre_idx = None
        for i in range(first - 1, (header_idx if header_idx is not None else prev), -1):
            t = paras[i].text.strip()
            if not t:
                continue
            if EXP_PAT.match(paras[i].text):
                continue
            if t[0] in '１２３４５' or t.startswith('問') or '解答' in t:
                break  # 実際の選択肢・見出しに到達 → 前文なし
            pre_idx = i
            break
        blocks.append({
            'question_num': qnum, 'subject': subject,
            'header_para_idx': header_idx, 'preamble_para_idx': pre_idx,
            'explanation_para_idxs': exps, 'answer_para_idx': ai,
        })
        prev = ai
    return blocks


def extract_info(doc):
    """後方互換: 最初の小問ブロックを単問形式で返す。"""
    blocks = segment_questions(doc)
    if not blocks:
        return {'question_num': None, 'preamble_para_idx': None,
                'explanation_para_idxs': [], 'answer_para_idx': None}
    b = blocks[0]
    return {'question_num': b['question_num'],
            'preamble_para_idx': b['preamble_para_idx'],
            'explanation_para_idxs': b['explanation_para_idxs'],
            'answer_para_idx': b['answer_para_idx']}


def _write_header(p, qnum, subject):
    """問番号ヘッダーを問題側と統一（問＋番号をボールド、ＭＳゴシック）。"""
    clear_runs(p)
    p._p.append(make_run('問', font=GOTHIC, bold=True))
    p._p.append(make_run(str(qnum or '000'), font=GOTHIC, bold=True))
    if subject:
        p._p.append(make_run('（' + subject + '）', font=GOTHIC))


def write_block(doc, block, expl):
    """1つの小問ブロックに前文・選択肢解説・解答を書き込む（削除はしない）。"""
    paras = doc.paragraphs
    # ヘッダー（「問N（科目）」形式の短い見出しのみ、問番号をボールド化して統一）
    # 問題文の設問行（長文）を誤って書き換えないよう文字数でガードする
    if block.get('header_para_idx') is not None:
        hp = paras[block['header_para_idx']]
        if len(hp.text.strip()) <= 12 and re.match(r'^問\d+', hp.text.strip()):
            _write_header(hp, block.get('question_num'), block.get('subject'))

    # 選択肢解説
    for idx, item in zip(block['explanation_para_idxs'], expl.get('選択肢解説', [])):
        p = paras[idx]
        clear_runs(p)
        p._p.append(make_run('', font=GOTHIC, bold=True))
        p._p.append(make_run(item.get('番号', ''), font=GOTHIC))
        p._p.append(make_run('　', font=HIRAGINO))
        p._p.append(make_run(item.get('正誤', '誤') + '：'))
        add_runs(p._p, parse_markup(item.get('内容', '')))

    # 解答行（右揃え・全体ボールド）
    p_ans = paras[block['answer_para_idx']]
    clear_runs(p_ans)
    p_ans.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    q = block.get('question_num', '000')
    for seg in ['問', str(q), '　', '解答　']:
        p_ans._p.append(make_run(seg, font=GOTHIC, bold=True))
    for i, ans in enumerate(expl.get('解答', [])):
        p_ans._p.append(make_run(ans, font=GOTHIC, bold=True))
        if i < len(expl['解答']) - 1:
            p_ans._p.append(make_run('、', font=GOTHIC, bold=True))

    # 前文
    if block.get('preamble_para_idx') is not None:
        p_pre = paras[block['preamble_para_idx']]
        clear_runs(p_pre)
        前文 = expl.get('前文', '')
        add_runs(p_pre._p, parse_markup(前文) if 前文 else [{'text': ''}])
        pPr = p_pre._p.find(qn('w:pPr'))
        last = p_pre._p
        for line in expl.get('前文_追加行', []):
            np = OxmlElement('w:p')
            if pPr is not None:
                np.append(copy.deepcopy(pPr))
            last.addnext(np)
            add_runs(np, parse_markup(line))
            last = np


def write_all(doc, blocks, expls):
    """全ブロックを書き込み、最後の解答行より後ろのテンプレート雛形を削除する。"""
    # 末尾雛形の削除（最後の解答行以降）
    last_ans = doc.paragraphs[blocks[-1]['answer_para_idx']]._p
    del_elems, found = [], False
    for p in doc.paragraphs:
        if found:
            del_elems.append(p._p)
        if p._p is last_ans:
            found = True
    # 書き込みは段落インデックスに依存するため、削除より前に実施
    for block, expl in zip(blocks, expls):
        write_block(doc, block, expl)
    for e in del_elems:
        if e.getparent() is not None:
            e.getparent().remove(e)


def write_to_doc(doc, info, expl):
    """後方互換: 単問1つ分を書き込む。"""
    block = {'question_num': info.get('question_num'), 'subject': None,
             'header_para_idx': None,
             'preamble_para_idx': info.get('preamble_para_idx'),
             'explanation_para_idxs': info['explanation_para_idxs'],
             'answer_para_idx': info['answer_para_idx']}
    write_all(doc, [block], [expl])


SYSTEM_PROMPT = """あなたは薬剤師国家試験の解説作成の専門家です。
以下のJSON形式のみで回答してください（他のテキスト不要）:
{
  "前文": "概念整理・公式・前提（純粋な知識問題で不要なら空文字）",
  "前文_追加行": ["前文が複数段落になる場合の2段落目", "3段落目"],
  "選択肢解説": [
    {"番号":"１","正誤":"誤","内容":"解説文"},
    {"番号":"２","正誤":"正","内容":"解説文"},
    {"番号":"３","正誤":"誤","内容":"解説文"},
    {"番号":"４","正誤":"誤","内容":"解説文"},
    {"番号":"５","正誤":"正","内容":"解説文"}
  ],
  "解答": ["２","５"]
}

【文体・スタイルのルール】
- 文末：「〜である。」「〜となる。」「〜と考えられる。」「〜と推定できる。」など断定調。「〜です」「〜ます」は使わない。
- 「誤」の選択肢：なぜ誤りかを説明し、「なお、正しくは〜である。」や「〜ではなく、〜である。」で正しい情報も必ず示す。
- 前文で説明済みの内容は「上記参照。」でよい。
- 他の選択肢を参照するときは「選択肢X参照。」と書く。
- 計算問題では前文に「●Step1：」「●Step2：」などの段階的な解法を示す。
- 薬学専門用語・薬物名・受容体名は正確に記載する。
- 計算は数値を具体的に示し、途中式を省略しない。

【記号・表記の統一ルール】
- ハイフンとマイナス：数式・モデル名・化合物名・略号内の連結記号にはすべて「－」（全角ハイフンマイナス、U+FF0D）を使用し、半角ハイフン「-」は一切使わない。
  例：「1－コンパートメントモデル」「P－糖タンパク質」「P－gp」「UDP－グルクロン酸転移酵素」「L－カルボシステイン」「α－ヘリックス」（「1-」「P-」のような半角ハイフンは禁止）
- 引き算の記号はすべて「－」（全角マイナス）。例：50－30＝20
- クリアランスの区別：
  ・CL = 全身クリアランス（肝・腎などをすべて含む）
  ・CLr = 腎クリアランス（renal clearance）
  ・CLh = 肝クリアランス（hepatic clearance）
  ・問題文でCLtotと明示されている場合のみCLtotを使用
- 投与量：D（一般）、Dpo（経口）、Div（静注）。"Dose"は使わない。
- 消失速度定数：kel（keは使わない）
- 吸収速度定数：ka
- 分布容積：Vd
- バイオアベイラビリティ：F
- 最高血中濃度：Cmax
- 定常状態血中濃度：Css（平均はCss,av）
- 消失半減期：t1/2（T1/2は使わない）
- 投与間隔：τ
- 平均滞留時間：MRT、平均吸収時間：MAT、平均溶出時間：MDT
- 式番号：①②③…を用いて文中で相互参照
- 単位：h、L、mg、h－1（上付き－1）など

【良い解説の例①：薬剤（薬物動態）知識問題（第110回 問173）】
正解：３、５
前文：「薬物の尿中排泄は、糸球体におけるろ過、尿細管における分泌、再吸収という三つの過程によって行われる。血液中に含まれる薬物のうちタンパクと結合していない非結合形の薬物が糸球体でろ過を受ける。次に、近位尿細管においてトランスポーター等によって認識される薬物は分泌を受け、尿細管に流入する。その後、尿細管内に流入した薬物のうち再吸収を受けなかった薬物が尿中に排泄される。このうち、糸球体ろ過における単位時間あたりの血漿のろ過量を糸球体ろ過速度（GFR）といい、通常成人では約100 mL/min/1.73 m2である。ろ過クリアランスは、GFR×fp（fp：血漿タンパク非結合率）で表される。イヌリンは血漿タンパクと結合せず（fp=1）、尿細管分泌や再吸収を受けないため、イヌリンクリアランス＝GFRとなる。一方、クレアチニンは尿細管において若干の分泌を受けるため、クレアチニンクリアランス＞GFRとなる。」
選択肢１：「誤：GFR＝イヌリンクリアランス＝30 mL/min/1.73 m2と推定できる。」
選択肢２：「誤：イヌリンは尿細管で分泌・再吸収を受けないため、再吸収クリアランスは存在しない。」
選択肢３：「正：クレアチニンの尿細管分泌クリアランス＝クレアチニンクリアランス－イヌリンクリアランス＝50－30＝20 mL/min/1.73 m2と推定できる。」
選択肢４：「誤：正常成人のGFRは約100 mL/min/1.73 m2であるため、本患者のイヌリンクリアランス30 mL/min/1.73 m2は正常時より小さいと考えられる。」
選択肢５：「正：本患者のクレアチニンクリアランス50 mL/min/1.73 m2は正常時（約100 mL/min/1.73 m2）より小さいと考えられる。」

【良い解説の例②：薬理知識問題（第110回 問160）】
正解：１、５
前文：「」（純粋知識問題のため前文なし）
選択肢１：「正：オキシメテバノールは麻薬性鎮咳薬であり、オピオイド受容体を刺激して鎮咳作用を示す。」
選択肢２：「誤：L－カルボシステインは構造中にSH基を有さず、ムコタンパク質の構成成分であるシアル酸・フコースのバランスを改善して去痰作用を示す。なお、SH基を有しムコタンパク質のペプチド鎖の連結を切断して去痰作用を示す薬剤はアセチルシステインである。」
選択肢３：「誤：フルマゼニルはベンゾジアゼピン受容体に結合し、ベンゾジアゼピン系薬の作用に拮抗することで呼吸抑制を改善する。なお、末梢性化学受容器を刺激して間接的に呼吸中枢を興奮させる薬剤はドキサプラムである。」
選択肢４：「誤：アンブロキソールはブロムヘキシンの活性代謝物であり、肺サーファクタント分泌の促進・線毛運動の亢進により去痰作用を示す。（本選択肢は親薬物と活性代謝物が逆）」
選択肢５：「正：ニンテダニブは低分子チロシンキナーゼ阻害薬であり、VEGFR・FGFR・PDGFRのチロシンキナーゼを阻害して肺の線維化を抑制する。」

【良い解説の例③：生化学知識問題（第110回 問115）】
正解：２、３
前文：「」（純粋知識問題のため前文なし）
選択肢１：「誤：α－ヘリックスとは1本のポリペプチド鎖が分子内水素結合することで形成される細長いらせん構造を指す。コラーゲンの三重らせん構造は3本のポリペプチド鎖が形成する特殊ならせん構造であり、α－ヘリックスとは異なる。」
選択肢２：「正：Xはプロリン（Pro）の翻訳後修飾で生じたアミノ酸であることから、ヒドロキシプロリンであると考えられる。」
選択肢３：「正：グリシン（Gly）は側鎖が水素原子であるため空間を占める割合が小さく、コラーゲンの三重らせん構造における混み合った部位に安定して収まることができるため、三重らせん構造の形成に重要な役割を担う。」
選択肢４：「誤：ビタミンCがコラーゲン遺伝子の転写を促進する際に核内に移行するかどうかは現在不明である。」
選択肢５：「誤：コラーゲンは細胞外マトリックスの代表的な構成タンパク質であり組織の強度を保つが、細胞内の細胞骨格を構成するわけではない。」

【良い解説の例④：グラフ問題（第111回 問170）】
正解：１、５
前文：「リボフラビンは、食事の有無による胃内容排出速度（GER）の違いにより吸収量が変化する。空腹時に服用するとGERの増大により、十二指腸に存在する吸収トランスポーターが飽和しやすくなり、吸収量が低下する。一方、食後に服用するとGERの低下により、吸収トランスポーターの飽和が起こりにくくなることで吸収量が増大する。なお、グラフの縦軸は累積尿中排泄量を示しているが、累積尿中排泄量≒吸収量と考える必要があるため、Aが朝食後服用（吸収量が多い）、Bが空腹時服用（吸収量が少ない）と読み取れる。」
選択肢１：「正：上記参照。」
選択肢２：「誤：上記参照。」
選択肢３：「誤：BがAより低値となるのは、リボフラビンのGERが空腹時に増大することによる吸収トランスポーターの飽和が原因である。」
選択肢４：「誤：選択肢３参照。」
選択肢５：「正：メトクロプラミドはドパミンD2受容体遮断薬であり、D2受容体を遮断することでコリン作動性神経を興奮させGERを増大させる。そのため、メトクロプラミドを前投与した時の曲線は空腹時服用時と類似し、AよりBに近くなる。」

マークアップ（数式・変数名に使用）:
{{CL:r}}=CLr {{f:e}}=fe {{K:sp}}=Ksp {{t:1/2}}=t1/2 {{Vd}}=Vd
{{mu}}=µ {{-}}=全角マイナス（－） {{sup:2}}=上付き2"""


RENMON_INSTRUCTION = """
【連問（問196以降の実践問題）の指示】
これは複数の小問が連続する「連問」である。以下の小問すべての解説を、必ず次のJSON形式（連問配列）のみで一括生成すること:
{{
  "連問": [
    {{"問番号":"{first}", "前文":"...", "前文_追加行":[], "選択肢解説":[{{"番号":"１","正誤":"誤","内容":"..."}}, ...], "解答":["３"]}},
    ...（小問の数だけ続ける）...
  ]
}}
対象の小問（この順序・この問番号で出力すること）: {targets}
- 連問は前の小問の内容を前提とする（例：「前問の処方提案の理由は〜」）。共有の症例・検査値・処方を踏まえ、各小問の整合性を保つこと。
- 各小問の「選択肢解説」は、その小問の選択肢数に一致させること。
- 「1つ選べ」の小問は解答を1つ、「2つ選べ」は2つ、という具合に設問の指示に従うこと。"""


def _extract_json(text):
    depth, start = 0, None
    for i, ch in enumerate(text):
        if ch == '{':
            if depth == 0:
                start = i
            depth += 1
        elif ch == '}':
            depth -= 1
            if depth == 0 and start is not None:
                return json.loads(text[start:i + 1])
    raise ValueError("JSONが見つかりません")


def call_api(question_text, api_key, sub_questions=None):
    """単問なら解説JSONを、連問なら {'連問':[...]} を返す。
    sub_questions: 連問時の [{'番号':..,'科目':..,'選択肢数':..}] のリスト。"""
    user_content = f"解説を生成してください:\n\n{question_text}"
    if sub_questions and len(sub_questions) > 1:
        targets = "、".join(
            f"問{s['番号']}（{s.get('科目') or ''}・選択肢{s.get('選択肢数', 5)}）"
            for s in sub_questions)
        user_content += RENMON_INSTRUCTION.format(
            first=sub_questions[0]['番号'], targets=targets)
    payload = {
        "model": "claude-opus-4-8",
        "max_tokens": 8192,
        "system": SYSTEM_PROMPT,
        "messages": [{"role": "user", "content": user_content}]
    }
    req = urllib.request.Request(
        "https://api.anthropic.com/v1/messages",
        data=json.dumps(payload).encode(),
        headers={'Content-Type':'application/json',
                 'x-api-key': api_key,
                 'anthropic-version':'2023-06-01'},
        method='POST')
    with urllib.request.urlopen(req, timeout=180) as resp:
        text = json.loads(resp.read())['content'][0]['text']
        return _extract_json(text)


# ──────────────────────────────────────────────
# Streamlit UI
# ──────────────────────────────────────────────
def main():
    st.set_page_config(
        page_title='薬剤師国家試験 解説生成システム',
        page_icon='📋',
        layout='centered'
    )

    st.title('📋 薬剤師国家試験 解説生成システム')
    st.markdown('テンプレートdocxをアップロードして解説を自動生成します。')
    st.divider()

    # ① ファイルアップロード
    st.subheader('① テンプレートファイルをアップロード')
    uploaded = st.file_uploader(
        'docxファイルをドラッグ＆ドロップ、またはクリックして選択',
        type=['docx'],
        help='問題と「ああああ」プレースホルダーが入ったテンプレートdocxファイル'
    )

    # ② モード
    st.subheader('② モードを選択')
    mode = st.radio('', ['🤖 AIが自動生成', '✏️ 手動入力（無料）'],
                    horizontal=True, label_visibility='collapsed')

    # ③ 入力
    st.subheader('③ 解説内容')
    explanation = None

    if 'AI' in mode:
        # APIキー：Streamlit secretsにあれば自動使用
        if 'ANTHROPIC_API_KEY' in st.secrets:
            api_key = st.secrets['ANTHROPIC_API_KEY']
            st.success('✅ APIキー設定済み（管理者設定）')
        else:
            api_key = st.text_input(
                'Anthropic APIキー',
                type='password',
                placeholder='sk-ant-api03-...',
                help='console.anthropic.com で取得。1問あたり約1〜2円。'
            )
    else:
        api_key = None
        st.markdown('各選択肢の解説を入力してください。')
        st.caption('マークアップ: `{{CL:r}}`=CLr　`{{f:e}}`=fe　`{{-}}`=全角マイナス　`{{mu}}`=µ')

        manual_data = {}
        for char in NUMBERS:
            col1, col2 = st.columns([1, 5])
            with col1:
                seigo = st.selectbox(char, ['誤', '正'], key=f'seigo_{char}', label_visibility='visible')
            with col2:
                content = st.text_input(f'選択肢{char}の解説', key=f'content_{char}',
                                        label_visibility='collapsed',
                                        placeholder=f'選択肢{char}の解説を入力...')
            manual_data[char] = {'seigo': seigo, 'content': content}

        st.markdown('')
        col_a, col_b = st.columns(2)
        with col_a:
            ans_input = st.text_input('解答（例: ２,５）', placeholder='２,５')
        with col_b:
            preamble = st.text_input('前文（任意）', placeholder='計算式など')

    st.divider()

    # 生成ボタン
    generate = st.button('🚀 解説を生成する', type='primary', use_container_width=True)

    if generate:
        # バリデーション
        if not uploaded:
            st.error('ファイルをアップロードしてください。')
            st.stop()

        if 'AI' in mode and not api_key:
            st.error('APIキーを入力してください。')
            st.stop()

        if '手動' in mode:
            missing = [c for c in NUMBERS if not manual_data[c]['content'].strip()]
            if missing:
                st.error(f'選択肢{"、".join(missing)}の解説を入力してください。')
                st.stop()
            if not ans_input.strip():
                st.error('解答を入力してください。')
                st.stop()

            explanation = {
                '前文': preamble.strip() if preamble else '',
                '前文_追加行': [],
                '選択肢解説': [
                    {'番号': c, '正誤': manual_data[c]['seigo'], '内容': manual_data[c]['content']}
                    for c in NUMBERS
                ],
                '解答': [a.strip() for a in ans_input.split(',')]
            }

        # 処理
        with st.spinner('処理中...'):
            try:
                doc = Document(io.BytesIO(uploaded.read()))
                blocks = segment_questions(doc)

                if not blocks:
                    st.error('テンプレートの構造が読み取れませんでした。ファイルを確認してください。')
                    st.stop()

                is_renmon = len(blocks) > 1
                q_labels = '・'.join(f"問{b['question_num']}" for b in blocks)
                if is_renmon:
                    st.info(f'連問を検出しました（{q_labels}／{len(blocks)}問）')
                else:
                    st.info(f'問{blocks[0]["question_num"]} を検出しました')

                if 'AI' in mode:
                    # 問題文のみ抽出（解説プレースホルダー・解答行は除外）
                    exp_pat2 = re.compile(r'^[１２３４５][\s　](誤|正)：')
                    lines = []
                    for p in doc.paragraphs:
                        t = p.text.strip()
                        if not t:
                            continue
                        if exp_pat2.match(t):
                            continue
                        if _is_answer_line(t):
                            continue
                        # 雛形プレースホルダ（「ああ」「ひな型」等）を除外
                        if 'ひな型' in t or ('解説（' in t and 'ああ' in t):
                            continue
                        lines.append(t)
                    for table in doc.tables:
                        lines.append('')
                        for row in table.rows:
                            cells = [c.text.replace('\n', ' ').strip() for c in row.cells]
                            if any(cells):
                                lines.append(' | '.join(cells))
                    question_text = '\n'.join(lines)

                    st.info('Claude API で解説生成中...')
                    uploaded.seek(0)
                    doc = Document(io.BytesIO(uploaded.read()))
                    blocks = segment_questions(doc)

                    sub_meta = [{'番号': b['question_num'], '科目': b.get('subject'),
                                 '選択肢数': len(b['explanation_para_idxs'])}
                                for b in blocks]
                    result = call_api(question_text, api_key,
                                      sub_questions=sub_meta if is_renmon else None)

                    if is_renmon:
                        items = result.get('連問', []) if isinstance(result, dict) else []
                        if not items:
                            raise ValueError('連問形式の応答が得られませんでした。')
                        by_num = {str(it.get('問番号')): it for it in items}
                        expls = []
                        for i, b in enumerate(blocks):
                            it = by_num.get(str(b['question_num']))
                            if it is None:
                                it = items[i] if i < len(items) else {}
                            expls.append(it)
                    else:
                        expls = [result]
                else:
                    if is_renmon:
                        st.error('連問は手動入力に未対応です。AIモードをご利用ください。')
                        st.stop()
                    expls = [explanation]

                write_all(doc, blocks, expls)

                # docxをメモリに保存
                buf = io.BytesIO()
                doc.save(buf)
                buf.seek(0)

                fname = uploaded.name.replace('.docx', '_完成.docx')
                ans_str = ' ／ '.join(
                    f"問{b['question_num']}: " + '、'.join(e.get('解答', []))
                    for b, e in zip(blocks, expls))

                st.success(f'✅ 完成！（解答 → {ans_str}）')

                st.download_button(
                    label='📥 完成ファイルをダウンロード',
                    data=buf,
                    file_name=fname,
                    mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                    use_container_width=True
                )

            except urllib.error.HTTPError as e:
                st.error(f'APIエラー（{e.code}）: APIキーを確認してください。')
            except Exception as e:
                st.error(f'エラーが発生しました: {e}')
                st.exception(e)


if __name__ == '__main__':
    main()
