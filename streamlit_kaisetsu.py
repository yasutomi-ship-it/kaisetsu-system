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

# ──────────────────────────────────────────────
# フォーマットエンジン
# ──────────────────────────────────────────────
TNR      = 'Times New Roman'
CENTURY  = 'Century'
GOTHIC   = 'ＭＳ ゴシック'
HIRAGINO = 'ヒラギノ角ゴ ProN W3'
MARKUP   = re.compile(r'\{\{([^}]+)\}\}')
NUMBERS  = ['１', '２', '３', '４', '５']


def parse_markup(text):
    runs, pos = [], 0
    for m in MARKUP.finditer(text):
        if m.start() > pos:
            runs.append({'text': text[pos:m.start()]})
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
    if pos < len(text): runs.append({'text': text[pos:]})
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


def extract_info(doc):
    info = {'question_num': None, 'preamble_para_idx': None,
            'explanation_para_idxs': [], 'answer_para_idx': None}
    paras = doc.paragraphs
    exp_pat = re.compile(r'^[１２３４５]　(誤|正)：')
    for p in paras:
        m = re.match(r'問(\d+)', p.text)
        if m: info['question_num'] = m.group(1); break
    for i, p in enumerate(paras):
        t = p.text.strip()
        if exp_pat.match(p.text): info['explanation_para_idxs'].append(i)
        q = info['question_num']
        if '解答' in t and t.startswith('問'): info['answer_para_idx'] = i
    if info['explanation_para_idxs']:
        first = info['explanation_para_idxs'][0]
        for i in range(first-1, max(0, first-5), -1):
            t = paras[i].text.strip()
            if t and not t.startswith('問') and '解答' not in t:
                info['preamble_para_idx'] = i; break
    return info


def write_to_doc(doc, info, expl):
    paras = doc.paragraphs
    p_pre  = paras[info['preamble_para_idx']] if info['preamble_para_idx'] is not None else None
    p_exps = [paras[i] for i in info['explanation_para_idxs']]
    p_ans  = paras[info['answer_para_idx']]

    del_elems = []
    found = False
    for p in paras:
        if found: del_elems.append(p._p)
        if p is p_ans: found = True

    for p, item in zip(p_exps, expl.get('選択肢解説',[])):
        clear_runs(p)
        p._p.append(make_run('', font=GOTHIC, bold=True))
        p._p.append(make_run(item.get('番号',''), font=GOTHIC))
        p._p.append(make_run('　', font=HIRAGINO))
        p._p.append(make_run(item.get('正誤','誤') + '：'))
        add_runs(p._p, parse_markup(item.get('内容','')))

    clear_runs(p_ans)
    q = info.get('question_num','000')
    for seg in [('問',True),(q,True),('　',True),('解答　',True)]:
        p_ans._p.append(make_run(seg[0], font=GOTHIC, bold=seg[1]))
    for i, ans in enumerate(expl.get('解答',[])):
        p_ans._p.append(make_run(ans, font=GOTHIC, bold=True))
        if i < len(expl['解答'])-1:
            p_ans._p.append(make_run('、', font=GOTHIC, bold=True))

    if p_pre:
        clear_runs(p_pre)
        前文 = expl.get('前文','')
        add_runs(p_pre._p, parse_markup(前文) if 前文 else [{'text':''}])
        pPr = p_pre._p.find(qn('w:pPr'))
        last = p_pre._p
        for line in expl.get('前文_追加行',[]):
            np = OxmlElement('w:p')
            if pPr is not None: np.append(copy.deepcopy(pPr))
            last.addnext(np); add_runs(np, parse_markup(line)); last = np

    for e in del_elems:
        if e.getparent() is not None: e.getparent().remove(e)


SYSTEM_PROMPT = """あなたは薬剤師国家試験の解説作成の専門家です。
以下のJSON形式のみで回答してください（他のテキスト不要）:
{
  "前文": "計算や前提整理（不要なら空文字）",
  "前文_追加行": ["前文が複数行になる場合の2行目", "3行目"],
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
- 文末は「〜である。」「〜となる。」など体言止め・簡潔な断定調。「〜です」「〜ます」は使わない。
- 余計な前置き・繰り返しは禁止。各選択肢は簡潔に核心のみ述べる。
- 前文で説明済みの内容は「正：上記参照。」「誤：上記参照。」でよい。
- 他の選択肢を参照するときは「選択肢X参照。」と書く。
- 薬学専門用語は正確に使う（例：「消失速度定数」「胃内容排出速度（GER）」）。
- 計算は途中式を省略せず、数値を具体的に示す。
- 前文は概念の整理・公式・前提を提示し、個別解説の土台を作る。

【良い解説の例①：知識問題（第110回 問173）】
正解：３、５
前文：「薬物の尿中排泄は、糸球体におけるろ過、尿細管における分泌、再吸収という三つの過程によって行われる。血液中に含まれる薬物のうちタンパクと結合していない非結合形の薬物が糸球体でろ過を受ける。次に、近位尿細管においてトランスポーター等によって認識される薬物は分泌を受け、尿細管に流入する。その後、尿細管内に流入した薬物のうち再吸収を受けなかった薬物が尿中に排泄される。このうち、糸球体ろ過における単位時間あたりの血漿のろ過量を糸球体ろ過速度（GFR）といい、通常成人では約100 mL/min/1.73 m2である。血漿中の薬物のうち、糸球体でろ過を受けるのは非結合形の薬物のみなので、ろ過クリアランスは、GFR×fp（fp：血漿タンパク非結合率）で表される。イヌリンは、血漿タンパクと結合せず（fp=1）、尿細管分泌や再吸収を受けない。したがって、イヌリンのクリアランスは、ろ過クリアランスと一致し、GFRと等しくなる。一方、クレアチニンは血漿タンパクと結合せず、主に糸球体においてろ過され再吸収をほとんど受けないが、尿細管において若干の分泌を受ける。そのため、クレアチニンのクリアランスはイヌリンのクリアランスよりも大きくなる。」
選択肢１：「誤：GFRは、イヌリンクリアランスと等しくなる。よって、本問の患者におけるGFRは30 mL/min/1.73 m2と推定できる。」
選択肢２：「誤：イヌリンは尿細管で分泌や再吸収を受けない。そのため、再吸収に関わるクリアランスは存在しない。」
選択肢３：「正：イヌリンとクレアチニンの腎排泄の差は、尿細管分泌の有無であるため、クレアチニンの尿細管分泌クリアランスは、クレアチニンクリアランスとイヌリンクリアランスの差で表される。クレアチニンの尿細管分泌クリアランス＝50－30＝20 mL/min/1.73 m2と推定できる。」
選択肢４：「誤：糸球体ろ過速度の正常値は通常成人で約100 mL/min/1.73 m2である。本問の患者のイヌリンクリアランスは30 mL/min/1.73 m2であるため、糸球体ろ過速度が正常なときよりも、本患者のイヌリンクリアランスは小さいと考えられる。」
選択肢５：「正：本問の患者のクレアチニンクリアランスは50 mL/min/1.73 m2であるため、糸球体ろ過速度が正常なときよりも、本患者のクレアチニンクリアランスは小さいと考えられる。」

【良い解説の例②：グラフ問題（第111回 問170）】
正解：１、５
前文：「リボフラビンは、食事の有無による胃内容排出速度（GER）の違いにより吸収量が変化する。空腹時に服用するとGERの増大により、十二指腸に存在する吸収トランスポーターが飽和しやすくなり、吸収量が低下する。一方、食後に服用するとGERの低下により、吸収トランスポーターの飽和が起こりにくくなることで吸収量が増大する。なお、グラフの縦軸は累積尿中排泄量を示しているが、経口投与した薬物が尿中に排泄されるためには吸収されることが必要であるため、累積尿中排泄量≒吸収量と考える必要がある。そのため、Aが朝食後服用（吸収量が多い）、Bが空腹時服用（吸収量が少ない）と読み取れる。」
選択肢１：「正：上記参照。」
選択肢２：「誤：上記参照。」
選択肢３：「誤：BがAより低値となるのは、リボフラビンのGERが空腹時に増大することによる、吸収トランスポーターの飽和が原因である。」
選択肢４：「誤：選択肢３参照。」
選択肢５：「正：メトクロプラミドは、ドパミンD2受容体遮断薬であり、D2受容体を遮断することでコリン作動性神経を興奮させ、GERを増大させる作用をもつ。そのため、メトクロプラミドを前投与した時の曲線は、空腹時に服用した際のグラフと類似するため、AよりBに近くなる。」

【記号・表記の統一ルール】
- 投与量：D（一般）、Dpo（経口投与）、Div（静脈内投与）。"Dose"は使わない。
- 消失速度定数：kel（keは使わない）
- 吸収速度定数：ka
- 分布容積：Vd
- クリアランス：CL（CLtotは使わない）
- バイオアベイラビリティ：F
- 最高血中濃度：Cmax
- 定常状態血中濃度：Css（平均はCss,av）
- 消失半減期：t1/2（T1/2は使わない）
- 投与間隔：τ
- 平均滞留時間：MRT、平均吸収時間：MAT、平均溶出時間：MDT
- 式番号：①②③…を用いて相互参照する
- 単位：h（時間）、L（リットル）、mg、h⁻¹ など

マークアップ（数式・変数名に使用）:
{{CL:r}}=CLr(TNR斜体+下付き) {{f:e}}=fe(Century斜体+下付き)
{{K:sp}}=Ksp {{t:1/2}}=t1/2 {{Vd}}=Vd {{mu}}=µ {{-}}=全角マイナス
{{sup:2}}=上付き2 など"""


def call_api(question_text, api_key):
    payload = {
        "model": "claude-opus-4-8",
        "max_tokens": 4096,
        "system": SYSTEM_PROMPT,
        "messages": [{"role":"user","content":f"解説を生成してください:\n\n{question_text}"}]
    }
    req = urllib.request.Request(
        "https://api.anthropic.com/v1/messages",
        data=json.dumps(payload).encode(),
        headers={'Content-Type':'application/json',
                 'x-api-key': api_key,
                 'anthropic-version':'2023-06-01'},
        method='POST')
    with urllib.request.urlopen(req, timeout=120) as resp:
        text = json.loads(resp.read())['content'][0]['text']
        m = re.search(r'\{[\s\S]+\}', text)
        if m: return json.loads(m.group())
        raise ValueError("JSONが見つかりません")


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
                info = extract_info(doc)
                q_num = info['question_num']

                if not info['explanation_para_idxs'] or info['answer_para_idx'] is None:
                    st.error('テンプレートの構造が読み取れませんでした。ファイルを確認してください。')
                    st.stop()

                st.info(f'問{q_num} を検出しました')

                if 'AI' in mode:
                    # 問題文のみ抽出（解説プレースホルダーは除外）
                    exp_pat2 = re.compile(r'^[１２３４５][\s　]')
                    lines = []
                    for p in doc.paragraphs:
                        t = p.text.strip()
                        if not t: continue
                        # 選択肢解説行とプレースホルダーを除外
                        if exp_pat2.match(t): continue
                        if '解答' in t and t.startswith('問'): continue
                        lines.append(t)
                    # 表を整形して追加
                    for table in doc.tables:
                        lines.append('')
                        for row in table.rows:
                            cells = [c.text.replace('\n', ' ').strip() for c in row.cells]
                            if any(cells):
                                lines.append(' | '.join(cells))
                    question_text = '\n'.join(lines)
                    st.info('Claude API で解説生成中...')
                    # ファイルを再読み込み（上でread()済みのため）
                    uploaded.seek(0)
                    doc = Document(io.BytesIO(uploaded.read()))
                    info = extract_info(doc)
                    explanation = call_api(question_text, api_key)

                write_to_doc(doc, info, explanation)

                # docxをメモリに保存
                buf = io.BytesIO()
                doc.save(buf)
                buf.seek(0)

                fname = uploaded.name.replace('.docx', '_完成.docx')
                ans_str = '、'.join(explanation.get('解答', []))

                st.success(f'✅ 完成！（解答: {ans_str}）')

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
