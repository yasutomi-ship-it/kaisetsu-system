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
  "前文_追加行": [],
  "選択肢解説": [
    {"番号":"１","正誤":"誤","内容":"解説文"},
    {"番号":"２","正誤":"正","内容":"解説文"},
    {"番号":"３","正誤":"誤","内容":"解説文"},
    {"番号":"４","正誤":"誤","内容":"解説文"},
    {"番号":"５","正誤":"正","内容":"解説文"}
  ],
  "解答": ["２","５"]
}
マークアップ: {{CL:r}}=CLr(TNR斜体+下付き) {{f:e}}=fe(Century斜体+下付き)
{{K:sp}}=Ksp {{t:1/2}}=t1/2 {{mu}}=µ {{-}}=全角マイナス"""


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
