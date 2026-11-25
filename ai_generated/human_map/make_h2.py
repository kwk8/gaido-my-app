import copy, sys
from pptx import Presentation
from pptx.util import Pt
src, dst = sys.argv[1], sys.argv[2]
p = Presentation(src)
RNS = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'

def dup(idx):
    s = p.slides[idx]
    n = p.slides.add_slide(s.slide_layout)
    for sh in list(n.shapes): sh._element.getparent().remove(sh._element)
    rmap = {}
    for rid, rel in s.part.rels.items():
        if 'slideLayout' in rel.reltype or 'notesSlide' in rel.reltype: continue
        rmap[rid] = n.part.relate_to(rel._target, rel.reltype)
    for el in s.shapes._spTree:
        n.shapes._spTree.append(copy.deepcopy(el))
    for e in n.shapes._spTree.iter():
        for a in ('embed', 'link', 'id'):
            v = e.get(RNS + a)
            if v in rmap: e.set(RNS + a, rmap[v])
    return n

def set_text(sh, text):
    tf = sh.text_frame
    for para in tf.paragraphs[1:]:
        para._p.getparent().remove(para._p)
    para = tf.paragraphs[0]
    runs = para.runs
    runs[0].text = text
    for r in runs[1:]: r._r.getparent().remove(r._r)

def by_name(slide, name):
    return [s for s in slide.shapes if s.name == name]

def get(slide, name):
    return by_name(slide, name)[0]

def fill_cards(slide, group_names, people):
    for i, gn in enumerate(group_names):
        g = get(slide, gn)
        if i >= len(people):
            g._element.getparent().remove(g._element); continue
        tbs = sorted([c for c in g.shapes if c.has_text_frame], key=lambda c: c.top)
        set_text(tbs[0], people[i][0]); set_text(tbs[1], people[i][1])
        tbs[0].text_frame.paragraphs[0].runs[0].font.size = Pt(18)
        if len(people[i][1]) > 16:
            tbs[1].text_frame.paragraphs[0].runs[0].font.size = Pt(9 if len(people[i][1]) > 24 else 10)

def remove(slide, *names):
    for n in names:
        for s in by_name(slide, n): s._element.getparent().remove(s._element)

def retitle(slide):
    for sh in by_name(slide, 'タイトル 2'):
        for r in sh.text_frame.paragraphs[0].runs:
            r.text = r.text.replace('上半期', '下半期')

# ページ1: 技術開発部 / AI活用チーム / 若手出張チーム
a = dup(4); retitle(a)
set_text(get(a, 'テキスト ボックス 12'), '若手出張チーム')
fill_cards(a, ['グループ化 6', 'グループ化 7'], [
    ('山本さん', '事務局のまとめ役'),
    ('王さん', '出張の引率、調整')])
set_text(get(a, 'テキスト ボックス 53'), 'AI活用チーム')
fill_cards(a, ['グループ化 15', 'グループ化 54', 'グループ化 58'], [
    ('朱さん', 'リーダー'),
    ('冨樫さん', '情通wayの方を中心に'),
    ('手塚さん', 'POエビデンスの話を中心に')])
set_text(get(a, 'テキスト ボックス 33'), '技術開発部')
fill_cards(a, ['グループ化 17', 'グループ化 18', 'グループ化 22', 'グループ化 26', 'グループ化 35'], [
    ('吉田さん', 'meiji step3のPMとphysical AIの活動をしている'),
    ('本後さん', 'meiji step3のPMとphysical AIの活動をしている'),
    ('横井さん', '監査のやり方について教えてくれる'),
    ('安藤さん', '一緒に監査をやる')])

# ページ2: Huawei / Cisco
b = dup(4); retitle(b)
set_text(get(b, 'テキスト ボックス 12'), 'Huawei')
fill_cards(b, ['グループ化 6', 'グループ化 7'], [
    ('池田さん', 'メディア系に詳しい'),
    ('片山さん', 'ネットワークに詳しい')])
set_text(get(b, 'テキスト ボックス 53'), 'Cisco')
fill_cards(b, ['グループ化 15', 'グループ化 54', 'グループ化 58'], [
    ('椿さん', 'CTC担当営業'),
    ('手塚さん', 'CTC担当営業'),
    ('小森さん', 'splunk担当')])
remove(b, '四角形: 角を丸くする 108', 'テキスト ボックス 33',
       'グループ化 17', 'グループ化 18', 'グループ化 22', 'グループ化 26', 'グループ化 35')
p.save(dst)
