#!/usr/bin/env python3
"""Вспомогательные инструменты для навыка delo.

  extract <файл.pdf|docx|txt> <out.txt>
      извлечь текст документа

  build <out.docx> <папка_дела> [<папка_дела> ...] [--qa qa.json] [--title "..."] [--copy] [--print]
      проверить нормы по тексту и собрать ВСЕ дела в один Word-файл по шаблону template.docx.
      В каждой папке дела: analysis.json и source.txt.
"""
import argparse
import copy
import io
import json
import re
import subprocess
import sys
from pathlib import Path

NA = "Не указано в документе"
TEMPLATE = Path(__file__).parent / "template.docx"


# ---------- extract ----------

def extract_text(path: Path) -> str:
    ext = path.suffix.lower()
    data = path.read_bytes()
    if ext == ".pdf":
        from pypdf import PdfReader
        reader = PdfReader(io.BytesIO(data))
        if reader.is_encrypted:
            reader.decrypt("")
        text = "\n".join((p.extract_text() or "") for p in reader.pages)
    elif ext == ".docx":
        import docx
        d = docx.Document(io.BytesIO(data))
        parts = [p.text for p in d.paragraphs]
        for t in d.tables:
            for row in t.rows:
                parts.append(" | ".join(c.text.strip() for c in row.cells))
        text = "\n".join(parts)
    elif ext in (".txt", ".md"):
        text = data.decode("utf-8", errors="replace")
    elif ext == ".doc":
        sys.exit("ОШИБКА: формат .doc не поддерживается — сохраните как .docx или PDF.")
    else:
        sys.exit("ОШИБКА: поддерживаются PDF, .docx, .txt.")
    text = re.sub(r"[ \t]+", " ", text)
    lines = [l for l in text.split("\n") if l.strip()]
    if lines and sum(len(l) for l in lines) / len(lines) < 20:
        # PDF отдал текст «по слову на строку» — склеиваем в сплошной текст
        text = re.sub(r"\s+", " ", text)
        text = re.sub(r"\s+((?:ОПИСАТЕЛЬНАЯ|МОТИВИРОВОЧНАЯ|РЕЗОЛЮТИВНАЯ) ЧАСТЬ:|УСТАНОВИЛ[А]?:|РЕШИЛ:|ПОСТАНОВИЛ[А]?:|ТРЕБОВАНИЕ ИСТЦА:)", r"\n\n\1", text)
    text = re.sub(r"\n{3,}", "\n\n", text).strip()
    if len(text) < 200:
        sys.exit("ОШИБКА: в файле почти нет текста (вероятно, скан без текстового слоя). OCR не поддерживается.")
    return text


def cmd_extract(src, out):
    text = extract_text(Path(src))
    Path(out).write_text(text, encoding="utf-8")
    print(f"OK: {len(text)} символов → {out}")


# ---------- проверка норм ----------

def _norm(s):
    return re.sub(r"\s+", " ", str(s).replace("«", '"').replace("»", '"')).strip().lower()


def verify_norms(r, source):
    """Помечает n['verified'], если цитата или номер статьи реально есть в документе."""
    src = _norm(source)
    bad = []
    for n in r.get("norms") or []:
        quote = _norm(n.get("quote", ""))
        quote_ok = len(quote) >= 10 and quote in src
        num = re.search(r"\d+(?:-\d+)?", str(n.get("article", "")))
        article_ok = False
        if num:
            pat = (r"(?:ст\.?|статьи|статья|статье|статьей|статьёй|статей|статьям|статьями|п\.?|пункт\w*)\s*"
                   + re.escape(num.group(0)) + r"(?!\d)")
            article_ok = re.search(pat, src) is not None
            if not article_ok:  # перечисления вида «статьями 272, 292»
                article_ok = re.search(r"стать\w*\s+(?:[\d\-]+\s*,\s*)*" + re.escape(num.group(0)) + r"(?!\d)", src) is not None
        n["verified"] = bool(quote_ok or article_ok)
        if not n["verified"]:
            bad.append(f"{n.get('act', '')} {n.get('article', '')}".strip())
    return bad


# ---------- содержимое: блоки (общие для Word и Markdown) ----------
# блок = (вид, [(текст, жирный|None), ...]); виды: h1 h2 h3 ital body b0 b1

def _na(v):
    return not v or (isinstance(v, str) and v.strip().lower().startswith("не указано"))


def _t(v):
    return NA if _na(v) else str(v).strip()


def _plain(t):
    return (t, None)


def _bold(t):
    return (t, True)


def case_blocks(r, n):
    B = []
    title = r.get("topic") or r.get("case_title") or "Без названия"
    B.append(("h2", [_plain(f"Дело №{n}. {title}")]))
    meta = ", ".join(x for x in (
        r.get("court"), r.get("case_number") and f"дело № {r['case_number']}",
        r.get("decision_date"), r.get("document_type")) if not _na(x))
    if meta:
        B.append(("ital", [_plain(f"({meta})")]))

    B.append(("h3", [_plain("1. Фабула дела")]))
    B.append(("body", [_plain(_t(r.get("fabula")))]))

    pl = r.get("plaintiff") or {}
    pl_label = r.get("plaintiff_label") or "Истец"
    df_label = r.get("defendant_label") or "Ответчик"
    B.append(("h3", [_plain(f"2. {pl_label}")]))
    B.append(("b0", [_bold("Кто: "), _plain(_t(pl.get("who")))]))
    B.append(("b0", [_bold("Чего требует: "), _plain(_t(pl.get("demands")))]))
    grounds = pl.get("grounds") or []
    if grounds:
        B.append(("b0", [_bold("Почему считает права нарушенными:")]))
        B += [("b1", [_plain(str(g))]) for g in grounds]
    else:
        B.append(("b0", [_bold("Почему считает права нарушенными: "), _plain(NA)]))

    df = r.get("defendant") or {}
    B.append(("h3", [_plain(f"3. {df_label}")]))
    B.append(("b0", [_bold("Кто: "), _plain(_t(df.get("who")))]))
    B.append(("b0", [_bold("Позиция и доводы: "), _plain(_t(df.get("position")))]))
    if not _na(r.get("third_parties")):
        B.append(("b0", [_bold("Иные участники: "), _plain(str(r["third_parties"]).strip())]))

    dec = r.get("decision") or {}
    B.append(("h3", [_plain("4. Решение суда")]))
    if dec.get("intro"):
        B.append(("body", [_plain(str(dec["intro"]))]))
    pts = dec.get("points") or []
    B += [("b0", [_plain(str(p))]) for p in pts] if pts else [("body", [_plain(NA)])]

    B.append(("h3", [_plain("5. Правовое обоснование")]))
    norms = [n_ for n_ in (r.get("norms") or []) if isinstance(n_, dict)]
    if not norms:
        B.append(("body", [_plain(NA)]))
    for nm in norms:
        head = ", ".join(x for x in (nm.get("article"), nm.get("act")) if x)
        text = str(nm.get("explanation") or nm.get("role") or "").strip()
        segs = [_bold(head + ": "), _plain(text)]
        if not nm.get("verified"):
            segs.append(_bold("  ⚠ не найдена в тексте документа — проверьте"))
        B.append(("b0", segs))

    B.append(("h3", [_plain("6. Готовый устный ответ (1–2 минуты)")]))
    oral = r.get("oral_answer") or []
    if isinstance(oral, str):
        oral = [oral]
    B += [("body", [_plain(str(p))]) for p in oral] if oral else [("body", [_plain(NA)])]

    if r.get("warnings"):
        B.append(("h3", [_plain("Примечания")]))
        B += [("b0", [_plain(str(w))]) for w in r["warnings"]]
    return B


def build_blocks(cases, qa, title):
    B = [("h1", [_plain(title)])]
    for i, r in enumerate(cases, 1):
        B += case_blocks(r, i)
    if len(cases) > 1:
        B.append(("h2", [_plain("Сводная шпаргалка: Главная суть по каждому делу")]))
        for i, r in enumerate(cases, 1):
            short = r.get("short_name") or r.get("topic") or r.get("case_title") or ""
            B.append(("body", [_bold(f"Дело №{i} ({short}): "), _plain(_t(r.get("summary")))]))
    if qa:
        B.append(("h2", [_plain("Блок вопросов и ответов для преподавателя")]))
        for i, item in enumerate(qa, 1):
            B.append(("body", [_bold(f"{i}. Вопрос: {item.get('q', '')}")]))
            B.append(("body", [_bold("Ответ: "), _plain(str(item.get("a", "")))]))
    return B


# ---------- рендер: Markdown ----------

def to_markdown(blocks):
    L, prev = [], None
    for kind, segs in blocks:
        text = "".join(f"**{t}**" if b else t for t, b in segs)
        if kind == "h1":
            L += [f"# {text}", ""]
        elif kind == "h2":
            L += ["", f"## {text}"]
        elif kind == "h3":
            L += ["", f"### {text}"]
        elif kind == "ital":
            L += [f"_{text}_"]
        elif kind in ("b0", "b1"):
            if prev not in ("b0", "b1"):
                L.append("")
            L.append(("  " if kind == "b1" else "") + f"- {text}")
        else:
            L += ["", text]
        prev = kind
    return "\n".join(L).strip() + "\n"


# ---------- рендер: Word по шаблону ----------

def _load_template():
    from docx import Document
    from docx.oxml.ns import qn
    d = Document(str(TEMPLATE))
    ps = d.paragraphs

    def first(pred):
        return next((p for p in ps if pred(p)), None)

    def lvl(p):
        pPr = p._p.pPr
        if pPr is None or pPr.numPr is None or pPr.numPr.ilvl is None:
            return None
        return pPr.numPr.ilvl.val

    def head(name):
        return first(lambda p: p.style.name == name)

    h3 = head("Heading 3")
    body = ps[ps.index(h3) + 1]
    ital = first(lambda p: not p.style.name.startswith("Heading") and p.runs and p.runs[0].italic)
    protos = {"h1": head("Heading 1"), "h2": head("Heading 2"), "h3": h3, "ital": ital,
              "body": body, "b0": first(lambda p: lvl(p) == 0), "b1": first(lambda p: lvl(p) == 1)}
    for k, v in protos.items():
        if v is None:
            sys.exit(f"ОШИБКА: в template.docx не нашёлся образец абзаца «{k}».")
    saved = {}
    for k, p in protos.items():
        rpr = p.runs[0]._r.rPr if p.runs else None
        saved[k] = (copy.deepcopy(p._p), copy.deepcopy(rpr) if rpr is not None else None)
    plain_rpr = saved["body"][1]
    body_el = d.element.body
    for el in list(body_el):
        if el.tag != qn("w:sectPr"):
            body_el.remove(el)
    return d, saved, plain_rpr


def to_docx(blocks, path):
    from docx.shared import Pt
    from docx.text.paragraph import Paragraph
    from docx.oxml.ns import qn
    d, saved, plain_rpr = _load_template()
    body_el = d.element.body
    sect = body_el.find(qn("w:sectPr"))
    n = len(blocks)
    for i, (kind, segs) in enumerate(blocks):
        proto, rpr = saved[kind]
        p_el = copy.deepcopy(proto)
        for ch in list(p_el):
            if ch.tag != qn("w:pPr"):
                p_el.remove(ch)
        sect.addprevious(p_el)
        par = Paragraph(p_el, d._body)
        is_head = kind in ("h1", "h2", "h3", "ital")
        for text, bold in segs:
            run = par.add_run(text)
            base = rpr if is_head else plain_rpr
            if base is not None:
                if run._r.rPr is not None:
                    run._r.remove(run._r.rPr)
                run._r.insert(0, copy.deepcopy(base))
            if not is_head:
                run.bold = True if bold else None
        nxt = blocks[i + 1][0] if i + 1 < n else None
        if kind in ("b0", "b1") and nxt not in ("b0", "b1"):
            par.paragraph_format.space_after = Pt(12)
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    d.save(str(path))


# ---------- сборка ----------

def cmd_build(args):
    cases = []
    for cd in map(Path, args.dirs):
        r = json.loads((cd / "analysis.json").read_text(encoding="utf-8"))
        bad = verify_norms(r, (cd / "source.txt").read_text(encoding="utf-8"))
        print(f"{cd.name}: норм не найдено в тексте — {len(bad)}" + (f" ({'; '.join(bad)})" if bad else ""))
        cases.append(r)
    qa = json.loads(Path(args.qa).read_text(encoding="utf-8")) if args.qa else None
    title = args.title or ("Разбор судебного дела" if len(cases) == 1 else "Разбор судебной практики Республики Казахстан")
    blocks = build_blocks(cases, qa, title)
    to_docx(blocks, args.out)
    print(f"Готово: {args.out}")
    md = to_markdown(blocks)
    if args.copy:
        try:
            subprocess.run(["pbcopy"], input=md.encode("utf-8"), check=True)
            print("Скопировано в буфер обмена.")
        except Exception as e:  # noqa: BLE001
            print(f"Не удалось скопировать в буфер: {e}")
    if args.print_md or len(cases) == 1:
        print("-----")
        print(md)


if __name__ == "__main__":
    a = sys.argv[1:]
    if len(a) == 3 and a[0] == "extract":
        cmd_extract(a[1], a[2])
    elif a and a[0] == "build":
        ap = argparse.ArgumentParser(prog="court_tools.py build")
        ap.add_argument("out")
        ap.add_argument("dirs", nargs="+")
        ap.add_argument("--qa")
        ap.add_argument("--title")
        ap.add_argument("--copy", action="store_true")
        ap.add_argument("--print", dest="print_md", action="store_true")
        cmd_build(ap.parse_args(a[1:]))
    else:
        sys.exit(__doc__)
