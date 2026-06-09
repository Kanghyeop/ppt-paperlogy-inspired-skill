# -*- coding: utf-8 -*-
"""ppt-skill 빌드 CLI: 슬라이드 스펙 JSON -> 편집 가능한 .pptx (네이티브 차트/표/도형).

사용법:
  python scripts/build.py <spec.json> --out <output.pptx> [--config config.json]

spec.json 형식:
{
  "meta": {                       # 표지/감사 (자동 생성). 생략 시 기본값
    "eyebrow": "...",             # 표지 상단 라벨(영문 권장)
    "title1": "큰 제목 1줄",        # 또는 "title"
    "title2": "큰 제목 2줄(선택)",
    "subtitle": "표지 부제",
    "foot": "날짜 · 출처/작성",
    "cover": true,                # false면 표지 생략
    "thanks": true                # false면 감사 생략, 또는 {title,subtitle,foot}
  },
  "slides": [ {헤더/제목/부제/캡션/block ...}, ... ]   # 본문 슬라이드
}
"""
import sys, os, json, argparse
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import render

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

def load_config(path):
    cfg = {"accent": "1456F0", "template": os.path.join(ROOT, "assets", "ppt-template.pptx")}
    if path and os.path.exists(path):
        cfg.update(json.load(open(path, encoding="utf-8")))
    # 상대경로 -> skill 루트 기준 절대경로
    if not os.path.isabs(cfg["template"]):
        cfg["template"] = os.path.join(ROOT, cfg["template"])
    return cfg

def assemble(spec):
    """spec(dict 또는 list) -> render용 slides_data 리스트(표지+본문+감사)."""
    if isinstance(spec, list):
        meta, slides = {}, spec
    else:
        meta, slides = spec.get("meta", {}), spec.get("slides", [])
    out = []
    # 표지
    if meta.get("cover", True) and (meta.get("title1") or meta.get("title") or meta.get("title2")):
        t = meta.get("title1") or meta.get("title") or ""
        out.append({"cover": True, "eyebrow": meta.get("eyebrow", ""),
                    "title1": t, "title2": meta.get("title2", ""),
                    "subtitle": meta.get("subtitle", ""), "foot": meta.get("foot", "")})
    # 본문
    out.extend(slides)
    # 감사
    thx = meta.get("thanks", True)
    if thx:
        d = {"thanks": True}
        if isinstance(thx, dict):
            d.update(thx)
        d.setdefault("title", "감사합니다")
        d.setdefault("foot", meta.get("foot", ""))
        out.append(d)
    return out

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("spec")
    ap.add_argument("--out", required=True)
    ap.add_argument("--config", default=os.path.join(ROOT, "config.json"))
    args = ap.parse_args()

    cfg = load_config(args.config)
    spec = json.load(open(args.spec, encoding="utf-8"))
    slides_data = assemble(spec)

    # 이미지 블록의 상대경로를 spec 파일 위치 기준 절대경로로 해석
    spec_dir = os.path.dirname(os.path.abspath(args.spec))
    for d in slides_data:
        blk = d.get("block")
        if blk and blk.get("type") == "image" and blk.get("path") and not os.path.isabs(blk["path"]):
            blk["path"] = os.path.join(spec_dir, blk["path"])

    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    render.build(slides_data, cfg["template"], args.out, accent=cfg.get("accent"))

    n_content = len([d for d in slides_data if not d.get("cover") and not d.get("thanks")])
    print("OK -> %s | 본문 %d장 (+표지/감사), accent #%s" % (args.out, n_content, cfg.get("accent")))

if __name__ == "__main__":
    main()
