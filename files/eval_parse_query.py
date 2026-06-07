import json
import subprocess
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
WF = ROOT / "workflows" / "Housing Assistant v12.json"

# v12: без POI / геокодирования — только комнаты, цена, площадь, этаж, район
QUERIES = [
    ("Q1", "Нужна однушка до 3 миллионов", {"rooms": [1], "price_max": 3_000_000}),
    ("Q2", "2к квартира до 5 млн", {"rooms": [2], "price_max": 5_000_000}),
    ("Q3", "3-к до 8 млн, площадь от 70 до 90 кв", {"rooms": [3], "min_area": 70, "max_area": 90, "price_max": 8_000_000}),
    ("Q4", "Студия до 2.5 млн", {"rooms": [0], "price_max": 2_500_000}),
    ("Q5", "2к не первый этаж, до 6 млн", {"rooms": [2], "price_max": 6_000_000, "not_first": True}),
    ("Q6", "Однокомнатная в Ленинском районе до 4 млн", {"rooms": [1], "price_max": 4_000_000, "districts": ["ленинский"]}),
    ("Q7", "3-к до 7 млн, не последний этаж", {"rooms": [3], "price_max": 7_000_000, "not_last": True}),
    ("Q8", "Двушка до 5.5 млн, площадь от 45 до 60 кв", {"rooms": [2], "price_max": 5_500_000, "min_area": 45, "max_area": 60}),
    ("Q9", "1 комната до 3500000", {"rooms": [1], "price_max": 3_500_000}),
    ("Q10", "4 комнаты до 10 млн", {"rooms": [4], "price_max": 10_000_000}),
]


def load_parse_code() -> str:
    data = json.loads(WF.read_text(encoding="utf-8"))
    return next(n["parameters"]["jsCode"] for n in data["nodes"] if n.get("name") == "Parse Query")


def run_node(queries: list[str]) -> list[dict]:
    body = load_parse_code().split("const x = $input.first().json;", 1)[1]
    script = (
        "function runParse(user_text, chat_id='test') {\n"
        "  const x = { chat_id, user_text };\n"
        + body
        + "\n}\n"
        + "const tests = "
        + json.dumps(queries, ensure_ascii=False)
        + ";\n"
        + "for (const t of tests) console.log(JSON.stringify(runParse(t)[0].json.filters));\n"
    )
    with tempfile.NamedTemporaryFile("w", suffix=".mjs", delete=False, encoding="utf-8") as f:
        f.write(script)
        path = f.name
    proc = subprocess.run(["node", path], capture_output=True, text=True, encoding="utf-8")
    if proc.returncode != 0:
        raise RuntimeError(proc.stderr)
    return [json.loads(line) for line in proc.stdout.strip().splitlines()]


def matches(expected: dict, got: dict) -> tuple[bool, list[str]]:
    bad = []
    if "rooms" in expected and not any(r in got.get("rooms", []) for r in expected["rooms"]):
        bad.append("rooms")
    for key in ("price_min", "price_max", "min_area", "max_area", "not_first", "not_last"):
        if key in expected and got.get(key) != expected[key]:
            bad.append(key)
    if "districts" in expected and not all(d in got.get("districts", []) for d in expected["districts"]):
        bad.append("districts")
    return not bad, bad


def main() -> None:
    texts = [q[1] for q in QUERIES]
    parsed = run_node(texts)
    ok = 0
    rows = []
    for (qid, text, exp), got in zip(QUERIES, parsed):
        good, bad = matches(exp, got)
        ok += int(good)
        rows.append({"id": qid, "text": text, "ok": good, "bad": bad, "filters": got})
    acc = ok / len(QUERIES) * 100
    out = {"parsing_accuracy_pct": round(acc, 1), "correct": ok, "total": len(QUERIES), "rows": rows}
    print(json.dumps(out, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
