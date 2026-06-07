import json
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
WORKFLOW_PATH = ROOT / "workflows" / "Housing Assistant v12.json"
DATASET_PATH = ROOT / "files" / "housing_dataset.json"


def _find_dataset_array_span(js_code: str) -> tuple[int, int]:
    
    marker = "const DATASET"
    i = js_code.find(marker)
    if i == -1:
        raise ValueError("Не найдено 'const DATASET' в jsCode узла 'Filter Dataset'")

    eq = js_code.find("=", i)
    if eq == -1:
        raise ValueError("Не найден '=' после 'const DATASET'")

    start = js_code.find("[", eq)
    if start == -1:
        raise ValueError("Не найден '[' (начало массива) после 'const DATASET ='")

    depth = 0
    in_str = False
    esc = False
    end = None

    for pos in range(start, len(js_code)):
        ch = js_code[pos]

        if in_str:
            if esc:
                esc = False
            elif ch == "\\":
                esc = True
            elif ch == '"':
                in_str = False
            continue

        if ch == '"':
            in_str = True
            continue

        if ch == "[":
            depth += 1
        elif ch == "]":
            depth -= 1
            if depth == 0:
                end = pos
                break

    if end is None:
        raise ValueError("Не удалось найти закрывающую ']' массива DATASET")

    return start, end


def main() -> None:
    dataset = json.loads(DATASET_PATH.read_text(encoding="utf-8"))
    wf = json.loads(WORKFLOW_PATH.read_text(encoding="utf-8"))

    nodes = wf.get("nodes", [])
    target = None
    for n in nodes:
        if n.get("id") == "filter-dataset" or n.get("name") == "Filter Dataset":
            target = n
            break
    if not target:
        raise ValueError("Не найден узел 'Filter Dataset' (id=filter-dataset) в workflow")

    params = target.get("parameters") or {}
    js_code = params.get("jsCode")
    if not isinstance(js_code, str) or not js_code.strip():
        raise ValueError("В узле 'Filter Dataset' отсутствует parameters.jsCode")

    start, end = _find_dataset_array_span(js_code)

    new_array = json.dumps(dataset, ensure_ascii=False)
    new_js_code = js_code[:start] + new_array + js_code[end + 1 :]

    target["parameters"]["jsCode"] = new_js_code

    WORKFLOW_PATH.write_text(
        json.dumps(wf, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(f"OK: встроил {len(dataset)} записей в {WORKFLOW_PATH}")


if __name__ == "__main__":
    main()
