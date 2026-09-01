import json
from datetime import datetime, timezone

import pandas as pd

with open("payload.json", encoding="utf-8") as f:
    data = json.load(f)

items = data.get("iniciativas", [])

# Publish the raw data for the live dashboard to fetch on load (shared read path).
with open("data.json", "w", encoding="utf-8") as f:
    json.dump(
        {"items": items, "publishedAt": datetime.now(timezone.utc).isoformat()},
        f,
        ensure_ascii=False,
        indent=2,
    )

ini_rows, upd_rows = [], []

for i in items:
    ini_rows.append({
        "ID": i.get("id"), "Nombre": i.get("name"), "Área": i.get("area"),
        "Categoría": i.get("category"), "Estado": i.get("status"),
        "Prioridad": i.get("priority"), "Ejecución": i.get("executing"),
        "Fuente del dato": i.get("dataSource"), "Proceso anterior": i.get("processBefore"),
        "Tiempo antes (h)": i.get("timeBeforeH"), "Personas antes": i.get("peopleBefore"),
        "Frecuencia antes (veces/mes)": i.get("freqBeforeMonth"),
        "Tiempo después (h)": i.get("timeAfterH"), "Personas después": i.get("peopleAfter"),
        "Frecuencia después (veces/mes)": i.get("freqAfterMonth"),
        "Tipo de cambio": i.get("savingType"),
        "Métricas": ", ".join(i.get("metrics") or []),
        "Actualizado": i.get("updatedAt"),
    })
    for u in (i.get("updates") or []):
        upd_rows.append({
            "ID iniciativa": i.get("id"), "Iniciativa": i.get("name"),
            "Fecha": u.get("date"), "Nota": u.get("text"),
        })

with pd.ExcelWriter("iniciativas.xlsx", engine="openpyxl") as writer:
    pd.DataFrame(ini_rows).to_excel(writer, sheet_name="Iniciativas", index=False)
    pd.DataFrame(upd_rows).to_excel(writer, sheet_name="Seguimiento", index=False)
