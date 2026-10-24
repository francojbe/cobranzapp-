import urllib.request
import uuid
import json

filepath = r"C:\Users\franc\Downloads\mora.xls"
boundary = uuid.uuid4().hex

with open(filepath, "rb") as f:
    file_bytes = f.read()

part1 = f'--{boundary}\r\nContent-Disposition: form-data; name="file"; filename="mora.xls"\r\nContent-Type: application/vnd.ms-excel\r\n\r\n'.encode('utf-8')
part2 = f'\r\n--{boundary}--\r\n'.encode('utf-8')
body = part1 + file_bytes + part2

req = urllib.request.Request(
    "http://127.0.0.1:8000/api/upload",
    data=body,
    headers={"Content-Type": f"multipart/form-data; boundary={boundary}"}
)

with urllib.request.urlopen(req) as resp:
    res = json.loads(resp.read().decode("utf-8"))
    print("1. UPLOAD SUCCESS! HTTP Code:", resp.status)
    print("   Message:", res.get("message"))
    print("   Rows loaded:", res["metadata"]["total_rows"])
    print("   Unique clients:", res["metadata"]["total_clientes"])
    print("   Total amount:", res["totals"]["monto_total"])

# 2. Test filter processing (1 - 60 days)
filter_payload = json.dumps({"min_dias": 1, "max_dias": 60, "group_by": "Programa"}).encode("utf-8")
req_proc = urllib.request.Request("http://127.0.0.1:8000/api/process", data=filter_payload, headers={"Content-Type": "application/json"})
with urllib.request.urlopen(req_proc) as resp:
    proc_res = json.loads(resp.read().decode("utf-8"))
    print("2. PROCESS SUCCESS! Filtered records:", proc_res["total_filtered"])
    print("   Filtered amount:", proc_res["kpis"]["monto_total"])

# 3. Test PDF generation
with urllib.request.urlopen("http://127.0.0.1:8000/api/download/pdf") as resp:
    pdf_bytes = resp.read()
    print("3. PDF GENERATION SUCCESS! Bytes:", len(pdf_bytes))

# 4. Test Excel generation
with urllib.request.urlopen("http://127.0.0.1:8000/api/download/excel") as resp:
    excel_bytes = resp.read()
    print("4. EXCEL GENERATION SUCCESS! Bytes:", len(excel_bytes))

print("\n>>> ALL TESTS PASSED SUCCESSFULLY! <<<")
