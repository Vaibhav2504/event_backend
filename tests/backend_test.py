"""End-to-end backend tests for Student Event Evaluation Management System."""
import os, io, uuid, pytest, requests, openpyxl

BASE = os.environ.get("REACT_APP_BACKEND_URL", "http://localhost:8000").rstrip("/")
API = f"{BASE}/api"
ADMIN_EMAIL = "admin@evalsystem.in"
ADMIN_PW = "Admin@12345"

state = {}

def H(tok): return {"Authorization": f"Bearer {tok}"}

@pytest.fixture(scope="session")
def admin_token():
    r = requests.post(f"{API}/auth/login", json={"email": ADMIN_EMAIL, "password": ADMIN_PW})
    assert r.status_code == 200, r.text
    data = r.json()
    assert "access_token" in data and data["user"]["role"] == "super_admin"
    return data["access_token"]

def test_01_login_and_me(admin_token):
    r = requests.get(f"{API}/auth/me", headers=H(admin_token))
    assert r.status_code == 200
    assert r.json()["email"] == ADMIN_EMAIL

def test_02_logout():
    r = requests.post(f"{API}/auth/logout")
    assert r.status_code == 200

def test_03_event_crud(admin_token):
    r = requests.post(f"{API}/events", headers=H(admin_token), json={
        "event_name": "TEST_Event", "description": "d", "event_date": "2026-02-01", "status": "active"})
    assert r.status_code == 200
    eid = r.json()["id"]; state["event_id"] = eid
    assert requests.get(f"{API}/events", headers=H(admin_token)).status_code == 200
    r2 = requests.put(f"{API}/events/{eid}", headers=H(admin_token), json={
        "event_name": "TEST_Event2", "description": "x", "event_date": "2026-02-02", "status": "active"})
    assert r2.status_code == 200 and r2.json()["event_name"] == "TEST_Event2"

def test_04_parameters_weightage(admin_token):
    eid = state["event_id"]
    p1 = requests.post(f"{API}/events/{eid}/parameters", headers=H(admin_token), json={"parameter_name": "Innovation", "weightage": 40})
    assert p1.status_code == 200
    state["param1"] = p1.json()["id"]
    p2 = requests.post(f"{API}/events/{eid}/parameters", headers=H(admin_token), json={"parameter_name": "Presentation", "weightage": 60})
    assert p2.status_code == 200
    state["param2"] = p2.json()["id"]
    over = requests.post(f"{API}/events/{eid}/parameters", headers=H(admin_token), json={"parameter_name": "X", "weightage": 10})
    assert over.status_code == 400

def test_05_students_crud(admin_token):
    enr = f"TEST{uuid.uuid4().hex[:6]}"
    r = requests.post(f"{API}/students", headers=H(admin_token), json={
        "enrollment_no": enr, "student_name": "TEST_Student", "department": "CS",
        "semester": "6", "institute": "ABC", "email": "t@x.in"})
    assert r.status_code == 200
    state["student_id"] = r.json()["id"]
    assert requests.get(f"{API}/students?search=TEST_Student", headers=H(admin_token)).status_code == 200

def test_06_students_sample_and_upload(admin_token):
    r = requests.get(f"{API}/students/sample", headers=H(admin_token))
    assert r.status_code == 200 and "spreadsheet" in r.headers.get("content-type", "")
    wb = openpyxl.Workbook(); ws = wb.active
    ws.append(["enrollment_no", "student_name", "department", "semester", "institute", "email"])
    ws.append([f"TESTU{uuid.uuid4().hex[:6]}", "TEST_Up", "CS", "6", "ABC", "u@x.in"])
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    up = requests.post(f"{API}/students/upload", headers=H(admin_token),
                       files={"file": ("s.xlsx", buf, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")})
    assert up.status_code == 200 and up.json()["created"] >= 1

def test_07_evaluator_create(admin_token):
    em = f"eval_{uuid.uuid4().hex[:6]}@test.in"
    r = requests.post(f"{API}/evaluators", headers=H(admin_token), json={
        "name": "TEST_Eval", "email": em, "password": "Eval@12345",
        "department": "CS", "designation": "Prof"})
    assert r.status_code == 200
    state["eval_id"] = r.json()["id"]; state["eval_email"] = em

def test_08_assignments(admin_token):
    r = requests.post(f"{API}/assignments", headers=H(admin_token), json={
        "event_id": state["event_id"], "evaluator_id": state["eval_id"],
        "student_ids": [state["student_id"]]})
    assert r.status_code == 200 and r.json()["created"] == 1
    # dedup
    r2 = requests.post(f"{API}/assignments", headers=H(admin_token), json={
        "event_id": state["event_id"], "evaluator_id": state["eval_id"],
        "student_ids": [state["student_id"]]})
    assert r2.json()["created"] == 0

def test_09_role_based_access():
    # login as evaluator
    r = requests.post(f"{API}/auth/login", json={"email": state["eval_email"], "password": "Eval@12345"})
    assert r.status_code == 200
    tok = r.json()["access_token"]; state["eval_token"] = tok
    # forbidden
    bad = requests.post(f"{API}/events", headers=H(tok), json={"event_name": "X", "event_date": "2026-01-01"})
    assert bad.status_code == 403

def test_10_evaluator_dashboard_and_assigned():
    tok = state["eval_token"]
    d = requests.get(f"{API}/evaluator/dashboard", headers=H(tok))
    assert d.status_code == 200 and d.json()["assigned"] >= 1
    a = requests.get(f"{API}/evaluator/assigned", headers=H(tok))
    assert a.status_code == 200 and len(a.json()) >= 1

def test_11_evaluation_submit_validation():
    tok = state["eval_token"]
    # marks exceed weightage
    bad = requests.post(f"{API}/evaluations", headers=H(tok), json={
        "student_id": state["student_id"], "event_id": state["event_id"], "comments": "",
        "marks": [{"parameter_id": state["param1"], "marks": 999},
                  {"parameter_id": state["param2"], "marks": 50}]})
    assert bad.status_code == 400
    ok = requests.post(f"{API}/evaluations", headers=H(tok), json={
        "student_id": state["student_id"], "event_id": state["event_id"], "comments": "good",
        "marks": [{"parameter_id": state["param1"], "marks": 36},
                  {"parameter_id": state["param2"], "marks": 54}]})
    assert ok.status_code == 200 and ok.json()["total_marks"] == 90.0

def test_12_results_ranking(admin_token):
    r = requests.get(f"{API}/results?event_id=" + state["event_id"], headers=H(admin_token))
    assert r.status_code == 200
    rows = r.json()
    assert len(rows) >= 1
    row = rows[0]
    assert row["rank"] == 1 and row["grade"] == "A+" and row["score"] == 90.0

def test_13_reports(admin_token):
    for path in ["event-wise", "student-wise", "evaluator-wise"]:
        for fmt in ["xlsx", "pdf"]:
            r = requests.get(f"{API}/reports/{path}?fmt={fmt}", headers=H(admin_token))
            assert r.status_code == 200, f"{path} {fmt} failed"
            assert len(r.content) > 100
    for fmt in ["xlsx", "pdf"]:
        r = requests.get(f"{API}/reports/ranking?event_id={state['event_id']}&fmt={fmt}", headers=H(admin_token))
        assert r.status_code == 200 and len(r.content) > 100

def test_14_change_password_and_profile(admin_token):
    r = requests.put(f"{API}/auth/profile", headers=H(admin_token), json={"name": "Super Admin"})
    assert r.status_code == 200
    # change pw then revert
    new_tok = state["eval_token"]
    cp = requests.post(f"{API}/auth/change-password", headers=H(new_tok),
                       json={"current_password": "Eval@12345", "new_password": "Eval@99999"})
    assert cp.status_code == 200
    requests.post(f"{API}/auth/change-password",
                  headers=H(new_tok),
                  json={"current_password": "Eval@99999", "new_password": "Eval@12345"})

# ---- Iteration 3: status field + credentials sheet ----

def test_16_students_list_has_status(admin_token):
    r = requests.get(f"{API}/students", headers=H(admin_token))
    assert r.status_code == 200
    rows = r.json()
    assert len(rows) >= 1
    # find our test student
    me = next((s for s in rows if s["id"] == state["student_id"]), rows[0])
    assert "status" in me
    assert me["status"] in ("active", "inactive")


def test_17_student_status_patch(admin_token):
    sid = state["student_id"]
    r = requests.patch(f"{API}/students/{sid}/status", headers=H(admin_token), json={"status": "inactive"})
    assert r.status_code == 200, r.text
    assert r.json()["status"] == "inactive"
    # GET reflects
    g = requests.get(f"{API}/students?search=TEST_Student", headers=H(admin_token))
    rows = g.json()
    found = next((s for s in rows if s["id"] == sid), None)
    assert found and found["status"] == "inactive"
    # restore
    r2 = requests.patch(f"{API}/students/{sid}/status", headers=H(admin_token), json={"status": "active"})
    assert r2.status_code == 200 and r2.json()["status"] == "active"


def test_18_student_status_404(admin_token):
    r = requests.patch(f"{API}/students/nope-{uuid.uuid4().hex[:6]}/status", headers=H(admin_token), json={"status": "active"})
    assert r.status_code == 404


def test_19_evaluator_cannot_patch_student_status():
    tok = state["eval_token"]
    r = requests.patch(f"{API}/students/{state['student_id']}/status", headers=H(tok), json={"status": "active"})
    assert r.status_code == 403


def test_20_credentials_export(admin_token):
    r = requests.get(f"{API}/credentials/export", headers=H(admin_token))
    assert r.status_code == 200
    assert "spreadsheet" in r.headers.get("content-type", "")
    wb = openpyxl.load_workbook(io.BytesIO(r.content))
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    headers = list(rows[0])
    assert headers == ["role", "name", "email", "password", "department", "designation", "status"]
    assert len(rows) > 1
    # all rows should have password '(set)'
    emails = []
    for row in rows[1:]:
        d = dict(zip(headers, row))
        assert d["password"] == "(set)"
        emails.append((d["email"] or "").lower())
    assert ADMIN_EMAIL in emails
    assert state["eval_email"] in emails
    state["cred_xlsx"] = r.content


def test_21_credentials_export_forbidden_for_evaluator():
    tok = state["eval_token"]
    r = requests.get(f"{API}/credentials/export", headers=H(tok))
    assert r.status_code == 403


def test_22_credentials_import_forbidden_for_evaluator():
    tok = state["eval_token"]
    wb = openpyxl.Workbook(); ws = wb.active
    ws.append(["role", "name", "email", "password", "department", "designation", "status"])
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    r = requests.post(f"{API}/credentials/import", headers=H(tok),
                     files={"file": ("c.xlsx", buf, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")})
    assert r.status_code == 403


def test_23_credentials_import_roundtrip_updates_name(admin_token):
    # Open downloaded xlsx, change name on evaluator row, upload back
    wb = openpyxl.load_workbook(io.BytesIO(state["cred_xlsx"]))
    ws = wb.active
    headers = [c.value for c in ws[1]]
    name_col = headers.index("name") + 1
    email_col = headers.index("email") + 1
    new_name = f"TEST_Eval_{uuid.uuid4().hex[:4]}"
    for r in range(2, ws.max_row + 1):
        if (ws.cell(row=r, column=email_col).value or "").lower() == state["eval_email"]:
            ws.cell(row=r, column=name_col).value = new_name
            break
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    up = requests.post(f"{API}/credentials/import", headers=H(admin_token),
                       files={"file": ("c.xlsx", buf, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")})
    assert up.status_code == 200, up.text
    body = up.json()
    assert body["updated"] >= 1
    # verify in /api/evaluators
    ev = requests.get(f"{API}/evaluators", headers=H(admin_token)).json()
    me = next((e for e in ev if e.get("email") == state["eval_email"]), None)
    assert me and me["name"] == new_name


def test_24_credentials_import_password_update_then_login(admin_token):
    # Change evaluator password via credentials sheet, then login with new pw
    wb = openpyxl.load_workbook(io.BytesIO(state["cred_xlsx"]))
    ws = wb.active
    headers = [c.value for c in ws[1]]
    pw_col = headers.index("password") + 1
    email_col = headers.index("email") + 1
    new_pw = "NewEvalPw@2026"
    for r in range(2, ws.max_row + 1):
        if (ws.cell(row=r, column=email_col).value or "").lower() == state["eval_email"]:
            ws.cell(row=r, column=pw_col).value = new_pw
            break
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    up = requests.post(f"{API}/credentials/import", headers=H(admin_token),
                       files={"file": ("c.xlsx", buf, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")})
    assert up.status_code == 200 and up.json()["updated"] >= 1
    # login with new password
    lg = requests.post(f"{API}/auth/login", json={"email": state["eval_email"], "password": new_pw})
    assert lg.status_code == 200, lg.text
    # old pw should fail
    old = requests.post(f"{API}/auth/login", json={"email": state["eval_email"], "password": "Eval@12345"})
    assert old.status_code == 401
    # revert so other tests/iterations stay green
    tok = lg.json()["access_token"]
    rev = requests.post(f"{API}/auth/change-password", headers=H(tok),
                        json={"current_password": new_pw, "new_password": "Eval@12345"})
    assert rev.status_code == 200


def test_25_credentials_import_skips_new_without_password(admin_token):
    wb = openpyxl.Workbook(); ws = wb.active
    ws.append(["role", "name", "email", "password", "department", "designation", "status"])
    new_email = f"newuser_{uuid.uuid4().hex[:6]}@test.in"
    ws.append(["evaluator", "TEST_NewNoPw", new_email, "", "CS", "Prof", "active"])  # no password -> skipped
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    up = requests.post(f"{API}/credentials/import", headers=H(admin_token),
                       files={"file": ("c.xlsx", buf, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")})
    assert up.status_code == 200
    body = up.json()
    assert body["skipped"] >= 1 and body["created"] == 0
    # ensure no login possible
    lg = requests.post(f"{API}/auth/login", json={"email": new_email, "password": "anything"})
    assert lg.status_code == 401


def test_26_credentials_import_creates_new_evaluator(admin_token):
    wb = openpyxl.Workbook(); ws = wb.active
    ws.append(["role", "name", "email", "password", "department", "designation", "status"])
    new_email = f"createduser_{uuid.uuid4().hex[:6]}@test.in"
    state["created_email"] = new_email
    ws.append(["evaluator", "TEST_Created", new_email, "Created@12345", "CS", "Prof", "active"])
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    up = requests.post(f"{API}/credentials/import", headers=H(admin_token),
                       files={"file": ("c.xlsx", buf, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")})
    assert up.status_code == 200 and up.json()["created"] >= 1
    lg = requests.post(f"{API}/auth/login", json={"email": new_email, "password": "Created@12345"})
    assert lg.status_code == 200
    # cleanup
    ev = requests.get(f"{API}/evaluators", headers=H(admin_token)).json()
    me = next((e for e in ev if e.get("email") == new_email), None)
    if me:
        requests.delete(f"{API}/evaluators/{me['id']}", headers=H(admin_token))


def test_99_cleanup(admin_token):
    requests.delete(f"{API}/events/{state['event_id']}", headers=H(admin_token))
    requests.delete(f"{API}/students/{state['student_id']}", headers=H(admin_token))
    requests.delete(f"{API}/evaluators/{state['eval_id']}", headers=H(admin_token))
