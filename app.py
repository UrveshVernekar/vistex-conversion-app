import os
import shutil
import tempfile
import uuid
import jwt
import datetime
from typing import Optional

from fastapi import FastAPI, Request, Form, UploadFile, File, Depends, HTTPException, status
from fastapi.responses import HTMLResponse, RedirectResponse, FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

import converter

APP_USER = os.getenv("APP_USER", "admin")
APP_PASS = os.getenv("APP_PASS", "vistex2026")
SECRET_KEY = os.getenv("SECRET_KEY", "vistex_secret_key_2026_internal_server")
ALGORITHM = "HS256"

NEW_USERS_PASS = "Ifbgoa123$#"

USERS_DB = {
    "admin": {"password": APP_PASS, "name": "Administrator"},
    "gaurav_gupta@ifbglobal.com": {"password": NEW_USERS_PASS, "name": "Gaurav Gupta"},
    "priya_vishvas@ifbglobal.com": {"password": NEW_USERS_PASS, "name": "Priya Dhawan"},
    "lekhika_gupta@ifbglobal.com": {"password": NEW_USERS_PASS, "name": "Lekhika Gupta"},
    "siddharath_bhatia@ifbglobal.com": {"password": NEW_USERS_PASS, "name": "Siddharth Bhatia"},
    "manoj_sharma@ifbglobal.com": {"password": NEW_USERS_PASS, "name": "Manoj Sharma"},
    "kundan_jha@ifbglobal.com": {"password": NEW_USERS_PASS, "name": "Kundan Jha"},
    "arnim_saini@ifbglobal.com": {"password": NEW_USERS_PASS, "name": "Arnim Saini"},
    "aneesh_aggarwal@ifbglobal.com": {"password": NEW_USERS_PASS, "name": "Aneesh Aggarwal"},
    "dineshwar_pathania@ifbglobal.com": {"password": NEW_USERS_PASS, "name": "Dineshwar Singh Pathania"},
    "chirag_gautam@ifbglobal.com": {"password": NEW_USERS_PASS, "name": "Chirag Gautam"},
    "dinesh_verma@ifbglobal.com": {"password": NEW_USERS_PASS, "name": "Dinesh Verma"},
    "puneet_mehra@ifbglobal.com": {"password": NEW_USERS_PASS, "name": "Puneet Mehra"},
    "rahul@ifbglobal.com": {"password": NEW_USERS_PASS, "name": "Rahul Mahajan"},
    "piyush_jain@ifbglobal.com": {"password": NEW_USERS_PASS, "name": "Piyush Jain"},
    "ankit_dubey@ifbglobal.com": {"password": NEW_USERS_PASS, "name": "Ankit Dube"},
    "rishabh_bhardwaj@ifbglobal.com": {"password": NEW_USERS_PASS, "name": "Rishabh Bhardwaj"},
    "ankit@ifbglobal.com": {"password": NEW_USERS_PASS, "name": "Ankit Goyal"},
    "bhavya_sanghavi@ifbglobal.com": {"password": NEW_USERS_PASS, "name": "Bhavya Sanghvi"},
    "janmesh_shah@ifbglobal.com": {"password": NEW_USERS_PASS, "name": "Janmesh Shah"},
    "chetan_chopra@ifbglobal.com": {"password": NEW_USERS_PASS, "name": "Chetan Chopra"},
    "rohit_sanyal@ifbglobal.com": {"password": NEW_USERS_PASS, "name": "Rohit Sanyal"},
    "kousik_chatterjee@ifbglobal.com": {"password": NEW_USERS_PASS, "name": "Kousik Chatterjee"},
    "kommineni_sriharsha@ifbglobal.com": {"password": NEW_USERS_PASS, "name": "Kommineni Sriharsha"},
    "shaubhick_paul@ifbglobal.com": {"password": NEW_USERS_PASS, "name": "Shaubhick Paul"},
    "pranjal_dey@ifbglobal.com": {"password": NEW_USERS_PASS, "name": "Pranjal Dey"},
    "manish_nawani@ifbglobal.com": {"password": NEW_USERS_PASS, "name": "Manish Kumar"},
    "orissa_accounts@ifbglobal.com": {"password": NEW_USERS_PASS, "name": "Ved Prakash"},
    "prosenjit_chowdhury@ifbglobal.com": {"password": NEW_USERS_PASS, "name": "Prasenjit Chowdhury"},
    "prasanth_j@ifbglobal.com": {"password": NEW_USERS_PASS, "name": "J Prasanth"},
    "murugesa_pandian@ifbglobal.com": {"password": NEW_USERS_PASS, "name": "Murugesh.J"},
    "sailendra_singh@ifbglobal.com": {"password": NEW_USERS_PASS, "name": "Shailendra Singh"},
    "konki_sunil@ifbglobal.com": {"password": NEW_USERS_PASS, "name": "Konki Sunil"},
    "manapragada_lalita@ifbglobal.com": {"password": NEW_USERS_PASS, "name": "Manapragada Lalita"},
    "modadugu_kumar@ifbglobal.com": {"password": NEW_USERS_PASS, "name": "Chandra Kumar Modadugu"},
    "d_prabhakar@ifbglobal.com": {"password": NEW_USERS_PASS, "name": "Prabhakar Dasagani"},
    "usman_m@ifbglobal.com": {"password": NEW_USERS_PASS, "name": "Usman Maniyal"},
    "joby_jose@ifbglobal.com": {"password": NEW_USERS_PASS, "name": "Joby Jose"},
    "som@ifbglobal.com": {"password": NEW_USERS_PASS, "name": "Som Mascharak"},
    "a_jaiswal@ifbglobal.com": {"password": NEW_USERS_PASS, "name": "Ashish Jaiswal"},
    "aaron_victor@ifbglobal.com": {"password": NEW_USERS_PASS, "name": "Aaron Victor"},
    "dinesh_joshi@ifbglobal.com": {"password": NEW_USERS_PASS, "name": "Dinesh Joshi"},
    "yatish_chum@ifbglobal.com": {"password": NEW_USERS_PASS, "name": "Yatish Chum"},
    "chaitra_shetti@ifbglobal.com": {"password": NEW_USERS_PASS, "name": "Chaitra Shetti"}
}

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMP_STORAGE_DIR = os.path.join(BASE_DIR, "Temp")
os.makedirs(TEMP_STORAGE_DIR, exist_ok=True)

app = FastAPI(title="Vistex Format Converter Portal")

app.mount("/static", StaticFiles(directory=os.path.join(BASE_DIR, "static")), name="static")
templates = Jinja2Templates(directory=os.path.join(BASE_DIR, "templates"))


def create_session_token(username: str, user_name: str) -> str:
    expiration = datetime.datetime.utcnow() + datetime.timedelta(hours=24)
    payload = {"sub": username, "name": user_name, "exp": expiration}
    return jwt.encode(payload, SECRET_KEY, algorithm=ALGORITHM)


def get_current_user(request: Request) -> Optional[dict]:
    token = request.cookies.get("vistex_session")
    if not token:
        return None
    try:
        payload = jwt.decode(token, SECRET_KEY, algorithms=[ALGORITHM])
        return {"username": payload.get("sub"), "name": payload.get("name", payload.get("sub"))}
    except Exception:
        return None


@app.get("/", response_class=HTMLResponse)
async def root(request: Request):
    user = get_current_user(request)
    if user:
        return RedirectResponse(url="/dashboard", status_code=status.HTTP_302_FOUND)
    return RedirectResponse(url="/login", status_code=status.HTTP_302_FOUND)


@app.get("/login", response_class=HTMLResponse)
async def login_page(request: Request):
    user = get_current_user(request)
    if user:
        return RedirectResponse(url="/dashboard", status_code=status.HTTP_302_FOUND)
    return templates.TemplateResponse(request, "login.html")


@app.post("/api/login")
async def api_login(username: str = Form(...), password: str = Form(...)):
    uname_clean = username.strip().lower()
    user_entry = USERS_DB.get(uname_clean)
    if user_entry and user_entry["password"] == password:
        token = create_session_token(uname_clean, user_entry["name"])
        response = JSONResponse({"success": True, "message": "Login successful"})
        response.set_cookie(
            key="vistex_session",
            value=token,
            httponly=True,
            max_age=86400,
            samesite="lax"
        )
        return response
    return JSONResponse(
        status_code=status.HTTP_401_UNAUTHORIZED,
        content={"success": False, "detail": "Invalid username or password"}
    )


@app.get("/api/logout")
async def api_logout():
    response = RedirectResponse(url="/login", status_code=status.HTTP_302_FOUND)
    response.delete_cookie(key="vistex_session")
    return response


@app.get("/dashboard", response_class=HTMLResponse)
async def dashboard_page(request: Request):
    user = get_current_user(request)
    if not user:
        return RedirectResponse(url="/login", status_code=status.HTTP_302_FOUND)
    return templates.TemplateResponse(request, "dashboard.html", {"user": user})


@app.get("/manuals", response_class=HTMLResponse)
async def manuals_page(request: Request):
    user = get_current_user(request)
    if not user:
        return RedirectResponse(url="/login", status_code=status.HTTP_302_FOUND)
    return templates.TemplateResponse(request, "manuals.html", {"user": user})


@app.get("/api/manuals/download/{filename}")
async def api_manuals_download(filename: str, request: Request):
    user = get_current_user(request)
    if not user:
        raise HTTPException(status_code=401, detail="Unauthorized")

    safe_filename = os.path.basename(filename)
    file_path = os.path.join(BASE_DIR, safe_filename)

    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="Manual file not found")

    return FileResponse(
        path=file_path,
        filename=safe_filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.post("/api/convert")
async def api_convert(
    request: Request,
    file: UploadFile = File(...),
    mode: str = Form("auto")
):
    user = get_current_user(request)
    if not user:
        raise HTTPException(status_code=401, detail="Unauthorized")

    if not file.filename.lower().endswith((".xlsx", ".xls", ".csv")):
        raise HTTPException(status_code=400, detail="Invalid file format. Only .xlsx, .xls, or .csv files are supported.")

    session_id = str(uuid.uuid4())
    session_dir = os.path.join(TEMP_STORAGE_DIR, session_id)
    os.makedirs(session_dir, exist_ok=True)

    input_file_path = os.path.join(session_dir, file.filename)
    with open(input_file_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    try:
        res = converter.process_conversion(
            input_file_path=input_file_path,
            output_dir=session_dir,
            mode=mode
        )

        xlsx_url = f"/api/download/{session_id}/{res['xlsx_filename']}"
        csv_urls = [f"/api/download/{session_id}/{cf}" for cf in res['csv_filenames']]
        zip_url = f"/api/download/{session_id}/{res['zip_filename']}"

        return {
            "success": True,
            "session_id": session_id,
            "mode_used": res["mode_used"],
            "mode_label": res["mode_label"],
            "xlsx_filename": res["xlsx_filename"],
            "csv_filenames": res["csv_filenames"],
            "zip_filename": res["zip_filename"],
            "sheets": res["sheets"],
            "xlsx_url": xlsx_url,
            "csv_urls": csv_urls,
            "zip_url": zip_url
        }

    except Exception as e:
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Conversion error: {str(e)}")


@app.get("/api/download/{session_id}/{filename}")
async def api_download(session_id: str, filename: str, request: Request):
    user = get_current_user(request)
    if not user:
        raise HTTPException(status_code=401, detail="Unauthorized")

    file_path = os.path.join(TEMP_STORAGE_DIR, session_id, filename)
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="File not found")

    # Determine media type
    if filename.endswith(".xlsx"):
        media_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    elif filename.endswith(".csv"):
        media_type = "text/csv"
    elif filename.endswith(".zip"):
        media_type = "application/zip"
    else:
        media_type = "application/octet-stream"

    return FileResponse(path=file_path, filename=filename, media_type=media_type)
