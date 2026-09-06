import os
import openpyxl
import json
import base64

def get_base64_image(file_path):
    if not os.path.exists(file_path):
        return ""
    ext = os.path.splitext(file_path)[1].lower().replace(".", "")
    if ext == "jpg": ext = "jpeg"
    with open(file_path, "rb") as f:
        encoded = base64.b64encode(f.read()).decode("utf-8")
    return f"data:image/{ext};base64,{encoded}"

def build_portal():
    folder = r"G:\Exium\2026\Award\May-Jun"
    
    # 1. Base64 Encode Logos & Award Criteria
    jun_png_path = os.path.join(folder, "Jun_Award_Criteria_p1.png")
    jun_pdf_path = os.path.join(folder, "Jun Award Criteria.pdf")
    if not os.path.exists(jun_png_path) and os.path.exists(jun_pdf_path):
        try:
            import fitz
            doc = fitz.open(jun_pdf_path)
            if len(doc) > 0:
                pix = doc[0].get_pixmap(dpi=150)
                pix.save(jun_png_path)
        except Exception as e:
            print("PDF conversion warning:", e)

    criteria_images = {
        "may": get_base64_image(os.path.join(folder, "Award Criteria (May'26).png")),
        "june": get_base64_image(jun_png_path)
    }

    logos = {
        "main_logo": get_base64_image(os.path.join(folder, "Exium MUPS Logo.png")),
        "aarong": get_base64_image(os.path.join(folder, "Brand Logo", "Aarong Logo.png")),
        "apex": get_base64_image(os.path.join(folder, "Brand Logo", "Apex Logo.png")),
        "bata": get_base64_image(os.path.join(folder, "Brand Logo", "Bata logo.jpg")),
        "bestbuy": get_base64_image(os.path.join(folder, "Brand Logo", "Best Buy Logo.png")),
        "infinity": get_base64_image(os.path.join(folder, "Brand Logo", "Infinity Logo.png")),
        "catseye": get_base64_image(os.path.join(folder, "Brand Logo", "Cats Eye Logo.jpeg"))
    }
    print("Encoded all logos and award criteria images to base64.")

    # 2. Parse FF list.xlsx (Current Master Hierarchy)
    wb_ff = openpyxl.load_workbook(os.path.join(folder, "FF list.xlsx"), data_only=True)
    ws_ff = wb_ff["FF-RPL"]

    ff_by_mio = {}
    ff_by_terr = {}
    zones_map = {}
    regions_map = {}

    for r in range(2, ws_ff.max_row + 1):
        terr_code = str(ws_ff.cell(r, 1).value or "").strip()
        terr_name = str(ws_ff.cell(r, 2).value or "").strip()
        mio_code = str(ws_ff.cell(r, 3).value or "").strip()
        mio_name = str(ws_ff.cell(r, 4).value or "").strip()
        desig = str(ws_ff.cell(r, 5).value or "").strip()
        sap_reg = str(ws_ff.cell(r, 6).value or "").strip()
        reg_name = str(ws_ff.cell(r, 7).value or "").strip()
        rh = str(ws_ff.cell(r, 8).value or "").strip()
        sap_zone = str(ws_ff.cell(r, 9).value or "").strip()
        zone = str(ws_ff.cell(r, 10).value or "").strip()
        zh = str(ws_ff.cell(r, 11).value or "").strip()

        info = {
            "current_terr_code": terr_code,
            "current_terr_name": terr_name,
            "sap_mio_code": mio_code,
            "current_mio_name": mio_name,
            "desig": desig,
            "current_sap_reg_code": sap_reg,
            "current_region": reg_name,
            "current_rh": rh,
            "current_sap_zone_code": sap_zone,
            "current_zone": zone,
            "current_zh": zh
        }
        if mio_code and mio_code not in ["None", "0", ""]:
            ff_by_mio[mio_code] = info
        if terr_code:
            ff_by_terr[terr_code] = info

        # Zones map
        if sap_zone and sap_zone not in zones_map:
            zones_map[sap_zone] = {
                "sap_zone_code": sap_zone,
                "zone_name": zone,
                "zonal_head": zh,
                "regions": set()
            }
        if sap_zone:
            zones_map[sap_zone]["regions"].add(reg_name)

        # Regions map
        if sap_reg and sap_reg not in regions_map:
            regions_map[reg_name] = {
                "region_name": reg_name,
                "sap_region_code": sap_reg,
                "regional_head": rh,
                "zone_name": zone,
                "sap_zone_code": sap_zone,
                "zonal_head": zh,
                "may_count": 0,
                "jun_count": 0
            }

    print(f"Loaded FF List: {len(ff_by_mio)} MIOs, {len(ff_by_terr)} Territories, {len(regions_map)} Regions, {len(zones_map)} Zones.")

    # 3. Extract and Map May 2026 Achievers to Current FF Positions
    wb_may = openpyxl.load_workbook(os.path.join(folder, "Exium Award Achiever List_May 2026.xlsx"), data_only=True)
    ws_may = wb_may["Award_May"]
    may_achievers = []
    may_transferred = 0

    for r in range(2, ws_may.max_row + 1):
        ac = str(ws_may.cell(r, 4).value or "").strip()
        an = str(ws_may.cell(r, 5).value or "").strip()
        mc = str(ws_may.cell(r, 6).value or "").strip()
        mn = str(ws_may.cell(r, 7).value or "").strip()
        reg = str(ws_may.cell(r, 2).value or "").strip()
        z = str(ws_may.cell(r, 1).value or "").strip()
        rh = str(ws_may.cell(r, 3).value or "").strip()
        rx = int(ws_may.cell(r, 8).value or 0)
        ach = ws_may.cell(r, 9).value
        ach_str = f"{ach:.1f}%" if isinstance(ach, (int, float)) else str(ach or "")
        award = int(ws_may.cell(r, 10).value or 0)

        if not ac and not mn: continue

        ff_info = ff_by_mio.get(mc) or ff_by_terr.get(ac)
        is_transferred = False
        transfer_details = None

        if ff_info:
            curr_terr = ff_info["current_terr_code"]
            curr_reg = ff_info["current_region"]
            if curr_terr != ac or curr_reg != reg:
                is_transferred = True
                may_transferred += 1
                transfer_details = {
                    "prev_area_code": ac,
                    "prev_area_name": an,
                    "prev_region": reg,
                    "prev_rh": rh,
                    "prev_zone": z
                }
            curr_zone = ff_info["current_zone"]
            curr_sap_zone = ff_info["current_sap_zone_code"]
            curr_zh = ff_info["current_zh"]
            curr_reg_name = ff_info["current_region"]
            curr_sap_reg = ff_info["current_sap_reg_code"]
            curr_rh_name = ff_info["current_rh"]
            curr_area_code = ff_info["current_terr_code"]
            curr_area_name = ff_info["current_terr_name"]
            curr_desig = ff_info["desig"]
        else:
            curr_zone = z
            curr_sap_zone = ""
            curr_zh = ""
            curr_reg_name = reg
            curr_sap_reg = ""
            curr_rh_name = rh
            curr_area_code = ac
            curr_area_name = an
            curr_desig = "MIO"

        if curr_reg_name in regions_map:
            regions_map[curr_reg_name]["may_count"] += 1

        may_achievers.append({
            "id": f"MAY_{r-1}_{curr_area_code}_{mc}",
            "month": "May_2026",
            "month_label": "May 2026",
            "zone": curr_zone,
            "sap_zone_code": curr_sap_zone,
            "zonal_head": curr_zh,
            "region": curr_reg_name,
            "sap_region_code": curr_sap_reg,
            "regional_head": curr_rh_name,
            "area_code": curr_area_code,
            "area_name": curr_area_name,
            "mio_code": mc,
            "mio_name": mn,
            "desig": curr_desig,
            "rx": rx,
            "metric_label": "Ach%",
            "metric_val": ach_str,
            "award": award,
            "is_transferred": is_transferred,
            "transfer_details": transfer_details,
            "choice": "",
            "timestamp": "",
            "status": "Not Started"
        })

    # 4. Extract and Map June 2026 Achievers to Current FF Positions
    wb_jun = openpyxl.load_workbook(os.path.join(folder, "Exium Award Achiever List_June 2026.xlsx"), data_only=True)
    ws_jun = wb_jun["Award_Jun"]
    jun_achievers = []
    jun_transferred = 0

    for r in range(2, ws_jun.max_row + 1):
        ac = str(ws_jun.cell(r, 4).value or "").strip()
        an = str(ws_jun.cell(r, 5).value or "").strip()
        mc = str(ws_jun.cell(r, 6).value or "").strip()
        mn = str(ws_jun.cell(r, 7).value or "").strip()
        reg = str(ws_jun.cell(r, 2).value or "").strip()
        z = str(ws_jun.cell(r, 1).value or "").strip()
        rh = str(ws_jun.cell(r, 3).value or "").strip()
        rx = int(ws_jun.cell(r, 8).value or 0)
        growth = ws_jun.cell(r, 11).value
        growth_str = f"{growth:.1f}%" if isinstance(growth, (int, float)) else str(growth or "")
        sales_jun = ws_jun.cell(r, 10).value
        sales_str = f"{sales_jun/100000:.2f} Lac" if isinstance(sales_jun, (int, float)) else str(sales_jun or "").replace("৳", "").strip()
        award = int(ws_jun.cell(r, 12).value or 0)

        if not ac and not mn: continue

        ff_info = ff_by_mio.get(mc) or ff_by_terr.get(ac)
        is_transferred = False
        transfer_details = None

        if ff_info:
            curr_terr = ff_info["current_terr_code"]
            curr_reg = ff_info["current_region"]
            if curr_terr != ac or curr_reg != reg:
                is_transferred = True
                jun_transferred += 1
                transfer_details = {
                    "prev_area_code": ac,
                    "prev_area_name": an,
                    "prev_region": reg,
                    "prev_rh": rh,
                    "prev_zone": z
                }
            curr_zone = ff_info["current_zone"]
            curr_sap_zone = ff_info["current_sap_zone_code"]
            curr_zh = ff_info["current_zh"]
            curr_reg_name = ff_info["current_region"]
            curr_sap_reg = ff_info["current_sap_reg_code"]
            curr_rh_name = ff_info["current_rh"]
            curr_area_code = ff_info["current_terr_code"]
            curr_area_name = ff_info["current_terr_name"]
            curr_desig = ff_info["desig"]
        else:
            curr_zone = z
            curr_sap_zone = ""
            curr_zh = ""
            curr_reg_name = reg
            curr_sap_reg = ""
            curr_rh_name = rh
            curr_area_code = ac
            curr_area_name = an
            curr_desig = "MIO"

        if curr_reg_name in regions_map:
            regions_map[curr_reg_name]["jun_count"] += 1

        jun_achievers.append({
            "id": f"JUN_{r-1}_{curr_area_code}_{mc}",
            "month": "June_2026",
            "month_label": "June 2026",
            "zone": curr_zone,
            "sap_zone_code": curr_sap_zone,
            "zonal_head": curr_zh,
            "region": curr_reg_name,
            "sap_region_code": curr_sap_reg,
            "regional_head": curr_rh_name,
            "area_code": curr_area_code,
            "area_name": curr_area_name,
            "mio_code": mc,
            "mio_name": mn,
            "desig": curr_desig,
            "rx": rx,
            "metric_label": "Growth%",
            "metric_val": f"{growth_str} (Sales: {sales_str})",
            "award": award,
            "is_transferred": is_transferred,
            "transfer_details": transfer_details,
            "choice": "",
            "timestamp": "",
            "status": "Not Started"
        })

    print(f"May Achievers: {len(may_achievers)} (Transferred: {may_transferred})")
    print(f"June Achievers: {len(jun_achievers)} (Transferred: {jun_transferred})")

    # 5. Format Zones & Regions (Alphabetical by Zone Name)
    formatted_zones = []
    for sz_code, zdata in sorted(zones_map.items(), key=lambda x: x[1]["zone_name"].strip().lower()):
        reg_list = []
        for rname in sorted(list(zdata["regions"])):
            if rname in regions_map:
                reg_list.append(regions_map[rname])
        formatted_zones.append({
            "sap_zone_code": sz_code,
            "zone_name": zdata["zone_name"],
            "zonal_head": zdata["zonal_head"],
            "regions": reg_list
        })

    embedded_database = {
        "months": [
            {
                "key": "May_2026", 
                "label": "May 2026 Award", 
                "badge": "May 2026",
                "criteria": "Performance Metric: No. of Rx & Ach%", 
                "award_range": "1,000 / 2,000 Voucher",
                "total": len(may_achievers)
            },
            {
                "key": "June_2026", 
                "label": "June 2026 Award", 
                "badge": "June 2026",
                "criteria": "Performance Metric: Sales Growth% over April 2026", 
                "award_range": "2,000 / 3,000 Voucher",
                "total": len(jun_achievers)
            }
        ],
        "zones": formatted_zones,
        "regions_map": regions_map,
        "achievers_may": may_achievers,
        "achievers_jun": jun_achievers,
        "logos": logos,
        "criteria_images": criteria_images
    }

    json_db_str = json.dumps(embedded_database, ensure_ascii=False)

    html_content = generate_html_portal(json_db_str)

    out_paths = [
        os.path.join(folder, "index.html"),
        os.path.join(folder, "Award_Choice_Portal.html")
    ]
    for p in out_paths:
        with open(p, "w", encoding="utf-8") as f:
            f.write(html_content)
        print(f"Generated Portal: {p} ({os.path.getsize(p)} bytes)")

def generate_html_portal(json_db_str):
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
    <meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">
    <meta http-equiv="Pragma" content="no-cache">
    <meta http-equiv="Expires" content="0">
    <title>Exium MUPS - Award for Sr. / MIO</title>
    <!-- Tailwind CSS CDN -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- SheetJS for Client-Side Excel Export -->
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
    <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800;900&display=swap" rel="stylesheet">
    <style>
        * {{ box-sizing: border-box; }}
        html {{ scroll-behavior: smooth; }}
        body {{
            font-family: 'Plus Jakarta Sans', sans-serif;
            background-color: #f1f5f9;
            color: #1e293b;
            -webkit-tap-highlight-color: transparent;
        }}
        .brand-card {{
            transition: all 0.2s ease, box-shadow 0.2s ease;
        }}
        .brand-card:hover {{
            transform: translateY(-2px);
            box-shadow: 0 8px 16px -4px rgba(0, 0, 0, 0.08);
        }}
        .brand-card.selected {{
            border-color: #059669 !important;
            background: #f0fdf4 !important;
            box-shadow: 0 0 0 2px #059669, 0 8px 16px -4px rgba(16, 185, 129, 0.2) !important;
        }}
        ::-webkit-scrollbar {{ width: 5px; height: 5px; }}
        ::-webkit-scrollbar-track {{ background: #f8fafc; }}
        ::-webkit-scrollbar-thumb {{ background: #cbd5e1; border-radius: 4px; }}
        ::-webkit-scrollbar-thumb:hover {{ background: #94A3B8; }}
    </style>
</head>
<body class="min-h-screen flex flex-col bg-slate-100 text-slate-800 antialiased pb-24">

    <!-- Top Navigation Header (Identical to Exium MUPS Header Standard) -->
    <header class="bg-white border-b border-slate-200 shadow-sm sticky top-0 z-40">
        <div class="max-w-7xl mx-auto px-3 sm:px-6 lg:px-8 py-2 sm:py-2.5">
            <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                
                <!-- Line 1: Logo + Title + Orange Badge -->
                <div class="flex items-center justify-between sm:justify-start gap-2.5">
                    <div class="flex items-center gap-2 sm:gap-2.5 min-w-0">
                        <img id="header-logo" alt="Exium MUPS Logo" class="h-8 sm:h-10 w-auto object-contain flex-shrink-0">
                        <div class="border-l-2 border-slate-300 pl-2 sm:pl-2.5 flex items-center min-w-0">
                            <h1 class="text-sm sm:text-base md:text-lg font-black text-slate-900 tracking-tight leading-none whitespace-nowrap">
                                Award for Sr. / MIO
                            </h1>
                        </div>
                    </div>

                    <!-- Desktop Buttons -->
                    <div class="hidden sm:flex items-center gap-2">
                        <button onclick="onHeaderCloudSyncClick()" id="btn-cloud-sync" title="Live Google Sheets Real-Time Sync" class="px-3 py-1.5 bg-emerald-50 hover:bg-emerald-100 text-emerald-800 rounded-xl text-xs font-bold border border-emerald-200 transition flex items-center gap-1.5 shadow-xs">
                            <span id="header-sync-dot" class="w-2 h-2 rounded-full bg-emerald-500 animate-pulse"></span>
                            <svg id="sync-icon-desktop" class="w-3.5 h-3.5 text-emerald-600" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15"/></svg>
                            <span id="header-sync-status">Live Sync</span>
                        </button>
                        <button onclick="openZonalModal()" class="px-3.5 py-1.5 bg-orange-50 hover:bg-orange-100 text-orange-700 rounded-xl text-xs font-bold border border-orange-200 transition flex items-center gap-1.5">
                            <svg class="w-3.5 h-3.5 text-orange-500" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 21V5a2 2 0 00-2-2H7a2 2 0 00-2 2v16m14 0h2m-2 0h-5m-9 0H3m2 0h5M9 7h1m-1 4h1m4-4h1m-1 4h1m-5 10v-5a1 1 0 011-1h2a1 1 0 011 1v5m-4 0h4"/></svg>
                            <span>Zonal Head</span>
                        </button>
                        <button onclick="openAwardCriteriaModal()" class="px-3.5 py-1.5 bg-blue-50 hover:bg-blue-100 text-blue-700 rounded-xl text-xs font-bold border border-blue-200 transition flex items-center gap-1.5 shadow-xs">
                            <svg class="w-3.5 h-3.5 text-blue-600" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/></svg>
                            <span>Award Criteria</span>
                        </button>
                        <button onclick="openAdminModal()" class="px-3.5 py-1.5 bg-slate-900 hover:bg-slate-800 text-white rounded-xl text-xs font-bold shadow transition flex items-center gap-1.5">
                            <svg class="w-3.5 h-3.5 text-orange-400" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 15v2m-6 4h12a2 2 0 002-2v-6a2 2 0 00-2-2H6a2 2 0 00-2 2v6a2 2 0 002 2zm10-10V7a4 4 0 00-8 0v4h8z"/></svg>
                            <span>Admin</span>
                        </button>
                    </div>
                </div>

                <!-- Mobile Buttons -->
                <div class="flex sm:hidden items-center gap-1.5 pt-1 border-t border-slate-100">
                    <button onclick="onHeaderCloudSyncClick()" id="btn-cloud-sync-mobile" class="py-1.5 px-2 bg-emerald-50 hover:bg-emerald-100 text-emerald-800 rounded-xl text-[11px] font-bold border border-emerald-200 transition flex items-center justify-center gap-1">
                        <span id="header-sync-dot-mobile" class="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-pulse"></span>
                        <svg class="w-3.5 h-3.5 text-emerald-600" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15"/></svg>
                        <span id="header-sync-status-mobile">Sync</span>
                    </button>
                    <button onclick="openAwardCriteriaModal()" class="flex-1 py-1.5 px-2 bg-blue-50 hover:bg-blue-100 text-blue-700 rounded-xl text-[11px] font-bold border border-blue-200 transition text-center flex items-center justify-center gap-1">
                        <svg class="w-3.5 h-3.5 text-blue-600" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/></svg>
                        <span>Criteria</span>
                    </button>
                    <button onclick="openZonalModal()" class="py-1.5 px-2.5 bg-orange-50 hover:bg-orange-100 text-orange-700 rounded-xl text-[11px] font-bold border border-orange-200 transition text-center flex items-center justify-center gap-1">
                        <svg class="w-3.5 h-3.5 text-orange-500" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 21V5a2 2 0 00-2-2H7a2 2 0 00-2 2v16m14 0h2m-2 0h-5m-9 0H3m2 0h5M9 7h1m-1 4h1m4-4h1m-1 4h1m-5 10v-5a1 1 0 011-1h2a1 1 0 011 1v5m-4 0h4"/></svg>
                        <span>Zonal</span>
                    </button>
                    <button onclick="openAdminModal()" class="py-1.5 px-2.5 bg-slate-900 hover:bg-slate-800 text-white rounded-xl text-[11px] font-bold shadow transition flex items-center gap-1">
                        <svg class="w-3.5 h-3.5 text-orange-400" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 15v2m-6 4h12a2 2 0 002-2v-6a2 2 0 00-2-2H6a2 2 0 00-2 2v6a2 2 0 002 2zm10-10V7a4 4 0 00-8 0v4h8z"/></svg>
                        <span>Admin</span>
                    </button>
                </div>

            </div>
        </div>
    </header>

    <!-- Main Content Container -->
    <main class="max-w-7xl mx-auto px-3 sm:px-6 lg:px-8 py-5 sm:py-6 flex-1 w-full">

        <!-- ============================================== -->
        <!-- VIEW 1: REGIONAL / ZONAL / MIO LOGIN VIEW      -->
        <!-- ============================================== -->
        <section id="login-view" class="flex-1 flex items-center justify-center py-4 sm:py-8">
            <div class="bg-white border border-slate-200 rounded-3xl shadow-xl p-5 sm:p-8 max-w-lg w-full">
                
                <div class="text-center space-y-1.5 mb-6">
                    <div class="inline-flex p-3 rounded-2xl bg-orange-50 text-orange-600 mb-1 border border-orange-100">
                        <svg class="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 8v13m0-13V6a2 2 0 112 2h-2zm0 0V5.5A2.5 2.5 0 109.5 8H12zm-7 4h14M5 12a2 2 0 110-4h14a2 2 0 110 4M5 12v7a2 2 0 002 2h10a2 2 0 002-2v-7"/></svg>
                    </div>
                    <h2 class="text-xl sm:text-2xl font-black text-slate-900 tracking-tight">Sr. / MIO Award Choice Portal</h2>
                    <p class="text-xs sm:text-sm text-slate-500">Select your Zone and Region to assign award catalogue vouchers</p>
                </div>

                <!-- Tabs: Regional Head Login, MIO Search -->
                <div class="flex border-b border-slate-200 mb-5 text-xs font-bold text-slate-500">
                    <button id="tab-reg" onclick="switchLoginTab('regional')" class="flex-1 py-2.5 text-center border-b-2 border-orange-500 text-orange-600 font-black">
                        Regional Head Login
                    </button>
                    <button id="tab-mio" onclick="switchLoginTab('mio')" class="flex-1 py-2.5 text-center border-b-2 border-transparent hover:text-slate-800">
                        MIO Search
                    </button>
                </div>

                <!-- 1. Regional Head Login Form -->
                <form id="form-regional" onsubmit="handleRegionalLogin(event)" class="space-y-4">
                    <div>
                        <label class="block text-xs font-bold text-slate-700 uppercase tracking-wider mb-1.5">
                            1. Select Zone
                        </label>
                        <select id="reg-zone-select" onchange="onRegionalZoneChange()" class="w-full bg-slate-50 border border-slate-300 rounded-xl px-3.5 py-2.5 text-xs sm:text-sm font-semibold text-slate-800 focus:ring-2 focus:ring-orange-500 outline-none cursor-pointer">
                            <option value="">-- Choose Zone --</option>
                        </select>
                    </div>

                    <div>
                        <label class="block text-xs font-bold text-slate-700 uppercase tracking-wider mb-1.5">
                            2. Select Region
                        </label>
                        <select id="reg-region-select" disabled onchange="onRegionalRegionChange()" class="w-full bg-slate-100 border border-slate-300 rounded-xl px-3.5 py-2.5 text-xs sm:text-sm font-semibold text-slate-800 focus:ring-2 focus:ring-orange-500 outline-none cursor-pointer disabled:cursor-not-allowed">
                            <option value="">-- Select Zone First --</option>
                        </select>
                    </div>

                    <!-- Region Preview Details -->
                    <div id="reg-details-box" class="hidden bg-slate-50 border border-slate-200 rounded-xl p-3.5 text-xs space-y-1.5">
                        <div class="flex justify-between items-center">
                            <span class="text-slate-500 font-bold">Regional Head:</span>
                            <span id="disp-reg-head" class="font-black text-slate-900">-</span>
                        </div>
                        <div class="flex justify-between items-center border-t border-slate-200 pt-1.5">
                            <span class="text-slate-500 font-bold">Eligible Award Achievers:</span>
                            <span id="disp-reg-count" class="font-extrabold text-orange-600 bg-orange-50 border border-orange-200 px-2 py-0.5 rounded-full">0</span>
                        </div>
                    </div>

                    <div>
                        <div class="flex justify-between items-center mb-1.5">
                            <label class="block text-xs font-bold text-slate-700 uppercase tracking-wider">
                                3. Region Password
                            </label>
                            <span class="text-[10px] text-orange-600 font-black bg-orange-50 px-1.5 py-0.5 rounded border border-orange-200">SAP Region Code</span>
                        </div>
                        <input type="password" id="reg-password-input" placeholder="Enter SAP Region Code" class="w-full bg-slate-50 border border-slate-300 rounded-xl px-3.5 py-2.5 text-xs sm:text-sm font-semibold text-slate-800 focus:ring-2 focus:ring-orange-500 outline-none">
                    </div>

                    <button type="submit" class="w-full py-3 px-4 rounded-xl font-black text-sm text-white bg-slate-900 hover:bg-slate-800 shadow-md transition transform active:scale-[0.99] flex items-center justify-center gap-2">
                        <span>Login to Region</span>
                        <svg class="w-4 h-4 text-orange-400" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M14 5l7 7m0 0l-7 7m7-7H3"/></svg>
                    </button>
                </form>

                <!-- 2. Direct MIO Search Form -->
                <form id="form-mio-search" onsubmit="searchMioDirect(event)" class="hidden space-y-3">
                    <div>
                        <label class="block text-xs font-bold text-slate-700 uppercase tracking-wider mb-1.5">
                            Search by SAP MIO Code, Area Code or Name
                        </label>
                        <div class="flex gap-2">
                            <input type="text" id="mio-quick-query" oninput="searchMioDirect(event)" placeholder="e.g. 6135 or 13027 or Name" class="flex-1 bg-slate-50 border border-slate-300 rounded-xl px-3.5 py-2.5 text-xs sm:text-sm font-semibold text-slate-800 focus:ring-2 focus:ring-orange-500 outline-none">
                            <button type="submit" class="px-4 py-2.5 bg-slate-900 hover:bg-slate-800 text-white font-bold text-xs rounded-xl transition flex items-center gap-1.5 shadow-sm">
                                <svg class="w-3.5 h-3.5 text-orange-400" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z"/></svg>
                                <span>Search</span>
                            </button>
                        </div>
                    </div>
                    <div id="mio-search-results-box" class="space-y-3 max-h-96 overflow-y-auto pr-1"></div>
                </form>

            </div>
        </section>

        <!-- ============================================== -->
        <!-- VIEW 2: REGIONAL MAIN VIEW                     -->
        <!-- ============================================== -->
        <section id="regional-view" class="hidden space-y-5">

            <!-- Locked Notification Banner (Shown when admin locks submissions) -->
            <div id="regview-locked-banner" class="hidden bg-rose-50 border-2 border-rose-300 rounded-2xl p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3 text-rose-950 shadow-sm">
                <div class="flex items-center gap-3">
                    <div class="w-10 h-10 rounded-xl bg-rose-100 border border-rose-200 flex items-center justify-center text-rose-700 text-xl font-black flex-shrink-0">
                        🔒
                    </div>
                    <div>
                        <div class="font-black text-xs sm:text-sm text-rose-950">Voucher Selection is Currently CLOSED (Read-Only Mode)</div>
                        <div class="text-[11px] text-rose-700 mt-0.5">The Central Administrator has locked submissions. Existing choices are preserved and cannot be modified.</div>
                    </div>
                </div>
                <span class="px-3 py-1 bg-rose-200 text-rose-900 text-xs font-black rounded-xl border border-rose-300 flex-shrink-0 self-start sm:self-center">Submissions Locked</span>
            </div>

            <!-- Regional Banner Card -->
            <div class="bg-white rounded-2xl p-4 sm:p-6 shadow-sm border border-slate-200 flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                <div>
                    <div class="flex items-center gap-2 flex-wrap mb-1">
                        <span id="regview-zone" class="px-2.5 py-0.5 rounded-full text-[10px] font-black uppercase tracking-wider bg-orange-100 text-orange-800 border border-orange-200">Zone</span>
                        <span id="regview-sap" class="px-2.5 py-0.5 rounded-full text-[10px] font-bold bg-slate-100 text-slate-600 border border-slate-200">SAP: 00000</span>
                    </div>
                    <h2 id="regview-name" class="text-xl sm:text-2xl font-black text-slate-900 tracking-tight">Region Name</h2>
                    <div class="text-xs text-slate-500 mt-0.5">
                        <span id="regview-rh">Regional Head: -</span> • <span id="regview-zh" class="text-orange-700 font-bold">Zonal Head: -</span>
                    </div>
                </div>

                <!-- Region Progress Bar -->
                <div class="bg-slate-50 border border-slate-200 rounded-xl p-3 sm:p-3.5 min-w-[280px] sm:min-w-[320px] space-y-2">
                    <div class="flex justify-between text-xs font-bold">
                        <span class="text-slate-600">Region Progress</span>
                        <span id="regview-progress-pct" class="text-emerald-700 font-black">0%</span>
                    </div>
                    <div class="w-full bg-slate-200 h-2 rounded-full overflow-hidden">
                        <div id="regview-progress-bar" class="bg-emerald-600 h-full w-0 transition-all duration-300"></div>
                    </div>
                    <div class="flex items-center justify-between gap-3 text-[11px] text-slate-500 font-medium">
                        <span id="regview-progress-count">0 of 0 Completed</span>
                        <button onclick="exitToLogin()" class="text-orange-600 hover:text-orange-700 hover:underline font-bold flex-shrink-0 ml-2">Change Region</button>
                    </div>
                </div>
            </div>

            <!-- REGIONAL CONTROLS & DUAL-MONTH CRITERIA REFERENCE -->
            <div class="bg-white rounded-2xl border border-slate-200 shadow-sm p-4 space-y-3.5">
                <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                    <div>
                        <h3 class="text-base font-black text-slate-900">Award Achievers</h3>
                    </div>
                    
                    <!-- Search Input & Export Region Excel -->
                    <div class="flex items-center gap-2">
                        <button onclick="exportToExcel('region')" class="px-3 py-2 bg-emerald-50 hover:bg-emerald-100 text-emerald-800 border border-emerald-300 rounded-xl text-xs font-bold flex items-center gap-1.5 transition shadow-xs whitespace-nowrap">
                            <svg class="w-3.5 h-3.5 text-emerald-700" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/></svg>
                            <span>Export Region Excel</span>
                        </button>
                        <div class="relative w-full sm:w-64">
                            <input type="text" id="reg-mio-filter-input" oninput="renderAchievers()" placeholder="Filter Territory or MIO..." class="w-full bg-slate-50 border border-slate-300 rounded-xl pl-8 pr-3.5 py-2 text-xs font-semibold text-slate-800 outline-none focus:ring-2 focus:ring-orange-500">
                            <svg class="w-4 h-4 text-slate-400 absolute left-2.5 top-2.5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z"/></svg>
                        </div>
                    </div>
                </div>

                <!-- Monthly Criteria Reference Banners -->
                <div class="grid grid-cols-1 sm:grid-cols-2 gap-2.5 text-xs">
                    <div class="rounded-xl p-2.5 bg-blue-50/70 border border-blue-200 text-blue-900 flex items-center justify-between">
                        <div>
                            <span class="font-black">May 2026:</span>
                            <span class="text-blue-800 text-[11px] ml-1">No. of Rx & Ach%</span>
                        </div>
                        <span class="font-bold text-[11px] bg-blue-100 px-2 py-0.5 rounded border border-blue-200">1,000 / 2,000</span>
                    </div>
                    <div class="rounded-xl p-2.5 bg-purple-50/70 border border-purple-200 text-purple-900 flex items-center justify-between">
                        <div>
                            <span class="font-black">June 2026:</span>
                            <span class="text-purple-800 text-[11px] ml-1">Sales Growth% over April</span>
                        </div>
                        <span class="font-bold text-[11px] bg-purple-100 px-2 py-0.5 rounded border border-purple-200">2,000 / 3,000</span>
                    </div>
                </div>
            </div>

            <!-- Achievers Cards Grid -->
            <div id="regview-achievers-container" class="space-y-5">
                <!-- Injected dynamically -->
            </div>

        </section>

        <!-- ============================================== -->
        <!-- VIEW 3: ZONAL HEAD DASHBOARD VIEW              -->
        <!-- ============================================== -->
        <section id="zonal-view" class="hidden space-y-5">

            <div class="bg-white rounded-2xl p-5 shadow-sm border border-slate-200 flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                <div>
                    <div class="flex items-center gap-2 mb-1">
                        <span class="bg-orange-500 text-white text-[10px] font-black uppercase px-2.5 py-0.5 rounded-full">Zonal Dashboard</span>
                        <span id="zoneview-sap" class="bg-slate-100 text-slate-700 text-[10px] font-mono font-bold px-2.5 py-0.5 rounded-full border border-slate-200">SAP Zone: -</span>
                    </div>
                    <h2 id="zoneview-name" class="text-xl sm:text-2xl font-black text-slate-900">Zone Name</h2>
                    <p id="zoneview-zh" class="text-xs text-slate-500">Zonal Head: -</p>
                </div>
                <div class="flex items-center gap-2">
                    <button onclick="exportToExcel('zone')" class="px-3.5 py-1.5 bg-emerald-700 hover:bg-emerald-800 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 shadow-sm transition">
                        <svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/></svg>
                        <span>Export Zone Excel</span>
                    </button>
                    <button onclick="exitToLogin()" class="px-3.5 py-1.5 bg-slate-100 hover:bg-slate-200 border border-slate-300 rounded-xl text-xs font-bold text-slate-700 transition">
                        Back to Login
                    </button>
                </div>
            </div>

            <!-- Zone KPI Summary -->
            <div class="grid grid-cols-2 sm:grid-cols-4 gap-3">
                <div class="bg-white border border-slate-200 rounded-xl p-3.5 shadow-sm text-center">
                    <div class="text-[11px] font-bold text-slate-500 uppercase">Total Achievers</div>
                    <div id="zone-stat-total" class="text-2xl font-black text-slate-900 mt-0.5">0</div>
                    <div id="zone-stat-mayjun" class="text-[10px] text-slate-500">May: 0 | June: 0</div>
                </div>
                <div class="bg-white border border-slate-200 rounded-xl p-3.5 shadow-sm text-center">
                    <div class="text-[11px] font-bold text-emerald-600 uppercase">Completed</div>
                    <div id="zone-stat-completed" class="text-2xl font-black text-emerald-700 mt-0.5">0</div>
                    <div id="zone-stat-pct" class="text-[10px] text-emerald-600 font-bold">0.0%</div>
                </div>
                <div class="bg-white border border-slate-200 rounded-xl p-3.5 shadow-sm text-center">
                    <div class="text-[11px] font-bold text-amber-600 uppercase">Pending</div>
                    <div id="zone-stat-pending" class="text-2xl font-black text-amber-700 mt-0.5">0</div>
                    <div class="text-[10px] text-slate-400">Awaiting choice</div>
                </div>
                <div class="bg-white border border-slate-200 rounded-xl p-3.5 shadow-sm text-center">
                    <div class="text-[11px] font-bold text-blue-600 uppercase">Total Regions</div>
                    <div id="zone-stat-regions" class="text-2xl font-black text-blue-900 mt-0.5">0</div>
                    <div class="text-[10px] text-blue-600">In this zone</div>
                </div>
            </div>

            <!-- Regions Breakdown Table -->
            <div class="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
                <div class="p-3.5 bg-slate-50 border-b border-slate-200 flex justify-between items-center">
                    <h3 class="font-black text-xs uppercase tracking-wider text-slate-800">Regions in this Zone</h3>
                    <span class="text-[11px] text-slate-500">Click to view region cards</span>
                </div>
                <div class="overflow-x-auto">
                    <table class="w-full text-left text-xs">
                        <thead class="bg-slate-100 text-slate-700 font-bold border-b border-slate-200">
                            <tr>
                                <th class="py-2.5 px-3">Region Name</th>
                                <th class="py-2.5 px-3">SAP Code</th>
                                <th class="py-2.5 px-3">Regional Head</th>
                                <th class="py-2.5 px-3 text-center">May</th>
                                <th class="py-2.5 px-3 text-center">June</th>
                                <th class="py-2.5 px-3 text-center">Total</th>
                                <th class="py-2.5 px-3 text-center">Completed</th>
                                <th class="py-2.5 px-3 text-right">Action</th>
                            </tr>
                        </thead>
                        <tbody id="zone-regions-tbody" class="divide-y divide-slate-100 font-medium"></tbody>
                    </table>
                </div>
            </div>

            <!-- Zone Voucher Breakdown -->
            <div class="bg-white rounded-2xl border border-slate-200 shadow-sm p-4 space-y-3">
                <h3 class="font-black text-xs uppercase tracking-wider text-slate-800">Voucher Choices in this Zone</h3>
                <div id="zone-voucher-cards-grid" class="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-2.5"></div>
            </div>

        </section>

    </main>

    <!-- Bottom Sticky Action Bar (For Regional View) -->
    <div id="sticky-bar" class="hidden fixed bottom-0 inset-x-0 bg-white/95 backdrop-blur-md border-t border-slate-200 shadow-xl py-2.5 sm:py-3 z-30">
        <div class="max-w-7xl mx-auto px-3 sm:px-6 lg:px-8 flex items-center justify-between gap-4">
            <div class="flex items-center space-x-2.5">
                <div class="w-2.5 h-2.5 rounded-full bg-emerald-500 animate-pulse"></div>
                <div class="text-xs">
                    <span class="font-black text-slate-900" id="sticky-completed-text">0 Completed</span>
                    <span class="text-slate-400 mx-1">•</span>
                    <span class="text-slate-500 font-medium" id="sticky-region-text">Region</span>
                </div>
            </div>
            <div class="flex items-center gap-2">
                <button onclick="exportToExcel('region')" class="hidden sm:flex px-3.5 py-1.5 bg-emerald-50 hover:bg-emerald-100 text-emerald-800 border border-emerald-300 rounded-xl text-xs font-bold transition items-center gap-1.5 shadow-xs">
                    <svg class="w-3.5 h-3.5 text-emerald-700" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/></svg>
                    <span>Export Excel</span>
                </button>
                <button onclick="exitToLogin()" class="px-3 py-1.5 border border-slate-300 text-slate-700 hover:bg-slate-100 rounded-xl text-xs font-bold transition">
                    Back to Login
                </button>
                <button onclick="manualSaveSync()" id="btn-save-sync" class="px-4 py-1.5 bg-slate-900 hover:bg-slate-800 text-white rounded-xl text-xs font-bold shadow transition flex items-center gap-1.5">
                    <svg class="w-3.5 h-3.5 text-emerald-400" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M8 7H5a2 2 0 00-2 2v9a2 2 0 002 2h14a2 2 0 002-2V9a2 2 0 00-2-2h-3m-1 4l-3 3m0 0l-3-3m3 3V4"/></svg>
                    <span>Save</span>
                </button>
            </div>
        </div>
    </div>

    <!-- ============================================== -->
    <!-- MODAL: AWARD CRITERIA (MAY & JUNE 2026)        -->
    <!-- ============================================== -->
    <div id="criteria-modal" class="hidden fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-3 sm:p-5">
        <div class="bg-white rounded-3xl max-w-4xl w-full max-h-[92vh] flex flex-col shadow-2xl overflow-hidden border border-slate-200">
            
            <!-- Modal Header -->
            <div class="bg-white p-4 sm:p-5 border-b border-slate-200 flex justify-between items-center flex-shrink-0">
                <div class="flex items-center space-x-2.5">
                    <div class="w-9 h-9 rounded-xl bg-blue-50 border border-blue-200 flex items-center justify-center text-blue-600 font-bold text-base">
                        📋
                    </div>
                    <div>
                        <h3 class="text-sm sm:text-base font-black text-slate-900">Official Award Criteria</h3>
                        <p class="text-[11px] text-slate-500">Exium MUPS Campaign Evaluation Criteria & Guidelines</p>
                    </div>
                </div>
                <button onclick="closeAwardCriteriaModal()" class="text-slate-400 hover:text-slate-600 p-1.5 rounded-lg hover:bg-slate-100 transition">
                    <svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"/></svg>
                </button>
            </div>

            <!-- Modal Tabs: May 2026 & June 2026 -->
            <div class="flex border-b border-slate-200 bg-slate-50 px-4 sm:px-6 pt-3 gap-2 flex-shrink-0">
                <button id="tab-crit-may" onclick="switchCriteriaTab('may')" class="px-4 py-2.5 rounded-t-xl text-xs sm:text-sm font-black transition border-b-2 border-blue-600 text-blue-700 bg-white shadow-xs">
                    May 2026 Criteria
                </button>
                <button id="tab-crit-jun" onclick="switchCriteriaTab('june')" class="px-4 py-2.5 rounded-t-xl text-xs sm:text-sm font-bold transition border-b-2 border-transparent text-slate-500 hover:text-slate-800">
                    June 2026 Criteria
                </button>
            </div>

            <!-- Modal Body (Scrollable Image Viewer) -->
            <div class="flex-1 overflow-y-auto p-4 sm:p-6 bg-slate-100/70 space-y-4 text-center">
                <!-- May Criteria View -->
                <div id="crit-view-may" class="space-y-3">
                    <div class="bg-blue-50 border border-blue-200 rounded-2xl p-3 text-xs text-blue-900 font-medium flex items-center justify-between flex-wrap gap-2 text-left">
                        <div>
                            <strong>May 2026 Campaign:</strong> Performance Metric based on No. of Rx & Ach%
                        </div>
                        <span class="font-black bg-blue-100 px-2.5 py-0.5 rounded text-[11px] text-blue-950">Award: 1,000 / 2,000</span>
                    </div>
                    <div class="bg-white p-2 sm:p-4 rounded-2xl border border-slate-200 shadow-sm flex justify-center items-center overflow-auto">
                        <img id="img-crit-may" alt="May 2026 Award Criteria" class="max-w-full h-auto rounded-xl shadow-xs object-contain cursor-zoom-in hover:opacity-95 transition" onclick="openImageFullscreen(this.src)">
                    </div>
                </div>

                <!-- June Criteria View -->
                <div id="crit-view-jun" class="hidden space-y-3">
                    <div class="bg-purple-50 border border-purple-200 rounded-2xl p-3 text-xs text-purple-900 font-medium flex items-center justify-between flex-wrap gap-2 text-left">
                        <div>
                            <strong>June 2026 Campaign:</strong> Performance Metric based on Sales Growth% over April 2026
                        </div>
                        <span class="font-black bg-purple-100 px-2.5 py-0.5 rounded text-[11px] text-purple-950">Award: 2,000 / 3,000</span>
                    </div>
                    <div class="bg-white p-2 sm:p-4 rounded-2xl border border-slate-200 shadow-sm flex justify-center items-center overflow-auto">
                        <img id="img-crit-jun" alt="June 2026 Award Criteria" class="max-w-full h-auto rounded-xl shadow-xs object-contain cursor-zoom-in hover:opacity-95 transition" onclick="openImageFullscreen(this.src)">
                    </div>
                </div>
            </div>

            <!-- Modal Footer -->
            <div class="bg-slate-50 p-3 sm:p-4 border-t border-slate-200 flex justify-between items-center flex-shrink-0">
                <span class="text-[11px] text-slate-400 italic">Click on image or download to inspect full circular</span>
                <div class="flex items-center gap-2">
                    <button onclick="downloadActiveCriteria()" class="px-3.5 py-1.5 bg-white hover:bg-slate-100 border border-slate-300 text-slate-700 rounded-xl text-xs font-bold transition flex items-center gap-1.5 shadow-xs">
                        <svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4"/></svg>
                        <span>Download</span>
                    </button>
                    <button onclick="closeAwardCriteriaModal()" class="px-4 py-1.5 bg-slate-900 hover:bg-slate-800 text-white rounded-xl text-xs font-bold transition">
                        Close
                    </button>
                </div>
            </div>

        </div>
    </div>

    <!-- ============================================== -->
    <!-- MODAL: ZONAL HEAD AUTHENTICATION              -->
    <!-- ============================================== -->
    <div id="zonal-modal" class="hidden fixed inset-0 bg-slate-900/50 backdrop-blur-sm z-50 flex items-center justify-center p-4">
        <div class="bg-white rounded-3xl max-w-md w-full flex flex-col shadow-2xl overflow-hidden border border-slate-200">
            
            <div class="bg-white p-4 sm:p-5 border-b border-slate-200 flex justify-between items-center">
                <div class="flex items-center space-x-2.5">
                    <div class="w-8 h-8 rounded-lg bg-orange-50 border border-orange-200 flex items-center justify-center text-orange-600 font-bold text-base">
                        🌐
                    </div>
                    <div>
                        <h3 class="text-sm sm:text-base font-black text-slate-900">Zonal Head Login</h3>
                        <p class="text-[11px] text-slate-500">Exium MUPS Zonal Award Dashboard</p>
                    </div>
                </div>
                <button onclick="closeZonalModal()" class="text-slate-400 hover:text-slate-600 p-1">
                    <svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"/></svg>
                </button>
            </div>

            <form id="form-zonal" onsubmit="handleZonalLogin(event)" class="p-5 sm:p-6 space-y-4">
                <div class="bg-orange-50 border border-orange-200 rounded-xl p-3 text-xs text-orange-900 font-medium">
                    Zonal Heads can login with their <strong>SAP Zone Code</strong> to inspect zone-wide award selections and completion rates.
                </div>

                <div>
                    <label class="block text-xs font-bold text-slate-700 uppercase tracking-wider mb-1.5">
                        1. Select Your Zone
                    </label>
                    <select id="zone-login-select" onchange="onZonalZoneChange()" class="w-full bg-slate-50 border border-slate-300 rounded-xl px-3.5 py-2.5 text-xs sm:text-sm font-semibold text-slate-800 focus:ring-2 focus:ring-orange-500 outline-none cursor-pointer">
                        <option value="">-- Choose Zone --</option>
                    </select>
                </div>

                <div id="zone-details-box" class="hidden bg-slate-50 border border-slate-200 rounded-xl p-3.5 text-xs space-y-1.5">
                    <div class="flex justify-between items-center"><span class="text-slate-500 font-bold">Zonal Head:</span><span id="disp-zh-name" class="font-black text-slate-900">-</span></div>
                    <div class="flex justify-between items-center"><span class="text-slate-500 font-bold">Total Regions:</span><span id="disp-zh-regcount" class="font-black text-slate-900">-</span></div>
                </div>

                <div>
                    <div class="flex justify-between items-center mb-1.5">
                        <label class="block text-xs font-bold text-slate-700 uppercase tracking-wider">
                            2. Zonal Password
                        </label>
                        <span class="text-[10px] text-orange-600 font-black bg-orange-50 px-1.5 py-0.5 rounded border border-orange-200">SAP Zone Code</span>
                    </div>
                    <input type="password" id="zonal-password-input" placeholder="Enter SAP Zone Code" class="w-full bg-slate-50 border border-slate-300 rounded-xl px-3.5 py-2.5 text-xs sm:text-sm font-semibold text-slate-800 focus:ring-2 focus:ring-orange-500 outline-none">
                </div>

                <button type="submit" class="w-full py-3 px-4 rounded-xl font-black text-sm text-white bg-slate-900 hover:bg-slate-800 shadow-md transition transform active:scale-[0.99] flex items-center justify-center gap-2">
                    <span>Open Zonal Dashboard</span>
                    <svg class="w-4 h-4 text-orange-400" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M14 5l7 7m0 0l-7 7m7-7H3"/></svg>
                </button>
            </form>
        </div>
    </div>

    <!-- ============================================== -->
    <!-- MODAL: ADMIN PANEL LOGIN & DASHBOARD           -->
    <!-- ============================================== -->
    <div id="admin-modal" class="hidden fixed inset-0 bg-slate-900/50 backdrop-blur-sm z-50 flex items-center justify-center p-4">
        <div class="bg-white rounded-3xl max-w-4xl w-full max-h-[90vh] flex flex-col shadow-2xl overflow-hidden border border-slate-200">
            
            <div class="bg-white p-4 sm:p-5 border-b border-slate-200 flex justify-between items-center">
                <div class="flex items-center space-x-2.5">
                    <div class="w-8 h-8 rounded-lg bg-orange-50 border border-orange-200 flex items-center justify-center text-orange-600 font-bold text-base">
                        🛡️
                    </div>
                    <div>
                        <h3 class="text-sm sm:text-base font-black text-slate-900">Admin Panel - National Award Summary</h3>
                        <p class="text-[11px] text-slate-500">Exium MUPS Central Award Management</p>
                    </div>
                </div>
                <button onclick="closeAdminModal()" class="text-slate-400 hover:text-slate-600 p-1">
                    <svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"/></svg>
                </button>
            </div>

            <div class="p-5 sm:p-6 overflow-y-auto space-y-5 flex-1 text-xs sm:text-sm">
                
                <!-- Password Prompt (If locked) -->
                <div id="admin-auth-box" class="space-y-3 max-w-sm mx-auto py-8 text-center">
                    <div class="w-10 h-10 bg-orange-50 text-orange-600 rounded-xl mx-auto flex items-center justify-center text-xl mb-1 border border-orange-100">
                        🔒
                    </div>
                    <h4 class="font-black text-sm text-slate-900">Admin Authentication</h4>
                    <p class="text-xs text-slate-500">Enter master password to access national summary</p>
                    <div class="flex gap-2">
                        <input type="password" id="admin-password-input" onkeydown="if(event.key==='Enter') verifyAdminPassword()" placeholder="Enter Admin Password..." class="flex-1 bg-slate-50 border border-slate-300 rounded-xl px-3.5 py-2 text-xs font-semibold text-slate-800 outline-none focus:ring-2 focus:ring-orange-500">
                        <button onclick="verifyAdminPassword()" class="px-4 py-2 bg-slate-900 hover:bg-slate-800 text-white rounded-xl text-xs font-bold transition">
                            Unlock
                        </button>
                    </div>
                </div>

                <!-- Admin Dashboard (When unlocked) -->
                <div id="admin-dashboard-box" class="hidden space-y-5">
                    
                    <div class="grid grid-cols-3 gap-3">
                        <div class="bg-slate-50 border border-slate-200 rounded-xl p-3.5 text-center">
                            <div class="text-[11px] font-bold text-slate-500 uppercase">National Achievers</div>
                            <div id="admin-stat-total" class="text-2xl font-black text-slate-900 mt-0.5">1,541</div>
                            <div class="text-[10px] text-slate-500">May (721) + June (820)</div>
                        </div>
                        <div class="bg-emerald-50 border border-emerald-200 rounded-xl p-3.5 text-center">
                            <div class="text-[11px] font-bold text-emerald-700 uppercase">Completed</div>
                            <div id="admin-stat-completed" class="text-2xl font-black text-emerald-800 mt-0.5">0</div>
                            <div id="admin-stat-pct" class="text-[10px] text-emerald-600 font-bold">0.0%</div>
                        </div>
                        <div class="bg-amber-50 border border-amber-200 rounded-xl p-3.5 text-center">
                            <div class="text-[11px] font-bold text-amber-700 uppercase">Pending</div>
                            <div id="admin-stat-pending" class="text-2xl font-black text-amber-800 mt-0.5">1,541</div>
                            <div class="text-[10px] text-amber-600">Awaiting choice</div>
                        </div>
                    </div>

                    <!-- Voucher Brand Breakdown Table -->
                    <div class="bg-white border border-slate-200 rounded-xl overflow-hidden shadow-sm">
                        <div class="p-3 bg-slate-50 border-b border-slate-200 font-black text-xs uppercase tracking-wider text-slate-800 flex justify-between items-center">
                            <span>National Voucher Brand Requirements (PO Planning)</span>
                            <span class="text-[11px] text-slate-500 font-normal">Real-time</span>
                        </div>
                        <table class="w-full text-left text-xs">
                            <thead class="bg-slate-100 text-slate-700 font-bold border-b border-slate-200">
                                <tr>
                                    <th class="py-2 px-3">Brand Logo & Name</th>
                                    <th class="py-2 px-3 text-center">May</th>
                                    <th class="py-2 px-3 text-center">June</th>
                                    <th class="py-2 px-3 text-center">Total</th>
                                    <th class="py-2 px-3 text-right">Share (%)</th>
                                </tr>
                            </thead>
                            <tbody id="admin-voucher-tbody" class="divide-y divide-slate-100 font-medium"></tbody>
                        </table>
                    </div>

                    <!-- Admin Controls: Regional Head Input Control (On/Off) & Data Reset -->
                    <div class="grid grid-cols-1 md:grid-cols-2 gap-3.5">
                        
                        <!-- 1. Regional Head Input On/Off Toggle -->
                        <div class="bg-slate-50 border border-slate-200 rounded-2xl p-4 space-y-2.5 flex flex-col justify-between">
                            <div>
                                <div class="flex items-center justify-between gap-2 mb-1.5">
                                    <div class="flex items-center gap-1.5">
                                        <span class="text-base">🔒</span>
                                        <h4 class="font-black text-xs sm:text-sm text-slate-900">Regional Head Input Control</h4>
                                    </div>
                                    <span id="admin-lock-badge" class="px-2.5 py-0.5 rounded-full text-[10px] font-black bg-emerald-100 text-emerald-800 border border-emerald-300">🟢 Input ON (Active)</span>
                                </div>
                                <p class="text-[11px] text-slate-500 leading-snug">
                                    Enable or freeze voucher choices for Regional Heads. When <strong>OFF (Locked)</strong>, the portal switches to Read-Only mode.
                                </p>
                            </div>
                            <div class="pt-1">
                                <button id="admin-toggle-lock-btn" onclick="toggleSubmissionLock()" class="w-full py-2 px-3 rounded-xl text-xs font-black transition flex items-center justify-center gap-2 shadow-xs bg-rose-600 hover:bg-rose-700 text-white">
                                    <span>🔒 Turn OFF Input (Lock Submissions)</span>
                                </button>
                            </div>
                        </div>

                        <!-- 2. Data Deletion & Reset Center -->
                        <div class="bg-rose-50/60 border border-rose-200 rounded-2xl p-4 space-y-2.5 flex flex-col justify-between">
                            <div>
                                <div class="flex items-center justify-between gap-2 mb-1.5">
                                    <div class="flex items-center gap-1.5">
                                        <span class="text-base">🗑️</span>
                                        <h4 class="font-black text-xs sm:text-sm text-rose-950">Data Deletion & Reset</h4>
                                    </div>
                                    <span class="px-2 py-0.5 rounded-md text-[10px] font-bold bg-rose-100 text-rose-800 border border-rose-200">Admin Only</span>
                                </div>
                                <p class="text-[11px] text-rose-900/80 leading-snug">
                                    Clear submitted choices and reset MIO records back to <strong>Pending</strong> (with double-confirmation protection).
                                </p>
                            </div>
                            <div class="grid grid-cols-2 gap-2 pt-1">
                                <button onclick="adminDeleteCurrentRegionData()" title="Clear choices only for active region" class="py-2 px-2.5 bg-white hover:bg-rose-50 text-rose-700 font-bold rounded-xl border border-rose-300 text-xs transition flex items-center justify-center gap-1 shadow-2xs">
                                    <span>🗑️</span>
                                    <span>Reset Region</span>
                                </button>
                                <button onclick="adminDeleteAllNationalData()" title="Clear choices for all MIOs nationwide" class="py-2 px-2.5 bg-rose-600 hover:bg-rose-700 text-white font-black rounded-xl text-xs transition flex items-center justify-center gap-1 shadow-xs">
                                    <span>⚠️</span>
                                    <span>Delete All Data</span>
                                </button>
                            </div>
                        </div>

                    </div>

                    <!-- Google Apps Script URL Endpoint -->
                    <div class="bg-slate-50 border border-slate-200 rounded-xl p-3.5 space-y-2.5">
                        <div class="flex items-center justify-between">
                            <label class="block text-xs font-bold text-slate-700">Google Apps Script Web App Endpoint URL</label>
                            <span id="admin-endpoint-status" class="text-[10px] font-bold text-emerald-700 bg-emerald-100 px-2 py-0.5 rounded border border-emerald-300 hidden">✓ Connected</span>
                        </div>
                        <div class="flex flex-wrap gap-2">
                            <input type="text" id="admin-endpoint-input" placeholder="Paste deployed Google Apps Script URL (https://script.google.com/macros/s/.../exec)" class="flex-1 min-w-[260px] bg-white border border-slate-300 rounded-lg px-3 py-1.5 text-xs font-mono text-slate-800 outline-none focus:ring-2 focus:ring-orange-500">
                            <button onclick="saveAdminEndpoint()" class="px-3 py-1.5 bg-slate-900 hover:bg-slate-800 text-white rounded-lg text-xs font-bold transition">Save</button>
                            <button onclick="syncFromCloud(false)" class="px-3 py-1.5 bg-blue-600 hover:bg-blue-700 text-white rounded-lg text-xs font-bold transition flex items-center gap-1 shadow-xs">
                                <svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15"/></svg>
                                <span>Sync Now</span>
                            </button>
                            <button onclick="fixGoogleSheetFormatting()" title="Fix May Ach% and BD Timestamps directly in your Google Sheet" class="px-3.5 py-1.5 bg-amber-600 hover:bg-amber-700 text-white rounded-lg text-xs font-bold transition flex items-center gap-1 shadow-xs">
                                <span>🛠️ Fix Google Sheet Formatting</span>
                            </button>
                        </div>

                        <!-- Universal Multi-Device Cloud Sync Link Banner -->
                        <div class="p-3 bg-emerald-50/90 border border-emerald-200 rounded-xl flex flex-wrap items-center justify-between gap-2">
                            <div class="space-y-0.5">
                                <span class="text-xs font-bold text-emerald-950 flex items-center gap-1.5">
                                    <span>🌐</span>
                                    <span>Universal Multi-Device Link (Auto-Connected)</span>
                                </span>
                                <p class="text-[11px] text-emerald-800">
                                    Share this standard link with all Regional Heads: <strong class="font-mono text-emerald-950">https://didarul2911-netizen.github.io/exium-award-campaign-2026/</strong>
                                </p>
                            </div>
                            <button onclick="copyShareableLink()" class="px-2.5 py-1.5 bg-emerald-700 hover:bg-emerald-800 text-white rounded-lg text-xs font-bold transition flex items-center gap-1 shrink-0 shadow-xs">
                                <span>📋 Copy Link</span>
                            </button>
                        </div>

                        <p class="text-[11px] text-slate-500 leading-normal">
                            Connects directly to your central Google Sheet for live 2-way real-time data synchronization across all devices. Use <strong>Fix Google Sheet Formatting</strong> to automatically correct May Ach% & BD timestamps in your sheet with one click.
                        </p>
                    </div>

                </div>

            </div>

            <div class="bg-slate-50 p-3.5 border-t border-slate-200 flex flex-wrap justify-between items-center gap-2">
                <div class="flex items-center gap-2">
                    <button onclick="exportToExcel('admin')" class="px-3.5 py-2 bg-emerald-700 hover:bg-emerald-800 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 transition shadow-sm">
                        <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/></svg>
                        <span>Export National Excel (.xlsx)</span>
                    </button>
                    <button onclick="downloadCsvExport()" class="px-3 py-2 bg-white border border-slate-300 hover:bg-slate-100 text-slate-700 rounded-xl text-xs font-bold flex items-center gap-1 transition shadow-xs">
                        <svg class="w-3.5 h-3.5 text-slate-500" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/></svg>
                        <span>CSV</span>
                    </button>
                </div>
                <button onclick="closeAdminModal()" class="px-4 py-2 bg-slate-900 hover:bg-slate-800 text-white rounded-xl text-xs font-bold transition">
                    Close
                </button>
            </div>

        </div>
    </div>

    <!-- Toast Notification -->
    <div id="toast-box" class="fixed top-16 right-4 bg-slate-900 text-white px-4 py-2.5 rounded-2xl shadow-xl flex items-center space-x-2 text-xs font-bold transform translate-x-32 opacity-0 pointer-events-none transition-all duration-300 z-50">
        <span id="toast-icon">✅</span>
        <span id="toast-text">Notification</span>
    </div>

    <!-- CLIENT LOGIC & EMBEDDED DATABASE -->
    <script>
        const DB = {json_db_str};

        const BRANDS = [
            {{
                name: 'Infinity Gift Voucher',
                subtitle: 'Mega Mall & Lifestyle',
                key: 'infinity',
                logo: DB.logos.infinity
            }},
            {{
                name: 'Apex Gift Voucher',
                subtitle: 'Premium Footwear & Leather',
                key: 'apex',
                logo: DB.logos.apex
            }},
            {{
                name: 'Bata Gift Voucher',
                subtitle: 'Footwear & Accessories',
                key: 'bata',
                logo: DB.logos.bata
            }},
            {{
                name: 'Aarong Gift Voucher',
                subtitle: 'Heritage & Lifestyle Fashion',
                key: 'aarong',
                logo: DB.logos.aarong
            }},
            {{
                name: 'Best Buy Gift Voucher',
                subtitle: 'RFL Home & Electronics',
                key: 'bestbuy',
                logo: DB.logos.bestbuy
            }},
            {{
                name: 'Cats Eye Gift Voucher',
                subtitle: 'Fashion & Lifestyle',
                key: 'catseye',
                logo: DB.logos.catseye
            }}
        ];

        // 12 Distinct Vibrant Light Color Themes for MIO Profile & Territory Headers (1 step darker: 100-series with crisp 300-series borders)
        const MIO_COLOR_PALETTES = [
            {{ bg: 'bg-sky-100', border: 'border-sky-300', avatar: 'bg-sky-600', desigBg: 'bg-white text-sky-950 border-sky-300' }},
            {{ bg: 'bg-emerald-100', border: 'border-emerald-300', avatar: 'bg-emerald-600', desigBg: 'bg-white text-emerald-950 border-emerald-300' }},
            {{ bg: 'bg-amber-100', border: 'border-amber-300', avatar: 'bg-amber-600', desigBg: 'bg-white text-amber-950 border-amber-300' }},
            {{ bg: 'bg-purple-100', border: 'border-purple-300', avatar: 'bg-purple-600', desigBg: 'bg-white text-purple-950 border-purple-300' }},
            {{ bg: 'bg-rose-100', border: 'border-rose-300', avatar: 'bg-rose-600', desigBg: 'bg-white text-rose-950 border-rose-300' }},
            {{ bg: 'bg-teal-100', border: 'border-teal-300', avatar: 'bg-teal-600', desigBg: 'bg-white text-teal-950 border-teal-300' }},
            {{ bg: 'bg-indigo-100', border: 'border-indigo-300', avatar: 'bg-indigo-600', desigBg: 'bg-white text-indigo-950 border-indigo-300' }},
            {{ bg: 'bg-cyan-100', border: 'border-cyan-300', avatar: 'bg-cyan-600', desigBg: 'bg-white text-cyan-950 border-cyan-300' }},
            {{ bg: 'bg-orange-100', border: 'border-orange-300', avatar: 'bg-orange-600', desigBg: 'bg-white text-orange-950 border-orange-300' }},
            {{ bg: 'bg-violet-100', border: 'border-violet-300', avatar: 'bg-violet-600', desigBg: 'bg-white text-violet-950 border-violet-300' }},
            {{ bg: 'bg-lime-100', border: 'border-lime-300', avatar: 'bg-lime-600', desigBg: 'bg-white text-lime-950 border-lime-300' }},
            {{ bg: 'bg-fuchsia-100', border: 'border-fuchsia-300', avatar: 'bg-fuchsia-600', desigBg: 'bg-white text-fuchsia-950 border-fuchsia-300' }}
        ];

        let currentRegionName = '';
        let currentActiveZoneCode = '';
        let currentMonthKey = 'May_2026';
        let savedChoices = JSON.parse(localStorage.getItem('EXIUM_AWARD_CHOICES') || '{{}}');

        // Multi-Device Auto-Connect Engine:
        // 1. Production Central Google Apps Script Web App Endpoint
        // 2. Optional override via URL query parameters: ?api=... or ?endpoint=...
        // 3. Fallback to localStorage
        const DEFAULT_CLOUD_ENDPOINT = 'https://script.google.com/macros/s/AKfycbypTb7igli6MB6JzH5U9htVWSKe8Vgd-njUahmuJXtT1b83CUaDLHzyFpnMHsAm58JRHA/exec';
        let queryEndpoint = '';
        try {{
            const urlParams = new URLSearchParams(window.location.search);
            queryEndpoint = (urlParams.get('api') || urlParams.get('endpoint') || '').trim();
            if (queryEndpoint) {{
                localStorage.setItem('EXIUM_AWARD_ENDPOINT', queryEndpoint);
            }}
        }} catch (e) {{}}

        let cloudEndpoint = queryEndpoint || localStorage.getItem('EXIUM_AWARD_ENDPOINT') || DEFAULT_CLOUD_ENDPOINT;
        if (!cloudEndpoint || !cloudEndpoint.startsWith('http')) {{
            cloudEndpoint = DEFAULT_CLOUD_ENDPOINT;
            localStorage.setItem('EXIUM_AWARD_ENDPOINT', DEFAULT_CLOUD_ENDPOINT);
        }}
        let isAdminUnlocked = false;
        let isSubmissionsLocked = (localStorage.getItem('EXIUM_SUBMISSIONS_LOCKED') === 'true');

        // BANGLADESH STANDARD TIME (BDT / UTC+6) GENERATOR & FORMATTER
        function getBDTimestamp() {{
            try {{
                const d = new Date();
                const parts = new Intl.DateTimeFormat('en-GB', {{
                    timeZone: 'Asia/Dhaka',
                    day: '2-digit',
                    month: 'short',
                    year: 'numeric',
                    hour: '2-digit',
                    minute: '2-digit',
                    second: '2-digit',
                    hour12: true
                }}).formatToParts(d);

                let day = '', month = '', year = '', hour = '', min = '', sec = '', dayPeriod = '';
                for (const p of parts) {{
                    if (p.type === 'day') day = p.value;
                    else if (p.type === 'month') month = p.value;
                    else if (p.type === 'year') year = p.value;
                    else if (p.type === 'hour') hour = p.value;
                    else if (p.type === 'minute') min = p.value;
                    else if (p.type === 'second') sec = p.value;
                    else if (p.type === 'dayPeriod') dayPeriod = p.value.replace(/\\./g, '').toUpperCase();
                }}
                if (day && month && year && hour && min && sec && dayPeriod) {{
                    return `${{day}}-${{month}}-${{year}} ${{hour}}:${{min}}:${{sec}} ${{dayPeriod}}`;
                }}
            }} catch (e) {{}}

            // Deterministic UTC+6 fallback: Shift UTC epoch by exactly +6 hours
            const d = new Date();
            const bdDate = new Date(d.getTime() + (6 * 60 * 60 * 1000));
            const pad = n => String(n).padStart(2, '0');
            const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
            let h = bdDate.getUTCHours();
            const ampm = h >= 12 ? 'PM' : 'AM';
            h = h % 12 || 12;
            return `${{pad(bdDate.getUTCDate())}}-${{months[bdDate.getUTCMonth()]}}-${{bdDate.getUTCFullYear()}} ${{pad(h)}}:${{pad(bdDate.getUTCMinutes())}}:${{pad(bdDate.getUTCSeconds())}} ${{ampm}}`;
        }}

        function formatToBDTime(isoOrStr) {{
            if (!isoOrStr) return '';
            const trimmed = String(isoOrStr).trim();
            if (/^\\d{{2}}-[A-Za-z]{{3}}-\\d{{4}}\\s+\\d{{1,2}}:\\d{{2}}:\\d{{2}}\\s+(AM|PM)/i.test(trimmed)) {{
                return trimmed;
            }}
            try {{
                const d = new Date(trimmed);
                if (isNaN(d.getTime())) return trimmed;
                try {{
                    const parts = new Intl.DateTimeFormat('en-GB', {{
                        timeZone: 'Asia/Dhaka',
                        day: '2-digit',
                        month: 'short',
                        year: 'numeric',
                        hour: '2-digit',
                        minute: '2-digit',
                        second: '2-digit',
                        hour12: true
                    }}).formatToParts(d);
                    let day = '', month = '', year = '', hour = '', min = '', sec = '', dayPeriod = '';
                    for (const p of parts) {{
                        if (p.type === 'day') day = p.value;
                        else if (p.type === 'month') month = p.value;
                        else if (p.type === 'year') year = p.value;
                        else if (p.type === 'hour') hour = p.value;
                        else if (p.type === 'minute') min = p.value;
                        else if (p.type === 'second') sec = p.value;
                        else if (p.type === 'dayPeriod') dayPeriod = p.value.replace(/\\./g, '').toUpperCase();
                    }}
                    if (day && month && year && hour && min && sec && dayPeriod) {{
                        return `${{day}}-${{month}}-${{year}} ${{hour}}:${{min}}:${{sec}} ${{dayPeriod}}`;
                    }}
                }} catch (e) {{}}

                const bdDate = new Date(d.getTime() + (6 * 60 * 60 * 1000));
                const pad = n => String(n).padStart(2, '0');
                const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
                let h = bdDate.getUTCHours();
                const ampm = h >= 12 ? 'PM' : 'AM';
                h = h % 12 || 12;
                return `${{pad(bdDate.getUTCDate())}}-${{months[bdDate.getUTCMonth()]}}-${{bdDate.getUTCFullYear()}} ${{pad(h)}}:${{pad(bdDate.getUTCMinutes())}}:${{pad(bdDate.getUTCSeconds())}} ${{ampm}}`;
            }} catch (e) {{
                return trimmed;
            }}
        }}

        document.addEventListener('DOMContentLoaded', () => {{
            if (DB.logos.main_logo) {{
                document.getElementById('header-logo').src = DB.logos.main_logo;
            }}
            initLoginDropdowns();

            // Clean visible URL query string if ?api= was supplied, keeping address bar clean
            if (queryEndpoint && window.history && window.history.replaceState) {{
                try {{
                    const cleanUrl = window.location.pathname;
                    window.history.replaceState({{}}, document.title, cleanUrl);
                }} catch (e) {{}}
            }}

            if (cloudEndpoint) {{
                const input = document.getElementById('admin-endpoint-input');
                if (input) input.value = cloudEndpoint;
                const statusBadge = document.getElementById('admin-endpoint-status');
                if (statusBadge) statusBadge.classList.remove('hidden');
                updateShareableLinkBox();
                // Flush any offline pending changes and pull fresh cloud data
                flushPendingQueue();
                syncFromCloud(true);
            }} else {{
                updateSyncStatusBadge('offline');
            }}

            // Real-Time Background Sync: Poll cloud every 10 seconds for instant updates from other devices
            setInterval(() => {{
                if (cloudEndpoint && !isSyncing) {{
                    syncFromCloud(true);
                }}
            }}, 10000);

            // Instant Sync when device becomes active or regains internet
            window.addEventListener('focus', () => {{
                if (cloudEndpoint && !isSyncing) syncFromCloud(true);
            }});
            document.addEventListener('visibilitychange', () => {{
                if (document.visibilityState === 'visible' && cloudEndpoint && !isSyncing) {{
                    syncFromCloud(true);
                }}
            }});
            window.addEventListener('online', () => {{
                flushPendingQueue();
                if (cloudEndpoint && !isSyncing) syncFromCloud(true);
            }});

            updateSubmissionLockUI();

            const session = JSON.parse(localStorage.getItem('EXIUM_AWARD_SESSION') || 'null');
            if (session && session.region && DB.regions_map[session.region]) {{
                openRegionalView(session.region);
            }}
        }});

        function initLoginDropdowns() {{
            const sortedZones = [...DB.zones].sort((a, b) => a.zone_name.localeCompare(b.zone_name));

            const regZoneSel = document.getElementById('reg-zone-select');
            regZoneSel.innerHTML = '<option value="">-- Choose Zone --</option>';
            sortedZones.forEach(z => {{
                const opt = document.createElement('option');
                opt.value = z.sap_zone_code;
                opt.textContent = `${{z.zone_name}} (${{z.sap_zone_code}})`;
                regZoneSel.appendChild(opt);
            }});

            const zoneLoginSel = document.getElementById('zone-login-select');
            zoneLoginSel.innerHTML = '<option value="">-- Choose Zone --</option>';
            sortedZones.forEach(z => {{
                const opt = document.createElement('option');
                opt.value = z.sap_zone_code;
                opt.textContent = `${{z.zone_name}} - ${{z.zonal_head}} (${{z.sap_zone_code}})`;
                zoneLoginSel.appendChild(opt);
            }});
        }}

        function switchLoginTab(mode) {{
            const tabReg = document.getElementById('tab-reg');
            const tabMio = document.getElementById('tab-mio');
            const formReg = document.getElementById('form-regional');
            const formMio = document.getElementById('form-mio-search');

            if (tabReg && tabMio) {{
                [tabReg, tabMio].forEach(t => {{
                    t.className = 'flex-1 py-2.5 text-center border-b-2 border-transparent hover:text-slate-800 text-slate-500 font-bold';
                }});
            }}
            if (formReg) formReg.classList.add('hidden');
            if (formMio) formMio.classList.add('hidden');

            if (mode === 'regional') {{
                if (tabReg) tabReg.className = 'flex-1 py-2.5 text-center border-b-2 border-orange-500 text-orange-600 font-black';
                if (formReg) formReg.classList.remove('hidden');
            }} else if (mode === 'mio') {{
                if (tabMio) tabMio.className = 'flex-1 py-2.5 text-center border-b-2 border-orange-500 text-orange-600 font-black';
                if (formMio) formMio.classList.remove('hidden');
            }}
        }}

        function onRegionalZoneChange() {{
            const zoneCode = document.getElementById('reg-zone-select').value;
            const regSel = document.getElementById('reg-region-select');
            regSel.innerHTML = '<option value="">-- Select Region --</option>';
            document.getElementById('reg-details-box').classList.add('hidden');

            if (!zoneCode) {{
                regSel.disabled = true;
                regSel.classList.add('bg-slate-100');
                return;
            }}

            const zoneObj = DB.zones.find(z => z.sap_zone_code === zoneCode);
            if (!zoneObj) return;

            regSel.disabled = false;
            regSel.classList.remove('bg-slate-100');

            zoneObj.regions.forEach(r => {{
                const opt = document.createElement('option');
                opt.value = r.region_name;
                const totalAch = (r.may_count || 0) + (r.jun_count || 0);
                opt.textContent = `${{r.region_name}} (${{r.sap_region_code}}) - ${{totalAch}} Achievers`;
                regSel.appendChild(opt);
            }});
        }}

        function onRegionalRegionChange() {{
            const regName = document.getElementById('reg-region-select').value;
            const box = document.getElementById('reg-details-box');
            if (!regName || !DB.regions_map[regName]) {{
                box.classList.add('hidden');
                return;
            }}

            const r = DB.regions_map[regName];
            document.getElementById('disp-reg-head').textContent = r.regional_head || 'Not Assigned';
            const total = (r.may_count || 0) + (r.jun_count || 0);
            document.getElementById('disp-reg-count').textContent = `${{total}} Achievers (May: ${{r.may_count}}, June: ${{r.jun_count}})`;
            box.classList.remove('hidden');
        }}

        function handleRegionalLogin(e) {{
            e.preventDefault();
            const regName = document.getElementById('reg-region-select').value;
            const pass = document.getElementById('reg-password-input').value.trim();

            if (!regName) {{
                alert('Please select a Region!');
                return;
            }}

            const r = DB.regions_map[regName];
            const expectedSap = String(r.sap_region_code).trim();

            if (pass !== expectedSap) {{
                alert('Invalid Region Password! Regional Head login strictly requires the SAP Region Code (' + expectedSap + ').');
                return;
            }}

            openRegionalView(regName);
        }}

        function openRegionalView(regName) {{
            currentRegionName = regName;
            const r = DB.regions_map[regName];

            localStorage.setItem('EXIUM_AWARD_SESSION', JSON.stringify({{ region: regName }}));

            document.getElementById('login-view').classList.add('hidden');
            document.getElementById('zonal-view').classList.add('hidden');
            document.getElementById('regional-view').classList.remove('hidden');
            document.getElementById('sticky-bar').classList.remove('hidden');

            document.getElementById('regview-zone').textContent = r.zone_name;
            document.getElementById('regview-sap').textContent = `SAP: ${{r.sap_region_code}}`;
            document.getElementById('regview-name').textContent = r.region_name;
            document.getElementById('regview-rh').innerHTML = `Regional Head: <strong class="text-slate-900">${{r.regional_head}}</strong>`;
            document.getElementById('regview-zh').innerHTML = `Zonal Head: <strong class="text-orange-700 font-bold">${{r.zonal_head}}</strong>`;
            document.getElementById('sticky-region-text').textContent = r.region_name;

            // Clear search filter input if any
            const filterInput = document.getElementById('reg-mio-filter-input');
            if (filterInput) filterInput.value = '';

            renderAchievers();
            updateRegionProgress();
            window.scrollTo({{ top: 0, behavior: 'smooth' }});
        }}

        function exitToLogin() {{
            localStorage.removeItem('EXIUM_AWARD_SESSION');
            document.getElementById('regional-view').classList.add('hidden');
            document.getElementById('zonal-view').classList.add('hidden');
            document.getElementById('sticky-bar').classList.add('hidden');
            document.getElementById('login-view').classList.remove('hidden');
            currentRegionName = '';
            currentActiveZoneCode = '';
        }}

        // GROUP ACHIEVERS IN REGION BY MIO PERSON
        function getRegionMios(regName) {{
            const miosMap = new Map();

            const mayList = DB.achievers_may.filter(a => a.region === regName);
            const junList = DB.achievers_jun.filter(a => a.region === regName);

            mayList.forEach(a => {{
                const key = (a.mio_code && a.mio_code.trim()) ? a.mio_code.trim() : (a.area_code + '_' + a.mio_name);
                if (!miosMap.has(key)) {{
                    miosMap.set(key, {{
                        key: key,
                        mio_name: a.mio_name,
                        mio_code: a.mio_code,
                        area_code: a.area_code,
                        area_name: a.area_name,
                        desig: a.desig,
                        is_transferred: a.is_transferred,
                        transfer_details: a.transfer_details,
                        may_award: a,
                        jun_award: null
                    }});
                }} else {{
                    const item = miosMap.get(key);
                    item.may_award = a;
                    if (a.is_transferred && !item.is_transferred) {{
                        item.is_transferred = true;
                        item.transfer_details = a.transfer_details;
                    }}
                }}
            }});

            junList.forEach(a => {{
                const key = (a.mio_code && a.mio_code.trim()) ? a.mio_code.trim() : (a.area_code + '_' + a.mio_name);
                if (!miosMap.has(key)) {{
                    miosMap.set(key, {{
                        key: key,
                        mio_name: a.mio_name,
                        mio_code: a.mio_code,
                        area_code: a.area_code,
                        area_name: a.area_name,
                        desig: a.desig,
                        is_transferred: a.is_transferred,
                        transfer_details: a.transfer_details,
                        may_award: null,
                        jun_award: a
                    }});
                }} else {{
                    const item = miosMap.get(key);
                    item.jun_award = a;
                    if (a.is_transferred && !item.is_transferred) {{
                        item.is_transferred = true;
                        item.transfer_details = a.transfer_details;
                    }}
                }}
            }});

            return Array.from(miosMap.values()).sort((a, b) => {{
                const cmpArea = (a.area_name || '').localeCompare(b.area_name || '');
                if (cmpArea !== 0) return cmpArea;
                const cmpCode = (a.area_code || '').localeCompare(b.area_code || '');
                if (cmpCode !== 0) return cmpCode;
                return (a.mio_name || '').localeCompare(b.mio_name || '');
            }});
        }}

        // CALCULATE MIO STATUS: COMPLETE ONLY WHEN ALL ELIGIBLE MONTHS ARE COMPLETED
        function getMioStatus(mio) {{
            let eligibleMonths = 0;
            let completedMonths = 0;

            if (mio.may_award) {{
                eligibleMonths++;
                if (savedChoices[mio.may_award.id] && savedChoices[mio.may_award.id].voucher) {{
                    completedMonths++;
                }}
            }}
            if (mio.jun_award) {{
                eligibleMonths++;
                if (savedChoices[mio.jun_award.id] && savedChoices[mio.jun_award.id].voucher) {{
                    completedMonths++;
                }}
            }}

            const isComplete = (eligibleMonths > 0 && completedMonths === eligibleMonths);
            return {{
                eligibleMonths,
                completedMonths,
                isComplete
            }};
        }}

        // RENDER A SINGLE MONTH AWARD BOX WITH VOUCHER SELECTION & REMOVE OPTION
        function renderMonthAwardBox(ach, monthKey) {{
            const isMay = (monthKey === 'May_2026');
            const chosenVoucher = (savedChoices[ach.id] && savedChoices[ach.id].voucher) ? savedChoices[ach.id].voucher : '';
            const isCompleted = Boolean(chosenVoucher);

            const themeBorder = isMay ? 'border-blue-200 bg-blue-50/20' : 'border-purple-200 bg-purple-50/20';
            const badgeMonth = isMay ? 'bg-blue-600 text-white' : 'bg-purple-700 text-white';
            const badgeValue = isMay ? 'bg-blue-100 text-blue-900 border-blue-200' : 'bg-purple-100 text-purple-900 border-purple-200';
            const monthLabel = isMay ? 'May 2026 Award' : 'June 2026 Award';

            let selectionStatusHtml = '';
            const choiceObj = savedChoices[ach.id];
            const tsFormatted = (choiceObj && choiceObj.timestamp) ? formatToBDTime(choiceObj.timestamp) : '';

            if (isCompleted) {{
                selectionStatusHtml = `
                    <div class="flex items-center gap-2 flex-wrap sm:flex-nowrap">
                        <span class="text-xs font-black text-emerald-800 bg-emerald-100/90 border border-emerald-300 px-3 py-1 rounded-xl flex items-center gap-1.5 shadow-sm">
                            <span>✓</span>
                            <span>Selected: <strong>${{chosenVoucher}}</strong></span>
                            ${{tsFormatted ? `<span class="text-[10px] font-medium text-emerald-700 font-mono ml-1">(${{tsFormatted}})</span>` : ''}}
                        </span>
                        ${{isSubmissionsLocked ? `
                            <span class="text-[11px] font-bold text-slate-400 bg-slate-100 border border-slate-200 px-2.5 py-1 rounded-xl cursor-not-allowed flex items-center gap-1 shadow-2xs" title="Voucher selections are locked by administrator">
                                <span>🔒</span>
                                <span>Locked</span>
                            </span>
                        ` : `
                            <button onclick="removeVoucher('${{ach.id}}')" title="Remove / Clear this voucher choice" class="text-[11px] font-bold text-rose-700 hover:text-white bg-rose-50 hover:bg-rose-600 border border-rose-200 hover:border-rose-600 px-2.5 py-1 rounded-xl transition flex items-center gap-1 shadow-sm">
                                <svg class="w-3 h-3" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16"/></svg>
                                <span>✕ Remove Choice</span>
                            </button>
                        `}}
                    </div>
                `;
            }} else {{
                selectionStatusHtml = `
                    <div class="flex items-center gap-2 flex-wrap">
                        <span class="text-xs font-bold text-amber-700 bg-amber-50 border border-amber-200 px-2.5 py-1 rounded-xl flex items-center gap-1">
                            <span>⏳</span>
                            <span>Selection Pending</span>
                        </span>
                        ${{tsFormatted ? `<span class="text-[10px] font-semibold text-slate-500 bg-slate-100 border border-slate-200 px-2 py-0.5 rounded-lg font-mono" title="Timestamp preserved from earlier selection">Recorded: ${{tsFormatted}}</span>` : ''}}
                    </div>
                `;
            }}

            const brandsGrid = BRANDS.map(b => {{
                const isSelected = (chosenVoucher === b.name);
                const clickTooltip = isSubmissionsLocked ? '🔒 Submissions are currently locked by Administrator' : (isSelected ? 'Click to Deselect / Remove Choice' : 'Click to Select ' + b.name);
                return `
                    <div onclick="selectVoucher('${{ach.id}}', '${{b.name}}')" 
                         title="${{clickTooltip}}"
                         class="brand-card ${{isSubmissionsLocked ? 'cursor-not-allowed opacity-90' : 'cursor-pointer'}} rounded-xl border-2 p-2 sm:p-2.5 flex flex-col justify-between relative transition-all ${{
                             isSelected ? 'selected ring-2 ring-emerald-500 ring-offset-1 bg-white' : 'border-slate-200 hover:border-slate-300 bg-white'
                         }}">
                        <div class="flex items-center justify-between mb-1">
                            <div class="h-6 sm:h-7 w-16 sm:w-20 flex items-center">
                                <img src="${{b.logo}}" alt="${{b.name}}" class="max-h-6 sm:max-h-7 max-w-full object-contain">
                            </div>
                            <div class="w-4 h-4 rounded-full border-2 flex items-center justify-center flex-shrink-0 ${{
                                isSelected ? 'border-emerald-600 bg-emerald-600 text-white shadow-xs' : 'border-slate-300 bg-slate-50'
                            }}">
                                ${{isSelected ? '<svg class="w-2.5 h-2.5" fill="currentColor" viewBox="0 0 20 20"><path fill-rule="evenodd" d="M16.707 5.293a1 1 0 010 1.414l-8 8a1 1 0 01-1.414 0l-4-4a1 1 0 011.414-1.414L8 12.586l7.293-7.293a1 1 0 011.414 0z" clip-rule="evenodd"></path></svg>' : ''}}
                            </div>
                        </div>
                        <div>
                            <div class="text-[11px] sm:text-xs font-black text-slate-900 tracking-tight leading-tight truncate">${{b.name.replace(' Gift Voucher', '')}}</div>
                            <div class="text-[9px] text-slate-400 font-medium truncate mt-0.5">${{b.subtitle}}</div>
                        </div>
                        ${{isSelected ? `
                            <div class="mt-1 pt-1 border-t border-emerald-100 flex items-center justify-between text-[9px] text-emerald-800 font-black">
                                <span>Selected</span>
                                ${{isSubmissionsLocked ? '<span class="text-slate-400">🔒 Locked</span>' : '<span class="text-rose-600 hover:underline">✕ Deselect</span>'}}
                            </div>
                        ` : ''}}
                    </div>
                `;
            }}).join('');

            return `
                <div class="rounded-2xl border-2 ${{themeBorder}} p-3.5 sm:p-4 space-y-3">
                    <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-2.5 border-b border-slate-200/60 pb-2.5">
                        <div class="flex items-center gap-2 flex-wrap">
                            <span class="px-2.5 py-1 rounded-lg text-xs font-black ${{badgeMonth}} shadow-sm">
                                ${{monthLabel}}
                            </span>
                            <span class="px-2.5 py-1 rounded-lg text-xs font-black border ${{badgeValue}}">
                                ${{Number(ach.award).toLocaleString()}}
                            </span>
                            <span class="text-xs text-slate-600 font-medium ml-1">
                                ${{ach.metric_label}}: <strong class="text-slate-900 font-bold">${{ach.metric_val}}</strong> ${{ach.rx ? `(Rx: ${{ach.rx}})` : ''}}
                            </span>
                        </div>
                        ${{selectionStatusHtml}}
                    </div>

                    <!-- 6 Brand Options Grid -->
                    <div class="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-2 sm:gap-2.5">
                        ${{brandsGrid}}
                    </div>
                </div>
            `;
        }}

        // RENDER ACHIEVERS WITH TERRITORY + MIO NAME TOGETHER AND DUAL-MONTH AWARDS
        function renderAchievers() {{
            const container = document.getElementById('regview-achievers-container');
            container.innerHTML = '';

            let mios = getRegionMios(currentRegionName);

            // Filter Query
            const filterInput = document.getElementById('reg-mio-filter-input');
            const q = (filterInput ? filterInput.value : '').trim().toLowerCase();
            if (q) {{
                mios = mios.filter(m => 
                    (m.mio_name && m.mio_name.toLowerCase().includes(q)) ||
                    (m.mio_code && m.mio_code.toLowerCase().includes(q)) ||
                    (m.area_name && m.area_name.toLowerCase().includes(q)) ||
                    (m.area_code && m.area_code.toLowerCase().includes(q))
                );
            }}

            if (mios.length === 0) {{
                container.innerHTML = `
                    <div class="bg-white rounded-2xl p-12 text-center border border-slate-200 text-slate-400">
                        <span class="text-3xl">📭</span>
                        <h4 class="font-bold text-sm text-slate-700 mt-2">${{q ? 'No matching Territory or MIO found' : 'No Eligible Achievers in this Region'}}</h4>
                        <p class="text-xs text-slate-500 mt-1">${{q ? `No achiever matches "${{q}}" in ${{currentRegionName}}` : `There are no achievers listed under ${{currentRegionName}}.`}}</p>
                    </div>
                `;
                return;
            }}

            mios.forEach((mio, mioIdx) => {{
                const status = getMioStatus(mio);
                const totalAwardVal = (mio.may_award ? Number(mio.may_award.award) : 0) + (mio.jun_award ? Number(mio.jun_award.award) : 0);
                const theme = MIO_COLOR_PALETTES[mioIdx % MIO_COLOR_PALETTES.length];

                const card = document.createElement('div');
                card.id = `mio-card-${{mio.key}}`;
                card.className = `bg-white rounded-2xl sm:rounded-3xl shadow-sm border transition-all p-4 sm:p-6 space-y-4 ${{
                    status.isComplete 
                        ? 'border-emerald-400 ring-2 ring-emerald-400/20 bg-emerald-50/5' 
                        : (status.completedMonths > 0 ? 'border-amber-300 ring-1 ring-amber-300/30' : 'border-slate-200')
                }}`;

                let transferBadgeHtml = '';
                if (mio.is_transferred && mio.transfer_details) {{
                    const td = mio.transfer_details;
                    transferBadgeHtml = `
                        <div class="bg-amber-50 border border-amber-200 rounded-xl p-3 text-xs text-amber-950 flex items-start gap-2">
                            <span class="text-sm flex-shrink-0">🔄</span>
                            <div class="space-y-0.5">
                                <div class="font-black text-amber-900">Transferred MIO Information</div>
                                <div class="text-[11px] text-amber-800">
                                    Award achieved while in: <strong>${{td.prev_area_name}} (${{td.prev_area_code}})</strong> • Region: <strong>${{td.prev_region}}</strong> (RH: <strong>${{td.prev_rh}}</strong>) • Zone: <strong>${{td.prev_zone}}</strong>
                                </div>
                            </div>
                        </div>
                    `;
                }}

                // Status Badge
                let statusBadgeHtml = '';
                if (status.isComplete) {{
                    statusBadgeHtml = `
                        <span class="px-3.5 py-1.5 rounded-full text-xs font-black bg-emerald-100 text-emerald-800 border border-emerald-300 flex items-center gap-1.5 shadow-sm">
                            <span>✓</span>
                            <span>Complete (${{status.completedMonths}}/${{status.eligibleMonths}})</span>
                        </span>
                    `;
                }} else if (status.completedMonths > 0) {{
                    statusBadgeHtml = `
                        <span class="px-3.5 py-1.5 rounded-full text-xs font-black bg-amber-100 text-amber-800 border border-amber-300 flex items-center gap-1.5 shadow-sm">
                            <span>⏳</span>
                            <span>${{status.completedMonths}}/${{status.eligibleMonths}} Complete</span>
                        </span>
                    `;
                }} else {{
                    statusBadgeHtml = `
                        <span class="px-3.5 py-1.5 rounded-full text-xs font-black bg-white/95 text-slate-700 border border-slate-300/80 flex items-center gap-1.5 shadow-2xs">
                            <span>⚪</span>
                            <span>Pending (0/${{status.eligibleMonths}})</span>
                        </span>
                    `;
                }}

                // Month eligibility labels
                let eligibilityMonthsList = [];
                if (mio.may_award) eligibilityMonthsList.push('May');
                if (mio.jun_award) eligibilityMonthsList.push('June');
                const monthEligibilityText = eligibilityMonthsList.join(' & ');

                // Card Header: Territory and MIO Name Together with Individual Light Color Theme
                let headerHtml = `
                    <div class="rounded-2xl border-2 p-3.5 sm:p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3 shadow-xs ${{theme.bg}} ${{theme.border}}">
                        <div class="space-y-2">
                            <div class="flex items-center gap-2 flex-wrap">
                                <span class="bg-orange-500 text-white text-xs font-black px-2.5 py-1 rounded-lg flex items-center gap-1 shadow-xs">
                                    <span>📍</span>
                                    <span>${{mio.area_name}} (${{mio.area_code}})</span>
                                </span>
                                <span class="text-slate-400 font-light">•</span>
                                <h3 class="text-base sm:text-lg font-black text-slate-900 tracking-tight">
                                    ${{mio.mio_name}}
                                </h3>
                                <span class="bg-white text-blue-700 font-mono text-xs font-black px-2 py-0.5 rounded-md border border-blue-200 shadow-2xs">
                                    MIO: ${{mio.mio_code}}
                                </span>
                            </div>
                            <div class="flex items-center gap-2 text-xs text-slate-700 font-medium flex-wrap">
                                <span class="font-bold ${{theme.desigBg}} px-2.5 py-0.5 rounded text-[11px] border shadow-2xs">${{mio.desig || 'Medical Information Officer'}}</span>
                                <span class="text-slate-400">•</span>
                                <span>Eligible: <strong class="text-slate-900 font-bold">${{monthEligibilityText}}</strong> (${{status.eligibleMonths}} Month${{status.eligibleMonths > 1 ? 's' : ''}})</span>
                            </div>
                        </div>

                        <div class="flex items-center gap-2.5 flex-shrink-0 self-start sm:self-center flex-wrap">
                            <!-- Prominent Highlighted Total Award Value Badge -->
                            <div class="px-3.5 py-1.5 rounded-xl bg-white border-2 border-orange-400 text-orange-950 shadow-sm flex items-center gap-2" title="Total Bi-Monthly Award Value">
                                <span class="text-[10px] sm:text-[11px] font-extrabold text-slate-500 uppercase tracking-wider">Total Award</span>
                                <span class="text-base sm:text-lg font-black text-orange-600 font-mono tracking-tight">${{totalAwardVal.toLocaleString()}}</span>
                            </div>
                            ${{statusBadgeHtml}}
                        </div>
                    </div>
                `;

                // Month Award Boxes
                let awardBoxesHtml = '<div class="space-y-3.5">';
                if (mio.may_award) {{
                    awardBoxesHtml += renderMonthAwardBox(mio.may_award, 'May_2026');
                }}
                if (mio.jun_award) {{
                    awardBoxesHtml += renderMonthAwardBox(mio.jun_award, 'June_2026');
                }}
                awardBoxesHtml += '</div>';

                card.innerHTML = headerHtml + transferBadgeHtml + awardBoxesHtml;
                container.appendChild(card);
            }});
        }}

        // ==============================================
        // PERSISTENT OUTBOX QUEUE & MULTI-DEVICE SYNC
        // ==============================================
        function getPendingQueue() {{
            try {{
                return JSON.parse(localStorage.getItem('EXIUM_PENDING_QUEUE') || '{{}}');
            }} catch (e) {{
                return {{}};
            }}
        }}

        function savePendingQueue(q) {{
            try {{
                localStorage.setItem('EXIUM_PENDING_QUEUE', JSON.stringify(q));
            }} catch (e) {{}}
        }}

        function queueForUpload(achId, month, area_code, mio_code, voucher, timestamp, action = 'save') {{
            const q = getPendingQueue();
            q[achId] = {{
                id: achId,
                month: month,
                area_code: area_code,
                mio_code: mio_code,
                voucher: voucher,
                timestamp: timestamp,
                action: action,
                clientTime: Date.now()
            }};
            savePendingQueue(q);
        }}

        let isFlushingQueue = false;
        async function flushPendingQueue() {{
            if (!cloudEndpoint) return;
            if (isFlushingQueue) return;

            const q = getPendingQueue();
            const batch = Object.values(q);
            if (batch.length === 0) return;

            isFlushingQueue = true;
            updateSyncStatusBadge('syncing');

            try {{
                const bodyStr = JSON.stringify({{
                    action: 'save_choices',
                    choices: batch
                }});

                let success = false;
                try {{
                    const res = await fetch(cloudEndpoint, {{
                        method: 'POST',
                        headers: {{ 'Content-Type': 'text/plain;charset=utf-8' }},
                        body: bodyStr
                    }});
                    if (res.ok) success = true;
                }} catch (corsErr) {{
                    try {{
                        await fetch(cloudEndpoint, {{
                            method: 'POST',
                            mode: 'no-cors',
                            headers: {{ 'Content-Type': 'text/plain;charset=utf-8' }},
                            body: bodyStr
                        }});
                        success = true;
                    }} catch (err) {{
                        console.warn('Network send error:', err);
                    }}
                }}

                if (success) {{
                    const currentQ = getPendingQueue();
                    batch.forEach(item => {{
                        if (currentQ[item.id] && currentQ[item.id].clientTime === item.clientTime) {{
                            delete currentQ[item.id];
                        }}
                    }});
                    savePendingQueue(currentQ);
                    updateSyncStatusBadge('synced');
                }} else {{
                    updateSyncStatusBadge('pending');
                }}
            }} catch (e) {{
                console.warn('Queue flush error:', e);
                updateSyncStatusBadge('pending');
            }} finally {{
                isFlushingQueue = false;
            }}
        }}

        let debounceFlushTimer = null;
        function triggerDebouncedFlush() {{
            updateSyncStatusBadge('pending');
            clearTimeout(debounceFlushTimer);
            debounceFlushTimer = setTimeout(() => {{
                flushPendingQueue();
            }}, 350);
        }}

        function updateSyncStatusBadge(state, detail = '') {{
            const elDesk = document.getElementById('header-sync-status');
            const dotDesk = document.getElementById('header-sync-dot');
            const iconDesk = document.getElementById('sync-icon-desktop');
            const elMob = document.getElementById('header-sync-status-mobile');
            const dotMob = document.getElementById('header-sync-dot-mobile');

            if (!cloudEndpoint) {{
                if (elDesk) elDesk.textContent = 'Offline (Click to Connect)';
                if (dotDesk) dotDesk.className = 'w-2 h-2 rounded-full bg-amber-400';
                if (elMob) elMob.textContent = 'Offline';
                if (dotMob) dotMob.className = 'w-1.5 h-1.5 rounded-full bg-amber-400';
                return;
            }}

            if (state === 'syncing') {{
                if (elDesk) elDesk.textContent = 'Syncing...';
                if (dotDesk) dotDesk.className = 'w-2 h-2 rounded-full bg-blue-500 animate-ping';
                if (iconDesk) iconDesk.classList.add('animate-spin');
                if (elMob) elMob.textContent = 'Syncing';
                if (dotMob) dotMob.className = 'w-1.5 h-1.5 rounded-full bg-blue-500 animate-ping';
            }} else if (state === 'synced') {{
                const timeStr = detail || new Date().toLocaleTimeString('en-US', {{ timeZone: 'Asia/Dhaka', hour: '2-digit', minute: '2-digit', second: '2-digit' }});
                if (elDesk) elDesk.textContent = `Live: ${{timeStr}}`;
                if (dotDesk) dotDesk.className = 'w-2 h-2 rounded-full bg-emerald-500 animate-pulse';
                if (iconDesk) iconDesk.classList.remove('animate-spin');
                if (elMob) elMob.textContent = 'Live';
                if (dotMob) dotMob.className = 'w-1.5 h-1.5 rounded-full bg-emerald-500 animate-pulse';
            }} else if (state === 'pending') {{
                const q = getPendingQueue();
                const count = Object.keys(q).length;
                if (elDesk) elDesk.textContent = `Pending (${{count}})`;
                if (dotDesk) dotDesk.className = 'w-2 h-2 rounded-full bg-amber-500';
                if (iconDesk) iconDesk.classList.remove('animate-spin');
                if (elMob) elMob.textContent = `Queued`;
                if (dotMob) dotMob.className = 'w-1.5 h-1.5 rounded-full bg-amber-500';
            }}
        }}

        function onHeaderCloudSyncClick() {{
            if (!cloudEndpoint) {{
                const input = prompt('Enter Google Apps Script Web App URL to connect this device to live Google Sheets:');
                if (input && input.trim()) {{
                    cloudEndpoint = input.trim();
                    localStorage.setItem('EXIUM_AWARD_ENDPOINT', cloudEndpoint);
                    showToast('Cloud Endpoint Connected!', '☁️');
                    updateShareableLinkBox();
                    flushPendingQueue();
                    syncFromCloud(false);
                }}
                return;
            }}
            flushPendingQueue();
            syncFromCloud(false);
        }}

        function updateShareableLinkBox() {{
            // No-op: Central Google Apps Script URL is embedded directly as DEFAULT_CLOUD_ENDPOINT
        }}

        function copyShareableLink() {{
            const cleanUrl = 'https://didarul2911-netizen.github.io/exium-award-campaign-2026/';
            navigator.clipboard.writeText(cleanUrl).then(() => {{
                showToast('Standard Portal Link copied! Send directly to Regional Heads.', '📋');
            }}).catch(() => {{
                showToast(cleanUrl, '📋');
            }});
        }}

        function selectVoucher(achId, voucherName) {{
            if (isSubmissionsLocked) {{
                alert('🔒 Voucher choice submissions are currently CLOSED by Administrator. Portal is in Read-Only mode.');
                return;
            }}

            // Toggle Deselect: If this voucher is already chosen, clicking again removes it
            if (savedChoices[achId] && savedChoices[achId].voucher === voucherName) {{
                removeVoucher(achId);
                return;
            }}

            const all = DB.achievers_may.concat(DB.achievers_jun);
            const a = all.find(x => x.id === achId);
            if (!a) return;

            const bdTime = getBDTimestamp();
            savedChoices[achId] = {{
                voucher: voucherName,
                timestamp: bdTime,
                status: 'Complete'
            }};

            localStorage.setItem('EXIUM_AWARD_CHOICES', JSON.stringify(savedChoices));

            // Queue for cloud upload in persistent storage
            queueForUpload(achId, a.month, a.area_code, a.mio_code, voucherName, bdTime, 'save');

            renderAchievers();
            updateRegionProgress();
            triggerDebouncedFlush();
            showToast(`Selected ${{voucherName}}`, '✅');
        }}

        function removeVoucher(achId) {{
            if (isSubmissionsLocked) {{
                alert('🔒 Voucher choice submissions are currently CLOSED by Administrator. Portal is in Read-Only mode.');
                return;
            }}

            if (!savedChoices[achId]) return;

            const all = DB.achievers_may.concat(DB.achievers_jun);
            const a = all.find(x => x.id === achId);
            if (!a) return;

            const oldVoucher = savedChoices[achId].voucher;
            const existingTimestamp = (savedChoices[achId] && savedChoices[achId].timestamp) ? savedChoices[achId].timestamp : '';

            // User requirement: When deselected, only choice is removed/blank, but timestamp is preserved
            savedChoices[achId] = {{
                voucher: '',
                timestamp: existingTimestamp,
                status: 'Pending'
            }};

            localStorage.setItem('EXIUM_AWARD_CHOICES', JSON.stringify(savedChoices));

            // Queue removal in persistent storage
            queueForUpload(achId, a.month, a.area_code, a.mio_code, '', existingTimestamp, 'remove');

            renderAchievers();
            updateRegionProgress();
            triggerDebouncedFlush();
            showToast(`Removed ${{oldVoucher}}`, '🗑️');
        }}

        function updateRegionProgress() {{
            const regMay = DB.achievers_may.filter(a => a.region === currentRegionName);
            const regJun = DB.achievers_jun.filter(a => a.region === currentRegionName);
            const allAwards = regMay.concat(regJun);

            let completedAwards = 0;
            allAwards.forEach(a => {{
                if (savedChoices[a.id] && savedChoices[a.id].voucher) completedAwards++;
            }});

            const totalAwards = allAwards.length;
            const pct = totalAwards > 0 ? Math.round((completedAwards / totalAwards) * 100) : 0;

            const mios = getRegionMios(currentRegionName);
            let completeMios = 0;
            mios.forEach(m => {{
                const st = getMioStatus(m);
                if (st.isComplete) completeMios++;
            }});

            document.getElementById('regview-progress-pct').textContent = `${{pct}}%`;
            document.getElementById('regview-progress-bar').style.width = `${{pct}}%`;
            document.getElementById('regview-progress-count').textContent = `${{completedAwards}} of ${{totalAwards}} Awards (${{completeMios}}/${{mios.length}} MIOs Complete)`;
            document.getElementById('sticky-completed-text').textContent = `${{completedAwards}}/${{totalAwards}} Awards (${{pct}}%) • ${{completeMios}}/${{mios.length}} MIOs Complete`;
        }}

        async function manualSaveSync() {{
            if (isSubmissionsLocked) {{
                alert('🔒 Voucher choice submissions are currently CLOSED by Administrator. Selections cannot be modified or saved.');
                return;
            }}

            const btn = document.getElementById('btn-save-sync');
            btn.innerHTML = `<svg class="w-3.5 h-3.5 animate-spin text-white" fill="none" stroke="currentColor" viewBox="0 0 24 24"><circle class="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="4"></circle><path class="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path></svg> <span>Saving & Syncing...</span>`;

            // 1. Save local choices to localStorage
            localStorage.setItem('EXIUM_AWARD_CHOICES', JSON.stringify(savedChoices));

            // 2. Queue all choices in active region to guarantee they are queued
            if (currentRegionName && DB.regions_map[currentRegionName]) {{
                const mios = DB.regions_map[currentRegionName].mios;
                mios.forEach(m => {{
                    if (m.may_award && savedChoices[m.may_award.id]) {{
                        const ch = savedChoices[m.may_award.id];
                        queueForUpload(m.may_award.id, 'May_2026', m.may_award.area_code, m.may_award.mio_code, ch.voucher || '', ch.timestamp || getBDTimestamp(), ch.voucher ? 'save' : 'remove');
                    }}
                    if (m.jun_award && savedChoices[m.jun_award.id]) {{
                        const ch = savedChoices[m.jun_award.id];
                        queueForUpload(m.jun_award.id, 'June_2026', m.jun_award.area_code, m.jun_award.mio_code, ch.voucher || '', ch.timestamp || getBDTimestamp(), ch.voucher ? 'save' : 'remove');
                    }}
                }});
            }}

            // 3. Flush outbox to cloud
            await flushPendingQueue();

            // 4. Pull fresh data from cloud
            await syncFromCloud(true);

            btn.innerHTML = `<span>✓ Saved & Synced</span>`;
            showToast('All selections saved and synced with Google Sheets!', '💾');
            setTimeout(() => {{
                btn.innerHTML = `<svg class="w-3.5 h-3.5 text-emerald-400" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M8 7H5a2 2 0 00-2 2v9a2 2 0 002 2h14a2 2 0 002-2V9a2 2 0 00-2-2h-3m-1 4l-3 3m0 0l-3-3m3 3V4"/></svg><span>Save</span>`;
            }}, 2000);
        }}

        function onZonalZoneChange() {{
            const zoneCode = document.getElementById('zone-login-select').value;
            const box = document.getElementById('zone-details-box');
            if (!zoneCode) {{
                box.classList.add('hidden');
                return;
            }}
            const z = DB.zones.find(x => x.sap_zone_code === zoneCode);
            if (z) {{
                document.getElementById('disp-zh-name').textContent = z.zonal_head;
                document.getElementById('disp-zh-regcount').textContent = `${{z.regions.length}} Regions`;
                box.classList.remove('hidden');
            }}
        }}

        function handleZonalLogin(e) {{
            e.preventDefault();
            const zoneCode = document.getElementById('zone-login-select').value;
            const pass = document.getElementById('zonal-password-input').value.trim();

            if (!zoneCode) {{
                alert('Please select your Zone!');
                return;
            }}

            if (pass !== zoneCode) {{
                alert('Invalid Zonal Password! Zonal login strictly requires your SAP Zone Code (' + zoneCode + ').');
                return;
            }}

            closeZonalModal();
            openZonalDashboard(zoneCode);
        }}

        function openZonalDashboard(zoneCode) {{
            currentActiveZoneCode = zoneCode;
            const zoneObj = DB.zones.find(z => z.sap_zone_code === zoneCode);
            if (!zoneObj) return;

            document.getElementById('login-view').classList.add('hidden');
            document.getElementById('regional-view').classList.add('hidden');
            document.getElementById('sticky-bar').classList.add('hidden');
            document.getElementById('zonal-view').classList.remove('hidden');

            document.getElementById('zoneview-name').textContent = zoneObj.zone_name;
            document.getElementById('zoneview-sap').textContent = `SAP Zone: ${{zoneObj.sap_zone_code}}`;
            document.getElementById('zoneview-zh').textContent = `Zonal Head: ${{zoneObj.zonal_head}}`;

            const allAchievers = DB.achievers_may.concat(DB.achievers_jun);
            const zoneAchievers = allAchievers.filter(a => a.sap_zone_code === zoneCode);
            const mayCount = DB.achievers_may.filter(a => a.sap_zone_code === zoneCode).length;
            const junCount = DB.achievers_jun.filter(a => a.sap_zone_code === zoneCode).length;

            let completed = 0;
            const brandCounts = {{}};
            BRANDS.forEach(b => brandCounts[b.name] = 0);

            zoneAchievers.forEach(a => {{
                const ch = savedChoices[a.id];
                if (ch && ch.voucher) {{
                    completed++;
                    if (brandCounts[ch.voucher] !== undefined) {{
                        brandCounts[ch.voucher]++;
                    }}
                }}
            }});

            const total = zoneAchievers.length;
            const pct = total > 0 ? ((completed / total) * 100).toFixed(1) : '0.0';

            document.getElementById('zone-stat-total').textContent = total;
            document.getElementById('zone-stat-mayjun').textContent = `May: ${{mayCount}} | June: ${{junCount}}`;
            document.getElementById('zone-stat-completed').textContent = completed;
            document.getElementById('zone-stat-pct').textContent = `${{pct}}%`;
            document.getElementById('zone-stat-pending').textContent = total - completed;
            document.getElementById('zone-stat-regions').textContent = zoneObj.regions.length;

            const tbody = document.getElementById('zone-regions-tbody');
            tbody.innerHTML = '';

            zoneObj.regions.forEach(r => {{
                const regMay = DB.achievers_may.filter(a => a.region === r.region_name);
                const regJun = DB.achievers_jun.filter(a => a.region === r.region_name);
                const regTotal = regMay.length + regJun.length;
                let regDone = 0;
                regMay.concat(regJun).forEach(a => {{
                    if (savedChoices[a.id] && savedChoices[a.id].voucher) regDone++;
                }});
                const regPct = regTotal > 0 ? Math.round((regDone / regTotal) * 100) : 0;

                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td class="py-2.5 px-3 font-black text-slate-900">${{r.region_name}}</td>
                    <td class="py-2.5 px-3 font-mono text-slate-600">${{r.sap_region_code}}</td>
                    <td class="py-2.5 px-3 text-slate-700">${{r.regional_head}}</td>
                    <td class="py-2.5 px-3 text-center">${{regMay.length}}</td>
                    <td class="py-2.5 px-3 text-center">${{regJun.length}}</td>
                    <td class="py-2.5 px-3 text-center font-bold text-slate-900">${{regTotal}}</td>
                    <td class="py-2.5 px-3 text-center">
                        <span class="px-2 py-0.5 rounded-full text-[10px] font-bold ${{
                            regDone === regTotal && regTotal > 0 ? 'bg-emerald-100 text-emerald-800' : 'bg-slate-100 text-slate-700'
                        }}">${{regDone}}/${{regTotal}} (${{regPct}}%)</span>
                    </td>
                    <td class="py-2.5 px-3 text-right">
                        <button onclick="openRegionalView('${{r.region_name}}')" class="px-2.5 py-1 bg-slate-100 hover:bg-slate-200 text-slate-800 font-bold rounded-lg border border-slate-300 text-[11px] transition">
                            View Cards →
                        </button>
                    </td>
                `;
                tbody.appendChild(tr);
            }});

            const vGrid = document.getElementById('zone-voucher-cards-grid');
            vGrid.innerHTML = '';
            BRANDS.forEach(b => {{
                const count = brandCounts[b.name] || 0;
                const vCard = document.createElement('div');
                vCard.className = 'bg-slate-50 border border-slate-200 rounded-xl p-2.5 text-center space-y-1';
                vCard.innerHTML = `
                    <div class="h-6 flex items-center justify-center">
                        <img src="${{b.logo}}" alt="${{b.name}}" class="max-h-6 max-w-full object-contain">
                    </div>
                    <div class="text-[10px] font-bold text-slate-700 truncate">${{b.name.replace(' Gift Voucher', '')}}</div>
                    <div class="text-base font-black text-slate-900">${{count}}</div>
                `;
                vGrid.appendChild(vCard);
            }});

            window.scrollTo({{ top: 0, behavior: 'smooth' }});
        }}

        function searchMioDirect(e) {{
            if (e && e.preventDefault) e.preventDefault();
            const q = document.getElementById('mio-quick-query').value.trim().toLowerCase();
            const resBox = document.getElementById('mio-search-results-box');
            resBox.innerHTML = '';
            if (!q) return;

            const all = DB.achievers_may.concat(DB.achievers_jun);

            // Group achievers by person (using mio_code or unique fallback key)
            const personsMap = {{}};
            all.forEach(a => {{
                const key = a.mio_code || (a.area_code + '_' + a.mio_name);
                if (!personsMap[key]) {{
                    personsMap[key] = {{
                        mio_code: a.mio_code,
                        mio_name: a.mio_name,
                        desig: a.desig,
                        area_name: a.area_name,
                        area_code: a.area_code,
                        region: a.region,
                        sap_region_code: a.sap_region_code,
                        regional_head: a.regional_head,
                        zone: a.zone,
                        sap_zone_code: a.sap_zone_code,
                        zonal_head: a.zonal_head,
                        is_transferred: false,
                        transfer_records: [],
                        awards: []
                    }};
                }}
                if (a.is_transferred && a.transfer_details) {{
                    personsMap[key].is_transferred = true;
                    personsMap[key].transfer_records.push({{
                        month: a.month_label,
                        details: a.transfer_details
                    }});
                }}
                personsMap[key].awards.push(a);
            }});

            const matchedPersons = Object.values(personsMap).filter(p => 
                (p.mio_code && p.mio_code.toLowerCase().includes(q)) || 
                (p.area_code && p.area_code.toLowerCase().includes(q)) || 
                (p.mio_name && p.mio_name.toLowerCase().includes(q))
            );

            if (matchedPersons.length === 0) {{
                resBox.innerHTML = `<div class="p-4 bg-slate-50 border border-slate-200 rounded-2xl text-slate-500 text-xs text-center">No matching achiever found for "${{q}}".</div>`;
                return;
            }}

            matchedPersons.slice(0, 10).forEach((p, pIdx) => {{
                // Sort awards chronologically (May first, then June)
                p.awards.sort((a, b) => (a.month === 'May_2026' ? -1 : 1));
                const theme = MIO_COLOR_PALETTES[pIdx % MIO_COLOR_PALETTES.length];

                const card = document.createElement('div');
                card.className = 'p-4 sm:p-5 bg-white border border-slate-200 rounded-2xl shadow-sm text-xs space-y-3.5';

                // Separate Awards HTML for each month
                let awardsHtml = '';
                p.awards.forEach(m => {{
                    const choiceInfo = savedChoices[m.id];
                    const isCompleted = !!choiceInfo && choiceInfo.voucher;
                    const isMay = (m.month === 'May_2026');
                    const badgeBg = isMay ? 'bg-blue-100 text-blue-800 border-blue-200' : 'bg-purple-100 text-purple-800 border-purple-200';

                    awardsHtml += `
                        <div class="bg-slate-50/70 border border-slate-200 rounded-xl p-3 space-y-2">
                            <div class="flex items-center justify-between flex-wrap gap-1.5">
                                <div class="flex items-center gap-1.5">
                                    <span class="px-2 py-0.5 rounded-md font-black text-[10px] uppercase border ${{badgeBg}}">
                                        ${{m.month_label}} Award
                                    </span>
                                    <span class="px-2 py-0.5 rounded-md font-extrabold text-[10px] bg-emerald-50 text-emerald-800 border border-emerald-200">
                                        Value: ${{Number(m.award).toLocaleString()}}
                                    </span>
                                </div>
                                <span class="text-[10px] font-bold ${{isCompleted ? 'text-emerald-700 bg-emerald-50 border border-emerald-200' : 'text-amber-700 bg-amber-50 border border-amber-200'}} px-2 py-0.5 rounded-md">
                                    ${{isCompleted ? '✓ Completed' : '⏳ Pending'}}
                                </span>
                            </div>

                            <div class="text-[11px] text-slate-600 bg-white p-2 rounded-lg border border-slate-100 flex items-center justify-between flex-wrap gap-1">
                                <span>Criteria: <strong>${{m.metric_label}}</strong> = <strong class="text-slate-900">${{m.metric_val}}</strong></span>
                                ${{m.rx ? `<span class="text-slate-500">Rx: <strong class="text-slate-800">${{m.rx}}</strong></span>` : ''}}
                            </div>

                            <div>
                                ${{isCompleted ? `
                                    <div class="bg-emerald-50 border border-emerald-200 rounded-lg p-2 flex items-center justify-between gap-2">
                                        <div class="flex items-center gap-1.5">
                                            <span class="w-5 h-5 rounded-full bg-emerald-600 text-white flex items-center justify-center font-black text-[10px]">✓</span>
                                            <span class="text-[10px] font-bold text-slate-500 uppercase tracking-wider">Selected Voucher:</span>
                                            <strong class="text-xs font-black text-emerald-950">${{choiceInfo.voucher}}</strong>
                                        </div>
                                        ${{choiceInfo && choiceInfo.timestamp ? `<span class="text-[10px] font-medium text-emerald-700 font-mono">${{formatToBDTime(choiceInfo.timestamp)}}</span>` : ''}}
                                    </div>
                                ` : `
                                    <div class="bg-amber-50/60 border border-amber-200 rounded-lg p-2 flex items-center justify-between gap-2">
                                        <div class="flex items-center gap-1.5">
                                            <span class="w-5 h-5 rounded-full bg-amber-100 text-amber-800 flex items-center justify-center font-black text-[10px]">⏳</span>
                                            <span class="text-[11px] font-bold text-amber-900">Pending Selection by Regional Head</span>
                                        </div>
                                        ${{choiceInfo && choiceInfo.timestamp ? `<span class="text-[10px] font-semibold text-slate-500 bg-slate-100 border border-slate-200 px-2 py-0.5 rounded font-mono" title="Timestamp preserved from earlier selection">Recorded: ${{formatToBDTime(choiceInfo.timestamp)}}</span>` : ''}}
                                    </div>
                                `}}
                            </div>
                        </div>
                    `;
                }});

                // Transfer callout if any
                let transferHtml = '';
                if (p.is_transferred && p.transfer_records.length > 0) {{
                    const tList = p.transfer_records.map(tr => 
                        `<strong>${{tr.month}}</strong> in previous territory <em>${{tr.details.prev_area_name}} (${{tr.details.prev_area_code}})</em>, Region <em>${{tr.details.prev_region}}</em> (${{tr.details.prev_rh || 'RH'}})`
                    ).join('; ');
                    transferHtml = `
                        <div class="bg-amber-50 border border-amber-200 rounded-xl p-2.5 text-[11px] text-amber-900 leading-snug">
                            <span class="font-bold">🔄 Transfer Record:</span> Award earned in ${{tList}}.
                        </div>
                    `;
                }}

                card.innerHTML = `
                    <!-- 1. Person Header (Shown Once) with Light Theme -->
                    <div class="p-3 rounded-xl border flex items-start justify-between gap-2 shadow-2xs ${{theme.bg}} ${{theme.border}}">
                        <div>
                            <div class="flex items-center gap-1.5 flex-wrap mb-1">
                                <span class="bg-orange-500 text-white text-[11px] font-black px-2 py-0.5 rounded-md shadow-2xs">
                                    📍 ${{p.area_name}} (${{p.area_code}})
                                </span>
                                <span class="px-2 py-0.5 rounded-md font-black text-[10px] bg-slate-900 text-white shadow-2xs">
                                    SAP MIO: ${{p.mio_code}}
                                </span>
                                <span class="px-2 py-0.5 rounded-md font-bold text-[10px] bg-white/90 text-orange-800 border border-orange-200 shadow-2xs">
                                    ${{p.awards.length}} ${{p.awards.length > 1 ? 'Months Qualified' : 'Month Qualified'}}
                                </span>
                            </div>
                            <h4 class="text-base font-black text-slate-900 tracking-tight">${{p.mio_name}}</h4>
                            <div class="text-xs text-slate-600 font-semibold mt-0.5">${{p.desig}}</div>
                        </div>
                    </div>

                    <!-- 2. Posting / Hierarchy Details (Shown Once) -->
                    <div class="grid grid-cols-2 gap-2 bg-slate-50 p-2.5 rounded-xl border border-slate-200/80 text-[11px]">
                        <div>
                            <span class="text-slate-400 block text-[9px] font-bold uppercase tracking-wider">Current Territory</span>
                            <strong class="text-slate-800 text-xs">${{p.area_name}}</strong>
                            <span class="text-slate-500 text-[10px] block font-medium">Area Code: ${{p.area_code}}</span>
                        </div>
                        <div>
                            <span class="text-slate-400 block text-[9px] font-bold uppercase tracking-wider">Region & Zone</span>
                            <strong class="text-slate-800 text-xs">${{p.region}}</strong>
                            <span class="text-slate-500 text-[10px] block font-medium">Zone: ${{p.zone}}</span>
                        </div>
                        <div class="col-span-2 border-t border-slate-200 pt-1.5 flex flex-col sm:flex-row sm:justify-between gap-1 text-[10px]">
                            <span>Regional Head: <strong class="text-slate-800 font-bold">${{p.regional_head || 'N/A'}}</strong></span>
                            <span>Zonal Head: <strong class="text-slate-800 font-bold">${{p.zonal_head || 'N/A'}}</strong></span>
                        </div>
                    </div>

                    ${{transferHtml}}

                    <!-- 3. Separate Month Awards Section -->
                    <div class="space-y-2 pt-1">
                        <div class="text-[10px] font-black uppercase tracking-wider text-slate-400">
                            Award & Voucher Details
                        </div>
                        <div class="space-y-2">
                            ${{awardsHtml}}
                        </div>
                    </div>

                    <div class="text-[10px] text-slate-400 text-center italic border-t border-slate-100 pt-2">
                        🔒 Read-only view. Voucher selection is submitted by Regional Head.
                    </div>
                `;

                resBox.appendChild(card);
            }});
        }}

        let currentCriteriaTab = 'may';
        function openAwardCriteriaModal(initialTab = 'may') {{
            currentCriteriaTab = initialTab;
            const modal = document.getElementById('criteria-modal');
            const imgMay = document.getElementById('img-crit-may');
            const imgJun = document.getElementById('img-crit-jun');

            if (DB.criteria_images) {{
                if (imgMay && DB.criteria_images.may) imgMay.src = DB.criteria_images.may;
                if (imgJun && DB.criteria_images.june) imgJun.src = DB.criteria_images.june;
            }}

            switchCriteriaTab(initialTab);
            modal.classList.remove('hidden');
        }}

        function closeAwardCriteriaModal() {{
            document.getElementById('criteria-modal').classList.add('hidden');
        }}

        function switchCriteriaTab(tab) {{
            currentCriteriaTab = tab;
            const tabMay = document.getElementById('tab-crit-may');
            const tabJun = document.getElementById('tab-crit-jun');
            const viewMay = document.getElementById('crit-view-may');
            const viewJun = document.getElementById('crit-view-jun');

            if (tab === 'may') {{
                tabMay.className = 'px-4 py-2.5 rounded-t-xl text-xs sm:text-sm font-black transition border-b-2 border-blue-600 text-blue-700 bg-white shadow-xs';
                tabJun.className = 'px-4 py-2.5 rounded-t-xl text-xs sm:text-sm font-bold transition border-b-2 border-transparent text-slate-500 hover:text-slate-800';
                viewMay.classList.remove('hidden');
                viewJun.classList.add('hidden');
            }} else {{
                tabJun.className = 'px-4 py-2.5 rounded-t-xl text-xs sm:text-sm font-black transition border-b-2 border-purple-700 text-purple-700 bg-white shadow-xs';
                tabMay.className = 'px-4 py-2.5 rounded-t-xl text-xs sm:text-sm font-bold transition border-b-2 border-transparent text-slate-500 hover:text-slate-800';
                viewJun.classList.remove('hidden');
                viewMay.classList.add('hidden');
            }}
        }}

        function openImageFullscreen(src) {{
            if (!src) return;
            const win = window.open();
            win.document.write(`<html><head><title>Award Criteria Document</title><style>body{{margin:0;background:#0f172a;display:flex;justify-content:center;align-items:center;min-height:100vh;}}img{{max-width:100%;height:auto;box-shadow:0 10px 25px rgba(0,0,0,0.5);}}</style></head><body><img src="${{src}}"></body></html>`);
        }}

        function downloadActiveCriteria() {{
            const isMay = (currentCriteriaTab === 'may');
            const src = isMay ? DB.criteria_images?.may : DB.criteria_images?.june;
            if (!src) return;

            const link = document.createElement('a');
            link.href = src;
            link.download = isMay ? 'Exium_Award_Criteria_May_2026.png' : 'Exium_Award_Criteria_June_2026.png';
            link.click();
        }}

        function openZonalModal() {{
            document.getElementById('zonal-modal').classList.remove('hidden');
        }}

        function closeZonalModal() {{
            document.getElementById('zonal-modal').classList.add('hidden');
        }}

        function openAdminModal() {{
            document.getElementById('admin-modal').classList.remove('hidden');
            if (isAdminUnlocked) {{
                showAdminDashboard();
            }} else {{
                document.getElementById('admin-auth-box').classList.remove('hidden');
                document.getElementById('admin-dashboard-box').classList.add('hidden');
            }}
        }}

        function closeAdminModal() {{
            document.getElementById('admin-modal').classList.add('hidden');
        }}

        function verifyAdminPassword() {{
            const pass = document.getElementById('admin-password-input').value.trim();
            if (pass === 'Exium MUPS') {{
                isAdminUnlocked = true;
                showAdminDashboard();
            }} else {{
                alert('Invalid Admin Password! Please enter the correct password.');
            }}
        }}

        function showAdminDashboard() {{
            document.getElementById('admin-auth-box').classList.add('hidden');
            document.getElementById('admin-dashboard-box').classList.remove('hidden');
            updateSubmissionLockUI();
            refreshAdminStats();
        }}

        function refreshAdminStats() {{
            const all = DB.achievers_may.concat(DB.achievers_jun);
            const total = all.length;
            let completed = 0;
            const brandCounts = {{}};
            BRANDS.forEach(b => {{
                brandCounts[b.name] = {{ may: 0, jun: 0, total: 0 }};
            }});

            all.forEach(a => {{
                const ch = savedChoices[a.id];
                if (ch && ch.voucher) {{
                    completed++;
                    if (brandCounts[ch.voucher]) {{
                        if (a.month === 'May_2026') brandCounts[ch.voucher].may++;
                        else brandCounts[ch.voucher].jun++;
                        brandCounts[ch.voucher].total++;
                    }}
                }}
            }});

            const elTotal = document.getElementById('admin-stat-total');
            const elCompleted = document.getElementById('admin-stat-completed');
            const elPct = document.getElementById('admin-stat-pct');
            const elPending = document.getElementById('admin-stat-pending');
            if (elTotal) elTotal.textContent = total.toLocaleString();
            if (elCompleted) elCompleted.textContent = completed.toLocaleString();
            const pct = total > 0 ? ((completed / total) * 100).toFixed(1) : '0.0';
            if (elPct) elPct.textContent = `${{pct}}%`;
            if (elPending) elPending.textContent = (total - completed).toLocaleString();

            const tbody = document.getElementById('admin-voucher-tbody');
            if (!tbody) return;
            tbody.innerHTML = '';

            BRANDS.forEach(b => {{
                const data = brandCounts[b.name] || {{ may: 0, jun: 0, total: 0 }};
                const share = completed > 0 ? ((data.total / completed) * 100).toFixed(1) : '0.0';
                const row = document.createElement('tr');
                row.innerHTML = `
                    <td class="py-2 px-3 font-bold text-slate-900 flex items-center gap-2.5">
                        <img src="${{b.logo}}" alt="${{b.name}}" class="h-6 w-14 object-contain">
                        <span>${{b.name}}</span>
                    </td>
                    <td class="py-2 px-3 text-center">${{data.may}}</td>
                    <td class="py-2 px-3 text-center">${{data.jun}}</td>
                    <td class="py-2 px-3 text-center font-black text-slate-900">${{data.total}}</td>
                    <td class="py-2 px-3 text-right font-bold text-orange-600">${{share}}%</td>
                `;
                tbody.appendChild(row);
            }});
        }}

        function updateSubmissionLockUI() {{
            const banner = document.getElementById('regview-locked-banner');
            const badge = document.getElementById('admin-lock-badge');
            const toggleBtn = document.getElementById('admin-toggle-lock-btn');

            if (banner) {{
                if (isSubmissionsLocked) {{
                    banner.classList.remove('hidden');
                }} else {{
                    banner.classList.add('hidden');
                }}
            }}

            if (badge) {{
                if (isSubmissionsLocked) {{
                    badge.className = 'px-2.5 py-0.5 rounded-full text-[10px] font-black bg-rose-100 text-rose-800 border border-rose-300';
                    badge.textContent = '🔒 Input OFF (Locked / Read-Only)';
                }} else {{
                    badge.className = 'px-2.5 py-0.5 rounded-full text-[10px] font-black bg-emerald-100 text-emerald-800 border border-emerald-300';
                    badge.textContent = '🟢 Input ON (Active)';
                }}
            }}

            if (toggleBtn) {{
                if (isSubmissionsLocked) {{
                    toggleBtn.className = 'w-full py-2 px-3 rounded-xl text-xs font-black transition flex items-center justify-center gap-2 shadow-xs bg-emerald-600 hover:bg-emerald-700 text-white';
                    toggleBtn.innerHTML = '<span>🔓 Turn ON Input (Allow Submissions)</span>';
                }} else {{
                    toggleBtn.className = 'w-full py-2 px-3 rounded-xl text-xs font-black transition flex items-center justify-center gap-2 shadow-xs bg-rose-600 hover:bg-rose-700 text-white';
                    toggleBtn.innerHTML = '<span>🔒 Turn OFF Input (Lock Submissions)</span>';
                }}
            }}

            if (currentRegionName) {{
                renderAchievers();
            }}
        }}

        async function toggleSubmissionLock() {{
            const targetState = !isSubmissionsLocked;
            const confirmMsg = targetState
                ? 'Are you sure you want to Turn OFF Regional Head input?\\n\\nThis will LOCK voucher selection nationwide into Read-Only mode. Regional Heads will not be able to select or remove vouchers.'
                : 'Are you sure you want to Turn ON Regional Head input?\\n\\nThis will UNLOCK the portal and ALLOW Regional Heads to select and submit voucher choices.';
            
            if (!confirm(confirmMsg)) return;

            isSubmissionsLocked = targetState;
            localStorage.setItem('EXIUM_SUBMISSIONS_LOCKED', isSubmissionsLocked ? 'true' : 'false');
            updateSubmissionLockUI();
            showToast(isSubmissionsLocked ? 'Input Locked (Read-Only)' : 'Input Unlocked (Active)', isSubmissionsLocked ? '🔒' : '🔓');

            if (cloudEndpoint) {{
                try {{
                    const url = cloudEndpoint + (cloudEndpoint.includes('?') ? '&' : '?') + 't=' + Date.now();
                    await fetch(url, {{
                        method: 'POST',
                        mode: 'no-cors',
                        headers: {{ 'Content-Type': 'application/json' }},
                        body: JSON.stringify({{
                            action: 'set_submission_lock',
                            locked: isSubmissionsLocked
                        }})
                    }});
                }} catch (e) {{
                    console.warn('Could not sync lock state to cloud:', e);
                }}
            }}
        }}

        async function adminDeleteAllNationalData() {{
            const confirm1 = confirm('⚠️ CAUTION: Are you sure you want to DELETE ALL voucher choices nationwide?\\n\\nThis will reset ALL 1,023 MIO awards back to "Pending". This action cannot be undone.');
            if (!confirm1) return;

            const typed = prompt('To confirm permanent deletion, type "DELETE" in capital letters:');
            if (typed !== 'DELETE') {{
                alert('Deletion cancelled. You must type "DELETE" exactly.');
                return;
            }}

            savedChoices = {{}};
            localStorage.setItem('EXIUM_AWARD_CHOICES', JSON.stringify(savedChoices));
            savePendingQueue({{}});

            if (currentRegionName) {{
                renderAchievers();
                updateRegionProgress();
            }}
            if (currentActiveZoneCode) {{
                renderZonalDashboardData(currentActiveZoneCode);
            }}
            refreshAdminStats();

            showToast('Resetting all national data...', '⏳');

            if (cloudEndpoint) {{
                try {{
                    const url = cloudEndpoint + (cloudEndpoint.includes('?') ? '&' : '?') + 't=' + Date.now();
                    await fetch(url, {{
                        method: 'POST',
                        mode: 'no-cors',
                        headers: {{ 'Content-Type': 'text/plain;charset=utf-8' }},
                        body: JSON.stringify({{
                            action: 'delete_all_data'
                        }})
                    }});
                    showToast('All National Data Cleared on Sheet & Portal!', '🗑️');
                }} catch (err) {{
                    console.error('Delete error:', err);
                    showToast('Local reset done. Could not reach cloud sheet.', '⚠️');
                }}
            }} else {{
                showToast('All Local Choices Cleared! (No cloud endpoint set)', '🗑️');
            }}
        }}

        async function adminDeleteCurrentRegionData() {{
            let targetRegion = currentRegionName;
            if (!targetRegion) {{
                const input = prompt('Enter the exact Region Name to reset (e.g. Dhaka South, Sylhet, etc.):');
                if (!input) return;
                targetRegion = input.trim();
            }}

            if (!DB.regions_map[targetRegion]) {{
                alert(`Region "${{targetRegion}}" not found! Please check spelling.`);
                return;
            }}

            const regMay = DB.achievers_may.filter(a => a.region === targetRegion);
            const regJun = DB.achievers_jun.filter(a => a.region === targetRegion);
            const totalInReg = regMay.length + regJun.length;

            const confirm1 = confirm(`Are you sure you want to delete and reset all voucher choices for Region: "${{targetRegion}}" (${{totalInReg}} awards)?\\n\\nThis will reset them back to Pending.`);
            if (!confirm1) return;

            // Clear choices for this region from local choices and pending queue
            const curQ = getPendingQueue();
            [...regMay, ...regJun].forEach(a => {{
                if (savedChoices[a.id]) {{
                    delete savedChoices[a.id];
                }}
                if (curQ[a.id]) {{
                    delete curQ[a.id];
                }}
            }});
            localStorage.setItem('EXIUM_AWARD_CHOICES', JSON.stringify(savedChoices));
            savePendingQueue(curQ);

            if (currentRegionName === targetRegion) {{
                renderAchievers();
                updateRegionProgress();
            }}
            if (currentActiveZoneCode) {{
                renderZonalDashboardData(currentActiveZoneCode);
            }}
            refreshAdminStats();

            showToast(`Resetting Region ${{targetRegion}}...`, '⏳');

            if (cloudEndpoint) {{
                try {{
                    const url = cloudEndpoint + (cloudEndpoint.includes('?') ? '&' : '?') + 't=' + Date.now();
                    await fetch(url, {{
                        method: 'POST',
                        mode: 'no-cors',
                        headers: {{ 'Content-Type': 'text/plain;charset=utf-8' }},
                        body: JSON.stringify({{
                            action: 'delete_region_data',
                            region_name: targetRegion
                        }})
                    }});
                    showToast(`Region "${{targetRegion}}" Data Cleared on Sheet!`, '🗑️');
                }} catch (err) {{
                    console.error('Delete region error:', err);
                    showToast('Region choices reset locally.', 'ℹ️');
                }}
            }} else {{
                showToast(`Region "${{targetRegion}}" choices cleared locally!`, '🗑️');
            }}
        }}

        let isSyncing = false;
        async function syncFromCloud(silent = false) {{
            if (!cloudEndpoint) {{
                if (!silent) showToast('Please set Google Apps Script URL in Admin Panel first!', '⚠️');
                updateSyncStatusBadge('offline');
                return;
            }}
            if (isSyncing) return;
            isSyncing = true;
            updateSyncStatusBadge('syncing');

            try {{
                // 1. Flush any pending uncommitted choices first
                await flushPendingQueue();

                // 2. Fetch latest snapshot from Google Sheets
                const url = cloudEndpoint + (cloudEndpoint.includes('?') ? '&' : '?') + 'action=fetch_data&t=' + Date.now();
                const res = await fetch(url);
                if (!res.ok) throw new Error('HTTP ' + res.status);
                const json = await res.json();

                if (json && json.status === 'success' && json.data) {{
                    if (json.submissions_locked !== undefined) {{
                        const cloudLocked = (json.submissions_locked === true || json.submissions_locked === 'true');
                        if (isSubmissionsLocked !== cloudLocked) {{
                            isSubmissionsLocked = cloudLocked;
                            localStorage.setItem('EXIUM_SUBMISSIONS_LOCKED', isSubmissionsLocked ? 'true' : 'false');
                            updateSubmissionLockUI();
                        }}
                    }}

                    const mayData = json.data.May_2026 || {{}};
                    const junData = json.data.June_2026 || {{}};
                    const pendingQ = getPendingQueue();
                    let appliedCount = 0;

                    // Anti-Blanking Merge Logic
                    const mergeCloudAchievers = (achieverList, cloudDict, monthKey) => {{
                        achieverList.forEach(a => {{
                            const key = a.area_code + '_' + a.mio_code;
                            const cloudItem = cloudDict[key];
                            const localItem = savedChoices[a.id];
                            const isLocalPending = !!pendingQ[a.id];

                            // If this device just clicked and is waiting to upload, do NOT overwrite it
                            if (isLocalPending) {{
                                return;
                            }}

                            if (cloudItem && cloudItem.voucher) {{
                                // Cloud has confirmed voucher
                                const prevVoucher = localItem ? (localItem.voucher || '') : '';
                                const prevTime = localItem ? (localItem.timestamp || '') : '';
                                if (prevVoucher !== cloudItem.voucher || prevTime !== cloudItem.timestamp) {{
                                    savedChoices[a.id] = {{
                                        voucher: cloudItem.voucher,
                                        timestamp: cloudItem.timestamp || '',
                                        status: 'Complete'
                                    }};
                                    appliedCount++;
                                }}
                            }} else if (cloudItem && cloudItem.status === 'Pending' && !cloudItem.voucher) {{
                                // Cloud explicitly has Pending / Deselected
                                if (localItem && localItem.voucher) {{
                                    savedChoices[a.id] = {{
                                        voucher: '',
                                        timestamp: cloudItem.timestamp || localItem.timestamp || '',
                                        status: 'Pending'
                                    }};
                                    appliedCount++;
                                }}
                            }} else if (!cloudItem) {{
                                // Cloud does not have any record for this achiever yet
                                // ANTI-BLANKING RULE: NEVER ERASE LOCAL CHOICE!
                                // If this device has a valid choice, preserve it and queue for upload!
                                if (localItem && localItem.voucher) {{
                                    queueForUpload(a.id, monthKey, a.area_code, a.mio_code, localItem.voucher, localItem.timestamp || getBDTimestamp(), 'save');
                                }}
                            }}
                        }});
                    }};

                    mergeCloudAchievers(DB.achievers_may, mayData, 'May_2026');
                    mergeCloudAchievers(DB.achievers_jun, junData, 'June_2026');

                    localStorage.setItem('EXIUM_AWARD_CHOICES', JSON.stringify(savedChoices));

                    if (appliedCount > 0 || !silent) {{
                        if (currentRegionName) {{
                            renderAchievers();
                            updateRegionProgress();
                        }}
                        if (currentActiveZoneCode) {{
                            renderZonalDashboardData(currentActiveZoneCode);
                        }}
                        if (isAdminUnlocked) {{
                            refreshAdminStats();
                        }}
                    }}

                    const badgeStatus = document.getElementById('admin-endpoint-status');
                    if (badgeStatus) badgeStatus.classList.remove('hidden');
                    updateShareableLinkBox();
                    updateSyncStatusBadge('synced');

                    // If we queued any missing local choices, flush them immediately
                    const finalQ = getPendingQueue();
                    if (Object.keys(finalQ).length > 0) {{
                        flushPendingQueue();
                    }}

                    if (!silent) {{
                        showToast(`Cloud Synced! (${{appliedCount}} updated)`, '☁️');
                    }}
                }}
            }} catch (err) {{
                console.warn('Cloud sync error:', err);
                updateSyncStatusBadge('pending');
                if (!silent) showToast('Cloud sync failed. Check URL or internet.', '⚠️');
            }} finally {{
                isSyncing = false;
            }}
        }}

        function saveAdminEndpoint() {{
            const val = document.getElementById('admin-endpoint-input').value.trim();
            cloudEndpoint = val;
            localStorage.setItem('EXIUM_AWARD_ENDPOINT', val);
            const badgeStatus = document.getElementById('admin-endpoint-status');
            if (val) {{
                if (badgeStatus) badgeStatus.classList.remove('hidden');
                updateShareableLinkBox();
                showToast('Endpoint URL saved! Testing live sync...', '⚙️');
                flushPendingQueue();
                syncFromCloud(false);
            }} else {{
                if (badgeStatus) badgeStatus.classList.add('hidden');
                updateShareableLinkBox();
                updateSyncStatusBadge('offline');
                showToast('Endpoint URL cleared.', 'ℹ️');
            }}
        }}

        async function fixGoogleSheetFormatting() {{
            if (!cloudEndpoint) {{
                showToast('Please enter and save Google Apps Script URL first!', '⚠️');
                return;
            }}
            showToast('Repairing Google Sheet formatting & timestamps...', '⏳');
            try {{
                const url = cloudEndpoint + (cloudEndpoint.includes('?') ? '&' : '?') + 'action=fix_formatting&t=' + Date.now();
                const res = await fetch(url);
                const json = await res.json();
                if (json && json.status === 'success') {{
                    showToast(json.message || 'Google Sheet Ach% & Timestamps Fixed!', '✅');
                    syncFromCloud(false);
                }} else {{
                    showToast('Response: ' + (json.message || 'Complete'), 'ℹ️');
                }}
            }} catch (err) {{
                showToast('Fix command sent to Google Sheet!', '✅');
            }}
        }}

        function exportToExcel(type = 'admin') {{
            let mayList = [];
            let junList = [];
            let filePrefix = 'Exium_Award_Master';

            if (type === 'region') {{
                const reg = currentRegionName;
                if (!reg) {{
                    showToast('Please open a region first!', '⚠️');
                    return;
                }}
                mayList = DB.achievers_may.filter(x => x.region === reg);
                junList = DB.achievers_jun.filter(x => x.region === reg);
                filePrefix = `Exium_Award_Region_${{reg.replace(new RegExp('[ _/-]+', 'g'), '_')}}`;
            }} else if (type === 'zone') {{
                const zoneObj = DB.zones.find(z => z.sap_zone_code === currentActiveZoneCode);
                if (!zoneObj) {{
                    showToast('Please select a zone first!', '⚠️');
                    return;
                }}
                mayList = DB.achievers_may.filter(x => x.zone === zoneObj.zone_name);
                junList = DB.achievers_jun.filter(x => x.zone === zoneObj.zone_name);
                filePrefix = `Exium_Award_Zone_${{zoneObj.zone_name.replace(new RegExp('[ _/-]+', 'g'), '_')}}`;
            }} else {{
                mayList = DB.achievers_may;
                junList = DB.achievers_jun;
                filePrefix = 'Exium_Award_National_Master_2026';
            }}

            const formatAchieverRow = (a, isJune) => {{
                const ch = savedChoices[a.id] || {{}};
                const voucher = (ch.voucher !== undefined) ? ch.voucher : (a.choice || '');
                const rawTimestamp = ch.timestamp || a.timestamp || '';
                const timestamp = formatToBDTime(rawTimestamp);
                const status = voucher ? 'Complete' : 'Pending';
                const td = a.transfer_details || {{}};

                const r = {{
                    'Month': isJune ? 'June 2026' : 'May 2026',
                    'Current Zone': a.zone,
                    'Current SAP Zone': a.sap_zone_code || '',
                    'Zonal Head': a.zonal_head || '',
                    'Current Region': a.region,
                    'Current SAP Region': a.sap_region_code || '',
                    'Current Regional Head': a.regional_head || '',
                    'Current Area Code': a.area_code,
                    'Current Area Name': a.area_name,
                    'SAP MIO Code': a.mio_code,
                    'Sr./ MIO Name': a.mio_name,
                    'Designation': a.desig || 'MIO',
                    'No. of Rx': a.rx || 0
                }};

                if (isJune) {{
                    r['Metric (Growth% & Sales)'] = a.metric_val || '';
                }} else {{
                    r['Metric (Ach%)'] = a.metric_val || '';
                }}

                r['Award Amount (BDT)'] = a.award || 0;
                r['Selected Gift Voucher'] = voucher || 'Not Selected';
                r['Submission Timestamp'] = timestamp;
                r['Status'] = status;
                r['Is Transferred'] = a.is_transferred ? 'Yes' : 'No';
                r['Previous Area Code'] = td.prev_area_code || '';
                r['Previous Area Name'] = td.prev_area_name || '';
                r['Previous Region'] = td.prev_region || '';
                r['Previous Regional Head'] = td.prev_rh || '';
                r['Previous Zone'] = td.prev_zone || '';

                return r;
            }};

            const mayRows = mayList.map(a => formatAchieverRow(a, false));
            const junRows = junList.map(a => formatAchieverRow(a, true));
            const dateTag = new Date().toISOString().slice(0, 10);

            if (typeof XLSX !== 'undefined') {{
                const wb = XLSX.utils.book_new();

                if (mayRows.length > 0) {{
                    const wsMay = XLSX.utils.json_to_sheet(mayRows);
                    wsMay['!cols'] = [
                        {{ wch: 12 }}, {{ wch: 16 }}, {{ wch: 16 }}, {{ wch: 18 }}, {{ wch: 20 }},
                        {{ wch: 18 }}, {{ wch: 20 }}, {{ wch: 16 }}, {{ wch: 20 }}, {{ wch: 14 }},
                        {{ wch: 22 }}, {{ wch: 12 }}, {{ wch: 10 }}, {{ wch: 16 }}, {{ wch: 18 }},
                        {{ wch: 24 }}, {{ wch: 22 }}, {{ wch: 12 }}, {{ wch: 14 }}, {{ wch: 16 }},
                        {{ wch: 20 }}, {{ wch: 20 }}, {{ wch: 20 }}, {{ wch: 16 }}
                    ];
                    XLSX.utils.book_append_sheet(wb, wsMay, 'Award_May_2026');
                }}

                if (junRows.length > 0) {{
                    const wsJun = XLSX.utils.json_to_sheet(junRows);
                    wsJun['!cols'] = [
                        {{ wch: 12 }}, {{ wch: 16 }}, {{ wch: 16 }}, {{ wch: 18 }}, {{ wch: 20 }},
                        {{ wch: 18 }}, {{ wch: 20 }}, {{ wch: 16 }}, {{ wch: 20 }}, {{ wch: 14 }},
                        {{ wch: 22 }}, {{ wch: 12 }}, {{ wch: 10 }}, {{ wch: 24 }}, {{ wch: 18 }},
                        {{ wch: 24 }}, {{ wch: 22 }}, {{ wch: 12 }}, {{ wch: 14 }}, {{ wch: 16 }},
                        {{ wch: 20 }}, {{ wch: 20 }}, {{ wch: 20 }}, {{ wch: 16 }}
                    ];
                    XLSX.utils.book_append_sheet(wb, wsJun, 'Award_June_2026');
                }}

                XLSX.writeFile(wb, `${{filePrefix}}_${{dateTag}}.xlsx`);
                showToast('Excel exported successfully!', '📊');
            }} else {{
                // Fallback to UTF-8 BOM CSV
                downloadFallbackCsv(mayRows.concat(junRows), `${{filePrefix}}_${{dateTag}}.csv`);
                showToast('Exported CSV successfully!', '📊');
            }}
        }}

        function downloadFallbackCsv(rows, filename) {{
            if (!rows || !rows.length) return;
            const headers = Object.keys(rows[0]);
            let csv = '\\uFEFF' + headers.map(h => `"${{h}}"`).join(',') + '\\r\\n';
            rows.forEach(r => {{
                csv += headers.map(h => {{
                    let v = (r[h] === null || r[h] === undefined) ? '' : String(r[h]);
                    return `"${{v.replace(/"/g, '""')}}"`;
                }}).join(',') + '\\r\\n';
            }});
            const blob = new Blob([csv], {{ type: 'text/csv;charset=utf-8;' }});
            const link = document.createElement('a');
            link.href = URL.createObjectURL(blob);
            link.download = filename;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }}

        function downloadCsvExport() {{
            exportToExcel('admin');
        }}

        function showToast(msg, icon = '✅') {{
            const t = document.getElementById('toast-box');
            document.getElementById('toast-text').textContent = msg;
            document.getElementById('toast-icon').textContent = icon;
            t.classList.remove('translate-x-32', 'opacity-0', 'pointer-events-none');
            setTimeout(() => {{
                t.classList.add('translate-x-32', 'opacity-0', 'pointer-events-none');
            }}, 3000);
        }}
    </script>
</body>
</html>
"""

if __name__ == "__main__":
    build_portal()
