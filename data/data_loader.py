import os
import pandas as pd

ACTIVE_STATES = {1, "1", "在线", "正常", "ON", "on", "true", "True", "是", "Y"}

REQUIRED_FILES = {
    'sections': '断面关系表.xlsx',
    'stations': '电站信息表.xlsx',
    'users': '负荷信息表.xlsx',
    'template': '导入数据（测试）.xlsx'
}

def _find_col(df, candidates):
    for c in candidates:
        if c in df.columns:
            return c
    return None

def _read_excel_strict(path):
    try:
        return pd.read_excel(path, dtype=object)
    except Exception as e:
        raise ValueError(f"Failed to read Excel file {path}: {e}")

def _split_points(val):
    if pd.isna(val):
        return []
    if isinstance(val, (int, float)):
        val = str(int(val))
    return [p.strip() for p in str(val).replace('\r',';').replace('\n',';').replace(',', ';').split(';') if p.strip()]

def load_data(root_dir='.'):  # check files exist
    paths = {}
    for key, name in REQUIRED_FILES.items():
        p = os.path.join(root_dir, name)
        if not os.path.exists(p):
            raise ValueError(f"Required file missing: {name} (expected at repository root)")
        paths[key] = p

    # load sections
    df_sec = _read_excel_strict(paths['sections'])
    sec_id_col = _find_col(df_sec, ['断面编号', '断面ID', '断面'])
    sec_name_col = _find_col(df_sec, ['断面名称', '断面名'])
    sec_point_col = _find_col(df_sec, ['点位', '点位编码'])
    if not sec_id_col or not sec_name_col or not sec_point_col:
        raise ValueError(f"断面关系表.xlsx must contain columns: 断面编号, 断面名称, 点位 (alternatives supported)")

    sections = {}
    point_to_sections = {}
    for _, row in df_sec.iterrows():
        sec_id = str(row[sec_id_col]).strip()
        sec_name = str(row[sec_name_col]).strip()
        pts = _split_points(row[sec_point_col])
        sections[sec_id] = {'id': sec_id, 'name': sec_name, 'points': pts}
        for p in pts:
            point_to_sections.setdefault(p, []).append(sec_id)

    # load stations
    df_st = _read_excel_strict(paths['stations'])
    st_id_col = _find_col(df_st, ['电站编号', '编号', '电站ID'])
    st_name_col = _find_col(df_st, ['电站名称', '名称'])
    st_output_col = _find_col(df_st, ['出力', '出力(MW)', '容量'])
    st_point_col = _find_col(df_st, ['点位', '点位编码'])
    if not st_id_col or not st_name_col or not st_output_col or not st_point_col:
        raise ValueError('电站信息表.xlsx must contain columns: 电站编号/编号, 电站名称, 出力/出力(MW), 点位')

    stations = {}
    for _, row in df_st.iterrows():
        sid = str(row[st_id_col]).strip()
        sname = str(row[st_name_col]).strip()
        try:
            sout = float(row[st_output_col]) if not pd.isna(row[st_output_col]) else 0.0
        except Exception:
            sout = 0.0
        pts = _split_points(row[st_point_col])
        stations[sid] = {
            'id': sid,
            'name': sname,
            'output': float(sout),
            'points': pts,
            'participate': True,
            'adjustment': 0.0,
            'type': None,  # optional D/T
            'section_ids': []
        }

    # load users
    df_usr = _read_excel_strict(paths['users'])
    usr_id_col = _find_col(df_usr, ['用户编号', '编号', '用户ID'])
    usr_name_col = _find_col(df_usr, ['用电用户', '用户', '用户名'])
    usr_load_col = _find_col(df_usr, ['负荷', '负荷(MW)', '负荷(kW)'])
    usr_point_col = _find_col(df_usr, ['点位', '点位编码'])
    if not usr_id_col or not usr_name_col or not usr_load_col or not usr_point_col:
        raise ValueError('负荷信息表.xlsx must contain columns: 用户编号/编号, 用电用户/用户, 负荷/负荷(MW), 点位')

    users = {}
    for _, row in df_usr.iterrows():
        uid = str(row[usr_id_col]).strip()
        uname = str(row[usr_name_col]).strip()
        try:
            uload = float(row[usr_load_col]) if not pd.isna(row[usr_load_col]) else 0.0
        except Exception:
            uload = 0.0
        pts = _split_points(row[usr_point_col])
        users[uid] = {'id': uid, 'name': uname, 'load': float(uload), 'points': pts, 'section_id': None}

    # load template / point status
    df_tmp = _read_excel_strict(paths['template'])
    # the template may have point->status as two-column or columns are points with first row statuses
    point_status = {}
    if df_tmp.shape[1] == 2 and '点位' in df_tmp.columns and '状态' in df_tmp.columns:
        # two-column layout
        for _, r in df_tmp.iterrows():
            p = str(r[df_tmp.columns[0]]).strip()
            s = r[df_tmp.columns[1]]
            point_status[p] = s
    else:
        # try first row as statuses and columns are points
        first_row = df_tmp.iloc[0]
        for col in df_tmp.columns:
            p = str(col).strip()
            s = first_row[col]
            point_status[p] = s

    # normalize point status values
    def _is_active(v):
        if pd.isna(v):
            return False
        if isinstance(v, (int, float)):
            v2 = int(v)
        else:
            v2 = str(v).strip()
        return v2 in ACTIVE_STATES

    active_points = {p for p, s in point_status.items() if _is_active(s)}

    # map stations/users to sections strictly: station is assigned to a section only if it has a point listed that exists in point-status and is active
    for sid, s in stations.items():
        assigned = set()
        for p in s['points']:
            if p in active_points and p in point_to_sections:
                for sec in point_to_sections[p]:
                    assigned.add(sec)
        s['section_ids'] = sorted(list(assigned))

    for uid, u in users.items():
        assigned = set()
        for p in u['points']:
            if p in active_points and p in point_to_sections:
                for sec in point_to_sections[p]:
                    assigned.add(sec)
        # Strict rule: only assign if at least one active point maps to a section; pick first if multiple
        u['section_id'] = sorted(list(assigned))[0] if assigned else None

    return {'sections': sections, 'stations': stations, 'users': users, 'point_status': point_status}