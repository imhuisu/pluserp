import tkinter as tk
from tkinter import messagebox, filedialog, ttk
import datetime
import os
import subprocess
import tempfile

# ═══════════════════════════════════════════════════════
#  버전
# ═══════════════════════════════════════════════════════
VERSION = "ver.2606140605"

# ═══════════════════════════════════════════════════════
#  상수
# ═══════════════════════════════════════════════════════
BAR_ENDS   = 80
SAW        = 8
REG05_OPTS = [4300, 5000]
REG06_LEN  = 5700

HINGE_OPTIONS = ["좌측", "우측"]
FIX_OPTIONS   = ["없음", "우측", "좌측", "양측"]

# ═══════════════════════════════════════════════════════
#  REG-06 수식 계산 (BASE.xlsx 수식 기반)
#  frame  = 폭(C)
#  door   = 문폭(E)
#  dc_base= 245 (door closer 기준값)
#  reg06  = frame - 100
#  door_sz= frame-80 if FIX==없음 else door
# ═══════════════════════════════════════════════════════
DC_BASE = 245

def calc_reg06(frame, door, hinge, fix):
    """
    엑셀 수식 기반으로 REG-06 관련 값 계산
    반환: dict(reg06, dc, fix1, fix2)
    """
    reg06   = frame - 100
    door_sz = (frame - 80) if fix == '없음' else door

    fix1, fix2 = None, None

    if hinge == '좌측':
        if fix == '없음':
            dc = DC_BASE
        elif fix == '우측':
            dc   = DC_BASE
            fix1 = (door_sz - 20) + 17.5
        elif fix == '좌측':
            dc   = reg06 - (door_sz - 20) + DC_BASE
            fix1 = reg06 - (door_sz - 20) - 17.5
        elif fix == '양측':
            dc   = (reg06 / 2) - ((door_sz - 20) / 2) + DC_BASE
            fix1 = (reg06 / 2) - ((door_sz - 20) / 2) - 17.5
            fix2 = (reg06 / 2) + ((door_sz - 20) / 2) + 17.5
    else:  # 우측
        if fix == '없음':
            dc = reg06 - DC_BASE
        elif fix == '우측':
            dc   = door_sz - 10 - DC_BASE
            fix1 = (door_sz - 20) + 17.5
        elif fix == '좌측':
            dc   = reg06 - DC_BASE
            fix1 = reg06 - (door_sz - 20) - 17.5
        elif fix == '양측':
            dc   = (reg06 / 2) + ((door_sz - 20) / 2) - DC_BASE
            fix1 = (reg06 / 2) - ((door_sz - 20) / 2) - 17.5
            fix2 = (reg06 / 2) + ((door_sz - 20) / 2) + 17.5

    return dict(reg06=reg06, dc=dc, fix1=fix1, fix2=fix2)

# ═══════════════════════════════════════════════════════
#  REG-05 Bar 길이 선택
# ═══════════════════════════════════════════════════════
def usable_2piece(bar_len):
    return bar_len - BAR_ENDS - SAW - SAW - BAR_ENDS - SAW

def usable_1piece(bar_len):
    return bar_len - BAR_ENDS - SAW - BAR_ENDS - SAW

def choose_reg05_bar(reg05_len):
    for bl in REG05_OPTS:
        if 2 * reg05_len <= usable_2piece(bl):
            return bl, 2
    for bl in REG05_OPTS:
        if reg05_len <= usable_1piece(bl):
            return bl, 1
    return None, None

# ═══════════════════════════════════════════════════════
#  그리디 패킹
# ═══════════════════════════════════════════════════════
def pack_pieces(lengths, bar_len):
    limit = bar_len - BAR_ENDS - SAW - BAR_ENDS - SAW
    bars, current, cur_sum = [], [], 0
    for ln in sorted(lengths, reverse=True):
        n = len(current)
        if cur_sum + ln + n * SAW <= limit:
            current.append(ln)
            cur_sum += ln
        else:
            if current:
                bars.append(current)
            current, cur_sum = [ln], ln
    if current:
        bars.append(current)
    return bars

# ═══════════════════════════════════════════════════════
#  헬퍼
# ═══════════════════════════════════════════════════════
def size_str(w, h):
    return f"W{w}" if h <= 2000 else f"W{w}*H{h}"

def piece_id(seq, date_suffix):
    return f"{seq:05d}{date_suffix}"

def bar_angles():
    return [('LeftAngle','90'),('RightAngle','90'),
            ('LeftHeightAngle','0'),('RightHeightAngle','0'),
            ('LeftWidthAngle','0'),('RightWidthAngle','0')]

def frame_angles():
    return [('LeftAngle','0'),('RightAngle','0'),
            ('LeftHeightAngle','90'),('RightHeightAngle','90'),
            ('LeftWidthAngle','90'),('RightWidthAngle','90')]

def angles_block(pairs, indent):
    L = [f'{indent}<Angles>']
    for k, v in pairs:
        L.append(f'{indent}  <{k}>{v}</{k}>')
    L.append(f'{indent}</Angles>')
    return L

def info_fields_numbered(indent):
    L = [f'{indent}<InfoFields>']
    for i in range(1, 5):
        L.append(f'{indent}  <Info id="{i}" />')
    L.append(f'{indent}</InfoFields>')
    return L

def piece_info_fields(label, sz, indent):
    return [f'{indent}<InfoFields>',
            f'{indent}  <Info id="" />',
            f'{indent}  <Info id="{label}" />',
            f'{indent}  <Info id="{sz}" />',
            f'{indent}  <Info id="Area00" />',
            f'{indent}</InfoFields>']

def make_piece(pid, length, label, sz, macro_name, pos_x, date_suffix, indent='      '):
    i2 = indent + '  '
    L = [f'{indent}<Piece ID="{pid}" Length="{length}" Quantity="1" '
         f'enabled="true" mat_type="1" wash_loss="0" is_wash="3">']
    L += angles_block(frame_angles(), i2)
    L += piece_info_fields(label, sz, i2)
    L.append(f'{i2}<Machinings>')
    L.append(f'{i2}  <Macro Type="MCR" Name="{macro_name}" PositionX="{pos_x}" Comment="0" />')
    L.append(f'{i2}</Machinings>')
    L.append(f'{indent}</Piece>')
    return L

def make_piece_multi_macro(pid, length, label, sz, macros, date_suffix, indent='      '):
    i2 = indent + '  '
    L = [f'{indent}<Piece ID="{pid}" Length="{length}" Quantity="1" '
         f'enabled="true" mat_type="1" wash_loss="0" is_wash="3">']
    L += angles_block(frame_angles(), i2)
    L += piece_info_fields(label, sz, i2)
    L.append(f'{i2}<Machinings>')
    for mname, mpos in macros:
        L.append(f'{i2}  <Macro Type="MCR" Name="{mname}" PositionX="{mpos}" Comment="0" />')
    L.append(f'{i2}</Machinings>')
    L.append(f'{indent}</Piece>')
    return L

def make_bar(bar_id, bar_len, pieces_lines, indent='    '):
    i2 = indent + '  '
    L = [f'{indent}<Bar ID="{bar_id}" Length="{bar_len}" Quantity="1" enabled="true">']
    L += angles_block(bar_angles(), i2)
    L += info_fields_numbered(i2)
    L += pieces_lines
    L.append(f'{indent}</Bar>')
    return L

# ═══════════════════════════════════════════════════════
#  XML 생성
# ═══════════════════════════════════════════════════════
def generate_xml(rows, save_path):
    today       = datetime.date.today()
    date_str    = f"{today.day:02d}.{today.month:02d}.{today.year}"
    date_suffix = f"{today.year % 100:02d}{today.month:02d}{today.day:02d}14"

    rows_sorted = sorted(rows, key=lambda x: x['height'], reverse=True)
    n = len(rows_sorted) - 1

    L = []
    L.append('<?xml version="1.0" encoding="utf-8" standalone="no"?>')
    L.append('<Unilink>')
    L.append(f'  <FileInfo CreatedBy="string" CreationTime="{date_str}" />')
    L.append('')

    # ── REG-05 ──
    L.append('')
    L.append('')
    L.append('')
    L.append('  <Profile Serie="REGEN" Name="REG-05" Width="130" Height="70" enabled="1">')
    L.append('    <Color>')
    L.append('      <Inside Color="  ũ    " />')
    L.append('      <Outside Color="  ũ    " />')
    L.append('    </Color>')
    L += info_fields_numbered('    ')

    bar_id = 1
    for idx, row in enumerate(rows_sorted):
        inv       = n - idx
        r_seq     = inv * 32 + 8
        l_seq     = inv * 32 + 7
        reg05_len = float(row['height'] - 30)
        sz        = size_str(row['width'], row['height'])
        bar_len, per_bar = choose_reg05_bar(int(reg05_len))

        if bar_len is None:
            raise ValueError(f"'{row['name']}' REG-05 길이({reg05_len}mm) 원재료 범위 초과")

        if per_bar == 2:
            rh = make_piece(piece_id(r_seq, date_suffix), reg05_len,
                            f'{row["name"]}-FrameR', sz, 'REG-05_RH', '16.1', date_suffix)
            for li, line in enumerate(rh):
                if 'REG-05_RH' in line:
                    rh.insert(li+1, ''); break
            lh = make_piece(piece_id(l_seq, date_suffix), reg05_len,
                            f'{row["name"]}-FrameL', sz, 'REG-05_LH', '16.1', date_suffix)
            L += make_bar(bar_id, bar_len, rh + lh); bar_id += 1
        else:
            rh = make_piece(piece_id(r_seq, date_suffix), reg05_len,
                            f'{row["name"]}-FrameR', sz, 'REG-05_RH', '16.1', date_suffix)
            for li, line in enumerate(rh):
                if 'REG-05_RH' in line:
                    rh.insert(li+1, ''); break
            L += make_bar(bar_id, bar_len, rh); bar_id += 1
            lh = make_piece(piece_id(l_seq, date_suffix), reg05_len,
                            f'{row["name"]}-FrameL', sz, 'REG-05_LH', '16.1', date_suffix)
            L += make_bar(bar_id, bar_len, lh); bar_id += 1

    L.append('  </Profile>')
    L.append('')

    # ── REG-08 (T/B 분리) ──
    L.append('')
    L.append('')
    L.append('')
    L.append('  ')
    L.append('')
    L.append('')
    L.append('')
    L.append('')
    L.append('  <Profile Serie="REGEN" Name="REG-08" Width="130" Height="15" enabled="1">')
    L.append('    <Color>')
    L.append('      <Inside Color="  ũ    " />')
    L.append('      <Outside Color="  ũ    " />')
    L.append('    </Color>')
    L += info_fields_numbered('    ')

    for frame_type, macro_name, pos_x in [('T','REG-08_T','5.5'), ('B','REG-08_B','5.5')]:
        widths = sorted([r['width'] for r in rows], reverse=True)
        for bar_widths in pack_pieces(widths, 5700):
            pieces_lines = []
            for w in bar_widths:
                row = next(r for r in rows if r['width'] == w)
                orig_idx = next(i for i, r in enumerate(rows_sorted) if r['name'] == row['name'])
                inv  = n - orig_idx
                seq  = inv * 32 + (2 if frame_type == 'T' else 5)
                sz   = size_str(row['width'], row['height'])
                pieces_lines += make_piece(
                    piece_id(seq, date_suffix), float(w),
                    f'{row["name"]}-Frame{frame_type}', sz, macro_name, pos_x, date_suffix)
            L += make_bar(bar_id, 5700, pieces_lines); bar_id += 1

    L.append('')
    L.append('  </Profile>')
    L.append('')

    # ── REG-06 ──
    L.append('')
    L.append('')
    L.append('')
    L.append('  <Profile Serie="REGEN" Name="REG-06" Width="130" Height="20" enabled="1">')
    L.append('    <Color>')
    L.append('      <Inside Color="  ũ    " />')
    L.append('      <Outside Color="  ũ    " />')
    L.append('    </Color>')
    L += info_fields_numbered('    ')

    reg06_lengths = [r['lk']['reg06'] for r in rows_sorted]
    for bar_lengths in pack_pieces(reg06_lengths, REG06_LEN):
        pieces_lines = []
        for ln in bar_lengths:
            row = next(r for r in rows_sorted if r['lk']['reg06'] == ln)
            orig_idx = next(i for i, r in enumerate(rows_sorted) if r['name'] == row['name'])
            inv = n - orig_idx
            seq = inv * 32 + 6
            lk  = row['lk']
            sz  = size_str(row['width'], row['height'])

            macros = []
            # 1. door closer position
            macros.append(('REG-06_DC', str(lk['dc'])))
            # 2. fix position (1개 또는 2개)
            if lk['fix1'] is not None:
                macros.append(('REG-06_FIX', str(lk['fix1'])))
            if lk['fix2'] is not None:
                macros.append(('REG-06_FIX', str(lk['fix2'])))
            # 3. REG-06-20 (항상 0으로 출력)
            macros.append(('REG-06-20', '0'))

            pieces_lines += make_piece_multi_macro(
                piece_id(seq, date_suffix), float(ln),
                f'{row["name"]}-REG06', sz, macros, date_suffix)
        L += make_bar(bar_id, REG06_LEN, pieces_lines); bar_id += 1

    L.append('')
    L.append('  </Profile>')
    L.append('')
    L.append('</Unilink>')

    content = '\r\n'.join(L)
    with open(save_path, 'w', encoding='utf-8') as f:
        f.write('\ufeff')
        f.write(content)


# ═══════════════════════════════════════════════════════
#  GUI
# ═══════════════════════════════════════════════════════
class App:
    MAX_ROWS = 15

    def __init__(self, root):
        self.root = root
        root.title("REGEN XML 생성기")
        root.resizable(False, False)

        # 타이틀
        tf = tk.Frame(root, bg="#2C3E50", pady=10)
        tf.pack(fill='x')
        title_inner = tk.Frame(tf, bg="#2C3E50")
        title_inner.pack()
        tk.Label(title_inner, text="REGEN XML 생성기",
                 font=("맑은 고딕", 16, "bold"), fg="white", bg="#2C3E50").pack(side='left')
        tk.Label(title_inner, text=f"  {VERSION}",
                 font=("맑은 고딕", 9), fg="#95A5A6", bg="#2C3E50").pack(side='left', pady=(6,0))
        tk.Label(tf, text="노란색 셀에 데이터 입력 후 [XML 생성] 클릭",
                 font=("맑은 고딕", 9), fg="#BDC3C7", bg="#2C3E50").pack()

        # 헤더
        hf = tk.Frame(root, bg="#ECF0F1", pady=6)
        hf.pack(fill='x', padx=15)
        headers = [("No.",4),("업체명",12),("폭 W",7),("높이 H",7),("문폭",7),
                   ("경첩방향",10),("FIX",10),("REG-05",9),("원재료",8),("REG-08",9),("REG-06",9),("DC위치",9),("FIX위치",16),("REG-06-20",9)]
        for col, (text, w) in enumerate(headers):
            tk.Label(hf, text=text, font=("맑은 고딕", 9, "bold"),
                     bg="#ECF0F1", fg="#2C3E50", width=w, anchor='center').grid(
                row=0, column=col, padx=2)

        # 콤보박스 노란색 스타일
        style = ttk.Style()
        style.theme_use('default')
        style.configure('Yellow.TCombobox', fieldbackground='#FFFF00', background='#FFFF00',
                        selectbackground='#FFFF00', selectforeground='black')

        # 입력행
        self.input_frame = tk.Frame(root, bg="white")
        self.input_frame.pack(fill='x', padx=15, pady=2)

        YELLOW = "#FFFF00"
        self.name_vars   = []
        self.width_vars  = []
        self.height_vars = []
        self.door_vars   = []
        self.hinge_vars  = []
        self.fix_vars    = []
        self.reg05_lbls  = []
        self.barlen_lbls = []
        self.reg08_lbls  = []
        self.reg06_lbls  = []
        self.dc_lbls     = []
        self.fixpos_lbls = []
        self.r620_lbls   = []

        for i in range(self.MAX_ROWS):
            bg = "white" if i % 2 == 0 else "#F8F9FA"
            tk.Label(self.input_frame, text=str(i+1), font=("맑은 고딕", 9),
                     bg=bg, width=4, anchor='center').grid(row=i, column=0, padx=2, pady=2)

            nv  = tk.StringVar()
            wv  = tk.StringVar()
            hv  = tk.StringVar()
            dv  = tk.StringVar()
            hgv = tk.StringVar(value=HINGE_OPTIONS[0])
            fxv = tk.StringVar(value=FIX_OPTIONS[0])

            tk.Entry(self.input_frame, textvariable=nv, bg=YELLOW,
                     font=("맑은 고딕", 9), width=12, justify='center').grid(row=i, column=1, padx=2, pady=2)
            tk.Entry(self.input_frame, textvariable=wv, bg=YELLOW,
                     font=("맑은 고딕", 9), width=7, justify='center').grid(row=i, column=2, padx=2, pady=2)
            tk.Entry(self.input_frame, textvariable=hv, bg=YELLOW,
                     font=("맑은 고딕", 9), width=7, justify='center').grid(row=i, column=3, padx=2, pady=2)
            tk.Entry(self.input_frame, textvariable=dv, bg=YELLOW,
                     font=("맑은 고딕", 9), width=7, justify='center').grid(row=i, column=4, padx=2, pady=2)

            hg_cb = ttk.Combobox(self.input_frame, textvariable=hgv,
                                  values=HINGE_OPTIONS, state='readonly',
                                  width=8, font=("맑은 고딕", 9), style='Yellow.TCombobox')
            hg_cb.grid(row=i, column=5, padx=2, pady=2)

            fx_cb = ttk.Combobox(self.input_frame, textvariable=fxv,
                                  values=FIX_OPTIONS, state='readonly',
                                  width=8, font=("맑은 고딕", 9), style='Yellow.TCombobox')
            fx_cb.grid(row=i, column=6, padx=2, pady=2)

            r05 = tk.Label(self.input_frame, text="-", font=("맑은 고딕", 9),
                           bg=bg, width=9, anchor='center', fg="#555")
            rbl = tk.Label(self.input_frame, text="-", font=("맑은 고딕", 9),
                           bg=bg, width=8, anchor='center', fg="#555")
            r08 = tk.Label(self.input_frame, text="-", font=("맑은 고딕", 9),
                           bg=bg, width=9, anchor='center', fg="#555")
            r06 = tk.Label(self.input_frame, text="-", font=("맑은 고딕", 9),
                           bg=bg, width=9, anchor='center', fg="#555")

            r05.grid(row=i, column=7,  padx=2, pady=2)
            rbl.grid(row=i, column=8,  padx=2, pady=2)
            r08.grid(row=i, column=9,  padx=2, pady=2)
            r06.grid(row=i, column=10, padx=2, pady=2)

            dc_lbl  = tk.Label(self.input_frame, text="-", font=("맑은 고딕", 9),
                               bg=bg, width=9, anchor='center', fg="#555")
            fix_lbl = tk.Label(self.input_frame, text="-", font=("맑은 고딕", 9),
                               bg=bg, width=16, anchor='center', fg="#555")
            r620_lbl= tk.Label(self.input_frame, text="-", font=("맑은 고딕", 9),
                               bg=bg, width=9, anchor='center', fg="#555")
            dc_lbl.grid(row=i,   column=11, padx=2, pady=2)
            fix_lbl.grid(row=i,  column=12, padx=2, pady=2)
            r620_lbl.grid(row=i, column=13, padx=2, pady=2)

            wv.trace_add('write',  lambda *a, i=i: self.recalc(i))
            hv.trace_add('write',  lambda *a, i=i: self.recalc(i))
            dv.trace_add('write',  lambda *a, i=i: self.recalc(i))
            hgv.trace_add('write', lambda *a, i=i: self.recalc(i))
            fxv.trace_add('write', lambda *a, i=i: self.recalc(i))

            self.name_vars.append(nv)
            self.width_vars.append(wv)
            self.height_vars.append(hv)
            self.door_vars.append(dv)
            self.hinge_vars.append(hgv)
            self.fix_vars.append(fxv)
            self.reg05_lbls.append(r05)
            self.barlen_lbls.append(rbl)
            self.reg08_lbls.append(r08)
            self.reg06_lbls.append(r06)
            self.dc_lbls.append(dc_lbl)
            self.fixpos_lbls.append(fix_lbl)
            self.r620_lbls.append(r620_lbl)

        # 버튼
        bf = tk.Frame(root, bg="white", pady=12)
        bf.pack(fill='x', padx=15)
        tk.Button(bf, text="🗑  전체 초기화", font=("맑은 고딕", 10),
                  bg="#E74C3C", fg="white", padx=12, pady=6, relief='flat',
                  command=self.clear_all).pack(side='left', padx=5)
        tk.Button(bf, text="🖨  인쇄", font=("맑은 고딕", 10),
                  bg="#2980B9", fg="white", padx=12, pady=6, relief='flat',
                  command=self.print_screen).pack(side='right', padx=5)
        tk.Button(bf, text="⚙  XML 생성", font=("맑은 고딕", 11, "bold"),
                  bg="#27AE60", fg="white", padx=20, pady=6, relief='flat',
                  command=self.run).pack(side='right', padx=5)

        # 상태바
        self.status_var = tk.StringVar(value="준비 완료 — 노란색 셀에 데이터를 입력하세요.")
        tk.Label(root, textvariable=self.status_var, font=("맑은 고딕", 9),
                 bg="#2C3E50", fg="#ECF0F1", anchor='w', padx=10, pady=4).pack(
                 fill='x', side='bottom')

    def recalc(self, i):
        try:
            w  = int(self.width_vars[i].get())
            h  = int(self.height_vars[i].get())
            reg05 = h - 30
            bl, _ = choose_reg05_bar(reg05)
            self.reg05_lbls[i].config(text=str(reg05), fg="#27AE60")
            self.barlen_lbls[i].config(text=f"{bl}mm" if bl else "불가",
                                        fg="#2980B9" if bl else "#E74C3C")
            self.reg08_lbls[i].config(text=str(w), fg="#27AE60")
            d = int(self.door_vars[i].get()) if self.door_vars[i].get().strip() else w
            hinge = self.hinge_vars[i].get()
            fix   = self.fix_vars[i].get()
            lk = calc_reg06(w, d, hinge, fix)
            self.reg06_lbls[i].config(text=str(int(lk['reg06'])), fg="#8E44AD")
            # DC 위치
            self.dc_lbls[i].config(text=str(lk['dc']), fg="#C0392B")
            # FIX 위치 (1개 또는 2개)
            if lk['fix1'] is not None and lk['fix2'] is not None:
                fix_txt = f"{lk['fix1']} / {lk['fix2']}"
            elif lk['fix1'] is not None:
                fix_txt = str(lk['fix1'])
            else:
                fix_txt = "없음"
            self.fixpos_lbls[i].config(text=fix_txt, fg="#E67E22")
            # REG-06-20 항상 0
            self.r620_lbls[i].config(text="0", fg="#7F8C8D")
        except ValueError:
            for lbl in [self.reg05_lbls[i], self.barlen_lbls[i],
                        self.reg08_lbls[i], self.reg06_lbls[i],
                        self.dc_lbls[i], self.fixpos_lbls[i], self.r620_lbls[i]]:
                lbl.config(text="-", fg="#555")

    def clear_all(self):
        if messagebox.askyesno("초기화", "모든 입력값을 삭제하시겠습니까?"):
            for i in range(self.MAX_ROWS):
                self.name_vars[i].set("")
                self.width_vars[i].set("")
                self.height_vars[i].set("")
                self.door_vars[i].set("")
                self.hinge_vars[i].set(HINGE_OPTIONS[0])
                self.fix_vars[i].set(FIX_OPTIONS[0])
            self.status_var.set("초기화 완료.")

    def print_screen(self):
        try:
            rows_data = []
            for i in range(self.MAX_ROWS):
                rows_data.append((
                    i+1,
                    self.name_vars[i].get().strip(),
                    self.width_vars[i].get().strip(),
                    self.height_vars[i].get().strip(),
                    self.door_vars[i].get().strip(),
                    self.hinge_vars[i].get(),
                    self.fix_vars[i].get(),
                    self.reg05_lbls[i].cget('text'),
                    self.barlen_lbls[i].cget('text'),
                    self.reg08_lbls[i].cget('text'),
                    self.reg06_lbls[i].cget('text'),
                ))
            today = datetime.date.today().strftime("%Y-%m-%d")
            rows_html = ""
            for no, name, w, h, door, hinge, fix, r05, bl, r08, r06 in rows_data:
                is_input = name or w or h
                bg = "white" if no % 2 == 1 else "#F8F9FA"
                def yc(v): return f'<td style="background:#FFFF00;text-align:center">{v}</td>'
                def rc(v, c="#555"): return f'<td style="color:{c};text-align:center">{v}</td>'
                if is_input:
                    rows_html += (f'<tr style="background:{bg}"><td style="text-align:center">{no}</td>'
                                  f'{yc(name)}{yc(w)}{yc(h)}{yc(door)}{yc(hinge)}{yc(fix)}'
                                  f'{rc(r05,"#27AE60")}{rc(bl,"#2980B9")}{rc(r08,"#27AE60")}{rc(r06,"#8E44AD")}'
                                  f'</tr>\n')
                else:
                    rows_html += (f'<tr style="background:{bg}"><td style="text-align:center">{no}</td>'
                                  + '<td style="background:#FFFF00"></td>'*6
                                  + '<td style="text-align:center">-</td>'*4
                                  + '</tr>\n')

            html = f"""<!DOCTYPE html>
<html><head><meta charset="utf-8">
<style>
  body{{font-family:'맑은 고딕',Arial,sans-serif;margin:20px}}
  .tb{{background:#2C3E50;color:white;text-align:center;padding:14px;border-radius:4px;margin-bottom:10px}}
  .tb h2{{margin:0;font-size:18px}} .tb p{{margin:3px 0 0;font-size:10px;color:#BDC3C7}}
  table{{width:100%;border-collapse:collapse;font-size:11px}}
  th{{background:#ECF0F1;color:#2C3E50;padding:6px 3px;text-align:center;border:1px solid #ddd}}
  td{{padding:4px 3px;border:1px solid #e0e0e0}}
  .st{{background:#2C3E50;color:#ECF0F1;padding:5px 10px;font-size:10px;margin-top:8px;border-radius:3px}}
  .dt{{text-align:right;font-size:10px;color:#888;margin-bottom:5px}}
</style>
</head><body>
<div class="dt">출력일: {today}</div>
<div class="tb"><h2>REGEN XML 생성기 <span style="font-size:11px;color:#95A5A6">{VERSION}</span></h2></div>
<table><thead><tr>
  <th>No.</th><th>업체명</th><th>폭W</th><th>높이H</th><th>문폭</th>
  <th>경첩방향</th><th>FIX</th><th>REG-05</th><th>원재료</th><th>REG-08</th><th>REG-06</th>
</tr></thead><tbody>{rows_html}</tbody></table>
<div class="st">{self.status_var.get()}</div>
</body></html>"""

            tmp = tempfile.NamedTemporaryFile(suffix='.html', delete=False, mode='w', encoding='utf-8')
            tmp.write(html); tmp_path = tmp.name; tmp.close()
            import webbrowser
            webbrowser.open(f'file:///{tmp_path.replace(os.sep, "/")}')
            self.status_var.set("🖨 브라우저가 열렸습니다 — Ctrl+P 로 인쇄하세요.")
        except Exception as e:
            messagebox.showerror("인쇄 오류", f"인쇄 중 오류:\n{e}")

    def run(self):
        rows, errors = [], []
        for i in range(self.MAX_ROWS):
            name  = self.name_vars[i].get().strip()
            w_str = self.width_vars[i].get().strip()
            h_str = self.height_vars[i].get().strip()
            d_str = self.door_vars[i].get().strip()
            hinge = self.hinge_vars[i].get()
            fix   = self.fix_vars[i].get()

            if not name and not w_str and not h_str and not d_str:
                continue
            if not name:
                errors.append(f"행 {i+1}: 업체명을 입력해주세요."); continue
            if not w_str:
                errors.append(f"행 {i+1} ({name}): 폭(W)을 입력해주세요."); continue
            if not h_str:
                errors.append(f"행 {i+1} ({name}): 높이(H)를 입력해주세요."); continue
            if not d_str:
                errors.append(f"행 {i+1} ({name}): 문폭을 입력해주세요."); continue
            try:
                w, h, d = int(w_str), int(h_str), int(d_str)
            except ValueError:
                errors.append(f"행 {i+1} ({name}): 폭/높이/문폭은 숫자여야 합니다."); continue
            if w <= 0 or h <= 0 or d <= 0:
                errors.append(f"행 {i+1} ({name}): 양수를 입력해주세요."); continue

            reg05 = h - 30
            bl, _ = choose_reg05_bar(reg05)
            if bl is None:
                errors.append(f"행 {i+1} ({name}): REG-05 길이({reg05}mm) 원재료 범위 초과"); continue

            lk = calc_reg06(w, d, hinge, fix)

            rows.append({'name': name, 'width': w, 'height': h, 'door': d,
                         'hinge': hinge, 'fix': fix, 'lk': lk})

        if errors:
            messagebox.showerror("입력 오류", "\n".join(errors)); return
        if not rows:
            messagebox.showwarning("입력 없음", "데이터를 1개 이상 입력해주세요."); return

        today = datetime.date.today()
        save_path = filedialog.asksaveasfilename(
            title="XML 저장 위치 선택", defaultextension=".XML",
            initialfile=f"regen_{today.strftime('%Y%m%d')}.XML",
            filetypes=[("XML 파일", "*.XML *.xml"), ("모든 파일", "*.*")])
        if not save_path:
            return

        try:
            generate_xml(rows, save_path)
            self.status_var.set(f"✅ 저장 완료: {save_path}  ({len(rows)}개 업체)")
            messagebox.showinfo("완료",
                f"XML 파일이 생성되었습니다!\n\n업체 수: {len(rows)}개\n저장 위치: {save_path}")
        except Exception as e:
            messagebox.showerror("오류", f"XML 생성 중 오류:\n{e}")
            self.status_var.set("❌ 오류 발생")


# ═══════════════════════════════════════════════════════
if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
