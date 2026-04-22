import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import tkinter.scrolledtext as st
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from mpl_toolkits.mplot3d.art3d import Poly3DCollection
import matplotlib.ticker as ticker
import numpy as np
import sys
import os
import random
from collections import Counter
import pandas as pd
import ast

# --- 1. フォント設定・カラーパレット ---
def get_jp_font_family():
    system = sys.platform
    if system == "win32": return "MS Gothic"
    elif system == "darwin": return "Hiragino Sans"
    return "DejaVu Sans"

FONT_FAMILY = get_jp_font_family()
plt.rcParams['font.family'] = FONT_FAMILY

class Colors:
    BG_MAIN = "#0B1221"
    BG_PANEL = "#152036"
    BG_CARD = "#1E2A45"
    ACCENT_MAIN = "#00F0FF"
    ACCENT_HOT = "#FF0099"
    TEXT_MAIN = "#E0E6ED"
    TEXT_DIM = "#94A3B8"
    SUCCESS = "#00FF9D"
    WARNING = "#FFB800"
    ERROR = "#FF4444"

class Fonts:
    HEADER = ("Meiryo", 12, "bold")
    BODY = ("Meiryo", 10)
    BODY_BOLD = ("Meiryo", 10, "bold")
    SMALL = ("Meiryo", 8)
    MONO = ("Consolas", 11)

# --- 2. 部品マスタ ---
def load_parts_master(file_path='parts_master.xlsx'):
    if not os.path.isabs(file_path):
        base_dir = os.path.dirname(os.path.abspath(__file__))
        file_path = os.path.join(base_dir, file_path)
    df = pd.read_excel(file_path)
    parts_master = {}
    for index, row in df.iterrows():
        part_id = row.iloc[0]
        offset_str = row.iloc[7]
        try:
            offset = ast.literal_eval(offset_str) if isinstance(offset_str, str) else (0,0,0)
        except:
            offset = (0,0,0)
        parts_master[part_id] = {
            'name': row.iloc[1],
            'w': int(row.iloc[2]),
            'd': int(row.iloc[3]),
            'h': int(row.iloc[4]),
            'weight': int(row.iloc[5]), # will be overwritten by random cache
            'color': row.iloc[6],
            'offset': offset
        }
    return parts_master

try:
    PARTS_MASTER = load_parts_master()
except Exception as e:
    print(f"Error loading parts master: {e}")
    raise e

# --- 3. データクラス ---
class Item:
    def __init__(self, item_id, master_data, unique_suffix, assigned_weight):
        self.id = f"{item_id}-{unique_suffix}"
        self.name = master_data['name']
        self.w = master_data['w']
        self.d = master_data['d']
        self.h = master_data['h']
        self.weight = assigned_weight 
        self.color = master_data['color']
        self.offset = master_data['offset']
        self.position = None
        self.abs_cog = None
        self.is_overweight = False
        self.is_spaceover = False

    def set_position(self, x, y, z):
        self.position = (x, y, z)
        self.abs_cog = (
            x + self.w/2 + self.offset[0],
            y + self.d/2 + self.offset[1],
            z + self.h/2 + self.offset[2]
        )

# --- 4. コンテナクラス ---
class Container:
    def __init__(self):
        self.w, self.d, self.h = 12000, 2300, 2400
        self.max_weight = 15000
        self.door_w, self.door_h = 2300, 2400
        self.items = []
        self.unloaded_items = []
        self.total_weight = 0
        self.candidate_x = {0}; self.candidate_y = {0}; self.candidate_z = {0}
        self.max_items_limit = 10000
        self.placed_matrix = np.zeros((self.max_items_limit, 6), dtype=np.float32)
        self.placed_count = 0

    def load_items(self, item_list):
        # 現場の経験と勘：高さだけでなく、同じ種類のアイテム（名前）と面サイズで固めて安定したブロックを作る
        item_list.sort(key=lambda x: (x.h, max(x.w, x.d), x.name, x.weight), reverse=True)
        self.items = []
        self.unloaded_items = []
        self.total_weight = 0
        self.candidate_x = {0}; self.candidate_y = {0}; self.candidate_z = {0}
        self.placed_count = 0 

        for item in item_list:
            if not self._try_load_single_item(item):
                self.unloaded_items.append(item)

    def _try_load_single_item(self, item):
        # 縦置き・横置きそれぞれの向き（オリエンテーション）のリストを作成
        orientations = []
        # パターン1: オリジナル (w, d)
        if not (item.d > self.door_w or item.w > self.w):
            orientations.append((item.w, item.d))
        # パターン2: 90度回転 (d, w)
        if not (item.w > self.door_w or item.d > self.w):
            if item.w != item.d:
                orientations.append((item.d, item.w))

        if not orientations:
            # どちらの向きでも物理的に扉を通らない
            item.is_spaceover = True
            item.set_position(self.w + 100, self.d/2 - item.d/2, 0)
            return False

        coord = None
        best_w, best_d = item.w, item.d
        sorted_z = sorted(list(self.candidate_z))
        sorted_x = sorted(list(self.candidate_x))
        sorted_y = sorted(list(self.candidate_y))
        
        # 物理的に配置可能な場所を探す（経験と勘：回転も試行して検証）
        for z in sorted_z:
            if z + item.h > self.h: continue
            for x in sorted_x:
                # 縦横両方の向きを試して空間効率が良さそうな方を採用する
                for w_trial, d_trial in orientations:
                    if x + w_trial > self.w: continue
                    for y in sorted_y:
                        if y + d_trial > self.d: continue
                        
                        original_w, original_d = item.w, item.d
                        item.w = w_trial
                        item.d = d_trial
                        
                        if self._can_place_physically(item, x, y, z):
                            coord = (x, y, z)
                            best_w, best_d = w_trial, d_trial
                            # 見つかったら一旦サイズを戻す
                            item.w, item.d = original_w, original_d
                            break
                        item.w, item.d = original_w, original_d
                        
                    if coord: break
                if coord: break
            if coord: break

        # コンテナ内に空き空間がない場合、天井に乗せてはみ出させる (Space Over)
        if not coord:
            item.is_spaceover = True
            item.set_position(self.w/2 - item.w/2, self.d/2 - item.d/2, self.h + 50)
            return False

        # 採用された回転向きを確定する
        item.w, item.d = best_w, best_d

        # スペースはあるが、重量オーバーの場合 (Weight Over)
        if self.total_weight + item.weight > self.max_weight:
            item.is_overweight = True
            item.set_position(coord[0], coord[1], coord[2])
            return False

        # 正常配置
        self._place_item(item, coord[0], coord[1], coord[2])
        return True

    def _place_item(self, item, x, y, z):
        item.set_position(x, y, z)
        self.items.append(item)
        
        if self.placed_count < self.max_items_limit:
            self.placed_matrix[self.placed_count] = [x, y, z, x + item.w, y + item.d, z + item.h]
            self.placed_count += 1
        else:
            new_row = np.array([[x, y, z, x + item.w, y + item.d, z + item.h]], dtype=np.float32)
            self.placed_matrix = np.vstack([self.placed_matrix, new_row])
            self.placed_count += 1
            self.max_items_limit += 1

        self.total_weight += item.weight
        if x + item.w < self.w: self.candidate_x.add(x + item.w)
        if y + item.d < self.d: self.candidate_y.add(y + item.d)
        if z + item.h < self.h: self.candidate_z.add(z + item.h)

    def _can_place_physically(self, item, x, y, z):
        if self.placed_count > 0:
            active_placed = self.placed_matrix[:self.placed_count]
            ix1, iy1, iz1 = x, y, z
            ix2, iy2, iz2 = x + item.w, y + item.d, z + item.h
            
            px1, py1, pz1 = active_placed[:, 0], active_placed[:, 1], active_placed[:, 2]
            px2, py2, pz2 = active_placed[:, 3], active_placed[:, 4], active_placed[:, 5]
            
            collision_mask = (
                (ix2 > px1) & (ix1 < px2) &
                (iy2 > py1) & (iy1 < py2) &
                (iz2 > pz1) & (iz1 < pz2)
            )
            if np.any(collision_mask): return False

        if z == 0: return True
        
        if self.placed_count > 0:
            active_placed = self.placed_matrix[:self.placed_count]
            pz2 = active_placed[:, 5]
            support_candidates_idx = np.abs(pz2 - z) < 1.0
            if not np.any(support_candidates_idx): return False 
                
            supports = active_placed[support_candidates_idx]
            sx1, sy1 = supports[:, 0], supports[:, 1]
            sx2, sy2 = supports[:, 3], supports[:, 4]
            inter_x1 = np.maximum(x, sx1)
            inter_y1 = np.maximum(y, sy1)
            inter_x2 = np.minimum(x + item.w, sx2)
            inter_y2 = np.minimum(y + item.d, sy2)
            w_overlap = np.maximum(0, inter_x2 - inter_x1)
            d_overlap = np.maximum(0, inter_y2 - inter_y1)
            
            supported_area = np.sum(w_overlap * d_overlap)
            base_area = item.w * item.d
            if supported_area / base_area < 0.99: return False
            return True
        return False
    
    def get_loadable_counts(self, master_data):
        counts = {}
        saved_items_len = len(self.items)
        saved_weight = self.total_weight
        saved_cx = self.candidate_x.copy()
        saved_cy = self.candidate_y.copy()
        saved_cz = self.candidate_z.copy()
        saved_count = self.placed_count

        for key, val in master_data.items():
            self.items = self.items[:saved_items_len]
            self.total_weight = saved_weight
            self.candidate_x = saved_cx.copy()
            self.candidate_y = saved_cy.copy()
            self.candidate_z = saved_cz.copy()
            self.placed_count = saved_count
            
            count = 0
            while True:
                if count >= 50: break
                if self.total_weight + val['weight'] > self.max_weight: break
                
                temp_item = Item(key, val, f"trial-{count}", 1000) # dummy weight for trial calculation geometry
                temp_item.weight = val['weight'] # real average weight constraint? For suggestions, use max weight or assume average? Let's use 10,000kg.
                temp_item.weight = 10000 
                if self._try_load_single_item(temp_item):
                    count += 1
                else: break
            
            if count > 0: counts[key] = count
                
        self.items = self.items[:saved_items_len]
        self.total_weight = saved_weight
        self.candidate_x = saved_cx
        self.candidate_y = saved_cy
        self.candidate_z = saved_cz
        self.placed_count = saved_count
        return counts

    def get_cog_stats(self):
        if self.total_weight == 0: return (0,0,0), (0,0)
        cx = sum(i.weight * i.abs_cog[0] for i in self.items) / self.total_weight
        cy = sum(i.weight * i.abs_cog[1] for i in self.items) / self.total_weight
        cz = sum(i.weight * i.abs_cog[2] for i in self.items) / self.total_weight
        diff_x = (cx - self.w/2) / self.w * 100
        diff_y = (cy - self.d/2) / self.d * 100
        return (cx, cy, cz), (diff_x, diff_y)

# --- 5. GUIアプリ ---
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Vanning Optimizer v5.0 [NEO]")
        self.root.geometry("1400x900")
        self.root.configure(bg=Colors.BG_MAIN)
        
        self.container = None; self.fig = None; self.ax = None
        self.qty_vars = {}
        
        # [NEW] 重量の固定キャッシュ
        self.weight_cache = {} 
        self.log_messages = []
        self.last_quantities = {}

        self.left_frame = tk.Frame(root, width=480, bg=Colors.BG_PANEL)
        self.left_frame.pack(side=tk.LEFT, fill=tk.Y)
        self.left_frame.pack_propagate(False)
        
        self.right_frame = tk.Frame(root, bg=Colors.BG_MAIN)
        self.right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        self.canvas_frame = tk.Frame(self.right_frame, bg=Colors.BG_MAIN)
        self.canvas_frame.pack(fill=tk.BOTH, expand=True)

        self._create_left_panel()
        self._create_hud_controls()
        self.run_simulation()

    def _create_left_panel(self):
        # ヘッダー
        header_frame = tk.Frame(self.left_frame, bg=Colors.BG_PANEL, pady=10, padx=20)
        header_frame.pack(fill=tk.X)
        tk.Label(header_frame, text="VANNING SYSTEM", bg=Colors.BG_PANEL, fg=Colors.ACCENT_MAIN, 
                 font=("Meiryo", 16, "bold"), anchor="w").pack(fill=tk.X)
        tk.Label(header_frame, text="40ft Container Simulator", fg=Colors.TEXT_DIM, bg=Colors.BG_PANEL).pack(anchor=tk.W)

        # [NEW] 総重量ゲージ
        gauge_frame = tk.Frame(self.left_frame, bg=Colors.BG_PANEL, padx=20, pady=5)
        gauge_frame.pack(fill=tk.X)
        self.lbl_weight = tk.Label(gauge_frame, text="総重量: 0 / 26,500 kg (0%)", bg=Colors.BG_PANEL, fg="white", font=Fonts.BODY_BOLD)
        self.lbl_weight.pack(anchor="w")
        
        style = ttk.Style()
        style.theme_use('default')
        style.configure("Cyan.Horizontal.TProgressbar", background=Colors.ACCENT_MAIN, troughcolor="#222222")
        self.weight_progress = ttk.Progressbar(gauge_frame, orient="horizontal", length=440, mode="determinate", style="Cyan.Horizontal.TProgressbar")
        self.weight_progress.pack(fill=tk.X, pady=(5,0))

        # 年間予定選択エリア
        yearly_frame = tk.Frame(self.left_frame, bg=Colors.BG_PANEL, padx=20, pady=10)
        yearly_frame.pack(side=tk.TOP, fill=tk.X)
        
        tk.Label(yearly_frame, text="📅 年間シミュレーション設定", bg=Colors.BG_PANEL, fg=Colors.ACCENT_MAIN, 
                 font=Fonts.BODY_BOLD).pack(anchor=tk.W)
        
        sel_row = tk.Frame(yearly_frame, bg=Colors.BG_PANEL)
        sel_row.pack(fill=tk.X, pady=5)
        
        tk.Label(sel_row, text="対象月:", bg=Colors.BG_PANEL, fg="white", font=Fonts.SMALL).pack(side=tk.LEFT)
        self.month_combo = ttk.Combobox(sel_row, values=[f"{i}月" for i in range(1, 13)], width=10, state="readonly")
        self.month_combo.set("1月")
        self.month_combo.pack(side=tk.LEFT, padx=5)
        self.month_combo.bind("<<ComboboxSelected>>", self.on_month_selected)
        
        self.btn_yearly = tk.Button(sel_row, text="年間予定表を読込", bg="#224433", fg="white",
                                    font=("Meiryo", 8, "bold"), command=self.load_yearly_layout, cursor="hand2")
        self.btn_yearly.pack(side=tk.LEFT, padx=5)

        tk.Label(self.left_frame, text="PARTS SELECTION (読込結果)", bg=Colors.BG_PANEL, fg=Colors.TEXT_MAIN, 
                 font=Fonts.BODY_BOLD).pack(anchor=tk.W, padx=20, pady=(5, 5))

        # スクロールエリア
        canvas_frame = tk.Frame(self.left_frame, bg=Colors.BG_PANEL, padx=20)
        canvas_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(canvas_frame, bg=Colors.BG_PANEL, highlightthickness=0)
        scrollbar = tk.Scrollbar(canvas_frame, orient="vertical", command=canvas.yview)
        self.scrollable_frame = tk.Frame(canvas, bg=Colors.BG_PANEL)
        
        self.scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw", width=420) 
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        for i, (key, data) in enumerate(PARTS_MASTER.items()):
            # [NEW] Slider UI instead of Spinbox
            row_frame = tk.Frame(self.scrollable_frame, bg=Colors.BG_CARD, pady=5, padx=5)
            row_frame.pack(fill=tk.X, pady=3)
            
            color_box = tk.Label(row_frame, bg=data['color'], width=1); color_box.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 5))
            info_frame = tk.Frame(row_frame, bg=Colors.BG_CARD); info_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            tk.Label(info_frame, text=data['name'], bg=Colors.BG_CARD, fg="white", font=("Meiryo", 10, "bold"), anchor="w").pack(fill=tk.X)
            
            # [NEW] 寸法情報の追加表示（重量は外部ファイル依存）
            dim_text = f"サイズ: L{data['w']} x W{data['d']} x H{data['h']} / 重量: 外部ファイル参照"
            tk.Label(info_frame, text=dim_text, bg=Colors.BG_CARD, fg=Colors.TEXT_DIM, font=("Meiryo", 8), anchor="w").pack(fill=tk.X)
            
            var = tk.IntVar(value=0)
            self.qty_vars[key] = var
            self.last_quantities[key] = 0
            
            # 操作用ボタンコンテナ
            ctrl_frame = tk.Frame(info_frame, bg=Colors.BG_CARD)
            ctrl_frame.pack(side=tk.RIGHT, padx=5)
            
            # マイナスボタン
            btn_minus = tk.Button(ctrl_frame, text="－", font=("Meiryo", 9, "bold"), width=2,
                                  bg=Colors.BG_PANEL, fg=Colors.TEXT_MAIN, relief="flat", cursor="hand2",
                                  command=lambda k=key: self.on_slider_change(k, str(self.qty_vars[k].get() - 1)))
            btn_minus.pack(side=tk.LEFT, padx=2)
            
            # 数量ラベル
            val_lbl = tk.Label(ctrl_frame, textvariable=var, bg=Colors.BG_CARD, fg=Colors.ACCENT_HOT, font=("Meiryo", 12, "bold"), width=3)
            val_lbl.pack(side=tk.LEFT, padx=5)
            
            # プラスボタン
            btn_plus = tk.Button(ctrl_frame, text="＋", font=("Meiryo", 9, "bold"), width=2,
                                 bg=Colors.BG_PANEL, fg=Colors.TEXT_MAIN, relief="flat", cursor="hand2",
                                 command=lambda k=key: self.on_slider_change(k, str(self.qty_vars[k].get() + 1), weight=data['weight']))
            btn_plus.pack(side=tk.LEFT, padx=2)

        # 追加ログテキスト
        bottom_frame = tk.Frame(self.left_frame, bg="#111111", padx=20, pady=10)
        bottom_frame.pack(side=tk.BOTTOM, fill=tk.X)
        
        btn_run = tk.Button(bottom_frame, text="▶ バンニング一括実行", bg=Colors.ACCENT_MAIN, fg="black",
                            font=("Meiryo", 14, "bold"), command=self.run_simulation, cursor="hand2")
        btn_run.pack(fill=tk.X, pady=(0, 5))

        btn_csv = tk.Button(bottom_frame, text="📁 マニフェスト読込 (CSV/Excel)", bg="#334466", fg="white",
                              font=("Meiryo", 10, "bold"), command=self.load_manifest_file, cursor="hand2")
        btn_csv.pack(fill=tk.X, pady=(0, 10))

        tk.Label(bottom_frame, text="OPERATION HISTORY", bg="#111111", fg=Colors.TEXT_DIM, font=Fonts.SMALL).pack(anchor="w")
        self.log_text = st.ScrolledText(bottom_frame, width=30, height=5, font=("MS Gothic", 9), 
                                           bg="#050505", fg="white", borderwidth=0, highlightthickness=1, highlightbackground=Colors.ACCENT_MAIN)
        self.log_text.pack(fill=tk.X, pady=(2, 0))

    def _create_hud_controls(self):
        hud_frame = tk.Frame(self.right_frame, bg=Colors.BG_MAIN)
        hud_frame.place(relx=0.98, rely=0.98, anchor="se")
        
        btn_base_style = {"bg": "black", "fg": "#00FFFF", "activebackground": "#003333", "activeforeground": "white",
                          "font": ("Meiryo", 9, "bold"), "relief": "flat", "bd": 1, "highlightthickness": 0}
        
        def mk_hud_btn(parent, text, cmd, bg_color="black", fg_color="#00FFFF"):
            frm = tk.Frame(parent, bg="#00FFFF", padx=1, pady=1)
            frm.pack(side=tk.LEFT, padx=4)
            btn = tk.Button(frm, text=text, command=cmd, **btn_base_style)
            btn.configure(bg=bg_color, fg=fg_color)
            btn.pack(fill=tk.BOTH)
            return btn
        
        mk_hud_btn(hud_frame, " CLEAR ", self.clear_all_items, bg_color="#220000", fg_color="#FF3333")
        mk_hud_btn(hud_frame, " ⟲ ", lambda: self.rotate_view(10))
        mk_hud_btn(hud_frame, " ⟳ ", lambda: self.rotate_view(-10))

    def append_log(self, text, color="white"):
        self.log_text.insert(tk.END, text + "\n")
        self.log_text.see(tk.END)

    def on_slider_change(self, key, val_str, weight=0):
        # 0〜40の範囲制限
        try:
            new_qty = max(0, min(40, int(float(val_str))))
        except:
            return
            
        self.qty_vars[key].set(new_qty) # ボタン操作時用に明示的にセット
        old_qty = self.last_quantities[key]
        
        if new_qty > old_qty:
            diff = new_qty - old_qty
            for _ in range(diff):
                # 外部ファイルまたはマスタの確定重量を使用
                w = int(weight)
                if key not in self.weight_cache: self.weight_cache[key] = []
                self.weight_cache[key].append(w)
                self.append_log(f"📝 {PARTS_MASTER[key]['name']} をリスト追加 ({w:,}kg)")
                
        elif new_qty < old_qty:
            diff = old_qty - new_qty
            for _ in range(diff):
                if key in self.weight_cache and len(self.weight_cache[key]) > 0:
                    w = self.weight_cache[key].pop()
                    self.append_log(f"❌ {PARTS_MASTER[key]['name']} をリスト削除 (-{w:,}kg)")
                    
        self.last_quantities[key] = new_qty

    def load_manifest_file(self):
        file_path = filedialog.askopenfilename(
            title="予定リストを選択",
            filetypes=[("Excel & CSV files", "*.xlsx *.xls *.csv")]
        )
        if not file_path: return
        
        try:
            if file_path.endswith('.csv'):
                try:
                    df = pd.read_csv(file_path, encoding='utf-8')
                except Exception:
                    df = pd.read_csv(file_path, encoding='shift_jis')
            else:
                df = pd.read_excel(file_path)
        except Exception as e:
            messagebox.showerror("エラー", f"ファイルの読み込みに失敗しました:\n{e}")
            return
            
        weight_col = next((col for col in df.columns if "重量" in str(col) or "重さ" in str(col) or "Weight" in str(col)), None)
        if not weight_col:
            messagebox.showerror("エラー", "ファイル内に「重量」列が見つかりません。今回の仕様では各ケースの重量指定が必須です。")
            return
            
        name_col = next((col for col in df.columns if "名称" in str(col) or "名前" in str(col) or "品名" in str(col) or "資材" in str(col) or "Name" in str(col)), None)
        if not name_col:
            name_col = next((col for col in df.columns if df[col].dtype == object), None)
            
        if not name_col:
            messagebox.showerror("エラー", "部品名（資材名称）が記載された列が見つかりません。")
            return
            
        self.clear_all_items(run_sim=False)
        self.append_log(f"📁 ファイルからマニフェスト（重量確定済み）を読み込んでいます...")
        
        name_to_key = {v['name'].replace(" ", "").replace("　",""): k for k, v in PARTS_MASTER.items()}
        loaded_count = 0
        unknown_parts = []
        
        qty_col = next((col for col in df.columns if "数量" in str(col) or "個数" in str(col) or "数" in str(col)), None)

        for _, row in df.iterrows():
            part_name = str(row[name_col]).strip()
            clean_name = part_name.replace(" ", "").replace("　","")
            
            if clean_name in name_to_key:
                matched_key = name_to_key[clean_name]
                w = 0
                try:
                    w = int(row[weight_col])
                except:
                    continue
                
                qty = 1
                if qty_col and not pd.isna(row[qty_col]):
                    try:
                        qty = int(row[qty_col])
                    except:
                        qty = 1
                
                for _ in range(qty):
                    cur = self.qty_vars[matched_key].get()
                    self.qty_vars[matched_key].set(cur + 1)
                    self.on_slider_change(matched_key, str(cur + 1), weight=w)
                    loaded_count += 1
            else:
                unknown_parts.append(part_name)
                
        if unknown_parts:
            self.append_log(f"⚠️ マスタにない部品を無視しました: {', '.join(set(unknown_parts))}", "yellow")
            
        self.append_log(f"✅ 合計 {loaded_count} 個のケース（実重量）を読み込みました！", "green")
        self.run_simulation()

    def load_yearly_layout(self):
        base_dir = os.path.dirname(os.path.abspath(__file__))
        file_path = os.path.join(base_dir, "vanning_layout_2026.xlsx")
        
        if not os.path.exists(file_path):
            messagebox.showerror("エラー", f"'{file_path}' が見つかりません。\nパス: {file_path}")
            return
        
        self.yearly_xl = pd.ExcelFile(file_path)
        self.append_log(f"📅 年間予定表を読み込みました。")
        self.on_month_selected(None)

    def on_month_selected(self, event):
        if not hasattr(self, 'yearly_xl'): 
            return
            
        month_str = self.month_combo.get()
        sheet_name = month_str
        if sheet_name not in self.yearly_xl.sheet_names:
            sheet_name = month_str.replace("月", "")
            if sheet_name not in self.yearly_xl.sheet_names:
                messagebox.showerror("エラー", f"シート '{month_str}' が見つかりません。")
                return
        
        df = pd.read_excel(self.yearly_xl, sheet_name=sheet_name, header=None)
        
        self.clear_all_items(run_sim=False)
        self.append_log(f"🔍 {month_str} の全データを抽出中...")
        
        loaded_count = 0
        name_to_key = {v['name'].replace(" ", "").replace("　",""): k for k, v in PARTS_MASTER.items()}
        
        for i, row in df.iterrows():
            if pd.isna(row[0]) or "Container" in str(row[0]) or "種別" in str(row[0]):
                continue
            
            # IDによる紐付け (Excelの2列目がID 1, 2, ... 31 となっている)
            matched_key = None
            try:
                raw_id = int(row[1])
                test_key = f"CASE_{raw_id:02d}"
                if test_key in PARTS_MASTER:
                    matched_key = test_key
            except:
                pass
            
            # IDで見つからなかった場合は名前で探す
            if not matched_key:
                part_name = str(row[2]).strip()
                clean_name = part_name.replace(" ", "").replace("　","")
                if clean_name in name_to_key:
                    matched_key = name_to_key[clean_name]
                else:
                    for m_name, m_key in name_to_key.items():
                        if clean_name in m_name or m_name in clean_name:
                            matched_key = m_key
                            break
            
            if matched_key:
                try:
                    w = int(row[6])
                    cur = self.qty_vars[matched_key].get()
                    self.qty_vars[matched_key].set(cur + 1)
                    self.on_slider_change(matched_key, str(cur + 1), weight=w)
                    loaded_count += 1
                except:
                    continue
                    
        self.append_log(f"✅ {month_str} から計 {loaded_count} 個の荷物を抽出しました。")
        if loaded_count > 0:
            self.append_log("💡 集約シミュレーションを開始します...")
            self.run_simulation()
        else:
            self.append_log("⚠️ 読み込める荷物が見つかりませんでした。", "yellow")

    def clear_all_items(self, run_sim=True):
        for key, var in self.qty_vars.items():
            var.set(0)
            self.last_quantities[key] = 0
            self.weight_cache[key] = []
        self.append_log("🔄 リストをクリアしました", "yellow")
        if run_sim:
            self.run_simulation()

    def run_simulation(self):
        items_to_load = []
        for key, var in self.qty_vars.items():
            qty = var.get()
            if qty > 0:
                master = PARTS_MASTER[key]
                weights = self.weight_cache.get(key, [])
                for i in range(qty):
                    # 安全対策
                    assigned_weight = weights[i] if i < len(weights) else random.randint(1000, 15000)
                    items_to_load.append(Item(key, master, i, assigned_weight))
        
        self.container = Container()
        self.container.load_items(items_to_load)
        cog, devs = self.container.get_cog_stats()

        # Update gauge
        tot_w = self.container.total_weight
        mx_w = self.container.max_weight
        pct_w = (tot_w / mx_w) * 100
        
        # [NEW] 容積充填率の計算
        tot_v = sum(item.w * item.d * item.h for item in self.container.items)
        mx_v = self.container.w * self.container.d * self.container.h
        pct_v = (tot_v / mx_v) * 100
        
        self.lbl_weight.config(text=f"重量: {tot_w:,}/{mx_w:,}kg ({pct_w:.1f}%) | 容積: {pct_v:.1f}%")
        self.weight_progress['value'] = pct_w
        
        if pct_w > 100: self.lbl_weight.config(fg=Colors.ERROR)
        else: self.lbl_weight.config(fg="white")

        self.draw_3d_graph(cog, devs)

    def draw_3d_graph(self, cog, devs):
        if self.fig: plt.close(self.fig); 
        for w in self.canvas_frame.winfo_children(): w.destroy()
        
        plt.style.use('dark_background')
        self.fig = plt.figure(figsize=(8, 6), dpi=100)
        self.fig.patch.set_facecolor(Colors.BG_MAIN)
        self.ax = self.fig.add_subplot(111, projection='3d')
        self.ax.set_facecolor(Colors.BG_MAIN)
        self.ax.set_title("Vanning Optimizer: Layout View", fontsize=14, color=Colors.ACCENT_MAIN)
        c = self.container
        self.ax.set_xlim([0, c.w]); self.ax.set_ylim([0, c.d]); self.ax.set_zlim([0, c.h])
        self.ax.set_box_aspect((c.w, c.d, c.h))

        # コンテナ境界線
        edges = [([0, c.w, c.w, 0, 0], [0, 0, c.d, c.d, 0], [0, 0, 0, 0, 0]), ([0, c.w, c.w, 0, 0], [0, 0, c.d, c.d, 0], [c.h]*5)]
        for x in [0, c.w]:
            for y in [0, c.d]: edges.append(([x, x], [y, y], [0, c.h]))
        for xs, ys, zs in edges: 
            self.ax.plot(xs, ys, zs, color=Colors.ACCENT_MAIN, lw=1.5, alpha=0.3, zorder=0)

        def draw_box(item, is_error=False):
            x, y, z = item.position; dx, dy, dz = item.w, item.d, item.h
            xx = [x, x+dx, x+dx, x, x, x+dx, x+dx, x]; yy = [y, y, y+dy, y+dy, y, y, y+dy, y+dy]; zz = [z, z, z, z, z+dz, z+dz, z+dz, z+dz]
            verts = [[(xx[i], yy[i], zz[i]) for i in [0, 1, 5, 4]], [(xx[i], yy[i], zz[i]) for i in [7, 6, 2, 3]],
                     [(xx[i], yy[i], zz[i]) for i in [0, 3, 7, 4]], [(xx[i], yy[i], zz[i]) for i in [1, 2, 6, 5]],
                     [(xx[i], yy[i], zz[i]) for i in [0, 1, 2, 3]], [(xx[i], yy[i], zz[i]) for i in [4, 5, 6, 7]]]
            
            if is_error:
                poly = Poly3DCollection(verts, facecolors=Colors.ERROR, linewidths=1.0, edgecolors='white', alpha=0.6, zorder=2)
            else:
                poly = Poly3DCollection(verts, facecolors=item.color, linewidths=0.5, edgecolors='black', alpha=0.9, zorder=1)
                
            poly._item_info = f"{item.name}\n重量: {item.weight:,}kg\nサイズ: W{item.w} D{item.d} H{item.h}"
            poly.set_picker(True)
            self.ax.add_collection3d(poly)

        # 正常に積載されたアイテム
        for item in c.items:
            draw_box(item, False)
            
        # エラーではみ出た・重量オーバーしたアイテムの描画は非表示にする
        # for e_item in c.unloaded_items:
        #     draw_box(e_item, True)

        # [NEW] 重心ボール表示
        if c.total_weight > 0:
            self.ax.scatter([cog[0]], [cog[1]], [cog[2]], color=Colors.ACCENT_HOT, s=400, marker='o', 
                            edgecolors='white', linewidths=2, label="重心\nCOG", zorder=100)
        
        # [NEW] 目盛りの間引き
        self.ax.yaxis.set_major_locator(ticker.MaxNLocator(5))
        self.ax.xaxis.set_major_locator(ticker.MaxNLocator(8))
        self.ax.zaxis.set_major_locator(ticker.MaxNLocator(5))

        self.ax.grid(False)
        self.ax.set_xlabel('L（奥/手前）', color=Colors.TEXT_DIM); self.ax.set_ylabel('W（横幅）', color=Colors.TEXT_DIM); self.ax.set_zlabel('H（高さ）', color=Colors.TEXT_DIM)
        self.ax.tick_params(colors=Colors.TEXT_DIM)
        self.ax.xaxis.pane.fill = False; self.ax.yaxis.pane.fill = False; self.ax.zaxis.pane.fill = False
        
        self.fig.canvas.mpl_connect('pick_event', self.on_pick)
        
        self.ax.legend(loc='upper right', facecolor=Colors.BG_CARD, edgecolor=Colors.BG_PANEL, labelcolor=Colors.TEXT_MAIN)
        canvas = FigureCanvasTkAgg(self.fig, master=self.canvas_frame); canvas.draw()
        canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)
    
    def on_pick(self, event):
        info = getattr(event.artist, '_item_info', None)
        if info:
            messagebox.showinfo("貨物詳細", info)

    def rotate_view(self, angle):
        if self.ax: self.ax.azim += angle; self.fig.canvas.draw_idle()

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()