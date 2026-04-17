import tkinter as tk
from tkinter import ttk, messagebox
import tkinter.scrolledtext as st
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from mpl_toolkits.mplot3d.art3d import Poly3DCollection
import numpy as np
import sys
import os
import random
from collections import Counter # 集計用

# --- 1. フォント設定・カラーパレット ---
def get_jp_font_family():
    system = sys.platform
    if system == "win32": return "MS Gothic"
    elif system == "darwin": return "Hiragino Sans"
    return "DejaVu Sans"

import pandas as pd
import ast

FONT_FAMILY = get_jp_font_family()
plt.rcParams['font.family'] = FONT_FAMILY

# デザイン定数 (近未来 / サイバーパンクテーマ)
class Colors:
    BG_MAIN = "#0B1221"      # Deep Dark Blue/Black
    BG_PANEL = "#152036"     # Panel Background
    BG_CARD = "#1E2A45"      # Item Card Background
    ACCENT_MAIN = "#00F0FF"  # Cyber Cyan
    ACCENT_HOT = "#FF0099"   # Neon Pink
    TEXT_MAIN = "#E0E6ED"    # White-ish
    TEXT_DIM = "#94A3B8"     # Gray
    SUCCESS = "#00FF9D"      # Neon Green
    WARNING = "#FFB800"      # Amber
    ERROR = "#FF4444"        # Red

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
        # Excelファイルのカラムマッピング確認済み
        # 0: ID, 1: Name, 2: W, 3: D, 4: H, 5: Weight, 6: Color, 7: Offset
        part_id = row.iloc[0]
        
        # オフセット文字列 "(x,y,z)" をタプルに変換
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
            'weight': int(row.iloc[5]),
            'color': row.iloc[6],
            'offset': offset
        }
    return parts_master

try:
    PARTS_MASTER = load_parts_master()
except Exception as e:
    print(f"Error loading parts master: {e}")
    # Fallback or exit? For now let's just use an empty dict or re-raise
    # Re-raising to make sure the user knows something went wrong
    raise e


# --- 3. データクラス ---
class Item:
    def __init__(self, item_id, master_data, unique_suffix):
        self.id = f"{item_id}-{unique_suffix}"
        self.name = master_data['name']
        self.w = master_data['w']
        self.d = master_data['d']
        self.h = master_data['h']
        self.weight = random.randint(1000, 15000) # [NEW] 1,000kg〜15,000kgのランダム重量
        self.color = master_data['color']
        self.offset = master_data['offset']
        self.position = None
        self.abs_cog = None

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
        # 40ft専用 (提供された実寸法に合わせる)
        self.w, self.d, self.h = 12000, 2300, 2400
        self.max_weight = 26500
        self.door_w, self.door_h = 2300, 2400 # 庫内寸法とぴったり同じに設定して制限を開放
        self.items = []
        self.unloaded_items = []
        self.total_weight = 0
        self.candidate_x = {0}; self.candidate_y = {0}; self.candidate_z = {0}
        # [NEW] Pre-allocate large array to avoid resize overhead
        # Max 10000 items is sufficient for this simulation scale
        self.max_items_limit = 10000
        self.placed_matrix = np.zeros((self.max_items_limit, 6), dtype=np.float32)
        self.placed_count = 0

    def load_items(self, item_list):
        item_list.sort(key=lambda x: (x.h, x.d, x.w, x.weight), reverse=True)
        self.items = []
        self.unloaded_items = []
        self.total_weight = 0
        self.candidate_x = {0}; self.candidate_y = {0}; self.candidate_z = {0}
        self.placed_count = 0 
        # Note: placed_matrix contents don't need clearing, just reset count

        for item in item_list:
            if not self._try_load_single_item(item):
                self.unloaded_items.append(item)

    def _try_load_single_item(self, item):
        # 扉開口部のチェック (幅と高さが扉サイズ以下であること)
        if item.d > self.door_w or item.h > self.door_h:
            return False

        # 単品積載試行ロジック (成功したらTrue)
        sorted_z = sorted(list(self.candidate_z))
        sorted_x = sorted(list(self.candidate_x))
        sorted_y = sorted(list(self.candidate_y))
        
        for z in sorted_z:
            if z + item.h > self.h: continue
            for x in sorted_x:
                if x + item.w > self.w: continue
                for y in sorted_y:
                    if y + item.d > self.d: continue
                    if self._can_place_physically(item, x, y, z):
                        self._place_item(item, x, y, z)
                        return True
        return False

    def _place_item(self, item, x, y, z):
        item.set_position(x, y, z)
        self.items.append(item)
        
        # NumPy配列へ登録 (Pre-allocated)
        if self.placed_count < self.max_items_limit:
            self.placed_matrix[self.placed_count] = [x, y, z, x + item.w, y + item.d, z + item.h]
            self.placed_count += 1
        else:
            # Fallback or resize if exceeded (simplified: ignore or just append slowly?)
            # For safety, let's resize
            new_row = np.array([[x, y, z, x + item.w, y + item.d, z + item.h]], dtype=np.float32)
            self.placed_matrix = np.vstack([self.placed_matrix, new_row])
            self.placed_count += 1
            self.max_items_limit += 1

        self.total_weight += item.weight
        if x + item.w < self.w: self.candidate_x.add(x + item.w)
        if y + item.d < self.d: self.candidate_y.add(y + item.d)
        if z + item.h < self.h: self.candidate_z.add(z + item.h)

    def _can_place_physically(self, item, x, y, z):
        # 0. 重量チェック
        if self.total_weight + item.weight > self.max_weight: return False
        
        if self.placed_count > 0:
            # Active checking area
            active_placed = self.placed_matrix[:self.placed_count]
            
            # 1. 干渉チェック (ベクトル化)
            ix1, iy1, iz1 = x, y, z
            ix2, iy2, iz2 = x + item.w, y + item.d, z + item.h
            
            px1, py1, pz1 = active_placed[:, 0], active_placed[:, 1], active_placed[:, 2]
            px2, py2, pz2 = active_placed[:, 3], active_placed[:, 4], active_placed[:, 5]
            
            collision_mask = (
                (ix2 > px1) & (ix1 < px2) &
                (iy2 > py1) & (iy1 < py2) &
                (iz2 > pz1) & (iz1 < pz2)
            )
            if np.any(collision_mask):
                return False

        if z == 0: return True
        
        # 2. 支持チェック (ベクトル化)
        if self.placed_count > 0:
            active_placed = self.placed_matrix[:self.placed_count]
            
            pz2 = active_placed[:, 5]
            support_candidates_idx = np.abs(pz2 - z) < 1.0
            
            if not np.any(support_candidates_idx):
                return False 
                
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
            
            if supported_area / base_area < 0.99: return False # [CHANGED] 80% -> 99% (Effectively 100%)
            return True
            
        return False
    
    # 【新機能】今の状態であと何が積めるかチェックする
    def get_loadable_counts(self, master_data):
        counts = {}
        # 現状のバックアップ
        saved_items_len = len(self.items)
        saved_weight = self.total_weight
        saved_cx = self.candidate_x.copy()
        saved_cy = self.candidate_y.copy()
        saved_cz = self.candidate_z.copy()
        saved_count = self.placed_count # Only need to save the count!

        for key, val in master_data.items():
            # 状態復元
            self.items = self.items[:saved_items_len]
            self.total_weight = saved_weight
            self.candidate_x = saved_cx.copy()
            self.candidate_y = saved_cy.copy()
            self.candidate_z = saved_cz.copy()
            self.placed_count = saved_count # Restore count (data beyond this is ignored)
            
            count = 0
            while True:
                if count >= 50: break # 安全のため上限
                if self.total_weight + val['weight'] > self.max_weight: break
                
                # 仮想アイテム生成と積載試行
                temp_item = Item(key, val, f"trial-{count}")
                if self._try_load_single_item(temp_item):
                    count += 1
                else:
                    break
            
            if count > 0:
                counts[key] = count
                
        # 最終的な復元
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
        self.root.title("Vanning Optimizer v4.1 [NEO]")
        self.root.geometry("1400x900")
        self.root.configure(bg=Colors.BG_MAIN)
        
        self.container = None; self.fig = None; self.ax = None
        self.qty_vars = {}
        self.suggestion_timer = None # For debouncing 
        
        # 左側: 操作パネル
        self.left_frame = tk.Frame(root, width=480, bg=Colors.BG_PANEL)
        self.left_frame.pack(side=tk.LEFT, fill=tk.Y)
        self.left_frame.pack_propagate(False)
        
        # 右側: 可視化エリア
        self.right_frame = tk.Frame(root, bg=Colors.BG_MAIN)
        self.right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        self.canvas_frame = tk.Frame(self.right_frame, bg=Colors.BG_MAIN)
        self.canvas_frame.pack(fill=tk.BOTH, expand=True)

        self._create_left_panel()
        self._create_hud_controls()
        self.run_simulation()

    def _create_left_panel(self):
        # ヘッダーエリア
        header_frame = tk.Frame(self.left_frame, bg=Colors.BG_PANEL, pady=20, padx=20)
        header_frame.pack(fill=tk.X)
        tk.Label(header_frame, text="VANNING SYSTEM", bg=Colors.BG_PANEL, fg=Colors.ACCENT_MAIN, 
                 font=("Meiryo", 16, "bold"), anchor="w").pack(fill=tk.X)
        tk.Label(header_frame, text="Optimization & Loading", bg=Colors.BG_PANEL, fg=Colors.TEXT_DIM, 
                 font=Fonts.SMALL, anchor="w").pack(fill=tk.X)

        # コンテナ選択
        ctrl_frame = tk.Frame(self.left_frame, bg=Colors.BG_PANEL, padx=20)
        ctrl_frame.pack(fill=tk.X)
        
        tk.Label(ctrl_frame, text="CONTAINER TYPE: 40ft (Max 26.5t)", bg=Colors.BG_PANEL, fg=Colors.ACCENT_MAIN, 
                 font=Fonts.BODY_BOLD).pack(anchor=tk.W, pady=(10, 5))

        # 部品リストヘッダー
        tk.Label(ctrl_frame, text="PARTS SELECTION", bg=Colors.BG_PANEL, fg=Colors.TEXT_MAIN, 
                 font=Fonts.BODY_BOLD).pack(anchor=tk.W, pady=(20, 5))

        # 部品用スクロールエリア
        canvas_frame = tk.Frame(self.left_frame, bg=Colors.BG_PANEL, padx=20)
        canvas_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(canvas_frame, bg=Colors.BG_PANEL, highlightthickness=0)
        scrollbar = tk.Scrollbar(canvas_frame, orient="vertical", command=canvas.yview)
        self.scrollable_frame = tk.Frame(canvas, bg=Colors.BG_PANEL)
        
        self.scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw", width=420) # 内部フレームの固定幅
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 下部アクションエリア
        bottom_frame = tk.Frame(self.left_frame, bg="#111111", padx=20, pady=20)
        bottom_frame.pack(side=tk.BOTTOM, fill=tk.X)
        
        # 実行ボタン (完全撤去)
        
        # 視点操作 (HUDに移動済みなので削除)

        for i, (key, data) in enumerate(PARTS_MASTER.items()):
            row_frame = tk.Frame(self.scrollable_frame, bg="white", pady=5, padx=5, highlightbackground="#eeeeee", highlightthickness=1)
            row_frame.pack(fill=tk.X, pady=2)
            color_box = tk.Label(row_frame, bg=data['color'], width=2); color_box.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 5))
            info_frame = tk.Frame(row_frame, bg="white"); info_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            tk.Label(info_frame, text=data['name'], bg="white", font=("Meiryo", 10, "bold"), anchor="w").pack(fill=tk.X)
            spec_text = f"{key} | {data['w']}x{data['d']}x{data['h']}mm | {data['weight']}kg"
            tk.Label(info_frame, text=spec_text, bg="white", fg="gray", font=("Meiryo", 8), anchor="w").pack(fill=tk.X)
            var = tk.IntVar(value=0); self.qty_vars[key] = var
            # リアルタイム更新のバインディング
            spin = tk.Spinbox(row_frame, from_=0, to=40, textvariable=var, width=5, font=("Meiryo", 12), command=self.run_simulation)
            spin.bind("<KeyRelease>", lambda e: self.run_simulation())
            spin.pack(side=tk.RIGHT, padx=5); tk.Label(row_frame, text="個", bg="white").pack(side=tk.RIGHT)

        # 結果表示エリア（近未来風：黒背景・ネオン文字） → 高さ拡張
        self.result_text = st.ScrolledText(bottom_frame, width=30, height=12, font=("MS Gothic", 9), 
                                           bg="#000000", fg="#00FF00",  # 完全な黒背景
                                           insertbackground="white", relief="flat", highlightthickness=1, highlightbackground="#00FFFF") # シアン枠
        self.result_text.pack(fill=tk.X, pady=(10, 10))
        
        # タグ設定
        self.result_text.tag_config("red", foreground="#FF3333", font=("MS Gothic", 9, "bold"))
        self.result_text.tag_config("blue", foreground="#33FFFF", font=("MS Gothic", 9, "bold"))
        self.result_text.tag_config("green", foreground="#66FF66", font=("MS Gothic", 9, "bold"))
        self.result_text.tag_config("yellow", foreground="#FFCC00", font=("MS Gothic", 9, "bold"))
        self.result_text.tag_config("white", foreground="#FFFFFF", font=("MS Gothic", 9))

        # 追加提案・決定エリア（ダークテーマ化）
        suggest_frame = tk.Frame(bottom_frame, bg="#111111", pady=5)
        suggest_frame.pack(fill=tk.X)
        
        # アイコン風 decoration
        tk.Label(suggest_frame, text="✚", bg="#111111", fg="#00FFFF", font=("Meiryo", 10)).pack(side=tk.LEFT)
        tk.Label(suggest_frame, text="提案追加:", bg="#111111", fg="white", font=("Meiryo", 9, "bold")).pack(side=tk.LEFT, padx=(0, 5))
        
        self.suggest_var = tk.StringVar()
        self.suggest_combo = ttk.Combobox(suggest_frame, textvariable=self.suggest_var, state="readonly", width=18)
        self.suggest_combo.pack(side=tk.LEFT, padx=2)
        
        self.add_qty_var = tk.IntVar(value=1)
        tk.Spinbox(suggest_frame, from_=1, to=40, textvariable=self.add_qty_var, width=3, font=("Meiryo", 10)).pack(side=tk.LEFT, padx=5)
        
        # 決定ボタン（HUDスタイル）
        decide_btn_frame = tk.Frame(suggest_frame, bg="#00FFFF", padx=1, pady=1) # Border
        decide_btn_frame.pack(side=tk.LEFT, padx=5)
        tk.Button(decide_btn_frame, text="決定", command=self.add_suggested_item, 
                  bg="black", fg="#00FFFF", activebackground="#003333", activeforeground="white",
                  font=("Meiryo", 9, "bold"), relief="flat", bd=0).pack(fill=tk.BOTH)

    def _create_hud_controls(self):
        # 右下HUD (Overlay風に配置)
        hud_frame = tk.Frame(self.right_frame, bg=Colors.BG_MAIN) # 背景色をメインに合わせる
        hud_frame.place(relx=0.98, rely=0.98, anchor="se")
        
        # 近未来風ボタンスタイル
        btn_base_style = {
            "bg": "black", 
            "fg": "#00FFFF",  # Cyan Text
            "activebackground": "#003333",
            "activeforeground": "white",
            "font": ("Meiryo", 9, "bold"),
            "relief": "flat",
            "bd": 1,
            "highlightthickness": 0
        }
        
        def mk_hud_btn(parent, text, cmd, bg_color="black", fg_color="#00FFFF"):
            # 枠線風に見せるためのフレーム
            frm = tk.Frame(parent, bg="#00FFFF", padx=1, pady=1) # Cyan Border
            frm.pack(side=tk.LEFT, padx=4)
            
            btn = tk.Button(frm, text=text, command=cmd, **btn_base_style)
            btn.configure(bg=bg_color, fg=fg_color)
            btn.pack(fill=tk.BOTH)
            return btn
        
        # クリアボタン (警告色っぽく)
        mk_hud_btn(hud_frame, " CLEAR ", self.clear_all_items, bg_color="#220000", fg_color="#FF3333") # Red accent
        
        # 操作系
        mk_hud_btn(hud_frame, " ⟲ ", lambda: self.rotate_view(10))
        mk_hud_btn(hud_frame, " ⟳ ", lambda: self.rotate_view(-10))
        mk_hud_btn(hud_frame, " RESET VIEW ", self.reset_view)

    def clear_all_items(self):
        for var in self.qty_vars.values():
            var.set(0)
        self.run_simulation()

    def on_var_change(self, *args):
        # This method is no longer used as spinboxes directly call run_simulation
        pass

    def run_simulation(self):
        # Manual run button triggers everything immediately
        if self.suggestion_timer: self.root.after_cancel(self.suggestion_timer)
        self.run_basic_simulation()
        self.run_suggestions()

    def run_basic_simulation(self):
        # Part 1: Fast simulation (Loading & Status & 3D View)
        items_to_load = []
        for key, var in self.qty_vars.items():
            try:
                qty = var.get()
            except:
                qty = 0 # Handle empty input during typing
            
            if qty > 0:
                master = PARTS_MASTER[key]
                for i in range(qty): items_to_load.append(Item(key, master, i))
        
        self.container = Container()
        self.container.load_items(items_to_load)
        cog, devs = self.container.get_cog_stats()

        self.result_text.delete("1.0", tk.END)
        if len(items_to_load) == 0:
            self.result_text.insert(tk.END, "個数を入力して実行してください。\n")
            self.draw_empty_container()
            return

        is_weight_ok = self.container.total_weight <= self.container.max_weight
        weight_ratio = (self.container.total_weight / self.container.max_weight) * 100
        # 左右の偏荷重(devs[1])は条件に基づき考慮外とする
        is_cog_ok = abs(devs[0]) < 5
        
        if not is_weight_ok: status = "重量オーバー"; color = "red"
        elif not is_cog_ok: status = "偏荷重注意"; color = "red"
        elif len(self.container.unloaded_items) > 0: status = "一部積載不可"; color = "blue"
        else: status = "安全"; color = "blue"

        self.result_text.insert(tk.END, f"判定: [{status}]\n", color)
        self.result_text.insert(tk.END, f"総重量: {self.container.total_weight:,}kg ({weight_ratio:.1f}%)\n")
        
        if len(items_to_load) > 40:
            self.result_text.insert(tk.END, "※警告: 合計ケース数が通常の上限(40)を上回っています\n", "yellow")
        
        if len(self.container.unloaded_items) > 0:
            self.result_text.insert(tk.END, "-"*30 + "\n")
            self.result_text.insert(tk.END, f"⚠ 積載不可 (計{len(self.container.unloaded_items)}個):\n", "red")
            unloaded_counts = Counter([item.name for item in self.container.unloaded_items])
            for name, count in unloaded_counts.items():
                self.result_text.insert(tk.END, f" ・{name}: {count}個\n", "red")
                
        self.draw_3d_graph(cog, devs)

    def run_suggestions(self):
        # Part 2: Heavy simulation (Additional Suggestions)
        if not self.container or len(self.container.items) == 0: return

        # Note: UI update must happen on main thread. 
        # Since this calculation is synchronous and optimized (now fast), we run it directly.
        # If it was still slow, we would need threading.
        
        loadable_counts = self.container.get_loadable_counts(PARTS_MASTER)
        
        self.result_text.insert(tk.END, "-"*30 + "\n")
        self.result_text.insert(tk.END, "【空きスペース提案】\n")
        
        combo_values = []
        if not loadable_counts:
            self.result_text.insert(tk.END, " 追加積載は難しい状態です\n")
            self.suggest_combo['values'] = []
            self.suggest_var.set("")
        else:
            self.result_text.insert(tk.END, " ↓以下の部品なら追加可能:\n", "green")
            for key, count in loadable_counts.items():
                name = PARTS_MASTER[key]['name']
                self.result_text.insert(tk.END, f" ・{name}: あと{count}個\n", "green")
                combo_values.append(f"{name} ({key})")
            
            self.suggest_combo['values'] = combo_values
            if combo_values: self.suggest_combo.current(0)

    def add_suggested_item(self):
        selection = self.suggest_var.get()
        if not selection: return
        
        # "Name (Key)" から Key を抽出
        tokens = selection.split('(')
        if len(tokens) >= 2:
            key = tokens[-1].strip(')')
            try:
                add_qty = self.add_qty_var.get()
            except:
                return 
            
            if key in self.qty_vars:
                current_val = self.qty_vars[key].get()
                self.qty_vars[key].set(current_val + add_qty)
                self.run_simulation()

    def draw_empty_container(self):
        if self.fig: plt.close(self.fig); 
        for w in self.canvas_frame.winfo_children(): w.destroy()
        
        plt.style.use('dark_background')
        self.fig = plt.figure(figsize=(8, 6), dpi=100)
        self.fig.patch.set_facecolor(Colors.BG_MAIN)
        
        self.ax = self.fig.add_subplot(111, projection='3d')
        self.ax.set_facecolor(Colors.BG_MAIN) # Plot bg
        
        c = Container()
        self.ax.set_title("バンニング計画図 (40ft)", fontsize=14, color=Colors.ACCENT_MAIN)
        self.ax.set_xlim([0, c.w]); self.ax.set_ylim([0, c.d]); self.ax.set_zlim([0, c.h])
        self.ax.set_box_aspect((c.w, c.d, c.h))
        
        # コンテナの縁 (シアン色のネオン風)
        edges = [([0, c.w, c.w, 0, 0], [0, 0, c.d, c.d, 0], [0, 0, 0, 0, 0]), ([0, c.w, c.w, 0, 0], [0, 0, c.d, c.d, 0], [c.h]*5)]
        for x in [0, c.w]:
            for y in [0, c.d]: edges.append(([x, x], [y, y], [0, c.h]))
        
        for xs, ys, zs in edges: 
            self.ax.plot(xs, ys, zs, color=Colors.ACCENT_MAIN, lw=1.5, alpha=0.5, zorder=0)
            
        # 見た目をきれいにするためにグリッドを隠す
        self.ax.grid(False)
        self.ax.set_xlabel('L (mm)', color=Colors.TEXT_DIM)
        self.ax.set_ylabel('W (mm)', color=Colors.TEXT_DIM)
        self.ax.set_zlabel('H (mm)', color=Colors.TEXT_DIM)
        self.ax.tick_params(colors=Colors.TEXT_DIM)
        
        # 各面の背景を透明に
        self.ax.xaxis.pane.fill = False
        self.ax.yaxis.pane.fill = False
        self.ax.zaxis.pane.fill = False
        
        canvas = FigureCanvasTkAgg(self.fig, master=self.canvas_frame); canvas.draw()
        canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)

    def draw_3d_graph(self, cog, devs):
        if self.fig: plt.close(self.fig); 
        for w in self.canvas_frame.winfo_children(): w.destroy()
        
        plt.style.use('dark_background')
        self.fig = plt.figure(figsize=(8, 6), dpi=100)
        self.fig.patch.set_facecolor(Colors.BG_MAIN)
        
        self.ax = self.fig.add_subplot(111, projection='3d')
        self.ax.set_facecolor(Colors.BG_MAIN)
        
        self.ax.set_title("PLAN: 40ft", fontsize=14, color=Colors.ACCENT_MAIN)
        c = self.container
        self.ax.set_xlim([0, c.w]); self.ax.set_ylim([0, c.d]); self.ax.set_zlim([0, c.h]); self.ax.set_box_aspect((c.w, c.d, c.h))

        # コンテナの縁
        edges = [([0, c.w, c.w, 0, 0], [0, 0, c.d, c.d, 0], [0, 0, 0, 0, 0]), ([0, c.w, c.w, 0, 0], [0, 0, c.d, c.d, 0], [c.h]*5)]
        for x in [0, c.w]:
            for y in [0, c.d]: edges.append(([x, x], [y, y], [0, c.h]))
        for xs, ys, zs in edges: 
            self.ax.plot(xs, ys, zs, color=Colors.ACCENT_MAIN, lw=1.5, alpha=0.3, zorder=0)

        # 部品の描画
        for item in c.items:
            x, y, z = item.position; dx, dy, dz = item.w, item.d, item.h
            xx = [x, x+dx, x+dx, x, x, x+dx, x+dx, x]; yy = [y, y, y+dy, y+dy, y, y, y+dy, y+dy]; zz = [z, z, z, z, z+dz, z+dz, z+dz, z+dz]
            verts = [[(xx[i], yy[i], zz[i]) for i in [0, 1, 5, 4]], [(xx[i], yy[i], zz[i]) for i in [7, 6, 2, 3]],
                     [(xx[i], yy[i], zz[i]) for i in [0, 3, 7, 4]], [(xx[i], yy[i], zz[i]) for i in [1, 2, 6, 5]],
                     [(xx[i], yy[i], zz[i]) for i in [0, 1, 2, 3]], [(xx[i], yy[i], zz[i]) for i in [4, 5, 6, 7]]]
            self.ax.add_collection3d(Poly3DCollection(verts, facecolors=item.color, linewidths=0.5, edgecolors='k', alpha=0.8, zorder=1))
            
        # 重心と中心
        self.ax.plot([cog[0]], [cog[1]], [cog[2]], marker='o', markersize=10, color=Colors.ACCENT_HOT, 
                     markeredgecolor='white', markeredgewidth=1, linestyle='None', label='COG', zorder=100)
        self.ax.plot([c.w/2], [c.d/2], [c.h/2], marker='x', markersize=10, color=Colors.ACCENT_MAIN, 
                     markeredgecolor=Colors.ACCENT_MAIN, markeredgewidth=2, linestyle='None', label='Center', zorder=100)
        
        self.ax.grid(False)
        self.ax.set_xlabel('L', color=Colors.TEXT_DIM); self.ax.set_ylabel('W', color=Colors.TEXT_DIM); self.ax.set_zlabel('H', color=Colors.TEXT_DIM)
        self.ax.tick_params(colors=Colors.TEXT_DIM)
        self.ax.xaxis.pane.fill = False; self.ax.yaxis.pane.fill = False; self.ax.zaxis.pane.fill = False
        
        self.ax.legend(loc='upper right', facecolor=Colors.BG_CARD, edgecolor=Colors.BG_PANEL, labelcolor=Colors.TEXT_MAIN)
        canvas = FigureCanvasTkAgg(self.fig, master=self.canvas_frame); canvas.draw()
        canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)
    
    def rotate_view(self, angle):
        if self.ax: self.ax.azim += angle; self.fig.canvas.draw_idle()
    def reset_view(self):
        if self.ax: self.ax.view_init(elev=30, azim=-60); self.fig.canvas.draw_idle()

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()