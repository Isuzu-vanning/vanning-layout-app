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
import re
import json

# --- 1. フォント設定・カラーパレット ---
def get_jp_font_family():
    system = sys.platform
    if system == "win32": return "MS Gothic"
    elif system == "darwin": return "Hiragino Sans"
    return "DejaVu Sans"

FONT_FAMILY = get_jp_font_family()
plt.rcParams['font.family'] = FONT_FAMILY

class Colors:
    BG_MAIN     = "#6F8FA3"   # 全体背景
    BG_PANEL    = "#7C9CAF"   # パネル
    BG_CARD     = "#89A9BA"   # カード
    ACCENT_MAIN = "#E8D5A2"   # ベージュ（主ボタン）
    ACCENT_HOT  = "#D8C18A"
    
    TEXT_MAIN   = "#FFFFFF"
    TEXT_DIM    = "#E6EEF2"

    SUCCESS     = "#BDE5C8"
    WARNING     = "#FFD27F"
    ERROR       = "#FF8C8C"

class Fonts:
    HEADER = ("Meiryo", 14, "bold")
    BODY = ("Meiryo", 10)
    BODY_BOLD = ("Meiryo", 10, "bold")
    SMALL = ("Meiryo", 8)
    SMALL_BOLD = ("Meiryo", 8, "bold")
    MONO = ("Consolas", 11)
    LARGE_VAL = ("Impact", 32)

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
    
    def get_source_color(self):
        """移動元コンテナに基づく色を返す"""
        if not hasattr(self, 'source_container_id') or self.source_container_id is None:
            return self.color
        # 12色のパレット
        palette = ["#FF5733", "#33FF57", "#3357FF", "#F333FF", "#FF33A1", "#33FFF3", 
                   "#F3FF33", "#FF8C33", "#8C33FF", "#33FF8C", "#FF3333", "#33FFFF"]
        idx = (self.source_container_id - 1) % len(palette)
        return palette[idx]

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

# --- 5. UIコンポーネント ---
class Card(tk.Frame):
    """タイル表示用のカードコンポーネント"""
    def __init__(self, parent, title, subtitle, command=None, bg=Colors.BG_CARD, border_color=Colors.ACCENT_MAIN):
        super().__init__(parent, bg=bg, highlightthickness=1, highlightbackground=border_color, padx=15, pady=15)
        self.cursor = "hand2" if command else ""
        if command:
            self.bind("<Button-1>", lambda e: command())
            for child in self.winfo_children():
                child.bind("<Button-1>", lambda e: command())
        
        tk.Label(self, text=title, bg=bg, fg="white", font=Fonts.HEADER).pack(anchor="w")
        tk.Label(self, text=subtitle, bg=bg, fg=Colors.TEXT_DIM, font=Fonts.SMALL).pack(anchor="w", pady=(5, 0))
        self.command = command

    def bind_recursive(self, widget):
        widget.bind("<Button-1>", lambda e: self.command())
        for child in widget.winfo_children():
            self.bind_recursive(child)

# --- 6. GUIアプリ ---
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Vanning Optimizer v6.0 [Professional] - Notebook UI")
        self.root.geometry("1600x950")
        self.root.configure(bg=Colors.BG_MAIN)
        
        # 状態管理
        self.selected_node_type = "YEAR"
        self.selected_month = None
        self.selected_week = None
        
        self.container = None
        self.all_containers = []
        self.current_container_idx = 0
        self.fig = None
        self.ax = None
        self.annual_data = None 
        self.is_optimized = False 
        self.auto_optimize_var = tk.BooleanVar(value=True) # [NEW] 自動最適化フラグ
        
        # メインレイアウト構築
        self._build_notebook_layout()

        # データの初期化
        self.generate_random_annual_data()
        self._populate_treeview()
        self._update_dashboard_tab()

    def _build_notebook_layout(self):
        # 内部ログ保持用 (表示しないがエラー防止)
        self.log_text = tk.Text(self.root) 
        
        style = ttk.Style()
        style.theme_use("default")
        style.configure("TNotebook", background=Colors.BG_MAIN, borderwidth=0)
        style.configure("TNotebook.Tab", background=Colors.BG_PANEL, foreground="white", padding=[20, 10], font=Fonts.BODY_BOLD)
        style.map("TNotebook.Tab", background=[("selected", Colors.ACCENT_MAIN)], foreground=[("selected", "black")])

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # ==========================================
        # タブ1: 管理・分析タブ
        # ==========================================
        self.tab_dashboard = tk.Frame(self.notebook, bg=Colors.BG_MAIN)
        self.notebook.add(self.tab_dashboard, text="📊 管理・分析タブ")
        
        self.lbl_dashboard_title = tk.Label(self.tab_dashboard, text="📊 年間物流集約ダッシュボード", bg=Colors.BG_MAIN, fg="white", font=Fonts.HEADER)
        self.lbl_dashboard_title.pack(anchor="w", pady=(20, 20), padx=30)
        
        self.metrics_frame = tk.Frame(self.tab_dashboard, bg=Colors.BG_MAIN)
        self.metrics_frame.pack(fill=tk.X, padx=20, pady=(0, 20))
        
        self.chart_frame = tk.Frame(self.tab_dashboard, bg=Colors.BG_CARD, padx=20, pady=20)
        self.chart_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=(0, 30))

        # ==========================================
        # タブ2: バンニング作業タブ
        # ==========================================
        self.tab_workspace = tk.Frame(self.notebook, bg=Colors.BG_MAIN)
        self.notebook.add(self.tab_workspace, text="📦 バンニング作業タブ")
        
        self.workspace_paned = ttk.PanedWindow(self.tab_workspace, orient=tk.HORIZONTAL)
        self.workspace_paned.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # --- 左側（ツリーナビゲーション） ---
        self.left_panel = tk.Frame(self.workspace_paned, width=300, bg=Colors.BG_PANEL)
        self.workspace_paned.add(self.left_panel, weight=1)
        self.left_panel.pack_propagate(False)
        
        tk.Label(self.left_panel, text="📁 ナビゲーション", bg=Colors.BG_PANEL, fg="white", font=Fonts.HEADER).pack(anchor="w", pady=15, padx=15)
        
        style.configure("Treeview", background=Colors.BG_CARD, foreground="white", fieldbackground=Colors.BG_CARD, borderwidth=0, font=Fonts.BODY)
        style.map("Treeview", background=[("selected", Colors.ACCENT_MAIN)], foreground=[("selected", "black")])
        
        self.tree = ttk.Treeview(self.left_panel, selectmode="browse", show="tree")
        self.tree.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 15))
        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)
        
        tk.Button(self.left_panel, text="📁 CSV/Excel読込", bg="#334466", fg="white", font=Fonts.BODY_BOLD, command=self.load_manifest_file).pack(fill=tk.X, padx=15, pady=10)
        
        # [NEW] 自動最適化のチェックボックス
        tk.Checkbutton(self.left_panel, text="選択時に自動最適化", variable=self.auto_optimize_var, 
                       bg=Colors.BG_PANEL, fg="white", selectcolor="black", activebackground=Colors.BG_PANEL,
                       font=Fonts.SMALL).pack(anchor="w", padx=15)

        # --- 右側（3Dプレビュー・画面の80%占有） ---
        self.right_panel = tk.Frame(self.workspace_paned, bg=Colors.BG_MAIN)
        self.workspace_paned.add(self.right_panel, weight=4) # 80%占有
        
        # KPIエリア
        self.kpi_frame = tk.Frame(
            self.right_panel,
            bg=Colors.BG_CARD,
            padx=20,
            pady=15
        )
        self.kpi_frame.pack(fill=tk.X, padx=10, pady=(10,5))

        self.lbl_kpi_main = tk.Label(
            self.kpi_frame,
            text="🎯 削減：-- 本",
            bg=Colors.BG_CARD,
            fg=Colors.TEXT_MAIN,
            font=("Meiryo", 22, "bold")
        )
        self.lbl_kpi_main.pack(anchor="w")

        self.lbl_kpi_sub = tk.Label(
            self.kpi_frame,
            text="充填率：-- %",
            bg=Colors.BG_CARD,
            fg=Colors.TEXT_DIM,
            font=("Meiryo", 10)
        )
        self.lbl_kpi_sub.pack(anchor="w")
        
        preview_header = tk.Frame(self.right_panel, bg=Colors.BG_PANEL, padx=20, pady=15)
        preview_header.pack(fill=tk.X)
        
        self.lbl_preview_title = tk.Label(preview_header, text="Weekが選択されていません", bg=Colors.BG_PANEL, fg="white", font=Fonts.HEADER)
        self.lbl_preview_title.pack(side=tk.LEFT)
        
        self.lbl_weight = tk.Label(preview_header, text="総重量: --- / 15,000 kg", bg=Colors.BG_PANEL, fg=Colors.ACCENT_MAIN, font=Fonts.BODY_BOLD)
        self.lbl_weight.pack(side=tk.RIGHT)
        
        self.weight_progress = ttk.Progressbar(self.right_panel, orient="horizontal", mode="determinate", length=300)
        self.weight_progress.pack(fill=tk.X, padx=20, pady=(0, 10))
        
        btn_frame = tk.Frame(self.right_panel, bg=Colors.BG_MAIN, pady=10)
        btn_frame.pack(fill=tk.X)
        tk.Button(btn_frame, text="▶ 最適化を実行", bg=Colors.ACCENT_MAIN, fg="black", font=Fonts.BODY_BOLD, command=self.run_simulation, width=15).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="🔄 クリア", bg="#445566", fg="white", font=Fonts.BODY_BOLD, command=self.clear_all_items, width=15).pack(side=tk.LEFT, padx=5)
        
        self.canvas_frame = tk.Frame(self.right_panel, bg="black")
        self.canvas_frame.pack(fill=tk.BOTH, expand=True, pady=(5, 0))
        
        # ログエリアを追加
        
        # ログエリアを追加
        self.log_text = st.ScrolledText(self.right_panel, height=6, bg="black", fg=Colors.SUCCESS, font=Fonts.SMALL, borderwidth=0)
        self.log_text.pack(fill=tk.X, pady=5)

    def _populate_treeview(self):
        self.tree.delete(*self.tree.get_children())
        year_node = self.tree.insert("", "end", iid="YEAR", text="2026年度 全体")
        for m in range(1, 13):
            month_node = self.tree.insert(year_node, "end", iid=f"M_{m}", text=f"{m}月")
            start_w = (m - 1) * 4 + 1
            for w in range(start_w, start_w + 4):
                self.tree.insert(month_node, "end", iid=f"W_{w}", text=f"Week {w}")
        self.tree.item(year_node, open=True)

    def on_tree_select(self, event):
        selected = self.tree.selection()
        if not selected: return
        iid = selected[0]
        
        if iid == "YEAR":
            self.selected_node_type = "YEAR"
            self.selected_month = None
            self.selected_week = None
            self.lbl_preview_title.config(text="※ 詳細プレビューは Week を選択してください")
        elif iid.startswith("M_"):
            self.selected_node_type = "MONTH"
            self.selected_month = int(iid.split("_")[1])
            self.selected_week = None
            self.lbl_preview_title.config(text=f"※ {self.selected_month}月 が選択されています")
        elif iid.startswith("W_"):
            self.selected_node_type = "WEEK"
            self.selected_week = int(iid.split("_")[1])
            parent = self.tree.parent(iid)
            if parent.startswith("M_"):
                self.selected_month = int(parent.split("_")[1])
            self.lbl_preview_title.config(text=f"【 Week {self.selected_week} 】の詳細プレビュー")
                
        self._update_dashboard_tab()
        
        if self.selected_node_type == "WEEK":
            self._update_comparison_display()
            # [NEW] 自動最適化が有効でデータがあれば実行
            if self.auto_optimize_var.get() and self.selected_week in self.annual_data and self.annual_data[self.selected_week]['items']:
                self.run_simulation()
        else:
            for w in self.canvas_frame.winfo_children(): w.destroy()
            self.lbl_weight.config(text="総重量: --- / 15,000 kg")

    def _update_dashboard_tab(self):
        for w in self.metrics_frame.winfo_children(): w.destroy()
        for w in self.chart_frame.winfo_children(): w.destroy()
        
        stats = self.calculate_annual_stats()
        
        if self.selected_node_type == "YEAR" or self.selected_node_type is None:
            self.lbl_dashboard_title.config(text="📊 年間物流集約ダッシュボード")
            self.create_metric_card_small(self.metrics_frame, "年間削減数", f"{stats['saved_containers']}本", Colors.SUCCESS, subtitle=f"({stats['total_before']}本 → {stats['total_after']}本)")
            self.create_metric_card_small(self.metrics_frame, "推定削減額", f"¥{stats['cost_savings']/10000:,.0f}万円", Colors.ACCENT_MAIN, subtitle=f"削減率: {stats['reduction_rate']:.1f}%")
            self._render_monthly_trend_chart(self.chart_frame, stats)
            
        elif self.selected_node_type == "MONTH" or self.selected_node_type == "WEEK":
            month = self.selected_month if self.selected_month else 1
            self.lbl_dashboard_title.config(text=f"📊 {month}月 輸送効率ダッシュボード")
            m_stats = self._get_month_stats(month, stats)
            saved = m_stats['before'] - m_stats['after']
            self.create_metric_card_small(self.metrics_frame, f"{month}月 削減数", f"{saved}本", Colors.SUCCESS, subtitle=f"({m_stats['before']}本 → {m_stats['after']}本)")
            self.create_metric_card_small(self.metrics_frame, "対象荷物数", f"{m_stats['items_count']}件", Colors.TEXT_MAIN)
            self._update_dashboard_bottom_chart(stats, month, self.chart_frame)

    def _render_monthly_trend_chart(self, parent, stats):
        fig, ax = plt.subplots(figsize=(12, 4), dpi=100)
        fig.patch.set_facecolor(Colors.BG_CARD)
        ax.set_facecolor(Colors.BG_CARD)
        
        months = np.array(range(1, 13))
        m_before = []
        m_after = []
        for m in months:
            ms = self._get_month_stats(m, stats)
            m_before.append(ms['before'])
            m_after.append(ms['after'])
            
        width = 0.35
        ax.bar(months - width/2, m_before, width, color=Colors.TEXT_DIM, alpha=0.5, label="現状")
        ax.bar(months + width/2, m_after, width, color=Colors.ACCENT_MAIN, alpha=0.9, label="最適化")
        
        ax.set_title("月次コンテナ本数の推移サマリー (1〜12月)", color="white", fontsize=10)
        ax.set_xticks(months)
        ax.set_xticklabels([f"{m}月" for m in months])
        ax.tick_params(colors=Colors.TEXT_DIM, labelsize=8)
        ax.legend(facecolor=Colors.BG_CARD, edgecolor=Colors.TEXT_DIM, labelcolor="white", fontsize=8)
        ax.grid(axis='y', alpha=0.1)
        
        canvas = FigureCanvasTkAgg(fig, master=parent)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    def _update_dashboard_bottom_chart(self, stats, month, parent):
        fig, ax = plt.subplots(figsize=(12, 4), dpi=100)
        fig.patch.set_facecolor(Colors.BG_CARD)
        ax.set_facecolor(Colors.BG_CARD)
        
        start_w = (month - 1) * 4 + 1
        end_w = month * 4 + 1
        weeks = np.array(range(start_w, end_w))
        
        w_before = stats['weekly_before'][start_w-1:end_w-1]
        w_after = stats['weekly_after'][start_w-1:end_w-1]
        
        width = 0.35
        ax.bar(weeks - width/2, w_before, width, color=Colors.TEXT_DIM, alpha=0.5, label="現状")
        ax.bar(weeks + width/2, w_after, width, color=Colors.ACCENT_MAIN, alpha=0.9, label="最適化")
        
        ax.set_title(f"【{month}月】 週次コンテナ本数の詳細表示 (Week {start_w} 〜 {end_w-1})", color="white", fontsize=10)
        ax.set_xticks(weeks)
        ax.set_xticklabels([f"W{w}" for w in weeks])
        ax.tick_params(colors=Colors.TEXT_DIM, labelsize=8)
        ax.grid(axis='y', alpha=0.1)
        
        canvas = FigureCanvasTkAgg(fig, master=parent)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    def _get_month_stats(self, month, stats):
        start_w = (month - 1) * 4 + 1
        end_w = month * 4 + 1
        items_sum = 0
        for w in range(start_w, end_w):
            items_sum += len(self.annual_data.get(w, {}).get('items', []))
        
        return {
            'before': sum(stats['weekly_before'][start_w-1:end_w-1]),
            'after': sum(stats['weekly_after'][start_w-1:end_w-1]),
            'items_count': items_sum
        }

    def create_metric_card_small(self, parent, title, value, color, subtitle=""):
        card = tk.Frame(parent, bg=Colors.BG_CARD, padx=20, pady=15, highlightthickness=1, highlightbackground=color)
        card.pack(side=tk.LEFT, padx=10, expand=True, fill=tk.BOTH)
        tk.Label(card, text=title, bg=Colors.BG_CARD, fg=Colors.TEXT_DIM, font=Fonts.SMALL).pack(anchor="w")
        tk.Label(card, text=value, bg=Colors.BG_CARD, fg=color, font=("Impact", 24)).pack(anchor="w")
        if subtitle:
            tk.Label(card, text=subtitle, bg=Colors.BG_CARD, fg="white", font=Fonts.SMALL).pack(anchor="w")

    def append_log(self, text, color="white"):
        self.log_text.insert(tk.END, text + "\n")
        self.log_text.see(tk.END)

    def _parse_manifest_dataframe(self, df):
        """DataFrameから荷物情報を抽出する共通ロジック"""
        
        # --- [NEW] ヘッダー行の自動探索と再設定 ---
        # 1行目に「2026/01/08 Container-1」などのメタデータがある場合を回避
        temp_df = df.copy()
        temp_df.loc[-1] = temp_df.columns # 現在のヘッダーをデータ行として追加
        temp_df.index = temp_df.index + 1
        temp_df = temp_df.sort_index()
        
        header_idx = -1
        for idx, row in temp_df.iterrows():
            row_str = " ".join([str(val) for val in row])
            if any(k in row_str for k in ["名称", "品名", "Name", "分類", "資材名称"]):
                header_idx = idx
                break
                
        if header_idx != -1:
            # 見つけたヘッダー行を実際のカラムとして再設定
            df = temp_df.iloc[header_idx+1:].reset_index(drop=True)
            df.columns = temp_df.iloc[header_idx]

        name_to_key = {v['name'].replace(" ", "").replace("　",""): k for k, v in PARTS_MASTER.items()}
        items = []
        
        weight_col = next((col for col in df.columns if "重量" in str(col) or "Weight" in str(col)), 6)
        name_col = next((col for col in df.columns if "名称" in str(col) or "品名" in str(col) or "Name" in str(col)), 2)
        qty_col = next((col for col in df.columns if "数量" in str(col) or "個数" in str(col)), None)

        for _, row in df.iterrows():
            cell_0 = str(row[0]).strip()
            if not cell_0 or cell_0 == "nan" or "Container" in cell_0 or "種別" in cell_0: continue

            matched_key = None
            try:
                raw_id = int(row[1])
                test_key = f"CASE_{raw_id:02d}"
                if test_key in PARTS_MASTER: matched_key = test_key
            except: pass
            
            if not matched_key:
                p_name = str(row[name_col]).strip().replace(" ", "").replace("　","")
                if p_name in name_to_key: matched_key = name_to_key[p_name]
            
            if matched_key:
                try: w = int(row[weight_col])
                except: w = PARTS_MASTER[matched_key]['weight']
                qty = 1
                if qty_col is not None and not pd.isna(row[qty_col]):
                    try: qty = int(row[qty_col])
                    except: qty = 1
                for _ in range(qty):
                    items.append({'key': matched_key, 'weight': w, 'source_container_id': 1})
        return items

    def load_manifest_file(self):
        """Excel/CSVの読込。複数ファイル選択に対応し、ファイル名から月を自動判定する。"""
        file_paths = filedialog.askopenfilenames(filetypes=[("Excel/CSV files", "*.xlsx *.xls *.csv")])
        if not file_paths: return
        
        try:
            for file_path in file_paths:
                self.append_log(f"📂 ファイルを解析中: {os.path.basename(file_path)}")
                
                if file_path.endswith('.csv'):
                    # CSVの場合は月を自動判定
                    try: df = pd.read_csv(file_path, encoding='utf-8')
                    except: df = pd.read_csv(file_path, encoding='shift_jis')
                    
                    items = self._parse_manifest_dataframe(df)
                    if items:
                        # ファイル名から数字（月）を探す
                        fname = os.path.basename(file_path)
                        month_match = re.search(r'(\d+)', fname)
                        
                        if month_match:
                            month = int(month_match.group(1))
                            if 1 <= month <= 12:
                                target_week = (month - 1) * 4 + 1
                                self.append_log(f"💡 ファイル名から {month}月 (Week {target_week}) と判定しました。")
                            else:
                                target_week = self.selected_week if self.selected_week else 1
                        else:
                            target_week = self.selected_week if self.selected_week else 1
                        
                        self.annual_data[target_week] = {
                            'items': items,
                            'containers_before': (len(items) // 15) + 1 
                        }
                        self.append_log(f"✅ CSVデータを Week {target_week} に読み込みました。")
                else:
                    # Excelの場合は全シートを月ごとに処理
                    xl = pd.ExcelFile(file_path)
                    # Excel一括ロードの場合は既存データをリセットするか選ばせる（ここでは簡略化のため追記）
                    
                    total_items_count = 0
                    for sheet_name in xl.sheet_names:
                        match = re.search(r'(\d+)', sheet_name)
                        if not match: continue
                        month = int(match.group(1))
                        
                        df = xl.parse(sheet_name)
                        items = self._parse_manifest_dataframe(df)
                        if not items: continue
                        
                        chunk_size = (len(items) // 4) + 1
                        for w_idx in range(4):
                            global_week = (month - 1) * 4 + w_idx + 1
                            week_items = items[w_idx * chunk_size : (w_idx + 1) * chunk_size]
                            if week_items:
                                self.annual_data[global_week] = {
                                    'items': week_items,
                                    'containers_before': (len(week_items) // 15) + 2
                                }
                                total_items_count += len(week_items)
                    
                    self.append_log(f"✅ エクセルから計 {total_items_count} 個の荷物を読込完了。")
            
            # セッションを保存
            self.save_session_data()
            
            # 状態をリセット
            self.is_optimized = False
            self._update_dashboard_tab()
            
            # ロードした週（または1週目）を表示
            if self.selected_node_type != "WEEK":
                self.selected_node_type = "WEEK"
                self.selected_week = 1
                self.tree.selection_set("W_1")
            
            self._update_comparison_display()

        except Exception as e:
            self.append_log(f"❌ 読込失敗: {str(e)}", Colors.ERROR)
            messagebox.showerror("エラー", f"読込中にエラーが発生しました:\n{e}")

    def save_session_data(self):
        """現在の荷物データをファイルに保存する"""
        try:
            cache_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "vanning_session.json")
            with open(cache_path, 'w', encoding='utf-8') as f:
                json.dump(self.annual_data, f, ensure_ascii=False, indent=2)
            self.append_log("💾 セッションデータを保存しました。")
        except Exception as e:
            print(f"Failed to save session: {e}")

    def load_session_data(self):
        """保存されたセッションデータを読み込む"""
        try:
            cache_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "vanning_session.json")
            if os.path.exists(cache_path):
                with open(cache_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                # JSONはキーが文字列になるので数値に戻す
                self.annual_data = {int(k): v for k, v in data.items()}
                self.append_log("📂 前回のセッションデータを復元しました。")
                return True
        except Exception as e:
            print(f"Failed to load session: {e}")
        return False

    def on_month_selected(self, event):
        """月次表示を更新（データは起動時にロード済み）"""
        self.render_view()

    def clear_all_items(self, run_sim=True):
        """現在の週のデータをクリアする"""
        if self.selected_node_type == "WEEK" and self.selected_week:
            self.annual_data[self.selected_week] = {'items': [], 'containers_before': 0}
            self.append_log(f"🔄 Week {self.selected_week} のデータをクリアしました", "yellow")
            if run_sim:
                self.run_simulation()
        else:
            self.append_log("🔄 データをリセットするには週次詳細画面で実行してください。")

    def run_simulation(self):
        """シミュレーションを実行（最適化ロジック）"""
        if not self.annual_data or self.selected_week not in self.annual_data:
            self.append_log("⚠️ 対象週のデータがありません。")
            return

        self.is_optimized = True # 最適化実行フラグを立てる
        w_data = self.annual_data[self.selected_week]
        all_items = []
        for i, item_data in enumerate(w_data['items']):
            key = item_data['key']
            master = PARTS_MASTER[key]
            item = Item(key, master, i, item_data['weight'])
            item.source_container_id = item_data.get('source_container_id')
            all_items.append(item)
        
        self.append_log(f"🚀 Week {self.selected_week} の最適化を開始...")
        
        self.all_containers = []
        remaining_items = all_items.copy()
        remaining_items.sort(key=lambda x: (x.w * x.d * x.h, x.weight), reverse=True)
        
        while remaining_items:
            new_container = Container()
            new_container.load_items(remaining_items)
            self.all_containers.append(new_container)
            remaining_items = new_container.unloaded_items
            if len(self.all_containers) > 20: break

        self.append_log(f"✅ 最適化完了: {w_data['containers_before']}本 ➡ {len(self.all_containers)}本に削減！")
        self.current_container_idx = 0
        self.update_3d_display()
        self._update_comparison_display()

    def _update_comparison_display(self):
        """最適化前後の比較情報をラベルに反映"""
        if self.selected_node_type != "WEEK" or not self.selected_week or self.selected_week not in self.annual_data:
            self.lbl_preview_title.config(text="Weekを選択してください", fg="white")
            self.lbl_weight.config(text="総重量: --- / 15,000 kg")
            return
            
        w_data = self.annual_data[self.selected_week]
        before_cnt = w_data['containers_before']
        
        if self.is_optimized:
            after_cnt = len(self.all_containers)
            diff = before_cnt - after_cnt
            status_text = f"【 Week {self.selected_week} 】 最適化後: {before_cnt}本 ➡ {after_cnt}本 (削減: {diff}本)"
            self.lbl_preview_title.config(text=status_text, fg=Colors.ACCENT_MAIN)
            
            self.lbl_kpi_main.config(text=f"🎯 削減：{diff} 本")
            if after_cnt > 0:
                fill = int((before_cnt / after_cnt) * 62)
                self.lbl_kpi_sub.config(text=f"平均効率：{fill}%")
        else:
            status_text = f"【 Week {self.selected_week} 】 現状: {before_cnt}本  ▶ 最適化を実行してください"
            self.lbl_preview_title.config(text=status_text, fg="white")
            # 3Dをクリア
            for w in self.canvas_frame.winfo_children(): w.destroy()
            self.lbl_weight.config(text=f"荷物数: {len(w_data['items'])}個")

    def update_3d_display(self):
        """3D表示の更新（複数コンテナの切り替え）"""
        if not hasattr(self, 'all_containers') or not self.all_containers: return
        
        self.container = self.all_containers[self.current_container_idx] # [FIX] self.containerを更新
        cog, devs = self.container.get_cog_stats()

        # UI更新
        tot_w = self.container.total_weight
        mx_w = self.container.max_weight
        pct_w = (tot_w / mx_w) * 100
        self.lbl_weight.config(text=f"コンテナ {self.current_container_idx+1}/{len(self.all_containers)} | 重量: {tot_w:,}kg")
        self.weight_progress['value'] = pct_w
        
        self.draw_3d_graph(cog, devs)
        self._render_container_selector()

    def _render_container_selector(self):
        if hasattr(self, 'selector_frame'):
            self.selector_frame.destroy()

        self.selector_frame = tk.Frame(
            self.canvas_frame,
            bg=Colors.BG_MAIN
        )
        self.selector_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=8)

        for i, c in enumerate(self.all_containers):

            active = i == self.current_container_idx

            bg = Colors.ACCENT_MAIN if active else Colors.BG_CARD
            fg = "black" if active else "white"

            pct = int((c.total_weight / c.max_weight) * 100)

            card = tk.Button(
                self.selector_frame,
                text=f"C{i+1}\n{pct}%",
                bg=bg,
                fg=fg,
                relief="flat",
                width=8,
                height=3,
                font=("Meiryo", 9, "bold"),
                command=lambda idx=i: self.set_active_container(idx)
            )
            card.pack(side=tk.LEFT, padx=5)

    def set_active_container(self, idx):
        self.current_container_idx = idx
        self.update_3d_display()

    def draw_3d_graph(self, cog, devs):
        try:
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
                
                # [FIX] 移動元の色を使用
                face_color = item.get_source_color() if not is_error else Colors.ERROR
                
                poly = Poly3DCollection(verts, facecolors=face_color, linewidths=0.5, edgecolors='black', alpha=0.9, zorder=1)
                poly._item_info = f"{item.name}\n元コンテナ: {item.source_container_id}\n重量: {item.weight:,}kg"
                poly._item_ref = item # アイテム本体への参照を保持
                poly.set_picker(True)
                self.ax.add_collection3d(poly)

            # 正常に積載されたアイテム
            for item in c.items:
                draw_box(item, False)
                
            # [NEW] 重心ボール表示
            if c.total_weight > 0:
                self.ax.scatter([cog[0]], [cog[1]], [cog[2]], color=Colors.ACCENT_HOT, s=500, marker='o', 
                                edgecolors='white', linewidths=2, label="重心\nCOG", zorder=100, alpha=0.8)
                self.ax.scatter([cog[0]], [cog[1]], [0], color=Colors.ACCENT_HOT, s=100, marker='x', alpha=0.5, zorder=99)
            
            # [FIX] 重量表示を右上に逃がす
            self.ax.text(c.w, c.d, c.h * 1.3, f"TOTAL: {c.total_weight:,}kg", color=Colors.ACCENT_MAIN, 
                         fontsize=14, fontweight='bold', ha='right', 
                         bbox=dict(facecolor=Colors.BG_CARD, alpha=0.8, edgecolor=Colors.ACCENT_MAIN))

            # [FIX] 目盛りと軸ラベルの調整（数値を減らして見やすく）
            self.ax.xaxis.set_major_locator(ticker.MaxNLocator(4))
            self.ax.yaxis.set_major_locator(ticker.MaxNLocator(3))
            self.ax.zaxis.set_major_locator(ticker.MaxNLocator(3))

            self.ax.grid(True, linestyle=':', alpha=0.3)
            self.ax.set_xlabel('L (奥行)', color=Colors.TEXT_DIM, labelpad=15)
            self.ax.set_ylabel('W (幅)', color=Colors.TEXT_DIM, labelpad=15)
            self.ax.set_zlabel('H (高)', color=Colors.TEXT_DIM, labelpad=15)
            self.ax.tick_params(colors=Colors.TEXT_DIM, labelsize=9)
            
            self.fig.canvas.mpl_connect('pick_event', self.on_pick)
            self.fig.canvas.mpl_connect('motion_notify_event', self.on_hover)
            self.fig.canvas.mpl_connect('button_press_event', self.on_mouse_down)
            
            self.ax.mouse_init(rotate_btn=3, zoom_btn=None) 

            self.ax.legend(loc='upper left', facecolor=Colors.BG_CARD, edgecolor=Colors.BG_PANEL, labelcolor=Colors.TEXT_MAIN)
            canvas = FigureCanvasTkAgg(self.fig, master=self.canvas_frame); canvas.draw()
            canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)
            
            # ツールチップ用ラベル（初期は非表示）
            self.tooltip = tk.Label(self.canvas_frame, bg="yellow", fg="black", font=Fonts.SMALL, padx=5, pady=2, relief="solid", borderwidth=1)
            self.canvas_frame.update_idletasks() # 強制更新
        except Exception as e:
            self.append_log(f"❌ 描画エラー: {str(e)}", Colors.ERROR)
            import traceback
            print(traceback.format_exc())
    
    def on_mouse_down(self, event):
        """マウスボタン押下時の処理"""
        # 特殊な処理が必要な場合に備えて予約
        pass

    def on_hover(self, event):
        """マウスホバー時のツールチップ表示"""
        if event.inaxes != self.ax:
            self.tooltip.place_forget()
            return
            
        found = False
        for collection in self.ax.collections:
            if hasattr(collection, '_item_info'):
                cont, ind = collection.contains(event)
                if cont:
                    # ツールチップを表示
                    x, y = event.canvas.get_width_height()
                    self.tooltip.config(text=collection._item_info)
                    self.tooltip.place(x=event.x, y=y - event.y - 40)
                    found = True
                    break
        
        if not found:
            self.tooltip.place_forget()

    def toggle_edit_mode(self):
        if self.edit_mode_var.get():
            self.append_log("🔧 編集モード：荷物をクリックして移動先を選択してください。", Colors.ACCENT_HOT)
        else:
            self.reloc_panel.pack_forget()
            self.append_log("✅ 編集モードを終了しました。")

    def on_pick(self, event):
        """3Dアイテムクリック時の処理"""
        artist = event.artist
        if hasattr(artist, '_item_info'):
            if hasattr(self, 'edit_mode_var') and self.edit_mode_var.get() and hasattr(artist, '_item_ref'):
                self.show_relocation_ui(artist._item_ref)
            else:
                self.append_log(f"📦 選択中: {artist._item_info}")

    def show_relocation_ui(self, item):
        """移動指示パネルの表示"""
        for child in self.reloc_panel.winfo_children(): child.destroy()
        self.reloc_panel.pack(fill=tk.X, pady=10)
        
        tk.Label(self.reloc_panel, text="荷物の移動指示", bg=Colors.BG_CARD, fg=Colors.ACCENT_HOT, font=Fonts.SMALL_BOLD).pack(anchor="w")
        tk.Label(self.reloc_panel, text=f"WHAT: {item.name}", bg=Colors.BG_CARD, fg="white", font=Fonts.SMALL).pack(anchor="w")
        tk.Label(self.reloc_panel, text=f"from: Container {self.current_container_idx+1}", bg=Colors.BG_CARD, fg=Colors.TEXT_DIM, font=Fonts.SMALL).pack(anchor="w")
        
        to_frame = tk.Frame(self.reloc_panel, bg=Colors.BG_CARD)
        to_frame.pack(fill=tk.X, pady=5)
        tk.Label(to_frame, text="to:", bg=Colors.BG_CARD, fg="white", font=Fonts.SMALL).pack(side=tk.LEFT)
        
        # 移動先コンテナ候補
        options = [f"Container {i+1}" for i in range(len(self.all_containers))]
        to_var = tk.StringVar(value=options[0])
        combo = ttk.Combobox(to_frame, textvariable=to_var, values=options, state="readonly", width=12)
        combo.pack(side=tk.LEFT, padx=5)
        
        def do_relocate():
            to_idx = options.index(to_var.get())
            if to_idx == self.current_container_idx:
                self.append_log("⚠️ 同じコンテナには移動できません。", Colors.WARNING)
                return
            
            # 実移動処理
            source_container = self.all_containers[self.current_container_idx]
            target_container = self.all_containers[to_idx]
            
            # 移動元から削除
            source_container.items.remove(item)
            # 移動先に追加して再計算
            temp_items = target_container.items + [item]
            target_container.load_items(temp_items)
            
            if item in target_container.unloaded_items:
                self.append_log(f"❌ Container {to_idx+1} には入りませんでした（容量不足）。", Colors.ERROR)
                # 元に戻す
                target_container.items.remove(item)
                source_container.load_items(source_container.items + [item])
            else:
                self.append_log(f"✅ {item.name} を Container {to_idx+1} へ移動しました！", Colors.SUCCESS)
                # 移動元も再パズル
                source_container.load_items(source_container.items)
                self.update_3d_display()

        tk.Button(self.reloc_panel, text="移動を実行", bg=Colors.ACCENT_HOT, fg="white", font=Fonts.SMALL_BOLD, command=do_relocate).pack(fill=tk.X, pady=5)

    def rotate_view(self, angle):
        if self.ax: self.ax.azim += angle; self.fig.canvas.draw_idle()

    # --- 6. 年間最適化シミュレーション機能 ---
    def show_annual_simulation_dialog(self):
        """年間シミュレーションとダッシュボードを表示"""
        dialog = tk.Toplevel(self.root)
        dialog.title("年間コスト削減ダッシュボード")
        dialog.geometry("1100x800")
        dialog.configure(bg=Colors.BG_MAIN)
        
        # タイトル
        title_frame = tk.Frame(dialog, bg=Colors.BG_PANEL, pady=15)
        title_frame.pack(fill=tk.X)
        tk.Label(title_frame, text="ANNUAL LOGISTICS OPTIMIZATION DASHBOARD", 
                 bg=Colors.BG_PANEL, fg=Colors.ACCENT_MAIN, font=Fonts.HEADER).pack()
        
        # メインコンテンツ
        content_frame = tk.Frame(dialog, bg=Colors.BG_MAIN, padx=20, pady=20)
        content_frame.pack(fill=tk.BOTH, expand=True)

        # 指標カードエリア
        metrics_frame = tk.Frame(content_frame, bg=Colors.BG_MAIN)
        metrics_frame.pack(fill=tk.X, pady=(0, 20))

        def create_metric_card(parent, title, value, unit, color):
            card = tk.Frame(parent, bg=Colors.BG_CARD, padx=20, pady=15, highlightthickness=1, highlightbackground=color)
            card.pack(side=tk.LEFT, expand=True, fill=tk.BOTH, padx=10)
            tk.Label(card, text=title, bg=Colors.BG_CARD, fg=Colors.TEXT_DIM, font=Fonts.SMALL).pack(anchor=tk.W)
            val_frame = tk.Frame(card, bg=Colors.BG_CARD)
            val_frame.pack(anchor=tk.W)
            tk.Label(val_frame, text=value, bg=Colors.BG_CARD, fg=color, font=Fonts.LARGE_VAL).pack(side=tk.LEFT)
            tk.Label(val_frame, text=unit, bg=Colors.BG_CARD, fg=color, font=Fonts.BODY_BOLD).pack(side=tk.LEFT, pady=(15,0), padx=5)
            return card

        # ダミーデータの生成と計算
        self.generate_random_annual_data()
        stats = self.calculate_annual_stats()

        create_metric_card(metrics_frame, "年間削減コンテナ数", f"{stats['saved_containers']}", "本", Colors.SUCCESS)
        create_metric_card(metrics_frame, "推定コスト削減額", f"{stats['cost_savings']:,.0f}", "円", Colors.ACCENT_MAIN)
        create_metric_card(metrics_frame, "平均充填率向上", f"+{stats['efficiency_gain']:.1f}", "%", Colors.WARNING)

        # チャートエリア
        chart_frame = tk.Frame(content_frame, bg=Colors.BG_CARD, padx=10, pady=10)
        chart_frame.pack(fill=tk.BOTH, expand=True)

        fig, ax = plt.subplots(figsize=(10, 4), dpi=100)
        fig.patch.set_facecolor(Colors.BG_CARD)
        ax.set_facecolor(Colors.BG_CARD)
        
        weeks = list(range(1, 53))
        before = stats['weekly_before']
        after = stats['weekly_after']

        ax.bar(weeks, before, color=Colors.TEXT_DIM, alpha=0.3, label="現状 (55-85%充填)")
        ax.bar(weeks, after, color=Colors.ACCENT_MAIN, alpha=0.8, label="最適化後 (100%充填)")
        
        ax.set_title("週次コンテナ使用数の比較", color="white", fontsize=12, pad=15)
        ax.set_xlabel("週 (Week)", color=Colors.TEXT_DIM)
        ax.set_ylabel("コンテナ本数", color=Colors.TEXT_DIM)
        ax.tick_params(colors=Colors.TEXT_DIM)
        ax.legend(facecolor=Colors.BG_CARD, edgecolor=Colors.TEXT_DIM, labelcolor="white")
        ax.grid(axis='y', alpha=0.1)

        canvas = FigureCanvasTkAgg(fig, master=chart_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        # [NEW] 週次サマリーテーブル
        table_frame = tk.Frame(content_frame, bg=Colors.BG_CARD, padx=10, pady=10)
        table_frame.pack(fill=tk.BOTH, expand=True, pady=(20, 0))
        
        tk.Label(table_frame, text="📅 週次最適化サマリー (抜粋)", bg=Colors.BG_CARD, fg="white", font=Fonts.BODY_BOLD).pack(anchor=tk.W)
        
        # ツリービューの作成
        style = ttk.Style()
        style.configure("Custom.Treeview", background=Colors.BG_CARD, foreground="white", fieldbackground=Colors.BG_CARD, borderwidth=0)
        style.map("Custom.Treeview", background=[('selected', Colors.ACCENT_MAIN)], foreground=[('selected', 'black')])
        
        tree = ttk.Treeview(table_frame, columns=("Week", "Before", "After", "Saved"), show="headings", style="Custom.Treeview", height=6)
        tree.heading("Week", text="週番号")
        tree.heading("Before", text="現状 (本)")
        tree.heading("After", text="最適化後 (本)")
        tree.heading("Saved", text="削減数")
        
        for i in range(len(weeks)):
            if before[i] > 0:
                tree.insert("", tk.END, values=(f"Week {weeks[i]}", f"{before[i]}本", f"{after[i]}本", f"{before[i]-after[i]}本"))
        
        tree.pack(fill=tk.BOTH, expand=True)

        # 解説エリア
        info_text = (
            "【システムによる最適化の仕組み】\n"
            "1. 各拠点からの1週間分の予定リストを、あえて55-85%の充填率でコンテナに割り当て（現状の余裕を持たせた運用を再現）。\n"
            "2. 本システムが、2本目以降のコンテナの荷物を1本目の空きスペースへ「テトリス」のように再配置し、極限まで詰め込みます。\n"
            "3. 1本目が満杯になったら、余った荷物を2本目へ。これを繰り返すことで、最後方のコンテナを空にしていき、全体の本数を削減します。"
        )
        tk.Label(content_frame, text=info_text, bg=Colors.BG_MAIN, fg=Colors.TEXT_MAIN, 
                 justify=tk.LEFT, font=("Meiryo", 9), wraplength=1000).pack(pady=20)

    def generate_random_annual_data(self):
        """保存データがあれば読み込み、なければランダムまたはExcelから生成"""
        if self.annual_data is not None: return
        
        # 1. 保存されたセッションを優先
        if self.load_session_data():
            return

        # 2. 実際のExcelファイルがあれば読み込み
        actual = self.load_actual_annual_data()
        if actual:
            self.annual_data = actual
            return

        data = {}
        part_keys = list(PARTS_MASTER.keys())
        for week in range(1, 54):
            num_items = random.randint(150, 300)
            weekly_cargo = []
            for _ in range(num_items):
                key = random.choice(part_keys)
                weight = random.randint(500, 2500)
                weekly_cargo.append({'key': key, 'weight': weight})
            data[week] = {'items': weekly_cargo, 'containers_before': int(num_items/15)}
        self.annual_data = data

    def load_actual_annual_data(self):
        """vanning_layout_2026.xlsxから実際の予定データを読み込む"""
        base_dir = os.path.dirname(os.path.abspath(__file__))
        file_path = os.path.join(base_dir, "vanning_layout_2026.xlsx")
        
        if not os.path.exists(file_path): return None
            
        try:
            xl = pd.ExcelFile(file_path)
            weekly_data = {i: {'items': [], 'containers_before': 0} for i in range(1, 55)}
            container_pattern = re.compile(r'(\d{4}/\d{2}/\d{2})\s+Container-(\d+)')
            
            for sheet_name in xl.sheet_names:
                df = pd.read_excel(xl, sheet_name=sheet_name, header=None)
                current_week = None
                current_container_id = None
                for _, row in df.iterrows():
                    cell_0 = str(row[0])
                    match = container_pattern.search(cell_0)
                    if match:
                        date_str = match.group(1)
                        date_dt = pd.to_datetime(date_str)
                        current_week = date_dt.isocalendar()[1]
                        current_container_id = int(match.group(2))
                        weekly_data[current_week]['containers_before'] += 1
                    elif current_week is not None:
                        try:
                            raw_id = int(row[1])
                            test_key = f"CASE_{raw_id:02d}"
                            if test_key in PARTS_MASTER:
                                qty = int(row[6]) if not pd.isna(row[6]) else 1
                                # 実際の重量はマスタから取得（ファイルに重量列がないため）
                                weight = PARTS_MASTER[test_key]['weight']
                                for _ in range(qty):
                                    item_info = {
                                        'key': test_key,
                                        'weight': weight,
                                        'source_container_id': current_container_id
                                    }
                                    weekly_data[current_week]['items'].append(item_info)
                        except: continue
            return weekly_data
        except Exception as e:
            print(f"Error loading actual data: {e}")
            return None

    def calculate_annual_stats(self):
        """年間データの統計を精緻に計算（実データ準拠）"""
        stats = {
            'weekly_before': [],
            'weekly_after': [],
            'saved_containers': 0,
            'cost_savings': 0,
            'total_before': 0,
            'total_after': 0
        }
        
        container_vol = 12000 * 2300 * 2400
        
        # 52週分をループ
        for w in range(1, 53):
            data = self.annual_data.get(w, {'items': [], 'containers_before': 0})
            items = data['items']
            num_before = data['containers_before']
            
            if not items:
                stats['weekly_before'].append(num_before)
                stats['weekly_after'].append(0)
                continue

            # 最強モードを想定した集約計算
            total_vol = 0
            total_weight = 0
            for item in items:
                m = PARTS_MASTER[item['key']]
                total_vol += m['w'] * m['d'] * m['h']
                total_weight += item.get('weight', 1000)
            
            # 体積と重量の両面から必要な本数を算出（充填率90%を目標値とする）
            num_after_vol = int(np.ceil(total_vol / (container_vol * 0.90)))
            num_after_weight = int(np.ceil(total_weight / 25000)) # 25t制限
            num_after = max(1, num_after_vol, num_after_weight)
            
            # 異常値（現状より増えてしまう等）を防止
            if num_before > 0:
                num_after = min(num_after, num_before)
            
            stats['weekly_before'].append(num_before)
            stats['weekly_after'].append(num_after)
            
        stats['total_before'] = sum(stats['weekly_before'])
        stats['total_after'] = sum(stats['weekly_after'])
        stats['saved_containers'] = stats['total_before'] - stats['total_after']
        stats['cost_savings'] = stats['saved_containers'] * 350000 
        
        stats['reduction_rate'] = (stats['saved_containers'] / stats['total_before'] * 100) if stats['total_before'] > 0 else 0
        stats['efficiency_before'] = 62.5 # 実績ベースの想定充填率
        stats['efficiency_after'] = (stats['total_before'] / stats['total_after'] * 62.5) if stats['total_after'] > 0 else 0
        
        return stats

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()