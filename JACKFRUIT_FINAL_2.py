# advanced_expense_tracker_with_SIP_and_PDF.py
# Enhanced Expense Tracker - 3 Separate Windows (Light Blue Theme)

import wx
import wx.adv
import csv
import os
import json
from collections import defaultdict
from datetime import datetime
from matplotlib.figure import Figure
from matplotlib.backends.backend_wxagg import FigureCanvasWxAgg as FigureCanvas
import math
import io
import threading
import sys
import subprocess
import re  # for safe_float currency handling

# Optional PDF library
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas as pdfcanvas
    from reportlab.lib.utils import ImageReader
    REPORTLAB_AVAILABLE = True
except Exception:
    REPORTLAB_AVAILABLE = False

# -------------------------
# File names / Preferences (GLOBAL)
# -------------------------
FILENAME = "expenses.csv"
INCOME_FILE = "income.csv"
PREF_FILE = "preferences.json"

# Global references to windows for cross-communication
WINDOWS = {}

# Theme colours
BG_MAIN = wx.Colour(230, 244, 255)   # light blue background
BG_PANEL = wx.Colour(214, 234, 248)
BG_INPUT = wx.Colour(245, 250, 255)
BTN_PRIMARY = wx.Colour(0, 123, 255)
BTN_SECONDARY = wx.Colour(0, 184, 148)

# Currency choices for dropdown
CURRENCY_OPTIONS = ["₹", "$", "€", "£", "¥", "AED", "AUD", "CAD", "SGD"]

# -------------------------
# Preferences helpers
# -------------------------
def load_preferences():
    if not os.path.exists(PREF_FILE):
        prefs = {
            "currency_symbol": "₹",
            "default_monthly_budget": 0.0
        }
        save_preferences(prefs)
        return prefs
    try:
        with open(PREF_FILE, "r") as f:
            return json.load(f)
    except Exception:
        return {"currency_symbol": "₹", "default_monthly_budget": 0.0}

def save_preferences(prefs):
    with open(PREF_FILE, "w") as f:
        json.dump(prefs, f, indent=2)

PREFS = load_preferences()

# -------------------------
# CSV helpers
# -------------------------
def ensure_csv_exists(filename, headers):
    if not os.path.exists(filename):
        with open(filename, "w", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(headers)

def read_expenses():
    ensure_csv_exists(FILENAME, ["Date", "Category", "Amount", "Note"])
    rows = []
    with open(FILENAME, "r", newline="") as f:
        reader = csv.DictReader(f)
        for r in reader:
            rows.append({
                "Date": (r.get("Date") or "").strip(),
                "Category": (r.get("Category") or "").strip(),
                "Amount": (r.get("Amount") or "0").strip(),
                "Note": (r.get("Note") or "").strip()
            })
    return rows

def write_expenses(expenses):
    ensure_csv_exists(FILENAME, ["Date", "Category", "Amount", "Note"])
    with open(FILENAME, "w", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["Date", "Category", "Amount", "Note"])
        for r in expenses:
            writer.writerow([r["Date"], r["Category"], r["Amount"], r.get("Note", "")])

def add_expense(date, category, amount, note=""):
    ensure_csv_exists(FILENAME, ["Date", "Category", "Amount", "Note"])
    with open(FILENAME, "a", newline="") as f:
        writer = csv.writer(f)
        writer.writerow([date, category, amount, note])

# Income CSV helpers
def read_incomes():
    ensure_csv_exists(INCOME_FILE, ["Month", "Income"])
    incomes = {}
    with open(INCOME_FILE, "r", newline="") as f:
        reader = csv.DictReader(f)
        for r in reader:
            m = (r.get("Month") or "").strip()
            try:
                incomes[m] = float((r.get("Income") or "0").strip())
            except Exception:
                incomes[m] = 0.0
    return incomes

def write_incomes(incomes):
    ensure_csv_exists(INCOME_FILE, ["Month", "Income"])
    with open(INCOME_FILE, "w", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["Month", "Income"])
        for month, amt in incomes.items():
            writer.writerow([month, f"{amt:.2f}"])

def set_income_for_month(month_key, amount):
    incomes = read_incomes()
    incomes[month_key] = float(amount)
    write_incomes(incomes)

# -------------------------
# Utilities
# -------------------------
def date_to_month_key(date_str):
    try:
        d = datetime.strptime(date_str, "%Y-%m-%d")
    except Exception:
        try:
            d = datetime.fromisoformat(date_str)
        except Exception:
            try:
                parts = [int(p) for p in date_str.split("-") if p.strip()]
                if len(parts) >= 2:
                    year = parts[0]; month = parts[1]
                    return f"{year:04d}-{month:02d}"
                return None
            except Exception:
                return None
    return f"{d.year:04d}-{d.month:02d}"

def safe_float(x):
    """
    Safely convert amount strings that may contain currency symbols, commas, etc.
    Examples:
        "₹500.00" -> 500.0
        "$1,234.50" -> 1234.5
        " 300 " -> 300.0
    """
    try:
        s = str(x).strip()
        # Remove everything except digits, dot, minus
        cleaned = "".join(ch for ch in s if (ch.isdigit() or ch in ".-"))
        if not cleaned:
            return 0.0
        return float(cleaned)
    except Exception:
        return 0.0

# Global function to refresh all windows
def refresh_all_windows():
    for win in list(WINDOWS.values()):
        if hasattr(win, 'update_charts'):
            win.update_charts()
        if hasattr(win, 'load_table'):
            win.load_table()
        if hasattr(win, 'update_text'):
            win.update_text(win.get_suggestion_text())

# -------------------------
# SIP Calculator dialog
# -------------------------
class SIPDialog(wx.Dialog):
    def __init__(self, parent):
        super().__init__(parent, title="Start SIP — Savings Calculator", size=(440, 400))
        self.SetBackgroundColour(BG_PANEL)
        s = wx.BoxSizer(wx.VERTICAL)

        title = wx.StaticText(self, label="Systematic Investment Plan (SIP)")
        title.SetFont(wx.Font(13, wx.DEFAULT, wx.NORMAL, wx.BOLD))
        s.Add(title, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 10)

        grid = wx.FlexGridSizer(4, 2, 8, 8)
        grid.AddGrowableCol(1)

        grid.Add(wx.StaticText(self, label="Monthly investment ({}):".format(PREFS.get("currency_symbol","₹"))), 0, wx.ALIGN_CENTER_VERTICAL)
        self.monthly_txt = wx.TextCtrl(self)
        self.monthly_txt.SetBackgroundColour(BG_INPUT)
        grid.Add(self.monthly_txt, 0, wx.EXPAND)

        grid.Add(wx.StaticText(self, label="Expected annual return (%):"), 0, wx.ALIGN_CENTER_VERTICAL)
        self.annual_txt = wx.TextCtrl(self, value="12")
        self.annual_txt.SetBackgroundColour(BG_INPUT)
        grid.Add(self.annual_txt, 0, wx.EXPAND)

        grid.Add(wx.StaticText(self, label="Investment period (years):"), 0, wx.ALIGN_CENTER_VERTICAL)
        self.years_txt = wx.TextCtrl(self, value="10")
        self.years_txt.SetBackgroundColour(BG_INPUT)
        grid.Add(self.years_txt, 0, wx.EXPAND)

        grid.Add(wx.StaticText(self, label="Lump-sum goal (optional, {}):".format(PREFS.get("currency_symbol","₹"))), 0, wx.ALIGN_CENTER_VERTICAL)
        self.goal_txt = wx.TextCtrl(self)
        self.goal_txt.SetBackgroundColour(BG_INPUT)
        grid.Add(self.goal_txt, 0, wx.EXPAND)

        s.Add(grid, 0, wx.ALL | wx.EXPAND, 12)

        calc_btn = wx.Button(self, label="Calculate & Suggest")
        calc_btn.SetBackgroundColour(BTN_PRIMARY)
        calc_btn.SetForegroundColour(wx.WHITE)
        calc_btn.Bind(wx.EVT_BUTTON, self.on_calculate)
        s.Add(calc_btn, 0, wx.ALIGN_CENTER | wx.ALL, 6)

        self.result_box = wx.TextCtrl(self, style=wx.TE_MULTILINE | wx.TE_READONLY)
        self.result_box.SetMinSize((380, 140))
        self.result_box.SetBackgroundColour(wx.Colour(250, 252, 255))
        s.Add(self.result_box, 0, wx.ALL | wx.EXPAND, 12)

        self.SetSizer(s)

    def on_calculate(self, event):
        P_txt = self.monthly_txt.GetValue().strip()
        r_txt = self.annual_txt.GetValue().strip()
        y_txt = self.years_txt.GetValue().strip()
        goal_txt = self.goal_txt.GetValue().strip()

        try:
            P = float(P_txt)
            annual = float(r_txt)/100.0
            years = float(y_txt)
            if P < 0 or years <= 0:
                raise ValueError
        except Exception:
            wx.MessageBox("Please enter valid numeric values.", "Invalid input", style=wx.OK | wx.ICON_ERROR)
            return

        periods = int(years * 12)
        monthly_rate = annual/12.0
        if monthly_rate == 0:
            fv = P * periods
        else:
            fv = P * (((1+monthly_rate)**periods - 1) / monthly_rate) * (1+monthly_rate)

        text = []
        text.append(f"Monthly SIP: {PREFS.get('currency_symbol','₹')}{P:.2f}")
        text.append(f"Annual return assumed: {annual*100:.2f}%")
        text.append(f"Period: {years:.1f} years ({periods} months)")
        text.append("")
        text.append(f"Estimated corpus at end: {PREFS.get('currency_symbol','₹')}{fv:,.2f}")

        if goal_txt:
            try:
                goal = float(goal_txt)
                if monthly_rate == 0:
                    req_monthly = goal / periods
                else:
                    req_monthly = goal / ((((1+monthly_rate)**periods - 1) / monthly_rate) * (1+monthly_rate))
                text.append(f"To reach goal {PREFS.get('currency_symbol','₹')}{goal:,.2f}, you need ~ {PREFS.get('currency_symbol','₹')}{req_monthly:,.2f}/month")
            except Exception:
                pass

        text.append("")
        text.append("Suggestion: Automate this SIP via your bank or mutual fund platform. Start small and increase regularly.")

        self.result_box.SetValue("\n".join(text))

# -------------------------
# WINDOW 1: Entry Panel (Input + Table)
# -------------------------
class EntryWindow(wx.Frame):
    def __init__(self):
        super().__init__(None, title="📝 Expense Entry", size=(900, 620))
        global WINDOWS
        WINDOWS['entry'] = self

        self.Bind(wx.EVT_CLOSE, self.on_close)

        self.SetBackgroundColour(BG_MAIN)
        outer = wx.BoxSizer(wx.VERTICAL)

        card = wx.Panel(self)
        card.SetBackgroundColour(BG_PANEL)
        sizer = wx.BoxSizer(wx.VERTICAL)

        header = wx.StaticText(card, label="Expense Entry & List")
        header.SetFont(wx.Font(18, wx.DEFAULT, wx.NORMAL, wx.BOLD))
        sizer.Add(header, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 10)

        subtitle = wx.StaticText(card, label="Add daily expenses and track income for each month.")
        subtitle.SetFont(wx.Font(10, wx.DEFAULT, wx.NORMAL, wx.LIGHT))
        sizer.Add(subtitle, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 2)

        # Input grid
        grid = wx.FlexGridSizer(2, 7, 8, 8)
        grid.AddGrowableCol(1)
        grid.AddGrowableCol(3)
        grid.AddGrowableCol(5)

        # Row 1: Date, Category, Amount, Currency
        lbl_date = wx.StaticText(card, label="Date:")
        grid.Add(lbl_date, 0, wx.ALIGN_CENTER_VERTICAL)
        self.date_picker = wx.adv.DatePickerCtrl(card, style=wx.adv.DP_DROPDOWN | wx.adv.DP_SHOWCENTURY)
        self.date_picker.SetBackgroundColour(BG_INPUT)
        grid.Add(self.date_picker, 0, wx.EXPAND)

        lbl_cat = wx.StaticText(card, label="Category:")
        grid.Add(lbl_cat, 0, wx.ALIGN_CENTER_VERTICAL)
        self.category_input = wx.TextCtrl(card)
        self.category_input.SetBackgroundColour(BG_INPUT)
        grid.Add(self.category_input, 0, wx.EXPAND)

        lbl_amt = wx.StaticText(card, label="Amount:")
        grid.Add(lbl_amt, 0, wx.ALIGN_CENTER_VERTICAL)
        self.amount_input = wx.TextCtrl(card)
        self.amount_input.SetBackgroundColour(BG_INPUT)
        grid.Add(self.amount_input, 0, wx.EXPAND)

        # Currency dropdown
        self.currency_choice = wx.Choice(card, choices=CURRENCY_OPTIONS)
        # Default selection based on preferences, fallback to first
        default_symbol = PREFS.get("currency_symbol", "₹")
        if default_symbol in CURRENCY_OPTIONS:
            self.currency_choice.SetSelection(CURRENCY_OPTIONS.index(default_symbol))
        else:
            self.currency_choice.SetSelection(0)
        self.currency_choice.SetBackgroundColour(BG_INPUT)
        grid.Add(self.currency_choice, 0, wx.EXPAND)

        # Row 2: Note + spacers
        lbl_note = wx.StaticText(card, label="Note (optional):")
        grid.Add(lbl_note, 0, wx.ALIGN_CENTER_VERTICAL)
        self.note_input = wx.TextCtrl(card)
        self.note_input.SetBackgroundColour(BG_INPUT)
        grid.Add(self.note_input, 0, wx.EXPAND)

        # Fill remaining cells of row 2 with spacers
        grid.Add((0, 0))
        grid.Add((0, 0))
        grid.Add((0, 0))
        grid.Add((0, 0))
        grid.Add((0, 0))

        sizer.Add(grid, 0, wx.ALL | wx.EXPAND, 12)

        # Buttons
        btn_row = wx.BoxSizer(wx.HORIZONTAL)
        add_btn = wx.Button(card, label="➕ Add Expense")
        add_btn.SetBackgroundColour(BTN_PRIMARY)
        add_btn.SetForegroundColour(wx.WHITE)
        add_btn.Bind(wx.EVT_BUTTON, self.on_add_expense)
        btn_row.Add(add_btn, 0, wx.ALL, 5)

        clear_btn = wx.Button(card, label="✖ Clear")
        clear_btn.SetBackgroundColour(wx.Colour(255, 118, 117))
        clear_btn.SetForegroundColour(wx.WHITE)
        clear_btn.Bind(wx.EVT_BUTTON, lambda e: self.clear_inputs())
        btn_row.Add(clear_btn, 0, wx.ALL, 5)
        sizer.Add(btn_row, 0, wx.ALIGN_LEFT | wx.LEFT, 12)

        # Income control
        inc_box = wx.BoxSizer(wx.HORIZONTAL)
        inc_label = wx.StaticText(card, label="Monthly income (for selected month): ")
        inc_box.Add(inc_label, 0, wx.ALIGN_CENTER_VERTICAL)
        self.income_input = wx.TextCtrl(card, size=(160,-1))
        self.income_input.SetBackgroundColour(BG_INPUT)
        inc_box.Add(self.income_input, 0, wx.ALL, 5)
        set_inc_btn = wx.Button(card, label="💾 Save Income")
        set_inc_btn.SetBackgroundColour(BTN_SECONDARY)
        set_inc_btn.SetForegroundColour(wx.WHITE)
        set_inc_btn.Bind(wx.EVT_BUTTON, self.on_set_income)
        inc_box.Add(set_inc_btn, 0, wx.ALL, 5)
        sizer.Add(inc_box, 0, wx.LEFT | wx.RIGHT | wx.TOP, 12)

        # Expense table
        self.table = wx.ListCtrl(card, style=wx.LC_REPORT | wx.BORDER_SUNKEN)
        self.table.InsertColumn(0, "Date", width=110)
        self.table.InsertColumn(1, "Category", width=180)
        self.table.InsertColumn(2, "Amount", width=130)
        self.table.InsertColumn(3, "Note", width=280)
        self.table.Bind(wx.EVT_LIST_ITEM_ACTIVATED, self.on_edit_expense)
        sizer.Add(self.table, 1, wx.EXPAND | wx.ALL, 12)

        card.SetSizer(sizer)
        outer.Add(card, 1, wx.ALL | wx.EXPAND, 10)
        self.SetSizer(outer)

        self.load_table()
        self.Show()

    def on_close(self, event):
        global WINDOWS
        if WINDOWS.get('entry') is self:
            del WINDOWS['entry']
        self.Destroy()

    def clear_inputs(self):
        self.category_input.SetValue("")
        self.amount_input.SetValue("")
        self.note_input.SetValue("")

    def on_add_expense(self, event):
        d = self.date_picker.GetValue()
        date = f"{d.GetYear():04d}-{d.GetMonth()+1:02d}-{d.GetDay():02d}"
        cat = self.category_input.GetValue().strip()
        amt_txt = self.amount_input.GetValue().strip()
        note = self.note_input.GetValue().strip()

        if not cat or not amt_txt:
            wx.MessageBox("Please enter Category and Amount.", "Missing Fields", style=wx.OK | wx.ICON_ERROR)
            return
        try:
            amt = float(amt_txt)
            if math.isfinite(amt) is False:
                raise ValueError
        except Exception:
            wx.MessageBox("Please enter a valid numeric amount.", "Invalid Amount", style=wx.OK | wx.ICON_ERROR)
            return

        currency = self.currency_choice.GetString(self.currency_choice.GetSelection())
        amount_text = f"{currency}{amt:.2f}"

        add_expense(date, cat, amount_text, note)
        self.clear_inputs()
        self.load_table()
        refresh_all_windows()

    def load_table(self):
        self.table.DeleteAllItems()
        for row in read_expenses():
            idx = self.table.InsertItem(self.table.GetItemCount(), row["Date"])
            self.table.SetItem(idx, 1, row["Category"])
            self.table.SetItem(idx, 2, row["Amount"])
            self.table.SetItem(idx, 3, row.get("Note",""))

    def on_edit_expense(self, event):
        idx = event.GetIndex()
        date = self.table.GetItemText(idx, 0)
        cat = self.table.GetItemText(idx, 1)
        amt = self.table.GetItemText(idx, 2)
        note = self.table.GetItemText(idx, 3)
        dlg = EditExpenseDialog(self, idx, date, cat, amt, note)
        dlg.ShowModal()
        dlg.Destroy()
        self.load_table()
        refresh_all_windows()

    def on_set_income(self, event):
        d = self.date_picker.GetValue()
        month_key = f"{d.GetYear():04d}-{d.GetMonth()+1:02d}"
        amt_txt = self.income_input.GetValue().strip()
        if not amt_txt:
            wx.MessageBox("Please enter an income amount.", "Missing", style=wx.OK | wx.ICON_ERROR)
            return
        try:
            amt = float(amt_txt)
        except Exception:
            wx.MessageBox("Income must be numeric.", "Invalid", style=wx.OK | wx.ICON_ERROR)
            return
        set_income_for_month(month_key, amt)
        wx.MessageBox(f"Income for {month_key} saved: {PREFS.get('currency_symbol','₹')}{amt:.2f}", "Saved", style=wx.OK | wx.ICON_INFORMATION)
        refresh_all_windows()

# -------------------------
# WINDOW 2: Charts Window
# -------------------------
class ChartsWindow(wx.Frame):
    def __init__(self):
        super().__init__(None, title="📊 Analytics Dashboard", size=(900, 720))
        global WINDOWS
        WINDOWS['charts'] = self

        self.Bind(wx.EVT_CLOSE, self.on_close)
        self.SetBackgroundColour(BG_MAIN)
        
        outer = wx.BoxSizer(wx.VERTICAL)
        card = wx.Panel(self)
        card.SetBackgroundColour(BG_PANEL)
        sizer = wx.BoxSizer(wx.VERTICAL)

        chart_title = wx.StaticText(card, label="Visual Insights")
        chart_title.SetFont(wx.Font(18, wx.DEFAULT, wx.NORMAL, wx.BOLD))
        sizer.Add(chart_title, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 8)

        subtitle = wx.StaticText(card, label="Understand your spending patterns across time and categories.")
        subtitle.SetFont(wx.Font(10, wx.DEFAULT, wx.NORMAL, wx.LIGHT))
        sizer.Add(subtitle, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 2)

        self.figure = Figure(figsize=(8,6), dpi=100)
        self.canvas = FigureCanvas(card, -1, self.figure)
        sizer.Add(self.canvas, 1, wx.EXPAND | wx.ALL, 8)

        # Control buttons
        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        refresh_btn = wx.Button(card, label="🔄 Refresh Charts")
        refresh_btn.SetBackgroundColour(BTN_PRIMARY)
        refresh_btn.SetForegroundColour(wx.WHITE)
        refresh_btn.Bind(wx.EVT_BUTTON, lambda e: self.update_charts())
        btn_sizer.Add(refresh_btn, 0, wx.ALL, 5)
        
        open_entry_btn = wx.Button(card, label="📝 Open Entry")
        open_entry_btn.SetBackgroundColour(BG_INPUT)
        open_entry_btn.Bind(wx.EVT_BUTTON, lambda e: self.open_window('entry'))
        btn_sizer.Add(open_entry_btn, 0, wx.ALL, 5)
        
        open_suggestions_btn = wx.Button(card, label="💡 Open Suggestions")
        open_suggestions_btn.SetBackgroundColour(BG_INPUT)
        open_suggestions_btn.Bind(wx.EVT_BUTTON, lambda e: self.open_window('suggestions'))
        btn_sizer.Add(open_suggestions_btn, 0, wx.ALL, 5)
        
        sizer.Add(btn_sizer, 0, wx.ALIGN_CENTER | wx.ALL, 8)

        card.SetSizer(sizer)
        outer.Add(card, 1, wx.ALL | wx.EXPAND, 10)
        self.SetSizer(outer)
        
        self.update_charts()
        self.Show()

    def on_close(self, event):
        global WINDOWS
        if WINDOWS.get('charts') is self:
            del WINDOWS['charts']
        self.Destroy()

    def open_window(self, win_type):
        if win_type not in WINDOWS:
            if win_type == 'entry':
                EntryWindow()
            elif win_type == 'suggestions':
                SuggestionsWindow()
        else:
            WINDOWS[win_type].Show()
            WINDOWS[win_type].Raise()

    def update_charts(self):
        expenses = read_expenses()
        incomes = read_incomes()

        category_totals = defaultdict(float)
        month_totals = defaultdict(float)
        month_category = defaultdict(lambda: defaultdict(float))

        for r in expenses:
            amt = safe_float(r["Amount"]) 
            cat = r["Category"] or "Uncategorized"
            mkey = date_to_month_key(r["Date"]) or datetime.now().strftime("%Y-%m")
            category_totals[cat] += amt
            month_totals[mkey] += amt
            month_category[mkey][cat] += amt

        months = sorted(set(list(month_totals.keys()) + list(incomes.keys())))
        if not months:
            months = [datetime.now().strftime("%Y-%m")]

        self.figure.clear()
        ax_income = self.figure.add_subplot(2,2,1)
        ax_cat = self.figure.add_subplot(2,2,3)
        ax_pie = self.figure.add_subplot(2,2,2)

        inc_vals = [incomes.get(m, 0.0) for m in months]
        exp_vals = [month_totals.get(m, 0.0) for m in months]
        x = range(len(months))
        w = 0.35
        ax_income.bar([i - w/2 for i in x], inc_vals, w, label="Income")
        ax_income.bar([i + w/2 for i in x], exp_vals, w, label="Expenditure")
        ax_income.set_xticks(list(x))
        ax_income.set_xticklabels(months, rotation=45, fontsize=8)
        ax_income.set_title("Monthly Income vs Expenditure")
        ax_income.legend(fontsize=8)

        sorted_categories = sorted(category_totals.items(), key=lambda kv: kv[1], reverse=True)
        top_n = 10
        cats = [c for c,v in sorted_categories[:top_n]] or ["No data"]
        vals = [v for c,v in sorted_categories[:top_n]] or [0]
        ax_cat.bar(cats, vals, color="#74b9ff")
        ax_cat.set_title("Top Categories (All time)")
        ax_cat.tick_params(axis='x', rotation=45, labelsize=8)

        pie_top = sorted_categories[:6]
        labels = [c for c,v in pie_top]
        sizes = [v for c,v in pie_top]
        if len(sorted_categories) > 6:
            others_sum = sum(v for c,v in sorted_categories[6:])
            labels.append("Others")
            sizes.append(others_sum)
        if sum(sizes) == 0:
            labels = ["No data"]; sizes = [1]
        ax_pie.pie(sizes, labels=labels, autopct="%1.1f%%", startangle=140)
        ax_pie.set_title("Category Share (Top)")

        self.figure.tight_layout()
        self.canvas.draw()

# -------------------------
# WINDOW 3: Suggestions Window
# -------------------------
class SuggestionsWindow(wx.Frame):
    def __init__(self):
        super().__init__(None, title="💡 Suggestions & Actions", size=(600, 620))
        global WINDOWS
        WINDOWS['suggestions'] = self

        self.Bind(wx.EVT_CLOSE, self.on_close)
        self.SetBackgroundColour(BG_MAIN)
        
        outer = wx.BoxSizer(wx.VERTICAL)
        card = wx.Panel(self)
        card.SetBackgroundColour(BG_PANEL)
        sizer = wx.BoxSizer(wx.VERTICAL)
        
        self.title = wx.StaticText(card, label="Suggestions & Actions")
        self.title.SetFont(wx.Font(18, wx.DEFAULT, wx.NORMAL, wx.BOLD))
        sizer.Add(self.title, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 8)

        subtitle = wx.StaticText(card, label="High-level summary for the selected month plus quick tools.")
        subtitle.SetFont(wx.Font(10, wx.DEFAULT, wx.NORMAL, wx.LIGHT))
        sizer.Add(subtitle, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 2)

        self.info = wx.TextCtrl(card, style=wx.TE_MULTILINE | wx.TE_READONLY)
        self.info.SetMinSize((520, 320))
        self.info.SetBackgroundColour(wx.Colour(250, 252, 255))
        sizer.Add(self.info, 1, wx.EXPAND | wx.ALL, 8)

        # Action buttons
        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        self.sip_btn = wx.Button(card, label="🚀 Start SIP")
        self.sip_btn.SetBackgroundColour(BTN_SECONDARY)
        self.sip_btn.SetForegroundColour(wx.WHITE)
        self.sip_btn.Bind(wx.EVT_BUTTON, self.on_start_sip)
        btn_sizer.Add(self.sip_btn, 0, wx.ALL, 5)

        self.pref_btn = wx.Button(card, label="⚙ Preferences")
        self.pref_btn.SetBackgroundColour(BG_INPUT)
        self.pref_btn.Bind(wx.EVT_BUTTON, self.on_preferences)
        btn_sizer.Add(self.pref_btn, 0, wx.ALL, 5)

        self.pdf_btn = wx.Button(card, label="📄 PDF Report")
        self.pdf_btn.SetBackgroundColour(BG_INPUT)
        self.pdf_btn.Bind(wx.EVT_BUTTON, self.on_generate_pdf)
        btn_sizer.Add(self.pdf_btn, 0, wx.ALL, 5)

        self.csv_btn = wx.Button(card, label="📂 Open CSV")
        self.csv_btn.SetBackgroundColour(BG_INPUT)
        self.csv_btn.Bind(wx.EVT_BUTTON, self.on_open_csv)
        btn_sizer.Add(self.csv_btn, 0, wx.ALL, 5)

        sizer.Add(btn_sizer, 0, wx.ALIGN_CENTER | wx.TOP, 4)

        # Navigation buttons
        nav_sizer = wx.BoxSizer(wx.HORIZONTAL)
        open_entry_btn = wx.Button(card, label="📝 Open Entry")
        open_entry_btn.SetBackgroundColour(BG_INPUT)
        open_entry_btn.Bind(wx.EVT_BUTTON, lambda e: self.open_window('entry'))
        nav_sizer.Add(open_entry_btn, 0, wx.ALL, 5)
        
        open_charts_btn = wx.Button(card, label="📊 Open Charts")
        open_charts_btn.SetBackgroundColour(BG_INPUT)
        open_charts_btn.Bind(wx.EVT_BUTTON, lambda e: self.open_window('charts'))
        nav_sizer.Add(open_charts_btn, 0, wx.ALL, 5)
        sizer.Add(nav_sizer, 0, wx.ALIGN_CENTER | wx.TOP | wx.BOTTOM, 8)

        card.SetSizer(sizer)
        outer.Add(card, 1, wx.ALL | wx.EXPAND, 10)
        self.SetSizer(outer)

        self.update_text(self.get_suggestion_text())
        self.Show()

    def on_close(self, event):
        global WINDOWS
        if WINDOWS.get('suggestions') is self:
            del WINDOWS['suggestions']
        self.Destroy()

    def open_window(self, win_type):
        if win_type not in WINDOWS:
            if win_type == 'entry':
                EntryWindow()
            elif win_type == 'charts':
                ChartsWindow()
        else:
            WINDOWS[win_type].Show()
            WINDOWS[win_type].Raise()

    def get_suggestion_text(self):
        current_month = datetime.now().strftime("%Y-%m")
        if 'entry' in WINDOWS:
            try:
                d = WINDOWS['entry'].date_picker.GetValue()
                current_month = f"{d.GetYear():04d}-{d.GetMonth()+1:02d}"
            except Exception:
                pass

        incomes = read_incomes()
        expenses = read_expenses()
        month_totals = defaultdict(float)
        
        for r in expenses:
            mkey = date_to_month_key(r["Date"]) or current_month
            month_totals[mkey] += safe_float(r["Amount"])
        
        inc = incomes.get(current_month, 0.0)
        exp = month_totals.get(current_month, 0.0)
        diff = inc - exp

        suggestion_text = f"📅 Month: {current_month}\n"
        suggestion_text += f"💰 Income: {PREFS.get('currency_symbol','₹')}{inc:.2f}\n"
        suggestion_text += f"💸 Expenditure: {PREFS.get('currency_symbol','₹')}{exp:.2f}\n"
        suggestion_text += f"⚖  Balance: {PREFS.get('currency_symbol','₹')}{diff:.2f}\n\n"

        if inc == 0:
            suggestion_text += "⚠ No income set for this month.\nSet monthly income to enable better insights.\n\n"

        mcat = defaultdict(float)
        for r in expenses:
            if date_to_month_key(r["Date"]) == current_month:
                mcat[r["Category"] or "Uncategorized"] += safe_float(r["Amount"])
        
        if mcat:
            suggestion_text += "📋 This month's top categories:\n"
            for c,v in sorted(mcat.items(), key=lambda kv: kv[1], reverse=True)[:6]:
                suggestion_text += f"  • {c}: {PREFS.get('currency_symbol','₹')}{v:.2f}\n"

        # Overspending alert
        if inc > 0 and exp > inc:
            over_amount = exp - inc
            suggestion_text += (
                f"\n🚨 You have spent too much!\n"
                f"Expenditure exceeds income by {PREFS.get('currency_symbol','₹')}{over_amount:.2f}\n"
                "💡 Review top spending categories above and cut back where possible.\n\n"
            )
        elif diff >= 1000:
            suggestion_text += (
                f"\n✅ Great! {PREFS.get('currency_symbol','₹')}{diff:.2f} available to save.\n"
                "💡 Consider automated SIP investments.\n\n"
            )

        suggestion_text += "🔧 Quick Actions:\n"
        suggestion_text += "• Click 'Start SIP' for investment calculator\n"
        suggestion_text += "• Set income in Entry window\n"
        suggestion_text += "• Generate PDF reports monthly\n"
        return suggestion_text

    def update_text(self, text=None):
        if text is None:
            text = self.get_suggestion_text()
        wx.CallAfter(self.info.SetValue, text)

    def on_start_sip(self, event):
        dlg = SIPDialog(self)
        dlg.ShowModal()
        dlg.Destroy()

    def on_preferences(self, event):
        dlg = PreferencesDialog(self)
        dlg.ShowModal()
        dlg.Destroy()
        global PREFS
        PREFS = load_preferences()
        refresh_all_windows()

    def on_generate_pdf(self, event):
        generate_pdf_report_dialog(self)

    def on_open_csv(self, event):
        if not os.path.exists(FILENAME):
            wx.MessageBox(
                f"File '{FILENAME}' not found.\nAdd at least one expense to create it.",
                "CSV not found",
                style=wx.OK | wx.ICON_INFORMATION
            )
            return
        try:
            if os.name == "nt":  # Windows
                os.startfile(FILENAME)
            elif sys.platform == "darwin":  # macOS
                subprocess.call(["open", FILENAME])
            else:  # Linux / others
                subprocess.call(["xdg-open", FILENAME])
        except Exception as e:
            wx.MessageBox(
                f"Could not open '{FILENAME}'.\nError: {e}",
                "Error opening CSV",
                style=wx.OK | wx.ICON_ERROR
            )

# -------------------------
# PDF Report Generator (Modal Dialog)
# -------------------------
def generate_pdf_report_dialog(parent):
    print("generate_pdf_report_dialog called")

    month_key = datetime.now().strftime("%Y-%m")

    if 'entry' in WINDOWS:
        try:
            wx_date = WINDOWS['entry'].date_picker.GetValue()
            iso = wx_date.FormatISODate()
            if iso:
                month_key = iso[:7]
        except Exception:
            pass

    expenses = read_expenses()
    rows = [r for r in expenses if date_to_month_key(r["Date"]) == month_key]
    if not rows:
        wx.MessageBox(f"No expenses found for {month_key}", "No Data", style=wx.OK | wx.ICON_INFORMATION)
        return

    with wx.FileDialog(parent, "Save PDF report",
                       wildcard="PDF files (.pdf)|.pdf",
                       style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT) as fd:
        if fd.ShowModal() == wx.ID_CANCEL:
            return
        path = fd.GetPath()
        if not path.lower().endswith(".pdf"):
            path += ".pdf"

    cat_totals = defaultdict(float)
    for r in rows:
        cat_totals[r["Category"] or "Uncategorized"] += safe_float(r["Amount"]) 

    fig = Figure(figsize=(6,4), dpi=150)
    ax = fig.add_subplot(111)
    labels = list(cat_totals.keys())[:6]
    sizes = [cat_totals[k] for k in labels]
    if sum(sizes) == 0:
        labels = ["No data"]; sizes = [1]
    ax.pie(sizes, labels=labels, autopct="%1.1f%%")
    ax.set_title(f"Category share — {month_key}")

    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    buf.seek(0)

    if REPORTLAB_AVAILABLE:
        try:
            c = pdfcanvas.Canvas(path, pagesize=A4)
            width, height = A4
            c.setFont("Helvetica-Bold", 14)
            c.drawString(40, height-60, f"Monthly Expense Report — {month_key}")

            total = sum(safe_float(r["Amount"]) for r in rows)
            incomes = read_incomes()
            income = incomes.get(month_key, 0.0)
            c.setFont("Helvetica", 10)
            c.drawString(40, height-90, f"Income: {PREFS.get('currency_symbol','₹')}{income:.2f}")
            c.drawString(200, height-90, f"Expenditure: {PREFS.get('currency_symbol','₹')}{total:.2f}")

            img = ImageReader(buf)
            c.drawImage(img, 40, height-420, width=500, height=300, preserveAspectRatio=True, anchor='sw')

            c.drawString(40, height-440, "Top expenses:")
            sorted_by_amt = sorted(cat_totals.items(), key=lambda kv: kv[1], reverse=True)
            for i, (cat, amt) in enumerate(sorted_by_amt[:10], start=1):
                c.drawString(50, height-460-20*i, f"{i}. {cat}: {PREFS.get('currency_symbol','₹')}{amt:.2f}")

            c.showPage()
            c.save()
            wx.MessageBox(f"PDF report saved to {path}", "Saved", style=wx.OK | wx.ICON_INFORMATION)
        except Exception as e:
            wx.MessageBox(f"Failed to create PDF: {e}", "Error", style=wx.OK | wx.ICON_ERROR)
    else:
        try:
            png_path = path.replace('.pdf', '.png')
            with open(png_path, 'wb') as f:
                f.write(buf.getvalue())
            csv_path = path.replace('.pdf', '.csv')
            with open(csv_path, 'w', newline='') as f:
                writer = csv.writer(f)
                writer.writerow(["Date","Category","Amount","Note"])
                for r in rows:
                    writer.writerow([r["Date"], r["Category"], r["Amount"], r.get("Note","")])
            wx.MessageBox(f"Saved PNG+CSV:\n{ png_path }\n{ csv_path }\n\nInstall 'reportlab' for PDF.", "Saved", style=wx.OK | wx.ICON_INFORMATION)
        except Exception as e:
            wx.MessageBox(f"Failed to save: {e}", "Error", style=wx.OK | wx.ICON_ERROR)

# -------------------------
# Edit & Preferences Dialogs
# -------------------------
class EditExpenseDialog(wx.Dialog):
    def __init__(self, parent, index, date, category, amount, note):
        super().__init__(parent, title="Edit Expense", size=(480,360))
        self.parent = parent
        self.index = index

        self.SetBackgroundColour(BG_PANEL)
        s = wx.BoxSizer(wx.VERTICAL)

        s.Add(wx.StaticText(self, label="Date (YYYY-MM-DD):"), 0, wx.ALL, 6)
        self.date_txt = wx.TextCtrl(self, value=date)
        self.date_txt.SetBackgroundColour(BG_INPUT)
        s.Add(self.date_txt, 0, wx.EXPAND | wx.ALL, 6)

        s.Add(wx.StaticText(self, label="Category:"), 0, wx.ALL, 6)
        self.cat_txt = wx.TextCtrl(self, value=category)
        self.cat_txt.SetBackgroundColour(BG_INPUT)
        s.Add(self.cat_txt, 0, wx.EXPAND | wx.ALL, 6)

        s.Add(wx.StaticText(self, label="Amount (can include currency symbol):"), 0, wx.ALL, 6)
        self.amt_txt = wx.TextCtrl(self, value=amount)
        self.amt_txt.SetBackgroundColour(BG_INPUT)
        s.Add(self.amt_txt, 0, wx.EXPAND | wx.ALL, 6)

        s.Add(wx.StaticText(self, label="Note (optional):"), 0, wx.ALL, 6)
        self.note_txt = wx.TextCtrl(self, value=note)
        self.note_txt.SetBackgroundColour(BG_INPUT)
        s.Add(self.note_txt, 0, wx.EXPAND | wx.ALL, 6)

        btns = wx.BoxSizer(wx.HORIZONTAL)
        save = wx.Button(self, label="Save")
        save.SetBackgroundColour(BTN_PRIMARY)
        save.SetForegroundColour(wx.WHITE)
        save.Bind(wx.EVT_BUTTON, self.on_save)
        btns.Add(save, 0, wx.ALL, 6)

        delete = wx.Button(self, label="Delete")
        delete.SetBackgroundColour(wx.Colour(255, 118, 117))
        delete.SetForegroundColour(wx.WHITE)
        delete.Bind(wx.EVT_BUTTON, self.on_delete)
        btns.Add(delete, 0, wx.ALL, 6)

        cancel = wx.Button(self, label="Cancel")
        cancel.Bind(wx.EVT_BUTTON, lambda e: self.Close())
        btns.Add(cancel, 0, wx.ALL, 6)

        s.add(btns, 0, wx.ALIGN_CENTER)
        self.SetSizer(s)

    def on_save(self, event):
        new_date = self.date_txt.GetValue().strip()
        new_cat = self.cat_txt.GetValue().strip()
        new_amt = self.amt_txt.GetValue().strip()
        new_note = self.note_txt.GetValue().strip()
        if not new_cat or not new_amt or not new_date:
            wx.MessageBox("Date, Category and Amount are required.", "Missing", style=wx.OK | wx.ICON_ERROR)
            return
        try:
            _ = safe_float(new_amt)
        except Exception:
            wx.MessageBox("Please enter a valid amount (optionally with currency symbol).", "Invalid", style=wx.OK | wx.ICON_ERROR)
            return
        rows = read_expenses()
        if 0 <= self.index < len(rows):
            rows[self.index]["Date"] = new_date
            rows[self.index]["Category"] = new_cat
            rows[self.index]["Amount"] = new_amt
            rows[self.index]["Note"] = new_note
            write_expenses(rows)
            wx.MessageBox("Saved.", "Saved", style=wx.OK | wx.ICON_INFORMATION)
            self.Close()
            refresh_all_windows()
        else:
            wx.MessageBox("Row unavailable.", "Error", style=wx.OK | wx.ICON_ERROR)

    def on_delete(self, event):
        confirm = wx.MessageBox("Delete this expense? This action cannot be undone.", "Confirm", style=wx.YES_NO | wx.ICON_WARNING)
        if confirm == wx.YES:
            rows = read_expenses()
            if 0 <= self.index < len(rows):
                del rows[self.index]
                write_expenses(rows)
                wx.MessageBox("Deleted.", "Deleted", style=wx.OK | wx.ICON_INFORMATION)
                self.Close()
                refresh_all_windows()
            else:
                wx.MessageBox("Row unavailable.", "Error", style=wx.OK | wx.ICON_ERROR)

class PreferencesDialog(wx.Dialog):
    def __init__(self, parent):
        super().__init__(parent, title="Preferences", size=(420,260))
        self.SetBackgroundColour(BG_PANEL)
        s = wx.BoxSizer(wx.VERTICAL)
        prefs = load_preferences()

        s.Add(wx.StaticText(self, label="Preferred currency symbol:"), 0, wx.ALL, 6)
        self.currency_txt = wx.TextCtrl(self, value=prefs.get('currency_symbol','₹'))
        self.currency_txt.SetBackgroundColour(BG_INPUT)
        s.Add(self.currency_txt, 0, wx.EXPAND | wx.ALL, 6)

        s.Add(wx.StaticText(self, label="Default monthly budget target (numeric):"), 0, wx.ALL, 6)
        self.budget_txt = wx.TextCtrl(self, value=str(prefs.get('default_monthly_budget',0.0)))
        self.budget_txt.SetBackgroundColour(BG_INPUT)
        s.Add(self.budget_txt, 0, wx.EXPAND | wx.ALL, 6)

        btns = wx.BoxSizer(wx.HORIZONTAL)
        save = wx.Button(self, label="Save Preferences")
        save.SetBackgroundColour(BTN_PRIMARY)
        save.SetForegroundColour(wx.WHITE)
        save.Bind(wx.EVT_BUTTON, self.on_save)
        btns.Add(save, 0, wx.ALL, 6)
        cancel = wx.Button(self, label="Cancel")
        cancel.Bind(wx.EVT_BUTTON, lambda e: self.Close())
        btns.Add(cancel, 0, wx.ALL, 6)

        s.Add(btns, 0, wx.ALIGN_CENTER)
        self.SetSizer(s)

    def on_save(self, event):
        sym = self.currency_txt.GetValue().strip() or '₹'
        try:
            bud = float(self.budget_txt.GetValue())
        except Exception:
            wx.MessageBox('Budget must be numeric.', 'Invalid', style=wx.OK | wx.ICON_ERROR)
            return
        prefs = {"currency_symbol": sym, "default_monthly_budget": bud}
        save_preferences(prefs)
        wx.MessageBox('Preferences saved.', 'Saved', style=wx.OK | wx.ICON_INFORMATION)
        self.Close()
        refresh_all_windows()

# -------------------------
# Launcher Window
# -------------------------
class LauncherWindow(wx.Frame):
    def __init__(self):
        super().__init__(None, title="Expense Tracker Launcher", size=(520, 320))
        self.SetBackgroundColour(BG_MAIN)
        sizer = wx.BoxSizer(wx.VERTICAL)
        
        title = wx.StaticText(self, label="🚀 Expense Tracker")
        title.SetFont(wx.Font(22, wx.DEFAULT, wx.NORMAL, wx.BOLD))
        sizer.Add(title, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 20)
        
        desc = wx.StaticText(self, label="Choose a panel to get started:")
        desc.SetFont(wx.Font(11, wx.DEFAULT, wx.NORMAL, wx.NORMAL))
        sizer.Add(desc, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)
        
        btn_sizer = wx.BoxSizer(wx.VERTICAL)
        
        entry_btn = wx.Button(self, label="📝 Entry Window")
        entry_btn.SetBackgroundColour(BG_PANEL)
        entry_btn.Bind(wx.EVT_BUTTON, lambda e: self.open_window('entry'))
        btn_sizer.Add(entry_btn, 0, wx.EXPAND | wx.ALL, 8)
        
        charts_btn = wx.Button(self, label="📊 Charts Window")
        charts_btn.SetBackgroundColour(BG_PANEL)
        charts_btn.Bind(wx.EVT_BUTTON, lambda e: self.open_window('charts'))
        btn_sizer.Add(charts_btn, 0, wx.EXPAND | wx.ALL, 8)
        
        sug_btn = wx.Button(self, label="💡 Suggestions Window")
        sug_btn.SetBackgroundColour(BG_PANEL)
        sug_btn.Bind(wx.EVT_BUTTON, lambda e: self.open_window('suggestions'))
        btn_sizer.Add(sug_btn, 0, wx.EXPAND | wx.ALL, 8)
        
        sizer.Add(btn_sizer, 1, wx.EXPAND | wx.LEFT | wx.RIGHT, 60)
        self.SetSizer(sizer)
        self.Centre()
        self.Show()

    def open_window(self, win_type):
        if win_type not in WINDOWS:
            if win_type == 'entry':
                EntryWindow()
            elif win_type == 'charts':
                ChartsWindow()
            elif win_type == 'suggestions':
                SuggestionsWindow()
        else:
            WINDOWS[win_type].Show()
            WINDOWS[win_type].Raise()

# -------------------------
# Run app
# -------------------------
if __name__ == '__main__':
    app = wx.App(False)
    LauncherWindow()
    app.MainLoop()
