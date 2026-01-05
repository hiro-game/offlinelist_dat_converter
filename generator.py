import csv
import json
import os
import sys
from datetime import datetime
import xml.etree.ElementTree as ET  # CSV→XML の生成に使用
from xml.dom import minidom
import tkinter as tk
from tkinter import filedialog, messagebox

from bs4 import BeautifulSoup  # XML→CSV 用（寛容パーサー）


# -------------------------------
# パス関連（exe 化対応）
# -------------------------------
def resource_path(relative_path: str) -> str:
    if hasattr(sys, "_MEIPASS"):
        base_path = sys._MEIPASS  # PyInstaller 用
    else:
        base_path = os.path.dirname(os.path.abspath(sys.argv[0]))
    return os.path.join(base_path, relative_path)


# -------------------------------
# 設定ファイル読み込み
# -------------------------------
def load_config():
    config_path = resource_path("config.json")
    if not os.path.exists(config_path):
        messagebox.showerror("エラー", f"config.json が見つかりません。\n場所: {config_path}")
        raise FileNotFoundError("config.json not found")

    with open(config_path, "r", encoding="utf-8") as f:
        return json.load(f)


# -------------------------------
# XML 整形ヘルパー（CSV→XML用）
# -------------------------------
def prettify_element(elem: ET.Element) -> str:
    rough_string = ET.tostring(elem, encoding="unicode")
    reparsed = minidom.parseString(rough_string)
    pretty = reparsed.toprettyxml(indent="  ")
    lines = pretty.splitlines()
    lines = [line for line in lines if line.strip()]
    if lines and lines[0].startswith("<?xml"):
        lines = lines[1:]
    return "\n".join(lines)


# -------------------------------
# CSV / XML 判定
# -------------------------------
def is_csv(path: str) -> bool:
    return path.lower().endswith(".csv")


def is_xml(path: str) -> bool:
    return path.lower().endswith(".xml")


# -------------------------------
# games 部分生成（CSV → XML）
# -------------------------------
def build_games_from_csv(csv_path: str, extension_with_dot: str) -> ET.Element:
    games_root = ET.Element("games")

    with open(csv_path, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)

        for row in reader:
            game = ET.SubElement(games_root, "game")

            for key, value in row.items():
                if key == "romCRC":
                    files = ET.SubElement(game, "files")
                    crc = ET.SubElement(files, "romCRC")
                    crc.text = value
                    crc.set("extension", row.get("extension", extension_with_dot or ""))
                    continue

                if key == "extension":
                    continue

                elem = ET.SubElement(game, key)
                elem.text = " " if not value else str(value)

    return games_root


# -------------------------------
# configuration テンプレート（共通）
# -------------------------------
CONFIGURATION_TEMPLATE = """  <configuration>
    <datName>{datName}</datName>
    <imFolder>{imFolder}</imFolder>
    <datVersion>{datVersion}</datVersion>
    <system>{system}</system>
    <screenshotsWidth>{screenshotsWidth}</screenshotsWidth>
    <screenshotsHeight>{screenshotsHeight}</screenshotsHeight>
    <infos>
      <title visible="false" inNamingOption="true" default="false" />
      <publisher visible="true" inNamingOption="true" default="true" />
      <sourceRom visible="true" inNamingOption="true" default="true" />
      <im1CRC visible="true" inNamingOption="true" default="true" />
      <saveType visible="false" inNamingOption="true" default="false" />
      <im2CRC visible="true" inNamingOption="true" default="true" />
      <comment visible="true" inNamingOption="true" default="true" />
      <location visible="true" inNamingOption="true" default="false"/>
      <languages visible="true" inNamingOption="true" default="true" />
      <releaseNumber visible="true" inNamingOption="true" default="true" />
      <romSize visible="true" inNamingOption="true" default="true"/>
      <romCRC visible="true" inNamingOption="true" default="true" />
      <imageNumber visible="false" inNamingOption="false" default="false" />
    </infos>
    <canOpen><extension>{extension}</extension></canOpen>
    <newDat>
      <datVersionURL>{base_url}{datCode}.txt</datVersionURL>
      <datURL fileName="{datCode}.zip">{base_url}{datCode}.zip</datURL>
      <imURL>{base_url}{datCode}/</imURL>
    </newDat>
    <search>
      <to value="publisher" default="true" auto="true" />
      <to value="sourceRom" default="true" auto="true" />
      <to value="saveType" default="true" auto="true" />
      <to value="location" default="false" auto="true" />
      <to value="languages" default="false" auto="true" />
    </search>
    <romTitle>%n</romTitle>
  </configuration>"""


def build_configuration_xml_string(config: dict,
                                   dat_name: str,
                                   im_folder: str,
                                   system: str,
                                   ss_width: str,
                                   ss_height: str,
                                   extension_with_dot: str,
                                   dat_code: str) -> str:
    base_url = config.get("base_url", "")
    dat_version = datetime.now().strftime("%Y%m%d")

    return CONFIGURATION_TEMPLATE.format(
        datName=dat_name,
        imFolder=im_folder,
        datVersion=dat_version,
        system=system,
        screenshotsWidth=ss_width,
        screenshotsHeight=ss_height,
        extension=extension_with_dot,
        base_url=base_url,
        datCode=dat_code,
    )


# -------------------------------
# CSV → XML 全体生成
# -------------------------------
def generate_xml_from_csv(csv_path: str,
                          dat_name: str,
                          im_folder: str,
                          system: str,
                          ss_width: str,
                          ss_height: str,
                          extension_suffix: str,
                          dat_code: str) -> str:
    if not csv_path:
        raise ValueError("CSV ファイルが指定されていません。")

    extension_with_dot = "." + extension_suffix.strip().lstrip(".") if extension_suffix.strip() else ""

    config = load_config()

    configuration_xml = build_configuration_xml_string(
        config=config,
        dat_name=dat_name,
        im_folder=im_folder,
        system=system,
        ss_width=ss_width,
        ss_height=ss_height,
        extension_with_dot=extension_with_dot,
        dat_code=dat_code,
    )

    games_elem = build_games_from_csv(csv_path, extension_with_dot)
    games_xml = prettify_element(games_elem)

    header = '<?xml version="1.0" encoding="UTF-8" standalone="no"?>\n' \
             '<dat xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" ' \
             'xsi:noNamespaceSchemaLocation="datas.xsd">\n'
    footer = "\n</dat>\n"

    full_xml = header + configuration_xml + "\n" + games_xml + footer

    # 出力ファイル名は CSV 名ベース
    csv_dir = os.path.dirname(csv_path)
    base_name = os.path.splitext(os.path.basename(csv_path))[0]
    output_name = base_name + ".xml"
    output_path = os.path.join(csv_dir, output_name)

    with open(output_path, "w", encoding="utf-8", newline="\n") as f:
        f.write(full_xml)

    return output_path


# -------------------------------
# XML → CSV（Shift-JIS/UTF-8 自動判別）
# -------------------------------
def read_xml_text_auto(xml_path: str) -> str:
    with open(xml_path, "rb") as f:
        data = f.read()
    try:
        return data.decode("cp932")
    except UnicodeDecodeError:
        return data.decode("utf-8", errors="replace")


def generate_csv_from_xml(xml_path: str) -> str:
    if not xml_path:
        raise ValueError("XML ファイルが指定されていません。")

    xml_text = read_xml_text_auto(xml_path)
    root = ET.fromstring(xml_text)

    games = root.findall(".//game")
    if not games:
        raise ValueError("<game> タグが見つかりません。")

    fieldnames = [
        "imageNumber",
        "releaseNumber",
        "title",
        "im1CRC",
        "im2CRC",
        "publisher",
        "sourceRom",
        "location",
        "comment",
        "language",
        "saveType",
        "romSize",
        "romCRC",
        "extension",
    ]

    base, _ = os.path.splitext(xml_path)
    csv_path = base + ".csv"

    with open(csv_path, "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()

        for game in games:
            row = {key: "" for key in fieldnames}

            files = game.find("files")
            if files is not None:
                romcrc = files.find("romCRC")
                if romcrc is not None:
                    row["romCRC"] = (romcrc.text or "").strip()
                    row["extension"] = (romcrc.get("extension") or "").strip()

            for tag_name in fieldnames:
                if tag_name in ("romCRC", "extension"):
                    continue
                tag = game.find(tag_name)
                if tag is not None and tag.text is not None:
                    row[tag_name] = tag.text.strip()

            writer.writerow(row)

    return csv_path

# -------------------------------
# GUI
# -------------------------------
class OfflineListGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("OfflineList CSV/XML Converter")

        self.dat_name_var = tk.StringVar(value="")
        self.system_var = tk.StringVar(value="")
        self.im_folder_var = tk.StringVar(value="")
        self.screenshots_width_var = tk.StringVar(value="320")
        self.screenshots_height_var = tk.StringVar(value="224")
        self.extension_suffix_var = tk.StringVar(value="")
        self.dat_code_var = tk.StringVar(value="")

        self.dat_code_manual_override = False

        self.input_path = ""
        self.mode_var = tk.StringVar(value="未選択")

        self.build_ui()
        self.setup_bindings()
        self.handle_initial_argv()

    # -------------------------------
    # プレースホルダー
    # -------------------------------
    def add_placeholder(self, entry, text):
        entry.insert(0, text)
        entry.config(fg="gray")

        def on_focus_in(event):
            if entry.get() == text:
                entry.delete(0, "end")
                entry.config(fg="black")

        def on_focus_out(event):
            if entry.get() == "":
                entry.insert(0, text)
                entry.config(fg="gray")

        entry.bind("<FocusIn>", on_focus_in)
        entry.bind("<FocusOut>", on_focus_out)

    # -------------------------------
    # ツールチップ
    # -------------------------------
    class ToolTip:
        def __init__(self, widget, text):
            self.widget = widget
            self.text = text
            self.tip = None
            widget.bind("<Enter>", self.show)
            widget.bind("<Leave>", self.hide)

        def show(self, event=None):
            if self.tip is not None:
                return
            x = self.widget.winfo_rootx() + 20
            y = self.widget.winfo_rooty() + 20
            self.tip = tw = tk.Toplevel(self.widget)
            tw.wm_overrideredirect(True)
            tw.wm_geometry(f"+{x}+{y}")
            label = tk.Label(
                tw,
                text=self.text,
                background="#ffffe0",
                relief="solid",
                borderwidth=1,
                font=("Arial", 10)
            )
            label.pack()

        def hide(self, event=None):
            if self.tip:
                self.tip.destroy()
                self.tip = None

    # -------------------------------
    # UI 構築
    # -------------------------------
    def build_ui(self):
        pad = 5
        row = 0

        # datName
        tk.Label(self.root, text="datName:").grid(row=row, column=0, padx=pad, pady=pad, sticky="e")
        datName_entry = tk.Entry(self.root, textvariable=self.dat_name_var, width=40)
        datName_entry.grid(row=row, column=1, padx=pad, pady=pad, sticky="w")
        self.add_placeholder(datName_entry, "例：ファミコン")
        self.ToolTip(datName_entry, "OfflineList の左下リストに表示される名称")
        row += 1

        # system
        tk.Label(self.root, text="system:").grid(row=row, column=0, padx=pad, pady=pad, sticky="e")
        system_entry = tk.Entry(self.root, textvariable=self.system_var, width=20)
        system_entry.grid(row=row, column=1, padx=pad, pady=pad, sticky="w")
        self.add_placeholder(system_entry, "例：fc（スペースなし）")
        self.ToolTip(system_entry, "ゲーム機の名称。imFolder や URL に使用されるためスペース不可")
        system_entry.bind("<FocusOut>", self.on_system_focus_out)
        row += 1

        # imFolder
        tk.Label(self.root, text="imFolder:").grid(row=row, column=0, padx=pad, pady=pad, sticky="e")
        imFolder_entry = tk.Entry(self.root, textvariable=self.im_folder_var, width=40)
        imFolder_entry.grid(row=row, column=1, padx=pad, pady=pad, sticky="w")
        self.add_placeholder(imFolder_entry, "例：fcimg")
        self.ToolTip(imFolder_entry, "サムネイル画像を格納するフォルダ名")
        row += 1

        # screenshots
        tk.Label(self.root, text="screenshots:").grid(row=row, column=0, padx=pad, pady=pad, sticky="e")
        frame_ss = tk.Frame(self.root)
        frame_ss.grid(row=row, column=1, padx=pad, pady=pad, sticky="w")

        ss_w = tk.Entry(frame_ss, textvariable=self.screenshots_width_var, width=8)
        ss_w.pack(side="left")
        tk.Label(frame_ss, text=" x ").pack(side="left")
        ss_h = tk.Entry(frame_ss, textvariable=self.screenshots_height_var, width=8)
        ss_h.pack(side="left")

        self.add_placeholder(ss_w, "")
        self.add_placeholder(ss_h, "")
        self.ToolTip(frame_ss, "サムネイル画像の解像度（幅×高さ）")
        row += 1

        # extension
        tk.Label(self.root, text="extension:").grid(row=row, column=0, padx=pad, pady=pad, sticky="e")
        frame_ext = tk.Frame(self.root)
        frame_ext.grid(row=row, column=1, padx=pad, pady=pad, sticky="w")
        tk.Label(frame_ext, text=".").pack(side="left")
        ext_entry = tk.Entry(frame_ext, textvariable=self.extension_suffix_var, width=10)
        ext_entry.pack(side="left")
        self.add_placeholder(ext_entry, "例：nes")
        self.ToolTip(ext_entry, "ROM の拡張子（nes / smc / gb など）")
        row += 1

        # datCode
        tk.Label(self.root, text="datCode:").grid(row=row, column=0, padx=pad, pady=pad, sticky="e")
        datCode_entry = tk.Entry(self.root, textvariable=self.dat_code_var, width=20)
        datCode_entry.grid(row=row, column=1, padx=pad, pady=pad, sticky="w")
        self.add_placeholder(datCode_entry, "例：fc")
        self.ToolTip(datCode_entry, "DAT 更新 URL で使用される識別コード")
        row += 1

        # ファイル選択
        tk.Button(self.root, text="CSV / XML を選択", command=self.select_input).grid(row=row, column=0, padx=pad, pady=pad, sticky="e")
        self.input_label = tk.Label(self.root, text="未選択", anchor="w", width=50)
        self.input_label.grid(row=row, column=1, padx=pad, pady=pad, sticky="w")
        row += 1

        # モード表示
        tk.Label(self.root, text="現在のモード:").grid(row=row, column=0, padx=pad, pady=pad, sticky="e")
        self.mode_label = tk.Label(self.root, textvariable=self.mode_var, anchor="w")
        self.mode_label.grid(row=row, column=1, padx=pad, pady=pad, sticky="w")
        row += 1

        # 変換ボタン
        tk.Button(self.root, text="変換実行", command=self.on_convert).grid(row=row, column=0, columnspan=2, padx=pad, pady=pad)
        row += 1

        # 説明ラベル
        info = tk.Label(
            self.root,
            text=(
                "・CSV を選ぶと CSV → XML（DAT）を生成（フォーム必須チェックあり）。\n"
                "・XML を選ぶと XML → CSV を生成（フォームは無視されます）。\n"
                "・system を変えると imFolder=system+'img' が自動設定されます。\n"
                "・system 入力完了後（フォーカスアウト時）に datCode が自動設定されます。\n"
                "・datCode を手動変更すると自動設定は停止します。\n"
                "・CSV/XML を exe にドラッグ＆ドロップして起動すると、そのファイルが自動選択されます。"
            ),
            justify="left"
        )
        info.grid(row=row, column=0, columnspan=2, padx=pad, pady=pad, sticky="w")

        self.root.resizable(False, False)

    # -------------------------------
    # datCode 手動変更検知
    # -------------------------------
    def setup_bindings(self):
        def on_datcode_change(*args):
            self.dat_code_manual_override = True

        self.dat_code_var.trace_add("write", on_datcode_change)

    # -------------------------------
    # system 入力後の自動設定
    # -------------------------------
    def on_system_focus_out(self, event):
        system = self.system_var.get().strip()
        if system:
            self.im_folder_var.set(system + "img")
            if not self.dat_code_manual_override:
                self.dat_code_var.set(system)

    # -------------------------------
    # exe ドラッグ＆ドロップ対応
    # -------------------------------
    def handle_initial_argv(self):
        if len(sys.argv) >= 2:
            path = sys.argv[1]
            if os.path.isfile(path) and (is_csv(path) or is_xml(path)):
                self.set_input_file(path)

    # -------------------------------
    # ファイル選択
    # -------------------------------
    def select_input(self):
        path = filedialog.askopenfilename(
            title="CSV または XML ファイルを選択",
            filetypes=[
                ("CSV and XML", "*.csv *.xml"),
                ("CSV files", "*.csv"),
                ("XML files", "*.xml"),
                ("All files", "*.*"),
            ]
        )
        if path:
            self.set_input_file(path)

    # -------------------------------
    # 入力ファイル設定
    # -------------------------------
    def set_input_file(self, path: str):
        self.input_path = path
        self.input_label.config(text=os.path.basename(path))

        if is_csv(path):
            self.mode_var.set("CSV → XML")
        elif is_xml(path):
            self.mode_var.set("XML → CSV")
        else:
            self.mode_var.set("未対応の拡張子")

    # -------------------------------
    # CSV → XML 必須項目チェック
    # -------------------------------
    def validate_csv_mode_required_fields(self):
        missing = []
        if not self.dat_name_var.get().strip():
            missing.append("datName")
        if not self.system_var.get().strip():
            missing.append("system")
        if not self.im_folder_var.get().strip():
            missing.append("imFolder")
        if not self.extension_suffix_var.get().strip():
            missing.append("extension")
        if not self.dat_code_var.get().strip():
            missing.append("datCode")

        if missing:
            msg = "以下の項目が未入力です:\n\n" + "\n".join(missing)
            raise ValueError(msg)

    # -------------------------------
    # 変換実行
    # -------------------------------
    def on_convert(self):
        if not self.input_path:
            messagebox.showerror("エラー", "CSV または XML ファイルが選択されていません。")
            return

        try:
            if is_csv(self.input_path):
                self.validate_csv_mode_required_fields()
                output_path = generate_xml_from_csv(
                    csv_path=self.input_path,
                    dat_name=self.dat_name_var.get(),
                    im_folder=self.im_folder_var.get(),
                    system=self.system_var.get(),
                    ss_width=self.screenshots_width_var.get(),
                    ss_height=self.screenshots_height_var.get(),
                    extension_suffix=self.extension_suffix_var.get(),
                    dat_code=self.dat_code_var.get(),
                )
                messagebox.showinfo("完了", f"CSV → XML 変換が完了しました。\n\n{output_path}")

            elif is_xml(self.input_path):
                output_path = generate_csv_from_xml(self.input_path)
                messagebox.showinfo("完了", f"XML → CSV 変換が完了しました。\n\n{output_path}")

            else:
                messagebox.showerror("エラー", "CSV / XML 以外のファイルです。")

        except Exception as e:
            messagebox.showerror("エラー", f"変換に失敗しました。\n\n{e}")

# -------------------------------
# main
# -------------------------------
def main():
    root = tk.Tk()
    app = OfflineListGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
