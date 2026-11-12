# -*- coding: utf-8 -*-
"""
PyInstaller 图形化封包工具

- 自动创建/复用虚拟环境，识别入口脚本的第三方依赖并自动 pip 安装
- 支持单文件 / 目录两种打包模式
- UPX 压缩默认开启，本地找不到 upx 时自动下载
- 图标支持预览；选择 PNG/JPG/BMP 时自动转换为含 16/20/24/32/40/48/64/96/128/256 多尺寸的 .ico，并校验尺寸完整性
- 可选写入 exe 版本号、公司名、产品名、版权等信息
- 支持附加数据（文件/文件夹）、附加 PyInstaller 参数、打包完成后自动打开目录
- 自动记忆上次配置
"""

import ast
import json
import os
import queue
import re
import shlex
import shutil
import subprocess
import sys
import threading
import time
import urllib.request
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

UPX_VERSION = "4.2.4"
UPX_URL = f"https://github.com/upx/upx/releases/download/v{UPX_VERSION}/upx-{UPX_VERSION}-win64.zip"
DATA_SEP = ";" if os.name == "nt" else ":"
CONFIG_PATH = Path(__file__).resolve().parent / "pybuild_tool_config.json"

ICON_SIZES = (16, 20, 24, 32, 40, 48, 64, 96, 128, 256)
ICON_CONVERT_CODE = (
    "import sys\n"
    "from PIL import Image\n"
    "src, dst = sys.argv[1], sys.argv[2]\n"
    "sizes = sorted({int(x) for x in sys.argv[3].split(',') if x.strip()})\n"
    "im = Image.open(src).convert('RGBA')\n"
    "w, h = im.size\n"
    "side = max(w, h)\n"
    "canvas = Image.new('RGBA', (side, side), (0, 0, 0, 0))\n"
    "canvas.paste(im, ((side - w) // 2, (side - h) // 2))\n"
    "frames = [canvas.resize((s, s), Image.LANCZOS) for s in sizes]\n"
    "frames[-1].save(dst, format='ICO', sizes=[(s, s) for s in sizes], append_images=frames[:-1])\n"
)

PIP_NAME_MAP = {
    "PIL": "pillow",
    "cv2": "opencv-python",
    "yaml": "pyyaml",
    "sklearn": "scikit-learn",
    "bs4": "beautifulsoup4",
    "docx": "python-docx",
    "pptx": "python-pptx",
    "fitz": "pymupdf",
    "win32api": "pywin32",
    "win32com": "pywin32",
    "win32con": "pywin32",
    "pythonwin": "pywin32",
    "serial": "pyserial",
    "usb": "pyusb",
    "Crypto": "pycryptodome",
    "OpenSSL": "pyOpenSSL",
    "dateutil": "python-dateutil",
    "MySQLdb": "mysqlclient",
    "psycopg2": "psycopg2-binary",
    "pymysql": "PyMySQL",
    "jwt": "PyJWT",
    "dotenv": "python-dotenv",
    "magic": "python-magic",
    "wx": "wxPython",
    "qtpy": "QtPy",
    "pkg_resources": "setuptools",
    "Levenshtein": "python-Levenshtein",
}


class BuildError(Exception):
    pass


@dataclass
class BuildConfig:
    script: str = ""
    name: str = ""
    out_dir: str = ""
    use_venv: bool = True
    venv_dir: str = ""
    onedir: bool = False
    windowed: bool = True
    clean: bool = True
    use_upx: bool = True
    upx_dir: str = ""
    icon: str = ""
    icon_as_data: bool = True
    extra_datas: list = field(default_factory=list)
    extra_args: str = ""
    version_enabled: bool = False
    version: str = "1.0.0.0"
    company: str = ""
    product: str = ""
    copyright: str = ""
    description: str = ""


def _child_env():
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUTF8"] = "1"
    return env


def _decode_line(raw):
    for enc in ("utf-8", "gbk"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", "replace")


def _run(cmd, log, cwd=None):
    log("> " + " ".join(f'"{c}"' if " " in c else c for c in cmd))
    proc = subprocess.Popen(
        cmd, cwd=cwd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
        env=_child_env()
    )
    for raw in iter(proc.stdout.readline, b""):
        line = _decode_line(raw).rstrip()
        if line:
            log(line)
    code = proc.wait()
    if code != 0:
        raise BuildError(f"命令执行失败（退出码 {code}）：{Path(cmd[0]).name} {cmd[1] if len(cmd) > 1 else ''}")


def venv_python(venv_dir):
    p = Path(venv_dir) / ("Scripts/python.exe" if os.name == "nt" else "bin/python")
    return str(p) if p.exists() else None


def detect_imports(script):
    src = Path(script).read_text(encoding="utf-8", errors="replace")
    tree = ast.parse(src)
    mods = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for alias in node.names:
                mods.add(alias.name.split(".")[0])
        elif isinstance(node, ast.ImportFrom):
            if node.level == 0 and node.module:
                mods.add(node.module.split(".")[0])
    return mods


def third_party_modules(script):
    mods = detect_imports(script)
    local = {p.stem for p in Path(script).parent.glob("*.py")}
    stdlib = set(sys.stdlib_module_names)
    return sorted(m for m in mods if m not in stdlib and m not in local)


def _missing_modules(py, mods):
    code = (
        "import importlib.util, sys\n"
        "print(' '.join(m for m in sys.argv[1:] if importlib.util.find_spec(m) is None))\n"
    )
    out = subprocess.run([py, "-c", code, *mods], capture_output=True, text=True, errors="replace")
    if out.returncode != 0:
        return list(mods)
    return out.stdout.split()


def ensure_venv(cfg, log):
    py = venv_python(cfg.venv_dir)
    if py:
        log(f"[环境] 复用已有虚拟环境：{py}")
        return py
    log(f"[环境] 正在创建虚拟环境：{cfg.venv_dir}")
    Path(cfg.venv_dir).parent.mkdir(parents=True, exist_ok=True)
    _run([sys.executable, "-m", "venv", cfg.venv_dir], log)
    py = venv_python(cfg.venv_dir)
    if not py:
        raise BuildError("虚拟环境创建失败：未找到解释器")
    log("[环境] 虚拟环境创建完成")
    return py


def ensure_pyinstaller(py, log):
    chk = subprocess.run([py, "-c", "import PyInstaller"], capture_output=True)
    if chk.returncode == 0:
        ver = subprocess.run([py, "-c", "import PyInstaller;print(PyInstaller.__version__)"],
                             capture_output=True, text=True)
        log(f"[依赖] PyInstaller 已安装（{ver.stdout.strip()}）")
        return
    log("[依赖] 未检测到 PyInstaller，正在自动下载安装...")
    _run([py, "-m", "pip", "install", "--disable-pip-version-check", "pyinstaller"], log)


def ensure_dependencies(cfg, py, log):
    mods = third_party_modules(cfg.script)
    if not mods:
        log("[依赖] 未发现第三方依赖")
        return
    log(f"[依赖] 识别到第三方库：{', '.join(mods)}")
    missing = _missing_modules(py, mods)
    if not missing:
        log("[依赖] 全部依赖已就绪")
        return
    names = sorted({PIP_NAME_MAP.get(m, m) for m in missing})
    log(f"[依赖] 缺失：{', '.join(missing)} → 自动安装：{', '.join(names)}")
    _run([py, "-m", "pip", "install", "--disable-pip-version-check", *names], log)
    still = _missing_modules(py, mods)
    if still:
        raise BuildError(f"以下依赖安装后仍不可用：{', '.join(still)}")
    log("[依赖] 缺失依赖安装完成")


def find_upx(cfg):
    if cfg.upx_dir:
        p = Path(cfg.upx_dir)
        if p.is_file() and p.name.lower() in ("upx.exe", "upx"):
            return str(p.parent)
        exe = p / ("upx.exe" if os.name == "nt" else "upx")
        if exe.is_file():
            return str(p)
        found = next(iter(p.glob(f"*/{'upx.exe' if os.name == 'nt' else 'upx'}")), None)
        if found:
            return str(found.parent)
        return None
    w = shutil.which("upx")
    return str(Path(w).parent) if w else None


def ensure_upx(cfg, log):
    if cfg.upx_dir:
        d = find_upx(cfg)
        if not d:
            raise BuildError(f"指定目录中未找到 upx 可执行文件：{cfg.upx_dir}")
        log(f"[UPX] 使用指定目录：{d}")
        return d

    w = shutil.which("upx")
    if w:
        log(f"[UPX] 已从 PATH 找到：{w}")
        return str(Path(w).parent)

    target = Path(__file__).resolve().parent / "upx"
    target.mkdir(exist_ok=True)
    zip_path = target / f"upx-{UPX_VERSION}-win64.zip"
    log(f"[UPX] 未找到 upx，自动下载 UPX {UPX_VERSION} ...")

    last = {"pct": -1}

    def hook(blocks, block_size, total):
        if total > 0:
            pct = min(100, blocks * block_size * 100 // total)
            if pct >= last["pct"] + 25:
                last["pct"] = pct
                log(f"[UPX] 下载进度 {pct}%")

    try:
        urllib.request.urlretrieve(UPX_URL, zip_path, hook)
        with zipfile.ZipFile(zip_path) as z:
            z.extractall(target)
    finally:
        zip_path.unlink(missing_ok=True)

    found = next(iter(target.rglob("upx.exe" if os.name == "nt" else "upx")), None)
    if not found:
        raise BuildError("UPX 下载完成但未找到可执行文件")
    log(f"[UPX] 下载完成：{found.parent}")
    return str(found.parent)


def _esc(s):
    return str(s).replace("\\", "\\\\").replace("'", "\\'")


def _ver_tuple(v):
    parts = [int(x) for x in re.findall(r"\d+", str(v))]
    return tuple((parts + [0, 0, 0, 0])[:4])


def build_version_file(cfg, path):
    fv = _ver_tuple(cfg.version)
    content = f"""# UTF-8
VSVersionInfo(
  ffi=FixedFileInfo(filevers={fv}, prodvers={fv}, mask=0x3f, flags=0x0, OS=0x40004, fileType=0x1, subtype=0x0, date=(0, 0)),
  kids=[
    StringFileInfo([
      StringTable('080404b0', [
        StringStruct('CompanyName', '{_esc(cfg.company)}'),
        StringStruct('FileDescription', '{_esc(cfg.description or cfg.name)}'),
        StringStruct('FileVersion', '{_esc(cfg.version)}'),
        StringStruct('InternalName', '{_esc(cfg.name)}'),
        StringStruct('LegalCopyright', '{_esc(cfg.copyright)}'),
        StringStruct('OriginalFilename', '{_esc(cfg.name)}.exe'),
        StringStruct('ProductName', '{_esc(cfg.product or cfg.name)}'),
        StringStruct('ProductVersion', '{_esc(cfg.version)}')
      ])
    ]),
    VarFileInfo([VarStruct('Translation', [2052, 1200])])
  ]
)
"""
    Path(path).write_text(content, encoding="utf-8")


def ico_sizes(path):
    data = Path(path).read_bytes()
    if len(data) < 6 or data[:4] != b"\x00\x00\x01\x00":
        return []
    count = int.from_bytes(data[4:6], "little")
    sizes = []
    for i in range(count):
        off = 6 + i * 16
        if off + 2 > len(data):
            break
        sizes.append((data[off] or 256, data[off + 1] or 256))
    return sizes


def convert_image_to_ico(src, dst, sizes=ICON_SIZES):
    from PIL import Image
    sizes = sorted(set(sizes))
    im = Image.open(src).convert("RGBA")
    w, h = im.size
    side = max(w, h)
    canvas = Image.new("RGBA", (side, side), (0, 0, 0, 0))
    canvas.paste(im, ((side - w) // 2, (side - h) // 2))
    frames = [canvas.resize((s, s), Image.LANCZOS) for s in sizes]
    Path(dst).parent.mkdir(parents=True, exist_ok=True)
    frames[-1].save(dst, format="ICO", sizes=[(s, s) for s in sizes], append_images=frames[:-1])
    return dst


def _log_icon_sizes(icon, log):
    sizes = sorted({w for w, _ in ico_sizes(icon)})
    if not sizes:
        log(f"[图标] 警告：无法解析图标尺寸信息：{icon}")
        return
    log(f"[图标] 图标包含尺寸：{'/'.join(map(str, sizes))}")
    missing = sorted({16, 32, 48, 64} - set(sizes))
    if missing:
        log(f"[图标] 提示：缺少常用尺寸 {'/'.join(map(str, missing))}，小图标显示可能模糊")


def normalize_icon(cfg, py, log, work_dir):
    if not cfg.icon:
        return None
    p = Path(cfg.icon)
    if not p.is_file():
        raise BuildError(f"图标文件不存在：{p}")
    if p.suffix.lower() == ".ico":
        _log_icon_sizes(str(p.resolve()), log)
        return str(p.resolve())

    if subprocess.run([py, "-c", "import PIL"], capture_output=True).returncode != 0:
        log("[图标] 图标转换需要 Pillow，正在安装到打包环境...")
        _run([py, "-m", "pip", "install", "--disable-pip-version-check", "pillow"], log)

    ico = Path(work_dir) / "app_icon.ico"
    r = subprocess.run(
        [py, "-c", ICON_CONVERT_CODE, str(p.resolve()), str(ico), ",".join(map(str, ICON_SIZES))],
        capture_output=True, text=True, encoding="utf-8", errors="replace", env=_child_env())
    if r.returncode != 0:
        raise BuildError(f"图标转换失败：{r.stderr.strip() or '未知错误'}")
    log(f"[图标] 已自动转换 {p.name} → 多尺寸 .ico")
    _log_icon_sizes(str(ico), log)
    return str(ico)


def build_version_and_datas(cfg, icon, work_dir):
    datas = []
    if icon and cfg.icon_as_data:
        data_icon = icon
        if Path(icon).name.lower() != "app_icon.ico":
            data_icon = str(Path(work_dir) / "app_icon.ico")
            shutil.copyfile(icon, data_icon)
        datas.append((data_icon, "."))
    for d in cfg.extra_datas:
        p = Path(d)
        if p.is_dir():
            datas.append((str(p.resolve()), p.name))
        elif p.is_file():
            datas.append((str(p.resolve()), "."))
        else:
            raise BuildError(f"附加数据不存在：{d}")
    return list(dict.fromkeys(datas))


def _parse_extra_args(raw):
    if not raw.strip():
        return []
    return [a.strip('"').strip("'") for a in shlex.split(raw, posix=False) if a.strip()]


def _artifact_size(p):
    p = Path(p)
    if p.is_file():
        return p.stat().st_size
    return sum(f.stat().st_size for f in p.rglob("*") if f.is_file())


def run_build(cfg, log):
    script = Path(cfg.script).resolve()
    if not script.is_file() or script.suffix.lower() != ".py":
        raise BuildError("请选择有效的 .py 入口脚本")

    try:
        compile(script.read_bytes(), str(script), "exec")
    except SyntaxError as e:
        raise BuildError(f"脚本存在语法错误（第 {e.lineno} 行）：{e.msg}")
    log("[检查] 入口脚本语法检查通过")

    out_dir = Path(cfg.out_dir).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    work = out_dir / "build"
    work.mkdir(parents=True, exist_ok=True)
    log(f"[打包] 入口脚本：{script}")
    log(f"[打包] 输出目录：{out_dir}")

    if cfg.use_venv:
        py = ensure_venv(cfg, log)
    else:
        py = sys.executable
        log(f"[环境] 使用当前解释器：{py}")

    ensure_pyinstaller(py, log)
    ensure_dependencies(cfg, py, log)

    icon = normalize_icon(cfg, py, log, work)
    if icon:
        log(f"[图标] 使用的图标：{icon}")

    cmd = [py, "-m", "PyInstaller", "--noconfirm", "--clean",
           "--name", cfg.name,
           "--distpath", str(out_dir),
           "--workpath", str(work),
           "--specpath", str(work),
           "--onedir" if cfg.onedir else "--onefile",
           "--windowed" if cfg.windowed else "--console"]
    if icon:
        cmd += ["--icon", icon]

    for src, dest in build_version_and_datas(cfg, icon, work):
        cmd += ["--add-data", f"{src}{DATA_SEP}{dest}"]
        log(f"[数据] 打包附加资源：{src} → {dest}")

    if cfg.use_upx:
        try:
            cmd += ["--upx-dir", ensure_upx(cfg, log)]
        except Exception as e:
            log(f"[UPX] 获取失败，本次打包不使用 UPX：{e}")
    else:
        log("[UPX] 已关闭 UPX 压缩")

    extra = _parse_extra_args(cfg.extra_args)
    if extra:
        log(f"[参数] 附加 PyInstaller 参数：{' '.join(extra)}")
        cmd += extra

    if cfg.version_enabled:
        vf = work / "version_info.txt"
        build_version_file(cfg, vf)
        cmd += ["--version-file", str(vf)]
        log(f"[版本] 已生成版本信息文件（{cfg.version}）")

    cmd.append(str(script))
    log("[打包] 开始执行 PyInstaller ...")
    t0 = time.time()
    _run(cmd, log, cwd=str(script.parent))
    elapsed = time.time() - t0

    artifact = out_dir / (cfg.name + (".exe" if os.name == "nt" else ""))
    if cfg.onedir:
        artifact = out_dir / cfg.name
    if not artifact.exists():
        raise BuildError("打包结束但未找到生成的产物")

    size_mb = _artifact_size(artifact) / 1024 / 1024
    log(f"[打包] 产物大小：{size_mb:.1f} MB，耗时：{elapsed:.1f} 秒")

    if cfg.clean:
        shutil.rmtree(work, ignore_errors=True)
        log("[清理] 已删除临时构建目录")

    log(f"[打包] 完成：{artifact}")
    return str(artifact)


class BuildToolGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PyInstaller 图形化封包工具")
        self.geometry("960x880")
        self.minsize(920, 780)

        self.style = ttk.Style(self)
        if "vista" in self.style.theme_names():
            self.style.theme_use("vista")
        elif "clam" in self.style.theme_names():
            self.style.theme_use("clam")

        self.msg_queue = queue.Queue()
        self.extra_data_files = []
        self.running = False
        self._preview_img = None

        self._init_vars()
        self._build_ui()
        self._load_config()
        self._update_icon_preview()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self._poll_id = self.after(80, self._poll_queue)

    def _init_vars(self):
        self.var_script = tk.StringVar()
        self.var_name = tk.StringVar()
        self.var_out = tk.StringVar()
        self.var_use_venv = tk.BooleanVar(value=True)
        self.var_venv = tk.StringVar()
        self.var_mode = tk.StringVar(value="onefile")
        self.var_windowed = tk.BooleanVar(value=True)
        self.var_clean = tk.BooleanVar(value=True)
        self.var_use_upx = tk.BooleanVar(value=True)
        self.var_upx_dir = tk.StringVar()
        self.var_icon = tk.StringVar()
        self.var_icon_data = tk.BooleanVar(value=True)
        self.var_extra_args = tk.StringVar()
        self.var_version_on = tk.BooleanVar(value=False)
        self.var_version = tk.StringVar(value="1.0.0.0")
        self.var_company = tk.StringVar()
        self.var_product = tk.StringVar()
        self.var_copyright = tk.StringVar(value="Copyright (C) 2026")
        self.var_description = tk.StringVar()
        self.var_open_after = tk.BooleanVar(value=True)
        self.var_status = tk.StringVar(value="就绪")

    def _form_row(self, parent, row, label, var, browse_cmd=None):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky=tk.W, pady=3)
        ent = ttk.Entry(parent, textvariable=var)
        ent.grid(row=row, column=1, sticky=tk.EW, padx=6, pady=3)
        if browse_cmd:
            ttk.Button(parent, text="浏览...", width=8, command=browse_cmd).grid(row=row, column=2, pady=3)
        return ent

    def _build_ui(self):
        main = ttk.Frame(self, padding=12)
        main.pack(fill=tk.BOTH, expand=True)

        g1 = ttk.LabelFrame(main, text=" 1. 基本信息 ", padding=8)
        g1.pack(fill=tk.X)
        g1.columnconfigure(1, weight=1)
        self._form_row(g1, 0, "入口脚本：", self.var_script, self.choose_script)
        self._form_row(g1, 1, "程序名称：", self.var_name)
        self._form_row(g1, 2, "输出目录：", self.var_out, self.choose_out_dir)

        mid = ttk.Frame(main)
        mid.pack(fill=tk.X, pady=6)

        g2 = ttk.LabelFrame(mid, text=" 2. 环境与压缩 ", padding=8)
        g2.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 6))
        g2.columnconfigure(1, weight=1)
        ttk.Checkbutton(g2, text="自动创建/复用虚拟环境", variable=self.var_use_venv).grid(
            row=0, column=0, columnspan=3, sticky=tk.W)
        self._form_row(g2, 1, "环境路径：", self.var_venv, self.choose_venv_dir)
        ttk.Checkbutton(g2, text="启用 UPX 压缩（推荐）", variable=self.var_use_upx).grid(
            row=2, column=0, columnspan=3, sticky=tk.W, pady=(6, 0))
        self._form_row(g2, 3, "UPX 目录：", self.var_upx_dir, self.choose_upx_dir)
        ttk.Label(g2, text=f"留空自动查找；找不到时自动下载 UPX {UPX_VERSION}",
                  foreground="#777777", font=("Microsoft YaHei UI", 8)).grid(
            row=4, column=0, columnspan=3, sticky=tk.W)

        g3 = ttk.LabelFrame(mid, text=" 3. 打包选项 ", padding=8)
        g3.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        g3.columnconfigure(1, weight=1)
        mode_bar = ttk.Frame(g3)
        mode_bar.grid(row=0, column=0, columnspan=3, sticky=tk.W)
        ttk.Radiobutton(mode_bar, text="单文件", variable=self.var_mode, value="onefile").pack(side=tk.LEFT)
        ttk.Radiobutton(mode_bar, text="目录", variable=self.var_mode, value="onedir").pack(side=tk.LEFT, padx=(10, 0))
        ttk.Checkbutton(g3, text="隐藏控制台窗口（GUI 程序推荐）", variable=self.var_windowed).grid(
            row=1, column=0, columnspan=3, sticky=tk.W, pady=(4, 0))
        ttk.Checkbutton(g3, text="打包后清理临时构建文件", variable=self.var_clean).grid(
            row=2, column=0, columnspan=3, sticky=tk.W)
        self._form_row(g3, 3, "程序图标：", self.var_icon, self.choose_icon)
        prev_row = ttk.Frame(g3)
        prev_row.grid(row=4, column=0, columnspan=3, sticky=tk.W, pady=(2, 0))
        ttk.Label(prev_row, text="图标预览：").pack(side=tk.LEFT)
        self.lbl_icon_preview = ttk.Label(prev_row, text="未选择图标", foreground="#999999")
        self.lbl_icon_preview.pack(side=tk.LEFT)
        ttk.Label(g3, text="支持 .ico；选择 PNG/JPG/BMP 后自动转换为含 16~256 多尺寸的 .ico",
                  foreground="#777777", font=("Microsoft YaHei UI", 8)).grid(
            row=5, column=0, columnspan=3, sticky=tk.W)
        ttk.Checkbutton(g3, text="图标同时以 app_icon.ico 打包进资源目录（程序内加载图标时需要）",
                        variable=self.var_icon_data).grid(row=6, column=0, columnspan=3, sticky=tk.W)
        self.var_icon.trace_add("write", lambda *a: self._update_icon_preview())

        g4 = ttk.LabelFrame(main, text=" 4. 附加数据（打包进程序资源目录） ", padding=8)
        g4.pack(fill=tk.X, pady=(0, 6))
        data_bar = ttk.Frame(g4)
        data_bar.pack(fill=tk.X)
        self.lst_data = tk.Listbox(data_bar, height=2, font=("Consolas", 9))
        self.lst_data.pack(side=tk.LEFT, fill=tk.X, expand=True)
        data_btns = ttk.Frame(data_bar)
        data_btns.pack(side=tk.RIGHT, padx=(6, 0))
        ttk.Button(data_btns, text="添加文件", width=10, command=self.add_data_file).pack(pady=2)
        ttk.Button(data_btns, text="添加文件夹", width=10, command=self.add_data_dir).pack(pady=2)
        ttk.Button(data_btns, text="移除选中", width=10, command=self.remove_data_file).pack()

        g5 = ttk.LabelFrame(main, text=" 5. 高级参数 ", padding=8)
        g5.pack(fill=tk.X, pady=(0, 6))
        g5.columnconfigure(1, weight=1)
        ttk.Label(g5, text="附加参数：").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(g5, textvariable=self.var_extra_args).grid(row=0, column=1, sticky=tk.EW, padx=6)
        ttk.Label(g5, text="直接传给 PyInstaller，例如：--hidden-import=xxx --exclude-module=yyy",
                  foreground="#777777", font=("Microsoft YaHei UI", 8)).grid(
            row=1, column=0, columnspan=2, sticky=tk.W)

        g6 = ttk.LabelFrame(main, text=" 6. exe 版本信息与版权 ", padding=8)
        g6.pack(fill=tk.X, pady=(0, 6))
        ttk.Checkbutton(g6, text="写入版本信息（可在文件右键属性中查看）", variable=self.var_version_on).grid(
            row=0, column=0, columnspan=4, sticky=tk.W, pady=(0, 4))
        ttk.Label(g6, text="版本号：").grid(row=1, column=0, sticky=tk.W)
        ttk.Entry(g6, textvariable=self.var_version, width=16).grid(row=1, column=1, sticky=tk.W, padx=(0, 16))
        ttk.Label(g6, text="公司名：").grid(row=1, column=2, sticky=tk.W)
        ttk.Entry(g6, textvariable=self.var_company).grid(row=1, column=3, sticky=tk.EW)
        ttk.Label(g6, text="产品名称：").grid(row=2, column=0, sticky=tk.W, pady=2)
        ttk.Entry(g6, textvariable=self.var_product).grid(row=2, column=1, sticky=tk.EW, padx=(0, 16), pady=2)
        ttk.Label(g6, text="版权信息：").grid(row=2, column=2, sticky=tk.W)
        ttk.Entry(g6, textvariable=self.var_copyright).grid(row=2, column=3, sticky=tk.EW)
        ttk.Label(g6, text="文件描述：").grid(row=3, column=0, sticky=tk.W)
        ttk.Entry(g6, textvariable=self.var_description).grid(row=3, column=1, columnspan=3, sticky=tk.EW)
        g6.columnconfigure(1, weight=1)
        g6.columnconfigure(3, weight=2)

        act = ttk.Frame(main)
        act.pack(fill=tk.X, pady=(2, 4))
        self.btn_run = ttk.Button(act, text="🚀 开始打包", command=self.start_build)
        self.btn_run.pack(side=tk.LEFT, ipadx=10, ipady=3)
        ttk.Button(act, text="📂 打开输出目录", command=self.open_out_dir).pack(side=tk.LEFT, padx=8)
        ttk.Checkbutton(act, text="完成后自动打开目录", variable=self.var_open_after).pack(side=tk.LEFT, padx=6)
        self.progress = ttk.Progressbar(act, mode="indeterminate", length=160)
        self.progress.pack(side=tk.RIGHT)
        ttk.Label(act, textvariable=self.var_status).pack(side=tk.RIGHT, padx=10)

        g7 = ttk.LabelFrame(main, text=" 打包日志 ", padding=6)
        g7.pack(fill=tk.BOTH, expand=True)
        self.log_text = ScrolledText(g7, height=6, font=("Consolas", 9))
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def log(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)

    def _update_icon_preview(self):
        self._preview_img = None
        path = self.var_icon.get().strip()
        if not path or not Path(path).is_file():
            self.lbl_icon_preview.configure(image="", text="未选择图标", foreground="#999999")
            return
        p = Path(path)
        try:
            from PIL import Image, ImageTk
            im = Image.open(p).convert("RGBA")
            im.thumbnail((32, 32))
            self._preview_img = ImageTk.PhotoImage(im)
            self.lbl_icon_preview.configure(image=self._preview_img, text=f" {p.name}",
                                            compound=tk.LEFT, foreground="#333333")
            return
        except Exception:
            pass
        if p.suffix.lower() == ".png":
            try:
                img = tk.PhotoImage(file=str(p))
                factor = max(1, max(img.width(), img.height()) // 32)
                self._preview_img = img.subsample(factor, factor)
                self.lbl_icon_preview.configure(image=self._preview_img, text=f" {p.name}",
                                                compound=tk.LEFT, foreground="#333333")
                return
            except tk.TclError:
                pass
        self.lbl_icon_preview.configure(image="", text=f"已选择 {p.name}", foreground="#0066cc")

    def choose_script(self):
        path = filedialog.askopenfilename(title="选择入口 Python 脚本",
                                          filetypes=[("Python 脚本", "*.py")])
        if not path:
            return
        self.var_script.set(path)
        stem = Path(path).stem
        if not self.var_name.get().strip():
            self.var_name.set(stem)
        script_dir = Path(path).parent
        if not self.var_out.get().strip():
            self.var_out.set(str(script_dir))
        if not self.var_venv.get().strip():
            self.var_venv.set(str(script_dir / ".build_venv"))
        if not self.var_icon.get().strip():
            for cand in ("app_icon.ico", "favicon.ico", "app_icon.png"):
                ico = script_dir / cand
                if ico.is_file():
                    self.var_icon.set(str(ico))
                    if ico.suffix.lower() == ".ico":
                        self._add_data_path(str(ico))
                    break

    def choose_out_dir(self):
        d = filedialog.askdirectory(title="选择输出目录",
                                    initialdir=self.var_out.get().strip() or None)
        if d:
            self.var_out.set(d)

    def choose_venv_dir(self):
        d = filedialog.askdirectory(title="选择虚拟环境目录",
                                    initialdir=self.var_venv.get().strip() or None)
        if d:
            self.var_venv.set(d)

    def choose_upx_dir(self):
        d = filedialog.askdirectory(title="选择包含 upx.exe 的目录",
                                    initialdir=self.var_upx_dir.get().strip() or None)
        if d:
            self.var_upx_dir.set(d)

    def choose_icon(self):
        path = filedialog.askopenfilename(title="选择程序图标",
                                          filetypes=[("图标/图片", "*.ico *.png *.jpg *.jpeg *.bmp"),
                                                     ("所有文件", "*.*")])
        if not path:
            return
        p = Path(path)
        if p.suffix.lower() != ".ico":
            converted = self._auto_convert_icon(p)
            if converted:
                path = converted
        self.var_icon.set(path)

    def _auto_convert_icon(self, src):
        try:
            from PIL import Image  # noqa: F401
        except ImportError:
            self.log("[图标] 当前环境未安装 Pillow，将在打包时自动转换")
            return None
        try:
            out_dir = Path(__file__).resolve().parent / "converted_icons"
            dst = out_dir / (src.stem + ".ico")
            convert_image_to_ico(src, dst)
            sizes = sorted({w for w, _ in ico_sizes(dst)})
            self.log(f"[图标] 已自动转换 {src.name} → {dst.name}，包含尺寸：{'/'.join(map(str, sizes))}")
            return str(dst)
        except Exception as e:
            self.log(f"[图标] 自动转换失败：{e}（打包时会再次尝试）")
            return None

    def _add_data_path(self, path):
        if path and path not in self.extra_data_files:
            self.extra_data_files.append(path)
            self.lst_data.insert(tk.END, path)

    def add_data_file(self):
        files = filedialog.askopenfilenames(title="选择需要打入程序的附加文件")
        for f in files:
            self._add_data_path(os.path.abspath(f))

    def add_data_dir(self):
        d = filedialog.askdirectory(title="选择需要整个打入程序的文件夹")
        if d:
            self._add_data_path(os.path.abspath(d))

    def remove_data_file(self):
        for idx in reversed(self.lst_data.curselection()):
            self.lst_data.delete(idx)
            del self.extra_data_files[idx]

    def _collect(self):
        script = self.var_script.get().strip()
        if not script:
            raise BuildError("请先选择入口 .py 脚本")
        if not Path(script).is_file():
            raise BuildError(f"入口脚本不存在：{script}")
        name = self.var_name.get().strip() or Path(script).stem
        name = re.sub(r'[<>:"/\\|?*]', "", name).strip() or Path(script).stem
        out_dir = self.var_out.get().strip() or str(Path(script).parent)
        venv_dir = self.var_venv.get().strip() or str(Path(script).parent / ".build_venv")
        version = self.var_version.get().strip() or "1.0.0.0"
        if self.var_version_on.get() and not re.fullmatch(r"\d+(\.\d+){0,3}", version):
            raise BuildError("版本号格式不正确，应为 1.0.0 或 1.0.0.0 形式")
        return BuildConfig(
            script=script,
            name=name,
            out_dir=out_dir,
            use_venv=self.var_use_venv.get(),
            venv_dir=venv_dir,
            onedir=(self.var_mode.get() == "onedir"),
            windowed=self.var_windowed.get(),
            clean=self.var_clean.get(),
            use_upx=self.var_use_upx.get(),
            upx_dir=self.var_upx_dir.get().strip(),
            icon=self.var_icon.get().strip(),
            icon_as_data=self.var_icon_data.get(),
            extra_datas=list(self.extra_data_files),
            extra_args=self.var_extra_args.get().strip(),
            version_enabled=self.var_version_on.get(),
            version=version,
            company=self.var_company.get().strip(),
            product=self.var_product.get().strip(),
            copyright=self.var_copyright.get().strip(),
            description=self.var_description.get().strip(),
        )

    def _config_snapshot(self):
        return {
            "script": self.var_script.get(),
            "name": self.var_name.get(),
            "out": self.var_out.get(),
            "use_venv": self.var_use_venv.get(),
            "venv": self.var_venv.get(),
            "mode": self.var_mode.get(),
            "windowed": self.var_windowed.get(),
            "clean": self.var_clean.get(),
            "upx": self.var_use_upx.get(),
            "upx_dir": self.var_upx_dir.get(),
            "icon": self.var_icon.get(),
            "icon_data": self.var_icon_data.get(),
            "extra_args": self.var_extra_args.get(),
            "version_on": self.var_version_on.get(),
            "version": self.var_version.get(),
            "company": self.var_company.get(),
            "product": self.var_product.get(),
            "copyright": self.var_copyright.get(),
            "description": self.var_description.get(),
            "open_after": self.var_open_after.get(),
            "extra_data_files": list(self.extra_data_files),
        }

    def _save_config(self):
        try:
            CONFIG_PATH.write_text(json.dumps(self._config_snapshot(), ensure_ascii=False, indent=2),
                                   encoding="utf-8")
        except OSError:
            pass

    def _load_config(self):
        try:
            data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except Exception:
            return
        pairs = [
            ("script", self.var_script), ("name", self.var_name), ("out", self.var_out),
            ("venv", self.var_venv), ("mode", self.var_mode), ("upx_dir", self.var_upx_dir),
            ("icon", self.var_icon), ("extra_args", self.var_extra_args),
            ("version", self.var_version), ("company", self.var_company),
            ("product", self.var_product), ("copyright", self.var_copyright),
            ("description", self.var_description),
        ]
        bools = [
            ("use_venv", self.var_use_venv), ("windowed", self.var_windowed),
            ("clean", self.var_clean), ("upx", self.var_use_upx),
            ("icon_data", self.var_icon_data), ("version_on", self.var_version_on),
            ("open_after", self.var_open_after),
        ]
        for key, var in pairs:
            if key in data and isinstance(data[key], str):
                var.set(data[key])
        for key, var in bools:
            if key in data:
                var.set(bool(data[key]))
        for f in data.get("extra_data_files", []):
            self._add_data_path(f)

    def _on_close(self):
        try:
            self.after_cancel(self._poll_id)
        except tk.TclError:
            pass
        self._save_config()
        self.destroy()

    def start_build(self):
        if self.running:
            return
        try:
            cfg = self._collect()
        except BuildError as e:
            messagebox.showwarning("提示", str(e))
            return

        self._save_config()
        self.running = True
        self.btn_run.config(state=tk.DISABLED)
        self.log_text.delete("1.0", tk.END)
        self.var_status.set("正在打包...")
        self.progress.start(12)

        threading.Thread(target=self._worker, args=(cfg,), daemon=True).start()

    def _worker(self, cfg):
        try:
            result = run_build(cfg, lambda m: self.msg_queue.put(("log", m)))
            self.msg_queue.put(("done", result))
        except Exception as e:
            self.msg_queue.put(("error", str(e)))

    def _poll_queue(self):
        try:
            while True:
                kind, payload = self.msg_queue.get_nowait()
                if kind == "log":
                    self.log(payload)
                else:
                    self._finish(kind == "done", payload)
        except queue.Empty:
            pass
        self._poll_id = self.after(80, self._poll_queue)

    def _finish(self, ok, payload):
        self.running = False
        self.progress.stop()
        self.btn_run.config(state=tk.NORMAL)
        if ok:
            self.var_status.set("打包完成")
            if self.var_open_after.get():
                self.open_out_dir()
            messagebox.showinfo("完成", f"打包成功！\n输出：\n{payload}")
        else:
            self.var_status.set("打包失败")
            messagebox.showerror("打包失败", payload)

    def open_out_dir(self):
        out_dir = self.var_out.get().strip()
        if not out_dir and self.var_script.get().strip():
            out_dir = str(Path(self.var_script.get().strip()).parent)
        if not out_dir:
            messagebox.showwarning("提示", "请先设置输出目录")
            return
        os.makedirs(out_dir, exist_ok=True)
        if os.name == "nt":
            os.startfile(out_dir)
        else:
            subprocess.Popen(["xdg-open", out_dir])


def main():
    app = BuildToolGUI()
    icon = Path(__file__).resolve().parent / "app_icon.ico"
    if icon.is_file():
        try:
            app.iconbitmap(str(icon))
        except tk.TclError:
            pass
    app.mainloop()


if __name__ == "__main__":
    main()
