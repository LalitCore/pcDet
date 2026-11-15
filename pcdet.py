"""
PCDet - every detail of your PC in one dashboard.
CPU, GPU, RAM modules, disk health (SMART), battery health, network, processes, motherboard/BIOS,
plus temperatures, voltages and FAN speeds (via LibreHardwareMonitor) - with alerts and report export.

Install:  pip install PySide6 psutil
Run:      python pcdet.py     (right-click > Run as administrator gives the most sensor data)
Fans/CPU temps: also run LibreHardwareMonitor (Options > Remote Web Server > Run) - PCDet picks it up.
"""
import sys, os, json, time, re, shutil, subprocess, threading, base64, platform, socket, urllib.request, zipfile, csv, glob, importlib
from collections import deque
import psutil
from PySide6.QtCore import Qt, QTimer, QRectF, QPointF, QUrl
from PySide6.QtGui import (QBrush, QColor, QDesktopServices, QFont, QLinearGradient, QPainter, QPainterPath, QPen)
from PySide6.QtWidgets import (QApplication, QFileDialog, QFrame, QGridLayout, QHBoxLayout, QHeaderView, QLabel,
                               QListWidget, QPlainTextEdit, QPushButton, QScrollArea, QStackedWidget,
                               QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget, QAbstractItemView)

WIN = sys.platform == "win32"
NOWIN = 0x08000000 if WIN else 0
LHM_URL = "http://localhost:8085/data.json"
LHM_DL = "https://github.com/LibreHardwareMonitor/LibreHardwareMonitor/releases"

DATA = dict(static={}, smi=[], lhm=None, acpi=[], static_time=0)
LIVE = {}
VOLS = []
HIST = {k: deque(maxlen=60) for k in ("cpu", "ram", "gpu", "down", "up", "read", "write", "temp")}
ENGINE_DIR = os.path.join(os.getenv("APPDATA", os.path.expanduser("~")), "PCDet", "lhm")
DATA.update(hw=None, setup="", engine=dict(state="not started", msg=""), gpuwin={})


# ---------------------------------------------------------------- helpers
def run(cmd, timeout=20):
    try:
        return subprocess.run(cmd, capture_output=True, text=True, timeout=timeout,
                              creationflags=NOWIN, errors="replace").stdout
    except Exception:
        return ""


def ps(script, timeout=90):
    enc = base64.b64encode(script.encode("utf-16-le")).decode()
    return run(["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
                "-EncodedCommand", enc], timeout) if WIN else ""


def lst(x):
    return x if isinstance(x, list) else ([x] if x else [])


def S(key):
    return lst(DATA["static"].get(key))


def first(x):
    return x[0] if x else {}


def v(d, k, suffix=""):
    x = d.get(k)
    return "—" if x in (None, "") else f"{x}{suffix}"


def gb(b):
    return "—" if not b else f"{b / 1024 ** 3:.1f} GB"


def rate(b):
    for u in ("B/s", "KB/s", "MB/s", "GB/s"):
        if b < 1024 or u == "GB/s":
            return f"{b:.0f} {u}" if u == "B/s" else f"{b:.1f} {u}"
        b /= 1024


def dur(s):
    s = int(s)
    d, s = divmod(s, 86400)
    h, s = divmod(s, 3600)
    return (f"{d}d " if d else "") + f"{h}h {s // 60}m"


def textbar(p, n=12):
    k = int(round(max(0, min(100, p)) / 100 * n))
    return "▰" * k + "▱" * (n - k)


def num_of(s):
    m = re.search(r"-?[\d.,]+", str(s or ""))
    if not m:
        return None
    t = m.group(0)
    t = t.replace(",", "") if "," in t and "." in t else t.replace(",", ".")
    try:
        return float(t)
    except ValueError:
        return None


RAM_TYPES = {20: "DDR", 21: "DDR2", 24: "DDR3", 26: "DDR4", 34: "DDR5", 35: "LPDDR5"}


# ---------------------------------------------------------------- data collectors
PS_STATIC = r"""
$ErrorActionPreference='SilentlyContinue'
function D($x){ if($x){ try{ ([datetime]$x).ToString('yyyy-MM-dd') }catch{ "$x" } } }
$o=[ordered]@{}
$o.cpu=@(Get-CimInstance Win32_Processor | ForEach-Object { [pscustomobject]@{Name=("$($_.Name)").Trim();Maker=$_.Manufacturer;Cores=$_.NumberOfCores;Threads=$_.NumberOfLogicalProcessors;MaxMHz=$_.MaxClockSpeed;L2KB=$_.L2CacheSize;L3KB=$_.L3CacheSize;Socket=$_.SocketDesignation;Virt=$_.VirtualizationFirmwareEnabled} })
$o.board=@(Get-CimInstance Win32_BaseBoard | ForEach-Object { [pscustomobject]@{Maker=$_.Manufacturer;Model=$_.Product;Version=$_.Version;Serial=$_.SerialNumber} })
$o.bios=@(Get-CimInstance Win32_BIOS | ForEach-Object { [pscustomobject]@{Maker=$_.Manufacturer;Version=$_.SMBIOSBIOSVersion;Date=(D $_.ReleaseDate)} })
$o.cs=@(Get-CimInstance Win32_ComputerSystem | ForEach-Object { [pscustomobject]@{Maker=$_.Manufacturer;Model=$_.Model;Type=$_.SystemType;RAM=$_.TotalPhysicalMemory;Name=$_.Name} })
$o.os=@(Get-CimInstance Win32_OperatingSystem | ForEach-Object { [pscustomobject]@{Name=$_.Caption;Version=$_.Version;Build=$_.BuildNumber;Arch=$_.OSArchitecture;Installed=(D $_.InstallDate)} })
$o.ram=@(Get-CimInstance Win32_PhysicalMemory | ForEach-Object { [pscustomobject]@{Slot=$_.DeviceLocator;Bytes=$_.Capacity;MHz=$_.Speed;Configured=$_.ConfiguredClockSpeed;Maker=$_.Manufacturer;Part=("$($_.PartNumber)").Trim();Type=$_.SMBIOSMemoryType} })
$o.gpu=@(Get-CimInstance Win32_VideoController | ForEach-Object { [pscustomobject]@{Name=$_.Name;Driver=$_.DriverVersion;DriverDate=(D $_.DriverDate);VRAM=$_.AdapterRAM;W=$_.CurrentHorizontalResolution;H=$_.CurrentVerticalResolution;Hz=$_.CurrentRefreshRate} })
$o.disks=@(Get-PhysicalDisk | ForEach-Object { $r=$_ | Get-StorageReliabilityCounter; [pscustomobject]@{Name=$_.FriendlyName;Media=[string]$_.MediaType;Bus=[string]$_.BusType;SizeGB=[math]::Round($_.Size/1GB,1);Health=[string]$_.HealthStatus;Status=[string]$_.OperationalStatus;Serial=$_.SerialNumber;TempC=$r.Temperature;Wear=$r.Wear;Hours=$r.PowerOnHours;ReadErr=$r.ReadErrorsTotal;WriteErr=$r.WriteErrorsTotal;Cycles=$r.StartStopCycleCount} })
$f=Get-CimInstance -Namespace root\wmi -ClassName BatteryFullChargedCapacity | Select-Object -First 1
$d=Get-CimInstance -Namespace root\wmi -ClassName BatteryStaticData | Select-Object -First 1
$o.batt=@([pscustomobject]@{Full=$f.FullChargedCapacity;Design=$d.DesignedCapacity;Name=$d.DeviceName})
$o | ConvertTo-Json -Depth 4 -Compress
"""
SMI_KEYS = ["index", "name", "temperature.gpu", "fan.speed", "utilization.gpu", "utilization.memory",
            "memory.total", "memory.used", "power.draw", "power.limit", "clocks.gr", "clocks.mem", "driver_version"]
SMI = shutil.which("nvidia-smi") or next(
    (p for p in (r"C:\Windows\System32\nvidia-smi.exe", r"C:\Program Files\NVIDIA Corporation\NVSMI\nvidia-smi.exe")
     if os.path.exists(p)), None)


def read_smi():
    if not SMI:
        return []
    rows = []
    for line in run([SMI, "--query-gpu=" + ",".join(SMI_KEYS), "--format=csv,noheader,nounits"], 6).strip().splitlines():
        p = [x.strip() for x in line.split(",")]
        if len(p) >= len(SMI_KEYS):
            rows.append({k: (num_of(x) if k not in ("name", "driver_version") else x) for k, x in zip(SMI_KEYS, p)})
    return rows


def type_from_img(img):
    i = (img or "").lower()
    for key, t in (("nvidia", "GpuNvidia"), ("ati", "GpuAmd"), ("amd", "GpuAmd"), ("cpu", "Cpu"), ("hdd", "Storage"),
                   ("ssd", "Storage"), ("mainboard", "Motherboard"), ("chip", "SuperIO"), ("ram", "Memory")):
        if key in i:
            return t
    return ""


def parse_lhm(j):
    out = []

    def walk(n, path, imgs):
        txt, ch = n.get("Text", ""), n.get("Children") or []
        if ch:
            for c in ch:
                walk(c, path + [txt], imgs + [n.get("ImageURL", "")])
        elif n.get("Value") not in (None, ""):
            out.append(dict(hw=path[-2] if len(path) >= 2 else "", group=path[-1] if path else "", name=txt,
                            type=type_from_img(imgs[-2]) if len(imgs) >= 2 else "",
                            value=n.get("Value"), min=n.get("Min") or "", max=n.get("Max") or ""))
    walk(j, [], [])
    return out


def read_lhm():
    try:
        with urllib.request.urlopen(LHM_URL, timeout=1.5) as r:
            return parse_lhm(json.loads(r.read().decode("utf-8", "replace")))
    except Exception:
        return None


def read_acpi():
    out = ps("(Get-CimInstance -Namespace root/wmi -ClassName MSAcpi_ThermalZoneTemperature | ForEach-Object "
             "{[math]::Round($_.CurrentTemperature/10-273.15,1)}) -join ','", 20)
    t = [num_of(x) for x in out.strip().split(",") if x.strip()]
    return [x for x in t if x and 0 < x < 130]


def static_loop():
    while True:
        try:
            j = json.loads(ps(PS_STATIC) or "{}")
            if j:
                DATA["static"] = j
        except Exception:
            pass
        DATA["static_time"] = time.time()
        time.sleep(300)


def sensor_loop():
    n = 0
    while True:
        try:
            DATA["smi"] = read_smi()
            DATA["lhm"] = read_lhm()
            if n % 2 == 0:
                DATA["gpuwin"] = read_gpu_win()
            if not has_sensors() and n % 5 == 0 and WIN:
                DATA["acpi"] = read_acpi()
            g = gpu_summary()
            HIST["gpu"].append(g.get("util") or 0)
            HIST["temp"].append(hottest())
        except Exception:
            pass
        n += 1
        time.sleep(3)


def fast_loop():
    pio, pnet, tp, n = psutil.disk_io_counters(), psutil.net_io_counters(pernic=True), time.time(), 0
    psutil.cpu_percent(percpu=True)
    while True:
        time.sleep(1)
        try:
            now = time.time()
            dt, tp = max(0.2, now - tp), now
            cores = psutil.cpu_percent(percpu=True)
            tot = sum(cores) / max(1, len(cores))
            fr, vm, sw = psutil.cpu_freq(), psutil.virtual_memory(), psutil.swap_memory()
            io, net = psutil.disk_io_counters(), psutil.net_io_counters(pernic=True)
            rd = (io.read_bytes - pio.read_bytes) / dt if io and pio else 0
            wr = (io.write_bytes - pio.write_bytes) / dt if io and pio else 0
            nets = {k: ((x.bytes_recv - pnet[k].bytes_recv) / dt, (x.bytes_sent - pnet[k].bytes_sent) / dt)
                    for k, x in net.items() if k in pnet}
            pio, pnet = io, net
            LIVE.update(cpu=tot, cores=cores, freq=fr.current if fr else None, fmax=fr.max if fr else None,
                        vm=vm, sw=sw, read=rd, write=wr, nets=nets, boot=psutil.boot_time(),
                        bat=psutil.sensors_battery())
            HIST["cpu"].append(tot)
            HIST["ram"].append(vm.percent)
            HIST["down"].append(sum(x[0] for x in nets.values()))
            HIST["up"].append(sum(x[1] for x in nets.values()))
            HIST["read"].append(rd)
            HIST["write"].append(wr)
            if n % 2 == 0:
                procs = []
                for p in psutil.process_iter(["pid", "name"]):
                    try:
                        procs.append((p.info["name"] or "?", p.info["pid"], p.cpu_percent(None) / (psutil.cpu_count() or 1),
                                      p.memory_info().rss / 1048576))
                    except Exception:
                        pass
                LIVE["procs"] = procs
            if n % 10 == 0:
                vols = []
                for part in psutil.disk_partitions(all=False):
                    if "cdrom" in part.opts or not part.fstype:
                        continue
                    try:
                        u = psutil.disk_usage(part.mountpoint)
                        vols.append((part.mountpoint, part.fstype, u.used, u.total, u.percent))
                    except Exception:
                        pass
                VOLS[:] = vols
        except Exception:
            pass
        n += 1


# ---------------------------------------------------------------- in-app sensor engine (LibreHardwareMonitor)
GH_API = "https://api.github.com/repos/LibreHardwareMonitor/LibreHardwareMonitor/releases/latest"
_CLR = [None]
UNITS = {"Temperature": "°C", "Fan": "RPM", "Load": "%", "Clock": "MHz", "Power": "W", "Voltage": "V", "Control": "%",
         "Data": "GB", "SmallData": "MB", "Throughput": "B/s", "Level": "%", "Energy": "mWh", "Current": "A"}
GROUPS = {"Temperature": "Temperatures", "Fan": "Fans", "Clock": "Clocks", "Voltage": "Voltages", "Control": "Controls",
          "Level": "Levels", "Current": "Currents", "SmallData": "Data"}


def has_pythonnet():
    if not _CLR[0]:
        try:
            import clr  # noqa: F401
            _CLR[0] = True
        except Exception:
            return False
    return True


def find_lib():
    h = glob.glob(os.path.join(ENGINE_DIR, "**", "LibreHardwareMonitorLib.dll"), recursive=True)
    return h[0] if h else None


def find_exe():
    h = glob.glob(os.path.join(ENGINE_DIR, "**", "LibreHardwareMonitor.exe"), recursive=True)
    return h[0] if h else None


def write_lhm_config(exe):
    cfg = os.path.splitext(exe)[0] + ".config"
    try:
        with open(cfg, "w", encoding="utf-8") as f:
            f.write('<?xml version="1.0" encoding="utf-8"?>\n<configuration><appSettings>\n'
                    '<add key="startMinMenuItem" value="true" />\n<add key="minTrayMenuItem" value="true" />\n'
                    '<add key="minCloseMenuItem" value="true" />\n<add key="runWebServerMenuItem" value="true" />\n'
                    '<add key="listenerPort" value="8085" />\n</appSettings></configuration>\n')
    except Exception:
        pass


def set_setup(msg):
    DATA["setup"] = msg


def install_engine():
    try:
        os.makedirs(ENGINE_DIR, exist_ok=True)
        if not find_lib():
            set_setup("Finding the latest LibreHardwareMonitor release…")
            hdr = {"User-Agent": "PCDet"}
            rel = json.loads(urllib.request.urlopen(urllib.request.Request(GH_API, headers=hdr), timeout=20).read().decode())
            zips = [a for a in rel.get("assets", []) if a["name"].lower().endswith(".zip")]
            pick = (next((a for a in zips if a["name"].lower() == "librehardwaremonitor.zip"), None)
                    or next((a for a in zips if "net" not in a["name"].lower()), None) or (zips[0] if zips else None))
            if not pick:
                raise RuntimeError("no zip found in the latest release")
            set_setup(f"Downloading {pick['name']}…")
            zp = os.path.join(ENGINE_DIR, "lhm.zip")
            with urllib.request.urlopen(urllib.request.Request(pick["browser_download_url"], headers=hdr), timeout=180) as r, \
                    open(zp, "wb") as f:
                shutil.copyfileobj(r, f)
            set_setup("Extracting…")
            with zipfile.ZipFile(zp) as z:
                z.extractall(ENGINE_DIR)
            os.remove(zp)
        exe = find_exe()
        if exe:
            write_lhm_config(exe)
        if not has_pythonnet():
            set_setup("Installing pythonnet (the .NET bridge)…")
            run([sys.executable, "-m", "pip", "install", "pythonnet"], 400)
            importlib.invalidate_caches()
        if find_lib() and has_pythonnet():
            set_setup("✅ Sensor engine ready. Click 'Restart as Administrator' to start reading fans and temperatures.")
        elif find_lib():
            set_setup("Engine files downloaded, but pythonnet could not be installed (very new Python?). "
                      "Click 'Launch LibreHardwareMonitor (admin)' instead — PCDet will connect to it.")
        else:
            set_setup("Download finished but LibreHardwareMonitorLib.dll was not found in the package.")
    except Exception as e:
        set_setup(f"Setup failed: {e}")


def launch_lhm():
    exe = find_exe()
    if not exe:
        set_setup("Run 'One-click setup' first so LibreHardwareMonitor gets downloaded.")
        return
    write_lhm_config(exe)
    if WIN:
        import ctypes
        ctypes.windll.shell32.ShellExecuteW(None, "runas", exe, None, os.path.dirname(exe), 1)
        set_setup("LibreHardwareMonitor started (it sits in the tray). PCDet connects in a few seconds.")


def open_computer():
    import clr
    from System.Reflection import Assembly
    lib = find_lib()
    d = os.path.dirname(lib)
    sys.path.append(d)
    for f in sorted(glob.glob(os.path.join(d, "*.dll"))):
        if f != lib:
            try:
                Assembly.LoadFrom(f)
            except Exception:
                pass
    clr.AddReference(lib)
    from LibreHardwareMonitor.Hardware import Computer
    c = Computer()
    for flag in ("IsCpuEnabled", "IsGpuEnabled", "IsMemoryEnabled", "IsMotherboardEnabled", "IsControllerEnabled",
                 "IsStorageEnabled", "IsBatteryEnabled", "IsPsuEnabled"):
        try:
            setattr(c, flag, True)
        except Exception:
            pass
    c.Open()
    return c


def read_computer(c):
    out = []

    def visit(hw):
        hw.Update()
        for s in hw.Sensors:
            val = s.Value
            if val is None:
                continue
            st = str(s.SensorType)
            u = UNITS.get(st, "")

            def fmt(x):
                if x is None:
                    return ""
                return (f"{x:.0f} {u}" if st in ("Fan", "Clock", "Load", "Control", "Level") else f"{x:.1f} {u}").strip()
            out.append(dict(hw=str(hw.Name), type=str(hw.HardwareType), group=GROUPS.get(st, st), name=str(s.Name),
                            value=fmt(val), min=fmt(s.Min), max=fmt(s.Max)))
        for sub in hw.SubHardware:
            visit(sub)
    for hw in c.Hardware:
        visit(hw)
    return out


def hw_loop():
    comp = None
    while True:
        try:
            if comp is None:
                if not find_lib():
                    DATA["engine"] = dict(state="not installed", msg="")
                elif not has_pythonnet():
                    DATA["engine"] = dict(state="pythonnet missing", msg="use the LibreHardwareMonitor app instead")
                else:
                    comp = open_computer()
            if comp is not None:
                DATA["hw"] = read_computer(comp)
                DATA["engine"] = dict(state="running", msg="")
                time.sleep(2)
                continue
        except Exception as e:
            comp, DATA["hw"] = None, None
            DATA["engine"] = dict(state="error", msg=str(e)[:140])
        time.sleep(5)


def read_gpu_win():
    """Any-brand GPU load + VRAM from Windows performance counters."""
    if not WIN:
        return {}
    out = run(["typeperf", r"\GPU Engine(*)\Utilization Percentage", r"\GPU Adapter Memory(*)\Dedicated Usage",
               "-sc", "1"], 20)
    lines = [ln for ln in out.splitlines() if ln.startswith('"')]
    if len(lines) < 2:
        return {}
    head, vals = next(csv.reader([lines[0]])), next(csv.reader([lines[1]]))
    eng, vram = {}, 0.0
    for h, x in zip(head[1:], vals[1:]):
        try:
            val = float(x)
        except ValueError:
            continue
        m = re.search(r"engtype_(\w+)", h)
        if "Utilization" in h and m:
            eng[m.group(1)] = eng.get(m.group(1), 0) + val
        elif "Dedicated Usage" in h:
            vram += val
    return dict(util=min(100.0, max(eng.values())) if eng else None, vram=vram or None)


def disk_temp_live(name):
    n = (name or "").lower()
    toks = {t for t in re.split(r"\W+", n) if len(t) > 2}
    for s in lhm():
        if s["group"] != "Temperatures" or s.get("type") != "Storage":
            continue
        h = s["hw"].lower()
        if h and (h in n or n in h or len({t for t in re.split(r"\W+", h) if len(t) > 2} & toks) >= 2):
            x = num_of(s["value"])
            if x is not None:
                return x
    return None


def ssd_temp():
    t = [d["TempC"] for d in S("disks") if d.get("TempC")]
    t += [num_of(s["value"]) for s in lhm() if s["group"] == "Temperatures" and s.get("type") == "Storage"]
    t = [x for x in t if x]
    return max(t) if t else None


def fans():
    return [(f"{s['hw']} · {s['name']}", num_of(s["value"])) for s in lhm()
            if s["group"] == "Fans" and num_of(s["value"]) is not None]


def hottest():
    vals = [x for x in (cpu_temp(), gpu_summary().get("temp"), ssd_temp()) if x]
    return max(vals) if vals else 0


def diag_rows():
    eng, fl = DATA["engine"], fans()
    temps = [s for s in lhm() if s["group"] == "Temperatures"]
    return [("Diagnostics", None),
            ("Administrator rights", "✅ yes" if is_admin() else "❌ no — CPU temps, fans & SSD data usually need it"),
            ("Sensor engine files", "✅ installed" if find_lib() else "❌ not installed (click One-click setup)"),
            ("pythonnet (.NET bridge)", "✅ available" if has_pythonnet() else "❌ not installed"),
            ("In-app engine", eng["state"] + (f" — {eng['msg']}" if eng.get("msg") else "")),
            ("LibreHardwareMonitor app", "✅ connected" if DATA.get("lhm") is not None else "not detected"),
            ("NVIDIA nvidia-smi", "✅ " + SMI if SMI else "not found"),
            ("Sensors found", f"{len(lhm())}  ({len(temps)} temperatures, {len(fl)} fans)"),
            ("Setup", DATA.get("setup") or "—")]


# ---------------------------------------------------------------- sensor interpretation
def lhm():
    return DATA.get("hw") or DATA.get("lhm") or []


def has_sensors():
    return bool(DATA.get("hw")) or DATA.get("lhm") is not None


def is_gpu_hw(h):
    return any(w in h.lower() for w in ("nvidia", "geforce", "radeon", "intel arc", "gpu", "graphics", "uhd", "iris"))


def is_cpu_hw(h):
    return not is_gpu_hw(h) and any(w in h.lower() for w in ("intel", "amd", "ryzen", "core i", "xeon", "athlon", "cpu"))


def is_gpu_s(s):
    return s["type"].startswith("Gpu") if s.get("type") else is_gpu_hw(s["hw"])


def is_cpu_s(s):
    return s["type"] == "Cpu" if s.get("type") else is_cpu_hw(s["hw"])


def cpu_temp():
    temps = [(s["name"].lower(), num_of(s["value"])) for s in lhm() if s["group"] == "Temperatures" and is_cpu_s(s)]
    temps = [(n, x) for n, x in temps if x is not None]
    for key in ("package", "tctl", "tdie", "cpu"):
        for n, x in temps:
            if key in n:
                return x
    if temps:
        return max(x for _, x in temps)
    return max(DATA["acpi"]) if DATA.get("acpi") else None


def gpu_summary():
    g = first(DATA.get("smi") or [])
    d = dict(name=g.get("name"), util=g.get("utilization.gpu"), temp=g.get("temperature.gpu"), fan=g.get("fan.speed"),
             power=g.get("power.draw"), vu=g.get("memory.used"), vt=g.get("memory.total"))
    for s in lhm():
        if not is_gpu_s(s):
            continue
        x, n = num_of(s["value"]), s["name"].lower()
        if s["group"] == "Temperatures" and "core" in n and d["temp"] is None:
            d["temp"] = x
        if s["group"] == "Load" and n == "gpu core" and d["util"] is None:
            d["util"] = x
        if s["group"] == "Fans" and d["fan"] is None:
            d["fan"] = f"{x:.0f} RPM" if x is not None else None
    gw = DATA.get("gpuwin") or {}
    if d["util"] is None and gw.get("util") is not None:
        d["util"] = gw["util"]
    if d["vu"] is None and gw.get("vram"):
        d["vu"] = gw["vram"] / 1048576
    return d


def alerts():
    a, vm, now = [], LIVE.get("vm"), time.time()
    t = cpu_temp()
    if t and t >= 90:
        a.append(("🔴", f"CPU is very hot ({t:.0f}°C)"))
    elif t and t >= 80:
        a.append(("🟠", f"CPU is running warm ({t:.0f}°C)"))
    gt = gpu_summary().get("temp")
    if gt and gt >= 85:
        a.append(("🟠", f"GPU is hot ({gt:.0f}°C)"))
    if vm and vm.percent >= 90:
        a.append(("🟠", f"Memory usage is high ({vm.percent:.0f}%)"))
    for m, _, _, _, pct in VOLS:
        if pct >= 90:
            a.append(("🟠", f"Drive {m} is {pct:.0f}% full"))
    for d in S("disks"):
        if d.get("Health") not in ("Healthy", "", None):
            a.append(("🔴", f"Disk '{d.get('Name')}' reports health: {d.get('Health')}"))
        if (d.get("Wear") or 0) >= 80:
            a.append(("🟠", f"Disk '{d.get('Name')}' is {d['Wear']}% worn"))
        dt_ = d.get("TempC") or disk_temp_live(d.get("Name"))
        if (dt_ or 0) >= 60:
            a.append(("🟠", f"Disk '{d.get('Name')}' is hot ({dt_:.0f}°C)"))
    b = first(S("batt"))
    if b.get("Full") and b.get("Design") and b["Full"] / b["Design"] < 0.6:
        a.append(("🟠", f"Battery health is low ({b['Full'] / b['Design'] * 100:.0f}% of original capacity)"))
    if LIVE.get("boot") and now - LIVE["boot"] > 14 * 86400:
        a.append(("🟡", f"PC hasn't restarted for {dur(now - LIVE['boot'])}"))
    return a or [("✅", "Everything looks healthy")]


# ---------------------------------------------------------------- detail rows (also used for the report)
def rows_cpu():
    c, r = first(S("cpu")), [("Processor", None)]
    if c:
        r += [("Name", v(c, "Name")), ("Manufacturer", v(c, "Maker")), ("Cores / Threads", f"{v(c, 'Cores')} / {v(c, 'Threads')}"),
              ("Max clock", v(c, "MaxMHz", " MHz")), ("L2 / L3 cache", f"{v(c, 'L2KB')} KB / {v(c, 'L3KB')} KB"),
              ("Socket", v(c, "Socket")), ("Virtualization", "Enabled" if c.get("Virt") else "Disabled / unknown")]
    else:
        r += [("Name", platform.processor() or "loading…"), ("Cores / Threads", f"{psutil.cpu_count(False)} / {psutil.cpu_count()}")]
    t = cpu_temp()
    r += [("Live", None), ("Usage", f"{LIVE.get('cpu', 0):.0f}%"),
          ("Current clock", f"{LIVE['freq']:.0f} MHz" if LIVE.get("freq") else "—"),
          ("Temperature", f"{t:.0f} °C" if t else "— (run LibreHardwareMonitor for CPU temps)")]
    for s in lhm():
        if is_cpu_s(s) and s["group"] in ("Temperatures", "Clocks", "Power", "Voltages"):
            r.append((f"{s['group'][:-1] if s['group'].endswith('s') else s['group']}: {s['name']}", s["value"]))
    return r


def rows_gpu():
    r = []
    for g in S("gpu"):
        vr = g.get("VRAM")
        r += [(v(g, "Name"), None), ("Driver", f"{v(g, 'Driver')}  ({v(g, 'DriverDate')})"),
              ("Display", f"{v(g, 'W')}×{v(g, 'H')} @ {v(g, 'Hz')} Hz"),
              ("Video memory (Windows)", gb(vr) + ("  — capped at 4 GB by Windows, see live data" if vr and vr >= 4294967295 else ""))]
    for g in DATA.get("smi") or []:
        r += [(f"NVIDIA live · {g['name']}", None), ("GPU load", f"{v(g, 'utilization.gpu', '%')}"),
              ("VRAM", f"{v(g, 'memory.used', ' MB')} / {v(g, 'memory.total', ' MB')}"),
              ("Temperature", v(g, "temperature.gpu", " °C")), ("Fan speed", v(g, "fan.speed", " %")),
              ("Power", f"{v(g, 'power.draw', ' W')} / {v(g, 'power.limit', ' W')}"),
              ("Clocks (core / mem)", f"{v(g, 'clocks.gr', ' MHz')} / {v(g, 'clocks.mem', ' MHz')}")]
    gs = [s for s in lhm() if is_gpu_s(s)]
    if gs:
        r.append(("LibreHardwareMonitor GPU sensors", None))
        r += [(f"{s['hw']} · {s['group']}: {s['name']}", s["value"]) for s in gs]
    gw = DATA.get("gpuwin") or {}
    if gw.get("util") is not None or gw.get("vram"):
        r += [("Windows GPU counters (any brand)", None),
              ("GPU load", f"{gw['util']:.0f}%" if gw.get("util") is not None else "—"), ("VRAM in use", gb(gw.get("vram")))]
    if not r:
        r = [("GPU", None), ("Info", "loading…")]
    return r


def rows_memory():
    vm, sw, r = LIVE.get("vm"), LIVE.get("sw"), [("Memory", None)]
    if vm:
        r += [("Total", gb(vm.total)), ("In use", f"{gb(vm.used)}  ({vm.percent:.0f}%)  {textbar(vm.percent)}"),
              ("Available", gb(vm.available))]
    if sw:
        r += [("Page file / swap", f"{gb(sw.used)} / {gb(sw.total)}  ({sw.percent:.0f}%)")]
    mods = S("ram")
    if mods:
        r.append((f"Installed modules ({len(mods)})", None))
        for m in mods:
            mhz = m.get("Configured") or m.get("MHz")
            r.append((v(m, "Slot"), f"{gb(m.get('Bytes'))} {RAM_TYPES.get(m.get('Type'), '')} @ {mhz} MHz · "
                                    f"{v(m, 'Maker')} · {v(m, 'Part')}"))
    return r


def rows_storage():
    r, disks = [], S("disks")
    for d in disks:
        r.append((f"{v(d, 'Name')}  ({v(d, 'Media')}, {v(d, 'Bus')}, {v(d, 'SizeGB', ' GB')})", None))
        r += [("Health", f"{v(d, 'Health')}  /  {v(d, 'Status')}")]
        if d.get("Wear") is not None:
            r.append(("Life remaining", f"≈ {100 - d['Wear']}%  (wear {d['Wear']}%)"))
        r += [("Temperature", f"{d.get('TempC') or disk_temp_live(d.get('Name'))} °C"
                            if (d.get("TempC") or disk_temp_live(d.get("Name"))) else "— (run as admin / sensor engine needed)"),
              ("Power-on time", f"{d['Hours']} h  (~{d['Hours'] / 24:.0f} days)" if d.get("Hours") is not None else "—"),
              ("Read / write errors", f"{v(d, 'ReadErr')} / {v(d, 'WriteErr')}"), ("Start/stop cycles", v(d, "Cycles")),
              ("Serial", v(d, "Serial"))]
        if all(d.get(k) is None for k in ("TempC", "Wear", "Hours")):
            r.append(("Note", "Detailed SMART data not available — try running as administrator"))
    if not disks:
        r.append(("Physical disks", None))
        r.append(("Info", "loading… (or Windows Storage module unavailable)"))
    ss = [x for x in lhm() if x.get("type") == "Storage"]
    if ss:
        r.append(("Drive sensors (live)", None))
        r += [(f"{x['hw']} · {x['name']}", x["value"]) for x in ss]
    r.append(("Volumes", None))
    for m, fs, used, tot, pct in VOLS:
        r.append((f"{m}  ({fs})", f"{gb(used)} / {gb(tot)}   {pct:.0f}%  {textbar(pct)}"))
    r += [("Disk activity", None), ("Read", rate(LIVE.get("read", 0))), ("Write", rate(LIVE.get("write", 0)))]
    return r


def rows_network():
    r, stats, addrs = [], psutil.net_if_stats(), psutil.net_if_addrs()
    for name, st in stats.items():
        if name.lower().startswith("loopback") or name == "lo":
            continue
        r.append((name, None))
        r.append(("Status", f"{'Connected' if st.isup else 'Disconnected'}" + (f" · {st.speed} Mbps" if st.speed else "")))
        for a in addrs.get(name, []):
            lab = {2: "IPv4", 23: "IPv6", 17: "MAC", 10: "IPv6", -1: "MAC"}.get(int(a.family), None)
            if lab:
                r.append((lab, a.address))
        d, u = LIVE.get("nets", {}).get(name, (0, 0))
        r.append(("Down / Up", f"{rate(d)} / {rate(u)}"))
    return r or [("Network", None), ("Info", "no adapters found")]


def rows_system():
    c, b, bios, o = first(S("cs")), first(S("board")), first(S("bios")), first(S("os"))
    boot = LIVE.get("boot")
    r = [("Computer", None), ("Name", c.get("Name") or socket.gethostname()), ("Maker / model", f"{v(c, 'Maker')} {v(c, 'Model')}"),
         ("System type", v(c, "Type")), ("Uptime", dur(time.time() - boot) if boot else "—"),
         ("Last boot", time.strftime("%Y-%m-%d %H:%M", time.localtime(boot)) if boot else "—"),
         ("Operating system", None), ("Name", o.get("Name") or platform.system()),
         ("Version / build", f"{o.get('Version') or platform.version()} / {v(o, 'Build')}"),
         ("Architecture", o.get("Arch") or platform.machine()), ("Installed on", v(o, "Installed")),
         ("Motherboard", None), ("Maker / model", f"{v(b, 'Maker')} {v(b, 'Model')}"), ("Version", v(b, "Version")),
         ("Serial", v(b, "Serial")), ("BIOS", None), ("Maker", v(bios, "Maker")), ("Version", v(bios, "Version")),
         ("Date", v(bios, "Date"))]
    bat, bh = LIVE.get("bat"), first(S("batt"))
    if bat:
        r += [("Battery", None), ("Charge", f"{bat.percent:.0f}%  ({'charging / plugged in' if bat.power_plugged else 'on battery'})")]
        if bat.secsleft and bat.secsleft > 0 and not bat.power_plugged:
            r.append(("Time left", dur(bat.secsleft)))
        if bh.get("Full") and bh.get("Design"):
            r.append(("Battery health", f"{bh['Full'] / bh['Design'] * 100:.0f}%  ({bh['Full']} / {bh['Design']} mWh)"))
    return r


def sensor_rows():
    rows = [("LibreHardwareMonitor", s["hw"], s["group"], s["name"], s["value"], s["min"], s["max"]) for s in lhm()]
    if not rows:
        for g in DATA.get("smi") or []:
            rows += [("nvidia-smi", g["name"], "Temperatures", "GPU", f"{g['temperature.gpu']} °C", "", ""),
                     ("nvidia-smi", g["name"], "Fans", "GPU fan", f"{g['fan.speed']} %", "", "")]
        rows += [("ACPI", "Thermal zone", "Temperatures", f"Zone {i}", f"{t} °C", "", "") for i, t in enumerate(DATA.get("acpi") or [])]
    return rows


def report_text():
    parts = [f"PCDet report — {time.strftime('%Y-%m-%d %H:%M')}", "=" * 50, "ALERTS"]
    parts += [f"  {i} {t}" for i, t in alerts()]
    for title, fn in (("CPU", rows_cpu), ("GPU", rows_gpu), ("MEMORY", rows_memory), ("STORAGE", rows_storage),
                      ("NETWORK", rows_network), ("SYSTEM", rows_system)):
        parts += ["", f"[{title}]"]
        parts += [f"  {k}" if val is None else f"    {k}: {val}" for k, val in fn()]
    sr = sensor_rows()
    if sr:
        parts += ["", "[SENSORS]"] + [f"  {a[1]} / {a[2]} / {a[3]}: {a[4]}" for a in sr]
    return "\n".join(parts)


# ---------------------------------------------------------------- widgets
A1, A2 = "#00d4ff", "#8a4dff"


class Bar(QWidget):
    def __init__(self):
        super().__init__()
        self.setFixedHeight(16)
        self.val, self.text = 0.0, ""

    def set(self, x, text=""):
        self.val, self.text = max(0.0, min(1.0, x)), text
        self.update()

    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        r = QRectF(self.rect()).adjusted(0, 1, 0, -1)
        p.setPen(Qt.NoPen)
        p.setBrush(QColor("#23253a"))
        p.drawRoundedRect(r, r.height() / 2, r.height() / 2)
        if self.val > 0:
            fr = QRectF(r.left(), r.top(), max(r.height(), r.width() * self.val), r.height())
            g = QLinearGradient(fr.topLeft(), fr.topRight())
            g.setColorAt(0, QColor("#ff9a3c" if self.val > .9 else A1))
            g.setColorAt(1, QColor("#ff3b3b" if self.val > .9 else A2))
            p.setBrush(QBrush(g))
            p.drawRoundedRect(fr, r.height() / 2, r.height() / 2)
        p.setPen(QColor("white"))
        p.setFont(QFont("Segoe UI", 7))
        p.drawText(r, Qt.AlignCenter, self.text)


class Spark(QWidget):
    def __init__(self, color):
        super().__init__()
        self.setFixedHeight(42)
        self.color, self.vals = QColor(color), []

    def set(self, vals, mx=None):
        m = mx or (max(vals) if vals else 1) or 1
        self.vals = [x / m for x in vals]
        self.update()

    def paintEvent(self, _):
        if len(self.vals) < 2:
            return
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        w, h = self.width(), self.height()
        step = w / 59
        x0 = w - step * (len(self.vals) - 1)
        pts = [QPointF(x0 + i * step, h - 2 - min(1, x) * (h - 6)) for i, x in enumerate(self.vals)]
        path = QPainterPath(pts[0])
        for pt in pts[1:]:
            path.lineTo(pt)
        p.setBrush(Qt.NoBrush)
        p.setPen(QPen(self.color, 2))
        p.drawPath(path)
        path.lineTo(pts[-1].x(), h)
        path.lineTo(pts[0].x(), h)
        g = QLinearGradient(0, 0, 0, h)
        c1, c0 = QColor(self.color), QColor(self.color)
        c1.setAlpha(90)
        c0.setAlpha(0)
        g.setColorAt(0, c1)
        g.setColorAt(1, c0)
        p.setPen(Qt.NoPen)
        p.setBrush(QBrush(g))
        p.drawPath(path)


class Tile(QFrame):
    def __init__(self, title, color):
        super().__init__()
        self.setObjectName("card")
        lay = QVBoxLayout(self)
        t = QLabel(title)
        t.setStyleSheet("color:#8d90b3;font-weight:bold;letter-spacing:1px")
        self.big, self.sub, self.spark = QLabel("—"), QLabel(""), Spark(color)
        self.big.setStyleSheet("font-size:30px;font-weight:600")
        self.sub.setStyleSheet("color:#aeb1d4")
        self.sub.setWordWrap(True)
        for w in (t, self.big, self.sub, self.spark):
            lay.addWidget(w)

    def set(self, big, sub, hist, mx=None):
        self.big.setText(big)
        self.sub.setText(sub)
        self.spark.set(list(hist), mx)


class KV(QWidget):
    def __init__(self):
        super().__init__()
        self.g = QGridLayout(self)
        self.g.setContentsMargins(0, 0, 0, 0)
        self.g.setHorizontalSpacing(20)
        self.g.setVerticalSpacing(5)
        self.g.setColumnStretch(1, 1)
        self.rows = []

    def set(self, rows):
        while len(self.rows) < len(rows):
            a, b = QLabel(), QLabel()
            b.setWordWrap(True)
            b.setTextInteractionFlags(Qt.TextSelectableByMouse)
            a.setAlignment(Qt.AlignTop)
            i = len(self.rows)
            self.g.addWidget(a, i, 0)
            self.g.addWidget(b, i, 1)
            self.rows.append((a, b))
        for i, (a, b) in enumerate(self.rows):
            if i >= len(rows):
                a.hide()
                b.hide()
                continue
            k, val = rows[i]
            a.setText(k)
            b.setText("" if val is None else str(val))
            a.setStyleSheet("color:#9be7ff;font-weight:bold;padding-top:12px" if val is None else "color:#8d90b3")
            a.show()
            b.show()


def scrolled(w):
    s = QScrollArea()
    s.setWidgetResizable(True)
    s.setFrameShape(QFrame.NoFrame)
    s.setWidget(w)
    return s


STYLE = """
QWidget{background:#10111a;color:#e8e9f5;font-family:"Segoe UI";font-size:13px}
#side{background:#0b0c13}
#title{font-size:22px;font-weight:bold;color:#9be7ff;padding:12px 10px}
#h{font-size:22px;font-weight:bold}
#card{background:#171928;border:1px solid #262944;border-radius:16px}
QListWidget{background:transparent;border:none;outline:none}
QListWidget::item{padding:11px 14px;border-radius:10px;margin:2px 6px}
QListWidget::item:selected{background:#1f2140;color:white}
QPushButton{background:#1d1f33;border:1px solid #33365a;border-radius:9px;padding:8px 14px}
QPushButton:hover{background:#262948}
QTableWidget{background:#171928;border:none;gridline-color:#23253a}
QHeaderView::section{background:#1d1f33;border:none;padding:6px;color:#9be7ff}
QPlainTextEdit{background:#0c0d15;border:1px solid #262944;border-radius:10px;font-family:Consolas,monospace;font-size:12px}
QScrollBar:vertical{background:#10111a;width:10px}QScrollBar::handle:vertical{background:#2a2c40;border-radius:5px}
"""
NAMES = ["🏠  Overview", "🧠  CPU", "🎮  GPU", "💾  Memory", "🗄️  Storage", "🌐  Network", "🌡️  Sensors & Fans",
         "📋  Processes", "🖥️  System", "📄  Report"]


def is_admin():
    try:
        return bool(WIN and __import__("ctypes").windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


class Main(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PCDet")
        self.resize(1040, 700)
        self.setStyleSheet(STYLE)
        root = QHBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)
        side = QWidget()
        side.setObjectName("side")
        side.setFixedWidth(220)
        sl = QVBoxLayout(side)
        t = QLabel("🖥️ PCDet")
        t.setObjectName("title")
        sl.addWidget(t)
        self.nav = QListWidget()
        for n in NAMES:
            self.nav.addItem(n)
        sl.addWidget(self.nav)
        self.status = QLabel()
        self.status.setStyleSheet("color:#8d90b3;font-size:11px;padding:6px")
        self.status.setWordWrap(True)
        sl.addWidget(self.status)
        self.stack = QStackedWidget()
        root.addWidget(side)
        root.addWidget(self.stack, 1)
        self.refreshers = {}
        for i, b in enumerate((self.p_overview, self.p_kv(rows_cpu, True), self.p_kv(rows_gpu), self.p_kv(rows_memory),
                               self.p_kv(rows_storage), self.p_kv(rows_network), self.p_sensors, self.p_procs,
                               self.p_kv(rows_system), self.p_report)):
            b(i)
        self.nav.currentRowChanged.connect(self.go)
        self.nav.setCurrentRow(0)
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.tick)
        self.timer.start(1000)

    def go(self, i):
        self.stack.setCurrentIndex(i)
        self.refreshers[i]()

    def tick(self):
        self.status.setText(f"● Live · {'Admin' if is_admin() else 'Standard user'}\n"
                            f"Sensors: {len(lhm())}  ·  engine: {DATA['engine']['state']}")
        self.refreshers[self.stack.currentIndex()]()

    def page(self, title):
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.setContentsMargins(26, 22, 26, 22)
        h = QLabel(title)
        h.setObjectName("h")
        lay.addWidget(h)
        self.stack.addWidget(scrolled(w))
        return lay

    def p_kv(self, fn, cores=False):
        def build(i):
            lay = self.page(NAMES[i][3:].strip())
            card = QFrame()
            card.setObjectName("card")
            cl = QVBoxLayout(card)
            kv = KV()
            cl.addWidget(kv)
            lay.addWidget(card)
            bars = []
            if cores:
                cc = QFrame()
                cc.setObjectName("card")
                grid = QGridLayout(cc)
                for c in range(psutil.cpu_count() or 1):
                    lab, bar = QLabel(f"Core {c}"), Bar()
                    grid.addWidget(lab, c // 2, (c % 2) * 2)
                    grid.addWidget(bar, c // 2, (c % 2) * 2 + 1)
                    bars.append(bar)
                lay.addWidget(cc)
            lay.addStretch()

            def refresh():
                kv.set(fn())
                for b, x in zip(bars, LIVE.get("cores", [])):
                    b.set(x / 100, f"{x:.0f}%")
            self.refreshers[i] = refresh
        return build

    def p_overview(self, i):
        lay = self.page("Overview")
        self.sub = QLabel()
        self.sub.setStyleSheet("color:#aeb1d4")
        lay.addWidget(self.sub)
        grid = QGridLayout()
        self.tiles = {k: Tile(t, c) for k, t, c in (("cpu", "CPU", "#00d4ff"), ("gpu", "GPU", "#8a4dff"),
                                                    ("ram", "MEMORY", "#00ff9d"), ("disk", "STORAGE", "#ff9a3c"),
                                                    ("net", "NETWORK", "#ff4fd8"), ("pow", "POWER / BATTERY", "#ffd84d"),
                                                    ("temps", "TEMPS & FANS", "#ff5e62"))}
        for n, t in enumerate(self.tiles.values()):
            grid.addWidget(t, n // 4, n % 4)
        lay.addLayout(grid)
        card = QFrame()
        card.setObjectName("card")
        cl = QVBoxLayout(card)
        h = QLabel("Health check")
        h.setStyleSheet("font-weight:bold;color:#9be7ff")
        self.alerts = KV()
        cl.addWidget(h)
        cl.addWidget(self.alerts)
        lay.addWidget(card)
        lay.addStretch()
        self.refreshers[i] = self.ref_overview

    def ref_overview(self):
        c, o, L, T = first(S("cs")), first(S("os")), LIVE, self.tiles
        self.sub.setText(f"{v(c, 'Maker')} {v(c, 'Model')}  ·  {o.get('Name') or platform.system()}  ·  "
                         f"up {dur(time.time() - L['boot']) if L.get('boot') else '—'}")
        t = cpu_temp()
        T["cpu"].set(f"{L.get('cpu', 0):.0f}%", "  ·  ".join(x for x in (f"{L['freq']:.0f} MHz" if L.get("freq") else "",
                                                                          f"{t:.0f}°C" if t else "") if x), HIST["cpu"], 100)
        g = gpu_summary()
        T["gpu"].set(f"{g['util']:.0f}%" if g.get("util") is not None else "—",
                     "  ·  ".join(x for x in (f"{g['temp']:.0f}°C" if g.get("temp") else "",
                                              f"fan {g['fan']}{'%' if isinstance(g['fan'], (int, float)) else ''}" if g.get("fan") else "",
                                              str(g.get("name") or "")) if x) or "no live GPU data", HIST["gpu"], 100)
        vm = L.get("vm")
        if vm:
            T["ram"].set(f"{vm.percent:.0f}%", f"{gb(vm.used)} of {gb(vm.total)}", HIST["ram"], 100)
        main = VOLS[0] if VOLS else None
        T["disk"].set(f"{main[4]:.0f}%" if main else "—", f"R {rate(L.get('read', 0))}  ·  W {rate(L.get('write', 0))}",
                      HIST["read"])
        T["net"].set(rate(L.get("nets") and sum(x[0] for x in L["nets"].values()) or 0),
                     f"▲ {rate(sum(x[1] for x in L.get('nets', {}).values()))}", HIST["down"])
        bat, bh = L.get("bat"), first(S("batt"))
        if bat:
            health = f"  ·  health {bh['Full'] / bh['Design'] * 100:.0f}%" if bh.get("Full") and bh.get("Design") else ""
            T["pow"].set(f"{bat.percent:.0f}%", ("plugged in" if bat.power_plugged else "on battery") + health, [])
        else:
            T["pow"].set("AC", "desktop / no battery", [])
        fl = fans()
        temps = [("CPU", t), ("GPU", g.get("temp")), ("SSD", ssd_temp())]
        shown = "  ·  ".join(f"{n} {x:.0f}°C" for n, x in temps if x) or "no temperature sensors yet"
        fan_txt = "Fans: " + " / ".join(f"{x:.0f}" for _, x in fl[:4]) + " RPM" if fl else "Fans: none detected"
        hot = hottest()
        T["temps"].set(f"{hot:.0f}°C" if hot else "—", f"{shown}\n{fan_txt}", HIST["temp"], 100)
        rows = alerts()
        if not has_sensors() and not DATA.get("smi"):
            rows.append(("💡", "Fan speeds & more temperatures need the sensor engine: Sensors & Fans → One-click setup"))
        self.alerts.set(rows)

    def p_sensors(self, i):
        lay = self.page("Sensors & Fans")
        card = QFrame()
        card.setObjectName("card")
        cl = QVBoxLayout(card)
        self.diag = KV()
        cl.addWidget(self.diag)
        lay.addWidget(card)
        row = QHBoxLayout()
        btns = [("⚡ One-click setup", lambda: threading.Thread(target=install_engine, daemon=True).start()),
                ("Launch LibreHardwareMonitor (admin)", launch_lhm)]
        if WIN and not is_admin():
            btns.insert(1, ("🛡️ Restart as Administrator", self.as_admin))
        for text, fn in btns:
            b = QPushButton(text)
            b.clicked.connect(fn)
            row.addWidget(b)
        row.addStretch()
        lay.addLayout(row)
        self.table = QTableWidget(0, 6)
        self.table.setHorizontalHeaderLabels(["Source / hardware", "Group", "Sensor", "Value", "Min", "Max"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setMinimumHeight(420)
        lay.addWidget(self.table)
        self.refreshers[i] = self.ref_sensors

    def as_admin(self):
        if WIN:
            import ctypes
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, f'"{os.path.abspath(sys.argv[0])}"', None, 1)
            QApplication.quit()

    def ref_sensors(self):
        self.diag.set(diag_rows())
        rows = sensor_rows()
        self.table.setRowCount(len(rows))
        for r, row in enumerate(sorted(rows, key=lambda x: (x[2], x[1]))):
            for c, val in enumerate((f"{row[1]}" if row[0] == "LibreHardwareMonitor" else f"{row[0]} · {row[1]}",) + tuple(row[2:])):
                it = self.table.item(r, c)
                if it is None:
                    it = QTableWidgetItem()
                    self.table.setItem(r, c, it)
                it.setText(str(val))

    def p_procs(self, i):
        lay = self.page("Processes")
        self.sort = 0
        row = QHBoxLayout()
        for text, k in (("Sort by CPU", 0), ("Sort by memory", 1)):
            b = QPushButton(text)
            b.clicked.connect(lambda _, k=k: setattr(self, "sort", k))
            row.addWidget(b)
        row.addStretch()
        lay.addLayout(row)
        self.ptable = QTableWidget(0, 4)
        self.ptable.setHorizontalHeaderLabels(["Name", "PID", "CPU %", "Memory (MB)"])
        self.ptable.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.ptable.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.ptable.setMinimumHeight(520)
        lay.addWidget(self.ptable)
        self.refreshers[i] = self.ref_procs

    def ref_procs(self):
        procs = sorted(LIVE.get("procs", []), key=lambda p: -p[2 if self.sort == 0 else 3])[:20]
        self.ptable.setRowCount(len(procs))
        for r, (n, pid, cpu, mem) in enumerate(procs):
            for c, val in enumerate((n, pid, f"{cpu:.1f}", f"{mem:.0f}")):
                it = self.ptable.item(r, c)
                if it is None:
                    it = QTableWidgetItem()
                    self.ptable.setItem(r, c, it)
                it.setText(str(val))

    def p_report(self, i):
        lay = self.page("Report")
        row = QHBoxLayout()
        for text, fn in (("Refresh", self.fill_report), ("Copy", self.copy_report), ("Save .txt", lambda: self.save("txt")),
                         ("Save .json", lambda: self.save("json"))):
            b = QPushButton(text)
            b.clicked.connect(fn)
            row.addWidget(b)
        row.addStretch()
        lay.addLayout(row)
        self.rep = QPlainTextEdit()
        self.rep.setReadOnly(True)
        self.rep.setMinimumHeight(540)
        lay.addWidget(self.rep)
        self.refreshers[i] = lambda: None if self.rep.toPlainText() else self.fill_report()

    def fill_report(self):
        self.rep.setPlainText(report_text())

    def copy_report(self):
        QApplication.clipboard().setText(report_text())

    def save(self, kind):
        path, _ = QFileDialog.getSaveFileName(self, "Save report", f"pcdet-report.{kind}", f"*.{kind}")
        if not path:
            return
        with open(path, "w", encoding="utf-8") as f:
            if kind == "txt":
                f.write(report_text())
            else:
                json.dump(dict(static=DATA["static"], smi=DATA["smi"], sensors=lhm(), acpi=DATA["acpi"],
                               volumes=VOLS, alerts=alerts()), f, indent=2, default=str)


def main():
    for fn in (static_loop, sensor_loop, fast_loop, hw_loop):
        threading.Thread(target=fn, daemon=True).start()
    app = QApplication(sys.argv)
    w = Main()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
