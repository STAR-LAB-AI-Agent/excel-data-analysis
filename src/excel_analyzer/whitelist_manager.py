"""
whitelist_manager.py
白名单管理器：持久化、校验、增删；区分交互TTY / Agent非交互模式
配置文件：项目根目录 config/excel_agent_config.json
"""
import os
import json
from pathlib import Path
from typing import Tuple, Dict, Any, List

def get_project_root() -> Path:
    """自动探测项目根目录，向上搜索包含tests文件夹，跨平台无硬编码"""
    current = Path(__file__).resolve()
    for _ in range(5):
        if (current / "tests").exists():
            return current
        current = current.parent
    return Path.cwd()

PROJECT_ROOT: Path = get_project_root()
CONFIG_DIR: Path = PROJECT_ROOT / "config"
CONFIG_PATH: Path = CONFIG_DIR / "excel_agent_config.json"

def _ensure_config() -> None:
    """确保配置目录和配置文件存在，不存在则初始化默认配置"""
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    if not CONFIG_PATH.exists():
        # 配置文件首次创建：把项目 tests/test_data 加入默认白名单
        default_test_dir = str((PROJECT_ROOT / "tests" / "test_data").resolve())
        init_config = {
            "allow_dirs": [default_test_dir]
        }
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(init_config, f, ensure_ascii=False, indent=2)

def load_whitelist() -> List[str]:
    """加载白名单，返回绝对路径字符串列表"""
    _ensure_config()
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    raw = cfg.get("allow_dirs", [])
    abs_list = [str(Path(p).resolve()) for p in raw]
    return abs_list

def save_whitelist(allow_dirs: List[str]) -> None:
    """保存白名单，自动去重、转为绝对路径"""
    _ensure_config()
    unique_set = {str(Path(d).resolve()) for d in allow_dirs}
    final_list = sorted(list(unique_set))
    cfg = {"allow_dirs": final_list}
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)

def is_file_allowed(file_path: str, whitelist: List[str]) -> Tuple[bool, str]:
    abs_file = Path(file_path).resolve()
    for dir_str in whitelist:
        abs_dir = Path(dir_str).resolve()
        # 使用 relative_to 判断是否在目录下
        try:
            abs_file.relative_to(abs_dir)
            return True, f"文件位于信任目录 {dir_str}"
        except ValueError:
            continue
    return False, f"文件不在信任白名单，文件完整路径: {abs_file}"

# ========= TTY终端交互模式（本地命令行，可input） =========
def interactive_check(file_path: str) -> Tuple[bool, List[str]]:
    """
    终端交互校验：
    1. 如果白名单为空，询问用户添加初始目录
    2. 文件不在白名单，询问是否添加该文件父目录
    返回 (是否允许访问, 更新后的白名单列表)
    """
    wl = load_whitelist()
    # 白名单为空，首次使用
    if len(wl) == 0:
        print("\n⚠️ 当前Excel分析信任白名单为空，请输入需要信任的目录路径：")
        user_in = input("信任目录 > ").strip()
        if user_in:
            p = str(Path(user_in).resolve())
            confirm = input(f"确认添加目录【{p}】进入白名单？(y/n):").strip().lower()
            if confirm == "y":
                wl.append(p)
                save_whitelist(wl)
                print(f"✅已添加信任目录：{p}")
        else:
            print("❌未输入目录，拒绝访问")
            return False, wl

    ok, _ = is_file_allowed(file_path, wl)
    if ok:
        return True, wl

    abs_fp = str(Path(file_path).resolve())
    candidate = str(Path(abs_fp).parent.resolve())
    print(f"\n⚠️ 文件 {abs_fp} 不在信任白名单")
    ans = input(f"是否将父目录【{candidate}】添加到信任白名单？(y/n):").strip().lower()
    if ans == "y":
        wl.append(candidate)
        save_whitelist(wl)
        print(f"✅已添加目录：{candidate}")
        return True, wl
    else:
        print("❌用户拒绝添加目录，禁止访问该文件")
        return False, wl

def list_whitelist() -> List[str]:
    """打印并返回白名单列表，CLI--list‑whitelist使用"""
    wl = load_whitelist()
    print("===== Excel‑Agent 信任白名单目录 =====")
    if len(wl) == 0:
        print("(白名单为空)")
        return wl
    for idx, item in enumerate(wl):
        print(f"[{idx}] {item}")
    return wl

def remove_whitelist_by_index(idx: int) -> List[str]:
    """按序号删除白名单条目，--remove‑whitelist‑idx 使用"""
    wl = load_whitelist()
    if idx < 0 or idx >= len(wl):
        raise IndexError(f"序号{idx}超出白名单范围，当前共{len(wl)}条")
    removed = wl.pop(idx)
    save_whitelist(wl)
    print(f"✅已移除白名单条目：{removed}")
    return wl

# ========= 非交互模式（Agent / Skill subprocess调用，无input） =========
def agent_precheck(file_path: str) -> Dict[str, Any]:
    """
    供subprocess/skill调用，无控制台输入；
    不在白名单返回 need_user_confirm=True，交由Agent向聊天用户确认
    """
    wl = load_whitelist()
    abs_fp = str(Path(file_path).resolve())
    ok, msg = is_file_allowed(file_path, wl)
    if ok:
        return {
            "allowed": True,
            "need_user_confirm": False,
            "file_abs": abs_fp,
            "candidate_dir": None,
            "whitelist": wl,
            "message": msg
        }
    candidate_dir = str(Path(abs_fp).parent.resolve())
    return {
        "allowed": False,
        "need_user_confirm": True,
        "file_abs": abs_fp,
        "candidate_dir": candidate_dir,
        "whitelist": wl,
        "message": msg
    }

def agent_apply_confirm(dir_path: str, user_confirm: bool) -> Dict[str, Any]:
    """
    Agent拿到用户聊天确认后调用；user_confirm=True才真正写入配置
    """
    wl = load_whitelist()
    abs_dir = str(Path(dir_path).resolve())
    if user_confirm:
        if abs_dir not in wl:
            wl.append(abs_dir)
            save_whitelist(wl)
        return {"success": True, "whitelist": wl, "msg": f"已将 {abs_dir} 加入信任白名单"}
    else:
        return {"success": False, "whitelist": wl, "msg": "用户拒绝添加目录，禁止访问文件"}