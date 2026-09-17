"""check_deps.py 查询pandas、openpyxl当前环境版本，用于生成requirements.txt"""
import importlib.metadata

def get_package_version(pkg_name: str):
    try:
        ver = importlib.metadata.version(pkg_name)
        return ver
    except importlib.metadata.PackageNotFoundError:
        return None

if __name__ == "__main__":
    pd_ver = get_package_version("pandas")
    openpyxl_ver = get_package_version("openpyxl")
    print(f"pandas版本: {pd_ver}")
    print(f"openpyxl版本: {openpyxl_ver}")

    lines = []
    if pd_ver:
        lines.append(f"pandas=={pd_ver}")
    else:
        lines.append("# pandas 当前环境未安装，推荐版本：pandas==2.2.3")
    if openpyxl_ver:
        lines.append(f"openpyxl=={openpyxl_ver}")
    else:
        lines.append("# openpyxl 当前环境未安装，推荐版本：openpyxl==3.1.5")

    with open("requirements.txt","w",encoding="utf-8") as f:
        f.write("\n".join(lines))

    print("\n✅已自动生成 requirements.txt，请查看文件内容")