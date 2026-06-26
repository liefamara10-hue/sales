"""
分销缺口计算脚本 v2
用产品代码精准匹配 + 核心名后备匹配，计算每个门店的分销缺口
输出 data.json 供 index.html 看板使用
"""
import json
import re
import os
import subprocess
from collections import defaultdict
from datetime import date

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ====== 门店类型映射 ======
STORE_TYPE_MAP = {
    '商超A': '商超AB', '商超B': '商超AB',
    '农贸菜场店': '农贸店',
    '生鲜店': '生鲜店', '小超市': '小超市',
    '粮油店': '粮油店', '食杂店': '食杂店', '便利店': '便利店',
    '流通批发': None, '批发流通渠道': None, '其它客户': None,
}


def find_file(*keywords):
    """在脚本目录下找包含所有关键词的 xlsx 文件"""
    for f in os.listdir(BASE_DIR):
        if f.endswith('.xlsx') and all(k in f for k in keywords):
            return os.path.join(BASE_DIR, f)
    return None


def simple_core(name):
    """简化版核心名提取（仅用于代码匹配失败时的后备）"""
    if not name:
        return None
    name = str(name).strip()
    name = name.replace('×', '*')
    name = re.sub(r'([0-9A-Za-z])x([0-9])', r'\1*\2', name)
    name = re.sub(r'\([^)]*\)', '', name)
    name = re.sub(r'（[^）]*）', '', name)
    name = re.sub(r'[()（）]', '', name)
    name = re.sub(r'\*\d+', '', name)
    name = re.sub(r'[/\s]?\d+\.?\d*\s*(?:ML|L|KG|G|克|毫升|升)\s*$', '', name, flags=re.IGNORECASE).strip()
    name = name.rstrip('/*- ').strip().upper()
    return name


# ====== 读取分销标准 ======
def read_distribution_standard(filepath):
    """
    新格式：Row1=产线/产品代码/产品名称/八大业态, Row2=各门店类型, 数据从Row3起
    Col A=产线, Col B=产品代码, Col C=产品名称, Col D-J=7种门店类型
    """
    wb = load_workbook(filepath, data_only=True)
    ws = wb['分销标准']

    # Row2 → 门店类型列映射
    col_map = {}
    for c in range(4, 11):  # D=4 to J=10
        header = ws.cell(row=2, column=c).value
        if header:
            col_map[c] = str(header).strip()

    # 建立合并单元格映射：任意被合并的单元格 → 其所属合并区域
    merged_map = {}  # (row, col) → MergedCellRange
    for mr in ws.merged_cells.ranges:
        for row in range(mr.min_row, mr.max_row + 1):
            for col in range(mr.min_col, mr.max_col + 1):
                merged_map[(row, col)] = mr

    def get_cell_value(row, col):
        """读取单元格值，自动处理合并单元格"""
        cell = ws.cell(row=row, column=col)
        if cell.value is not None and not isinstance(cell, MergedCell):
            return str(cell.value).strip()
        # 被合并覆盖，从合并区域左上角取值
        mr = merged_map.get((row, col))
        if mr:
            top_val = ws.cell(row=mr.min_row, column=mr.min_col).value
            if top_val is not None:
                return str(top_val).strip()
        return None

    current_line = None
    products = []

    for r in range(3, ws.max_row + 1):
        # 产线 (Col A) - 用合并单元格值
        line_val = get_cell_value(r, 1)
        if line_val:
            current_line = line_val

        # 产品代码 (Col B) 和 产品名称 (Col C)
        product_code = get_cell_value(r, 2)
        product_name = get_cell_value(r, 3)

        if not product_name:
            continue

        # 跳过非产品条目
        if any(kw in product_name for kw in ['不做强制要求', '散米桶', '设置建议']):
            continue

        # 读取各门店类型要求
        requirements = {}
        for col_idx, store_type in col_map.items():
            requirements[store_type] = get_cell_value(r, col_idx)

        products.append({
            'code': product_code,
            'name': product_name,
            'line': current_line or '',
            'requirements': requirements,
        })

    # 识别 N选1 分组 + 散米桶分组
    groups = []
    bulk_bins = []  # 散米桶合并显示
    for mr in ws.merged_cells.ranges:
        if mr.min_row < 3:
            continue
        if mr.min_col < 4 or mr.min_col > 10:
            continue

        val = ws.cell(row=mr.min_row, column=mr.min_col).value
        if val is None:
            continue
        val = str(val).strip()

        # 散米桶
        if '散米桶' in val:
            store_type = col_map.get(mr.min_col)
            if not store_type:
                continue
            group_products = []
            for r in range(mr.min_row, mr.max_row + 1):
                pname = get_cell_value(r, 3)
                if pname:
                    group_products.append(pname)
            if group_products:
                bulk_bins.append({
                    'store_type': store_type,
                    'label': val,
                    'products': group_products,
                })
            continue

        # N选1
        n_choose = None
        if '二选一' in val or val == '2选1':
            n_choose = 2
        elif val == '3选1':
            n_choose = 3
        else:
            continue

        store_type = col_map.get(mr.min_col)
        if not store_type:
            continue

        group_products = []
        for r in range(mr.min_row, mr.max_row + 1):
            pname = get_cell_value(r, 3)
            if pname:
                group_products.append(pname)

        if group_products:
            groups.append({
                'store_type': store_type,
                'n': n_choose,
                'products': group_products,
            })

    wb.close()
    return products, groups, bulk_bins


# ====== 读取门店销量 ======
def read_store_sales(filepath):
    wb = load_workbook(filepath, data_only=True)
    ws = wb.active

    headers = {}
    for c in range(1, ws.max_column + 1):
        val = ws.cell(row=1, column=c).value
        if val:
            headers[str(val).strip()] = c

    def find_col(*names):
        for n in names:
            if n in headers:
                return headers[n]
        return None

    col_store = find_col('标准客户名称')
    col_type = find_col('标准客户类型(整合)')
    col_code = find_col('物料编码')
    col_product = find_col('物料名称')
    col_manager = find_col('负责人', '负责人员')

    missing = []
    if not col_store: missing.append('标准客户名称')
    if not col_type: missing.append('标准客户类型(整合)')
    if not col_product: missing.append('物料名称')
    if not col_manager: missing.append('负责人/负责人员')
    if missing:
        raise ValueError(f'未找到列: {", ".join(missing)}，表头: {list(headers.keys())}')

    store_data = {}
    for r in range(2, ws.max_row + 1):
        store_name = ws.cell(row=r, column=col_store).value
        store_type = ws.cell(row=r, column=col_type).value
        product_code = ws.cell(row=r, column=col_code).value if col_code else None
        product_name = ws.cell(row=r, column=col_product).value
        manager = ws.cell(row=r, column=col_manager).value

        if not all([store_name, store_type, product_name]):
            continue

        store_name = str(store_name).strip()
        store_type = str(store_type).strip()
        product_name = str(product_name).strip()
        product_code = str(product_code).strip() if product_code else None
        manager = str(manager).strip() if manager else '未知'

        key = (store_name, store_type, manager)
        if key not in store_data:
            store_data[key] = {'codes': set(), 'names': set()}
        if product_code:
            store_data[key]['codes'].add(product_code)
        store_data[key]['names'].add(product_name)

    wb.close()

    stores = []
    for (name, stype, manager), data in store_data.items():
        stores.append({
            'name': name,
            'type': stype,
            'manager': manager,
            'codes': data['codes'],
            'names': data['names'],
        })
    return stores


# ====== 主流程 ======
def build_data():
    sales_file = find_file('销量')
    std_file = find_file('标准')

    if not sales_file:
        raise FileNotFoundError('未找到销量表（文件名需含"销量"）')
    if not std_file:
        raise FileNotFoundError('未找到标准表（文件名需含"标准"）')

    print(f'销量表: {os.path.basename(sales_file)}')
    print(f'标准表: {os.path.basename(std_file)}')

    print('读取分销标准...')
    std_products, std_groups, bulk_bins = read_distribution_standard(std_file)
    print(f'  产品数: {len(std_products)}, N选1分组: {len(std_groups)}, 散米桶: {len(bulk_bins)}')

    print('读取门店销量...')
    stores = read_store_sales(sales_file)
    print(f'  门店数: {len(stores)}')

    # 仅用产品代码匹配
    std_by_code = {}
    for p in std_products:
        if p['code']:
            c = p['code']
            if c not in std_by_code:
                std_by_code[c] = []
            std_by_code[c].append(p)

    all_sales_codes = set()
    for s in stores:
        all_sales_codes.update(s['codes'])

    code_matched = sum(1 for c in all_sales_codes if c in std_by_code)
    print(f'  产品代码匹配: {code_matched} 种')

    # 汇总到负责人
    manager_stores = defaultdict(list)

    for store in stores:
        stype = store['type']
        mapped_type = STORE_TYPE_MAP.get(stype)

        # 门店已售标准产品（仅代码匹配）
        current_std_names = set()
        for c in store['codes']:
            if c in std_by_code:
                for p in std_by_code[c]:
                    current_std_names.add(p['name'])

        # 计算缺口
        gaps = []
        if mapped_type is not None:
            for sp in std_products:
                req = sp['requirements'].get(mapped_type)
                if req is None or req in ('None', ''):
                    continue
                if req and req not in ('None', '') and '选' not in req:
                    if sp['name'] not in current_std_names:
                        gaps.append({
                            'productLine': sp['line'],
                            'productCode': sp['code'] or '',
                            'product': sp['name'],
                            'requirement': '必须',
                        })

            for g in std_groups:
                if g['store_type'] != mapped_type:
                    continue
                covered = any(p in current_std_names for p in g['products'])
                if not covered:
                    first_line = ''
                    for sp in std_products:
                        if sp['name'] == g['products'][0]:
                            first_line = sp['line']
                            break
                    gaps.append({
                        'productLine': first_line,
                        'productCode': '',
                        'requirement': f'{g["n"]}选1',
                        'group': g['products'],
                    })

        # 散米桶合并：同一散米桶组的缺口合并为一条
        if mapped_type:
            for bb in bulk_bins:
                if bb['store_type'] != mapped_type:
                    continue
                # 找出该组内哪些产品是缺口
                bb_gap_products = []
                remaining_gaps = []
                for g in gaps:
                    if g.get('product') in bb['products']:
                        bb_gap_products.append(g['product'])
                    else:
                        remaining_gaps.append(g)
                if bb_gap_products:
                    gaps = remaining_gaps
                    gaps.append({
                        'productLine': '米',
                        'productCode': '',
                        'requirement': '散米桶',
                        'note': bb['label'],
                        'group': bb_gap_products,
                    })

        gaps.sort(key=lambda g: (g.get('productLine', '') or '', g.get('product', '') or ''))

        manager_stores[store['manager']].append({
            'name': store['name'],
            'type': stype,
            'hasStandard': mapped_type is not None,
            'gaps': gaps,
        })

    # 构建输出
    managers = []
    for name, mstores in sorted(manager_stores.items()):
        gap_store_count = sum(1 for s in mstores if len(s['gaps']) > 0)
        managers.append({
            'name': name,
            'storeCount': len(mstores),
            'gapStoreCount': gap_store_count,
            'stores': sorted(mstores, key=lambda s: (-len(s['gaps']), s['name'])),
        })

    managers.sort(key=lambda m: (-m['gapStoreCount'], m['name']))

    data = {
        'generated': date.today().isoformat(),
        'sourceSales': os.path.basename(sales_file),
        'sourceStd': os.path.basename(std_file),
        'managers': managers,
    }

    output_file = os.path.join(BASE_DIR, 'data.json')
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f'\n输出: {output_file}')
    print(f'负责人: {len(managers)}, 有缺口门店: {sum(m["gapStoreCount"] for m in managers)}')
    print('完成！')

    auto_git_push()


def auto_git_push():
    import subprocess
    result = subprocess.run(
        ['git', '-C', BASE_DIR, 'status', '--porcelain', 'data.json', 'index.html'],
        capture_output=True, text=True
    )
    if not result.stdout.strip():
        print('data.json 和 index.html 均无变更，跳过提交')
        return

    print('\n提交到 GitHub...')
    cmds = [
        ['git', '-C', BASE_DIR, 'add', 'data.json', 'index.html'],
        ['git', '-C', BASE_DIR, 'commit', '-m', f'更新数据 {date.today().isoformat()}'],
        ['git', '-C', BASE_DIR, 'push'],
    ]
    for cmd in cmds:
        r = subprocess.run(cmd, capture_output=True, text=True)
        if r.returncode != 0:
            print(f'  失败: {" ".join(cmd)}\n  {r.stderr.strip()}')
            return
    print('  已推送到 GitHub')


if __name__ == '__main__':
    build_data()
    try:
        input("按 Enter 键退出...")
    except EOFError:
        pass
