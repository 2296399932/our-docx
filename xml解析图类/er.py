import json
from xml.etree.ElementTree import Element, SubElement, tostring, ElementTree
import networkx as nx

# 读取 er.json
with open('er.json', 'r', encoding='utf-8') as f:
    er = json.load(f)

tables = er['tables']
relations = er['relations']

mxGraphModel = Element('mxGraphModel', {
    'dx': '1200', 'dy': '800', 'grid': '1', 'gridSize': '10', 'guides': '1', 'tooltips': '1',
    'connect': '1', 'arrows': '1', 'fold': '1', 'page': '1', 'pageScale': '1', 'pageWidth': '850',
    'pageHeight': '1100', 'math': '0', 'shadow': '0'
})
root = SubElement(mxGraphModel, 'root')

# id 分配器
def id_gen():
    i = 0
    while True:
        yield str(i)
        i += 1
id_iter = id_gen()

def next_id():
    return next(id_iter)

# 预分配 id
cell_ids = {}

# id=0, id=1
cell0 = SubElement(root, 'mxCell', {'id': next_id()})
cell1 = SubElement(root, 'mxCell', {'id': next_id(), 'parent': '0'})

# 表和字段
for table in tables:
    table_id = next_id()
    cell_ids[(table['name'], 'table')] = table_id
    table_cell = SubElement(root, 'mxCell', {
        'id': table_id,
        'value': table['name'],
        'style': 'rounded=0;whiteSpace=wrap;html=1;fontSize=18;',
        'vertex': '1',
        'parent': '1'
    })
    SubElement(table_cell, 'mxGeometry', {
        'x': str(table['x']), 'y': str(table['y']), 'width': '180', 'height': '60', 'as': 'geometry'
    })
    # 字段
    for field in table['fields']:
        field_id = next_id()
        cell_ids[(table['name'], field['name'])] = field_id
        field_cell = SubElement(root, 'mxCell', {
            'id': field_id,
            'value': field['name'],
            'style': 'ellipse;whiteSpace=wrap;html=1;fontSize=16;',
            'vertex': '1',
            'parent': '1'
        })
        SubElement(field_cell, 'mxGeometry', {
            'x': str(field['x']), 'y': str(field['y']), 'width': '110', 'height': '40', 'as': 'geometry'
        })
        # 字段与表连接
        edge_id = next_id()
        SubElement(root, 'mxCell', {
            'id': edge_id,
            'style': 'edgeStyle=none;endArrow=none;',
            'edge': '1',
            'parent': '1',
            'source': table_id,
            'target': field_id
        }).append(Element('mxGeometry', {'relative': '1', 'as': 'geometry'}))

# 关联
for rel in relations:
    rhombus_id = next_id()
    cell_ids[(rel['from'], rel['to'], 'rhombus')] = rhombus_id
    # 菱形直接用"关联"二字
    rhombus_cell = SubElement(root, 'mxCell', {
        'id': rhombus_id,
        'value': '关联',
        'style': 'rhombus;whiteSpace=wrap;html=1;fontSize=16;',
        'vertex': '1',
        'parent': '1'
    })
    SubElement(rhombus_cell, 'mxGeometry', {
        'x': str(rel['x']), 'y': str(rel['y']), 'width': '50', 'height': '50', 'as': 'geometry'
    })
    # 判断N/1显示
    rel_type = rel.get('type', '')
    if 'N' in rel_type or 'n' in rel_type:
        from_label = 'n'
        to_label = '1'
    else:
        from_label = '1'
        to_label = 'n'
    # from表-菱形
    from_id = cell_ids.get((rel['from'], 'table'))
    if from_id is None:
        print(f"关联 from 表未找到: {rel['from']}")
        continue
    edge1_id = next_id()
    SubElement(root, 'mxCell', {
        'id': edge1_id,
        'value': from_label,
        'style': 'edgeStyle=none;endArrow=none;html=1;align=center;verticalAlign=middle;',
        'edge': '1',
        'parent': '1',
        'source': from_id,
        'target': rhombus_id
    }).append(Element('mxGeometry', {'relative': '1', 'as': 'geometry'}))
    # 菱形-to表
    to_id = cell_ids.get((rel['to'], 'table'))
    if to_id is None:
        print(f"关联 to 表未找到: {rel['to']}")
        continue
    edge2_id = next_id()
    SubElement(root, 'mxCell', {
        'id': edge2_id,
        'value': to_label,
        'style': 'edgeStyle=none;endArrow=none;html=1;align=center;verticalAlign=middle;',
        'edge': '1',
        'parent': '1',
        'source': rhombus_id,
        'target': to_id
    }).append(Element('mxGeometry', {'relative': '1', 'as': 'geometry'}))

# 写入 xml
ElementTree(mxGraphModel).write('output.drawio', encoding='utf-8', xml_declaration=False, short_empty_elements=False)

def auto_layout_tables(tables, relations, area=400, min_dist=500):
    G = nx.Graph()
    for t in tables:
        G.add_node(t['name'])
    for rel in relations:
        G.add_edge(rel['from'], rel['to'])
    pos = nx.spring_layout(G, k=min_dist/100, scale=2000)  # k越大节点越分散
    # 写回tables
    for t in tables:
        t['x'] = int(pos[t['name']][0])
        t['y'] = int(pos[t['name']][1])
    return tables
