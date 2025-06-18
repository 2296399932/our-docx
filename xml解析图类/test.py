import re
from google import genai
from google.genai import types
import json

client = genai.Client(api_key="AIzaSyBqmGlZHCAFViUE0Jw-ox_Hi171dCK-XQw")

def generate_er_json(client, model, sql_prompt):
    """
    让AI根据SQL语句，整理出表、字段、关联、坐标，返回JSON
    """
    er_prompt = (
        "According to the following SQL CREATE TABLE statements, organize the ER information with these requirements: "
        "1. Output each table's name (preferably use the comment), and each field's name (preferably use the comment). "
        "2. Output the foreign key relationships between tables, and determine whether it is 1-to-N or N-to-N. "
        "3. Each table should have x and y (center, absolute coordinates). "
        "   For each table, please reserve a 400x400 area centered at the table's x and y for the table and all its fields. "
        "   All fields of a table must be placed within this 400x400 area, distributed around the table (top, bottom, left, right, corners, etc.), and must not overlap. "
        "   The x and y of each field must be within the 400x400 area centered at the table's x and y. "
        "   The 400x400 areas of different tables (centered at each table's x and y) must not overlap, and the distance between the centers of any two tables must be at least 500 units. "
        "   When generating the x and y coordinates for tables, please try to place tables that are related (have a foreign key relationship) as close to each other as possible, but still keep at least 500 units between their centers. "
        "4. The x and y of a foreign key relationship (diamond) should be the midpoint between the two related tables. "
        "5. All fields' x and y must be unique (no duplicates). "
        "6. Output only JSON in the following format:\n"
        "{\n  'tables': [\n    {\n      'name': 'table name', 'x': 0, 'y': 0, 'fields': [\n        {'name': 'field name', 'x': 0, 'y': 0}\n      ]\n    }\n  ],\n  'relations': [\n    {\n      'from': 'TableA', 'to': 'TableB', 'type': '1-to-N', 'x': 0, 'y': 0, 'desc': 'foreign key' }\n  ]\n}\n"
        "Do not output any explanation or code block markers, only JSON.\n"
        f"SQL:\n{sql_prompt}"
    )
    print("请求AI生成ER结构JSON...")
    response = client.models.generate_content(
        model=model,
        contents=[er_prompt],
        config=types.GenerateContentConfig(
            max_output_tokens=8192,
            temperature=0.3
        )
    )
    print("AI响应：", response.text)
    # 只提取JSON
    json_match = re.search(r'\{[\s\S]*\}', response.text)
    if json_match:
        return json.loads(json_match.group(0).replace("'", '"'))
    # fallback: 直接尝试解析
    try:
        return json.loads(response.text.replace("'", '"'))
    except Exception:
        print("AI输出无法解析为JSON：", response.text)
        return None

# 示例SQL
sql_prompt = """‘
create table user
(
    id         int auto_increment
        primary key,
    username   varchar(50)  not null,
    password   varchar(100) not null,
    phone      varchar(15)  null,
    role       varchar(20)  not null,
    birth_date date         not null,
    sex        int          not null,
    name       varchar(200) not null
);

create table file
(
    id         int auto_increment
        primary key,
    name       varchar(255)                        not null,
    file_url   varchar(255)                        not null,
    file_path  varchar(255)                        not null,
    created_at timestamp default CURRENT_TIMESTAMP null,
    updated_at timestamp default CURRENT_TIMESTAMP null on update CURRENT_TIMESTAMP,
    user_id    int                                 null,
    constraint file_ibfk_1
        foreign key (user_id) references user (id)
);

create index user_id
    on file (user_id);

create table fileing
(
    id      int auto_increment
        primary key,
    user_id int                                not null,
    file_id int                                not null,
    ing_at  datetime default CURRENT_TIMESTAMP null,
    constraint user_id
        unique (user_id, file_id),
    constraint fileing_ibfk_1
        foreign key (user_id) references user (id)
            on delete cascade,
    constraint fileing_ibfk_2
        foreign key (file_id) references file (id)
            on delete cascade
);

create index file_id
    on fileing (file_id);

create table project
(
    id                  int auto_increment
        primary key,
    name                varchar(100)                        not null,
    folder_path         varchar(255)                        not null,
    directory_image_url varchar(255)                        null,
    created_at          timestamp default CURRENT_TIMESTAMP null,
    updated_at          timestamp default CURRENT_TIMESTAMP null on update CURRENT_TIMESTAMP,
    user_id             int                                 null,
    constraint project_ibfk_1
        foreign key (user_id) references user (id)
);

create table file_project_association
(
    id         int auto_increment
        primary key,
    file_id    int null,
    project_id int null,
    constraint file_project_association_ibfk_1
        foreign key (file_id) references file (id)
            on delete cascade,
    constraint file_project_association_ibfk_2
        foreign key (project_id) references project (id)
            on delete cascade
);

create index file_id
    on file_project_association (file_id);

create index project_id
    on file_project_association (project_id);

create index user_id
    on project (user_id);

create table project_user_association
(
    id         int auto_increment
        primary key,
    project_id int null,
    user_id    int null,
    constraint project_user_association_ibfk_1
        foreign key (project_id) references project (id)
            on delete cascade,
    constraint project_user_association_ibfk_2
        foreign key (user_id) references user (id)
            on delete cascade
);

create index project_id
    on project_user_association (project_id);

create index user_id
    on project_user_association (user_id);

"""

if __name__ == "__main__":
    er_json = generate_er_json(client, "gemini-2.0-flash", sql_prompt)
    print(json.dumps(er_json, ensure_ascii=False, indent=2))
    # 保存到er.json
    with open("er.json", "w", encoding="utf-8") as f:
        json.dump(er_json, f, ensure_ascii=False, indent=2)
