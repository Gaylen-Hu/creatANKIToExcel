# Anki Deck Creator from Excel

本项目使用 `genanki` 和 `openpyxl` 库从 Excel 文件创建 Anki 卡片组。通过自定义的 HTML/CSS 模板，可以在 Anki 中实现带有选择题选项和答案解析的卡片。

## 功能概述

- 解析 Excel 文件，生成 Anki 卡片组
- 支持单选、多选题模板
- 自定义 HTML、CSS，实现带样式的题目、选项及答案显示
- 使用 `sessionStorage` 进行持久化数据存储

## 安装要求

确保您的环境已安装以下库：

```bash
pip install genanki openpyxl

```

## 使用方法

1. **准备 Excel 文件**
    - Excel 文件应包含题目、答案、解析、分值、题目类型、题号和选项等列。
2. **运行程序**
    - 将程序文件和 Excel 文件放在同一目录下，修改 `excel_file_path`、`deck_name` 和 `output_file_name` 参数。
    - 运行代码以生成 Anki 卡片包（.apkg 文件），可以直接导入到 Anki。
3. **自定义选项**
    - 可根据需求自定义 HTML 模板，以更改卡片样式或交互。

## 文件说明

- `create_anki_deck.py`: 主程序文件，包含从 Excel 生成 Anki 卡片的代码。
- `anki_card_template`: 自定义的 HTML 和 CSS 模板，定义了卡片的显示样式。

## 示例代码

以下是如何使用该代码生成卡片包的示例：

```python
python
复制代码
from create_anki_deck import create_anki_deck

excel_file_path = 'path/to/your/excel_file.xlsx'
deck_name = 'My Anki Deck'
output_file_name = 'my_deck.apkg'

create_anki_deck(excel_file_path, deck_name, output_file_name)

```

## 注意事项

- 本项目使用了 `sessionStorage` 进行答案选择的持久化。
- HTML 模板包含交互式 JavaScript 脚本，在 Anki 中显示卡片时可实现选择题答案的自动对比。

## 截图

可以根据实际情况添加程序运行效果的截图。

## 许可证

MIT License