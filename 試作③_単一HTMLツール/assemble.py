# -*- coding: utf-8 -*-
"""
2つの単一HTMLツールを組み立てる。

1. filler_template.html + core_logic.js + grid_render.js + filler_app.js
   → 「入力フォームの雛形」テキスト（__STRUCTURE__・__FORM_TITLE__は未確定のまま残す）
   → このテキストをJS文字列としてbuilder_app.jsに埋め込む
2. builder_template.html + vendor/xlsx.core.min.js + core_logic.js + grid_render.js
   + （1のテンプレートを埋め込み済みの）builder_app.js
   → excel_form_builder.html（ツール1・配布不要、開発者/管轄部署が使う）

ツール2（実際に配布される入力フォーム）は、ツール1の「HTMLで書き出す」ボタンが
実行時に生成する。ビルド時にはファイルとして出力しない。
"""
import json
import os
import re

BASE = os.path.dirname(os.path.abspath(__file__))


def read(name):
    with open(os.path.join(BASE, name), encoding='utf-8') as f:
        return f.read()


def escape_script_close(js_string_literal):
    # HTML的に</script>とみなされる文字列がJS文字列リテラルの中に生でいると
    # ブラウザがスクリプトタグをそこで終端してしまうため、分割してエスケープする。
    return re.sub(r'</(script)', r'<\\/\1', js_string_literal, flags=re.IGNORECASE)


def build_filler_template_text():
    template = read('filler_template.html')
    core_logic = read('core_logic.js')
    grid_render = read('grid_render.js')
    filler_app = read('filler_app.js')

    out = template.replace('/* __CORE_LOGIC__ */', core_logic)
    out = out.replace('/* __GRID_RENDER__ */', grid_render)
    out = out.replace('/* __FILLER_APP__ */', filler_app)
    return out


def build_builder_html(filler_template_text):
    template = read('builder_template.html')
    sheetjs = read('vendor/xlsx.core.min.js')
    core_logic = read('core_logic.js')
    grid_render = read('grid_render.js')
    builder_app = read('builder_app.js')

    filler_json_literal = escape_script_close(json.dumps(filler_template_text))
    builder_app = builder_app.replace('"__FILLER_TEMPLATE_JSON__"', filler_json_literal)

    out = template.replace('/* __SHEETJS__ */', sheetjs)
    out = out.replace('/* __CORE_LOGIC__ */', core_logic)
    out = out.replace('/* __GRID_RENDER__ */', grid_render)
    out = out.replace('/* __BUILDER_APP__ */', builder_app)
    return out


def main():
    filler_template_text = build_filler_template_text()
    builder_html = build_builder_html(filler_template_text)

    out_path = os.path.join(BASE, 'excel_form_builder.html')
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(builder_html)
    print('書き出し完了:', out_path, f'({os.path.getsize(out_path)/1024:.0f} KB)')

    # 参考用：フィラー単体のテンプレート状態も保存しておく（ビルド成果物ではなく確認用）
    debug_path = os.path.join(BASE, '_filler_template_debug.html')
    with open(debug_path, 'w', encoding='utf-8') as f:
        f.write(filler_template_text)
    print('（参考）フィラーテンプレート:', debug_path, f'({os.path.getsize(debug_path)/1024:.0f} KB)')


if __name__ == '__main__':
    main()
