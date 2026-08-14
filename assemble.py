# -*- coding: utf-8 -*-
"""
単一HTMLツール一式を組み立てる。入れ子構造は3階層。

1. filler_template.html + core_logic.js + grid_render.js + filler_app.js
   → 「入力フォームの雛形」テキスト（__STRUCTURE__・__FORM_TITLE__は未確定のまま残す）
   → このテキストをJS文字列としてbuilder_app.jsに埋め込む
2. viewer_template.html + core_logic.js + grid_render.js + viewer_app.js
   → 「住民公開ページの雛形」テキスト（__STRUCTURE__・__RECORDS__・__PUBLIC_CONFIG__・
     __VIEWER_TITLE__は未確定のまま残す）
   → このテキストをJS文字列としてaggregator_app.jsに埋め込む
3. aggregator_template.html + core_logic.js + grid_render.js
   + （2のテンプレートを埋め込み済みの）aggregator_app.js
   → 「集約ツールの雛形」テキスト（__STRUCTURE__は未確定のまま残す）
   → このテキストをJS文字列としてbuilder_app.jsに埋め込む
4. builder_template.html + vendor/xlsx.core.min.js + core_logic.js + grid_render.js
   + （1・3のテンプレートを埋め込み済みの）builder_app.js
   → excel_form_builder.html（ツール1・配布不要、開発者/管轄部署が使う）

ツール2（入力フォーム）・ツール3（集約ツール）・ツール4（住民公開ページ）は、いずれも
上流のツールの「書き出す」ボタンが実行時に生成する。ビルド時にはファイルとして出力しない。
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


def build_viewer_template_text():
    template = read('viewer_template.html')
    core_logic = read('core_logic.js')
    grid_render = read('grid_render.js')
    viewer_app = read('viewer_app.js')

    out = template.replace('/* __CORE_LOGIC__ */', core_logic)
    out = out.replace('/* __GRID_RENDER__ */', grid_render)
    out = out.replace('/* __VIEWER_APP__ */', viewer_app)
    return out


def build_aggregator_template_text(viewer_template_text):
    template = read('aggregator_template.html')
    core_logic = read('core_logic.js')
    grid_render = read('grid_render.js')
    aggregator_app = read('aggregator_app.js')

    viewer_json_literal = escape_script_close(json.dumps(viewer_template_text))
    aggregator_app = aggregator_app.replace('"__VIEWER_TEMPLATE_JSON__"', viewer_json_literal)

    out = template.replace('/* __CORE_LOGIC__ */', core_logic)
    out = out.replace('/* __GRID_RENDER__ */', grid_render)
    out = out.replace('/* __AGGREGATOR_APP__ */', aggregator_app)
    return out


def build_builder_html(filler_template_text, aggregator_template_text):
    template = read('builder_template.html')
    sheetjs = read('vendor/xlsx.core.min.js')
    core_logic = read('core_logic.js')
    grid_render = read('grid_render.js')
    builder_app = read('builder_app.js')

    filler_json_literal = escape_script_close(json.dumps(filler_template_text))
    builder_app = builder_app.replace('"__FILLER_TEMPLATE_JSON__"', filler_json_literal)
    aggregator_json_literal = escape_script_close(json.dumps(aggregator_template_text))
    builder_app = builder_app.replace('"__AGGREGATOR_TEMPLATE_JSON__"', aggregator_json_literal)

    out = template.replace('/* __SHEETJS__ */', sheetjs)
    out = out.replace('/* __CORE_LOGIC__ */', core_logic)
    out = out.replace('/* __GRID_RENDER__ */', grid_render)
    out = out.replace('/* __BUILDER_APP__ */', builder_app)
    return out


def main():
    filler_template_text = build_filler_template_text()
    viewer_template_text = build_viewer_template_text()
    aggregator_template_text = build_aggregator_template_text(viewer_template_text)
    builder_html = build_builder_html(filler_template_text, aggregator_template_text)

    out_path = os.path.join(BASE, 'excel_form_builder.html')
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(builder_html)
    print('書き出し完了:', out_path, f'({os.path.getsize(out_path)/1024:.0f} KB)')

    # 参考用：各テンプレート単体の状態も保存しておく（ビルド成果物ではなく確認用）
    debug_path = os.path.join(BASE, '_filler_template_debug.html')
    with open(debug_path, 'w', encoding='utf-8') as f:
        f.write(filler_template_text)
    print('（参考）フィラーテンプレート:', debug_path, f'({os.path.getsize(debug_path)/1024:.0f} KB)')

    viewer_debug_path = os.path.join(BASE, '_viewer_template_debug.html')
    with open(viewer_debug_path, 'w', encoding='utf-8') as f:
        f.write(viewer_template_text)
    print('（参考）ビューアーテンプレート:', viewer_debug_path, f'({os.path.getsize(viewer_debug_path)/1024:.0f} KB)')

    aggregator_debug_path = os.path.join(BASE, '_aggregator_template_debug.html')
    with open(aggregator_debug_path, 'w', encoding='utf-8') as f:
        f.write(aggregator_template_text)
    print('（参考）集約ツールテンプレート:', aggregator_debug_path, f'({os.path.getsize(aggregator_debug_path)/1024:.0f} KB)')


if __name__ == '__main__':
    main()
