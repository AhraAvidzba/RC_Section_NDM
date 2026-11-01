"""Строит чёткие формулы справки из математической записи, сверенной с PDF в norms/."""
from pathlib import Path
import argparse
import hashlib
import json

import matplotlib
matplotlib.use('Agg')
from matplotlib import pyplot as plt
from matplotlib.font_manager import FontProperties
from matplotlib.mathtext import MathTextParser
from matplotlib.textpath import TextPath
from PIL import Image


def render_catalog(catalog_path: Path, output: Path, update_catalog: bool):
    """Рисует формулы заново при 4 пикселях на пункт; нормативные сканы не масштабирует."""
    catalog = json.loads(catalog_path.read_text(encoding='utf-8-sig'))
    output.mkdir(parents=True, exist_ok=True)
    plt.rcParams.update({'mathtext.fontset': 'stix', 'font.family': 'DejaVu Sans', 'svg.fonttype': 'path'})
    parser = MathTextParser('path')
    for item in catalog:
        expression = '$'+item['MathText']+'$'
        measured = parser.parse(expression, dpi=72, prop=FontProperties(size=16))
        reference = item['Reference']
        reference_width = TextPath((0, 0), reference, size=8, prop=FontProperties(family='DejaVu Sans')).get_extents().width
        width = float(measured.width)+reference_width+12+24
        height = max(float(measured.height), 10)+16
        figure = plt.figure(figsize=(width/72, height/72), dpi=288, facecolor='#fafafc')
        baseline = (height-float(measured.height))/2+float(measured.depth)
        figure.text(12/width, baseline/height, expression, fontsize=16, va='baseline', color='#1e1e1e')
        figure.text((24+float(measured.width))/width, 0.5, reference, fontsize=8, va='center', color='#1e1e1e')
        path = output/item['File']
        figure.savefig(path, dpi=288, facecolor=figure.get_facecolor())
        figure.savefig(path.with_suffix('.svg'), facecolor=figure.get_facecolor())
        # SVG path сохраняет разделение строками; конечные пробелы Matplotlib не нужны.
        svg = path.with_suffix('.svg')
        svg.write_text('\n'.join(line.rstrip() for line in svg.read_text(encoding='utf-8').splitlines())+'\n', encoding='utf-8')
        plt.close(figure)
        with Image.open(path) as picture:
            picture.convert('RGB').save(path, optimize=True)
        item.update(Renderer='MathText/STIX; freshly typeset from verified normative expression',
                    WidthPoints=width, HeightPoints=height, PixelsPerPoint=4,
                    ExpressionSHA256=hashlib.sha256((item['MathText']+'\n'+item['Reference']).encode('utf-8')).hexdigest().upper(),
                    SHA256=hashlib.sha256(path.read_bytes()).hexdigest().upper())
    if update_catalog:
        catalog_path.write_text(json.dumps(catalog, ensure_ascii=False, indent=2)+'\n', encoding='utf-8')
    return catalog


if __name__ == '__main__':
    cli = argparse.ArgumentParser(description=__doc__)
    cli.add_argument('--catalog', type=Path, default=Path(__file__).with_name('NormativeFormulaImages.json'))
    cli.add_argument('--output-directory', type=Path, default=Path(__file__).with_name('NormativeFormulaImages'))
    cli.add_argument('--update-catalog', action='store_true')
    args = cli.parse_args()
    values = render_catalog(args.catalog, args.output_directory, args.update_catalog)
    print('Rendered normative formulas:', len(values))
