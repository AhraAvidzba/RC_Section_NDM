"""Compare native Excel PNGs and bind visual review to the inspected images."""
import argparse
import hashlib
import importlib.util
import json
from pathlib import Path

from PIL import Image, ImageChops, ImageDraw, ImageOps

spec = importlib.util.spec_from_file_location('full_logs', Path(__file__).with_name('Summarize-GeometryFull.py'))
logs = importlib.util.module_from_spec(spec)
spec.loader.exec_module(logs)

FAMILIES = ('Circle', 'Rectangle', 'LSection', 'TwoRectangles', 'RoundedSimple',
            'RoundedTapered', 'RoundedMixed', 'HollowCentered', 'HollowOffset')
SPECIAL = ('PERF_01_Stress', 'PERF_01_Strain', 'PERF_02_Stress', 'PERF_02_Strain',
           'Worst_Stress', 'Worst_Strain', 'Resize_NoLabels', 'MissingState')


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--directory', type=Path, required=True)
    args = parser.parse_args()
    root = args.directory
    qa = root/'VisualQA'
    qa.mkdir(exist_ok=True)
    errors, images, pairs, sheets, reports = [], [], [], [], []

    def opened(relative):
        path = root/relative
        if not path.is_file():
            errors.append('Missing PNG: ' + relative)
            return None
        with Image.open(path) as native:
            native.load()
            img = native.convert('RGB')
        if img.width < 400 or img.height < 300 or ImageChops.difference(img, Image.new('RGB', img.size, 'white')).getbbox() is None:
            errors.append('Empty or undersized PNG: ' + relative)
        images.append({'path': relative, 'sha256': logs.sha(path), 'size': list(img.size)})
        return img

    def sheet(name, tiles):
        width, height = 620, 480
        canvas = Image.new('RGB', (2*width, ((len(tiles)+1)//2)*height), '#d8d8d8')
        draw = ImageDraw.Draw(canvas)
        for index, (label, img) in enumerate(tiles):
            x, y = (index % 2)*width, (index//2)*height
            draw.text((x+8, y+6), label, fill='black')
            view = ImageOps.contain(img, (width-10, height-30))
            canvas.paste(view, (x+(width-view.width)//2, y+25))
        path = qa/(name+'.png')
        canvas.save(path)
        sheets.append({'path': path.relative_to(root).as_posix(), 'sha256': logs.sha(path)})

    for mode in ('Off', 'On'):
        report = logs.inspect(root/f'FinalPlotModes_{mode}.txt', logs.sha(root/'AuxiliaryModeFinal.xlsm'), 1)
        reports.append(report)
        if report['status'] != 'PASS':
            errors.append('Incomplete plot-mode test: ' + mode)
        mode_report = root/f'FinalPlotModes_{mode}.txt'
        mode_text = mode_report.read_text(encoding='utf-8-sig') if mode_report.is_file() else ''
        expected_mode = 'No' if mode == 'Off' else 'Yes'
        if f'PLOT_DIAGRAM_MODE: {expected_mode}' not in mode_text:
            errors.append('Actual plot diagram mode is not proven: ' + mode)
        for quantity in ('Stress', 'Strain'):
            tiles = []
            for index, family in enumerate(FAMILIES):
                before = opened(f'FinalPlots_{mode}_Baseline/{family}_{quantity}.png')
                after = opened(f'FinalPlots_{mode}_Current/{family}_{quantity}.png')
                if before is not None and after is not None:
                    same_size = before.size == after.size
                    delta = ImageChops.difference(before, after) if same_size else None
                    bounds = delta.getbbox() if delta else None
                    pairs.append({'mode': mode, 'family': family, 'quantity': quantity,
                                  'sameSize': same_size, 'pixelsEqual': same_size and bounds is None,
                                  'differenceBounds': list(bounds) if bounds else None})
                    tiles.extend([(family+' / baseline', before), (family+' / current', after)])
                if (index+1) % 3 == 0 and tiles:
                    sheet(f'{mode}_{quantity}_{(index+1)//3}', tiles)
                    tiles = []
        special = []
        for name in SPECIAL:
            img = opened(f'FinalPlotModes_{mode}/{name}.png')
            if img is not None:
                special.append((name, img))
        if special:
            sheet(mode+'_special', special)

    if len(images) != 88 or len(pairs) != 36 or len(sheets) != 14:
        errors.append('Expected 88 images, 36 pairs and 14 contact sheets')
    review_path = qa/'Review.json'
    review = json.loads(review_path.read_text(encoding='utf-8-sig')) if review_path.is_file() else {}
    reviewed = review.get('status') == 'REVIEWED' and review.get('images') == images and review.get('contactSheets') == sheets
    if not reviewed:
        errors.append('Visual review is missing or bound to different image hashes')
    differences = [p for p in pairs if not p['pixelsEqual']]
    if differences and not review.get('differencesAcceptedWithReason'):
        errors.append('Pixel differences require a recorded visual explanation')
    result = {'status': 'PASS' if not errors else 'INCOMPLETE_OR_FAILED', 'modes': ['Off', 'On'],
              'images': images, 'pairs': pairs, 'contactSheets': sheets, 'reports': reports,
              'visualReview': review, 'errors': errors,
              'scope': 'Native Stress/Strain images of nine families and eight saved-state/resize/missing-state modes per ON/OFF'}
    (root/'FinalVisualEvidence.json').write_text(json.dumps(result, ensure_ascii=False, indent=2)+'\n', encoding='utf-8')
    print(json.dumps({'status': result['status'], 'images': len(images), 'pairs': len(pairs),
                      'pixelDifferences': len(differences), 'contactSheets': len(sheets), 'errors': errors}))
    return 0 if not errors else 1


if __name__ == '__main__':
    raise SystemExit(main())
