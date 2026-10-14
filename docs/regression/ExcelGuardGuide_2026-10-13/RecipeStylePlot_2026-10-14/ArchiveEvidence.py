"""Компактные доказательства без частных макрокниг и неудачных копий."""
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

directory = Path(__file__).resolve().parent


def archive(name, paths):
    target = directory/name
    if target.exists():
        raise RuntimeError('Existing evidence archive: '+str(target))
    with ZipFile(target, 'w', ZIP_DEFLATED) as out:
        for path in sorted(set(paths)):
            if path.suffix.lower() == '.xlsm':
                raise RuntimeError('Private workbook must not be archived: '+str(path))
            out.write(path, str(path.relative_to(directory)))
    print(target.name, target.stat().st_size)


archive('NativeVBAEvidence.zip', [directory/name for name in ('BaselineNativeVBA.txt', 'FinalNativeVBA.txt', 'PublishedNativeVBA.txt')])
paths = []
for name in ('BaselineObjects', 'CandidateObjects'):
    paths += [p for p in (directory/name).iterdir() if p.suffix.lower() in ('.txt', '.json', '.xlsx', '.png')]
paths += [directory/name for name in ('OpenArc.json', 'OpenArc.xlsx', 'OpenArc.png')]
archive('NativePlotEvidence.zip', paths)
paths = []
for name in ('CalculatedBaselineG', 'CalculatedGSection', 'CalculatedHollowRounded'):
    paths += [p for p in (directory/name).iterdir() if p.suffix.lower() in ('.txt', '.tsv', '.json', '.png')]
archive('CalculatedExamplesEvidence.zip', paths)
archive('DirectedTestEvidence.zip', [p for p in (directory/'CandidateDirected').iterdir() if p.suffix.lower() in ('.txt', '.json')])
