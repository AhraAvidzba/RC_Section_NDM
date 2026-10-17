"""Проверяет полноту часового плана и независимо интегрирует сохранённые НДС.

SciPy/HiGHS используются только в тестах. Невозможность равновесия доказывается
необходимыми границами усилий; неудача внешнего решателя доказательством не служит.
"""
from __future__ import annotations
import argparse
import hashlib
import json
import re
import sys
import struct
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "tmp/focused-dependencies"))


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def read(path):
    return json.loads(path.read_text(encoding="utf-8-sig"))


def fields(line):
    return dict(p.split("=", 1) for p in line.split("|")[1:] if "=" in p)

def number(value):
    if isinstance(value,str) and value.startswith('F64:'):
        return struct.unpack('>d',bytes.fromhex(value[4:]))[0]
    return float(value)


def reports(directory, run):
    output = []
    for item in run["Records"]:
        if item["Kind"] == "VBA":
            path = ROOT / item["Report"]
            if digest(path) != item["ReportSHA256"]:
                raise ValueError("Report changed: " + str(path))
            output.append((item, path.read_text(encoding="utf-8-sig").splitlines()))
    return output


class Model:
    def __init__(self, lines):
        import numpy as np
        fibers = [fields(x) for x in lines if x.startswith("FIBER|")]
        self.center = fields(next(x for x in lines if x.startswith("CENTER|")))
        self.tol = fields(next(x for x in lines if x.startswith("ACCEPTANCE_TOLERANCE|")))
        self.tolerance = np.array([number(self.tol[k]) for k in ("N", "Mx", "My")])
        self.xy = np.array([[number(f["x"]), number(f["y"])] for f in fibers])
        self.area = np.array([number(f["area"]) for f in fibers])
        self.steel = np.array([f["kind"] == "S" for f in fibers])
        self.matrix = np.array([np.ones(len(fibers)), self.xy[:, 1], self.xy[:, 0]]) * self.area
        self.length = max(np.max(np.abs(self.xy)), 1.0)
        self.points, self.flags = {}, {}
        for line in lines:
            if line.startswith(("MATERIAL|", "MATERIAL_FLAGS|")):
                d = fields(line)
                key = d["mode"], d["profile"], d["kind"]
                if line.startswith("MATERIAL|"):
                    self.points.setdefault(key, []).append([number(d["strain"]), number(d["stress"])])
                else:
                    self.flags[key] = d
        self.points = {key: np.array(value) for key, value in self.points.items()}

    def stress(self, strain, mode, role, kind):
        import numpy as np
        points, flags = self.points[mode, role, kind], self.flags[mode, role, kind]
        result = np.interp(strain, points[:, 0], points[:, 1])
        for index, other, flag, mask in ((0, 1, "compressionExtension", strain < points[0, 0]),
                                        (-1, -2, "tensionExtension", strain > points[-1, 0])):
            if flags[flag] == "True":
                slope = (points[index, 1] - points[other, 1]) / (points[index, 0] - points[other, 0])
                result[mask] = points[index, 1] + (strain[mask] - points[index, 0]) * slope
        return result

    def integrate(self, plane, mode, role):
        strain = plane[0] + self.xy[:, 1] * plane[1] + self.xy[:, 0] * plane[2]
        concrete = self.stress(strain, mode, role, "C")
        stress = concrete.copy()
        stress[self.steel] = (self.stress(strain, mode, role, "S") - concrete)[self.steel]
        return self.matrix @ stress

    def independent_failure(self, target, mode, role):
        import numpy as np
        from scipy.optimize import least_squares, linprog
        flags = [self.flags[mode, role, kind] for kind in ("C", "S")]
        scales = np.array([1., self.length, self.length])
        centered_tolerance = self.tolerance.copy()
        centered_tolerance[1:] += abs(np.array([number(self.center['y']), number(self.center['x'])])) * self.tolerance[0]
        # В OFF все напряжения ограничены. Независимые напряжения расширяют
        # множество плоских НДС: несовместность LP доказывает невозможность НДС.
        bounded = all(f['compressionExtension'] == 'False' and f['tensionExtension'] == 'False' for f in flags)
        if bounded:
            cp, sp = (self.points[mode, role, kind] for kind in ("C", "S"))
            knots = np.unique(np.r_[cp[:, 0], sp[:, 0]])
            net = np.interp(knots, sp[:, 0], sp[:, 1]) - np.interp(knots, cp[:, 0], cp[:, 1])
            lower = np.where(self.steel, min(net), min(cp[:, 1]))
            upper = np.where(self.steel, max(net), max(cp[:, 1]))
            matrix = self.matrix / scales[:, None]
            lp = linprog(np.zeros(len(lower)), A_ub=np.r_[matrix, -matrix],
                         b_ub=np.r_[(target + centered_tolerance) / scales,
                                    (-target + centered_tolerance) / scales],
                         bounds=list(zip(lower, upper)), method='highs')
            if lp.status == 2:
                # Явное доказательство: проекция заданных усилий превышает
                # максимум при любых независимо допустимых напряжениях.
                if not hasattr(self,'directions'):
                    columns=matrix.T
                    axes=[np.eye(3)]
                    for index in range(1,len(columns)):
                        normals=np.cross(columns[index],columns[:index])
                        sizes=np.linalg.norm(normals,axis=1)
                        valid=sizes>1e-12*np.linalg.norm(columns[index])*np.linalg.norm(columns[:index],axis=1)
                        if np.any(valid):axes.append(normals[valid]/sizes[valid,None])
                    directions=np.vstack(axes)
                    self.directions=np.r_[directions,-directions]
                projections=self.directions@matrix
                support=np.maximum(projections*lower,projections*upper).sum(axis=1)
                projected=self.directions@(target/scales)
                tolerance=abs(self.directions)@(centered_tolerance/scales)
                roundoff=1e-9*np.maximum(1,np.maximum(abs(projected),abs(support)))
                margin=projected-support-tolerance-roundoff
                index=int(np.argmax(margin))
                if margin[index]>0:
                    return {"Classification": "ProvenImpossibleByNecessaryForceBounds", "HiGHSStatus": lp.status,
                            "Passed": True, "Direction":self.directions[index].tolist(),
                            "ProjectedTargetN":float(projected[index]),"MaximumProjectionN":float(support[index]),
                            "AcceptanceProjectionN":float(tolerance[index]),"RoundoffAllowanceN":float(roundoff[index]),
                            "StrictMarginN":float(margin[index]),
                            "Meaning": "Full requested state; no claim about an entire load path."}
        norm = max(float(np.max(abs(target / scales))), 1.)
        def residual(z):
            return (self.integrate(z / scales, mode, role) - target) / scales / norm
        best = None
        for start in ([0., 0., 0.], [-0.001, 0., 0.], [0.001, .001, -.001], [-1., .1, -.1]):
            candidate = least_squares(residual, start, max_nfev=400, xtol=1e-12, ftol=1e-12, gtol=1e-12)
            error = abs(self.integrate(candidate.x / scales, mode, role) - target)
            metric = float(max(error / centered_tolerance))
            if best is None or metric < best[0]:
                best = metric, candidate.x.tolist(), error.tolist()
            if metric <= 1:
                return {"Classification": "IndependentEquilibriumFoundForNumFail", "Passed": False,
                        "Residual": error.tolist(), "CenteredPlaneScaled": candidate.x.tolist()}
        return {"Classification": "Inconclusive", "Passed": False, "BestResidualRatio": best[0],
                "Meaning": "Optimizer failure is not proof of impossibility."}


def independent(directory, run):
    import numpy as np
    import scipy
    verifier_snapshot=directory/'IndependentVerifier.py'
    verifier_snapshot.write_bytes(Path(__file__).read_bytes())
    all_reports = reports(directory, run)
    models = {item['ID'][6:]: Model(lines) for item, lines in all_reports if item['ID'].startswith('Model_')}
    controls, failures, unsupported, numerical_messages = [], [], [], []
    for item, lines in all_reports:
        if item['Macro'] != 'modTestBatchCalculation.RunFocusedAcceptanceCases' and item['ID'] not in ('Batch15_No', 'Batch15_Yes'):
            continue
        shape = item['Argument1'] or 'SavedWorkbook'
        model = models[shape]
        cx, cy = (number(model.center[k]) for k in ('x', 'y'))
        mode = item['Mode']
        for line in lines:
            if not line.startswith('FOCUSED_STATE|'):
                continue
            state = fields(line)
            role = {'StrengthState':'PR1', 'CapacityState':'PR1', 'CrackedState':'PR2',
                    'PreCrackState':'PR3', 'PostCrackState':'PR2'}.get(state['type'])
            if role is None:
                unsupported.append({'Report':item['ID'], 'State':state}); continue
            n, mx, my = (number(state[k]) for k in ('N','Mx','My'))
            target = np.array([n, mx-cy*n, my-cx*n])
            record = {'Report':item['ID'], 'ID':state['id'], 'Mode':mode, 'Type':state['type'], 'Role':role}
            if state['converged'] == 'True':
                eps, kx, ky = (number(state[k]) for k in ('eps0','kx','ky'))
                efforts = model.integrate(np.array([eps+cy*kx+cx*ky,kx,ky]), mode, role)
                # Возвращаем интеграл в исходные оси и проверяем исходные допуски.
                original = efforts.copy(); original[1:] += [cy*efforts[0],cx*efforts[0]]
                residual = abs(original-np.array([n,mx,my]))
                record.update(Passed=bool(np.all(residual <= model.tolerance)), Residual=residual.tolist())
                controls.append(record)
            elif state['status'] == 'NumFail':
                if state['type'] not in ('StrengthState','CrackedState'):
                    record.update(Passed=False,Classification='LimitPathNeedsSeparateReview')
                else:
                    record.update(model.independent_failure(target, mode, role))
                failures.append(record)
        proved={(r['ID'],r['Type']) for r in failures if r['Report']==item['ID'] and r['Passed']}
        for line in lines:
            if line.startswith('META:') and '|status=rsNumericalFailure|' in line:
                match=re.match(r'META: audit03\.comments\.(.+)\.(direct|current)\|',line)
                state_type='StrengthState' if match and match[2]=='direct' else 'CrackedState'
                prefix='plateau.batch15.' if item['ID'].startswith('Batch15_') else 'focused.'+shape+'.'
                covered=bool(match) and (prefix+match[1],state_type) in proved
                numerical_messages.append({'Report':item['ID'],'Line':line,'Covered':covered})
                if not covered:unsupported.append({'Report':item['ID'],'UnreviewedNumericalMeta':line})
            if line.startswith('FOCUSED_UNRESOLVED|'):
                message=fields(line)
                prefix='focused.'+shape+'.'
                covered=message['kind']=='1' and any(key[0]==prefix+message['id'] for key in proved)
                if not covered:unsupported.append({'Report':item['ID'],'UnreviewedNumericalMessage':line})
    result = {'Passed': bool(controls) and len(models)==5 and not unsupported and all(r['Passed'] for r in controls+failures),
              'SourceSHA256':run['SourceSHA256'],
              'VerifierSHA256':digest(verifier_snapshot),
              'ReportHashes':{item['ID']:item['ReportSHA256'] for item,_ in all_reports},
              'ScipyVersion':scipy.__version__, 'ModelCount':len(models), 'IndependentIntegrals':len(controls),
              'ReviewedNumFailStates':len(failures), 'NumericalMessages':numerical_messages,
              'Unsupported':unsupported, 'Controls':controls, 'Failures':failures}
    (directory/'Independent.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print(json.dumps({**{k:result[k] for k in ('Passed','ScipyVersion','ModelCount','IndependentIntegrals','ReviewedNumFailStates')},
                      'UnsupportedCount':len(unsupported)},ensure_ascii=False))
    return result


def verify(directory, run):
    errors, checks, cases, assertions = [], [], 0, 0
    if any(not item.get('Passed') for item in run['Records']):
        errors.append('A required stage reported an unsuccessful exit')
    preparation=read(directory/'Probe/Preparation.json')
    inputs=read(directory/'Probe/InputManifest.json')
    if (digest(Path(run['SourceWorkbook']))!=run['SourceSHA256'] or
        preparation['SourceSHA256']!=run['SourceSHA256'] or inputs['SourceSHA256']!=run['SourceSHA256'] or
        not preparation['Compiled'] or not inputs['Passed']):
        errors.append('Source, compilation or input provenance differs')
    if digest(directory/'Probe/RC_Section_NDM.xlsm')!=preparation['ProbeWorkbookSHA256']:
        errors.append('Frozen probe workbook changed')
    for file,key in [('FocusedAcceptanceProbe.bas','HelperSHA256'),('Combined.bas','ExtraHelperSHA256'),('SavedSectionFixture.json','SavedFixtureSHA256')]:
        if digest(directory/file)!=preparation[key]:errors.append('Executed helper/fixture snapshot differs: '+file)
    expected = {(shape,group,mode) for shape,group in [('RectSet','Generated'),('Circle','Generated'),
                ('RoundedRectangle','Generated'),('HollowRectangle','Generated'),('SavedWorkbook','Saved'),
                ('SavedWorkbook','Paths')] for mode in ('No','Yes')}
    seen=set()
    for item, lines in reports(directory,run):
        body='\n'.join(lines)
        if 'SOURCE_SHA256: '+preparation['ProbeWorkbookSHA256'] not in lines:
            errors.append(item['ID']+': test workbook provenance differs')
        if not item['Passed'] or 'WATCHDOG_COMPLETED: exit=0;' not in body:
            errors.append(item['ID']+': execution did not pass')
        if 'GLOBAL_MODE: General.DiagramExtension='+item['Mode']+';' not in body or 'MODE_AFTER_SUITE: '+item['Mode'] not in lines:
            errors.append(item['ID']+': requested ON/OFF mode is not confirmed')
        if item['ID'].startswith('Model_'):
            continue
        totals = [x for x in lines if x.startswith('TOTAL')]
        if not totals or not re.search(r'passed=(\d+);\s*failed=0\b',totals[-1]):
            errors.append(item['ID']+': total missing or failed')
        else:
            assertions+=int(re.search(r'passed=(\d+)',totals[-1])[1])
        if item['Macro'].endswith('.RunFocusedAcceptanceCases'):
            key=item['Argument1'],item['Argument2'],item['Mode']
            if key in seen:errors.append('Duplicate load group')
            seen.add(key)
            timings=[fields(x) for x in lines if x.startswith('FOCUSED_TIME|')]
            input_manifest=read(directory/'Probe/InputManifest.json')
            group=next(g for g in input_manifest['Groups'] if (g['Shape'],g['Group'])==key[:2])
            input_path=Path(group['Report'])
            if digest(input_path)!=group['ReportSHA256']:
                errors.append(item['ID']+': input manifest changed')
            selected=[fields(x)['id'] for x in input_path.read_text(encoding='utf-8-sig').splitlines() if x.startswith('FOCUSED_INPUT|')]
            if [x['id'] for x in timings]!=selected:
                errors.append(item['ID']+': exact selected load IDs/order differ')
            expected_count=8 if key[1]=='Generated' else 23 if key[1]=='Saved' else 6
            if len(timings)!=expected_count or len({x['id'] for x in timings})!=expected_count:
                errors.append(item['ID']+': selected case count differs')
            cases+=len(timings)
            if any(float(t['elapsedSec'])>10 for t in timings): errors.append(item['ID']+': time exceeds 10 s')
            compare=[fields(x)['id'] for x in lines if x.startswith('FOCUSED_COMPARE|')]
            if compare!=[x for x in selected if x!='zero']:
                errors.append(item['ID']+': calculated result coverage differs')
        if item['ID'].startswith('Batch15_'):
            compare=[fields(x)['id'] for x in lines if x.startswith('FOCUSED_COMPARE|')]
            if len(compare)!=15 or len(set(compare))!=15:
                errors.append(item['ID']+': batch must contain 15 distinct results')
            batch_time=[fields(x) for x in lines if x.startswith('CORE_BATCH_TIME|')]
            if len(batch_time)!=1 or float(batch_time[0]['seconds'])>10:
                errors.append(item['ID']+': batch exceeds the existing 10 s gate')
        if item['ID'].startswith(('Batch15_','Stability_','AsymmetricSteel_')) or item['Macro'].endswith('.RunFocusedAcceptanceCases'):
            if 'RESULTS_SAVE_REOPEN:' not in body or 'STATUS_STYLE_SAVE_REOPEN:' not in body or 'equal=False' in body:
                errors.append(item['ID']+': save/reopen proof missing')
        checks.append({'ID':item['ID'],'Passed':item['Passed']})
    if seen!=expected or cases!=122:errors.append('122 selected cases in both modes are incomplete')
    ids={r['ID'] for r in run['Records']}
    required={'Batch15_No','Batch15_Yes','Stability_No','Stability_Yes','AsymmetricSteel_No','AsymmetricSteel_Yes','OutputPresentation','NativeCAD','Independent'}
    if not required<=ids:errors.append('Required remaining stages are missing: '+str(sorted(required-ids)))
    for name in ('NativeCAD/Stage.json','Independent.json'):
        path=directory/name
        if not path.exists() or not read(path).get('Passed'):errors.append(name+': stage did not pass')
    native_path=directory/'NativeCAD/Stage.json'
    if native_path.exists():
        native=read(native_path)
        if digest(directory/'NativeCAD/NativeRunner.ps1')!=native.get('RunnerSHA256'):
            errors.append('Native runner snapshot differs')
        if {r['Macro'] for r in native.get('Results',[])}!={'RunRealAutoCADFarOriginTests','RunRealAutoCADFarCapacityTests'}:
            errors.append('Both required native CAD gates must finish')
        for item in native.get('Results',[]):
            folder=directory/'NativeCAD'/item['Macro']
            invocation=read(folder/'NativeInvocation.json')
            if (invocation['SourceSHA256']!=run['SourceSHA256'] or not invocation['SourceUnchanged'] or
                not invocation['Passed'] or digest(folder/'NativeCAD.txt')!=item['ReportSHA256'] or
                invocation.get('RunnerSHA256')!=native.get('RunnerSHA256')):
                errors.append('Native CAD provenance or report differs: '+item['Macro'])
    independent_path=directory/'Independent.json'
    if independent_path.exists():
        independent_result=read(independent_path)
        if independent_result.get('SourceSHA256')!=run['SourceSHA256'] or independent_result.get('ReportHashes')!={item['ID']:item['ReportSHA256'] for item,_ in reports(directory,run)}:
            errors.append('Independent evidence belongs to another source or reports')
        if any(item.get('Macro','').endswith('.DumpFocusedForceBoundsExact') for item in run['Records']):
            snapshot=directory/'IndependentVerifier.py'
            if not snapshot.exists() or digest(snapshot)!=independent_result.get('VerifierSHA256'):
                errors.append('Independent verifier snapshot is missing or differs')
    if run['Status']!='COMPLETED_PENDING_REVIEW':errors.append('Driver is incomplete: '+run['Status'])
    if not run.get('ProtectedInputsUnchanged'):errors.append('Protected input changed')
    if not run.get('FinishedUTC'):
        errors.append('Driver completion timestamp is missing')
    elif datetime.fromisoformat(run['FinishedUTC'])>datetime.fromisoformat(run['DeadlineUTC']):
        errors.append('One-hour budget exceeded')
    for entry in run['ProtectedInputs']:
        if digest(Path(entry['Path']))!=entry['SHA256']:errors.append('Protected input changed: '+entry['Path'])
    return {'Status':'PASS' if not errors else 'INCOMPLETE_OR_FAILED','SourceSHA256':run['SourceSHA256'],
            'IndividualCases':cases,'PassedAssertions':assertions,'ElapsedSeconds':run.get('ElapsedSeconds'),
            'Errors':errors,'Stages':checks}


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--directory',type=Path,required=True)
    parser.add_argument('--independent',action='store_true')
    parser.add_argument('--check',action='store_true')
    args=parser.parse_args()
    run=read(args.directory/'Run.json')
    if args.independent:
        raise SystemExit(0 if independent(args.directory,run)['Passed'] else 1)
    result=verify(args.directory,run)
    (args.directory/'Acceptance.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print(json.dumps(result,ensure_ascii=False))
    if args.check and result['Status']!='PASS':raise SystemExit(1)


if __name__=='__main__':main()
