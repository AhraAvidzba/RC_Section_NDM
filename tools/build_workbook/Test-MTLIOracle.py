"""ДЛЯ ТЕСТОВ: независимые интегралы площади для нативного MTLI R/K.

Многоугольники обрезаются по уровням деформаций и интегрируются по
треугольникам (формула степени 2). Круг интегрируется по ширине полосы
адаптивным Симпсоном. Обход Грина и код VBA здесь не используются.
"""
import argparse
import json
import math
from pathlib import Path

EPS = [-0.004, -0.002, -0.001, 0, 0.00005, 0.0001, 0.004]
SIG = [-40, -30, -20, 0, 2, 0, 0]


def law(e):
    """Кусочно-линейная диаграмма с постоянными продолжениями."""
    if e < EPS[0]:
        return SIG[0], 0
    for i in range(1, len(EPS)):
        if e <= EPS[i]:
            et = (SIG[i] - SIG[i - 1]) / (EPS[i] - EPS[i - 1])
            return SIG[i - 1] + et * (e - EPS[i - 1]), et
    return SIG[-1], 0


def clip(poly, plane, level, sign):
    """Оставляет одну полуплоскость; вторая обрезка выделяет полосу."""
    if not poly:
        return []
    e, kx, ky = plane
    out = []
    p = poly[-1]
    a = sign * (e + kx * p[1] + ky * p[0] - level)
    for q in poly:
        b = sign * (e + kx * q[1] + ky * q[0] - level)
        if (a >= 0) != (b >= 0):
            t = a / (a - b)
            out.append((p[0] + t * (q[0] - p[0]), p[1] + t * (q[1] - p[1])))
        if b >= 0:
            out.append(q)
        p, a = q, b
    return out


def polygon_integral(poly, plane):
    """Три точки внутри каждого треугольника точно дают R и K полосы."""
    out = [0.0] * 12
    e, kx, ky = plane
    if kx == ky == 0:
        bands = [(-math.inf, math.inf)]
    else:
        levels = [-math.inf] + EPS + [math.inf]
        bands = list(zip(levels, levels[1:]))
    for low, high in bands:
        part = poly
        if math.isfinite(low):
            part = clip(part, plane, low, 1)
        if math.isfinite(high):
            part = clip(part, plane, high, -1)
        for j in range(1, len(part) - 1):
            vertices = [part[0], part[j], part[j + 1]]
            a, b, c = vertices
            area = abs((b[0] - a[0]) * (c[1] - a[1]) - (b[1] - a[1]) * (c[0] - a[0])) / 2
            for heavy in range(3):
                x = sum(v[0] * (2 / 3 if i == heavy else 1 / 6) for i, v in enumerate(vertices))
                y = sum(v[1] * (2 / 3 if i == heavy else 1 / 6) for i, v in enumerate(vertices))
                s, et = law(e + kx * y + ky * x)
                bases = [1, y, x]
                values = [s * z for z in bases] + [et * v * w for v in bases for w in bases]
                for i, value in enumerate(values):
                    out[i] += area * value / 3
    return out


def simpson(f, a, b, tol=1e-11):
    """Адаптивное интегрирование гладкой скалярной функции между узлами."""
    def step(a, b, fa, fm, fb, whole, tolerance, depth):
        m = (a + b) / 2
        fl, fr = f((a + m) / 2), f((m + b) / 2)
        left = (m - a) * (fa + 4 * fl + fm) / 6
        right = (b - m) * (fm + 4 * fr + fb) / 6
        correction = left + right - whole
        if depth == 0 or abs(correction) <= 15 * tolerance:
            return left + right + correction / 15
        return step(a, m, fa, fl, fm, left, tolerance / 2, depth - 1) + step(m, b, fm, fr, fb, right, tolerance / 2, depth - 1)
    m = (a + b) / 2
    fa, fm, fb = f(a), f(m), f(b)
    return step(a, b, fa, fm, fb, (b - a) * (fa + 4 * fm + fb) / 6, tol, 24)


def circle_integral(radius, plane):
    """Полосы круга: ширина 2 sqrt(R²-v²), интеграл u² = ширина³/12."""
    e, kx, ky = plane
    g = math.hypot(kx, ky)
    nx, ny = (ky / g, kx / g) if g else (0, 1)
    cuts = [-math.pi / 2, math.pi / 2]
    if g:
        cuts += [math.asin((z - e) / (g * radius)) for z in EPS if -radius < (z - e) / g < radius]
    cuts.sort()
    # Dimensionless integration keeps the adaptive error estimate well scaled.
    result = [0.0] * 12
    for a, b in zip(cuts, cuts[1:]):
        # Tangent is constant inside this material band; endpoints do not
        # choose a neighbouring derivative at a discontinuity.
        _, et = law(e + g * radius * math.sin((a + b) / 2))
        def values(t):
            v = radius * math.sin(t)
            halfwidth = radius * math.cos(t)
            width = 2 * halfwidth
            s, _ = law(e + g * v)
            u2 = width**3 / 12
            jac = halfwidth
            first = [width, ny * v * width, nx * v * width]
            second = [width, ny * v * width, nx * v * width,
                      ny * v * width, nx**2 * u2 + ny**2 * v**2 * width, nx * ny * (v**2 * width - u2),
                      nx * v * width, nx * ny * (v**2 * width - u2), ny**2 * u2 + nx**2 * v**2 * width]
            return [s * q * jac for q in first] + [et * q * jac for q in second]
        scales = [40 * radius**2, 40 * radius**3, 40 * radius**3] + [40000 * radius**p for p in [2, 3, 3, 3, 4, 4, 3, 4, 4]]
        for i, scale in enumerate(scales):
            result[i] += scale * simpson(lambda t: values(t)[i] / scale, a, b)
    return result


def run(native):
    """Сравнивает фиксированные экспортированные плоскости, без решения НДС."""
    checks, seen_cases = [], []
    rect = [(-100, -200), (100, -200), (100, 200), (-100, 200)]
    hole = [(-50, -100), (50, -100), (50, 100), (-50, 100)]
    for line in native.read_text(encoding='utf-8-sig').splitlines():
        if 'DATA_MTLI_ORACLE|' not in line:
            continue
        line = line[line.index('DATA_MTLI_ORACLE|'):]  # Итог предыдущего вложенного набора может не заканчиваться переводом строки.
        fields = line.split('|')
        if len(fields) != 18:
            raise ValueError('Oracle record must contain exactly three strains and twelve R/K values')
        kind, case = map(int, fields[1:3])
        seen_cases.append((kind, case))
        plane = tuple(float(s.replace(',', '.')) for s in fields[3:6])
        actual = [float(s.replace(',', '.')) for s in fields[6:]]
        expected = polygon_integral(rect, plane) if kind <= 2 else circle_integral(100, plane)
        if kind in (2, 4):
            inner = polygon_integral(hole, plane) if kind == 2 else circle_integral(50, plane)
            expected = [a - b for a, b in zip(expected, inner)]
        # Fixed dimensional absolute floors only protect cancellation zeros.
        floors = [1e-5, 1e-3, 1e-3] + [1e-2, 1, 1, 1, 100, 100, 1, 100, 100]
        for i, (a, b) in enumerate(zip(actual, expected)):
            error = abs(a - b)
            limit = floors[i] + abs(b) * 1e-7
            checks.append(dict(kind=kind, case=case, component=i, actual=a, expected=b, error=error, limit=limit, passed=error <= limit))
    complete = len(seen_cases) == 16 and set(seen_cases) == {(kind, case) for kind in range(1, 5) for case in range(1, 5)}
    return dict(Status='PASS' if complete and len(checks) == 192 and all(v['passed'] for v in checks) else 'FAIL', Checks=len(checks), CompleteCases=complete, Cases=seen_cases, Results=checks)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('native', type=Path)
    parser.add_argument('output', type=Path)
    args = parser.parse_args()
    result = run(args.native)
    args.output.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    print(result['Status'], result['Checks'], 'checks')
    raise SystemExit(result['Status'] != 'PASS')
