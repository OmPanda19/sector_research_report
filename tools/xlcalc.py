#!/usr/bin/env python3
"""Workbook-wide Excel formula evaluator.

Recursive-descent parser with lazy IF / IFERROR / AND / OR, cross-sheet references,
defined names, ranges, array broadcasting and the function set this workbook uses.
Used to prove the model actually computes before it is shipped.
"""
import re
import math
from openpyxl.utils import column_index_from_string as ci, get_column_letter as gl


class Err(Exception):
    pass


def nf(v):
    if isinstance(v, bool):
        return 1.0 if v else 0.0
    if v is None:
        return 0.0
    if isinstance(v, str):
        if v.strip() == '':
            return 0.0
        raise Err('text in arithmetic: %r' % v[:40])
    return float(v)


def bc(a, b, f):
    la, lb = isinstance(a, list), isinstance(b, list)
    if not la and not lb:
        return f(a, b)
    n = max(len(a) if la else 1, len(b) if lb else 1)
    return [f(a[i] if la else a, b[i] if lb else b) for i in range(n)]


TOK = re.compile(r"""
    (?P<num>\d+\.?\d*(?:[eE][+-]?\d+)?)
  | (?P<str>"(?:[^"]|"")*")
  | (?P<sqref>'(?:[^']|'')+'!\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+)
  | (?P<sqcell>'(?:[^']|'')+'!\$?[A-Z]{1,3}\$?\d+)
  | (?P<shref>[A-Za-z0-9_.]+!\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+)
  | (?P<shcell>[A-Za-z0-9_.]+!\$?[A-Z]{1,3}\$?\d+)
  | (?P<rng>\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+)
  | (?P<cell>\$?[A-Z]{1,3}\$?\d+)
  | (?P<func>_xlfn\.[A-Za-z][A-Za-z0-9_.]*|[A-Za-z][A-Za-z0-9_.]*)
  | (?P<op><>|<=|>=|[-+*/^()<>=,&%:])
  | (?P<ws>\s+)
""", re.X)


def lex(s):
    out, i = [], 0
    while i < len(s):
        m = TOK.match(s, i)
        if not m:
            raise Err('lex error at %d in %r' % (i, s[:80]))
        i = m.end()
        if m.lastgroup == 'ws':
            continue
        out.append((m.lastgroup, m.group()))
    return out


class P:
    def __init__(self, toks, ev, sheet):
        self.t, self.i, self.ev, self.sheet = toks, 0, ev, sheet

    def peek(self):
        return self.t[self.i] if self.i < len(self.t) else (None, None)

    def eat(self, val=None):
        k, v = self.peek()
        if val and v != val:
            raise Err('expected %r got %r' % (val, v))
        self.i += 1
        return v

    def expr(self):
        left = self.concat()
        while self.peek()[1] in ('=', '<>', '<', '>', '<=', '>='):
            op = self.eat()
            right = self.concat()
            a, b = self.ev.val(left, self.sheet), self.ev.val(right, self.sheet)

            def cmp(x, y, _op=op):
                if isinstance(x, str) or isinstance(y, str):
                    x = '' if x is None else x
                    y = '' if y is None else y
                    x, y = str(x).upper(), str(y).upper()
                else:
                    x, y = nf(x), nf(y)
                return {'=': x == y, '<>': x != y, '<': x < y, '>': x > y,
                        '<=': x <= y, '>=': x >= y}[_op]
            left = bc(a, b, cmp)
        return left

    def concat(self):
        v = self.add()
        while self.peek()[1] == '&':
            self.eat()
            a = self.ev.val(v, self.sheet)
            b = self.ev.val(self.add(), self.sheet)
            v = _s(a) + _s(b)
        return v

    def add(self):
        v = self.mul()
        while self.peek()[1] in ('+', '-'):
            op = self.eat()
            a, b = self.ev.val(v, self.sheet), self.ev.val(self.mul(), self.sheet)
            v = bc(a, b, (lambda x, y: nf(x) + nf(y)) if op == '+' else (lambda x, y: nf(x) - nf(y)))
        return v

    def mul(self):
        v = self.pw()
        while self.peek()[1] in ('*', '/'):
            op = self.eat()
            a, b = self.ev.val(v, self.sheet), self.ev.val(self.pw(), self.sheet)
            if op == '*':
                v = bc(a, b, lambda x, y: nf(x) * nf(y))
            else:
                def dv(x, y):
                    if nf(y) == 0:
                        raise Err('div0')
                    return nf(x) / nf(y)
                v = bc(a, b, dv)
        return v

    def pw(self):
        v = self.pct()
        while self.peek()[1] == '^':
            self.eat()
            a, b = self.ev.val(v, self.sheet), self.ev.val(self.pct(), self.sheet)
            v = bc(a, b, lambda x, y: nf(x) ** nf(y))
        return v

    def pct(self):
        v = self.un()
        while self.peek()[1] == '%':
            self.eat()
            v = bc(self.ev.val(v, self.sheet), 0, lambda x, _y: nf(x) / 100.0)
        return v

    def un(self):
        if self.peek()[1] in ('-', '+'):
            op = self.eat()
            v = self.ev.val(self.un(), self.sheet)
            return bc(v, 0, (lambda x, _y: -nf(x)) if op == '-' else (lambda x, _y: nf(x)))
        return self.prim()

    def prim(self):
        k, v = self.peek()
        if v == '(':
            self.eat()
            e = self.expr()
            self.eat(')')
            return e
        if k == 'num':
            self.eat()
            return float(v)
        if k == 'str':
            self.eat()
            return v[1:-1].replace('""', '"')
        if k in ('sqref', 'shref'):
            self.eat()
            sh, ref = v.rsplit('!', 1)
            return ('RANGE', sh.strip("'").replace("''", "'"), ref.replace('$', ''))
        if k in ('sqcell', 'shcell'):
            self.eat()
            sh, ref = v.rsplit('!', 1)
            return ('CELL', sh.strip("'").replace("''", "'"), ref.replace('$', ''))
        if k == 'rng':
            self.eat()
            return ('RANGE', self.sheet, v.replace('$', ''))
        if k == 'cell':
            self.eat()
            return ('CELL', self.sheet, v.replace('$', ''))
        if k == 'func':
            name = self.eat()
            if self.peek()[1] != '(':
                # a defined name
                return self.ev.defined(name, self.sheet)
            self.eat('(')
            args, depth, start = [], 0, self.i
            while True:
                kk, vv = self.peek()
                if vv is None:
                    raise Err('unterminated call %s' % name)
                if vv == '(':
                    depth += 1
                elif vv == ')':
                    if depth == 0:
                        args.append(self.t[start:self.i])
                        self.eat()
                        break
                    depth -= 1
                elif vv == ',' and depth == 0:
                    args.append(self.t[start:self.i])
                    self.eat()
                    start = self.i
                    continue
                self.i += 1
            return self.ev.call(name.replace('_xlfn.', '').upper(), args, self.sheet)
        raise Err('unexpected %r' % v)


def _s(v):
    if v is None:
        return ''
    if isinstance(v, bool):
        return 'TRUE' if v else 'FALSE'
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v)


class Book:
    def __init__(self, wb, max_depth=200):
        self.wb = wb
        self.cache = {}
        self.errors = {}
        self.names = {}
        self.max_depth = max_depth
        self.depth = 0
        for n, dn in wb.defined_names.items():
            self.names[n.upper()] = dn.value

    # ------------------------------------------------------------ access
    def raw(self, sheet, coord):
        try:
            ws = self.wb[sheet]
        except KeyError:
            raise Err('no sheet %r' % sheet)
        c = ws[coord]
        v = c.value
        if v is not None and not isinstance(v, (str, int, float, bool)):
            v = getattr(v, 'text', None)
        return v

    def cell(self, sheet, coord):
        key = (sheet, coord)
        if key in self.cache:
            v = self.cache[key]
            if v == '__CYC__':
                raise Err('CIRCULAR at %s!%s' % (sheet, coord))
            return v
        self.cache[key] = '__CYC__'
        v = self.raw(sheet, coord)
        if isinstance(v, str) and v.startswith('='):
            self.depth += 1
            if self.depth > self.max_depth:
                self.depth -= 1
                self.cache[key] = '__CYC__'
                raise Err('depth exceeded at %s!%s' % (sheet, coord))
            try:
                v = self.run(v[1:], sheet)
            except Err as e:
                self.cache[key] = None
                self.errors['%s!%s' % (sheet, coord)] = str(e)
                self.depth -= 1
                raise
            finally:
                self.depth -= 1
        elif v is None:
            v = None
        self.cache[key] = v
        return v

    def get(self, sheet, coord, default=None):
        """Evaluate without raising."""
        try:
            return self.cell(sheet, coord)
        except (Err, ZeroDivisionError, ValueError, TypeError, KeyError, RecursionError) as e:
            self.errors['%s!%s' % (sheet, coord)] = str(e)
            return default

    def run(self, body, sheet):
        return self.val(P(lex(body), self, sheet).expr(), sheet)

    def val(self, v, sheet):
        if isinstance(v, tuple):
            if v[0] == 'CELL':
                return self.cell(v[1], v[2])
            if v[0] == 'RANGE':
                return self.rng(v[1], v[2])
        return v

    def rng(self, sheet, ref):
        a, b = ref.split(':')
        c1, r1 = re.match(r'([A-Z]+)(\d+)', a).groups()
        c2, r2 = re.match(r'([A-Z]+)(\d+)', b).groups()
        out = []
        for rr in range(int(r1), int(r2) + 1):
            for cc in range(ci(c1), ci(c2) + 1):
                out.append(self.cell(sheet, '%s%d' % (gl(cc), rr)))
        return out

    def defined(self, name, sheet):
        v = self.names.get(name.upper())
        if v is None:
            raise Err('unknown name %s' % name)
        if str(v).startswith('#'):
            raise Err('name %s is %s' % (name, v))
        return self.run(str(v), sheet)

    # ------------------------------------------------------------ functions
    def call(self, name, args, sheet):
        def sub(i):
            return P(args[i], self, sheet).expr()

        def E(i):
            return self.val(sub(i), sheet)

        def N(i):
            v = E(i)
            if isinstance(v, list):
                v = v[0] if v else 0.0
            return nf(v)

        def flat(i):
            v = E(i)
            return v if isinstance(v, list) else [v]

        def nums(i):
            return [x for x in flat(i) if isinstance(x, (int, float)) and not isinstance(x, bool)]

        na = len(args)
        if name == 'IF':
            c = E(0)
            if isinstance(c, list):
                c = c[0] if c else False
            c = (c != '') if isinstance(c, str) else bool(c)
            if c:
                return E(1)
            return E(2) if na > 2 else False
        if name == 'IFS':
            for i in range(0, na - 1, 2):
                c = E(i)
                if (c != '') if isinstance(c, str) else bool(c):
                    return E(i + 1)
            raise Err('IFS no match')
        if name == 'IFERROR' or name == 'IFNA':
            try:
                v = E(0)
                if isinstance(v, str) and v.startswith('#'):
                    return E(1)
                return v
            except (Err, ZeroDivisionError, ValueError, TypeError, KeyError):
                return E(1)
        if name == 'AND':
            for i in range(na):
                for v in flat(i):
                    if not ((v != '') if isinstance(v, str) else bool(v)):
                        return False
            return True
        if name == 'OR':
            for i in range(na):
                for v in flat(i):
                    if ((v != '') if isinstance(v, str) else bool(v)):
                        return True
            return False
        if name == 'NOT':
            return not bool(E(0))
        if name == 'CHOOSE':
            return E(int(round(N(0))))
        if name == 'INDEX':
            arr = flat(0)
            k = int(round(N(1)))
            if na > 2:
                k = max(k, int(round(N(2))))
            if k < 1 or k > len(arr):
                raise Err('INDEX %d out of range %d' % (k, len(arr)))
            return arr[k - 1]
        if name in ('MATCH', 'XMATCH'):
            look = E(0)
            arr = flat(1)
            for j, x in enumerate(arr):
                if isinstance(look, str) or isinstance(x, str):
                    if _s(x).upper() == _s(look).upper():
                        return float(j + 1)
                else:
                    try:
                        if nf(x) == nf(look):
                            return float(j + 1)
                    except Err:
                        pass
            raise Err('MATCH not found: %r' % (look,))
        if name == 'XLOOKUP':
            look = E(0)
            keys = flat(1)
            vals = flat(2)
            for j, k in enumerate(keys):
                hit = (_s(k).upper() == _s(look).upper()) if (isinstance(look, str) or isinstance(k, str)) \
                    else (nf(k) == nf(look))
                if hit:
                    return vals[j] if j < len(vals) else None
            if na > 3:
                return E(3)
            raise Err('XLOOKUP not found: %r' % (look,))
        if name in ('MIN', 'MAX', 'SUM', 'AVERAGE', 'MEDIAN', 'COUNT', 'COUNTA', 'PRODUCT', 'STDEV'):
            v = []
            raw = []
            for i in range(na):
                v += nums(i)
                raw += flat(i)
            if name == 'MIN':
                return min(v) if v else 0.0
            if name == 'MAX':
                return max(v) if v else 0.0
            if name == 'SUM':
                return float(sum(v))
            if name == 'PRODUCT':
                p = 1.0
                for x in v:
                    p *= x
                return p
            if name == 'AVERAGE':
                if not v:
                    raise Err('AVERAGE empty')
                return sum(v) / len(v)
            if name == 'MEDIAN':
                if not v:
                    raise Err('MEDIAN empty')
                s = sorted(v)
                n = len(s)
                return s[n // 2] if n % 2 else (s[n // 2 - 1] + s[n // 2]) / 2.0
            if name == 'COUNT':
                return float(len(v))
            if name == 'COUNTA':
                return float(len([x for x in raw if x is not None and x != '']))
            if name == 'STDEV':
                if len(v) < 2:
                    raise Err('STDEV n<2')
                m = sum(v) / len(v)
                return math.sqrt(sum((x - m) ** 2 for x in v) / (len(v) - 1))
        if name == 'SUMPRODUCT':
            arrs = [flat(i) for i in range(na)]
            n = max(len(a) for a in arrs)
            t = 0.0
            for j in range(n):
                p = 1.0
                for a in arrs:
                    x = a[j] if len(a) > 1 else a[0]
                    p *= 0.0 if isinstance(x, str) or x is None else nf(x)
                t += p
            return t
        if name == 'ABS':
            return bc(E(0), 0, lambda x, _y: abs(nf(x)))
        if name == 'SIGN':
            return bc(E(0), 0, lambda x, _y: (nf(x) > 0) - (nf(x) < 0))
        if name in ('ROUND', 'ROUNDUP', 'ROUNDDOWN'):
            x, d = N(0), int(round(N(1)))
            f = 10.0 ** d
            if name == 'ROUND':
                return math.floor(abs(x) * f + 0.5) / f * (1 if x >= 0 else -1)
            if name == 'ROUNDUP':
                return math.ceil(abs(x) * f) / f * (1 if x >= 0 else -1)
            return math.floor(abs(x) * f) / f * (1 if x >= 0 else -1)
        if name == 'INT':
            return float(math.floor(N(0)))
        if name == 'MOD':
            return math.fmod(N(0), N(1))
        if name == 'SQRT':
            return math.sqrt(N(0))
        if name == 'LN':
            return math.log(N(0))
        if name == 'EXP':
            return math.exp(N(0))
        if name == 'POWER':
            return N(0) ** N(1)
        if name == 'ISNUMBER':
            try:
                v = E(0)
            except (Err, ZeroDivisionError, ValueError, TypeError):
                return False
            if isinstance(v, list):
                v = v[0] if v else None
            return isinstance(v, (int, float)) and not isinstance(v, bool)
        if name == 'ISTEXT':
            v = E(0)
            return isinstance(v, str)
        if name == 'ISBLANK':
            return E(0) is None
        if name == 'ISERROR':
            try:
                E(0)
                return False
            except (Err, ZeroDivisionError, ValueError, TypeError):
                return True
        if name == 'N':
            return nf(E(0))
        if name == 'TEXT':
            v = E(0)
            fmt = E(1)
            try:
                x = nf(v)
            except Err:
                return _s(v)
            f = str(fmt)
            if '%' in f:
                dec = len(f.split('.')[1].replace('%', '')) if '.' in f else 0
                return ('{:,.%df}%%' % dec).format(x * 100)
            dec = 0
            if '.' in f:
                dec = len(re.sub(r'[^0#]', '', f.split('.')[1]))
            return ('{:,.%df}' % dec).format(x)
        if name in ('COUNTIF', 'COUNTIFS', 'SUMIF', 'SUMIFS', 'AVERAGEIF', 'AVERAGEIFS'):
            return self._ifs(name, args, sheet)
        if name in ('PERCENTILE', 'PERCENTILE.INC', 'QUARTILE'):
            v = sorted(nums(0))
            if not v:
                raise Err('PERCENTILE empty')
            k = N(1)
            if name == 'QUARTILE':
                k = k / 4.0
            pos = k * (len(v) - 1)
            lo = int(math.floor(pos))
            hi = min(lo + 1, len(v) - 1)
            return v[lo] + (pos - lo) * (v[hi] - v[lo])
        if name in ('SMALL', 'LARGE'):
            v = sorted(nums(0), reverse=(name == 'LARGE'))
            k = int(round(N(1)))
            if k < 1 or k > len(v):
                raise Err('%s out of range' % name)
            return v[k - 1]
        if name in ('RANK', 'RANK.EQ'):
            x = N(0)
            v = nums(1)
            desc = True
            if na > 2 and N(2) != 0:
                desc = False
            s = sorted(v, reverse=desc)
            return float(s.index(x) + 1) if x in s else float(len(s) + 1)
        if name == 'TODAY':
            return 46200.0
        if name == 'NPV':
            r = N(0)
            v = []
            for i in range(1, na):
                v += nums(i)
            return sum(x / (1 + r) ** (j + 1) for j, x in enumerate(v))
        if name == 'YEAR':
            return 2026.0
        if name in ('SEARCH', 'FIND'):
            needle = _s(E(0))
            hay = _s(E(1))
            if name == 'SEARCH':
                i = hay.upper().find(needle.upper())
            else:
                i = hay.find(needle)
            if i < 0:
                raise Err('%s not found' % name)
            return float(i + 1)
        if name == 'LEFT':
            return _s(E(0))[:int(round(N(1))) if na > 1 else 1]
        if name == 'RIGHT':
            k = int(round(N(1))) if na > 1 else 1
            return _s(E(0))[-k:] if k else ''
        if name == 'MID':
            s = _s(E(0))
            a = int(round(N(1)))
            return s[a - 1:a - 1 + int(round(N(2)))]
        if name == 'LEN':
            return float(len(_s(E(0))))
        if name == 'UPPER':
            return _s(E(0)).upper()
        if name == 'LOWER':
            return _s(E(0)).lower()
        if name == 'TRIM':
            return _s(E(0)).strip()
        if name == 'CONCATENATE':
            return ''.join(_s(E(i)) for i in range(na))
        if name == 'TEXTJOIN':
            sep = _s(E(0))
            parts = []
            for i in range(2, na):
                parts += [_s(x) for x in flat(i) if x not in (None, '')]
            return sep.join(parts)
        if name == 'LET':
            # LET(name, value, ..., calc) - evaluate sequentially with substitution
            raise Err('LET not supported')
        raise Err('unsupported function %s' % name)

    def _ifs(self, name, args, sheet):
        def toks(i):
            return args[i]

        def flat(i):
            v = self.val(P(args[i], self, sheet).expr(), sheet)
            return v if isinstance(v, list) else [v]

        if name in ('COUNTIF', 'SUMIF', 'AVERAGEIF'):
            rng = flat(0)
            crit = self.val(P(args[1], self, sheet).expr(), sheet)
            pairs = [(rng, crit)]
            target = flat(2) if len(args) > 2 else rng
        else:
            if name == 'COUNTIFS':
                target = None
                pairs = [(flat(i), self.val(P(args[i + 1], self, sheet).expr(), sheet))
                         for i in range(0, len(args) - 1, 2)]
                rng = pairs[0][0]
            else:
                target = flat(0)
                pairs = [(flat(i), self.val(P(args[i + 1], self, sheet).expr(), sheet))
                         for i in range(1, len(args) - 1, 2)]
                rng = target
        n = len(pairs[0][0])

        def match(x, crit):
            if isinstance(crit, str):
                m = re.match(r'^(<=|>=|<>|<|>|=)(.*)$', crit)
                if m:
                    op, rest = m.groups()
                    try:
                        a, b = nf(x), float(rest)
                    except (Err, ValueError):
                        a, b = _s(x).upper(), rest.upper()
                    return {'<': a < b, '>': a > b, '<=': a <= b, '>=': a >= b,
                            '=': a == b, '<>': a != b}[op]
                if '*' in crit or '?' in crit:
                    pat = '^' + re.escape(crit).replace(r'\*', '.*').replace(r'\?', '.') + '$'
                    return re.match(pat, _s(x), re.I) is not None
                return _s(x).upper() == crit.upper()
            try:
                return nf(x) == nf(crit)
            except Err:
                return False

        hits = []
        for j in range(n):
            if all(match(p[0][j] if j < len(p[0]) else None, p[1]) for p in pairs):
                hits.append(j)
        if name in ('COUNTIF', 'COUNTIFS'):
            return float(len(hits))
        vals = [target[j] for j in hits if j < len(target)
                and isinstance(target[j], (int, float)) and not isinstance(target[j], bool)]
        if name in ('SUMIF', 'SUMIFS'):
            return float(sum(vals))
        if not vals:
            raise Err('AVERAGEIFS empty')
        return sum(vals) / len(vals)
