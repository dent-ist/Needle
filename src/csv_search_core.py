# -*- coding: utf-8 -*-
"""
Needle - core logic (no UI)

- Detects CSV encoding and delimiter
- Recommends an engine (SQLite / DuckDB) from file size and intended use
- Picks the fastest search method for each engine automatically
- Query language:  ([Column]:"text" AND "any text") OR "prefix*"
- Search history stored as JSON in the user's app-data folder
"""
import codecs
import csv
import io
import json
import os
import re
import sqlite3
import threading
import time
from functools import lru_cache

# DuckDB's compiled engine imports these standard modules at run time.
# Importing them here makes PyInstaller bundle them into the exe.
import collections.abc  # noqa: F401
import datetime  # noqa: F401
import decimal  # noqa: F401
import ipaddress  # noqa: F401
import pathlib  # noqa: F401
import typing  # noqa: F401
import uuid  # noqa: F401
import zoneinfo  # noqa: F401

try:
    import duckdb
except Exception:  # without duckdb, only SQLite is available
    duckdb = None

csv.field_size_limit(2**31 - 1)

MB = 1024 ** 2
GB = 1024 ** 3

ENCODINGS = ["utf-8", "utf-8-sig", "cp949", "utf-16", "latin-1"]
DELIMITERS = {",": "Comma (,)", "\t": "Tab", ";": "Semicolon (;)", "|": "Pipe (|)"}

MATCH_TYPES = ["contains", "equals", "startswith", "regex"]


class Cancelled(Exception):
    pass


class QueryError(ValueError):
    pass


def human_size(n):
    n = float(n)
    for unit in ["B", "KB", "MB", "GB", "TB"]:
        if n < 1024 or unit == "TB":
            return f"{int(n)} B" if unit == "B" else f"{n:.1f} {unit}"
        n /= 1024


def qident(name):
    return '"' + str(name).replace('"', '""') + '"'


def qlit(s):
    return "'" + str(s).replace("'", "''") + "'"


def like_escape(s):
    return s.replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_")


def is_cancel_error(err):
    return isinstance(err, Cancelled) or "interrupt" in str(err).lower()


# ---------------------------------------------------------------------------
# CSV inspection
# ---------------------------------------------------------------------------
def detect_encoding(path, sample_bytes=4 * MB):
    with open(path, "rb") as f:
        raw = f.read(sample_bytes)
    if raw.startswith(codecs.BOM_UTF8):
        return "utf-8-sig"
    if raw.startswith(codecs.BOM_UTF16_LE) or raw.startswith(codecs.BOM_UTF16_BE):
        return "utf-16"
    for enc in ("utf-8", "cp949"):
        try:
            raw.decode(enc)
            return enc
        except UnicodeDecodeError as e:
            if e.start >= len(raw) - 3:  # a character cut off at the end of the sample
                return enc
    return "latin-1"


def detect_delimiter(path, encoding):
    with open(path, "r", encoding=encoding, errors="replace", newline="") as f:
        first = f.readline()
    counts = {d: first.count(d) for d in DELIMITERS}
    best = max(counts, key=counts.get)
    return best if counts[best] > 0 else ","


def read_preview(path, encoding, delimiter, n=20):
    rows = []
    with open(path, "r", encoding=encoding, errors="replace", newline="") as f:
        for i, row in enumerate(csv.reader(f, delimiter=delimiter)):
            rows.append(row)
            if i >= n:
                break
    return rows


def make_headers(raw, ncols):
    out, seen = [], set()
    raw = list(raw) + [""] * max(0, ncols - len(raw))
    for i, h in enumerate(raw):
        h = (h or "").replace("\ufeff", "").strip() or f"Column{i + 1}"
        base, k = h, 2
        while h in seen or h == "_row":
            h = f"{base}_{k}"
            k += 1
        seen.add(h)
        out.append(h)
    return out


class CsvInfo:
    def __init__(self, path):
        self.path = os.path.abspath(path)
        self.size = os.path.getsize(path)
        self.encoding = detect_encoding(path)
        self.delimiter = detect_delimiter(path, self.encoding)
        self.refresh()

    def refresh(self):
        self.preview = read_preview(self.path, self.encoding, self.delimiter)
        header = self.preview[0] if self.preview else []
        width = max([len(r) for r in self.preview] or [0])
        self.headers = make_headers(header, width)


def recommend(size, purpose):
    """purpose: 'keyword' (find words in any column) / 'filter' (column conditions, big exports)
    returns (engine, use_fts, reason)"""
    if duckdb is None:
        return "sqlite", purpose == "keyword", "DuckDB is not available, so SQLite will be used."
    if purpose == "keyword":
        if size <= 2 * GB:
            return ("sqlite", True,
                    "SQLite with a full-text (FTS) index. Importing takes longer, but afterwards "
                    "words are found almost instantly in any column. The database will be "
                    "about 3-4x the size of the CSV.")
        return ("duckdb", False,
                "The file is too large for a practical full-text index. DuckDB scans in "
                "parallel on all CPU cores, so keyword searches are still fast.")
    return ("duckdb", False,
            "DuckDB imports quickly, stores data smaller than the CSV, and is very fast "
            "at column conditions and exporting large results.")


def estimate_db_size(size, engine, use_fts):
    if engine == "duckdb":
        return int(size * 0.6)
    return int(size * (1.3 + (2.5 if use_fts else 0)))


def convert_to_utf8(src, encoding, dst, progress, cancel):
    size = max(1, os.path.getsize(src))
    dec = codecs.getincrementaldecoder(encoding)(errors="replace")
    done = 0
    with open(src, "rb") as fi, open(dst, "wb") as fo:
        while True:
            b = fi.read(16 * MB)
            if not b:
                break
            fo.write(dec.decode(b).encode("utf-8"))
            done += len(b)
            progress(done / size, f"Converting to UTF-8… {human_size(done)} / {human_size(size)}")
            if cancel.is_set():
                raise Cancelled()
        fo.write(dec.decode(b"", final=True).encode("utf-8"))


def remove_db_files(path):
    for p in (path, path + "-journal", path + "-wal", path + ".wal", path + "-shm"):
        try:
            if os.path.exists(p):
                os.remove(p)
        except OSError:
            pass


# ---------------------------------------------------------------------------
# Query language
#
#   word  or  "some text"     any column contains the text
#   [Col]:"text"  [Col]."text" that column contains the text
#   [Col]="text"              that column equals the text
#   [Col]~"regex"             that column matches a regular expression
#   "abc*"  "*abc"  "a*c"     starts with / ends with / wildcard pattern
#   AND (or just a space), OR, NOT, ( )      AND binds tighter than OR
# ---------------------------------------------------------------------------
class Term:
    __slots__ = ("col", "op", "text")

    def __init__(self, col, op, text):
        self.col, self.op, self.text = col, op, text

    def __repr__(self):
        return f"Term({self.col!r}, {self.op!r}, {self.text!r})"


def _tokenize(q):
    toks, i, n = [], 0, len(q)
    while i < n:
        ch = q[i]
        if ch.isspace():
            i += 1
            continue
        if ch in "()":
            toks.append((ch, None, i))
            i += 1
            continue
        if q.startswith("&&", i):
            toks.append(("AND", None, i))
            i += 2
            continue
        if q.startswith("||", i):
            toks.append(("OR", None, i))
            i += 2
            continue
        start, col, op = i, None, None
        if ch == "[":
            j = q.find("]", i + 1)
            if j < 0:
                raise QueryError(f"Missing ']' after the column name at position {i + 1}.")
            col = q[i + 1:j].strip()
            if not col:
                raise QueryError(f"Empty column name '[]' at position {i + 1}.")
            i = j + 1
            if i < n and q[i] in ".:=~":
                op = q[i]
                i += 1
        elif ch in "=~":
            op = ch
            i += 1
        quoted = False
        if i < n and q[i] == '"':
            j, buf = i + 1, []
            while True:
                if j >= n:
                    raise QueryError(f"Missing closing quote for the text starting at position {i + 1}.")
                c = q[j]
                if c == "\\" and j + 1 < n:
                    buf.append(q[j + 1])
                    j += 2
                    continue
                if c == '"':
                    if j + 1 < n and q[j + 1] == '"':
                        buf.append('"')
                        j += 2
                        continue
                    break
                buf.append(c)
                j += 1
            val, i, quoted = "".join(buf), j + 1, True
        else:
            j = i
            while j < n and not q[j].isspace() and q[j] not in '()"':
                j += 1
            val, i = q[i:j], j
        if col is None and op is None and not quoted and val.upper() in ("AND", "OR", "NOT"):
            toks.append((val.upper(), None, start))
            continue
        if val == "":
            if col is not None:
                raise QueryError(f'Column [{col}] needs search text, for example [{col}]:"text".')
            raise QueryError(f"Expected search text at position {start + 1}.")
        toks.append(("TERM", (col, op, val), start))
    return toks


class _Parser:
    def __init__(self, toks):
        self.t, self.i = toks, 0

    def peek(self):
        return self.t[self.i][0] if self.i < len(self.t) else None

    def take(self):
        tok = self.t[self.i]
        self.i += 1
        return tok

    def parse(self):
        if not self.t:
            raise QueryError("Type something to search for.")
        node = self.or_()
        if self.i < len(self.t):
            k, _, pos = self.t[self.i]
            raise QueryError(f"Unexpected '{k}' at position {pos + 1}.")
        return node

    def or_(self):
        node = self.and_()
        while self.peek() == "OR":
            self.take()
            node = ("or", node, self.and_())
        return node

    def and_(self):
        node = self.not_()
        while True:
            k = self.peek()
            if k == "AND":
                self.take()
                node = ("and", node, self.not_())
            elif k in ("TERM", "(", "NOT"):
                node = ("and", node, self.not_())  # a space means AND
            else:
                return node

    def not_(self):
        if self.peek() == "NOT":
            self.take()
            return ("not", self.not_())
        return self.primary()

    def primary(self):
        k = self.peek()
        if k == "TERM":
            return ("term", self.take()[1])
        if k == "(":
            pos = self.take()[2]
            node = self.or_()
            if self.peek() != ")":
                raise QueryError(f"Missing ')' for the '(' at position {pos + 1}.")
            self.take()
            return node
        if k is None:
            raise QueryError("The query ends too early. Add a search term after the last operator.")
        raise QueryError(f"Expected a search term but found '{k}' at position {self.t[self.i][2] + 1}.")


def _classify_wildcard(v):
    if "*" not in v:
        return "contains", v
    core = v.strip("*")
    if core == "":
        raise QueryError('A wildcard needs some text, for example "abc*".')
    if "*" not in core:
        if v.startswith("*") and v.endswith("*"):
            return "contains", core
        if v.endswith("*"):
            return "startswith", core
    return "pattern", v


def parse_query(q, headers):
    ast = _Parser(_tokenize(q)).parse()
    lower = {h.lower(): i for i, h in enumerate(headers)}

    def resolve(node):
        if node[0] != "term":
            return (node[0],) + tuple(resolve(c) for c in node[1:])
        col, op, val = node[1]
        ci = None
        if col is not None:
            if col in headers:
                ci = headers.index(col)
            elif col.lower() in lower:
                ci = lower[col.lower()]
            else:
                names = ", ".join(headers[:12]) + (", …" if len(headers) > 12 else "")
                raise QueryError(f"Unknown column [{col}]. Available: {names}")
        if op == "=":
            kind = "equals"
        elif op == "~":
            kind = "regex"
            try:
                re.compile(val)
            except re.error as e:
                raise QueryError(f'Invalid regular expression "{val}": {e}')
        else:
            kind, val = _classify_wildcard(val)
        return ("term", Term(ci, kind, val))

    return resolve(ast)


_VERBS = {"contains": "contains", "equals": "equals", "startswith": "starts with",
          "regex": "matches regex", "pattern": "matches"}


def describe_query(node, headers, parent=None):
    kind = node[0]
    if kind == "term":
        t = node[1]
        who = "any column" if t.col is None else f"[{headers[t.col]}]"
        return f'{who} {_VERBS[t.op]} "{t.text}"'
    if kind == "not":
        return "NOT " + describe_query(node[1], headers, "not")
    a = describe_query(node[1], headers, kind)
    b = describe_query(node[2], headers, kind)
    s = f"{a} {kind.upper()} {b}"
    if (kind == "or" and parent in ("and", "not")) or (kind == "and" and parent in ("or", "not")):
        return f"({s})"
    return s


def count_terms(node):
    if node[0] == "term":
        return 1
    return sum(count_terms(c) for c in node[1:])


# ---------------------------------------------------------------------------
# Search history
# ---------------------------------------------------------------------------
def app_data_dir():
    if os.name == "nt":
        base = os.environ.get("APPDATA") or os.path.expanduser("~")
    else:
        base = os.environ.get("XDG_CONFIG_HOME") or os.path.expanduser("~/.config")
    d = os.path.join(base, "Needle")
    os.makedirs(d, exist_ok=True)
    # carry over history from the earlier "CSV Fast Search" name
    old = os.path.join(base, "CsvFastSearch", "history.json")
    new = os.path.join(d, "history.json")
    if os.path.exists(old) and not os.path.exists(new):
        try:
            import shutil
            shutil.copyfile(old, new)
        except OSError:
            pass
    return d


class History:
    MAX = 1000
    KEY = ("db", "mode", "column", "match", "text", "case")

    def __init__(self, path=None):
        self.path = path or os.path.join(app_data_dir(), "history.json")
        self.items = []
        self.load()

    def load(self):
        try:
            with open(self.path, "r", encoding="utf-8") as f:
                data = json.load(f)
            self.items = [x for x in data if isinstance(x, dict) and "text" in x]
        except (OSError, ValueError):
            self.items = []

    def save(self):
        try:
            tmp = self.path + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(self.items, f, ensure_ascii=False, indent=1)
            os.replace(tmp, self.path)
        except OSError:
            pass

    def add(self, entry):
        entry = dict(entry)
        entry["time"] = time.strftime("%Y-%m-%d %H:%M:%S")
        key = tuple(entry.get(k) for k in self.KEY)
        self.items = [x for x in self.items if tuple(x.get(k) for k in self.KEY) != key]
        self.items.append(entry)
        del self.items[:-self.MAX]
        self.save()

    def remove(self, entries):
        ids = {id(e) for e in entries}
        self.items = [x for x in self.items if id(x) not in ids]
        self.save()

    def clear(self):
        self.items = []
        self.save()


# ---------------------------------------------------------------------------
# Engines (shared)
# ---------------------------------------------------------------------------
class BaseEngine:
    kind = ""
    label = ""
    ROWNUM = "rowid"

    def __init__(self, db_path):
        self.db_path = db_path
        self.con = None
        self.headers = []
        self.meta = {}
        self.row_count = 0

    @property
    def ncols(self):
        return len(self.headers)

    @property
    def cols(self):
        return [f"c{i}" for i in range(self.ncols)]

    # --- metadata / view ---
    def _create_view(self):
        parts = [f"{self.ROWNUM} AS _row"] + [
            f"c{i} AS {qident(h)}" for i, h in enumerate(self.headers)]
        self.con.execute(f"CREATE VIEW data_view AS SELECT {', '.join(parts)} FROM data")

    def _write_meta(self, info, extra):
        meta = {
            "headers": json.dumps(self.headers, ensure_ascii=False),
            "source": info.path,
            "encoding": info.encoding,
            "delimiter": info.delimiter,
            "created": time.strftime("%Y-%m-%d %H:%M:%S"),
            "rows": str(self.row_count),
        }
        meta.update({k: str(v) for k, v in extra.items()})
        self.con.execute("CREATE TABLE _csvtool_meta (k VARCHAR PRIMARY KEY, v VARCHAR)")
        self.con.executemany("INSERT INTO _csvtool_meta VALUES (?, ?)", list(meta.items()))
        self.meta = meta

    def _load_meta(self):
        try:
            rows = self.con.execute("SELECT k, v FROM _csvtool_meta").fetchall()
        except Exception:
            raise ValueError("This database was not created by Needle.")
        self.meta = dict(rows)
        self.headers = json.loads(self.meta["headers"])
        self.row_count = int(self.meta.get("rows", 0) or 0)

    # --- building the WHERE clause ---
    def _col_expr(self, c, op, text, cs):
        raise NotImplementedError

    def _term_sql(self, col, op, text, cs):
        """returns (sql, params, method) where method is 'fts', 'index' or 'scan'"""
        col_ids = list(range(self.ncols)) if col is None else [col]
        exprs, params = [], []
        for i in col_ids:
            e, p = self._col_expr(f"c{i}", op, text, cs)
            exprs.append(e)
            params += p
        where = exprs[0] if len(exprs) == 1 else "(" + " OR ".join(exprs) + ")"
        return where, params, "scan"

    def _compile(self, node, cs, stats):
        kind = node[0]
        if kind == "term":
            t = node[1]
            sql, params, method = self._term_sql(t.col, t.op, t.text, cs)
            stats[method] = stats.get(method, 0) + 1
            return sql, params
        if kind == "not":
            s, p = self._compile(node[1], cs, stats)
            return f"(NOT COALESCE({s}, FALSE))", p
        a, pa = self._compile(node[1], cs, stats)
        b, pb = self._compile(node[2], cs, stats)
        return f"({a} {kind.upper()} {b})", pa + pb

    def _method_text(self, stats, scope):
        raise NotImplementedError

    def build_where(self, spec):
        """returns (source, where_sql, params, description)"""
        text = spec.get("text", "")
        mode = spec["mode"]
        if mode == "sql":
            return ("view", text.strip() or "TRUE", [], "Runs your SQL condition as written.")
        if text.strip() == "" if mode == "query" else text == "":
            return ("table", "TRUE", [], "No search text: showing all rows from the start.")
        cs = spec.get("case", False)
        if mode == "query":
            ast = parse_query(text, self.headers)
            stats = {}
            where, params = self._compile(ast, cs, stats)
            n = count_terms(ast)
            desc = (f"Reads as: {describe_query(ast, self.headers)}   ·   "
                    f"{n} term{'s' if n > 1 else ''}, {self._method_text(stats, 'query')}")
            return ("table", where, params, desc)
        op = spec["match"]
        if op == "regex":
            try:
                re.compile(text)
            except re.error as e:
                raise QueryError(f"Invalid regular expression: {e}")
        col = spec["col"] if mode == "column" else None
        where, params, method = self._term_sql(col, op, text, cs)
        return ("table", where, params, self._simple_text(method, col, op, cs))

    def _simple_text(self, method, col, op, cs):
        return self._method_text({method: 1}, col)

    # --- running queries ---
    def _sql(self, spec, what):
        src, where, params, _ = self.build_where(spec)
        if src == "view":
            table, cols = "data_view", "*"
        else:
            table, cols = "data", f"{self.ROWNUM} AS _row, " + ", ".join(self.cols)
        if what == "count":
            cols = "count(*)"
        return f"SELECT {cols} FROM {table} WHERE {where}", params

    def search(self, spec, limit, offset):
        sql, params = self._sql(spec, "rows")
        sql += f" LIMIT {int(limit)} OFFSET {int(offset)}"
        return self.con.execute(sql, params).fetchall()

    def count(self, spec):
        sql, params = self._sql(spec, "count")
        return self.con.execute(sql, params).fetchone()[0]

    def export(self, spec, out_path, encoding, progress, cancel):
        sql, params = self._sql(spec, "rows")
        cur = self.con.execute(sql, params)
        n = 0
        with open(out_path, "w", encoding=encoding, errors="replace", newline="") as f:
            w = csv.writer(f)
            w.writerow(self.headers)
            while True:
                rows = cur.fetchmany(20000)
                if not rows:
                    break
                w.writerows(r[1:] for r in rows)  # skip the row-number column
                n += len(rows)
                progress(None, f"Saving… {n:,} rows")
                if cancel.is_set():
                    raise Cancelled()
        return n

    def can_index(self):
        return False

    def interrupt(self):
        try:
            if self.con is not None:
                self.con.interrupt()
        except Exception:
            pass

    def close(self):
        try:
            if self.con is not None:
                self.con.close()
        except Exception:
            pass
        self.con = None


# ---------------------------------------------------------------------------
# SQLite
# ---------------------------------------------------------------------------
@lru_cache(maxsize=64)
def _re_compile(pattern):
    return re.compile(pattern)


def _regexp(pattern, value):
    if value is None:
        return 0
    return 1 if _re_compile(pattern).search(value) else 0


class SqliteEngine(BaseEngine):
    kind = "sqlite"
    label = "SQLite"
    ROWNUM = "rowid"

    def __init__(self, db_path):
        super().__init__(db_path)
        self.fts = ""
        self.indexed = set()

    def _connect(self):
        self.con = sqlite3.connect(self.db_path, check_same_thread=False, isolation_level=None)
        self.con.create_function("REGEXP", 2, _regexp, deterministic=True)
        self.con.execute("PRAGMA cache_size=-262144")  # about 256 MB cache

    def open(self):
        self._connect()
        self._load_meta()
        self.fts = self.meta.get("fts", "")
        self._load_indexes()

    def _load_indexes(self):
        rows = self.con.execute(
            "SELECT name FROM sqlite_master WHERE type='index' AND tbl_name='data'").fetchall()
        self.indexed = {int(m.group(1)) for (name,) in rows
                        if (m := re.fullmatch(r"idx_c(\d+)", name or ""))}

    def import_csv(self, info, use_fts, progress, cancel):
        remove_db_files(self.db_path)
        self._connect()
        c = self.con
        c.execute("PRAGMA page_size=32768")
        c.execute("PRAGMA journal_mode=OFF")
        c.execute("PRAGMA synchronous=OFF")
        self.headers = list(info.headers)
        n = self.ncols
        if n == 0:
            raise ValueError("No header row found in the CSV.")
        c.execute(f"CREATE TABLE data ({', '.join(f'c{i} TEXT' for i in range(n))})")
        ins = f"INSERT INTO data VALUES ({','.join('?' * n)})"
        size = max(1, info.size)
        total = fixed = 0
        last = 0.0
        with open(info.path, "rb") as fb:
            ft = io.TextIOWrapper(fb, encoding=info.encoding, errors="replace", newline="")
            reader = csv.reader(ft, delimiter=info.delimiter)
            next(reader, None)
            batch = []
            c.execute("BEGIN")
            for row in reader:
                if not row:
                    continue
                L = len(row)
                if L != n:
                    fixed += 1
                    if L < n:
                        row = row + [None] * (n - L)
                    else:  # extra fields are joined into the last column so nothing is lost
                        row = row[:n - 1] + [info.delimiter.join(row[n - 1:])]
                batch.append(row)
                if len(batch) >= 20000:
                    c.executemany(ins, batch)
                    total += len(batch)
                    batch.clear()
                    if cancel.is_set():
                        raise Cancelled()
                    now = time.time()
                    if now - last > 0.2:
                        progress(fb.tell() / size, f"Importing… {total:,} rows")
                        last = now
            if batch:
                c.executemany(ins, batch)
                total += len(batch)
            c.execute("COMMIT")
        self.row_count = total

        self.fts = ""
        if use_fts:
            progress(None, f"{total:,} rows imported. Building the full-text index… (this can take a while)")
            try:
                c.execute(f"CREATE VIRTUAL TABLE data_fts USING fts5({', '.join(self.cols)}, "
                          f"content='data', content_rowid='rowid', tokenize='trigram')")
                c.execute("INSERT INTO data_fts(data_fts) VALUES('rebuild')")
                self.fts = "trigram"
            except sqlite3.OperationalError as e:
                if is_cancel_error(e):
                    raise
                self.fts = ""  # this SQLite build has no trigram tokenizer
        self._create_view()
        self._write_meta(info, {"engine": "sqlite", "fts": self.fts, "fixed_rows": fixed})
        self.indexed = set()
        return {"rows": total, "fixed_rows": fixed, "fts": self.fts}

    def can_index(self):
        return True

    def create_index(self, col):
        self.con.execute(f"CREATE INDEX IF NOT EXISTS idx_c{col} ON data(c{col})")
        self.indexed.add(col)

    def _col_expr(self, c, op, text, cs):
        if op == "contains":
            return (f"instr({c}, ?) > 0" if cs else f"instr(lower({c}), lower(?)) > 0"), [text]
        if op == "equals":
            return (f"{c} = ?" if cs else f"{c} = ? COLLATE NOCASE"), [text]
        if op == "startswith":
            if cs:  # range form so a column index can be used
                return f"({c} >= ? AND {c} < ?)", [text, text + "\U0010ffff"]
            return f"{c} LIKE ? ESCAPE '\\'", [like_escape(text) + "%"]
        if op == "regex":
            return f"{c} REGEXP ?", [text if cs else "(?i)" + text]
        if op == "pattern":
            if cs:
                glob = text.replace("[", "[[]").replace("?", "[?]")
                return f"{c} GLOB ?", [glob]
            return f"{c} LIKE ? ESCAPE '\\'", [like_escape(text).replace("*", "%")]
        raise ValueError(op)

    def _term_sql(self, col, op, text, cs):
        where, params, _ = super()._term_sql(col, op, text, cs)
        if col is not None and cs and op in ("equals", "startswith") and col in self.indexed:
            return where, params, "index"
        lit = max(text.split("*"), key=len) if op == "pattern" else text
        if self.fts == "trigram" and op != "regex" and len(lit) >= 3:
            q = '"' + lit.replace('"', '""') + '"'
            if col is not None:
                q = f"c{col} : {q}"
            fts = "rowid IN (SELECT rowid FROM data_fts WHERE data_fts MATCH ?)"
            if op == "contains" and not cs:
                return fts, [q], "fts"
            return f"({fts} AND {where})", [q] + params, "fts"
        return where, params, "scan"

    def _method_text(self, stats, scope):
        if len(stats) == 1:
            m = next(iter(stats))
            if m == "fts":
                return "uses the full-text index (fast)"
            if m == "index":
                return "uses the column index (instant)"
            return "scans every row" + ("" if self.fts else " (no full-text index in this database)")
        parts = []
        for m, label in (("fts", "full-text index"), ("index", "column index"), ("scan", "row scan")):
            if stats.get(m):
                parts.append(f"{stats[m]} by {label}")
        return ", ".join(parts)

    def _simple_text(self, method, col, op, cs):
        if method == "fts":
            exact = op == "contains" and not cs
            return "Method: full-text index" + (" (very fast)" if exact else ", then an exact check (fast)")
        if method == "index":
            return "Method: column index (instant)"
        scope = "every column" if col is None else "this column"
        hint = ""
        if col is not None and op in ("equals", "startswith"):
            if col not in self.indexed:
                hint = "   ·   Tip: create an index on this column and turn on Match case for instant results"
            elif not cs:
                hint = "   ·   Tip: turn on Match case to use the column index"
        elif self.fts == "trigram" and op in ("contains", "equals", "startswith"):
            hint = "   ·   Tip: 3 or more characters use the full-text index"
        elif not self.fts and op == "contains":
            hint = "   ·   This database has no full-text index"
        return f"Method: scanning {scope}{hint}"


# ---------------------------------------------------------------------------
# DuckDB
# ---------------------------------------------------------------------------
class DuckDbEngine(BaseEngine):
    kind = "duckdb"
    label = "DuckDB"
    ROWNUM = "(rowid + 1)"

    def _connect(self):
        if duckdb is None:
            raise RuntimeError("The DuckDB module is not available.")
        self.con = duckdb.connect(self.db_path)

    def open(self):
        self._connect()
        self._load_meta()

    def import_csv(self, info, use_fts, progress, cancel):
        remove_db_files(self.db_path)
        self.headers = list(info.headers)
        n = self.ncols
        if n == 0:
            raise ValueError("No header row found in the CSV.")
        src, tmp = info.path, None
        if info.encoding not in ("utf-8", "utf-8-sig"):
            tmp = self.db_path + ".utf8.tmp.csv"
            convert_to_utf8(info.path, info.encoding, tmp, progress, cancel)
            src = tmp
        try:
            self._connect()
            columns = "{" + ", ".join(f"'c{i}': 'VARCHAR'" for i in range(n)) + "}"
            sql = (f"CREATE TABLE data AS SELECT * FROM read_csv({qlit(src)}, "
                   f"delim={qlit(info.delimiter)}, header=true, columns={columns}, "
                   f"quote='\"', escape='\"', null_padding=true, strict_mode=false, "
                   f"max_line_size=100000000)")
            progress(None, "Importing with DuckDB…")
            done = threading.Event()

            def watch():  # show DuckDB's own progress
                while not done.wait(0.3):
                    try:
                        p = self.con.query_progress()
                        if p is not None and p >= 0:
                            progress(min(p, 100) / 100, f"Importing with DuckDB… {p:.0f}%")
                    except Exception:
                        pass

            t = threading.Thread(target=watch, daemon=True)
            t.start()
            try:
                self.con.execute(sql)
            finally:
                done.set()
                t.join(1)
        finally:
            if tmp and os.path.exists(tmp):
                try:
                    os.remove(tmp)
                except OSError:
                    pass
        self.row_count = self.con.execute("SELECT count(*) FROM data").fetchone()[0]
        self._create_view()
        self._write_meta(info, {"engine": "duckdb", "fts": ""})
        self.con.execute("CHECKPOINT")
        return {"rows": self.row_count, "fixed_rows": 0, "fts": ""}

    def _col_expr(self, c, op, text, cs):
        lc = c if cs else f"lower({c})"
        val = text if cs else text.lower()
        if op == "contains":
            return f"contains({lc}, ?)", [val]
        if op == "equals":
            return f"{lc} = ?", [val]
        if op == "startswith":
            return f"starts_with({lc}, ?)", [val]
        if op == "regex":
            return (f"regexp_matches({c}, ?)" if cs else f"regexp_matches({c}, ?, 'i')"), [text]
        if op == "pattern":
            like = "LIKE" if cs else "ILIKE"
            return f"{c} {like} ? ESCAPE '\\'", [like_escape(text).replace("*", "%")]
        raise ValueError(op)

    def _method_text(self, stats, scope):
        if scope == "query":
            return "DuckDB parallel scan (all CPU cores)"
        where = "every column" if scope is None else "this column"
        return f"DuckDB parallel scan of {where} (all CPU cores)"

    def _simple_text(self, method, col, op, cs):
        return "Method: " + self._method_text({method: 1}, col)


# ---------------------------------------------------------------------------
def new_engine(kind, db_path):
    return DuckDbEngine(db_path) if kind == "duckdb" else SqliteEngine(db_path)


def open_database(path):
    with open(path, "rb") as f:
        head = f.read(16)
    eng = SqliteEngine(path) if head.startswith(b"SQLite format 3") else DuckDbEngine(path)
    try:
        eng.open()
    except Exception:
        eng.close()
        raise
    return eng
