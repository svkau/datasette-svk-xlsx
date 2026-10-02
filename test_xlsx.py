import io
import sqlite3

import pytest
from datasette.app import Datasette
from openpyxl import load_workbook


@pytest.fixture
def ds(tmp_path):
    path = str(tmp_path / "test.db")
    conn = sqlite3.connect(path)
    conn.execute("create table personer (id integer primary key, namn text, nr integer, f text)")
    conn.executemany(
        "insert into personer (namn, nr, f) values (?, ?, ?)",
        [(f"Person {i}", 199001011234 + i, "=1+1" if i == 0 else "ö\x01k")
         for i in range(2500)],
    )
    conn.execute("create view v as select * from personer where id % 2 = 0")
    conn.commit()
    return Datasette([path], settings={"max_returned_rows": 1000},
                     metadata={"databases": {"test": {"queries": {
                         "udda": "select id, namn from personer where id % 2 = 1"}}}})


def sheet(resp):
    assert resp.status_code == 200, resp.text
    wb = load_workbook(io.BytesIO(resp.content))
    return wb, wb.worksheets[0]


@pytest.mark.asyncio
async def test_table_all_pages(ds):
    wb, ws = sheet(await ds.client.get("/test/personer.xlsx"))
    assert ws.max_row == 2501
    assert [c.value for c in ws[1]] == ["id", "namn", "nr", "f"]
    assert ws["D2"].value == "=1+1" and ws["D2"].data_type == "s"
    assert ws["D3"].value == "ök"
    assert ws["C2"].value == 199001011234
    assert "Om uttaget" in wb.sheetnames


@pytest.mark.asyncio
async def test_filtered_sorted(ds):
    wb, ws = sheet(await ds.client.get(
        "/test/personer.xlsx?id__gt=2000&_sort_desc=id&_col=namn"))
    assert ws.max_row == 501
    assert [c.value for c in ws[1]] == ["id", "namn"]
    assert ws["A2"].value == 2500


@pytest.mark.asyncio
async def test_view(ds):
    wb, ws = sheet(await ds.client.get("/test/v.xlsx"))
    assert ws.max_row == 1251


@pytest.mark.asyncio
async def test_sql_not_truncated(ds):
    wb, ws = sheet(await ds.client.get(
        "/test.xlsx?sql=select+*+from+personer+where+id+>+:min&min=10"))
    assert ws.max_row == 2491


@pytest.mark.asyncio
async def test_canned(ds):
    wb, ws = sheet(await ds.client.get("/test/udda.xlsx"))
    assert ws.max_row == 1251


@pytest.mark.asyncio
async def test_row(ds):
    wb, ws = sheet(await ds.client.get("/test/personer/5.xlsx"))
    assert ws.max_row == 2 and ws["A2"].value == 5


@pytest.mark.asyncio
async def test_link_on_page(ds):
    html = (await ds.client.get("/test/personer")).text
    assert "personer.xlsx" in html


@pytest.mark.asyncio
async def test_permissions_forwarded(tmp_path):
    path = str(tmp_path / "p.db")
    conn = sqlite3.connect(path)
    conn.execute("create table t (id integer primary key)")
    conn.executemany("insert into t values (?)", [(i,) for i in range(1500)])
    conn.commit()
    ds = Datasette([path], settings={"max_returned_rows": 1000},
                   metadata={"databases": {"p": {"allow": {"id": "henrik"}}}})
    assert (await ds.client.get("/p/t.xlsx")).status_code == 403
    cookies = {"ds_actor": ds.sign({"a": {"id": "henrik"}}, "actor")}
    resp = await ds.client.get("/p/t.xlsx", cookies=cookies)
    wb, ws = sheet(resp)
    assert ws.max_row == 1501
