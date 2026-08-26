"""XER yazici bicimlendirme kurallari + ERMHDR ayristirma (veritabanisiz).

ERMHDR testleri bir regresyondan geliyor: parser basligi bes alan varsayip
5. alani para birimi sayiyordu. Gercek bir P6 export'u sekiz alan tasir ve
para birimi SONDADIR; o dosyada 5. alan "Izzat Islomov" idi, yani
`currency_validator.extract_currency_code` bir kisi adini para birimi kodu
olarak donduruyordu. Cost/hours kararlari bu koda bakiyor.
"""
import datetime as _dt
import decimal

import pytest

import currency_validator
import xer_parser
from p6 import writer


# --- ERMHDR ----------------------------------------------------------------
REAL = ("ERMHDR\t19.12\t2026-08-21\tProject\tIzzat1199\tIzzat Islomov\t"
        "dbxDatabaseNoName\tProject Management Cloud\tUSD")


def test_gercek_p6_basligi_para_birimini_sondan_alir():
    h = xer_parser._parse_ermhdr(REAL)
    assert h["currency"] == "USD"
    assert h["version"] == "19.12"
    assert h["exported"] == "2026-08-21"
    assert h["user"] == "Izzat1199"
    assert h["user_name"] == "Izzat Islomov"
    assert h["app"] == "Project Management Cloud"


def test_para_birimi_kodu_kisi_adi_dondurmez():
    """Asil regresyon."""
    code = currency_validator.extract_currency_code(
        xer_parser._parse_ermhdr(REAL))
    assert code == "USD"
    assert code != "Izzat Islomov"


def test_kisa_baslik_da_calisir():
    h = xer_parser._parse_ermhdr("ERMHDR\t19.12\t2026-08-21\tadmin\tP6\tEUR")
    assert h["currency"] == "EUR"
    assert h["version"] == "19.12"


def test_ham_alanlar_saklanir():
    h = xer_parser._parse_ermhdr(REAL)
    assert h["fields"][0] == "19.12" and h["fields"][-1] == "USD"


def test_eksik_baslik_patlamaz():
    h = xer_parser._parse_ermhdr("ERMHDR")
    assert h["version"] == "" and h["currency"] == ""


def test_yazicinin_basligi_geri_okunabilir():
    line = writer._header("USD", "MCP", "PMDB")
    h = xer_parser._parse_ermhdr(line)
    assert h["currency"] == "USD"
    assert h["version"] == writer.HEADER_VERSION


# --- deger bicimlendirme ---------------------------------------------------
@pytest.mark.parametrize("value,expected", [
    (None, ""),
    ("", ""),
    ("Гранит", "Гранит"),
    (42, "42"),
    (True, "Y"),
    (False, "N"),
    (_dt.date(2026, 9, 24), "2026-09-24"),
    (_dt.datetime(2026, 9, 24, 0, 0), "2026-09-24 00:00"),
    (_dt.datetime(2026, 9, 24, 17, 30), "2026-09-24 17:30"),
])
def test_deger_bicimlendirme(value, expected):
    assert writer._fmt(value) == expected


def test_ondalik_kuyruk_sifirlari_kirpilir():
    assert writer._fmt(decimal.Decimal("160.000000")) == "160"
    assert writer._fmt(decimal.Decimal("5.0000")) == "5"
    assert writer._fmt(decimal.Decimal("0.80000000")) == "0.8"


def test_sifir_ondalik_sifir_kalir():
    assert writer._fmt(decimal.Decimal("0E-8")) == "0"
    assert writer._fmt(decimal.Decimal("0.000000")) == "0"


def test_sekme_ve_satir_sonu_satir_gramerini_bozmaz():
    """Kacan bir sekme XER satirini kaydirir; alani atmak daha kotu olurdu."""
    assert "\t" not in writer._fmt("a\tb")
    assert "\n" not in writer._fmt("a\nb")
    assert writer._fmt("a\tb") == "a b"
    assert writer._fmt("a\r\nb") == "a  b"


# --- tablo secimi ----------------------------------------------------------
def test_yerel_kolonlar_disari_yazilmaz():
    assert "delete_session_id" in writer.SKIP_COLUMNS
    assert "delete_date" in writer.SKIP_COLUMNS


def test_schedoptions_projprop_a_esler():
    """XER bolum adi ile veritabani tablo adi ayni degil."""
    assert writer.XER_TO_DB_TABLE["SCHEDOPTIONS"] == "PROJPROP"


def test_tablo_sirasi_ebeveyni_cocuktan_once_yazar():
    order = list(writer.TABLE_ORDER)
    assert order.index("PROJECT") < order.index("PROJWBS")
    assert order.index("PROJWBS") < order.index("TASK")
    assert order.index("TASK") < order.index("TASKPRED")
    assert order.index("TASK") < order.index("TASKRSRC")
    assert order.index("RSRC") < order.index("RSRCRATE")
    assert order.index("RSRC") < order.index("TASKRSRC")


def test_proje_kapsamli_tablolar_isaretli():
    for t in ("PROJECT", "PROJWBS", "TASK", "TASKPRED", "TASKRSRC"):
        assert t in writer.PROJECT_SCOPED
    for t in ("CURRTYPE", "RSRC", "CALENDAR"):
        assert t not in writer.PROJECT_SCOPED


# --- korumalar -------------------------------------------------------------
def test_yol_yoksa_reddedilir():
    with pytest.raises(writer.WriterError, match="Cikti yolu"):
        writer.write_xer({"proj_id": 1})


def test_proj_id_yoksa_reddedilir():
    with pytest.raises(writer.WriterError, match="proj_id"):
        writer.write_xer({"path": "x.xer"})


def test_var_olan_dosyanin_uzerine_yazilmaz(tmp_path):
    p = tmp_path / "var.xer"
    p.write_bytes(b"eski")
    with pytest.raises(writer.WriterError, match="overwrite"):
        writer.write_xer({"proj_id": 1, "path": str(p)})
    assert p.read_bytes() == b"eski"
