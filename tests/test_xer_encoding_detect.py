"""XER kod sayfasi saptama testleri (offline).

Regresyon kaynagi: BOM'suz ANSI bir XER, `utf-8 errors='replace'` ile
okunuyordu; cp1251 Kiril metnin TAMAMI sessizce U+FFFD oluyordu. 950 aktivitelik
bukhtourcity programinda 529 gorev + 785 WBS adi bu yolla kayboldu. Testler
hem dogru saptamayi hem de "emin degilsen sus" davranisini kilitler.
"""
import os
import tempfile

import pytest

import xer_parser


HEADER = "ERMHDR\t19.12\t2026-08-25\tadmin\tP6\tUSD\n"


def _xer_body(task_name: str) -> str:
    return (
        HEADER
        + "%T\tPROJECT\n"
        + "%F\tproj_id\tproj_short_name\tclndr_id\tlast_recalc_date\n"
        + "%R\t1\tTEST\t9\t2026-09-02\n"
        + "%T\tCALENDAR\n"
        + "%F\tclndr_id\tclndr_name\tday_hr_cnt\tdefault_flag\n"
        + "%R\t9\tStandard\t8\tY\n"
        + "%T\tTASK\n"
        + "%F\ttask_id\tproj_id\twbs_id\ttask_code\ttask_name\ttask_type"
          "\tstatus_code\ttarget_drtn_hr_cnt\ttotal_float_hr_cnt\n"
        + "%R\t100\t1\t10\tA001\t" + task_name + "\tTT_Task\tTK_NotStart\t80\t0\n"
        + "%E\n"
    )


def _write(content: str, encoding: str, name: str, bom: bytes = b"") -> str:
    path = os.path.join(tempfile.gettempdir(), name)
    with open(path, "wb") as fh:
        fh.write(bom)
        fh.write(content.encode(encoding))
    return path


def _read_name(path: str, **kw) -> tuple[str, xer_parser.XerFile]:
    x = xer_parser.XerFile(path, **kw)
    return x.tables["TASK"]["rows"][0]["task_name"], x


# --- dogru cozumleme -------------------------------------------------------
def test_utf16le_bom_okunur():
    path = _write(_xer_body("Гранит"), "utf-16-le", "enc_u16bom.xer", b"\xff\xfe")
    try:
        name, x = _read_name(path)
        assert name == "Гранит"
        assert x.encoding == "utf-16-le" and x.encoding_source == "BOM"
    finally:
        os.remove(path)


def test_utf8_bom_okunur():
    path = _write(_xer_body("Şantiye Şefliği"), "utf-8", "enc_u8bom.xer", b"\xef\xbb\xbf")
    try:
        name, x = _read_name(path)
        assert name == "Şantiye Şefliği"
        assert x.encoding == "utf-8" and x.encoding_source == "BOM"
    finally:
        os.remove(path)


def test_bomsuz_utf8_okunur():
    path = _write(_xer_body("Кирилл ÖÇŞ"), "utf-8", "enc_u8raw.xer")
    try:
        name, x = _read_name(path)
        assert name == "Кирилл ÖÇŞ"
        assert x.encoding == "utf-8"
    finally:
        os.remove(path)


def test_ansi_cp1251_kiril_saptanir():
    """Asil regresyon: bu dosya eskiden bastan sona U+FFFD donuyordu."""
    path = _write(_xer_body("ЗЕМЛЯНЫЕ РАБОТЫ / EARTHWORKS"), "cp1251",
                  "enc_1251.xer")
    try:
        name, x = _read_name(path)
        assert name == "ЗЕМЛЯНЫЕ РАБОТЫ / EARTHWORKS"
        assert "�" not in name
        assert x.encoding == "cp1251"
        assert x.encoding_source.startswith("sezgisel")
        assert x.encoding_scores["cp1251"] > x.encoding_scores["cp1252"]
    finally:
        os.remove(path)


def test_ansi_turkce_kirilin_onune_gecer():
    """cp1254 metni cp1251 olarak okunmamali.

    cp1252/cp1254/cp1250 birbirinden AYIRT EDILEMEZ (ayni baytlar her birinde
    gecerli bir kelime uretir: Şantiye / Þantiye / Ţantiye), o yuzden aralarinda
    esitlik beklenir; onemli olan Kiril adayinin elenmesi.
    """
    path = _write(_xer_body("Şantiye Şefliği ÇÖĞÜ İşleri Ölçüm Beton"), "cp1254",
                  "enc_1254.xer")
    try:
        _name, x = _read_name(path)
        assert x.encoding != "cp1251", x.encoding_scores
        assert x.encoding_scores["cp1254"] > x.encoding_scores["cp1251"]
        assert x.encoding_confidence == "dusuk"
    finally:
        os.remove(path)


def test_latin_kod_sayfalari_esit_ve_dusuk_guven():
    path = _write(_xer_body("Ölçüm Çalışması Beton"), "cp1254", "enc_tie.xer")
    try:
        _name, x = _read_name(path)
        tied = [e for e, s in x.encoding_scores.items()
                if s == max(x.encoding_scores.values())]
        assert len(tied) > 1 and "cp1251" not in tied
        assert x.encoding_confidence == "dusuk"
        assert "esitlik" in x.encoding_source
    finally:
        os.remove(path)


def test_saf_ascii_bomsuz_bozulmaz():
    path = _write(_xer_body("EARTHWORKS PHASE 1"), "ascii", "enc_ascii.xer")
    try:
        name, _x = _read_name(path)
        assert name == "EARTHWORKS PHASE 1"
    finally:
        os.remove(path)


# --- acik parametre --------------------------------------------------------
def test_encoding_parametresi_sezgiyi_ezer():
    path = _write(_xer_body("Гранит"), "cp1251", "enc_param.xer")
    try:
        _name, x = _read_name(path, encoding="cp1251")
        assert x.encoding_source == "parametre"
    finally:
        os.remove(path)


# --- belirsizlikte sus -----------------------------------------------------
def test_cozulemeyen_icerik_sessizce_bos_donmez():
    """Anlamsiz yuksek baytlar: her cp bunlari cozer ama skor negatif kalmali.

    cp1250/1251/1252/1254'un dordunu birden reddeden bir bayt dizisi yok, o
    yuzden "hata firlat" yerine dogru invariant su: boyle bir icerik yuksek
    guvenle secilmis gibi raporlanmaz.
    """
    junk = "".join(chr(b) for b in range(0xA0, 0xC0)) * 5
    assert xer_parser.XerFile._score_ansi(junk) < 0


def test_dusuk_guven_isaretlenir():
    """Esitlikte secim yapiliyorsa bu acikca bildirilmeli."""
    path = _write(_xer_body("Ölçüm"), "cp1254", "enc_lowconf.xer")
    try:
        _name, x = _read_name(path)
        assert x.encoding_confidence in ("dusuk", "yuksek")
        if x.encoding_confidence == "dusuk":
            assert "esitlik" in x.encoding_source
    finally:
        os.remove(path)


def test_utf16le_yanlis_pozitif_vermez():
    """Kisa ASCII icerik utf-16-le olarak 'gecerli' cozulur ama XER degildir."""
    path = _write(_xer_body("ABCD"), "ascii", "enc_u16fp.xer")
    try:
        _name, x = _read_name(path)
        assert x.encoding != "utf-16-le"
    finally:
        os.remove(path)


# --- skorlama --------------------------------------------------------------
def test_skor_kiril_latin1_uzerinde_tutar():
    text_cyr = "Гранит Брусчатка"
    text_mojibake = text_cyr.encode("cp1251").decode("cp1252")
    assert (xer_parser.XerFile._score_ansi(text_cyr)
            > xer_parser.XerFile._score_ansi(text_mojibake))


def test_skor_simgeleri_cezalandirir():
    assert xer_parser.XerFile._score_ansi("×÷¤°±µ") < 0
