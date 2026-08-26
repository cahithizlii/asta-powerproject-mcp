"""p6/compare.py eslesme kurallari -- veritabanisiz birim testleri.

Tek bir kural bu dosyanin tamamini hakli cikariyor: **iki taraf task_code ile
eslesir, task_id ile DEGIL.** P6 bir program her sinir gectiginde (CLI import,
baseline kopyasi, export-yeniden import) id'leri yeniden numaralandirir; ayni
aktivite XER'de 3274452, import sonrasi veritabaninda 35847, o projenin
baseline'inda ucuncu bir sayidir. id ile eslesen bir karsilastirma 950
aktivitenin 950'sini birden "silinmis" ve "eklenmis" gosterir -- kesin
gorunen ama hicbir sey soylemeyen bir rapor uretir.
"""
import pytest

from p6 import compare


def task(code, tid, **kw):
    base = {"id": tid, "code": code, "name": "A", "duration_h": 80,
            "percent_complete": 0, "summary": False}
    base.update(kw)
    return base


def data(tasks, links=(), **kw):
    d = {"tasks": list(tasks), "links": list(links), "units": "qty",
         "day_hr_cnt": 8.0, "status_date": "2026-11-01", "task_count": len(tasks),
         "source": {"type": "db"}}
    d.update(kw)
    return d


# --- yeniden anahtarlama ---------------------------------------------------
def test_gorevler_koda_gore_yeniden_anahtarlanir():
    tasks, _links, notes = compare._rekey(
        data([task("A001", 111), task("A002", 222)]))
    assert [t["id"] for t in tasks] == ["A001", "A002"]
    assert [t["task_id_original"] for t in tasks] == [111, 222]
    assert notes["duplicate_codes"] == 0


def test_baglar_da_koda_cevrilir():
    d = data([task("A001", 111), task("A002", 222)],
             links=[{"from_id": 111, "to_id": 222, "type": "FS", "lag_days": 0}])
    _tasks, links, notes = compare._rekey(d)
    assert links == [{"from_id": "A001", "to_id": "A002", "type": "FS",
                      "lag_days": 0}]
    assert notes["links_outside_project"] == 0


def test_proje_disina_giden_bag_atlanir_ve_sayilir():
    d = data([task("A001", 111)],
             links=[{"from_id": 111, "to_id": 999, "type": "FS", "lag_days": 0}])
    _t, links, notes = compare._rekey(d)
    assert links == []
    assert notes["links_outside_project"] == 1


def test_kodsuz_gorev_atlanir():
    tasks, _l, _n = compare._rekey(data([task("A001", 111), task("", 222)]))
    assert [t["id"] for t in tasks] == ["A001"]


def test_tekrar_eden_kod_bir_kez_alinir_ve_bildirilir():
    tasks, _l, notes = compare._rekey(
        data([task("A001", 111), task("A001", 222), task("A002", 333)]))
    assert [t["id"] for t in tasks] == ["A001", "A002"]
    assert notes["duplicate_codes"] == 1


def test_id_farkli_olsa_da_iki_taraf_eslesir():
    """Asil regresyon: yeniden numaralandirma karsilastirmayi bozmamali."""
    import xer_compare

    a, _la, _na = compare._rekey(data([task("A001", 1), task("A002", 2)]))
    b, _lb, _nb = compare._rekey(
        data([task("A001", 90001), task("A002", 90002, duration_h=120)]))
    diff = xer_compare.diff_tasks(a, b, fields=["duration_h"])
    assert diff["added"] == [] and diff["removed"] == []
    assert len(diff["changed"]) == 1
    assert diff["changed"][0]["id"] == "A002"


# --- kaynak cozumleme ------------------------------------------------------
def test_a_b_nesne_bicimi():
    p = compare._side_params({"a": {"proj_id": 368}, "b": {"proj_id": 369}}, "a")
    assert p["proj_id"] == 368


def test_duz_sonek_bicimi():
    p = compare._side_params({"proj_id_a": 368, "path_b": "x.xer"}, "b")
    assert p["path"] == "x.xer" and "proj_id" not in p


def test_baseline_proj_id_projeye_cevrilir():
    """Baseline kendi basina bir projedir; proje+baseline olarak yuklenmemeli."""
    p = compare._side_params({"a": {"baseline_proj_id": 369}}, "a")
    assert p["proj_id"] == 369 and "baseline_proj_id" not in p


def test_ortak_parametreler_taraflara_gecer():
    p = compare._side_params({"a": {"proj_id": 368}, "alias": "X",
                              "units": "cost"}, "a")
    assert p["alias"] == "X" and p["units"] == "cost"


def test_taraf_ozel_parametre_ortaki_ezer():
    p = compare._side_params({"a": {"proj_id": 368, "units": "qty"},
                              "units": "cost"}, "a")
    assert p["units"] == "qty"


def test_kaynaksiz_taraf_reddedilir():
    with pytest.raises(compare.CompareError, match="kaynak verilmedi"):
        compare._side_params({"b": {"proj_id": 1}}, "a")


# --- baglam ve uyarilar ----------------------------------------------------
def test_farkli_birim_uyarisi():
    ctx = compare._context(data([], units="cost"), data([], units="qty"), {}, {})
    assert any("farkli birimde" in w for w in ctx["warnings"])


def test_farkli_veri_tarihi_uyarisi():
    ctx = compare._context(data([], status_date="2026-09-02"),
                           data([], status_date="2026-11-01"), {}, {})
    assert any("Veri tarihleri farkli" in w for w in ctx["warnings"])


def test_farkli_takvim_uyarisi():
    ctx = compare._context(data([], day_hr_cnt=8.0), data([], day_hr_cnt=9.0),
                           {}, {})
    assert any("Gun-saat farkli" in w for w in ctx["warnings"])


def test_ayni_taraflarda_uyari_yok():
    ctx = compare._context(data([]), data([]), {}, {})
    assert ctx["warnings"] == []
    assert ctx["join_key"] == "task_code"


def test_tekrar_eden_kod_baglamda_uyari_olur():
    ctx = compare._context(data([]), data([]), {"duplicate_codes": 3}, {})
    assert any("tekrar ediyor" in w for w in ctx["warnings"])


# --- kirpma ----------------------------------------------------------------
def test_trim_sayilari_korur_listeyi_keser():
    out = compare._trim({"added": list(range(10)), "note": "x"}, 3)
    assert out["added_count"] == 10
    assert out["added"] == [0, 1, 2]
    assert out["added_truncated"] is True
    assert out["note"] == "x"


def test_trim_kisa_listede_truncated_koymaz():
    out = compare._trim({"added": [1, 2]}, 5)
    assert out["added_count"] == 2 and "added_truncated" not in out
