"""p6/analysis.py karar kurallari -- veritabanisiz birim testleri.

Uc karar burada kilitleniyor; ucu de yanlis yapildiginda program bozuk degil
SESSIZCE YANLIS raporlanir:

1. **Yuzde tamamlanma tabani.** P6 aktivite basina uc ayri tamamlanma tutar ve
   hangisinin gecerli oldugunu `complete_pct_type` soyler. phys_complete_pct'i
   dogrudan okumak sure bazli (CP_Drtn) bir programda her aktiviteyi %0
   gosterir -- bukhtourcity'nin tamami CP_Drtn'dir, yani EV sifir cikardi.
2. **BAC'in birimi.** Ayni program XER'den 353.160 (maliyet), veritabanindan
   70.632 (saat) BAC veriyordu; birim bildirilmeden verilen BAC, ALFB1 9x
   hatasinin ta kendisidir (RULE 16.A).
3. **WBS yol eslesmesi.** wbs_short_name benzersiz DEGILDIR; '1', '2', '3'
   farkli ust dugumler altinda tekrar eder.
"""
import datetime as _dt

import pytest

from p6 import analysis


class FakeBag:
    def __init__(self, tasks=(), assignments=(), project=()):
        self.tables = {
            "TASK": {"headers": [], "rows": list(tasks)},
            "TASKRSRC": {"headers": [], "rows": list(assignments)},
            "PROJECT": {"headers": [], "rows": list(project)},
        }


def t(task_id, pct_type="CP_Drtn", target=80, remain=80, phys=0,
      status="TK_NotStart"):
    return {"task_id": str(task_id), "complete_pct_type": pct_type,
            "target_drtn_hr_cnt": str(target), "remain_drtn_hr_cnt": str(remain),
            "phys_complete_pct": str(phys), "status_code": status}


def a(task_id, act=0, remain=0):
    return {"task_id": str(task_id), "act_reg_qty": str(act),
            "act_ot_qty": "0", "remain_qty": str(remain)}


# --- 1) yuzde tamamlanma tabani -------------------------------------------
def test_cp_drtn_kalan_sureden_hesaplar():
    pct, basis = analysis.resolve_percent_complete(
        FakeBag(tasks=[t(1, "CP_Drtn", target=240, remain=72,
                         status="TK_Active")]))
    assert pct[1] == pytest.approx(70.0)
    assert basis["CP_Drtn"] == 1


def test_cp_phys_kayitli_degeri_kullanir():
    pct, _ = analysis.resolve_percent_complete(
        FakeBag(tasks=[t(1, "CP_Phys", target=240, remain=240, phys=45,
                         status="TK_Active")]))
    assert pct[1] == pytest.approx(45.0)


def test_cp_units_atamalardan_hesaplar():
    pct, _ = analysis.resolve_percent_complete(
        FakeBag(tasks=[t(1, "CP_Units", status="TK_Active")],
                assignments=[a(1, act=30, remain=10)]))
    assert pct[1] == pytest.approx(75.0)


def test_hesaplanamayinca_phys_complete_pct_e_duser():
    pct, basis = analysis.resolve_percent_complete(
        FakeBag(tasks=[t(1, "CP_Drtn", target=0, remain=0, phys=33,
                         status="TK_Active")]))
    assert pct[1] == pytest.approx(33.0)
    assert any("dusuldu" in k for k in basis)


def test_tamamlanan_her_zaman_yuz():
    pct, _ = analysis.resolve_percent_complete(
        FakeBag(tasks=[t(1, "CP_Drtn", target=80, remain=40,
                         status="TK_Complete")]))
    assert pct[1] == 100.0


def test_baslamamis_her_zaman_sifir():
    pct, _ = analysis.resolve_percent_complete(
        FakeBag(tasks=[t(1, "CP_Drtn", target=80, remain=0, phys=90,
                         status="TK_NotStart")]))
    assert pct[1] == 0.0


def test_yuzde_0_100_araliginda_kalir():
    pct, _ = analysis.resolve_percent_complete(
        FakeBag(tasks=[t(1, "CP_Drtn", target=80, remain=-40, status="TK_Active"),
                       t(2, "CP_Drtn", target=80, remain=200, status="TK_Active")]))
    assert pct[1] == 100.0 and pct[2] == 0.0


def test_taban_dagilimi_raporlanir():
    _pct, basis = analysis.resolve_percent_complete(
        FakeBag(tasks=[t(1, "CP_Drtn"), t(2, "CP_Drtn"), t(3, "CP_Phys")]))
    assert basis["CP_Drtn"] == 2 and basis["CP_Phys"] == 1


# --- 2) birim secimi -------------------------------------------------------
def _tasks(duration_h=100):
    return [{"duration_h": duration_h, "summary": False}]


def _asgs(qty=0.0, cost=0.0):
    return [{"target_qty": qty, "target_cost": cost}]


def test_maliyet_yuklu_ise_cost_secilir():
    u = analysis.resolve_units(None, _tasks(), _asgs(qty=100, cost=500))
    assert u["units"] == "cost"
    assert u["candidate_bac"] == {"cost": 500.0, "qty": 100.0, "duration_h": 100.0}


def test_maliyet_sifirsa_saate_duser_ve_uyarir():
    """detect_mode 'cost' der ama toplam target_cost 0 -- celiski bildirilmeli."""
    u = analysis.resolve_units(None, _tasks(), _asgs(qty=100, cost=0))
    assert u["units"] == "qty"
    assert u["cost_loading_mode"] in ("cost", "mixed")
    assert any("target_cost = 0" in wmsg for wmsg in u["units_warnings"])


def test_atama_yoksa_gorev_suresine_duser():
    u = analysis.resolve_units(None, _tasks(120), [])
    assert u["units"] == "duration_h"
    assert u["candidate_bac"]["duration_h"] == 120.0


def test_sure_tabani_farkli_kaynak_uyarisi_verir():
    u = analysis.resolve_units("duration_h", _tasks(), _asgs(qty=50, cost=0))
    assert u["units"] == "duration_h"
    assert any("iki farkli" in wmsg for wmsg in u["units_warnings"])


def test_acik_birim_sezgiyi_ezer():
    u = analysis.resolve_units("cost", _tasks(), _asgs(qty=100, cost=500))
    assert u["units"] == "cost" and u["units_reason"] == "parametre"


def test_gecersiz_birim_reddedilir():
    with pytest.raises(analysis.AnalysisError, match="gecersiz"):
        analysis.resolve_units("saat", _tasks(), _asgs())


def test_ozet_gorevler_sure_toplamina_girmez():
    tasks = [{"duration_h": 100, "summary": False},
             {"duration_h": 900, "summary": True}]
    u = analysis.resolve_units(None, tasks, [])
    assert u["candidate_bac"]["duration_h"] == 100.0


# --- 3) WBS yol eslesmesi --------------------------------------------------
def test_wbs_yolu_tekrar_eden_kisa_adlari_ayirir():
    rows = [
        {"wbs_id": "1", "parent_wbs_id": "", "wbs_short_name": "PROJ"},
        {"wbs_id": "2", "parent_wbs_id": "1", "wbs_short_name": "A"},
        {"wbs_id": "3", "parent_wbs_id": "1", "wbs_short_name": "B"},
        {"wbs_id": "4", "parent_wbs_id": "2", "wbs_short_name": "1"},
        {"wbs_id": "5", "parent_wbs_id": "3", "wbs_short_name": "1"},
    ]
    paths = analysis.wbs_paths(rows)
    assert paths["4"] == "PROJ/A/1" and paths["5"] == "PROJ/B/1"
    assert len(set(paths.values())) == len(rows)


# --- S-egrisi kovalari -----------------------------------------------------
def test_haftalik_kovalar_araligi_tam_kaplar():
    start, finish = _dt.date(2026, 1, 1), _dt.date(2026, 1, 31)
    b = analysis.buckets(start, finish, "week")
    assert b[0][0] == start and b[-1][1] == finish
    for i in range(len(b) - 1):
        assert b[i][1] + _dt.timedelta(days=1) == b[i + 1][0]


@pytest.mark.parametrize("bucket,expected_first_len", [
    ("day", 1), ("week", 7), ("month", 30)])
def test_kova_boyutlari(bucket, expected_first_len):
    b = analysis.buckets(_dt.date(2026, 1, 1), _dt.date(2026, 6, 30), bucket)
    assert (b[0][1] - b[0][0]).days + 1 == expected_first_len


def test_gecersiz_kova_reddedilir():
    with pytest.raises(analysis.AnalysisError, match="bucket"):
        analysis.buckets(_dt.date(2026, 1, 1), _dt.date(2026, 2, 1), "yil")


def test_tek_gunluk_aralik_tek_kova():
    b = analysis.buckets(_dt.date(2026, 1, 1), _dt.date(2026, 1, 1), "week")
    assert len(b) == 1 and b[0] == (_dt.date(2026, 1, 1), _dt.date(2026, 1, 1))


# --- tarih donusturme ------------------------------------------------------
@pytest.mark.parametrize("value,expected", [
    ("2026-09-24", _dt.date(2026, 9, 24)),
    ("2026-09-24 17:00:00", _dt.date(2026, 9, 24)),
    (_dt.date(2026, 9, 24), _dt.date(2026, 9, 24)),
    (_dt.datetime(2026, 9, 24, 17), _dt.date(2026, 9, 24)),
    (None, None), ("", None), ("bozuk", None),
])
def test_to_date(value, expected):
    assert analysis.to_date(value) == expected


# --- toplama ---------------------------------------------------------------
def test_aggregate_veri_tarihi_yoksa_reddeder():
    data = {"tasks": [], "status_date": None}
    with pytest.raises(analysis.AnalysisError, match="Veri tarihi"):
        analysis.aggregate(data)


def test_aggregate_temel_toplamlar():
    data = {
        "status_date": "2026-11-01",
        "tasks": [
            {"baseline_work": 100, "percent_complete": 50, "actual_work": 60,
             "baseline_start": "2026-09-01", "baseline_finish": "2026-09-30"},
            {"baseline_work": 200, "percent_complete": 0, "actual_work": 0,
             "baseline_start": "2026-12-01", "baseline_finish": "2026-12-31"},
        ],
    }
    agg = analysis.aggregate(data)
    assert agg["bac"] == 300
    assert agg["ev"] == 50          # 100*0.5
    assert agg["ac"] == 60
    assert agg["pv"] == 100         # ilk gorev veri tarihinde bitmis, ikincisi baslamamis
    assert agg["tasks_without_baseline_dates"] == 0


def test_aggregate_tarihsiz_gorevleri_sayar():
    data = {"status_date": "2026-11-01",
            "tasks": [{"baseline_work": 10, "percent_complete": 0,
                       "actual_work": 0, "baseline_start": None,
                       "baseline_finish": None}]}
    agg = analysis.aggregate(data)
    assert agg["tasks_without_baseline_dates"] == 1
    assert agg["pv"] == 0.0


def test_project_bounds_tarih_yoksa_hata():
    with pytest.raises(analysis.AnalysisError, match="baseline"):
        analysis.project_bounds({"tasks": [{"baseline_start": None,
                                            "baseline_finish": None}]})
